// Microsoft identity (MSAL) — authorization code + PKCE via system browser + loopback redirect.
// Tokens are cached on disk in the app's userData folder and silently refreshed.
const { PublicClientApplication, CryptoProvider } = require('@azure/msal-node');
const { shell } = require('electron');
const http = require('http');
const fs = require('fs');
const path = require('path');

const SCOPES = [
  'User.Read',
  'Calendars.ReadWrite',
  'Calendars.Read.Shared',
  'Mail.Read',
  'Mail.Send'
];

let pca = null;
let cachePath = null;
let currentSettings = null;

function cachePlugin() {
  return {
    beforeCacheAccess: async (ctx) => {
      try {
        ctx.tokenCache.deserialize(fs.readFileSync(cachePath, 'utf8'));
      } catch { /* no cache yet */ }
    },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) {
        fs.mkdirSync(path.dirname(cachePath), { recursive: true });
        fs.writeFileSync(cachePath, ctx.tokenCache.serialize(), 'utf8');
      }
    }
  };
}

function configure(settings, userDataDir) {
  const clientId = settings && settings.clientId;
  if (!clientId) { pca = null; currentSettings = settings; return; }
  // Rebuild only if config changed
  if (pca && currentSettings &&
      currentSettings.clientId === settings.clientId &&
      currentSettings.tenant === settings.tenant) return;
  currentSettings = settings;
  cachePath = path.join(userDataDir, 'msal-cache.json');
  pca = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${settings.tenant || 'common'}`
    },
    cache: { cachePlugin: cachePlugin() }
  });
}

function requireClient() {
  if (!pca) {
    const err = new Error('NOT_CONFIGURED');
    err.code = 'NOT_CONFIGURED';
    throw err;
  }
}

async function getAccount() {
  if (!pca) return null;
  const accounts = await pca.getTokenCache().getAllAccounts();
  return accounts[0] || null;
}

// Interactive sign-in: open system browser, catch redirect on a loopback port.
async function signIn() {
  requireClient();
  const crypto = new CryptoProvider();
  const { verifier, challenge } = await crypto.generatePkceCodes();

  // Loopback server on an ephemeral port
  const { server, port, codePromise } = await startLoopback();
  const redirectUri = `http://localhost:${port}`;

  try {
    const authUrl = await pca.getAuthCodeUrl({
      scopes: SCOPES,
      redirectUri,
      codeChallenge: challenge,
      codeChallengeMethod: 'S256',
      prompt: 'select_account'
    });
    shell.openExternal(authUrl);

    const code = await codePromise; // throws on timeout/denial
    const result = await pca.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri,
      codeVerifier: verifier
    });
    return {
      ok: true,
      account: {
        username: result.account.username,
        name: result.account.name
      }
    };
  } finally {
    server.close();
  }
}

function startLoopback() {
  return new Promise((resolveOuter, rejectOuter) => {
    let resolveCode, rejectCode;
    const codePromise = new Promise((res, rej) => { resolveCode = res; rejectCode = rej; });
    const timeout = setTimeout(
      () => rejectCode(new Error('Sign-in timed out after 5 minutes.')),
      5 * 60 * 1000
    );

    const server = http.createServer((req, res) => {
      const url = new URL(req.url, 'http://localhost');
      const code = url.searchParams.get('code');
      const error = url.searchParams.get('error');
      if (!code && !error) { res.writeHead(404); res.end(); return; }
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(`<html><body style="font-family:sans-serif;background:#052c46;color:#ffffff;
        display:flex;align-items:center;justify-content:center;height:100vh">
        <div style="text-align:center"><h2>${code ? 'Signed in to Planit' : 'Sign-in failed'}</h2>
        <p>You can close this tab and return to the app.</p></div></body></html>`);
      clearTimeout(timeout);
      if (code) resolveCode(code);
      else rejectCode(new Error(url.searchParams.get('error_description') || error));
    });

    server.on('error', rejectOuter);
    server.listen(0, '127.0.0.1', () => {
      resolveOuter({ server, port: server.address().port, codePromise });
    });
  });
}

async function getToken() {
  requireClient();
  const account = await getAccount();
  if (!account) {
    const err = new Error('NOT_SIGNED_IN');
    err.code = 'NOT_SIGNED_IN';
    throw err;
  }
  const result = await pca.acquireTokenSilent({ account, scopes: SCOPES });
  return result.accessToken;
}

async function signOut() {
  if (!pca) return true;
  const cache = pca.getTokenCache();
  const accounts = await cache.getAllAccounts();
  for (const a of accounts) await cache.removeAccount(a);
  return true;
}

module.exports = { configure, signIn, signOut, getAccount, getToken };
