/* Planit renderer — all views */
(() => {
  const $ = (sel, el) => (el || document).querySelector(sel);
  const $$ = (sel, el) => Array.from((el || document).querySelectorAll(sel));

  const state = {
    settings: { clientId: '', tenant: 'common' },
    account: null,          // MSAL account when signed in
    connected: false,       // signed in to Microsoft
    calCursor: new Date(),  // month being viewed
    calMode: 'month',       // 'month' | 'week'
    calSource: 'default',   // 'default' = your calendar, or a group/shared calendar id
    calendars: [],          // non-default calendars (group calendars, shared calendars)
    events: [],             // events for the visible range
    polls: [],
    groupPeople: [],
    mail: [],
    mailSel: null
  };

  /* ---------------- helpers ---------------- */
  function toast(msg, isError) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.toggle('error', !!isError);
    t.classList.remove('hidden');
    clearTimeout(t._h);
    t._h = setTimeout(() => t.classList.add('hidden'), 3400);
  }

  function modal(html) {
    $('#modal-card').innerHTML = html;
    $('#modal-root').classList.remove('hidden');
  }
  function closeModal() { $('#modal-root').classList.add('hidden'); }
  $('#modal-backdrop').addEventListener('click', closeModal);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Graph may return "2026-07-28T09:30:00.0000000" (no zone) — treat as UTC.
  function parseDT(v) {
    if (!v) return new Date(NaN);
    const s = typeof v === 'string' ? v : v.dateTime;
    if (/Z$|[+-]\d\d:\d\d$/.test(s)) return new Date(s);
    return new Date(s + 'Z');
  }
  const fmtTime = d => d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  const fmtDate = d => d.toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' });
  const fmtDateTime = d => `${fmtDate(d)} ${fmtTime(d)}`;
  const sameDay = (a, b) =>
    a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  const pad = n => String(n).padStart(2, '0');
  const toLocalInput = d =>
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const uid = () => 'x' + Math.random().toString(36).slice(2, 10);
  const localTZ = Intl.DateTimeFormat().resolvedOptions().timeZone;

  function graphErr(e) {
    const msg = (e && e.message) || String(e);
    if (msg.includes('NOT_CONFIGURED')) return 'Add your Azure client ID in Settings first.';
    if (msg.includes('NOT_SIGNED_IN')) return 'Sign in to Microsoft in Settings first.';
    return msg.replace(/^Error invoking remote method '[^']+': (Error: )?/, '');
  }

  /* ---------------- navigation ---------------- */
  $$('.nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      $$('.nav-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const view = btn.dataset.view;
      $$('.view').forEach(v => v.classList.remove('active'));
      $('#view-' + view).classList.add('active');
      renderers[view]();
    });
  });

  /* ================= CALENDAR ================= */
  async function loadCalendarList() {
    try {
      if (state.connected) {
        const res = await window.planit.graph.listCalendars();
        state.calendars = (res.value || []).filter(c => !c.isDefaultCalendar);
      } else {
        const local = await window.planit.store.get('localGroupCals', []);
        state.calendars = PlanitDemo.groupCals.concat(local);
      }
    } catch { state.calendars = []; }
    if (state.calSource !== 'default' && !state.calendars.find(c => c.id === state.calSource)) {
      state.calSource = 'default';
    }
  }

  async function loadEvents(rangeStart, rangeEnd) {
    const inRange = ev => {
      const s = parseDT(ev.start), e = parseDT(ev.end);
      return e >= rangeStart && s <= rangeEnd;
    };
    if (state.connected) {
      const res = state.calSource === 'default'
        ? await window.planit.graph.calendarView(rangeStart.toISOString(), rangeEnd.toISOString())
        : await window.planit.graph.calendarViewIn(state.calSource, rangeStart.toISOString(), rangeEnd.toISOString());
      state.events = res.value || [];
    } else {
      const local = await window.planit.store.get('localEvents', []);
      const base = state.calSource === 'default'
        ? PlanitDemo.events.concat(local.filter(ev => !ev.calId))
        : PlanitDemo.groupCalEvents.filter(ev => ev.calId === state.calSource)
            .concat(local.filter(ev => ev.calId === state.calSource));
      state.events = base.filter(inRange);
    }
  }

  function calRange() {
    const c = state.calCursor;
    if (state.calMode === 'month') {
      const first = new Date(c.getFullYear(), c.getMonth(), 1);
      const start = new Date(first);
      start.setDate(start.getDate() - start.getDay());
      const end = new Date(start);
      end.setDate(end.getDate() + 42);
      return { start, end };
    }
    const start = new Date(c);
    start.setDate(start.getDate() - start.getDay());
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setDate(end.getDate() + 7);
    return { start, end };
  }

  async function renderCalendar() {
    const root = $('#view-calendar');
    await loadCalendarList();
    const label = state.calMode === 'month'
      ? state.calCursor.toLocaleDateString([], { month: 'long', year: 'numeric' })
      : `Week of ${fmtDate(calRange().start)}`;
    root.innerHTML = `
      <div class="view-header">
        <div class="view-title">Calendar</div>
        <select id="cal-source" style="width:auto;max-width:190px" title="Which calendar to view">
          <option value="default">My calendar</option>
          ${state.calendars.map(c =>
            `<option value="${esc(c.id)}" ${state.calSource === c.id ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}
        </select>
        <div class="cal-nav">
          <button class="btn small" id="cal-prev">&#8249;</button>
          <div class="cal-label">${esc(label)}</div>
          <button class="btn small" id="cal-next">&#8250;</button>
          <button class="btn small" id="cal-today">Today</button>
        </div>
        <select id="cal-mode" style="width:auto">
          <option value="month" ${state.calMode === 'month' ? 'selected' : ''}>Month</option>
          <option value="week" ${state.calMode === 'week' ? 'selected' : ''}>Week</option>
        </select>
        <button class="btn primary" id="cal-new">+ New event</button>
      </div>
      <div id="cal-body" style="flex:1;display:flex;flex-direction:column;min-height:0;overflow:auto">
        <div class="spin">Loading events…</div>
      </div>`;

    $('#cal-source').onchange = e => { state.calSource = e.target.value; renderCalendar(); };
    $('#cal-prev').onclick = () => { shiftCal(-1); };
    $('#cal-next').onclick = () => { shiftCal(1); };
    $('#cal-today').onclick = () => { state.calCursor = new Date(); renderCalendar(); };
    $('#cal-mode').onchange = e => { state.calMode = e.target.value; renderCalendar(); };
    $('#cal-new').onclick = () => eventModal();

    const { start, end } = calRange();
    try {
      await loadEvents(start, end);
    } catch (e) {
      $('#cal-body').innerHTML = `<div class="empty">Couldn't load events: ${esc(graphErr(e))}</div>`;
      return;
    }
    if (state.calMode === 'month') renderMonthGrid(start);
    else renderWeekGrid(start);
  }

  function shiftCal(dir) {
    const c = state.calCursor;
    if (state.calMode === 'month') c.setMonth(c.getMonth() + dir);
    else c.setDate(c.getDate() + 7 * dir);
    renderCalendar();
  }

  function eventsOnDay(day) {
    return state.events.filter(ev => {
      const s = parseDT(ev.start), e = parseDT(ev.end);
      if (ev.isAllDay) {
        const eAdj = new Date(e.getTime() - 1);
        return day >= new Date(s.getFullYear(), s.getMonth(), s.getDate()) && day <= eAdj;
      }
      return sameDay(s, day);
    }).sort((a, b) => parseDT(a.start) - parseDT(b.start));
  }

  function renderMonthGrid(start) {
    const today = new Date();
    const month = state.calCursor.getMonth();
    const dows = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    let html = `<div class="cal-grid-head">${dows.map(d => `<div>${d}</div>`).join('')}</div><div class="cal-grid">`;
    for (let i = 0; i < 42; i++) {
      const day = new Date(start);
      day.setDate(day.getDate() + i);
      const evs = eventsOnDay(day);
      const shown = evs.slice(0, 3);
      html += `<div class="cal-cell ${day.getMonth() !== month ? 'other' : ''} ${sameDay(day, today) ? 'today' : ''}"
                    data-date="${day.toISOString()}">
        <div class="cal-daynum">${day.getDate()}</div>
        ${shown.map(ev => `<div class="cal-ev ${ev.isAllDay ? 'allday' : ''}" data-ev="${esc(ev.id)}">
            ${ev.isOnlineMeeting ? '🎥 ' : ''}${ev.isAllDay ? '' : fmtTime(parseDT(ev.start)) + ' '}${esc(ev.subject || '(no title)')}</div>`).join('')}
        ${evs.length > 3 ? `<div class="cal-more">+${evs.length - 3} more</div>` : ''}
      </div>`;
    }
    html += '</div>';
    $('#cal-body').innerHTML = html;

    $$('.cal-cell').forEach(cell => {
      cell.addEventListener('click', e => {
        const evEl = e.target.closest('.cal-ev');
        if (evEl) {
          const ev = state.events.find(x => String(x.id) === evEl.dataset.ev);
          if (ev) { eventModal(ev); return; }
        }
        const d = new Date(cell.dataset.date);
        d.setHours(9, 0, 0, 0);
        eventModal(null, d);
      });
    });
  }

  function renderWeekGrid(start) {
    const H0 = 6, H1 = 22, SLOT = 44; // px per hour
    const today = new Date();
    let head = '<div></div>';
    const days = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(start); d.setDate(d.getDate() + i); days.push(d);
      head += `<div style="text-align:center;padding:6px;font-size:12px;color:${sameDay(d, today) ? 'var(--accent)' : 'var(--muted)'}">
        ${d.toLocaleDateString([], { weekday: 'short' })} ${d.getDate()}</div>`;
    }
    let hoursCol = '';
    for (let h = H0; h < H1; h++) hoursCol += `<div class="week-slot"><div class="week-hour">${h}:00</div></div>`;
    let cols = '';
    for (const d of days) {
      const evs = eventsOnDay(d).filter(e => !e.isAllDay);
      let evHtml = '';
      for (const ev of evs) {
        const s = parseDT(ev.start), e = parseDT(ev.end);
        const top = Math.max(0, (s.getHours() + s.getMinutes() / 60 - H0) * SLOT);
        const height = Math.max(20, ((e - s) / 3600000) * SLOT - 3);
        evHtml += `<div class="week-ev" data-ev="${esc(ev.id)}" style="top:${top}px;height:${height}px">
          <b>${fmtTime(s)}</b> ${ev.isOnlineMeeting ? '🎥 ' : ''}${esc(ev.subject || '')}</div>`;
      }
      let slots = '';
      for (let h = H0; h < H1; h++) slots += `<div class="week-slot" data-date="${d.toISOString()}" data-hour="${h}"></div>`;
      cols += `<div class="week-col">${slots}${evHtml}</div>`;
    }
    $('#cal-body').innerHTML =
      `<div class="week-grid" style="grid-template-rows:auto">${head}</div>
       <div class="week-grid" style="flex:1">
         <div>${hoursCol}</div>${cols}
       </div>`;
    $$('.week-ev').forEach(el => el.addEventListener('click', () => {
      const ev = state.events.find(x => String(x.id) === el.dataset.ev);
      if (ev) eventModal(ev);
    }));
    $$('.week-slot[data-date]').forEach(el => el.addEventListener('click', () => {
      const d = new Date(el.dataset.date);
      d.setHours(Number(el.dataset.hour), 0, 0, 0);
      eventModal(null, d);
    }));
  }

  function eventModal(ev, defaultStart) {
    const isNew = !ev;
    const s = ev ? parseDT(ev.start) : (defaultStart || new Date());
    const e = ev ? parseDT(ev.end) : new Date((defaultStart || new Date()).getTime() + 3600000);
    modal(`
      <h2>${isNew ? 'New event' : 'Event'}</h2>
      <label class="field"><span>Title</span>
        <input id="ev-title" value="${esc(ev ? ev.subject : '')}" placeholder="What's happening?" /></label>
      <div class="row">
        <label class="field"><span>Starts</span><input id="ev-start" type="datetime-local" value="${toLocalInput(s)}" /></label>
        <label class="field"><span>Ends</span><input id="ev-end" type="datetime-local" value="${toLocalInput(e)}" /></label>
      </div>
      <label class="field"><span>Location</span>
        <input id="ev-loc" value="${esc(ev && ev.location ? ev.location.displayName : '')}" /></label>
      <label class="field"><span>Invite (comma-separated emails — sent via Outlook when connected)</span>
        <input id="ev-att" value="${esc(ev && ev.attendees ? ev.attendees.map(a => a.emailAddress.address).join(', ') : '')}" /></label>
      <label class="field" style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
        <input type="checkbox" id="ev-teams" style="width:auto" ${ev && ev.isOnlineMeeting ? 'checked' : ''} />
        <span style="margin:0;font-size:13px;color:var(--text)">Microsoft Teams meeting (adds a join link)</span>
      </label>
      ${ev && ev.onlineMeeting && ev.onlineMeeting.joinUrl
        ? `<button class="btn small" id="ev-join" style="margin-bottom:6px">🎥 Join Teams meeting</button>` : ''}
      ${state.calSource !== 'default'
        ? `<p class="hint" style="margin-bottom:6px">This event will be planned in the group calendar
           “${esc((state.calendars.find(c => c.id === state.calSource) || {}).name || '')}”.</p>` : ''}
      <div class="modal-actions">
        ${!isNew ? '<button class="btn danger" id="ev-delete">Delete</button>' : ''}
        <button class="btn" id="ev-cancel">Cancel</button>
        <button class="btn primary" id="ev-save">${isNew ? 'Create' : 'Save'}</button>
      </div>`);
    $('#ev-cancel').onclick = closeModal;
    const joinBtn = $('#ev-join');
    if (joinBtn) joinBtn.onclick = () => window.open(ev.onlineMeeting.joinUrl);
    if (!isNew) $('#ev-delete').onclick = async () => {
      try {
        if (state.connected) await window.planit.graph.deleteEvent(ev.id);
        else {
          if (ev.demo) { toast('Built-in demo events can’t be deleted.'); return; }
          const local = await window.planit.store.get('localEvents', []);
          await window.planit.store.set('localEvents', local.filter(x => x.id !== ev.id));
        }
        closeModal(); toast('Event deleted'); renderCalendar();
      } catch (err) { toast(graphErr(err), true); }
    };
    $('#ev-save').onclick = async () => {
      const title = $('#ev-title').value.trim() || '(no title)';
      const start = new Date($('#ev-start').value);
      const end = new Date($('#ev-end').value);
      if (!(start < end)) { toast('End must be after start.', true); return; }
      const attendees = $('#ev-att').value.split(',').map(x => x.trim()).filter(Boolean)
        .map(a => ({ emailAddress: { address: a }, type: 'required' }));
      const teams = $('#ev-teams').checked;
      try {
        if (state.connected) {
          const body = {
            subject: title,
            start: { dateTime: toLocalInput(start) + ':00', timeZone: localTZ },
            end: { dateTime: toLocalInput(end) + ':00', timeZone: localTZ },
            location: { displayName: $('#ev-loc').value.trim() },
            attendees
          };
          if (teams) { body.isOnlineMeeting = true; body.onlineMeetingProvider = 'teamsForBusiness'; }
          const doCreate = b => state.calSource === 'default'
            ? window.planit.graph.createEvent(b)
            : window.planit.graph.createEventIn(state.calSource, b);
          if (isNew) {
            try { await doCreate(body); }
            catch (err) {
              // Personal accounts don't support Teams-for-business links — retry with Skype provider.
              if (teams) { body.onlineMeetingProvider = 'skypeForConsumer'; await doCreate(body); }
              else throw err;
            }
          } else {
            await window.planit.graph.updateEvent(ev.id, body);
          }
        } else {
          const local = await window.planit.store.get('localEvents', []);
          const rec = {
            id: isNew ? uid() : ev.id, subject: title,
            calId: state.calSource !== 'default' ? state.calSource : undefined,
            isOnlineMeeting: teams,
            start: { dateTime: start.toISOString() }, end: { dateTime: end.toISOString() },
            location: { displayName: $('#ev-loc').value.trim() }, attendees
          };
          const next = isNew ? local.concat([rec]) : local.map(x => x.id === ev.id ? rec : x);
          if (!isNew && ev.demo) { toast('Built-in demo events can’t be edited — create a new one.'); return; }
          await window.planit.store.set('localEvents', next);
        }
        closeModal(); toast(isNew ? 'Event created' : 'Event updated'); renderCalendar();
      } catch (err) { toast(graphErr(err), true); }
    };
  }

  /* ================= POLLS ================= */
  async function loadPolls() {
    let polls = await window.planit.store.get('polls', null);
    if (polls === null) {
      polls = state.connected ? [] : PlanitDemo.polls;
      await window.planit.store.set('polls', polls);
    }
    state.polls = polls;
  }
  const savePolls = () => window.planit.store.set('polls', state.polls);

  const VOTE_NEXT = { '': 'yes', yes: 'maybe', maybe: 'no', no: '' };
  const VOTE_ICON = { yes: '✓', maybe: '~', no: '✗', '': '·' };

  function tally(poll, slotId) {
    let yes = 0, maybe = 0;
    for (const p of poll.participants) {
      const v = (poll.votes[p] || {})[slotId];
      if (v === 'yes') yes++;
      else if (v === 'maybe') maybe++;
    }
    return { yes, maybe, score: yes * 2 + maybe };
  }

  async function renderPolls() {
    const root = $('#view-polls');
    root.innerHTML = `
      <div class="view-header">
        <div class="view-title">Scheduling polls</div>
        <button class="btn primary" id="poll-new">+ New poll</button>
      </div>
      <div class="view-body" id="polls-body"><div class="spin">Loading…</div></div>`;
    $('#poll-new').onclick = () => pollModal();
    await loadPolls();
    const body = $('#polls-body');
    if (!state.polls.length) {
      body.innerHTML = `<div class="empty">No polls yet. Create one to find a time that works for everyone —
        like Doodle, but yours.</div>`;
      return;
    }
    body.innerHTML = state.polls.map(pollCard).join('');
    bindPollCards();
  }

  function pollCard(poll) {
    const best = poll.slots.reduce((a, b) => tally(poll, b.id).score > tally(poll, a.id).score ? b : a, poll.slots[0]);
    const rows = poll.participants.map(p => `
      <tr>
        <td>${esc(p)}</td>
        ${poll.slots.map(s => {
          const v = (poll.votes[p] || {})[s.id] || '';
          return `<td class="vote-cell vote-${v || 'none'} ${v ? 'vote-' + v : ''}"
            data-poll="${poll.id}" data-p="${esc(p)}" data-slot="${s.id}">${VOTE_ICON[v]}</td>`;
        }).join('')}
      </tr>`).join('');
    return `
      <div class="poll-card" id="poll-${poll.id}">
        <div class="poll-head">
          <div class="poll-title">${esc(poll.title)}</div>
          ${poll.finalized
            ? `<span class="badge done">Booked: ${esc(fmtDateTime(parseDT(poll.finalized.start)))}</span>`
            : '<span class="badge">Open</span>'}
        </div>
        ${poll.description ? `<div class="poll-desc">${esc(poll.description)}</div>` : ''}
        <table class="poll-table">
          <tr>
            <th style="min-width:160px">Participant</th>
            ${poll.slots.map(s => `<th class="${!poll.finalized && s === best ? 'winner' : ''}">
              ${esc(fmtDate(parseDT(s.start)))}<br>${fmtTime(parseDT(s.start))}–${fmtTime(parseDT(s.end))}
              <div class="slot-tally">${tally(poll, s.id).yes}✓ ${tally(poll, s.id).maybe}~</div></th>`).join('')}
          </tr>
          ${rows}
        </table>
        <div class="poll-actions">
          ${!poll.finalized ? `<button class="btn small" data-act="email" data-poll="${poll.id}">Email participants</button>
          <button class="btn primary small" data-act="finalize" data-poll="${poll.id}">Book best slot</button>` : ''}
          <button class="btn danger small" data-act="delete" data-poll="${poll.id}">Delete</button>
        </div>
        <div class="hint" style="margin-top:8px">Click a cell to cycle a vote: ✓ yes → ~ maybe → ✗ no.</div>
      </div>`;
  }

  function bindPollCards() {
    $$('.vote-cell').forEach(td => td.addEventListener('click', async () => {
      const poll = state.polls.find(p => p.id === td.dataset.poll);
      if (!poll || poll.finalized) return;
      const votes = poll.votes[td.dataset.p] || (poll.votes[td.dataset.p] = {});
      votes[td.dataset.slot] = VOTE_NEXT[votes[td.dataset.slot] || ''];
      await savePolls();
      renderPolls();
    }));
    $$('[data-act]').forEach(btn => btn.addEventListener('click', async () => {
      const poll = state.polls.find(p => p.id === btn.dataset.poll);
      if (!poll) return;
      const act = btn.dataset.act;
      if (act === 'delete') {
        state.polls = state.polls.filter(p => p.id !== poll.id);
        await savePolls(); renderPolls();
      } else if (act === 'email') {
        await emailPoll(poll);
      } else if (act === 'finalize') {
        await finalizePoll(poll);
      }
    }));
  }

  async function emailPoll(poll) {
    if (!state.connected) { toast('Connect your Microsoft account in Settings to send mail.', true); return; }
    const lines = poll.slots.map((s, i) =>
      `  Option ${i + 1}: ${fmtDateTime(parseDT(s.start))} – ${fmtTime(parseDT(s.end))}`).join('<br>');
    const message = {
      subject: `Vote: ${poll.title}`,
      body: {
        contentType: 'HTML',
        content: `<p>Hi! Please reply with which options work for you:</p><p>${lines}</p>
          <p>${esc(poll.description || '')}</p><p>— sent from Planit</p>`
      },
      toRecipients: poll.participants
        .filter(p => p !== (state.account && state.account.username))
        .map(a => ({ emailAddress: { address: a } }))
    };
    try {
      await window.planit.graph.sendMail(message);
      toast('Poll emailed to participants.');
    } catch (e) { toast(graphErr(e), true); }
  }

  async function finalizePoll(poll) {
    const best = poll.slots.reduce((a, b) => tally(poll, b.id).score > tally(poll, a.id).score ? b : a, poll.slots[0]);
    const s = parseDT(best.start), e = parseDT(best.end);
    try {
      if (state.connected) {
        const body = {
          subject: poll.title,
          body: { contentType: 'Text', content: (poll.description || '') + '\n\nBooked via Planit poll.' },
          start: { dateTime: toLocalInput(s) + ':00', timeZone: localTZ },
          end: { dateTime: toLocalInput(e) + ':00', timeZone: localTZ },
          attendees: poll.participants
            .filter(p => p !== (state.account && state.account.username))
            .map(a => ({ emailAddress: { address: a }, type: 'required' }))
        };
        if (poll.teams) { body.isOnlineMeeting = true; body.onlineMeetingProvider = 'teamsForBusiness'; }
        try { await window.planit.graph.createEvent(body); }
        catch (err) {
          if (poll.teams) { body.onlineMeetingProvider = 'skypeForConsumer'; await window.planit.graph.createEvent(body); }
          else throw err;
        }
      } else {
        const local = await window.planit.store.get('localEvents', []);
        local.push({
          id: uid(), subject: poll.title,
          start: { dateTime: s.toISOString() }, end: { dateTime: e.toISOString() }
        });
        await window.planit.store.set('localEvents', local);
      }
      poll.finalized = { start: best.start, end: best.end };
      await savePolls();
      toast(state.connected ? 'Event created and invites sent 🎉' : 'Event added to your calendar');
      renderPolls();
    } catch (err) { toast(graphErr(err), true); }
  }

  function pollModal() {
    const base = new Date(); base.setDate(base.getDate() + 1); base.setHours(18, 0, 0, 0);
    const mkSlot = (d) => `
      <div class="slot-row">
        <input type="datetime-local" class="slot-start" value="${toLocalInput(d)}" />
        <input type="number" class="slot-dur" value="60" min="15" step="15" style="width:90px" title="minutes" />
        <button class="btn small slot-del" title="Remove">✕</button>
      </div>`;
    modal(`
      <h2>New scheduling poll</h2>
      <label class="field"><span>Title</span><input id="pl-title" placeholder="Team dinner, sprint planning…" /></label>
      <label class="field"><span>Description (optional)</span><input id="pl-desc" /></label>
      <label class="field"><span>Participants (comma-separated emails)</span>
        <input id="pl-people" value="${esc(state.account ? state.account.username : 'you@example.com')}" /></label>
      <label class="field"><span>Proposed time slots (start + duration in minutes)</span></label>
      <div id="pl-slots">${mkSlot(base)}${mkSlot(new Date(base.getTime() + 86400000))}</div>
      <button class="btn small" id="pl-add">+ Add slot</button>
      <label class="field" style="display:flex;align-items:center;gap:8px;margin-top:12px">
        <input type="checkbox" id="pl-teams" style="width:auto" />
        <span style="margin:0;font-size:13px;color:var(--text)">Book the winning slot as a Microsoft Teams meeting</span>
      </label>
      <div class="modal-actions">
        <button class="btn" id="pl-cancel">Cancel</button>
        <button class="btn primary" id="pl-create">Create poll</button>
      </div>`);
    $('#pl-add').onclick = () => {
      $('#pl-slots').insertAdjacentHTML('beforeend', mkSlot(new Date(base.getTime() + 2 * 86400000)));
      bindSlotDel();
    };
    const bindSlotDel = () => $$('.slot-del').forEach(b => b.onclick = () => b.parentElement.remove());
    bindSlotDel();
    $('#pl-cancel').onclick = closeModal;
    $('#pl-create').onclick = async () => {
      const title = $('#pl-title').value.trim();
      if (!title) { toast('Give the poll a title.', true); return; }
      const participants = $('#pl-people').value.split(',').map(x => x.trim()).filter(Boolean);
      if (!participants.length) { toast('Add at least one participant.', true); return; }
      const slots = $$('#pl-slots .slot-row').map(row => {
        const start = new Date($('.slot-start', row).value);
        const dur = Number($('.slot-dur', row).value) || 60;
        return { id: uid(), start: start.toISOString(), end: new Date(start.getTime() + dur * 60000).toISOString() };
      }).filter(s => !isNaN(new Date(s.start)));
      if (!slots.length) { toast('Add at least one time slot.', true); return; }
      state.polls.unshift({
        id: uid(), title, description: $('#pl-desc').value.trim(),
        organizer: state.account ? state.account.username : 'you@example.com',
        teams: $('#pl-teams').checked,
        participants, slots, votes: {}, finalized: null
      });
      await savePolls();
      closeModal(); renderPolls();
    };
  }

  /* ================= GROUP ================= */
  const GROUP_COLORS = ['#1170a8', '#d99a2b', '#7b5bd6', '#34c98e', '#d8465f', '#2ab3ad', '#e2790f', '#5b8def'];
  const personColor = i => GROUP_COLORS[i % GROUP_COLORS.length];

  async function renderGroup() {
    const root = $('#view-group');
    if (!state.groupPeople.length) {
      state.groupPeople = await window.planit.store.get('groupPeople',
        state.connected ? [] : PlanitDemo.people.map(p => p.email));
    }
    if (!state.groupMode) state.groupMode = 'events';
    const today = new Date();
    root.innerHTML = `
      <div class="view-header">
        <div class="view-title">Group</div>
        <div class="seg">
          <button class="btn small ${state.groupMode === 'events' ? 'seg-on' : ''}" data-gm="events">Shared events</button>
          <button class="btn small ${state.groupMode === 'cals' ? 'seg-on' : ''}" data-gm="cals">Group calendars</button>
          <button class="btn small ${state.groupMode === 'busy' ? 'seg-on' : ''}" data-gm="busy">Free/busy</button>
        </div>
        <input type="date" id="gr-date" style="width:auto" value="${today.toISOString().slice(0, 10)}" />
        <select id="gr-days" style="width:auto">
          <option value="3">3 days</option>
          <option value="5">5 days</option>
          <option value="7" selected>7 days</option>
        </select>
        <button class="btn primary" id="gr-refresh">Refresh</button>
      </div>
      <div class="view-body">
        <div class="row" style="margin-bottom:10px;align-items:center">
          <div class="chips" id="gr-chips" style="flex:3"></div>
          <input id="gr-add" placeholder="add member email + Enter" style="flex:1;min-width:200px" />
        </div>
        <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap">
          <button class="btn small" id="gr-share">Share my calendar with the group</button>
          <span class="hint">Each member shares once (here or from Outlook) — then everyone's plans show up for the whole group.</span>
        </div>
        <div id="gr-grid" class="group-grid"><div class="empty">Loading…</div></div>
        <div id="gr-note" class="hint" style="margin-top:10px"></div>
      </div>`;

    const renderChips = () => {
      $('#gr-chips').innerHTML = state.groupPeople.map((p, i) =>
        `<span class="chip"><span class="dot" style="background:${personColor(i)}"></span>${esc(p)} <b data-p="${esc(p)}">✕</b></span>`).join('') ||
        '<span class="hint">No members yet — add emails on the right</span>';
      $$('#gr-chips b').forEach(b => b.onclick = async () => {
        state.groupPeople = state.groupPeople.filter(x => x !== b.dataset.p);
        await window.planit.store.set('groupPeople', state.groupPeople);
        renderChips(); loadGroup();
      });
    };
    renderChips();

    $('#gr-add').addEventListener('keydown', async e => {
      if (e.key !== 'Enter') return;
      const v = e.target.value.trim();
      if (v && !state.groupPeople.includes(v)) {
        state.groupPeople.push(v);
        await window.planit.store.set('groupPeople', state.groupPeople);
        renderChips(); loadGroup();
      }
      e.target.value = '';
    });
    $$('[data-gm]').forEach(b => b.onclick = () => { state.groupMode = b.dataset.gm; renderGroup(); });
    $('#gr-refresh').onclick = loadGroup;
    $('#gr-share').onclick = shareWithGroup;
    if (state.groupPeople.length) loadGroup();
    else $('#gr-grid').innerHTML = '<div class="empty">Add at least one member email above.</div>';
  }

  function loadGroup() {
    if (state.groupMode === 'busy') return loadGroupGrid();
    if (state.groupMode === 'cals') return renderGroupCals();
    return loadSharedEvents();
  }

  // -------- group calendars: shared calendars everyone can plan events in --------
  async function renderGroupCals() {
    const grid = $('#gr-grid');
    const note = $('#gr-note');
    grid.innerHTML = '<div class="spin">Loading calendars…</div>';
    note.textContent = '';
    let cals = [];
    try {
      if (state.connected) {
        const res = await window.planit.graph.listCalendars();
        cals = (res.value || []).filter(c => !c.isDefaultCalendar);
      } else {
        const local = await window.planit.store.get('localGroupCals', []);
        cals = PlanitDemo.groupCals.concat(local);
      }
    } catch (err) { grid.innerHTML = `<div class="empty">${esc(graphErr(err))}</div>`; return; }

    const mine = state.connected ? (state.account.username || '').toLowerCase() : 'you@example.com';
    grid.innerHTML = `
      <div style="padding:14px">
        <button class="btn primary small" id="gc-new">+ New group calendar</button>
        <div style="margin-top:12px">
        ${cals.length ? cals.map(c => {
          const owner = (c.owner && c.owner.address) || mine;
          const isMine = owner.toLowerCase() === mine;
          return `<div class="gc-row">
            <div><b>${esc(c.name)}</b>
              <div class="hint">${isMine ? 'created by you' : 'shared by ' + esc(owner)} ·
                ${c.canEdit !== false ? 'members with write access plan their own events here' : 'view only'}</div></div>
            ${isMine && state.connected ? `<button class="btn small" data-gc-add="${esc(c.id)}">Add member</button>` : ''}
            <button class="btn small" data-gc-open="${esc(c.id)}">Open</button>
          </div>`;
        }).join('') : '<div class="empty">No group calendars yet — create one and invite the crew.</div>'}
        </div>
      </div>`;
    $('#gc-new').onclick = groupCalModal;
    $$('[data-gc-open]').forEach(b => b.onclick = () => {
      state.calSource = b.dataset.gcOpen;
      $('.nav-btn[data-view="calendar"]').click();
    });
    $$('[data-gc-add]').forEach(b => b.onclick = () => addMemberModal(b.dataset.gcAdd));
    note.textContent = state.connected
      ? 'Invited members get an Outlook invite; once accepted, the calendar appears in their Planit calendar picker and they can plan their own events in it — Teams meetings included.'
      : 'Demo mode — group calendars are stored locally until you connect your Microsoft account.';
  }

  function groupCalModal() {
    modal(`
      <h2>New group calendar</h2>
      <label class="field"><span>Name</span><input id="gc-name" placeholder="Family, Trip crew, Dev team…" /></label>
      <label class="field"><span>Members (comma-separated emails) — they get write access</span>
        <input id="gc-members" value="${esc(state.groupPeople.join(', '))}" /></label>
      <div class="modal-actions">
        <button class="btn" id="gc-cancel">Cancel</button>
        <button class="btn primary" id="gc-create">Create &amp; invite</button>
      </div>`);
    $('#gc-cancel').onclick = closeModal;
    $('#gc-create').onclick = async () => {
      const name = $('#gc-name').value.trim();
      if (!name) { toast('Give the calendar a name.', true); return; }
      const members = $('#gc-members').value.split(',').map(x => x.trim()).filter(Boolean);
      try {
        if (state.connected) {
          const cal = await window.planit.graph.createCalendar(name);
          const mine = (state.account.username || '').toLowerCase();
          const fail = [];
          for (const m of members.filter(x => x.toLowerCase() !== mine)) {
            try { await window.planit.graph.shareCalendarIn(cal.id, m, 'write'); }
            catch (e) { fail.push(m); }
          }
          toast(fail.length
            ? `Calendar created — couldn't invite ${fail.join(', ')} (add them from Outlook → Calendar → Share).`
            : 'Group calendar created and invites sent 🎉');
        } else {
          const local = await window.planit.store.get('localGroupCals', []);
          local.push({ id: uid(), name, owner: { address: 'you@example.com' }, canEdit: true });
          await window.planit.store.set('localGroupCals', local);
          toast('Group calendar created (demo — stored locally).');
        }
        closeModal(); renderGroupCals();
      } catch (err) { toast(graphErr(err), true); }
    };
  }

  function addMemberModal(calId) {
    modal(`
      <h2>Add member</h2>
      <label class="field"><span>Email</span><input id="am-mail" placeholder="friend@outlook.com" /></label>
      <label class="field"><span>Access</span>
        <select id="am-role">
          <option value="write">Can plan events (write)</option>
          <option value="read">View only</option>
        </select></label>
      <div class="modal-actions">
        <button class="btn" id="am-cancel">Cancel</button>
        <button class="btn primary" id="am-add">Send invite</button>
      </div>`);
    $('#am-cancel').onclick = closeModal;
    $('#am-add').onclick = async () => {
      const email = $('#am-mail').value.trim();
      if (!email) { toast('Enter an email.', true); return; }
      try {
        await window.planit.graph.shareCalendarIn(calId, email, $('#am-role').value);
        toast('Invite sent to ' + email);
        closeModal();
      } catch (err) { toast(graphErr(err), true); }
    };
  }

  async function shareWithGroup() {
    if (!state.connected) { toast('Connect your Microsoft account in Settings first.', true); return; }
    const mine = (state.account.username || '').toLowerCase();
    const others = state.groupPeople.filter(p => p.toLowerCase() !== mine);
    if (!others.length) { toast('Add group members first.', true); return; }
    let ok = 0; const fail = [];
    for (const email of others) {
      try { await window.planit.graph.shareCalendar(email, 'read'); ok++; }
      catch (e) { fail.push(email); }
    }
    if (ok) toast(`Calendar shared with ${ok} member${ok > 1 ? 's' : ''} — they'll get an Outlook invite.`);
    if (fail.length) toast(`Couldn't share with ${fail.join(', ')} — already shared, or share from Outlook (Calendar → Share).`, !ok);
  }

  // -------- shared planned events (the group calendar) --------
  function eventsOfDay(events, day) {
    return events.filter(ev => {
      const s = parseDT(ev.start), e = parseDT(ev.end);
      if (ev.isAllDay) {
        const eAdj = new Date(e.getTime() - 1);
        return day >= new Date(s.getFullYear(), s.getMonth(), s.getDate()) && day <= eAdj;
      }
      return sameDay(s, day);
    });
  }

  async function loadSharedEvents() {
    const grid = $('#gr-grid');
    const note = $('#gr-note');
    grid.innerHTML = '<div class="spin">Loading shared calendars…</div>';
    note.textContent = '';
    const start = new Date($('#gr-date').value + 'T00:00:00');
    const days = Number($('#gr-days').value);
    const end = new Date(start); end.setDate(end.getDate() + days);
    let perPerson = [];
    try {
      if (state.connected) {
        const cals = ((await window.planit.graph.listCalendars()) || {}).value || [];
        const mine = (state.account.username || '').toLowerCase();
        for (const email of state.groupPeople) {
          const low = email.toLowerCase();
          if (low === mine) {
            const res = await window.planit.graph.calendarView(start.toISOString(), end.toISOString());
            perPerson.push({ email, label: 'You', events: res.value || [] });
            continue;
          }
          const cal = cals.find(c => c.owner && c.owner.address && c.owner.address.toLowerCase() === low);
          if (!cal) { perPerson.push({ email, label: email, events: [], missing: true }); continue; }
          const res = await window.planit.graph.calendarViewIn(cal.id, start.toISOString(), end.toISOString());
          perPerson.push({ email, label: cal.owner.name || email, events: res.value || [] });
        }
        const missing = perPerson.filter(p => p.missing).map(p => p.email);
        note.innerHTML = missing.length
          ? `Waiting on shared calendars from: <b>${missing.map(esc).join(', ')}</b> — once they accept your invite and
             share back (their Planit's “Share my calendar” button, or Outlook → Calendar → Share), their plans appear here.`
          : 'Showing your calendar plus the calendars group members shared with you.';
      } else {
        perPerson = state.groupPeople.map(email => ({
          email,
          label: (PlanitDemo.people.find(p => p.email === email) || { name: email }).name,
          events: PlanitDemo.sharedEvents(email, start.toISOString(), end.toISOString())
        }));
        note.textContent = 'Demo data — connect your Microsoft account in Settings to see real shared calendars.';
      }
    } catch (err) {
      grid.innerHTML = `<div class="empty">Couldn't load shared calendars: ${esc(graphErr(err))}</div>`;
      return;
    }

    let cols = '';
    const today = new Date();
    for (let d = 0; d < days; d++) {
      const day = new Date(start); day.setDate(day.getDate() + d);
      const items = [];
      perPerson.forEach((p, pi) => {
        for (const ev of eventsOfDay(p.events, day)) items.push({ ev, pi, p });
      });
      items.sort((a, b) => parseDT(a.ev.start) - parseDT(b.ev.start));
      cols += `
        <div class="ge-col">
          <div class="ge-day ${sameDay(day, today) ? 'ge-today' : ''}">
            ${day.toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' })}</div>
          ${items.map(({ ev, pi, p }) => `
            <div class="ge-ev" style="border-left-color:${personColor(pi)}">
              <div class="t">${ev.isAllDay ? 'All day' : fmtTime(parseDT(ev.start))} · ${esc(p.label)}</div>
              ${esc(ev.subject || '(no title)')}
            </div>`).join('') || '<div class="ge-none">—</div>'}
        </div>`;
    }
    grid.innerHTML = `<div class="ge-cols">${cols}</div>`;
  }

  async function loadGroupGrid() {
    const grid = $('#gr-grid');
    $('#gr-note').textContent = '';
    if (!state.groupPeople.length) { grid.innerHTML = '<div class="empty">Add at least one email above.</div>'; return; }
    grid.innerHTML = '<div class="spin">Checking schedules…</div>';
    const dateStr = $('#gr-date').value;
    const days = Number($('#gr-days').value);
    const H0 = 8, H1 = 18, INT = 30;
    const start = new Date(dateStr + 'T00:00:00');
    start.setHours(H0, 0, 0, 0);

    const perDay = [];
    try {
      for (let d = 0; d < days; d++) {
        const s = new Date(start); s.setDate(s.getDate() + d);
        const e = new Date(s); e.setHours(H1, 0, 0, 0);
        if (state.connected) {
          const res = await window.planit.graph.getSchedule(
            state.groupPeople, s.toISOString(), e.toISOString(), INT);
          perDay.push({ day: s, schedules: res.value || [] });
        } else {
          perDay.push({ day: s, schedules: PlanitDemo.demoSchedule(state.groupPeople, s.toISOString(), e.toISOString(), INT) });
        }
      }
    } catch (err) {
      grid.innerHTML = `<div class="empty">Couldn't fetch availability: ${esc(graphErr(err))}<br><br>
        <span class="hint">Free/busy lookup requires work/school (Microsoft 365) accounts — for personal accounts,
        use the “Shared events” tab instead.</span></div>`;
      return;
    }

    const slotsPerDay = (H1 - H0) * 60 / INT;
    const cls = { '0': 'gg-free', '1': 'gg-tent', '2': 'gg-busy', '3': 'gg-oof', '4': 'gg-tent' };
    let head = '<tr><th class="gg-name">Person</th>';
    for (const { day } of perDay) {
      head += `<th colspan="${slotsPerDay}">${esc(day.toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' }))}
        <div class="slot-tally">${H0}:00 – ${H1}:00</div></th>`;
    }
    head += '</tr>';
    let rows = '';
    for (const person of state.groupPeople) {
      let cells = '';
      for (const { schedules } of perDay) {
        const sched = schedules.find(s => (s.scheduleId || '').toLowerCase() === person.toLowerCase());
        const view = (sched && sched.availabilityView) || ''.padEnd(slotsPerDay, '9');
        for (let i = 0; i < slotsPerDay; i++) {
          const c = view[i];
          const hh = H0 + Math.floor(i * INT / 60), mm = (i * INT) % 60;
          cells += `<td class="gg-slot ${cls[c] || 'gg-unknown'}" title="${esc(person)} · ${pad(hh)}:${pad(mm)}"></td>`;
        }
      }
      rows += `<tr><td class="gg-name">${esc(person)}</td>${cells}</tr>`;
    }
    grid.innerHTML = `<table class="gg-table">${head}${rows}</table>`;
  }

  /* ================= MAIL ================= */
  async function renderMail() {
    const root = $('#view-mail');
    root.innerHTML = `
      <div class="view-header">
        <div class="view-title">Mail</div>
        <button class="btn" id="mail-refresh">Refresh</button>
        <button class="btn primary" id="mail-compose">+ Compose</button>
      </div>
      <div class="mail-layout">
        <div class="mail-list" id="mail-list"><div class="spin">Loading…</div></div>
        <div class="mail-read" id="mail-read"><div class="empty">Select a message</div></div>
      </div>`;
    $('#mail-refresh').onclick = renderMail;
    $('#mail-compose').onclick = composeModal;
    try {
      if (state.connected) {
        const res = await window.planit.graph.listMail(30, 0);
        state.mail = res.value || [];
      } else {
        state.mail = PlanitDemo.mail;
      }
    } catch (e) {
      $('#mail-list').innerHTML = `<div class="empty">${esc(graphErr(e))}</div>`;
      return;
    }
    const list = $('#mail-list');
    if (!state.mail.length) { list.innerHTML = '<div class="empty">Inbox empty</div>'; return; }
    list.innerHTML = state.mail.map(m => `
      <div class="mail-item ${m.isRead ? '' : 'unread'}" data-id="${esc(m.id)}">
        <div class="m-from"><span>${esc(m.from && m.from.emailAddress ? (m.from.emailAddress.name || m.from.emailAddress.address) : '?')}</span>
          <span class="m-date">${esc(new Date(m.receivedDateTime).toLocaleDateString([], { month: 'short', day: 'numeric' }))}</span></div>
        <div class="m-subj">${esc(m.subject || '(no subject)')}</div>
        <div class="m-prev">${esc(m.bodyPreview || '')}</div>
      </div>`).join('');
    $$('.mail-item').forEach(el => el.addEventListener('click', () => openMail(el.dataset.id)));
  }

  async function openMail(id) {
    $$('.mail-item').forEach(el => el.classList.toggle('sel', el.dataset.id === id));
    const pane = $('#mail-read');
    pane.innerHTML = '<div class="spin">Opening…</div>';
    let m;
    try {
      m = state.connected
        ? await window.planit.graph.readMail(id)
        : state.mail.find(x => x.id === id);
    } catch (e) { pane.innerHTML = `<div class="empty">${esc(graphErr(e))}</div>`; return; }
    if (!m) { pane.innerHTML = '<div class="empty">Message not found</div>'; return; }
    const from = m.from && m.from.emailAddress ? m.from.emailAddress : {};
    const isHtml = m.body && m.body.contentType && m.body.contentType.toLowerCase() === 'html';
    const bodyContent = m.body ? m.body.content : (m.bodyPreview || '');
    pane.innerHTML = `
      <h2>${esc(m.subject || '(no subject)')}</h2>
      <div class="mail-meta">From: <b>${esc(from.name || '')}</b> &lt;${esc(from.address || '')}&gt;
        · ${esc(new Date(m.receivedDateTime).toLocaleString())}</div>
      <div class="mail-body" id="mail-body-slot"></div>
      <div style="margin-top:16px"><button class="btn" id="mail-reply">Reply</button></div>`;
    if (isHtml) {
      const iframe = document.createElement('iframe');
      iframe.setAttribute('sandbox', '');
      iframe.srcdoc = bodyContent;
      $('#mail-body-slot').appendChild(iframe);
    } else {
      $('#mail-body-slot').textContent = typeof bodyContent === 'string' ? bodyContent : (m.body || '');
    }
    $('#mail-reply').onclick = () => composeModal({
      to: from.address || '',
      subject: /^re:/i.test(m.subject || '') ? m.subject : 'Re: ' + (m.subject || '')
    });
  }

  function composeModal(prefill) {
    prefill = prefill || {};
    modal(`
      <h2>New message</h2>
      <label class="field"><span>To (comma-separated)</span><input id="cm-to" value="${esc(prefill.to || '')}" /></label>
      <label class="field"><span>Subject</span><input id="cm-subj" value="${esc(prefill.subject || '')}" /></label>
      <label class="field"><span>Message</span><textarea id="cm-body" rows="9"></textarea></label>
      <div class="modal-actions">
        <button class="btn" id="cm-cancel">Cancel</button>
        <button class="btn primary" id="cm-send">Send</button>
      </div>`);
    $('#cm-cancel').onclick = closeModal;
    $('#cm-send').onclick = async () => {
      if (!state.connected) { toast('Connect your Microsoft account in Settings to send mail.', true); return; }
      const to = $('#cm-to').value.split(',').map(x => x.trim()).filter(Boolean);
      if (!to.length) { toast('Add a recipient.', true); return; }
      try {
        await window.planit.graph.sendMail({
          subject: $('#cm-subj').value.trim(),
          body: { contentType: 'Text', content: $('#cm-body').value },
          toRecipients: to.map(a => ({ emailAddress: { address: a } }))
        });
        closeModal(); toast('Sent ✉');
      } catch (e) { toast(graphErr(e), true); }
    };
  }

  /* ================= SETTINGS ================= */
  async function renderSettings() {
    const root = $('#view-settings');
    const s = state.settings;
    root.innerHTML = `
      <div class="view-header"><div class="view-title">Settings</div></div>
      <div class="view-body">
        <div class="settings-card">
          <h3>Microsoft account</h3>
          <p class="hint" style="margin-bottom:14px">
            Planit talks directly to Microsoft 365 (Outlook mail + calendar) via Microsoft Graph.
            To enable it you need a free <b>app registration</b> — see README.md for the 5-minute setup.
            Paste the Application (client) ID below.
          </p>
          <label class="field"><span>Application (client) ID</span>
            <input id="st-client" value="${esc(s.clientId || '')}" placeholder="00000000-0000-0000-0000-000000000000" /></label>
          <label class="field"><span>Tenant</span>
            <select id="st-tenant">
              <option value="common" ${s.tenant === 'common' ? 'selected' : ''}>common — any account (work/school or personal)</option>
              <option value="consumers" ${s.tenant === 'consumers' ? 'selected' : ''}>consumers — personal accounts only</option>
              <option value="organizations" ${s.tenant === 'organizations' ? 'selected' : ''}>organizations — work/school only</option>
            </select></label>
          <div class="row" style="align-items:center">
            <button class="btn" id="st-save">Save</button>
            <button class="btn primary" id="st-signin">${state.connected ? 'Switch account' : 'Sign in to Microsoft'}</button>
            ${state.connected ? '<button class="btn danger" id="st-signout">Sign out</button>' : ''}
          </div>
          <p class="hint" style="margin-top:12px" id="st-status">
            ${state.connected
              ? `<span class="status-ok">● Connected</span> as ${esc(state.account.username)}`
              : '<span class="status-warn">● Not connected</span> — the app is showing demo data.'}
          </p>
        </div>
        <div class="settings-card">
          <h3>About &amp; updates</h3>
          <p class="hint">Planit <span id="st-version">…</span> — personal + group calendar planner with Doodle-style polls.<br>
          Local data (polls, settings, offline events) is stored on this computer only.<br>
          Planit checks for updates automatically on launch and installs them on restart.</p>
          <div style="margin-top:10px"><button class="btn small" id="st-update">Check for updates now</button></div>
        </div>
      </div>`;
    window.planit.app.version().then(v => { $('#st-version').textContent = 'v' + v; });
    $('#st-update').onclick = async () => {
      const r = await window.planit.app.checkUpdates();
      if (!r.ok && r.reason === 'dev') { toast('Updates only apply to the installed app (not npm start).'); return; }
      if (!r.ok) { toast('Could not check for updates: ' + r.reason, true); return; }
      if (r.latest && r.latest !== r.current) toast(`Update ${r.latest} found — downloading in the background.`);
      else toast(`You're up to date (v${r.current}).`);
    };
    $('#st-save').onclick = saveSettings;
    $('#st-signin').onclick = async () => {
      await saveSettings(true);
      if (!state.settings.clientId) { toast('Enter your client ID first.', true); return; }
      toast('Opening Microsoft sign-in in your browser…');
      try {
        const res = await window.planit.auth.signIn();
        if (res && res.ok) {
          await refreshAccount();
          toast(`Signed in as ${res.account.username}`);
          renderSettings();
        }
      } catch (e) { toast(graphErr(e), true); }
    };
    const so = $('#st-signout');
    if (so) so.onclick = async () => {
      await window.planit.auth.signOut();
      await refreshAccount();
      toast('Signed out');
      renderSettings();
    };
  }

  async function saveSettings(silent) {
    state.settings.clientId = $('#st-client').value.trim();
    state.settings.tenant = $('#st-tenant').value;
    await window.planit.settings.set(state.settings);
    if (!silent) toast('Settings saved');
  }

  /* ================= init ================= */
  async function refreshAccount() {
    state.account = await window.planit.auth.account();
    state.connected = !!state.account;
    const chip = $('#account-chip');
    const banner = $('#demo-banner');
    if (state.connected) {
      chip.classList.remove('hidden');
      banner.classList.add('hidden');
      $('#account-name').textContent = state.account.name || state.account.username;
      $('#account-mail').textContent = state.account.username;
      $('#account-avatar').textContent = (state.account.name || state.account.username || '?')
        .split(' ').map(w => w[0]).slice(0, 2).join('').toUpperCase();
    } else {
      chip.classList.add('hidden');
      banner.classList.remove('hidden');
    }
  }

  const renderers = {
    calendar: renderCalendar,
    polls: renderPolls,
    group: renderGroup,
    mail: renderMail,
    settings: renderSettings
  };

  (async function init() {
    state.settings = await window.planit.settings.get();
    await window.planit.settings.set(state.settings); // ensures auth is configured
    await refreshAccount();
    renderCalendar();
  })();
})();
