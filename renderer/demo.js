// Demo data — used until a Microsoft account is connected.
window.PlanitDemo = (() => {
  const day = (offset, h, m) => {
    const d = new Date();
    d.setDate(d.getDate() + offset);
    d.setHours(h, m || 0, 0, 0);
    return d;
  };
  const iso = (d) => d.toISOString();

  const people = [
    { name: 'You', email: 'you@example.com' },
    { name: 'Maya Chen', email: 'maya@example.com' },
    { name: 'Luis Ortega', email: 'luis@example.com' },
    { name: 'Sam Patel', email: 'sam@example.com' }
  ];

  const events = [
    { id: 'd1', subject: 'Team stand-up', start: iso(day(0, 9, 30)), end: iso(day(0, 9, 45)) },
    { id: 'd2', subject: 'Lunch with Maya', start: iso(day(0, 12, 30)), end: iso(day(0, 13, 30)), location: 'Verde Café' },
    { id: 'd3', subject: 'Sprint review', start: iso(day(1, 15, 0)), end: iso(day(1, 16, 0)) },
    { id: 'd4', subject: 'Dentist', start: iso(day(2, 8, 0)), end: iso(day(2, 9, 0)) },
    { id: 'd5', subject: 'Project Falcon planning', start: iso(day(3, 10, 0)), end: iso(day(3, 11, 30)) },
    { id: 'd6', subject: 'Gym', start: iso(day(4, 18, 0)), end: iso(day(4, 19, 0)) },
    { id: 'd7', subject: 'Monthly all-hands', start: iso(day(7, 11, 0)), end: iso(day(7, 12, 0)) },
    { id: 'd8', subject: 'Weekend trip planning call', start: iso(day(5, 20, 0)), end: iso(day(5, 20, 45)) },
    { id: 'd9', subject: 'Code review block', start: iso(day(-1, 14, 0)), end: iso(day(-1, 16, 0)) },
    { id: 'd10', subject: 'Birthday — Luis', start: iso(day(9, 0, 0)), end: iso(day(10, 0, 0)), isAllDay: true }
  ].map(e => ({
    id: e.id, subject: e.subject, isAllDay: !!e.isAllDay,
    start: { dateTime: e.start }, end: { dateTime: e.end },
    location: { displayName: e.location || '' }, demo: true
  }));

  const polls = [
    {
      id: 'p1',
      title: 'Q3 kickoff dinner',
      description: 'Pick an evening that works for the whole crew.',
      organizer: 'you@example.com',
      participants: people.map(p => p.email),
      finalized: null,
      slots: [
        { id: 's1', start: iso(day(4, 19, 0)), end: iso(day(4, 21, 0)) },
        { id: 's2', start: iso(day(5, 19, 0)), end: iso(day(5, 21, 0)) },
        { id: 's3', start: iso(day(6, 18, 30)), end: iso(day(6, 20, 30)) }
      ],
      votes: {
        'you@example.com': { s1: 'yes', s2: 'yes', s3: 'no' },
        'maya@example.com': { s1: 'yes', s2: 'maybe', s3: 'no' },
        'luis@example.com': { s1: 'no', s2: 'yes', s3: 'yes' },
        'sam@example.com': { s1: 'yes', s2: 'yes', s3: 'maybe' }
      },
      demo: true
    }
  ];

  // Availability grid: per person, per 30-min slot from 8:00–18:00 → status
  function demoSchedule(emails, startISO, endISO, intervalMin) {
    const start = new Date(startISO), end = new Date(endISO);
    const slots = Math.max(0, Math.round((end - start) / 60000 / intervalMin));
    return emails.map((email, pi) => {
      let view = '';
      for (let i = 0; i < slots; i++) {
        const t = new Date(start.getTime() + i * intervalMin * 60000);
        const h = t.getHours(), dow = t.getDay();
        let c = '0'; // free
        if (dow === 0 || dow === 6) c = '0';
        else if (h < 8 || h >= 18) c = '0';
        else {
          // deterministic pseudo-pattern per person
          const seed = (t.getDate() * 7 + h * 2 + Math.floor(t.getMinutes() / 30) + pi * 3) % 11;
          if (seed < 3) c = '2';        // busy
          else if (seed === 3) c = '1'; // tentative
          else if (seed === 4 && pi === 2) c = '3'; // oof
          else c = '0';
        }
        view += c;
      }
      return { scheduleId: email, availabilityView: view };
    });
  }

  // Demo group calendars — writable shared calendars where members plan events
  const groupCals = [
    { id: 'gc-demo', name: 'Trip crew', owner: { address: 'you@example.com', name: 'You' }, canEdit: true, demo: true }
  ];
  const groupCalEvents = [
    { id: 'gce1', calId: 'gc-demo', subject: 'Book the cabins — Sam', start: iso(day(2, 19, 0)), end: iso(day(2, 20, 0)) },
    { id: 'gce2', calId: 'gc-demo', subject: 'Pay deposit — You', start: iso(day(3, 12, 0)), end: iso(day(3, 12, 30)) },
    { id: 'gce3', calId: 'gc-demo', subject: 'Playlist + packing list — Maya', start: iso(day(5, 18, 0)), end: iso(day(5, 19, 0)) },
    { id: 'gce4', calId: 'gc-demo', subject: 'Road trip departure', start: iso(day(6, 8, 0)), end: iso(day(6, 10, 0)) },
    { id: 'gce5', calId: 'gc-demo', subject: 'Trip kickoff call (Teams)', start: iso(day(1, 20, 0)), end: iso(day(1, 20, 30)), isOnlineMeeting: true }
  ].map(e => ({
    id: e.id, calId: e.calId, subject: e.subject, isAllDay: false,
    isOnlineMeeting: !!e.isOnlineMeeting,
    start: { dateTime: e.start }, end: { dateTime: e.end },
    location: { displayName: '' }, demo: true
  }));

  // Shared planned events per group member (deterministic per day)
  function sharedEvents(email, startISO, endISO) {
    const titles = {
      'you@example.com': ['Focus block', 'Team stand-up', 'Gym', 'Groceries run'],
      'maya@example.com': ['Design review', 'Yoga class', 'Client call', 'Book club'],
      'luis@example.com': ['Build pipeline', 'School pickup', '5-a-side football', 'Guitar practice'],
      'sam@example.com': ['Roadmap sync', 'Study group', 'Family dinner', 'Morning run']
    };
    const pool = titles[email] || ['Busy'];
    const out = [];
    const start = new Date(startISO), end = new Date(endISO);
    for (let d = new Date(start); d < end; d.setDate(d.getDate() + 1)) {
      const n = (d.getDate() + email.length) % 3; // 0–2 events/day
      for (let i = 0; i < n; i++) {
        const h = 9 + ((d.getDate() * (i + 2) + email.length) % 9);
        const s = new Date(d); s.setHours(h, 0, 0, 0);
        const e2 = new Date(s.getTime() + 3600000);
        out.push({
          id: 'se-' + email + '-' + d.getDate() + '-' + i,
          subject: pool[(d.getDate() + i) % pool.length],
          start: { dateTime: s.toISOString() }, end: { dateTime: e2.toISOString() }
        });
      }
    }
    return out;
  }

  const mail = [
    {
      id: 'm1', subject: 'Re: Q3 kickoff dinner — I voted!',
      from: { emailAddress: { name: 'Maya Chen', address: 'maya@example.com' } },
      receivedDateTime: iso(day(0, 8, 12)), isRead: false, hasAttachments: false,
      bodyPreview: 'Thursday works best for me, Friday is a maybe…',
      body: 'Hey!\n\nThursday works best for me, Friday is a maybe. Saturday I have family plans.\n\nExcited for this!\nMaya'
    },
    {
      id: 'm2', subject: 'Sprint review moved to 3pm',
      from: { emailAddress: { name: 'Luis Ortega', address: 'luis@example.com' } },
      receivedDateTime: iso(day(-1, 16, 40)), isRead: true, hasAttachments: false,
      bodyPreview: 'Heads up — moved tomorrow’s review to 3pm so the design team can join.',
      body: 'Heads up — moved tomorrow’s review to 3pm so the design team can join.\n\nLuis'
    },
    {
      id: 'm3', subject: 'Weekend trip: cabin options',
      from: { emailAddress: { name: 'Sam Patel', address: 'sam@example.com' } },
      receivedDateTime: iso(day(-2, 21, 5)), isRead: true, hasAttachments: true,
      bodyPreview: 'Found three cabins near the lake, attaching pics. Vote in the poll when you can!',
      body: 'Found three cabins near the lake, attaching pics.\n\nVote in the poll when you can!\n\nSam'
    },
    {
      id: 'm4', subject: 'Your Planit demo inbox',
      from: { emailAddress: { name: 'Planit', address: 'hello@planit.app' } },
      receivedDateTime: iso(day(-3, 9, 0)), isRead: true, hasAttachments: false,
      bodyPreview: 'This is sample mail. Connect your Microsoft account in Settings to see your real inbox.',
      body: 'This is sample mail.\n\nConnect your Microsoft account in Settings → Microsoft account to see your real Outlook inbox, calendar, and group availability.\n\n— Planit'
    }
  ];

  return { people, events, polls, demoSchedule, sharedEvents, groupCals, groupCalEvents, mail };
})();
