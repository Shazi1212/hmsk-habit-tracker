/* Habit Tracker core: dates, schema, streaks, stats. No DOM, no dependencies.
   Works in the browser (window.HTCore) and in Node (require) for the tests.

   Dates are calendar keys "YYYY-MM-DD" in the user's local time. Day arithmetic is done on the
   calendar (via Date.UTC on the y/m/d parts), never by adding 86400000 ms to a local Date, so
   daylight-saving changes and time zones cannot shift a day. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HTCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SCHEMA = 2;
  const DAY_MS = 86400000;

  // ---------- dates ----------
  const pad = n => String(n).padStart(2, '0');
  function keyOf(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  /** Local calendar day of a Date (or now). */
  function todayKey(now) {
    const d = now instanceof Date ? now : new Date();
    return keyOf(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }
  function isKey(k) {
    if (typeof k !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(k)) return false;
    const [y, m, d] = parts(k);
    return m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m);
  }
  function parts(k) { return [Number(k.slice(0, 4)), Number(k.slice(5, 7)), Number(k.slice(8, 10))]; }
  /** Days since 1970-01-01 on the calendar (time-zone free). */
  function dayNum(k) { const [y, m, d] = parts(k); return Math.round(Date.UTC(y, m - 1, d) / DAY_MS); }
  function fromDayNum(n) { const d = new Date(n * DAY_MS); return keyOf(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
  function addDays(k, n) { return fromDayNum(dayNum(k) + n); }
  function diffDays(a, b) { return dayNum(b) - dayNum(a); }
  /** ISO weekday, Monday = 1 ... Sunday = 7. */
  function weekday(k) { const w = new Date(dayNum(k) * DAY_MS).getUTCDay(); return w === 0 ? 7 : w; }
  /** Monday of the week containing k (weeks start on Monday, as in Germany / ISO 8601). */
  function weekStart(k) { return addDays(k, 1 - weekday(k)); }
  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function monthStart(k) { return k.slice(0, 8) + '01'; }
  function monthEnd(k) { const [y, m] = parts(k); return keyOf(y, m, daysInMonth(y, m)); }
  function range(a, b) { const out = []; for (let n = dayNum(a), e = dayNum(b); n <= e; n++) out.push(fromDayNum(n)); return out; }
  /** ISO 8601 week number and week-year. */
  function isoWeek(k) {
    const thu = addDays(k, 4 - weekday(k));
    const y = Number(thu.slice(0, 4));
    return { year: y, week: Math.floor(diffDays(keyOf(y, 1, 1), thu) / 7) + 1 };
  }
  /** Excel serial date (1900 system, days since 1899-12-30). */
  function excelSerial(k) { return dayNum(k) + 25569; }

  // ---------- schema ----------
  function uid() { return 'h' + Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-4); }
  function emptyState(today) {
    return { v: SCHEMA, habits: [], checks: {}, notes: {}, settings: { lang: '' }, meta: { firstUse: today || todayKey(), useDays: [], emailOfferHidden: false } };
  }
  function normTarget(t) {
    if (t && t.type === 'weekly') { const n = Math.round(Number(t.n)); if (n >= 1 && n <= 6) return { type: 'weekly', n }; }
    return { type: 'daily' };
  }
  function normHabit(h, i, today) {
    return {
      id: typeof h.id === 'string' && h.id ? h.id : uid(),
      name: String(h.name || '').trim().slice(0, 60) || 'Habit',
      icon: String(h.icon || '').slice(0, 4),
      target: normTarget(h.target),
      archived: !!h.archived,
      created: isKey(h.created) ? h.created : today,
      order: Number.isFinite(h.order) ? h.order : i
    };
  }

  /** Bring any stored object (v1 prototype format or v2) to the current schema. Throws on garbage. */
  function migrate(raw, today) {
    today = today || todayKey();
    if (!raw || typeof raw !== 'object') throw new Error('not a habit tracker backup');
    let s = raw;
    const v = Number(s.v || s.version || 1);
    if (v > SCHEMA) throw new Error('backup is from a newer version (' + v + ')');
    if (v < 2) {
      // v1: { version: 1, habits: [{ id, name, emoji, perWeek (0 = daily), done: [dates] }], notes: { 'YYYY-Www': text } }
      if (!Array.isArray(s.habits)) throw new Error('not a habit tracker backup');
      const checks = {}; const notes = {};
      const habits = s.habits.map((h, i) => {
        const nh = normHabit({ id: h.id, name: h.name, icon: h.emoji || h.icon, archived: h.archived, created: h.created,
          target: h.perWeek > 0 && h.perWeek < 7 ? { type: 'weekly', n: h.perWeek } : { type: 'daily' }, order: i }, i, today);
        checks[nh.id] = {};
        for (const d of h.done || []) if (isKey(d)) checks[nh.id][d] = 1;
        return nh;
      });
      for (const [wk, text] of Object.entries(s.notes || {})) {
        const m = /^(\d{4})-W(\d{2})$/.exec(wk);
        if (m && text) notes[isoWeekMonday(Number(m[1]), Number(m[2]))] = String(text);
      }
      s = { v: 2, habits, checks, notes, settings: s.settings || {}, meta: s.meta || {} };
    }
    // v2 clean-up (also validates imports)
    if (!Array.isArray(s.habits)) throw new Error('not a habit tracker backup');
    const out = emptyState(today);
    out.habits = s.habits.slice(0, 50).map((h, i) => normHabit(h, i, today)).sort((a, b) => a.order - b.order).map((h, i) => ({ ...h, order: i }));
    const ids = new Set(out.habits.map(h => h.id));
    for (const [id, days] of Object.entries(s.checks || {})) {
      if (!ids.has(id) || !days || typeof days !== 'object') continue;
      out.checks[id] = {};
      for (const d of Object.keys(days)) if (days[d] && isKey(d)) out.checks[id][d] = 1;
    }
    for (const h of out.habits) if (!out.checks[h.id]) out.checks[h.id] = {};
    for (const [wk, text] of Object.entries(s.notes || {})) if (isKey(wk) && text) out.notes[weekStart(wk)] = String(text).slice(0, 4000);
    out.settings.lang = s.settings && (s.settings.lang === 'de' || s.settings.lang === 'en') ? s.settings.lang : '';
    const meta = s.meta || {};
    out.meta.firstUse = isKey(meta.firstUse) ? meta.firstUse : today;
    out.meta.useDays = Array.isArray(meta.useDays) ? meta.useDays.filter(isKey).slice(-30) : [];
    out.meta.emailOfferHidden = !!meta.emailOfferHidden;
    return out;
  }
  function isoWeekMonday(y, w) {
    const jan4 = keyOf(y, 1, 4);
    return addDays(weekStart(jan4), (w - 1) * 7);
  }

  /** Record that the app was used today (distinct days, last 30 kept). Returns true if it is a new day. */
  function markUse(state, today) {
    const u = state.meta.useDays;
    if (u[u.length - 1] === today || u.includes(today)) return false;
    u.push(today); if (u.length > 30) u.splice(0, u.length - 30);
    return true;
  }

  // ---------- checks ----------
  function isDone(state, id, k) { return !!(state.checks[id] && state.checks[id][k]); }
  function setDone(state, id, k, on) {
    const c = state.checks[id] || (state.checks[id] = {});
    if (on) c[k] = 1; else delete c[k];
  }
  function toggle(state, id, k, today) {
    if (today && k > today) return false; // no ticking the future
    const on = !isDone(state, id, k);
    setDone(state, id, k, on);
    return on;
  }
  function doneDays(state, id) { return Object.keys(state.checks[id] || {}).sort(); }
  function countIn(state, id, a, b) { let n = 0; for (const d of Object.keys(state.checks[id] || {})) if (d >= a && d <= b) n++; return n; }
  function habitStart(state, h) { const d = doneDays(state, h.id); return d.length && d[0] < h.created ? d[0] : h.created; }

  // ---------- streaks ----------
  /** Daily: consecutive days. Today counts if done, but an open today does not break the streak.
      Weekly (n per week): consecutive weeks with at least n days; the running week counts once met. */
  function streaks(state, h, today) {
    return h.target.type === 'weekly' ? weeklyStreaks(state, h, today) : dailyStreaks(state, h, today);
  }
  function dailyStreaks(state, h, today) {
    const days = doneDays(state, h.id).filter(d => d <= today).map(dayNum);
    let best = 0, run = 0, prev = null;
    for (const n of days) { run = prev !== null && n === prev + 1 ? run + 1 : 1; prev = n; if (run > best) best = run; }
    const set = new Set(days);
    let cur = 0, n = dayNum(today);
    if (!set.has(n)) n--;
    while (set.has(n)) { cur++; n--; }
    return { current: cur, best, unit: 'days' };
  }
  function weeklyStreaks(state, h, today) {
    const need = h.target.n;
    const counts = new Map();
    for (const d of doneDays(state, h.id)) { if (d > today) continue; const w = weekStart(d); counts.set(w, (counts.get(w) || 0) + 1); }
    const met = w => (counts.get(w) || 0) >= need;
    const weeks = [...counts.keys()].filter(met).sort().map(dayNum);
    let best = 0, run = 0, prev = null;
    for (const n of weeks) { run = prev !== null && n === prev + 7 ? run + 1 : 1; prev = n; if (run > best) best = run; }
    let cur = 0, w = weekStart(today);
    if (!met(w)) w = addDays(w, -7);
    while (met(w)) { cur++; w = addDays(w, -7); }
    return { current: cur, best, unit: 'weeks' };
  }

  // ---------- completion ----------
  /** Share of the target reached between a and b (inclusive), only counting days up to today and since the habit started.
      Returns { done, expected, pct } or null when the period has not started for this habit. */
  function completion(state, h, a, b, today) {
    const start = habitStart(state, h);
    const from = a > start ? a : start;
    const to = b < today ? b : today;
    if (from > to) return null;
    const done = countIn(state, h.id, from, to);
    const days = diffDays(from, to) + 1;
    let expected;
    if (h.target.type === 'weekly') {
      // one calendar week: n days are the goal (so a running week reads as progress, 3/5);
      // longer periods (a month): n per 7 elapsed days, pro rata
      const fullWeek = a === weekStart(a) && diffDays(a, b) === 6;
      expected = fullWeek ? h.target.n : h.target.n * days / 7;
    } else expected = days;
    const pct = expected > 0 ? Math.min(1, done / expected) : 0;
    return { done, expected, pct };
  }
  function weekProgress(state, h, today) {
    const ws = weekStart(today);
    const done = countIn(state, h.id, ws, addDays(ws, 6));
    return { done, need: h.target.type === 'weekly' ? h.target.n : 7 };
  }
  function overall(state, a, b, today) {
    const vals = active(state).map(h => completion(state, h, a, b, today)).filter(Boolean);
    if (!vals.length) return null;
    return vals.reduce((s, v) => s + v.pct, 0) / vals.length;
  }
  function active(state) { return state.habits.filter(h => !h.archived).sort((a, b) => a.order - b.order); }

  /** 12-week heatmap: array of weeks (oldest first), each 7 cells { key, done, future }. */
  function heatmap(state, h, today, weeks) {
    weeks = weeks || 12;
    const first = addDays(weekStart(today), -7 * (weeks - 1));
    const out = [];
    for (let w = 0; w < weeks; w++) {
      const col = [];
      for (let d = 0; d < 7; d++) { const k = addDays(first, w * 7 + d); col.push({ key: k, done: isDone(state, h.id, k), future: k > today }); }
      out.push(col);
    }
    return out;
  }

  /** Last 7 days of a week (Mon..Sun) for the review: per habit days + count. */
  function weekReview(state, ws, today) {
    const days = range(ws, addDays(ws, 6));
    return {
      days,
      rows: active(state).map(h => ({ habit: h, cells: days.map(k => ({ key: k, done: isDone(state, h.id, k), future: k > today })),
        count: days.filter(k => isDone(state, h.id, k)).length,
        need: h.target.type === 'weekly' ? h.target.n : 7 })),
      overall: overall(state, ws, addDays(ws, 6), today)
    };
  }

  // ---------- suggestions / demo ----------
  const SUGGESTIONS = {
    en: [{ name: 'Drink 2 l of water', icon: '💧', target: { type: 'daily' } },
         { name: 'Walk 30 minutes', icon: '🚶', target: { type: 'weekly', n: 5 } },
         { name: 'Read 10 pages', icon: '📖', target: { type: 'daily' } }],
    de: [{ name: '2 l Wasser trinken', icon: '💧', target: { type: 'daily' } },
         { name: '30 Minuten gehen', icon: '🚶', target: { type: 'weekly', n: 5 } },
         { name: '10 Seiten lesen', icon: '📖', target: { type: 'daily' } }]
  };
  function addHabit(state, h, today) {
    const nh = normHabit({ ...h, id: uid(), order: state.habits.length, created: today }, state.habits.length, today);
    state.habits.push(nh); state.checks[nh.id] = {};
    return nh;
  }
  function move(state, id, dir) {
    const list = state.habits.slice().sort((a, b) => a.order - b.order);
    const i = list.findIndex(h => h.id === id), j = i + dir;
    if (i < 0 || j < 0 || j >= list.length) return false;
    [list[i], list[j]] = [list[j], list[i]];
    list.forEach((h, n) => { h.order = n; });
    return true;
  }
  function removeHabit(state, id) {
    state.habits = state.habits.filter(h => h.id !== id).map((h, i) => ({ ...h, order: i }));
    delete state.checks[id];
  }

  /** Deterministic demo data relative to today (for the screenshot and ?demo=1). */
  function demoState(today, lang) {
    const s = emptyState(addDays(today, -95));
    const names = lang === 'de'
      ? [['Laufen gehen', '🏃', { type: 'weekly', n: 3 }], ['10 Seiten lesen', '📖', { type: 'daily' }], ['Kein Handy nach 22 Uhr', '📵', { type: 'daily' }], ['2 l Wasser trinken', '💧', { type: 'daily' }], ['Deutsch lernen', '🗣', { type: 'weekly', n: 5 }]]
      : [['Go for a run', '🏃', { type: 'weekly', n: 3 }], ['Read 10 pages', '📖', { type: 'daily' }], ['No phone after 10 pm', '📵', { type: 'daily' }], ['Drink 2 l of water', '💧', { type: 'daily' }], ['Practise German', '🗣', { type: 'weekly', n: 5 }]];
    let seed = 7;
    const rnd = () => { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; };
    const rate = [0.5, 0.86, 0.7, 0.8, 0.72];
    names.forEach(([name, icon, target], i) => {
      const h = addHabit(s, { name, icon, target }, addDays(today, -90));
      for (let n = -90; n <= 0; n++) {
        const k = addDays(today, n);
        let p = rate[i];
        if (target.type === 'weekly' && i === 0) p = [1, 3, 5].includes(weekday(k)) ? 0.92 : 0.06;
        if (n > -13 && i === 1) p = 1; // a good current streak
        if (n === 0) p = i % 2 ? 1 : 0;
        if (rnd() < p) setDone(s, h.id, k, true);
      }
    });
    const lw = addDays(weekStart(today), -7);
    s.notes[lw] = lang === 'de' ? 'Zwei späte Abende im Büro, danach kein Lesen mehr. Buch auf den Nachttisch legen.' : 'Two late evenings at work, no reading after that. Put the book on the nightstand.';
    s.meta.useDays = range(addDays(today, -9), today);
    s.meta.firstUse = addDays(today, -95);
    return s;
  }

  return { SCHEMA, todayKey, isKey, keyOf, parts, dayNum, fromDayNum, addDays, diffDays, weekday, weekStart, daysInMonth, monthStart, monthEnd,
    range, isoWeek, isoWeekMonday, excelSerial, emptyState, migrate, markUse, isDone, setDone, toggle, doneDays, countIn, habitStart,
    streaks, completion, weekProgress, overall, active, heatmap, weekReview, SUGGESTIONS, addHabit, move, removeHabit, demoState, normTarget };
});
