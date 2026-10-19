// Run: node --test tests/
process.env.TZ = 'Europe/Berlin'; // DST: 2026-03-29 (23 h day) and 2026-10-25 (25 h day)
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../public/core.js');

const T = '2026-10-21'; // a Wednesday

function withHabit(target, days, created = '2026-01-01') {
  const s = C.emptyState(T);
  const h = C.addHabit(s, { name: 'x', target }, created);
  for (const d of days) C.setDone(s, h.id, d, true);
  return { s, h };
}

test('date keys: local calendar day, no UTC shift', () => {
  assert.equal(C.todayKey(new Date(2026, 9, 25, 0, 30)), '2026-10-25'); // 00:30 local, 22:30 UTC the day before
  assert.equal(C.todayKey(new Date(2026, 2, 29, 23, 59)), '2026-03-29');
  assert.equal(C.todayKey(new Date(2026, 0, 1, 0, 0)), '2026-01-01');
});

test('day arithmetic across DST change days', () => {
  assert.equal(C.addDays('2026-03-28', 1), '2026-03-29');
  assert.equal(C.addDays('2026-03-29', 1), '2026-03-30');
  assert.equal(C.addDays('2026-10-24', 1), '2026-10-25');
  assert.equal(C.addDays('2026-10-25', 1), '2026-10-26');
  assert.equal(C.diffDays('2026-03-01', '2026-04-01'), 31);
  assert.equal(C.diffDays('2026-10-01', '2026-11-01'), 31);
  assert.deepEqual(C.range('2026-10-24', '2026-10-27'), ['2026-10-24', '2026-10-25', '2026-10-26', '2026-10-27']);
  assert.equal(C.range('2026-01-01', '2026-12-31').length, 365);
  assert.equal(C.range('2028-01-01', '2028-12-31').length, 366);
});

test('weekday, week start, ISO week, month helpers', () => {
  assert.equal(C.weekday('2026-10-19'), 1);
  assert.equal(C.weekday('2026-10-25'), 7);
  assert.equal(C.weekStart('2026-10-25'), '2026-10-19');
  assert.equal(C.weekStart('2026-03-29'), '2026-03-23');
  assert.deepEqual(C.isoWeek('2026-01-01'), { year: 2026, week: 1 });
  assert.deepEqual(C.isoWeek('2027-01-01'), { year: 2026, week: 53 });
  assert.equal(C.isoWeekMonday(2026, 43), '2026-10-19');
  assert.equal(C.daysInMonth(2026, 2), 28);
  assert.equal(C.daysInMonth(2028, 2), 29);
  assert.equal(C.monthEnd('2026-02-10'), '2026-02-28');
  assert.ok(C.isKey('2026-02-28') && !C.isKey('2026-02-30') && !C.isKey('2026-13-01') && !C.isKey('x'));
  assert.equal(C.excelSerial('1900-03-01'), 61);
  assert.equal(C.excelSerial('2026-01-01'), 46023);
});

test('daily streak: today open does not break it, gaps do', () => {
  const { s, h } = withHabit({ type: 'daily' }, ['2026-10-18', '2026-10-19', '2026-10-20']);
  assert.deepEqual(C.streaks(s, h, T), { current: 3, best: 3, unit: 'days' });
  C.setDone(s, h.id, T, true);
  assert.equal(C.streaks(s, h, T).current, 4);
  C.setDone(s, h.id, '2026-10-19', false);
  assert.deepEqual(C.streaks(s, h, T), { current: 2, best: 2, unit: 'days' });
  assert.equal(C.streaks(s, h, '2026-10-23').current, 0); // two days missed
});

test('daily streak runs through DST change days', () => {
  const { s, h } = withHabit({ type: 'daily' }, ['2026-03-28', '2026-03-29', '2026-03-30', '2026-10-24', '2026-10-25', '2026-10-26']);
  assert.equal(C.streaks(s, h, '2026-03-30').current, 3);
  assert.equal(C.streaks(s, h, '2026-10-26').current, 3);
  assert.equal(C.streaks(s, h, '2026-10-26').best, 3);
});

test('best streak ignores future ticks', () => {
  const { s, h } = withHabit({ type: 'daily' }, ['2026-10-20', '2026-10-22', '2026-10-23']);
  assert.equal(C.streaks(s, h, T).best, 1);
});

test('weekly target: counts weeks met, running week counts once met', () => {
  // need 3/week. Weeks: 05 Oct (3 days, met), 12 Oct (3 days, met), 19 Oct (2 so far)
  const days = ['2026-10-05', '2026-10-07', '2026-10-09', '2026-10-12', '2026-10-13', '2026-10-18', '2026-10-19', '2026-10-20'];
  const { s, h } = withHabit({ type: 'weekly', n: 3 }, days);
  assert.deepEqual(C.streaks(s, h, T), { current: 2, best: 2, unit: 'weeks' });
  C.setDone(s, h.id, T, true);
  assert.equal(C.streaks(s, h, T).current, 3);
  // week of 28 Sep not met: streak starts 05 Oct
  C.setDone(s, h.id, '2026-09-28', true);
  assert.equal(C.streaks(s, h, T).best, 3);
  // a missed full week breaks the streak
  assert.equal(C.streaks(s, h, '2026-11-04').current, 0);
});

test('weekly target across the DST week (Oct 19 to 25)', () => {
  const { s, h } = withHabit({ type: 'weekly', n: 2 }, ['2026-10-24', '2026-10-25', '2026-10-26', '2026-10-27']);
  assert.equal(C.streaks(s, h, '2026-10-27').current, 2);
});

test('completion: daily counts elapsed days only, from habit start', () => {
  const { s, h } = withHabit({ type: 'daily' }, ['2026-10-19', '2026-10-20'], '2026-10-01');
  const w = C.completion(s, h, '2026-10-19', '2026-10-25', T); // Mon..Wed elapsed
  assert.equal(w.done, 2); assert.equal(w.expected, 3); assert.ok(Math.abs(w.pct - 2 / 3) < 1e-9);
  const m = C.completion(s, h, '2026-10-01', '2026-10-31', T);
  assert.equal(m.expected, 21);
  assert.equal(C.completion(s, h, '2026-09-01', '2026-09-30', T), null); // before the habit existed
});

test('completion: weekly target is n per week, capped at 100%', () => {
  const { s, h } = withHabit({ type: 'weekly', n: 3 }, ['2026-10-19', '2026-10-20', '2026-10-21', '2026-10-15'], '2026-10-01');
  const w = C.completion(s, h, '2026-10-19', '2026-10-25', T);
  assert.equal(w.expected, 3); assert.equal(w.pct, 1);
  const p = C.weekProgress(s, h, T);
  assert.deepEqual(p, { done: 3, need: 3 });
  const m = C.completion(s, h, '2026-10-01', '2026-10-31', T);
  assert.ok(Math.abs(m.expected - 9) < 1e-9); // 21 days * 3/7
});

test('toggle refuses future days', () => {
  const { s, h } = withHabit({ type: 'daily' }, []);
  assert.equal(C.toggle(s, h.id, '2026-10-22', T), false);
  assert.equal(C.isDone(s, h.id, '2026-10-22'), false);
  assert.equal(C.toggle(s, h.id, T, T), true);
  assert.equal(C.toggle(s, h.id, T, T), false);
});

test('heatmap is 12 weeks x 7 days ending with the current week', () => {
  const { s, h } = withHabit({ type: 'daily' }, ['2026-10-20']);
  const hm = C.heatmap(s, h, T, 12);
  assert.equal(hm.length, 12); assert.ok(hm.every(c => c.length === 7));
  assert.equal(hm[11][0].key, '2026-10-19');
  assert.equal(hm[0][0].key, '2026-08-03');
  assert.equal(hm[11][1].done, true);
  assert.equal(hm[11][3].future, true);
});

test('reorder, archive and remove', () => {
  const s = C.emptyState(T);
  const a = C.addHabit(s, { name: 'a' }, T), b = C.addHabit(s, { name: 'b' }, T), c = C.addHabit(s, { name: 'c' }, T);
  assert.ok(C.move(s, c.id, -1));
  assert.deepEqual(C.active(s).map(h => h.name), ['a', 'c', 'b']);
  assert.equal(C.move(s, a.id, -1), false);
  b.archived = true;
  assert.deepEqual(C.active(s).map(h => h.name), ['a', 'c']);
  C.removeHabit(s, a.id);
  assert.equal(s.checks[a.id], undefined);
  assert.equal(s.habits.length, 2);
});

test('migration from v1 prototype format', () => {
  const v1 = { version: 1, habits: [
    { id: 'a1', name: 'Run', emoji: '🏃', perWeek: 3, done: ['2026-10-19', '2026-10-20', 'garbage', '2026-02-30'] },
    { id: 'b2', name: 'Read', perWeek: 0, done: ['2026-10-21'] }],
    notes: { '2026-W43': 'late nights', '2026-W99x': 'bad' } };
  const s = C.migrate(v1, T);
  assert.equal(s.v, C.SCHEMA);
  assert.deepEqual(s.habits[0].target, { type: 'weekly', n: 3 });
  assert.equal(s.habits[0].icon, '🏃');
  assert.deepEqual(s.habits[1].target, { type: 'daily' });
  assert.deepEqual(Object.keys(s.checks.a1).sort(), ['2026-10-19', '2026-10-20']);
  assert.deepEqual(s.notes, { '2026-10-19': 'late nights' });
  assert.ok(Array.isArray(s.meta.useDays));
});

test('migration keeps v2 data and cleans bad values', () => {
  const s0 = C.emptyState(T);
  const h = C.addHabit(s0, { name: '  Water  ', target: { type: 'weekly', n: 9 } }, T);
  C.setDone(s0, h.id, T, true);
  s0.checks.ghost = { '2026-10-01': 1 };
  s0.notes['2026-10-21'] = 'note on a Wednesday';
  const s = C.migrate(JSON.parse(JSON.stringify(s0)), T);
  assert.equal(s.habits[0].name, 'Water');
  assert.deepEqual(s.habits[0].target, { type: 'daily' }); // 9/week is not a weekly target
  assert.equal(s.checks.ghost, undefined);
  assert.ok(C.isDone(s, h.id, T));
  assert.deepEqual(s.notes, { '2026-10-19': 'note on a Wednesday' });
});

test('migration rejects garbage and newer versions', () => {
  assert.throws(() => C.migrate(null, T));
  assert.throws(() => C.migrate({ foo: 1 }, T));
  assert.throws(() => C.migrate({ v: 99, habits: [] }, T), /newer/);
});

test('use days are distinct and capped', () => {
  const s = C.emptyState(T);
  assert.equal(C.markUse(s, T), true);
  assert.equal(C.markUse(s, T), false);
  for (let i = 1; i <= 40; i++) C.markUse(s, C.addDays(T, i));
  assert.equal(s.meta.useDays.length, 30);
});

test('demo data is generated without touching anything else', () => {
  const s = C.demoState(T, 'en');
  assert.equal(s.habits.length, 5);
  assert.ok(C.streaks(s, s.habits[1], T).current >= 12);
  assert.ok(Object.keys(s.notes).length === 1);
});
