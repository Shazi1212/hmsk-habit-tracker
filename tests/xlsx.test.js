process.env.TZ = process.env.TZ || 'Europe/Berlin';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const C = require('../public/core.js');
const X = require('../public/xlsx.js');

test('crc32 matches known vectors', () => {
  const enc = s => new TextEncoder().encode(s);
  assert.equal(X.crc32(enc('')), 0);
  assert.equal(X.crc32(enc('123456789')), 0xCBF43926);
  assert.equal(X.crc32(enc('The quick brown fox jumps over the lazy dog')), 0x414FA339);
});

test('zip has valid local headers, central directory and end record', () => {
  const z = X.zip([{ name: 'a.txt', data: 'hello' }, { name: 'dir/ü.xml', data: '<x/>' }], new Date(2026, 9, 25, 2, 30, 10));
  const dv = new DataView(z.buffer);
  assert.equal(dv.getUint32(0, true), 0x04034b50);
  const eocd = z.length - 22;
  assert.equal(dv.getUint32(eocd, true), 0x06054b50);
  assert.equal(dv.getUint16(eocd + 10, true), 2);
  const cdOff = dv.getUint32(eocd + 16, true);
  assert.equal(dv.getUint32(cdOff, true), 0x02014b50);
  assert.equal(dv.getUint32(cdOff + 16, true), X.crc32(new TextEncoder().encode('hello')));
});

test('column names', () => {
  assert.equal(X.colName(0), 'A'); assert.equal(X.colName(25), 'Z'); assert.equal(X.colName(26), 'AA'); assert.equal(X.colName(701), 'ZZ'); assert.equal(X.colName(702), 'AAA');
});

function sampleState(today) {
  const s = C.demoState(today, 'en');
  s.habits[2].archived = true;
  s.habits[0].name = 'Run & <stretch> "daily"';
  return s;
}

test('exported workbook opens with python openpyxl and holds the real data', (t) => {
  const py = spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' });
  if (py.status !== 0) { t.skip('python3 with openpyxl not installed'); return; }
  const today = '2026-10-21';
  const s = sampleState(today);
  const bytes = X.fromState(s, C, today, k => k);
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'ht-')), 'export.xlsx');
  fs.writeFileSync(file, bytes);
  const script = `
import openpyxl, json, sys
wb = openpyxl.load_workbook(sys.argv[1])
out = {"sheets": wb.sheetnames}
log = wb[wb.sheetnames[0]]
out["head"] = [c.value for c in log[1]]
out["first"] = [str(log.cell(2, 1).value)] + [log.cell(2, c).value for c in range(2, log.max_column + 1)]
out["rows"] = log.max_row
out["freeze"] = log.freeze_panes
hb = wb[wb.sheetnames[1]]
out["habits"] = [[c.value for c in r] for r in hb.iter_rows(min_row=2, max_col=5)]
nt = wb[wb.sheetnames[2]]
out["note"] = nt.cell(2, 2).value
print(json.dumps(out, default=str))
`;
  const r = spawnSync('python3', ['-c', script, file], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
  const o = JSON.parse(r.stdout);
  assert.deepEqual(o.sheets, ['x_sheet_log', 'x_sheet_habits', 'x_sheet_notes']);
  assert.equal(o.freeze, 'B2');
  // habits sorted: active by order, archived last
  const active = s.habits.filter(h => !h.archived), arch = s.habits.filter(h => h.archived);
  const order = [...active, ...arch];
  assert.deepEqual(o.head.slice(1, -1), order.map(h => h.icon + ' ' + h.name));
  // first data row is today, ticks match the state
  assert.ok(o.first[0].startsWith('2026-10-21'));
  order.forEach((h, i) => assert.equal(o.first[i + 1], C.isDone(s, h.id, today) ? 1 : null, h.name));
  const start = order.map(h => C.habitStart(s, h)).sort()[0];
  assert.equal(o.rows, C.diffDays(start, today) + 2);
  // streaks in the summary
  order.forEach((h, i) => { const st = C.streaks(s, h, today); assert.equal(o.habits[i][2], st.current); assert.equal(o.habits[i][3], st.best); });
  assert.match(o.note, /nightstand/);
});

test('export with the real German UI texts opens with openpyxl (sheet names, umlauts, emoji)', (t) => {
  const py = spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' });
  if (py.status !== 0) { t.skip('python3 with openpyxl not installed'); return; }
  const vm = require('node:vm');
  const ctx = { window: {} };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../public/i18n.js'), 'utf8'), ctx);
  const de = ctx.window.HTI18N.de;
  const tr = (k) => de[k];
  const today = '2026-10-25'; // DST change day in Europe
  const s = C.demoState(today, 'de');
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'ht-')), 'de.xlsx');
  fs.writeFileSync(file, X.fromState(s, C, today, tr));
  const r = spawnSync('python3', ['-c', `
import openpyxl, sys, json
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb.worksheets[0]
print(json.dumps({"names": wb.sheetnames, "a2": str(ws["A2"].value), "b1": ws["B1"].value, "n": ws.max_row}))
`, file], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
  const o = JSON.parse(r.stdout);
  assert.deepEqual(o.names, ['Protokoll', 'Gewohnheiten', 'Wochennotizen']);
  assert.ok(o.a2.startsWith('2026-10-25'), o.a2);
  assert.equal(o.b1, '🏃 Laufen gehen');
});
