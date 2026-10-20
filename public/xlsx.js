/* Minimal .xlsx writer, no library: SpreadsheetML parts packed in a ZIP (method "stored", CRC-32 computed here).
   Enough for values, dates, percentages, a bold header, column widths and frozen panes.
   Browser: window.HTXlsx. Node: require('./xlsx.js'). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HTXlsx = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------- CRC-32 (IEEE 802.3, as used by ZIP) ----------
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  const enc = s => new TextEncoder().encode(s);

  // ---------- ZIP (stored, no compression) ----------
  /** files: [{ name, data: string | Uint8Array }], date: Date for the entry timestamps. Returns Uint8Array. */
  function zip(files, date) {
    date = date || new Date();
    const dosTime = (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1);
    const dosDate = ((Math.max(1980, date.getFullYear()) - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
    const chunks = []; const central = []; let offset = 0;
    for (const f of files) {
      const name = enc(f.name);
      const data = typeof f.data === 'string' ? enc(f.data) : f.data;
      const crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      chunks.push(new Uint8Array(lh.buffer), name, data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true);
      ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
      ch.setUint16(30, 0, true); ch.setUint16(32, 0, true); ch.setUint16(34, 0, true); ch.setUint16(36, 0, true); ch.setUint32(38, 0, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(4, 0, true); end.setUint16(6, 0, true);
    end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
    const all = [...chunks, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((s, c) => s + c.length, 0));
    let p = 0; for (const c of all) { out.set(c, p); p += c.length; }
    return out;
  }

  // ---------- SpreadsheetML ----------
  const xesc = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function colName(i) { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  // style ids (see STYLES): 0 normal, 1 header, 2 date, 3 percent, 4 tick, 5 wrap text
  const S = { normal: 0, head: 1, date: 2, pct: 3, tick: 4, wrap: 5 };
  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd ddd"/></numFmts>' +
    '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FF151B24"/><name val="Calibri"/></font></fonts>' +
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1E4C8A"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFA8C8F2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  /** cell: null | string | number | { v, s } (s = style id). Strings are written inline (no shared strings part). */
  function cellXml(ref, c) {
    if (c === null || c === undefined || c === '') return '';
    const o = typeof c === 'object' ? c : { v: c };
    if (o.v === null || o.v === undefined || o.v === '') return o.s ? `<c r="${ref}" s="${o.s}"/>` : '';
    const s = o.s ? ` s="${o.s}"` : '';
    if (typeof o.v === 'number' && Number.isFinite(o.v)) return `<c r="${ref}"${s}><v>${o.v}</v></c>`;
    return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xesc(o.v)}</t></is></c>`;
  }
  function sheetXml(sh) {
    const rows = sh.rows || [];
    const width = rows.reduce((m, r) => Math.max(m, r.length), 1);
    let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    x += `<dimension ref="A1:${colName(width - 1)}${Math.max(1, rows.length)}"/>`;
    x += '<sheetViews><sheetView workbookViewId="0"' + (sh.first ? ' tabSelected="1"' : '') + '>';
    if (sh.freeze) {
      const m = /^([A-Z]+)(\d+)$/.exec(sh.freeze);
      const col = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, row = Number(m[2]) - 1;
      const pane = col && row ? 'bottomRight' : row ? 'bottomLeft' : 'topRight';
      x += `<pane${col ? ` xSplit="${col}"` : ''}${row ? ` ySplit="${row}"` : ''} topLeftCell="${sh.freeze}" activePane="${pane}" state="frozen"/><selection pane="${pane}" activeCell="${sh.freeze}" sqref="${sh.freeze}"/>`;
    }
    x += '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>';
    if (sh.cols && sh.cols.length) x += '<cols>' + sh.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
    x += '<sheetData>';
    rows.forEach((r, ri) => {
      const cells = r.map((c, ci) => cellXml(colName(ci) + (ri + 1), c)).join('');
      x += `<row r="${ri + 1}">${cells}</row>`;
    });
    x += '</sheetData>';
    if (sh.autoFilter && rows.length > 1) x += `<autoFilter ref="A1:${colName(width - 1)}${rows.length}"/>`;
    x += '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>';
    x += '<pageSetup paperSize="9" orientation="landscape"/></worksheet>';
    return x;
  }
  function safeSheetName(n, used) {
    let s = String(n).replace(/[\\/?*[\]:]/g, ' ').replace(/^'+|'+$/g, '').trim().slice(0, 31) || 'Sheet';
    let base = s, i = 2;
    while (used.has(s.toLowerCase())) { const suf = ' ' + i++; s = base.slice(0, 31 - suf.length) + suf; }
    used.add(s.toLowerCase());
    return s;
  }
  /** sheets: [{ name, rows, cols, freeze, autoFilter }]. Returns Uint8Array (.xlsx). */
  function workbook(sheets, opts) {
    opts = opts || {};
    const used = new Set();
    const names = sheets.map(s => safeSheetName(s.name, used));
    const files = [];
    files.push({ name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>' });
    files.push({ name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>' });
    const iso = (opts.date || new Date()).toISOString().replace(/\.\d+Z$/, 'Z');
    files.push({ name: 'docProps/core.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${xesc(opts.title || 'Habit Tracker')}</dc:title><dc:creator>Habit Tracker (hmsk.life)</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified></cp:coreProperties>` });
    files.push({ name: 'docProps/app.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Habit Tracker</Application></Properties>' });
    files.push({ name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<bookViews><workbookView activeTab="0"/></bookViews><sheets>' + names.map((n, i) => `<sheet name="${xesc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>' });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` });
    files.push({ name: 'xl/styles.xml', data: STYLES });
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml({ ...s, first: i === 0 }) }));
    return zip(files, opts.date);
  }

  /** The user's data as a workbook: Log (one row per day), Habits (streaks, totals), Weekly notes. */
  function fromState(state, core, today, t) {
    t = t || (k => k);
    const habits = state.habits.slice().sort((a, b) => (a.archived - b.archived) || (a.order - b.order));
    let start = today;
    for (const h of habits) { const s = core.habitStart(state, h); if (s < start) start = s; }
    if (core.diffDays(start, today) > 3650) start = core.addDays(today, -3650);
    const days = core.range(start, today).reverse(); // newest first, like the app
    const head = [{ v: t('x_date'), s: S.head }, ...habits.map(h => ({ v: (h.icon ? h.icon + ' ' : '') + h.name, s: S.head })), { v: t('x_done'), s: S.head }];
    const log = [head];
    for (const k of days) {
      let n = 0;
      const cells = habits.map(h => { const d = core.isDone(state, h.id, k); if (d) n++; return d ? { v: 1, s: S.tick } : null; });
      log.push([{ v: core.excelSerial(k), s: S.date }, ...cells, n]);
    }
    const ms = core.monthStart(today), me = core.monthEnd(today);
    const ws = core.weekStart(today);
    const summary = [[t('x_habit'), t('x_target'), t('x_cur'), t('x_best'), t('x_unit'), t('x_week'), t('x_month'), t('x_total'), t('x_since'), t('x_archived')].map(v => ({ v, s: S.head }))];
    for (const h of habits) {
      const st = core.streaks(state, h, today);
      const w = core.completion(state, h, ws, core.addDays(ws, 6), today);
      const m = core.completion(state, h, ms, me, today);
      summary.push([(h.icon ? h.icon + ' ' : '') + h.name,
        h.target.type === 'weekly' ? t('x_perweek').replace('{n}', h.target.n) : t('x_daily'),
        st.current, st.best, t(st.unit === 'weeks' ? 'x_weeks' : 'x_days'),
        w ? { v: Math.round(w.pct * 100) / 100, s: S.pct } : null,
        m ? { v: Math.round(m.pct * 100) / 100, s: S.pct } : null,
        core.doneDays(state, h.id).length,
        { v: core.excelSerial(core.habitStart(state, h)), s: S.date },
        h.archived ? t('x_yes') : '']);
    }
    const notes = [[{ v: t('x_weekof'), s: S.head }, { v: t('x_note'), s: S.head }]];
    for (const k of Object.keys(state.notes).sort().reverse()) notes.push([{ v: core.excelSerial(k), s: S.date }, { v: state.notes[k], s: S.wrap }]);
    return workbook([
      { name: t('x_sheet_log'), rows: log, cols: [16, ...habits.map(h => Math.min(28, Math.max(10, h.name.length + 4))), 8], freeze: 'B2', autoFilter: true },
      { name: t('x_sheet_habits'), rows: summary, cols: [30, 14, 14, 12, 8, 11, 12, 10, 16, 10], freeze: 'B2' },
      { name: t('x_sheet_notes'), rows: notes, cols: [16, 90], freeze: 'A2' }
    ], { title: 'Habit Tracker' });
  }

  return { crc32, zip, workbook, fromState, colName, sheetXml, S };
});
