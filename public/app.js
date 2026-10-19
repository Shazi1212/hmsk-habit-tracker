/* Habit Tracker UI. State lives in localStorage (key below), versioned via core.migrate.
   ?demo=1 shows generated demo data in memory only and never touches the stored data. ?lang=de|en picks the language. */
(function () {
  'use strict';
  const C = window.HTCore, X = window.HTXlsx, I18N = window.HTI18N, CFG = window.HTCONFIG || { site: '', template: '/assets/downloads/habit-tracker.xlsx' };
  const KEY = 'hmsk-habit-tracker';
  const params = new URLSearchParams(location.search);
  const DEMO = params.get('demo') === '1';
  const $ = s => document.querySelector(s);
  const main = $('#main');

  let today = C.todayKey();
  let state, canSave = true;
  let view = 'today';
  let monthCursor = C.monthStart(today);
  // the weekly review looks back: last week until Friday, the running week on the weekend
  const defaultWeek = k => C.weekday(k) >= 6 ? C.weekStart(k) : C.addDays(C.weekStart(k), -7);
  let weekCursor = defaultWeek(today);
  let editing = null; // habit id being edited in the habits view
  let lang = 'en';

  // ---------- storage ----------
  function load() {
    if (DEMO) { canSave = false; return C.demoState(today, pickLang(null)); }
    let raw = null;
    try { raw = localStorage.getItem(KEY); } catch (e) { canSave = false; }
    if (raw) {
      try { return C.migrate(JSON.parse(raw), today); }
      catch (e) {
        // keep the unreadable copy instead of overwriting it
        try { localStorage.setItem(KEY + '-broken-' + Date.now(), raw); } catch (e2) { /* ignore */ }
      }
    }
    return C.emptyState(today);
  }
  function save() {
    if (!canSave) return;
    try { localStorage.setItem(KEY, JSON.stringify(state)); }
    catch (e) { canSave = false; toast(t('saved_err', { e: e.message })); showBanner(); }
  }

  // ---------- i18n ----------
  function pickLang(st) {
    const q = params.get('lang');
    if (q === 'de' || q === 'en') return q;
    if (st && st.settings && st.settings.lang) return st.settings.lang;
    return /^de\b/i.test(navigator.language || '') ? 'de' : 'en';
  }
  function t(k, vars) {
    let s = (I18N[lang] && I18N[lang][k] !== undefined) ? I18N[lang][k] : I18N.en[k];
    if (vars && typeof s === 'string') for (const [a, b] of Object.entries(vars)) s = s.split('{' + a + '}').join(b);
    return s;
  }
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function fmtDate(k, withYear) {
    const [y, m, d] = C.parts(k);
    const wd = t('wd')[C.weekday(k) - 1], mon = t('months')[m - 1].slice(0, 3);
    return lang === 'de' ? `${wd} ${d}. ${mon}${withYear ? ' ' + y : ''}` : `${wd} ${d} ${mon}${withYear ? ' ' + y : ''}`;
  }
  function shortDate(k) { const [, m, d] = C.parts(k); return lang === 'de' ? `${d}.${m}.` : `${d} ${t('months')[m - 1].slice(0, 3)}`; }
  const pct = v => v === null || v === undefined ? '-' : Math.round(v * 100) + '%';
  function targetLabel(h) { return h.target.type === 'weekly' ? t('perweek', { n: h.target.n }) : t('daily'); }
  function unitLabel(n, unit) { return unit === 'weeks' ? (n === 1 ? t('week1') : t('weeks')) : (n === 1 ? t('day1') : t('days')); }
  const hname = h => (h.icon ? h.icon + ' ' : '') + h.name;
  const hnameHtml = h => (h.icon ? `<span class="emo" aria-hidden="true">${esc(h.icon)}</span>` : '') + esc(h.name);

  function applyChrome() {
    document.documentElement.lang = lang;
    const site = CFG.site || '';
    $('#lnkHome').href = site + (lang === 'de' ? '/de/tools/habit-tracker/' : '/tools/habit-tracker/');
    $('#lnkHome').textContent = t('back');
    $('#lnkImprint').href = site + (lang === 'de' ? '/impressum/' : '/imprint/');
    $('#lnkImprint').textContent = t('imprint');
    $('#lnkPrivacy').href = site + (lang === 'de' ? '/datenschutz/' : '/privacy/');
    $('#lnkPrivacy').textContent = t('privacy');
    const lb = $('#langBtn');
    lb.textContent = t('lang_switch'); lb.setAttribute('aria-label', t('lang_switch_label')); lb.lang = lang === 'de' ? 'en' : 'de';
    $('#skipLink').textContent = lang === 'de' ? 'Zum Inhalt' : 'Skip to content';
    document.querySelectorAll('#tabs button').forEach(b => {
      b.textContent = t('tab_' + b.dataset.view);
      if (b.dataset.view === view) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current');
    });
    document.title = t('app') + ' · hmsk.life';
    showBanner();
  }
  function showBanner() {
    const b = $('#banner');
    if (DEMO && params.get('shot') === '1') b.hidden = true; // clean screenshots for the tool page
    else if (DEMO) {
      const u = new URL(location.href); u.searchParams.delete('demo');
      b.innerHTML = `<span>${esc(t('demo_note'))}</span> <a href="${esc(u.pathname + u.search)}">${esc(t('demo_exit'))}</a>`;
      b.hidden = false;
    } else if (!canSave) { b.textContent = t('storage_off'); b.hidden = false; }
    else b.hidden = true;
  }
  let toastTimer;
  function toast(msg) { const el = $('#toast'); el.textContent = msg; el.classList.add('on'); clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('on'), 2600); }

  // ---------- pieces ----------
  function heat(h) {
    const weeks = C.heatmap(state, h, today, 12);
    const cells = weeks.map(col => col.map(c => `<i class="${c.future ? 'f' : c.done ? 'on' : ''}"></i>`).join('')).join('');
    const n = weeks.flat().filter(c => c.done).length;
    return `<div class="heat" role="img" aria-label="${esc(t('heat_label', { name: h.name }))}: ${n}">${cells}</div>`;
  }
  function emailOffer(force) {
    const used = state.meta.useDays.length >= 7;
    if (DEMO && !force) return '';
    if (!force && (!used || state.meta.emailOfferHidden)) return '';
    const href = (CFG.site || '') + (lang === 'de' ? '/de/?topic=habit-tracker#contact' : '/?topic=habit-tracker#contact');
    return `<aside class="offer"><p><span class="ps">&gt;</span> ${esc(t('email_offer'))} <a href="${esc(href)}">${esc(t('email_link'))}</a></p>
      <p class="dim small">${esc(t('email_honest'))}</p>${force ? '' : `<button type="button" class="linkish small" data-act="hide-offer">${esc(t('email_hide'))}</button>`}</aside>`;
  }

  // ---------- views ----------
  function renderToday() {
    const list = C.active(state);
    if (!list.length) return renderEmpty();
    const ws = C.weekStart(today);
    const wk = C.overall(state, ws, C.addDays(ws, 6), today);
    const mo = C.overall(state, C.monthStart(today), C.monthEnd(today), today);
    const rows = list.map((h, i) => {
      const on = C.isDone(state, h.id, today);
      const st = C.streaks(state, h, today);
      const wp = C.weekProgress(state, h, today);
      const prog = h.target.type === 'weekly' ? `<span class="tag${wp.done >= wp.need ? ' tag--ok' : ''}">${esc(t('week_prog', { d: wp.done, n: wp.need }))}</span>` : '';
      return `<li class="habit${on ? ' is-on' : ''}">
        <button type="button" class="tick" data-act="tick" data-id="${h.id}" aria-pressed="${on}" aria-label="${esc(t('mark_today', { name: h.name }))}">
          <span class="tick__box" aria-hidden="true">${on ? '<svg viewBox="0 0 16 16" class="ic"><path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/></svg>' : ''}</span>
          <span class="tick__txt"><span class="tick__name">${hnameHtml(h)}</span>
          <span class="tick__meta">${i < 9 ? `<kbd aria-hidden="true">${i + 1}</kbd>` : ''}${esc(targetLabel(h))}</span></span>
        </button>
        <div class="habit__stats">
          <span class="streak"><b>${st.current}</b> ${esc(unitLabel(st.current, st.unit))} <span class="dim">${esc(t('streak'))}</span></span>
          <span class="dim">${esc(t('best'))} ${st.best}</span>${prog}
        </div>
        ${heat(h)}
      </li>`;
    }).join('');
    return `<section class="pad">
      <div class="head"><h1 class="h1">${esc(fmtDate(today, false))}</h1>
        <p class="kpis"><span><b>${pct(wk)}</b> ${esc(t('week_pct'))}</span><span><b>${pct(mo)}</b> ${esc(t('month_pct'))}</span></p></div>
      <ul class="habits">${rows}</ul>
      <p class="dim small hint-keys">${esc(t('keys_hint'))}</p>
      ${emailOffer(false)}
    </section>`;
  }
  function renderEmpty() {
    const sug = C.SUGGESTIONS[lang] || C.SUGGESTIONS.en;
    return `<section class="pad empty">
      <p class="cmd"><span class="ps">$</span> ${esc(t('today_title').replace(/^\$ /, ''))}</p>
      <h1 class="h1">${esc(t('empty_title'))}</h1>
      <p class="lead">${esc(t('empty_text'))}</p>
      <ul class="sugs">${sug.map((s, i) => `<li><button type="button" class="sug" data-act="suggest" data-i="${i}">
        <span class="emo" aria-hidden="true">${esc(s.icon)}</span><span><b>${esc(s.name)}</b><span class="dim small"> · ${esc(s.target.type === 'weekly' ? t('perweek', { n: s.target.n }) : t('daily'))}</span></span>
        <span class="sug__add">+ ${esc(t('empty_add'))}</span></button></li>`).join('')}</ul>
      <p><button type="button" class="btn btn--ghost" data-act="goto" data-view="habits">${esc(t('empty_own'))}</button></p>
    </section>`;
  }

  function renderMonth() {
    const list = C.active(state);
    const [y, m] = C.parts(monthCursor);
    const days = C.range(monthCursor, C.monthEnd(monthCursor));
    const title = `${t('months')[m - 1]} ${y}`;
    const nav = `<div class="nav-row"><h1 class="h1 h1--sm" aria-live="polite">${esc(title)}</h1><span class="nav-btns">
      <button type="button" class="btn btn--ghost btn--sm" data-act="month" data-d="-1" aria-label="${esc(t('prev'))}">&lt;</button>
      <button type="button" class="btn btn--ghost btn--sm" data-act="month" data-d="1" aria-label="${esc(t('next'))}" ${C.monthStart(today) <= monthCursor ? 'disabled' : ''}>&gt;</button>
      ${monthCursor !== C.monthStart(today) ? `<button type="button" class="btn btn--ghost btn--sm" data-act="month" data-d="0">${esc(t('today_btn'))}</button>` : ''}</span></div>`;
    if (!list.length) return `<section class="pad">${nav}<p class="lead">${esc(t('no_active'))}</p></section>`;
    const head = days.map(k => { const d = C.parts(k)[2], wd = C.weekday(k);
      return `<th scope="col" class="${wd > 5 ? 'we' : ''}${k === today ? ' now' : ''}"><span>${esc(t('wd')[wd - 1].slice(0, 2))}</span>${d}</th>`; }).join('');
    const rows = list.map(h => {
      const c = C.completion(state, h, monthCursor, C.monthEnd(monthCursor), today);
      const cells = days.map(k => {
        const on = C.isDone(state, h.id, k), fut = k > today;
        return `<td class="${C.weekday(k) > 5 ? 'we' : ''}${k === today ? ' now' : ''}"><button type="button" class="cell${on ? ' on' : ''}" data-act="cell" data-id="${h.id}" data-k="${k}" aria-pressed="${on}" ${fut ? 'disabled' : ''} aria-label="${esc(t('cell_label', { name: h.name, date: fmtDate(k, true) }))}"></button></td>`;
      }).join('');
      return `<tr><th scope="row" class="rh"><span class="rh__in" title="${esc(h.name)}">${hnameHtml(h)}</span></th>${cells}<td class="pc">${pct(c && c.pct)}</td></tr>`;
    }).join('');
    const foot = days.map(k => {
      if (k > today) return '<td></td>';
      const act = list.filter(h => C.habitStart(state, h) <= k);
      if (!act.length) return '<td></td>';
      const n = act.filter(h => C.isDone(state, h.id, k)).length;
      return `<td class="dp" title="${n}/${act.length}">${Math.round(n / act.length * 100)}</td>`;
    }).join('');
    const mo = C.overall(state, monthCursor, C.monthEnd(monthCursor), today);
    return `<section class="pad">${nav}
      <p class="dim small">${esc(t('month_hint'))}</p>
      <div class="grid-wrap" tabindex="0" aria-label="${esc(title)}"><table class="grid">
        <thead><tr><th scope="col" class="rh"></th>${head}<th scope="col" class="pc">%</th></tr></thead>
        <tbody>${rows}</tbody>
        <tfoot><tr><th scope="row" class="rh dim">%</th>${foot}<td class="pc"><b>${pct(mo)}</b></td></tr></tfoot>
      </table></div>
    </section>`;
  }

  function renderReview() {
    const r = C.weekReview(state, weekCursor, today);
    const iw = C.isoWeek(weekCursor);
    const end = C.addDays(weekCursor, 6);
    const nav = `<div class="nav-row"><h1 class="h1 h1--sm" aria-live="polite">${esc(t('review_week', { w: iw.week, from: shortDate(weekCursor), to: shortDate(end) }))}</h1><span class="nav-btns">
      <button type="button" class="btn btn--ghost btn--sm" data-act="week" data-d="-1" aria-label="${esc(t('prev'))}">&lt;</button>
      <button type="button" class="btn btn--ghost btn--sm" data-act="week" data-d="1" aria-label="${esc(t('next'))}" ${C.weekStart(today) <= weekCursor ? 'disabled' : ''}>&gt;</button>
      ${weekCursor !== C.weekStart(today) ? `<button type="button" class="btn btn--ghost btn--sm" data-act="week" data-d="0">${esc(t('today_btn'))}</button>` : ''}</span></div>`;
    const any = Object.values(state.checks).some(c => Object.keys(c).length);
    const head = r.days.map(k => `<th scope="col"${k === today ? ' class="now"' : ''}>${esc(t('wd')[C.weekday(k) - 1])}</th>`).join('');
    const rows = r.rows.map(row => `<tr><th scope="row" class="rh"><span class="rh__in" title="${esc(row.habit.name)}">${hnameHtml(row.habit)}</span></th>${row.cells.map(c =>
      `<td><span class="dot${c.done ? ' on' : c.future ? ' f' : ''}" role="img" aria-label="${esc(fmtDate(c.key))}: ${c.done ? '✓' : '·'}"></span></td>`).join('')}
      <td class="pc${row.count >= row.need ? ' ok' : ''}">${row.count}/${row.need}</td></tr>`).join('');
    const note = state.notes[weekCursor] || '';
    return `<section class="pad">${nav}
      ${any && r.rows.length ? `<p class="kpis"><span><b>${pct(r.overall)}</b> ${esc(t('review_overall'))}</span></p>
      <div class="grid-wrap"><table class="grid grid--week"><thead><tr><th class="rh"></th>${head}<th class="pc"></th></tr></thead><tbody>${rows}</tbody></table></div>`
      : `<p class="lead">${esc(t('review_none'))}</p>`}
      <div class="field"><label class="flag" for="note">${esc(t('review_note'))}</label>
        <textarea id="note" rows="4" maxlength="4000" placeholder="${esc(t('review_note_ph'))}" data-week="${weekCursor}">${esc(note)}</textarea>
        <span class="dim small" id="noteState" aria-live="polite"></span></div>
      ${emailOffer(true)}
    </section>`;
  }

  function habitForm(h) {
    const tg = h ? h.target : { type: 'daily' };
    const id = h ? h.id : 'new';
    return `<form class="hform" data-form="${id}" novalidate>
      <div class="hform__row">
        <div class="field field--icon"><label class="flag" for="ic-${id}">${esc(t('h_icon'))}</label><input id="ic-${id}" name="icon" maxlength="4" value="${esc(h ? h.icon : '')}" placeholder="${esc(t('h_icon_ph'))}" autocomplete="off"></div>
        <div class="field field--grow"><label class="flag" for="nm-${id}">${esc(t('h_name'))}</label><input id="nm-${id}" name="name" maxlength="60" value="${esc(h ? h.name : '')}" placeholder="${esc(t('h_name_ph'))}" autocomplete="off" required></div>
      </div>
      <div class="hform__row">
        <div class="field"><label class="flag" for="tg-${id}">${esc(t('h_target'))}</label>
          <select id="tg-${id}" name="target">
            <option value="7"${tg.type === 'daily' ? ' selected' : ''}>${esc(t('t_daily'))}</option>
            ${[6, 5, 4, 3, 2, 1].map(n => `<option value="${n}"${tg.type === 'weekly' && tg.n === n ? ' selected' : ''}>${esc(t('perweek', { n }))}</option>`).join('')}
          </select></div>
        <div class="hform__btns">${h ? `<button type="submit" class="btn btn--amber">${esc(t('save'))}</button><button type="button" class="btn btn--ghost" data-act="cancel-edit">${esc(t('cancel'))}</button>`
          : `<button type="submit" class="btn btn--amber">${esc(t('add'))}</button>`}</div>
      </div>
      <p class="form-err small" aria-live="polite"></p>
    </form>`;
  }
  function renderHabits() {
    const list = C.active(state);
    const arch = state.habits.filter(h => h.archived);
    const row = (h, i, n) => editing === h.id ? `<li class="hrow hrow--edit">${habitForm(h)}</li>` : `<li class="hrow">
      <span class="hrow__name">${hnameHtml(h)} <span class="dim small">· ${esc(targetLabel(h))}</span></span>
      <span class="hrow__btns">
        <button type="button" class="ibtn" data-act="up" data-id="${h.id}" aria-label="${esc(t('up'))}: ${esc(h.name)}" ${i === 0 ? 'disabled' : ''}><svg viewBox="0 0 16 16" class="ic" aria-hidden="true"><path d="M7.646 4.646a.5.5 0 0 1 .708 0l6 6a.5.5 0 0 1-.708.708L8 5.707l-5.646 5.647a.5.5 0 0 1-.708-.708l6-6z"/></svg></button>
        <button type="button" class="ibtn" data-act="down" data-id="${h.id}" aria-label="${esc(t('down'))}: ${esc(h.name)}" ${i === n - 1 ? 'disabled' : ''}><svg viewBox="0 0 16 16" class="ic" aria-hidden="true"><path d="M1.646 4.646a.5.5 0 0 1 .708 0L8 10.293l5.646-5.647a.5.5 0 0 1 .708.708l-6 6a.5.5 0 0 1-.708 0l-6-6a.5.5 0 0 1 0-.708z"/></svg></button>
        <button type="button" class="btn btn--ghost btn--sm" data-act="edit" data-id="${h.id}">${esc(t('edit'))}</button>
        <button type="button" class="btn btn--ghost btn--sm" data-act="archive" data-id="${h.id}">${esc(t('archive'))}</button>
      </span></li>`;
    return `<section class="pad">
      <p class="cmd"><span class="ps">$</span> ${esc(t('habits_title').replace(/^\$ /, ''))}</p>
      ${habitForm(null)}
      <ul class="hlist">${list.map((h, i) => row(h, i, list.length)).join('')}</ul>
      ${arch.length ? `<h2 class="h2">${esc(t('archived_head'))}</h2><ul class="hlist">${arch.map(h => `<li class="hrow"><span class="hrow__name dim">${hnameHtml(h)}</span>
        <span class="hrow__btns"><button type="button" class="btn btn--ghost btn--sm" data-act="unarchive" data-id="${h.id}">${esc(t('unarchive'))}</button>
        <button type="button" class="btn btn--ghost btn--sm btn--danger" data-act="delete" data-id="${h.id}">${esc(t('del'))}</button></span></li>`).join('')}</ul>` : ''}
    </section>`;
  }

  function renderData() {
    return `<section class="pad">
      <p class="cmd"><span class="ps">$</span> ${esc(t('data_title').replace(/^\$ /, ''))}</p>
      <p class="lead">${esc(t('data_local'))}</p>
      <div class="btn-col">
        <button type="button" class="btn btn--amber" data-act="xlsx">${esc(t('exp_xlsx'))}</button>
        <button type="button" class="btn btn--ghost" data-act="csv">${esc(t('exp_csv'))}</button>
        <button type="button" class="btn btn--ghost" data-act="json">${esc(t('exp_json'))}</button>
        <label class="btn btn--ghost file-btn">${esc(t('imp_json'))}<input type="file" accept="application/json,.json" data-act="import" class="sr"></label>
      </div>
      <p class="small">${esc(t('template'))} <a href="${esc(CFG.template)}" download>${esc(t('template_link'))}</a></p>
      <p><button type="button" class="linkish small danger" data-act="clear">${esc(t('clear'))}</button></p>
    </section>`;
  }

  function render(focusSel) {
    applyChrome();
    const fn = { today: renderToday, month: renderMonth, review: renderReview, habits: renderHabits, data: renderData }[view] || renderToday;
    main.innerHTML = fn();
    if (focusSel) { const el = main.querySelector(focusSel); if (el) el.focus(); }
    scrollToToday();
  }
  /** On a phone the month grid is wider than the screen: show the days around today. */
  function scrollToToday() {
    const wrap = main.querySelector('.grid-wrap'), now = wrap && wrap.querySelector('thead th.now');
    if (!now) return;
    const rh = wrap.querySelector('.rh');
    wrap.scrollLeft = Math.max(0, now.offsetLeft - (rh ? rh.offsetWidth : 0) - (wrap.clientWidth - (rh ? rh.offsetWidth : 0)) / 2);
  }
  /** Re-render and put the focus back on the same control (keyboard users keep their place). */
  function rerenderKeepFocus(el) {
    const a = el && el.dataset;
    let sel = null;
    if (a && a.act) sel = `[data-act="${a.act}"]` + (a.id ? `[data-id="${a.id}"]` : '') + (a.k ? `[data-k="${a.k}"]` : '') + (a.d ? `[data-d="${a.d}"]` : '');
    const wrap = main.querySelector('.grid-wrap'); const sx = wrap ? wrap.scrollLeft : 0;
    render();
    const w2 = main.querySelector('.grid-wrap'); if (w2) w2.scrollLeft = sx;
    if (sel) {
      let n = main.querySelector(sel);
      if ((!n || n.disabled) && a.id) n = [...main.querySelectorAll(`[data-id="${a.id}"]`)].find(x => !x.disabled); // e.g. "up" became disabled at the top
      if (n && !n.disabled) n.focus({ preventScroll: true });
    }
  }

  // ---------- downloads ----------
  function download(name, data, type) {
    const blob = data instanceof Blob ? data : new Blob([data], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    toast(t('exported'));
  }
  function csv() {
    const sep = lang === 'de' ? ';' : ',';
    const q = v => { const s = String(v); return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const habits = state.habits.slice().sort((a, b) => (a.archived - b.archived) || (a.order - b.order));
    let start = today; for (const h of habits) { const s = C.habitStart(state, h); if (s < start) start = s; }
    const lines = [[t('x_date'), ...habits.map(h => h.name)].map(q).join(sep)];
    for (const k of C.range(start, today)) lines.push([k, ...habits.map(h => C.isDone(state, h.id, k) ? 1 : 0)].join(sep));
    return '﻿' + lines.join('\r\n') + '\r\n';
  }

  // ---------- events ----------
  function changed(el) { save(); rerenderKeepFocus(el); }
  document.getElementById('tabs').addEventListener('click', e => {
    const b = e.target.closest('button[data-view]'); if (!b) return;
    view = b.dataset.view; editing = null; render(); main.focus({ preventScroll: true });
  });
  $('#langBtn').addEventListener('click', () => {
    lang = lang === 'de' ? 'en' : 'de';
    if (!DEMO) { state.settings.lang = lang; save(); }
    const u = new URL(location.href); if (u.searchParams.has('lang')) { u.searchParams.set('lang', lang); history.replaceState(null, '', u); }
    render();
  });
  main.addEventListener('click', e => {
    const el = e.target.closest('[data-act]'); if (!el || el.tagName === 'INPUT') return;
    const a = el.dataset.act, id = el.dataset.id;
    const h = id && state.habits.find(x => x.id === id);
    switch (a) {
      case 'tick': C.toggle(state, id, today, today); changed(el); break;
      case 'cell': C.toggle(state, id, el.dataset.k, today); changed(el); break;
      case 'suggest': { const s = (C.SUGGESTIONS[lang] || C.SUGGESTIONS.en)[Number(el.dataset.i)]; C.addHabit(state, s, today); save(); render('[data-act="tick"]'); break; }
      case 'goto': view = el.dataset.view; render('input[name="name"]'); break;
      case 'month': { const d = Number(el.dataset.d); const [y, m] = C.parts(monthCursor);
        monthCursor = d === 0 ? C.monthStart(today) : C.keyOf(m + d < 1 ? y - 1 : m + d > 12 ? y + 1 : y, ((m - 1 + d + 12) % 12) + 1, 1);
        if (monthCursor > C.monthStart(today)) monthCursor = C.monthStart(today); rerenderKeepFocus(el); break; }
      case 'week': { const d = Number(el.dataset.d); weekCursor = d === 0 ? C.weekStart(today) : C.addDays(weekCursor, 7 * d);
        if (weekCursor > C.weekStart(today)) weekCursor = C.weekStart(today); rerenderKeepFocus(el); break; }
      case 'up': case 'down': C.move(state, id, a === 'up' ? -1 : 1); changed(el); break;
      case 'edit': editing = id; render(`[data-form="${id}"] input[name="name"]`); break;
      case 'cancel-edit': editing = null; render(); break;
      case 'archive': if (h) { h.archived = true; editing = null; changed(); } break;
      case 'unarchive': if (h) { h.archived = false; h.order = C.active(state).length; C.active(state).forEach((x, i) => { x.order = i; }); changed(); } break;
      case 'delete': if (h && confirm(t('confirm_del', { name: h.name }))) { C.removeHabit(state, id); changed(); } break;
      case 'hide-offer': state.meta.emailOfferHidden = true; changed(); break;
      case 'xlsx': download('habit-tracker-' + today + '.xlsx', new Blob([X.fromState(state, C, today, t)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); break;
      case 'csv': download('habit-tracker-' + today + '.csv', csv(), 'text/csv;charset=utf-8'); break;
      case 'json': download('habit-tracker-backup-' + today + '.json', JSON.stringify(state, null, 1), 'application/json'); break;
      case 'clear': if (confirm(t('confirm_clear'))) { state = C.emptyState(today); state.settings.lang = lang; C.markUse(state, today); save(); view = 'today'; render(); } break;
    }
  });
  main.addEventListener('submit', e => {
    const f = e.target.closest('form[data-form]'); if (!f) return;
    e.preventDefault();
    const fld = n => f.querySelector(`[name="${n}"]`);
    const name = fld('name').value.trim();
    if (!name) { f.querySelector('.form-err').textContent = t('name_required'); fld('name').focus(); return; }
    const n = Number(fld('target').value);
    const target = n >= 7 ? { type: 'daily' } : { type: 'weekly', n };
    const icon = fld('icon').value.trim();
    if (f.dataset.form === 'new') { C.addHabit(state, { name, icon, target }, today); save(); render('#nm-new'); toast('+ ' + name); }
    else { const h = state.habits.find(x => x.id === f.dataset.form); if (h) { h.name = name.slice(0, 60); h.icon = icon.slice(0, 4); h.target = C.normTarget(target); } editing = null; save(); render(`[data-act="edit"][data-id="${f.dataset.form}"]`); }
  });
  let noteTimer;
  main.addEventListener('input', e => {
    if (e.target.id !== 'note') return;
    const wk = e.target.dataset.week, v = e.target.value;
    clearTimeout(noteTimer);
    noteTimer = setTimeout(() => { if (v.trim()) state.notes[wk] = v; else delete state.notes[wk]; save(); const s = $('#noteState'); if (s) s.textContent = canSave ? t('review_saved') : ''; }, 400);
  });
  main.addEventListener('change', e => {
    if (e.target.dataset.act !== 'import') return;
    const file = e.target.files && e.target.files[0]; if (!file) return;
    const r = new FileReader();
    r.onload = () => {
      try {
        const next = C.migrate(JSON.parse(r.result), today);
        if (!confirm(t('confirm_import', { n: next.habits.length }))) return;
        state = next; C.markUse(state, today); save(); lang = pickLang(state); toast(t('imported', { n: next.habits.length })); view = 'today'; render();
      } catch (err) { toast(t('import_err', { e: err.message })); }
    };
    r.readAsText(file);
    e.target.value = '';
  });
  document.addEventListener('keydown', e => {
    if (view !== 'today' || e.ctrlKey || e.metaKey || e.altKey) return;
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select' || e.target.isContentEditable) return;
    if (!/^[1-9]$/.test(e.key)) return;
    const h = C.active(state)[Number(e.key) - 1]; if (!h) return;
    e.preventDefault();
    C.toggle(state, h.id, today, today); save(); render();
    const b = main.querySelector(`[data-act="tick"][data-id="${h.id}"]`); if (b) b.focus({ preventScroll: true });
  });
  // a tab left open past midnight moves on to the new day
  function checkDay() {
    const k = C.todayKey();
    if (k === today) return;
    if (monthCursor === C.monthStart(today)) monthCursor = C.monthStart(k);
    if (weekCursor === defaultWeek(today)) weekCursor = defaultWeek(k);
    today = k;
    if (!DEMO && C.markUse(state, today)) save();
    render();
  }
  document.addEventListener('visibilitychange', () => { if (!document.hidden) checkDay(); });
  setInterval(checkDay, 60000);
  window.addEventListener('storage', e => { if (e.key === KEY && e.newValue && !DEMO) { try { state = C.migrate(JSON.parse(e.newValue), today); render(); } catch (err) { /* ignore */ } } });

  // ---------- start ----------
  state = load();
  lang = pickLang(DEMO ? null : state);
  if (!DEMO && C.markUse(state, today)) save();
  const v0 = params.get('view'); if (['today', 'month', 'review', 'habits', 'data'].includes(v0)) view = v0;
  render();
})();
