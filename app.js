'use strict';
/* S-21 Списки — разделы, колонки, люди, поиск, выгрузка в Excel. Работает офлайн.
   Формат данных такой же, как у s21_data.json на компьютере: можно переносить файлом туда и обратно. */

const LS_KEY = 's21_data_v1', LS_UI = 's21_ui_v1', LS_PREV = 's21_data_prev_v1';
const DEFAULT_COLUMNS = [{key:'group',label:'Группа'},{key:'status',label:'Статус'},
  {key:'birth',label:'Дата рождения'},{key:'baptism',label:'Дата крещения'}];
const RESERVED = new Set(['name', 'reports']);
const $app = document.getElementById('app');

let db, ui, state = {screen: 'home'}, draft = null, installEvt = null;

/* ───────── утилиты ───────── */
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const nameKey = (s) => String(s || '').replace(/ё/g, 'е').replace(/Ё/g, 'Е').replace(/\s+/g, ' ').trim().toLowerCase();
const clone = (o) => JSON.parse(JSON.stringify(o));
const today = () => { const d = new Date(), p = (n) => String(n).padStart(2, '0'); return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`; };

function toast(text, ms = 2200) {
  const t = document.getElementById('toast');
  t.textContent = text; t.classList.add('show');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), ms);
}

function slugify(label, used) {
  const base = String(label).trim().toLowerCase().replace(/[^0-9a-zа-яё]+/g, '_').replace(/^_+|_+$/g, '') || 'col';
  let cand = base, n = 2;
  while (used.has(cand) || RESERVED.has(cand)) cand = `${base}_${n++}`;
  used.add(cand);
  return cand;
}

/* ───────── окна ───────── */
function dialog({title, msg, input, ok = 'OK', cancel = 'Отмена', danger = false, single = false}) {
  return new Promise((resolve) => {
    const root = document.getElementById('modal-root');
    root.innerHTML = `<div class="dialog-back"><div class="dialog" role="dialog">
      ${title ? `<h3>${esc(title)}</h3>` : ''}
      ${msg ? `<div class="msg">${esc(msg)}</div>` : ''}
      ${input !== undefined ? `<input id="dlg-input" value="${esc(input)}" autocomplete="off">` : ''}
      <div class="btns">${single ? '' : `<button id="dlg-cancel">${esc(cancel)}</button>`}
        <button id="dlg-ok" class="${danger ? 'danger' : 'accent'}">${esc(ok)}</button></div></div></div>`;
    const inp = document.getElementById('dlg-input');
    const done = (v) => { root.innerHTML = ''; resolve(v); };
    document.getElementById('dlg-ok').onclick = () => done(inp ? inp.value : true);
    const c = document.getElementById('dlg-cancel'); if (c) c.onclick = () => done(inp ? null : false);
    root.querySelector('.dialog-back').onclick = (e) => { if (e.target.classList.contains('dialog-back') && !single) done(inp ? null : false); };
    if (inp) { inp.focus(); inp.select(); inp.onkeydown = (e) => { if (e.key === 'Enter') done(inp.value); }; }
  });
}
const askText = (title, initial = '') => dialog({title, input: initial});
const confirmBox = (title, msg, ok = 'OK', danger = false) => dialog({title, msg, ok, danger});
const alertBox = (title, msg) => dialog({title, msg, single: true});

function sheet(title, items) {
  const root = document.getElementById('modal-root');
  root.innerHTML = `<div class="sheet-back"><div class="sheet">${title ? `<h3>${esc(title)}</h3>` : ''}
    ${items.map((it, i) => `<button class="item" data-i="${i}"><span>${esc(it.icon || '')}</span>
      <span><b>${esc(it.text)}</b>${it.sub ? `<small>${esc(it.sub)}</small>` : ''}</span></button>`).join('')}</div></div>`;
  const close = () => { root.innerHTML = ''; };
  root.querySelector('.sheet-back').onclick = (e) => { if (e.target.classList.contains('sheet-back')) close(); };
  root.querySelectorAll('.item').forEach((b) => b.onclick = () => { close(); items[+b.dataset.i].fn(); });
}

/* ───────── данные ───────── */
function sortPeople(sec) { sec.people.sort((a, b) => nameKey(a.name).localeCompare(nameKey(b.name), 'ru')); }

function normalizeDb(d) {
  if (!d || typeof d !== 'object' || !Array.isArray(d.sections)) throw new Error('В файле нет списка разделов («sections»)');
  d.sections.forEach((s) => {
    s.name = String(s.name || 'Без названия');
    if (!Array.isArray(s.columns)) s.columns = clone(DEFAULT_COLUMNS);
    if (!Array.isArray(s.people)) s.people = [];
    s.columns = s.columns.map((c) => ({key: String(c.key), label: String(c.label ?? '')}));
    s.people = s.people.filter((p) => p && typeof p === 'object' && p.name != null);
    sortPeople(s);
  });
  if (!d.sections.length) d.sections.push({name: 'Общий список', columns: clone(DEFAULT_COLUMNS), people: []});
  return d;
}

function loadAll() {
  try { db = normalizeDb(JSON.parse(localStorage.getItem(LS_KEY))); }
  catch (e) { db = normalizeDb({sections: [{name: 'Общий список', columns: clone(DEFAULT_COLUMNS), people: []}]}); }
  try { ui = JSON.parse(localStorage.getItem(LS_UI)) || {}; } catch (e) { ui = {}; }
  ui.section = Math.min(Math.max(0, ui.section | 0), db.sections.length - 1);
  ui.view = ui.view === 'cards' ? 'cards' : 'table';
  ui.q = '';
}
function saveDb() {
  try { localStorage.setItem(LS_KEY, JSON.stringify(db)); return true; }
  catch (e) { alertBox('Не удалось сохранить', 'Память браузера переполнена или запрещена. Сделайте резервную копию (меню → Резервная копия) и освободите место.'); return false; }
}
function saveUi() { try { localStorage.setItem(LS_UI, JSON.stringify({section: ui.section, view: ui.view})); } catch (e) {} }
const curSec = () => db.sections[ui.section];
const countPeople = () => db.sections.reduce((n, s) => n + s.people.length, 0);

function knownNames() {
  const m = new Map();
  db.sections.forEach((s) => s.people.forEach((p) => {
    const k = nameKey(p.name); if (!k) return;
    if (!m.has(k)) m.set(k, {person: p, sections: []});
    m.get(k).sections.push(s.name);
  }));
  return m;
}
const uniqueSectionName = (base, skip) => {
  let n = base, i = 2;
  while (db.sections.some((s) => s !== skip && nameKey(s.name) === nameKey(n))) n = `${base} ${i++}`;
  return n;
};

/* ───────── Excel (.xlsx) без внешних библиотек ───────── */
const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = (u8) => { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };

function zipStore(files) {
  const enc = new TextEncoder(), parts = [], central = []; let offset = 0;
  const d = new Date(), dosT = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
        dosD = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  const u16 = (v) => [v & 255, (v >> 8) & 255], u32 = (v) => [v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255];
  for (const f of files) {
    const name = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
    const head = new Uint8Array([0x50,0x4b,3,4, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(dosT), ...u16(dosD),
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0)]);
    parts.push(head, name, data);
    central.push(new Uint8Array([0x50,0x4b,1,2, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(dosT), ...u16(dosD),
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(0), ...u32(offset)]), name);
    offset += head.length + name.length + data.length;
  }
  const cdSize = central.reduce((n, a) => n + a.length, 0);
  const end = new Uint8Array([0x50,0x4b,5,6, ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cdSize), ...u32(offset), ...u16(0)]);
  return new Blob([...parts, ...central, end], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

const colLetter = (n) => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
const xml = (s) => String(s).replace(/[&<>"]/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');

// style ids: 0 обычный, 1 заголовок, 2 название, 3 данные с рамкой, 4 итог
function sheetXml(rows, merges, widths) {
  let body = '';
  rows.forEach((row, r) => {
    const cells = row.cells.map((v, c) => {
      const ref = colLetter(c) + (r + 1), s = row.style;
      if (v === null || v === '' || v === undefined) return `<c r="${ref}" s="${s}"/>`;
      if (typeof v === 'number') return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
      return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xml(v)}</t></is></c>`;
    }).join('');
    body += `<row r="${r + 1}"${row.ht ? ` ht="${row.ht}" customHeight="1"` : ''}>${cells}</row>`;
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>
<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>
<sheetData>${body}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}</worksheet>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
<font><b/><sz val="13"/><color rgb="FF1F3864"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FF1F3864"/><name val="Arial"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFC0C8D8"/></left><right style="thin"><color rgb="FFC0C8D8"/></right><top style="thin"><color rgb="FFC0C8D8"/></top><bottom style="thin"><color rgb="FFC0C8D8"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf fontId="2" fillId="3" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf borderId="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf fontId="3" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

function buildXlsx(secIdxs) {
  const used = new Set(), sheets = [];
  secIdxs.forEach((si) => {
    const s = db.sections[si], cols = s.columns, n = 2 + cols.length;
    let base = s.name.replace(/[\\/*?:\[\]]/g, ' ').trim().slice(0, 31) || 'Лист', nm = base, k = 2;
    while (used.has(nm.toLowerCase())) nm = base.slice(0, 28) + ' ' + k++;
    used.add(nm.toLowerCase());
    const rows = [{style: 2, ht: 26, cells: [`${s.name}  •  ${today()}`, ...Array(n - 1).fill('')]},
      {style: 1, ht: 22, cells: ['№', 'ФИО', ...cols.map((c) => c.label)]}];
    s.people.forEach((p, i) => rows.push({style: 3, cells: [i + 1, p.name, ...cols.map((c) => p[c.key] ?? '')]}));
    rows.push({style: 4, cells: [`Всего: ${s.people.length}`, ...Array(n - 1).fill('')]});
    const widths = [5, Math.min(48, Math.max(14, ...s.people.map((p) => String(p.name).length)) + 2),
      ...cols.map((c) => Math.min(40, Math.max(10, String(c.label).length + 2, ...s.people.map((p) => String(p[c.key] ?? '').length + 2))))];
    sheets.push({name: nm, xml: sheetXml(rows, [`A1:${colLetter(n - 1)}1`, `A${rows.length}:${colLetter(Math.min(3, n - 1))}${rows.length}`], widths)});
  });
  const files = [
    {name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`},
    {name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`},
    {name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`},
    {name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`},
    {name: 'xl/styles.xml', data: STYLES},
    ...sheets.map((s, i) => ({name: `xl/worksheets/sheet${i + 1}.xml`, data: s.xml})),
  ];
  return zipStore(files);
}

// how: 'share' — меню «Отправить», 'save' — сразу в «Загрузки» (на телефоне, откуда файл можно взять вручную)
async function deliver(blob, filename, how = 'share') {
  if (how !== 'save') try {
    const file = new File([blob], filename, {type: blob.type});
    if (navigator.canShare && navigator.canShare({files: [file]})) { await navigator.share({files: [file]}); return 'shared'; }
  } catch (e) { if (e && e.name === 'AbortError') return 'cancelled'; }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000);
  return 'downloaded';
}

/* ───────── навигация ───────── */
function go(s) { history.pushState(s, ''); state = s; render(); }
function back() { history.back(); }
window.addEventListener('popstate', (e) => { state = e.state || {screen: 'home'}; draft = state.screen === 'home' ? null : draft; render(); });
function render() {
  document.getElementById('modal-root').innerHTML = '';
  ({home: renderHome, person: renderPerson, section: renderSection, export: renderExport, backup: renderBackup, order: renderOrder})[state.screen || 'home']();
}

/* ───────── главный экран ───────── */
function renderHome() {
  const sec = curSec();
  $app.innerHTML = `<div class="screen">
    <div class="topbar"><h1>${esc(sec.name)} <span class="sub">${sec.people.length}</span></h1>
      <button class="iconbtn ghost" id="b-view" title="Вид">${ui.view === 'table' ? '☰' : '▦'}</button>
      <button class="iconbtn ghost" id="b-menu" title="Меню">⋮</button></div>
    <div class="chips" id="chips">${db.sections.map((s, i) => `<button class="chip${i === ui.section ? ' on' : ''}" data-i="${i}">${esc(s.name)} · ${s.people.length}</button>`).join('')}
      <button class="chip" id="b-newsec">＋ Раздел</button></div>
    <div class="toolbar"><input id="q" type="search" placeholder="🔍 Поиск по всей базе" value="${esc(ui.q)}" autocomplete="off"></div>
    <div class="content" id="list"></div>
    <button class="fab" id="b-add">＋ Человек</button></div>`;
  renderList();
  document.getElementById('chips').onclick = (e) => {
    const b = e.target.closest('.chip'); if (!b || b.id === 'b-newsec') return;
    ui.section = +b.dataset.i; ui.q = ''; saveUi(); renderHome();
  };
  document.getElementById('b-newsec').onclick = () => openSection(null);
  document.getElementById('b-add').onclick = () => openPerson(ui.section, null);
  document.getElementById('b-view').onclick = () => { ui.view = ui.view === 'table' ? 'cards' : 'table'; saveUi(); renderHome(); };
  document.getElementById('b-menu').onclick = openMenu;
  document.getElementById('q').oninput = (e) => { ui.q = e.target.value; renderList(); };
  const on = $app.querySelector('.chip.on'); if (on) on.scrollIntoView({inline: 'center', block: 'nearest'});
}

function kvLine(sec, p, max = 6) {
  const parts = sec.columns.map((c) => [c.label, p[c.key]]).filter(([, v]) => v !== undefined && v !== null && String(v).trim() !== '');
  return parts.slice(0, max).map(([l, v]) => `${l ? esc(l) + ': ' : ''}<span>${esc(v)}</span>`).join(' · ');
}

function renderList() {
  const box = document.getElementById('list'); if (!box) return;
  const q = nameKey(ui.q);
  if (q) {
    const res = [];
    db.sections.forEach((s, si) => s.people.forEach((p, pi) => { if (nameKey(p.name).includes(q)) res.push({s, si, p, pi}); }));
    box.innerHTML = res.length ? `<div class="cards" style="padding-top:4px"><div class="muted" style="margin:0 2px 8px">Найдено: ${res.length}</div>${res.slice(0, 300).map((r) => `<div class="card" data-s="${r.si}" data-p="${r.pi}"><b>${esc(r.p.name)}<span class="tag">${esc(r.s.name)}</span></b><div class="kv">${kvLine(r.s, r.p)}</div></div>`).join('')}</div>`
      : '<div class="empty">Никого не найдено</div>';
    box.onclick = (e) => { const c = e.target.closest('.card'); if (c) openPerson(+c.dataset.s, +c.dataset.p); };
    return;
  }
  const sec = curSec();
  if (!sec.people.length) { box.innerHTML = `<div class="empty">В этом разделе пока никого нет.<br>Нажмите «＋ Человек».</div>`; box.onclick = null; return; }
  if (ui.view === 'cards') {
    box.innerHTML = `<div class="cards" style="padding-top:4px">${sec.people.map((p, i) => `<div class="card" data-s="${ui.section}" data-p="${i}"><b>${esc(p.name)}</b><div class="kv">${kvLine(sec, p) || '—'}</div></div>`).join('')}</div>`;
    box.onclick = (e) => { const c = e.target.closest('.card'); if (c) openPerson(+c.dataset.s, +c.dataset.p); };
  } else {
    box.innerHTML = `<div class="tblwrap"><table><thead><tr><th>ФИО</th>${sec.columns.map((c) => `<th>${esc(c.label)}</th>`).join('')}</tr></thead>
      <tbody>${sec.people.map((p, i) => `<tr data-p="${i}"${ui.hl === p ? ' class="hl"' : ''}><td>${esc(p.name)}</td>${sec.columns.map((c) => `<td>${esc(p[c.key] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
    box.onclick = (e) => { const r = e.target.closest('tr[data-p]'); if (r) openPerson(ui.section, +r.dataset.p); };
    const hl = box.querySelector('tr.hl'); if (hl) hl.scrollIntoView({block: 'center'});
  }
}

function openMenu() {
  const items = [
    {icon: '⬇', text: 'Выгрузить в Excel', sub: 'выбрать разделы и отправить файл', fn: () => go({screen: 'export'})},
    {icon: '⚙', text: 'Колонки и название раздела', sub: `«${curSec().name}»`, fn: () => openSection(ui.section)},
    {icon: '↕', text: 'Порядок разделов', fn: () => go({screen: 'order'})},
    {icon: '💾', text: 'Резервная копия и перенос', sub: 'сохранить / загрузить файл данных', fn: () => go({screen: 'backup'})},
  ];
  if (installEvt) items.push({icon: '📲', text: 'Установить приложение', sub: 'значок на экране телефона', fn: async () => { installEvt.prompt(); try { await installEvt.userChoice; } catch (e) {} installEvt = null; }});
  items.push({icon: 'ℹ', text: 'О приложении', fn: () => alertBox('S-21 Списки', `Людей: ${countPeople()}, разделов: ${db.sections.length}.\nДанные хранятся только в этом телефоне и работают без интернета.\nДелайте резервную копию (меню → Резервная копия): если очистить данные браузера, список пропадёт.`)});
  sheet('Меню', items);
}

/* ───────── человек ───────── */
function openPerson(si, pi) {
  const sec = db.sections[si], p = pi == null ? null : sec.people[pi];
  draft = {si, p, name: p ? p.name : '', vals: p ? {...p} : {}, again: false};
  go({screen: 'person', si, pi});
}

function renderPerson() {
  if (!draft) { state = {screen: 'home'}; return renderHome(); }
  const sec = db.sections[draft.si], isEdit = !!draft.p, names = [...new Set(db.sections.flatMap((s) => s.people.map((p) => p.name)))].sort();
  $app.innerHTML = `<div class="screen">
    <div class="topbar"><button class="iconbtn ghost" id="b-back">←</button><h1>${isEdit ? 'Редактировать' : 'Новый человек'}</h1>
      ${isEdit ? '<button class="iconbtn ghost" id="b-del" title="Удалить">🗑</button>' : ''}</div>
    <div class="form" id="form">
      <div class="field"><label>ФИО</label><input id="f-name" list="dl-names" autocomplete="off" value="${esc(draft.name)}">
        <datalist id="dl-names">${names.map((n) => `<option value="${esc(n)}">`).join('')}</datalist><div id="hint" class="hint"></div></div>
      <div class="field"><label>Раздел</label><select id="f-sec">${db.sections.map((s, i) => `<option value="${i}"${i === draft.si ? ' selected' : ''}>${esc(s.name)}</option>`).join('')}</select></div>
      ${sec.columns.map((c, i) => `<div class="field"><label${c.label ? '' : ' class="dim"'}>${esc(c.label || `(колонка ${i + 1})`)}</label><input data-k="${esc(c.key)}" autocomplete="off" value="${esc(draft.vals[c.key] ?? '')}"></div>`).join('')}
    </div>
    <div class="actionbar"><button id="b-cancel">Отмена</button>
      ${isEdit ? '<button class="primary" id="b-save">Сохранить</button>' : '<button id="b-save2">Сохранить и ещё</button><button class="primary" id="b-save">Сохранить</button>'}</div></div>`;
  const form = document.getElementById('form');
  const readDom = () => { draft.name = document.getElementById('f-name').value; form.querySelectorAll('[data-k]').forEach((i) => { draft.vals[i.dataset.k] = i.value; }); };
  document.getElementById('b-back').onclick = back;
  document.getElementById('b-cancel').onclick = back;
  document.getElementById('f-sec').onchange = (e) => { readDom(); draft.si = +e.target.value; renderPerson(); };
  const hint = () => {
    const h = document.getElementById('hint'), nm = nameKey(document.getElementById('f-name').value), km = knownNames().get(nm);
    h.className = 'hint'; h.textContent = '';
    if (!nm || !km) return;
    const here = db.sections[draft.si].people.some((x) => x !== draft.p && nameKey(x.name) === nm);
    if (here) { h.className = 'hint warn'; h.textContent = `⚠ Уже есть в разделе «${db.sections[draft.si].name}»`; return; }
    const others = [...new Set(km.sections)];
    h.textContent = `Уже есть в: ${others.join(', ')}` + (isEdit ? '' : ' — данные подставлены');
    if (!isEdit) form.querySelectorAll('[data-k]').forEach((i) => { if (!i.value && km.person[i.dataset.k]) i.value = km.person[i.dataset.k]; });
  };
  document.getElementById('f-name').addEventListener('change', hint);
  document.getElementById('f-name').addEventListener('blur', hint);
  const save = async (again) => {
    readDom();
    const name = draft.name.trim(); if (!name) { toast('Введите ФИО'); return; }
    const target = db.sections[draft.si];
    const dup = target.people.some((x) => x !== draft.p && nameKey(x.name) === nameKey(name));
    if (dup && !(await confirmBox('Дубликат', `«${name}» уже есть в разделе «${target.name}».\nВсё равно сохранить ещё раз?`, 'Сохранить'))) return;
    const person = draft.p || {};
    person.name = name;
    target.columns.forEach((c) => { person[c.key] = (draft.vals[c.key] ?? '').trim(); });
    if (draft.p) { db.sections.forEach((s) => { const k = s.people.indexOf(person); if (k >= 0 && s !== target) s.people.splice(k, 1); }); }
    if (!target.people.includes(person)) target.people.push(person);
    sortPeople(target);
    if (!saveDb()) return;
    ui.section = draft.si; ui.hl = person; ui.q = ''; saveUi();
    if (again) {
      draft = {si: draft.si, p: null, name: '', vals: {group: draft.vals.group ?? ''}, again: true};
      renderPerson(); document.getElementById('f-name').focus(); toast(`Сохранено: ${name}`);
    } else { toast(`Сохранено: ${name}`); back(); }
  };
  document.getElementById('b-save').onclick = () => save(false);
  const s2 = document.getElementById('b-save2'); if (s2) s2.onclick = () => save(true);
  const del = document.getElementById('b-del');
  if (del) del.onclick = async () => {
    if (!(await confirmBox('Удалить человека', `Удалить «${draft.p.name}» из раздела «${sec.name}»?\nОтменить нельзя.`, 'Удалить', true))) return;
    sec.people.splice(sec.people.indexOf(draft.p), 1); saveDb(); toast('Удалено'); back();
  };
}

/* ───────── раздел и его колонки ───────── */
function openSection(idx) {
  const s = idx == null ? null : db.sections[idx];
  draft = {idx, name: s ? s.name : '', cols: (s ? s.columns : DEFAULT_COLUMNS).map((c) => ({key: c.key, label: c.label}))};
  go({screen: 'section', idx});
}

function renderSection() {
  if (!draft) { state = {screen: 'home'}; return renderHome(); }
  const isNew = draft.idx == null;
  $app.innerHTML = `<div class="screen">
    <div class="topbar"><button class="iconbtn ghost" id="b-back">←</button><h1>${isNew ? 'Новый раздел' : 'Раздел и колонки'}</h1></div>
    <div class="form" id="form">
      <div class="field"><label>Название раздела</label><input id="s-name" value="${esc(draft.name)}" autocomplete="off"></div>
      <div class="sec-title">Колонки (кроме ФИО — оно всегда есть). Название можно оставить пустым.</div>
      <div id="cols"></div>
      <button id="b-addcol" style="margin-top:4px">＋ Колонка</button>
      ${isNew ? '' : `<div class="sec-title">Действия с разделом</div>
      <button id="b-copy" style="width:100%;margin-bottom:8px">⧉ Сделать копию раздела (с людьми)</button>
      <button id="b-delsec" class="danger" style="width:100%">🗑 Удалить раздел</button>`}
    </div>
    <div class="actionbar"><button id="b-cancel">Отмена</button><button class="primary" id="b-save">${isNew ? 'Создать' : 'Сохранить'}</button></div></div>`;
  const drawCols = () => {
    const box = document.getElementById('cols');
    box.innerHTML = draft.cols.map((c, i) => `<div class="colrow" data-i="${i}"><input value="${esc(c.label)}" placeholder="(без названия)" autocomplete="off">
      <button data-a="up"${i === 0 ? ' disabled' : ''}>▲</button><button data-a="dn"${i === draft.cols.length - 1 ? ' disabled' : ''}>▼</button><button data-a="rm" class="danger">✕</button></div>`).join('')
      || '<div class="muted" style="padding:6px 2px">Колонок нет — будет только ФИО.</div>';
  };
  const sync = () => { document.querySelectorAll('#cols .colrow').forEach((r) => { draft.cols[+r.dataset.i].label = r.querySelector('input').value; }); draft.name = document.getElementById('s-name').value; };
  drawCols();
  document.getElementById('cols').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return; sync();
    const i = +b.closest('.colrow').dataset.i, a = b.dataset.a;
    if (a === 'rm') draft.cols.splice(i, 1);
    if (a === 'up' && i > 0) [draft.cols[i - 1], draft.cols[i]] = [draft.cols[i], draft.cols[i - 1]];
    if (a === 'dn' && i < draft.cols.length - 1) [draft.cols[i + 1], draft.cols[i]] = [draft.cols[i], draft.cols[i + 1]];
    drawCols();
  };
  document.getElementById('b-addcol').onclick = () => {
    sync(); draft.cols.push({key: null, label: ''}); drawCols();
    const inputs = document.querySelectorAll('#cols input'), last = inputs[inputs.length - 1];
    last.scrollIntoView({block: 'center'}); last.focus();
  };
  document.getElementById('b-back').onclick = back;
  document.getElementById('b-cancel').onclick = back;
  document.getElementById('b-save').onclick = () => {
    sync();
    const name = draft.name.replace(/\s+/g, ' ').trim();
    if (!name) { toast('Введите название раздела'); return; }
    if (db.sections.some((s, i) => i !== draft.idx && nameKey(s.name) === nameKey(name))) { toast('Раздел с таким названием уже есть'); return; }
    const used = new Set(draft.cols.map((c) => c.key).filter(Boolean));
    const cols = draft.cols.map((c) => ({key: c.key || slugify(c.label, used), label: c.label.trim()}));
    if (isNew) { db.sections.push({name, columns: cols, people: []}); ui.section = db.sections.length - 1; }
    else { const s = db.sections[draft.idx]; s.name = name; s.columns = cols; ui.section = draft.idx; }
    if (!saveDb()) return; saveUi(); toast(isNew ? 'Раздел создан' : 'Сохранено'); back();
  };
  const cp = document.getElementById('b-copy');
  if (cp) cp.onclick = async () => {
    const src = db.sections[draft.idx];
    const nm = await askText('Название копии', uniqueSectionName(`${src.name} (копия)`));
    if (nm == null || !nm.trim()) return;
    if (db.sections.some((s) => nameKey(s.name) === nameKey(nm))) { toast('Раздел с таким названием уже есть'); return; }
    db.sections.splice(draft.idx + 1, 0, {name: nm.replace(/\s+/g, ' ').trim(), columns: clone(src.columns), people: clone(src.people)});
    ui.section = draft.idx + 1; saveDb(); saveUi(); toast(`Копия создана: ${src.people.length} чел.`); back();
  };
  const ds = document.getElementById('b-delsec');
  if (ds) ds.onclick = async () => {
    if (db.sections.length <= 1) { alertBox('Нельзя', 'Должен остаться хотя бы один раздел.'); return; }
    const s = db.sections[draft.idx];
    if (!(await confirmBox('Удалить раздел', `Удалить раздел «${s.name}» и всех людей в нём (${s.people.length})?\nОтменить нельзя.`, 'Удалить', true))) return;
    db.sections.splice(draft.idx, 1); ui.section = Math.min(ui.section, db.sections.length - 1); saveDb(); saveUi(); toast('Раздел удалён'); back();
  };
}

/* ───────── порядок разделов ───────── */
function renderOrder() {
  $app.innerHTML = `<div class="screen"><div class="topbar"><button class="iconbtn ghost" id="b-back">←</button><h1>Порядок разделов</h1></div>
    <div class="form"><ul class="list" id="ord">${db.sections.map((s, i) => `<li data-i="${i}"><div class="grow">${esc(s.name)} <span class="muted">· ${s.people.length}</span></div>
      <button data-a="up"${i === 0 ? ' disabled' : ''}>▲</button><button data-a="dn"${i === db.sections.length - 1 ? ' disabled' : ''}>▼</button></li>`).join('')}</ul></div></div>`;
  document.getElementById('b-back').onclick = back;
  document.getElementById('ord').onclick = (e) => {
    const b = e.target.closest('button'); if (!b) return;
    const i = +b.closest('li').dataset.i, j = b.dataset.a === 'up' ? i - 1 : i + 1;
    if (j < 0 || j >= db.sections.length) return;
    const cur = db.sections[ui.section];
    [db.sections[i], db.sections[j]] = [db.sections[j], db.sections[i]];
    ui.section = db.sections.indexOf(cur); saveDb(); saveUi(); renderOrder();
  };
}

/* ───────── выгрузка в Excel ───────── */
function renderExport() {
  const checked = state.checked || (state.checked = new Set([ui.section]));
  $app.innerHTML = `<div class="screen"><div class="topbar"><button class="iconbtn ghost" id="b-back">←</button><h1>Выгрузка в Excel</h1></div>
    <div class="form"><div class="muted" style="margin:0 2px 10px">Отметьте разделы — каждый попадёт на отдельный лист.</div>
      ${db.sections.map((s, i) => `<label class="chk"><input type="checkbox" data-i="${i}"${checked.has(i) ? ' checked' : ''}><span>${esc(s.name)} <span class="muted">· ${s.people.length}</span></span></label>`).join('')}
      <div style="display:flex;gap:8px;margin-top:6px"><button id="b-all" style="flex:1">Выбрать все</button><button id="b-none" style="flex:1">Снять все</button></div></div>
    <div class="actionbar"><button class="primary" id="b-go">Сохранить в «Загрузки»</button><button class="accent" id="b-send">Отправить…</button></div></div>`;
  document.getElementById('b-back').onclick = back;
  const sync = () => { checked.clear(); document.querySelectorAll('[data-i]').forEach((c) => c.checked && checked.add(+c.dataset.i)); };
  document.querySelectorAll('[data-i]').forEach((c) => c.onchange = sync);
  document.getElementById('b-all').onclick = () => { document.querySelectorAll('[data-i]').forEach((c) => c.checked = true); sync(); };
  document.getElementById('b-none').onclick = () => { document.querySelectorAll('[data-i]').forEach((c) => c.checked = false); sync(); };
  const run = async (how) => {
    sync(); if (!checked.size) { toast('Отметьте хотя бы один раздел'); return; }
    const idx = [...checked].sort((a, b) => a - b), d = new Date(), p = (n) => String(n).padStart(2, '0');
    const stamp = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`;
    const fname = (idx.length === 1 ? db.sections[idx[0]].name.replace(/[\\/:*?"<>|]+/g, ' ').trim() : 'S21_экспорт') + `_${stamp}.xlsx`;
    try { const r = await deliver(buildXlsx(idx), fname, how); if (r !== 'cancelled') toast(r === 'shared' ? 'Файл отправлен' : `Файл «${fname}» сохранён в «Загрузки»`, 4500); }
    catch (e) { alertBox('Ошибка', 'Не удалось создать файл: ' + e.message); }
  };
  document.getElementById('b-go').onclick = () => run('save');
  document.getElementById('b-send').onclick = () => run('share');
}

/* ───────── резервная копия / перенос ───────── */
function renderBackup() {
  const hasPrev = !!localStorage.getItem(LS_PREV);
  $app.innerHTML = `<div class="screen"><div class="topbar"><button class="iconbtn ghost" id="b-back">←</button><h1>Резервная копия</h1></div>
    <div class="form">
      <div class="card"><b>Сейчас в телефоне</b><div class="kv">Разделов: <span>${db.sections.length}</span> · людей: <span>${countPeople()}</span></div></div>
      <div class="sec-title">Сохранить</div>
      <button id="b-save" class="primary" style="width:100%">⬆ Сохранить копию в «Загрузки» (s21_data.json)</button>
      <button id="b-send" style="width:100%;margin-top:8px">📤 Отправить копию…</button>
      <div class="muted" style="margin:6px 2px 0;font-size:13px">Файл можно отправить в мессенджер или облако и положить рядом с программой на компьютере.</div>
      <div class="sec-title">Загрузить</div>
      <button id="b-load" style="width:100%">⬇ Загрузить данные из файла…</button>
      <input type="file" id="file" accept=".json,application/json" style="display:none">
      <div class="muted" style="margin:6px 2px 0;font-size:13px">Текущие данные будут заменены. Перед этим телефон запомнит их — их можно вернуть.</div>
      ${hasPrev ? '<button id="b-undo" style="width:100%;margin-top:12px">↩ Вернуть данные, бывшие до загрузки</button>' : ''}
    </div></div>`;
  document.getElementById('b-back').onclick = back;
  const saveCopy = async (how) => {
    const blob = new Blob([JSON.stringify(db, null, 1)], {type: 'application/json'});
    const r = await deliver(blob, 's21_data.json', how); if (r !== 'cancelled') toast(r === 'shared' ? 'Копия отправлена' : 'Копия сохранена в «Загрузки»', 4000);
  };
  document.getElementById('b-save').onclick = () => saveCopy('save');
  document.getElementById('b-send').onclick = () => saveCopy('share');
  const file = document.getElementById('file');
  document.getElementById('b-load').onclick = () => file.click();
  file.onchange = async () => {
    const f = file.files[0]; if (!f) return;
    try {
      const incoming = normalizeDb(JSON.parse(await f.text()));
      const n = incoming.sections.reduce((a, s) => a + s.people.length, 0);
      if (!(await confirmBox('Заменить данные?', `В файле: разделов ${incoming.sections.length}, людей ${n}.\nСейчас в телефоне: разделов ${db.sections.length}, людей ${countPeople()}.\n\nТекущие данные будут заменены.`, 'Заменить'))) return;
      try { localStorage.setItem(LS_PREV, JSON.stringify(db)); } catch (e) {}
      db = incoming; ui.section = 0; ui.q = ''; if (saveDb()) { saveUi(); toast('Данные загружены'); renderBackup(); }
    } catch (e) { alertBox('Не получилось прочитать файл', String(e.message || e)); }
    file.value = '';
  };
  const undo = document.getElementById('b-undo');
  if (undo) undo.onclick = async () => {
    if (!(await confirmBox('Вернуть прежние данные?', 'Сейчас загруженные данные будут заменены теми, что были до загрузки.', 'Вернуть'))) return;
    try { db = normalizeDb(JSON.parse(localStorage.getItem(LS_PREV))); localStorage.removeItem(LS_PREV); ui.section = 0; saveDb(); saveUi(); toast('Прежние данные возвращены'); renderBackup(); }
    catch (e) { alertBox('Ошибка', String(e.message || e)); }
  };
}

/* ───────── запуск ───────── */
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); installEvt = e; });
loadAll();
history.replaceState({screen: 'home'}, '');
render();
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
if ('serviceWorker' in navigator && /^https?:/.test(location.protocol)) navigator.serviceWorker.register('sw.js').catch(() => {});

// для проверок
window.S21 = {get db() { return db; }, get ui() { return ui; }, buildXlsx, normalizeDb, saveDb, loadAll, renderHome, nameKey, zipStore};
