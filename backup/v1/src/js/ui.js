'use strict';
/* 화면: S3 기록 목록 · S1 업로드 · S2 대시보드(결과/진단/비교) · S4 설정 */

const $ = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

const S = {
  screen: 'list',
  trainings: [],
  modelCache: {},
  expanded: new Set(),
  listSel: [],
  cur: null,        // {info, model, opinions, raw}
  dirty: false,
  nodeId: null, side: '전체', tab: 'result',
  charts: [], panels: [],
  unitMetric: 'core',
  cmp: { sel: null, level: 2, side: '훈련부대', unit: '', ind: 'fire_hit' },
  upload: null,
  present: false, presentIdx: 0, scale: 1,
};
const CUR_KEY = '__cur__';

/* ======================= 공통 ======================= */
function go(screen) {
  if (screen === 'dash' && !S.cur) return;
  S.screen = screen;
  for (const s of ['list', 'upload', 'dash', 'settings']) $('#scr-' + s).hidden = s !== screen;
  $$('.nav-btn').forEach(b => b.classList.toggle('on', b.dataset.go === screen));
  $('#navDash').disabled = !S.cur;
  if (screen === 'list') renderList();
  if (screen === 'dash') renderDash();
  if (screen === 'settings') renderSettings();
  window.scrollTo(0, 0);
}

function updateStoreBadge() {
  $('#storeBadge').textContent = Store.label;
  $('#reconnect').hidden = !(Store.mode === 'A' && Store.needsReconnect);
}

async function ensureStore() {
  if (Store.ready) return true;
  if (Store.mode === 'A') {
    try {
      if (Store.needsReconnect) await Store.reconnect(); else await Store.pickFolder();
      updateStoreBadge();
      return true;
    } catch (e) { U.toast('저장 폴더가 지정되지 않았습니다.', 'err'); return false; }
  }
  return false;
}

async function refreshList() {
  if (!Store.ready) { S.trainings = []; return; }
  try { S.trainings = await Repo.list(); } catch (e) { S.trainings = []; U.toast('기록 목록을 읽지 못했습니다: ' + e.message, 'err'); }
}

function confirmLeave() {
  return !S.dirty || confirm('현재 훈련에 저장하지 않은 변경 사항이 있습니다. 계속하면 변경 사항이 사라집니다. 계속할까요?');
}

/* 이전 훈련(시작일이 더 이른 저장된 훈련)들의 수준별 통계 */
function prevStatsFor(info) {
  return S.trainings
    .filter(t => t.id !== info.id && t.levelStats && String(t.시작일 || '') < String(info.시작일 || ''))
    .map(t => t.levelStats);
}

/* ======================= S3 기록 목록 ======================= */
function renderList() {
  const body = $('#listBody');
  if (!Store.ready) {
    body.innerHTML = `<div class="empty">저장 폴더가 지정되지 않았습니다.<br><br>
      <button class="btn primary" id="btnPickFolder2">저장 폴더 지정</button>
      <p class="hint">공유 폴더(예: AAR_데이터)를 한 번 지정하면 이후 자동으로 읽고 씁니다.</p></div>`;
    $('#btnPickFolder2').onclick = async () => { if (await ensureStore()) { await refreshList(); renderList(); } };
    return;
  }
  const q = $('#listSearch').value.trim().toLowerCase();
  const list = S.trainings.filter(t => !q || (t.훈련명 + ' ' + (t.비고 || '') + ' ' + (t.훈련부대 || '')).toLowerCase().includes(q));
  S.listSel = S.listSel.filter(f => S.trainings.some(t => t.폴더 === f));
  $('#btnCompareSel').disabled = S.listSel.length < 2;
  if (!S.trainings.length) {
    body.innerHTML = '<div class="empty">저장된 훈련이 없습니다. [+ 새 훈련]으로 첫 훈련을 분석하세요.</div>';
    return;
  }
  let h = `<table class="tlist"><thead><tr><th style="width:34px"></th><th style="width:26px"></th><th>훈련명</th><th>기간</th><th>훈련부대</th><th>저장일</th><th>소견</th><th>분석기준</th><th></th></tr></thead><tbody>`;
  for (const t of list.slice().reverse()) {
    const f = U.esc(t.폴더);
    const sel = S.listSel.includes(t.폴더);
    const opN = (t.소견제대 || []).length;
    h += `<tr>
      <td><input type="checkbox" data-sel="${f}" ${sel ? 'checked' : ''} ${!sel && S.listSel.length >= 5 ? 'disabled' : ''} title="비교 대상 선택"></td>
      <td><button class="caret" data-exp="${f}">${S.expanded.has(t.폴더) ? '▼' : '▶'}</button></td>
      <td><b>${U.esc(t.훈련명)}</b>${t.비고 ? `<div class="hint">${U.esc(t.비고)}</div>` : ''}</td>
      <td>${U.esc(t.시작일 || '')} ~ ${U.esc(t.종료일 || '')}</td>
      <td>${U.esc(t.훈련부대 || '')}</td>
      <td>${U.esc(t.수정일 || '')}</td>
      <td>${opN ? `<span class="ok">작성 ${opN}개 제대</span>` : '<span class="hint">미작성</span>'}</td>
      <td>v${U.esc(t.분석기준버전 || '')}</td>
      <td class="acts">
        <button class="btn sm primary" data-open="${f}">열기</button>
        <button class="btn sm" data-reanalyze="${f}" title="현재 설정(지표 정의·기준값)으로 원천 데이터를 다시 분석">재분석</button>
        <button class="btn sm" data-export="${f}">내보내기</button>
        <button class="btn sm danger" data-del="${f}">삭제</button>
      </td></tr>`;
    if (S.expanded.has(t.폴더)) h += `<tr class="sub"><td colspan="9"><div class="tree" data-tree="${f}">불러오는 중…</div></td></tr>`;
  }
  h += '</tbody></table>';
  body.innerHTML = h;
  $$('[data-tree]', body).forEach(el => renderTree(el, el.dataset.tree));
}

async function renderTree(el, folder) {
  try {
    const model = S.modelCache[folder] || (S.modelCache[folder] = await Repo.loadModel(folder));
    const info = S.trainings.find(t => t.폴더 === folder) || {};
    const ops = new Set(info.소견제대 || []);
    const nodes = model.nodes;
    const li = n => {
      const kids = n.children.map(id => nodes[id]).filter(Boolean);
      return `<li><a data-open="${U.esc(folder)}" data-node="${U.esc(n.id)}">${U.esc(n.name)}</a>${ops.has(n.id) ? '<span class="op-flag">✎소견</span>' : ''}` +
        (kids.length ? `<ul>${kids.map(li).join('')}</ul>` : '') + '</li>';
    };
    let h = `<ul><li><a data-open="${U.esc(folder)}" data-node="ROOT|전체">훈련 전체</a>${ops.has('ROOT|전체') ? '<span class="op-flag">✎소견</span>' : ''}<ul>`;
    for (const side of ['훈련부대', '대항군']) {
      const r = nodes['ROOT|' + side];
      if (!r) continue;
      h += `<li><span class="side">${side}</span> <a data-open="${U.esc(folder)}" data-node="ROOT|${side}" data-side="${side}">(진영 전체)</a><ul>${r.children.map(id => li(nodes[id])).join('')}</ul></li>`;
    }
    el.innerHTML = h + '</ul></li></ul>';
  } catch (e) { el.textContent = '제대 트리를 읽지 못했습니다: ' + e.message; }
}

async function openTraining(folder, nodeId) {
  if (!(S.cur && S.cur.info.폴더 === folder)) {
    if (!confirmLeave()) return;
    U.busy('훈련 불러오는 중…');
    try {
      const r = await Repo.load(folder);
      S.cur = { info: r.info, model: r.model, opinions: r.opinions, raw: null };
      S.dirty = false;
      S.modelCache[folder] = r.model;
    } catch (e) { U.busy(); U.toast('열기 실패: ' + e.message, 'err'); return; }
    U.busy();
  }
  S.side = nodeId && nodeId.startsWith('ROOT|') ? nodeId.slice(5) : '전체';
  S.nodeId = nodeId && S.cur.model.nodes[nodeId] ? nodeId : 'ROOT|' + S.side;
  if (!S.cur.model.nodes[S.nodeId]) { S.side = '전체'; S.nodeId = 'ROOT|전체'; }
  S.tab = 'result';
  go('dash');
}

async function reanalyze(folder) {
  if (S.cur && S.cur.info.폴더 === folder && !confirmLeave()) return;
  U.busy('원천 데이터 다시 분석 중…');
  try {
    const r = await Repo.load(folder).catch(async () => ({ info: Object.assign(await Store.readJSON(folder + '/_훈련정보.json'), { 폴더: folder }), opinions: {} }));
    const buf = await Repo.loadRaw(folder);
    if (!buf) throw new Error('원천 엑셀(_원천/원천데이터.xlsx)이 없습니다.');
    await U.nextFrame();
    const parsed = Engine.parse(Engine.readWorkbook(buf), Config.cur, U.dayMs(r.info.시작일));
    const model = Engine.analyze(parsed.tables, Config.cur, r.info);
    S.cur = { info: r.info, model, opinions: r.opinions || {}, raw: null };
    const info = await Repo.save({ info: r.info, model, opinions: S.cur.opinions, raw: null, prevStats: prevStatsFor(r.info) });
    S.cur.info = info;
    S.dirty = false;
    S.modelCache[folder] = model;
    await refreshList();
    U.busy();
    const nErr = parsed.errors.length;
    U.toast(`재분석 완료 (분석기준 v${model.분석기준버전}${nErr ? `, 오류 ${nErr}건 제외` : ''})`, 'ok');
    S.side = '전체'; S.nodeId = 'ROOT|전체'; S.tab = 'result';
    go('dash');
  } catch (e) { U.busy(); U.toast('재분석 실패: ' + e.message, 'err'); }
}

async function removeTraining(folder) {
  const t = S.trainings.find(x => x.폴더 === folder);
  if (!confirm(`'${t ? t.훈련명 : folder}' 훈련의 원천 데이터, 분석 결과, 소견이 모두 삭제됩니다.\n삭제 후에는 되돌릴 수 없습니다. 삭제할까요?`)) return;
  try {
    await Repo.remove(folder);
    if (S.cur && S.cur.info.폴더 === folder) { S.cur = null; S.dirty = false; }
    delete S.modelCache[folder];
    S.listSel = S.listSel.filter(f => f !== folder);
    await refreshList();
    renderList();
    $('#navDash').disabled = !S.cur;
    U.toast('삭제했습니다.');
  } catch (e) { U.toast('삭제 실패: ' + e.message, 'err'); }
}

async function exportTraining(folder) {
  U.busy('내보내기 파일 만드는 중…');
  try {
    const blob = await Repo.exportZip(folder);
    U.download(blob, folder + '.zip');
  } catch (e) { U.toast('내보내기 실패: ' + e.message, 'err'); }
  U.busy();
}

function bindList() {
  $('#listSearch').addEventListener('input', renderList);
  $('#listBody').addEventListener('click', e => {
    const b = e.target.closest('[data-open],[data-exp],[data-reanalyze],[data-export],[data-del]');
    if (!b) return;
    if (b.dataset.exp) { S.expanded.has(b.dataset.exp) ? S.expanded.delete(b.dataset.exp) : S.expanded.add(b.dataset.exp); renderList(); }
    else if (b.dataset.open) openTraining(b.dataset.open, b.dataset.node);
    else if (b.dataset.reanalyze) reanalyze(b.dataset.reanalyze);
    else if (b.dataset.export) exportTraining(b.dataset.export);
    else if (b.dataset.del) removeTraining(b.dataset.del);
  });
  $('#listBody').addEventListener('change', e => {
    const f = e.target.dataset.sel;
    if (!f) return;
    if (e.target.checked) { if (S.listSel.length < 5) S.listSel.push(f); } else S.listSel = S.listSel.filter(x => x !== f);
    renderList();
  });
  $('#btnCompareSel').onclick = async () => {
    const sel = S.trainings.filter(t => S.listSel.includes(t.폴더));
    const latest = sel[sel.length - 1];
    const keepCur = S.cur && S.listSel.includes(S.cur.info.폴더);
    if (!keepCur) { await openTraining(latest.폴더); if (!S.cur || S.cur.info.폴더 !== latest.폴더) return; }
    S.cmp.sel = S.listSel.slice();
    S.tab = 'cmp';
    go('dash');
  };
  $('#btnImportZip').onclick = async () => { if (await ensureStore()) $('#fileZip').click(); };
  $('#fileZip').onchange = async e => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    U.busy('불러오는 중…');
    try {
      const info = await Repo.importZip(await f.arrayBuffer());
      await refreshList(); renderList();
      U.toast(`'${info.훈련명}' 훈련을 불러왔습니다.`, 'ok');
    } catch (err) { U.toast('불러오기 실패: ' + err.message, 'err'); }
    U.busy();
  };
}

/* ======================= S1 업로드 ======================= */
function bindUpload() {
  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) pickFile(f); });
  $('#fileXlsx').addEventListener('change', e => { const f = e.target.files[0]; if (f) pickFile(f); e.target.value = ''; });
  $('#infoForm').addEventListener('submit', e => { e.preventDefault(); runParse(); });
}

async function pickFile(file) {
  if (!/\.xlsx?$/i.test(file.name)) { U.toast('엑셀 파일(.xlsx)만 올릴 수 있습니다.', 'err'); return; }
  setUploadFile(file.name, await file.arrayBuffer());
}
function setUploadFile(name, buf) {
  S.upload = { name, buf };
  $('#drop').classList.add('has');
  $('#dropName').innerHTML = `<b>${U.esc(name)}</b> (${(buf.byteLength / 1048576).toFixed(1)} MB) · 다른 파일을 선택하려면 다시 누르세요`;
  const form = $('#infoForm');
  if (!form.훈련명.value) {
    const m = name.replace(/\.xlsx?$/i, '').replace(/^샘플\d+_/, '');
    const d = m.match(/^(\d{4}-\d{2})_(.+)$/);
    form.훈련명.value = d ? d[2] : m;
  }
  $('#btnAnalyze').disabled = false;
  $('#validation').innerHTML = '';
}

function formInfo() {
  const f = $('#infoForm');
  return { id: null, 훈련명: f.훈련명.value.trim() || '이름없는 훈련', 시작일: f.시작일.value, 종료일: f.종료일.value, 훈련부대: f.훈련부대.value.trim(), 비고: f.비고.value.trim() };
}

async function runParse() {
  if (!S.upload) return;
  if (!confirmLeave()) return;
  const t0 = performance.now();
  U.busy('엑셀 읽는 중…');
  await U.nextFrame();
  let parsed;
  try {
    const info = formInfo();
    parsed = Engine.parse(Engine.readWorkbook(S.upload.buf), Config.cur, U.dayMs(info.시작일));
  } catch (e) { U.busy(); U.toast('엑셀을 읽지 못했습니다: ' + e.message, 'err'); return; }
  U.busy();
  S.upload.parsed = parsed;
  S.upload.t0 = t0;
  const problems = parsed.errors.length || parsed.missingSheets.length || parsed.missingCols.length;
  if (!problems) return runAnalyze();
  renderValidation(parsed);
}

function renderValidation(p) {
  const nErrRows = Object.values(p.errorRows).reduce((s, x) => s + x.size, 0);
  const usable = Object.values(p.tables).some(t => t.length);
  let h = `<div class="card"><h3>데이터 검증 결과</h3><table class="vtable"><thead><tr><th>시트</th><th>관련 기능</th><th class="num">전체 행</th><th class="num">정상</th><th class="num">오류 행</th><th>상태</th></tr></thead><tbody>`;
  for (const key of Object.keys(SHEET_DEFS)) {
    const miss = p.missingSheets.includes(key);
    const mc = p.missingCols.find(m => m.sheet === key);
    const errN = p.errorRows[key].size;
    let st = '<span class="ok">정상</span>';
    if (miss) st = '<span class="err">시트 없음 → 관련 지표 "자료 없음"</span>';
    else if (mc) st = `<span class="err">필수 열 누락: ${mc.cols.map(U.esc).join(', ')}</span>`;
    else if (errN) st = `<span class="warn">오류 행 ${errN}개</span>`;
    h += `<tr><td>${key}</td><td>${SHEET_DEFS[key].fn}</td><td class="num">${p.counts[key] != null ? p.counts[key].toLocaleString() : '–'}</td><td class="num">${p.tables[key].length.toLocaleString()}</td><td class="num">${errN || ''}</td><td>${st}</td></tr>`;
  }
  h += '</tbody></table>';
  if (p.missingSheets.length || p.missingCols.length) {
    h += `<p class="hint">시트 이름이나 열 이름이 다르면 <a href="#" id="toSettings">[설정 › 열 매칭]</a>에서 실제 이름을 지정한 뒤 다시 [분석]을 누르세요. (엑셀 시트: ${p.sheetNames.map(U.esc).join(', ')})</p>`;
  }
  if (p.errors.length) {
    h += `<div class="errlist"><table class="vtable"><thead><tr><th>시트</th><th class="num">행 번호</th><th>열</th><th>값</th><th>사유</th></tr></thead><tbody>`;
    for (const e of p.errors.slice(0, 500)) h += `<tr><td>${e.sheet}</td><td class="num">${e.row}</td><td>${U.esc(e.col)}</td><td>${U.esc(e.value)}</td><td>${e.reason}</td></tr>`;
    if (p.errors.length > 500) h += `<tr><td colspan="5" class="hint">… 외 ${p.errors.length - 500}건</td></tr>`;
    h += '</tbody></table></div>';
  }
  h += `<div class="toolbar" style="margin-top:12px">` +
    (usable ? `<button class="btn primary" id="btnProceed">${nErrRows ? `오류 행 ${nErrRows}개 제외하고 분석 진행` : '누락 항목 없이 분석 진행'}</button>` : '<span class="err">분석할 수 있는 데이터가 없습니다.</span>') +
    `<button class="btn" id="btnCancelVal">취소</button></div></div>`;
  $('#validation').innerHTML = h;
  if ($('#btnProceed')) $('#btnProceed').onclick = runAnalyze;
  $('#btnCancelVal').onclick = () => { $('#validation').innerHTML = ''; };
  if ($('#toSettings')) $('#toSettings').onclick = e => { e.preventDefault(); go('settings'); };
}

async function runAnalyze() {
  U.busy('분석 중…');
  await U.nextFrame();
  try {
    const p = S.upload.parsed;
    const info = formInfo();
    info.원천파일명 = S.upload.name;
    info.제외행수 = p.errors.length ? Object.values(p.errorRows).reduce((s, x) => s + x.size, 0) : 0;
    const model = Engine.analyze(p.tables, Config.cur, info);
    if (!info.시작일) info.시작일 = U.dateStr(model.tStart);
    if (!info.종료일) info.종료일 = U.dateStr(model.tEnd);
    S.cur = { info, model, opinions: {}, raw: S.upload.buf };
    S.dirty = true;
    S.side = '전체'; S.nodeId = 'ROOT|전체'; S.tab = 'result'; S.cmp.sel = null;
    const sec = ((performance.now() - S.upload.t0) / 1000).toFixed(1);
    $('#validation').innerHTML = '';
    $('#infoForm').reset();
    $('#drop').classList.remove('has');
    $('#dropName').textContent = '시트: 교전기록 · 탐지기록 · 위치기록 · 피해기록 · 지휘통신 · 지속지원';
    $('#btnAnalyze').disabled = true;
    S.upload = null;
    U.busy();
    go('dash');
    U.toast(`분석 완료 (${sec}초). 소견을 보완한 뒤 [저장]을 누르세요.`, 'ok');
  } catch (e) { U.busy(); console.error(e); U.toast('분석 실패: ' + e.message, 'err'); }
}

/* ======================= S2 대시보드 ======================= */
function curNode() {
  const nodes = S.cur.model.nodes;
  if (!nodes[S.nodeId]) S.nodeId = nodes['ROOT|' + S.side] ? 'ROOT|' + S.side : 'ROOT|전체';
  return nodes[S.nodeId];
}
function kids(n) {
  const k = n.children.map(id => S.cur.model.nodes[id]).filter(Boolean);
  return k.length ? k : [n];
}
function unitLabel(n) { return n.level === 0 ? (S.side === '전체' ? '전체' : S.side) : n.path.join('-'); }
function crumbText() {
  const n = curNode();
  const parts = [S.cur.info.훈련명 + (S.side !== '전체' ? ` (${S.side})` : '')];
  if (n.level > 0) parts.push(...n.path);
  return parts.join(' > ');
}
function drill(id) { if (id && id !== S.nodeId && S.cur.model.nodes[id]) { S.nodeId = id; renderDash(); } }
function markDirty() { S.dirty = true; $('#dirtyMark').hidden = false; }

function renderDash() {
  const n = curNode();
  // 경로 표시
  let h = `<a data-node="ROOT|${S.side}">${U.esc(S.cur.info.훈련명)}</a>`;
  if (S.side !== '전체') h += ` <span class="hint">(${S.side})</span>`;
  for (let i = 0; i < n.path.length; i++) {
    const id = n.side + '|' + n.path.slice(0, i + 1).join('|');
    h += `<span class="sep">›</span>` + (i === n.path.length - 1 ? `<span>${U.esc(n.path[i])}</span>` : `<a data-node="${U.esc(id)}">${U.esc(n.path[i])}</a>`);
  }
  $('#crumb').innerHTML = h;
  $$('#sideSeg button').forEach(b => {
    b.classList.toggle('on', b.dataset.side === S.side);
    b.disabled = b.dataset.side !== '전체' && !S.cur.model.nodes['ROOT|' + b.dataset.side];
  });
  $('#dirtyMark').hidden = !S.dirty;
  $$('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === S.tab));
  for (const t of ['result', 'diag', 'cmp']) $('#tab-' + t).hidden = t !== S.tab;
  renderTab();
}

function destroyCharts() { S.charts.forEach(c => c.destroy()); S.charts = []; }

/* 패널: {id, title, sub, wide, chart(s)→config, html()→string, draw(ctx,x,y,w,h,s), cls, headExtra, after(el)} */
function mountPanels(container, panels) {
  destroyCharts();
  S.panels = panels;
  container.innerHTML = panels.map(p => `
    <div class="panel ${p.wide ? 'wide' : ''} ${p.cls || ''}" data-pid="${p.id}">
      <div class="panel-h"><h3>${U.esc(p.title)}</h3>${p.sub ? `<span class="sub">${U.esc(p.sub)}</span>` : ''}<span class="grow"></span>${p.headExtra || ''}
        <button class="btn sm" data-png="${p.id}" title="1920×1080 PNG로 저장">PNG</button></div>
      ${p.chart ? '<div class="chart-box"><canvas></canvas></div>' : p.html ? p.html() : ''}
    </div>`).join('');
  for (const p of panels) {
    const el = $(`[data-pid="${p.id}"]`, container);
    if (p.chart) {
      const cfg = p.chart(S.scale);
      p.instance = new Chart($('canvas', el), cfg);
      S.charts.push(p.instance);
    }
    if (p.after) p.after(el);
  }
  if (S.present) showPresentPanel();
}

function renderTab() {
  if (S.tab === 'result') renderResult();
  else if (S.tab === 'diag') renderDiag();
  else renderCompare();
}

/* ---------- ① 결과 분석 ---------- */
function onBarClick(list) {
  return {
    onClick: (evt, els) => { if (els.length) drill(list[els[0].index].id); },
    onHover: (evt, els) => { evt.native.target.style.cursor = els.length && list[els[0].index].id !== S.nodeId ? 'pointer' : 'default'; },
  };
}

function renderResult() {
  const n = curNode();
  const ch = kids(n);
  const labels = ch.map(c => c.name);
  const lv = ch[0].level;
  const childLevel = ch[0] === n ? LEVELS[n.level] : LEVELS[lv];
  const pct = v => (v == null ? null : +(v * 100).toFixed(1));
  const panels = [];

  const kpiItems = [
    ['총 교전 수', U.fmt(n.engagements, 'int'), `교전 간격 ${S.cur.model.params.교전간격분}분 기준`],
    ['총 사격 수', U.fmt(n.shots, 'int'), ''],
    ['명중 수', U.fmt(n.hits, 'int'), `명중률 ${U.fmt(n.metrics.fire_hit, 'pct')}`],
    ['피해 유발', U.fmt(n.dmgHits, 'int'), `사격 1회당 ${U.fmt(n.metrics.fire_dmg_per_shot, 'num3')}`],
    ['인원', U.fmt(n.personnel, 'int'), `사망 ${n.cas.사망} · 중상 ${n.cas.중상} · 경상 ${n.cas.경상}`],
    ['생존율', U.fmt(n.metrics.prot_survival, 'pct'), '(인원 − 사망) ÷ 인원'],
  ];
  panels.push({
    id: 'kpi', title: '교전 현황', sub: crumbText(), wide: true,
    html: () => `<div class="kpis">${kpiItems.map(k => `<div class="kpi"><div class="k">${k[0]}</div><div class="v">${k[1]}</div><div class="d">${k[2]}</div></div>`).join('')}</div>`,
    draw: (ctx, x, y, w, h, s) => {
      const cols = 3, rows = 2, gap = 16 * s, bw = (w - gap * (cols - 1)) / cols, bh = Math.min(220 * s, (h - gap) / rows);
      kpiItems.forEach((k, i) => {
        const bx = x + (i % cols) * (bw + gap), by = y + Math.floor(i / cols) * (bh + gap);
        ctx.fillStyle = '#f4f4f1'; Draw.round(ctx, bx, by, bw, bh, 8); ctx.fill();
        Draw.text(ctx, k[0], bx + 16 * s, by + 14 * s, { size: 15 * s, color: COL.ink2 });
        Draw.text(ctx, k[1], bx + 16 * s, by + 40 * s, { size: 38 * s, weight: 700 });
        Draw.text(ctx, k[2], bx + 16 * s, by + 92 * s, { size: 12 * s, color: COL.muted, maxW: bw - 32 * s });
      });
    },
  });

  panels.push({
    id: 'hit', title: '명중률', sub: `${childLevel}별 · 막대를 누르면 하위 제대로 이동`,
    chart: s => ({
      type: 'bar',
      data: { labels, datasets: [Charts.bar('명중률(%)', ch.map(c => pct(c.metrics.fire_hit)), COL.series[0])] },
      options: Charts.opts(s, Object.assign({
        plugins: { legend: { display: false }, valueLabels: { enabled: true, size: 12 * s, fmt: v => v + '%' }, tooltip: { callbacks: { label: c => `명중률 ${c.raw}% (사격 ${ch[c.dataIndex].shots.toLocaleString()}회, 명중 ${ch[c.dataIndex].hits.toLocaleString()})` } } },
        scales: { y: { ticks: { callback: v => v + '%' } } },
      }, onBarClick(ch))),
    }),
  });

  const alive = c => Math.max(0, c.personnel - c.cas.사망 - c.cas.중상 - c.cas.경상);
  panels.push({
    id: 'cas', title: '피해/생존 현황', sub: `${childLevel}별 인원(명)`,
    chart: s => ({
      type: 'bar',
      data: {
        labels, datasets: [
          ['사망', c => c.cas.사망], ['중상', c => c.cas.중상], ['경상', c => c.cas.경상], ['무피해(생존)', alive],
        ].map(([k, fn]) => Object.assign(Charts.bar(k, ch.map(fn), COL[k] || COL.무피해), { borderRadius: 0, borderColor: COL.surface, borderWidth: { top: 2 }, borderSkipped: false })),
      },
      options: Charts.opts(s, Object.assign({
        scales: { x: { stacked: true }, y: { stacked: true } },
        plugins: { tooltip: { mode: 'index', callbacks: { footer: items => `인원 ${ch[items[0].dataIndex].personnel}명 · 생존율 ${U.fmt(ch[items[0].dataIndex].metrics.prot_survival, 'pct')}` } } },
      }, onBarClick(ch))),
    }),
  });

  const m = S.cur.model;
  const tLabels = n.tsEng.map((_, i) => U.timeStr(m.t0 + i * m.binMin * 60000));
  panels.push({
    id: 'ts', title: '시간대별 추이', sub: `${m.binMin}분 단위 · ${U.dateStr(m.t0)}`, wide: true,
    chart: s => ({
      type: 'line',
      data: {
        labels: tLabels, datasets: [
          { label: '교전 수(개시)', data: n.tsEng, borderColor: COL.series[0], backgroundColor: COL.series[0], borderWidth: 2 * Math.max(1, s * 0.8), pointRadius: 3 * s, tension: 0.25 },
          { label: '피해 수', data: n.tsDmg, borderColor: COL.series[1], backgroundColor: COL.series[1], borderWidth: 2 * Math.max(1, s * 0.8), pointRadius: 3 * s, tension: 0.25 },
        ],
      },
      options: Charts.opts(s, { interaction: { mode: 'index', intersect: false }, plugins: { tooltip: { mode: 'index', intersect: false } } }),
    }),
  });

  const wNames = Object.keys(n.weapons).sort((a, b) => n.weapons[b].shots - n.weapons[a].shots);
  const wLabels = wNames.map(w => [w, `명중률 ${U.fmt(n.weapons[w].shots ? n.weapons[w].hits / n.weapons[w].shots : null, 'pct')}`]);
  panels.push({
    id: 'weapon', title: '화기별 분석', sub: '사격 수 · 명중 수 · 피해 기여(건)',
    chart: s => ({
      type: 'bar',
      data: {
        labels: wLabels, datasets: [
          Charts.bar('사격 수', wNames.map(w => n.weapons[w].shots), COL.series[0]),
          Charts.bar('명중 수', wNames.map(w => n.weapons[w].hits), COL.series[1]),
          Charts.bar('피해 기여', wNames.map(w => n.weapons[w].dmg), COL.series[2]),
        ],
      },
      options: Charts.opts(s, { plugins: { tooltip: { callbacks: { title: it => wNames[it[0].dataIndex] } } } }),
    }),
  });

  const core = ['fire_hit', 'prot_survival', 'int_predetect', 'sus_ammo'];
  const metricSel = `<select id="unitMetric"><option value="core">핵심 비율 지표 4종</option>${Config.cur.indicators.map(i => `<option value="${i.id}" ${S.unitMetric === i.id ? 'selected' : ''}>${U.esc(i.fn)} · ${U.esc(i.name)}</option>`).join('')}</select>`;
  panels.push({
    id: 'units', title: '제대별 비교', sub: `하위 ${childLevel} 간 지표`, headExtra: metricSel,
    chart: s => {
      const ids = S.unitMetric === 'core' ? core : [S.unitMetric];
      const val = (c, id) => { const ind = Config.ind(id); const v = c.metrics[id]; return v == null ? null : ind.fmt === 'pct' ? pct(v) : +v.toFixed(3); };
      const single = ids.length === 1 ? Config.ind(ids[0]) : null;
      return {
        type: 'bar',
        data: { labels, datasets: ids.map((id, i) => Charts.bar(Config.ind(id).name + (Config.ind(id).fmt === 'pct' ? '(%)' : Config.ind(id).unit ? `(${Config.ind(id).unit})` : ''), ch.map(c => val(c, id)), COL.series[i])) },
        options: Charts.opts(s, Object.assign({
          plugins: {
            legend: { display: !single },
            valueLabels: { enabled: !!single, size: 12 * s, fmt: v => (single && single.fmt === 'pct' ? v + '%' : String(v)) },
            subtitle: single ? { display: true, text: `${single.dir === 'low' ? '낮을수록' : '높을수록'} 좋음`, align: 'end', font: { size: 11 * s } } : { display: false },
          },
        }, onBarClick(ch))),
      };
    },
    after: el => { $('#unitMetric', el).value = S.unitMetric; $('#unitMetric', el).onchange = e => { S.unitMetric = e.target.value; renderResult(); }; },
  });

  mountPanels($('#tab-result'), panels);
}

/* ---------- ② 6대 기능 진단 ---------- */
function diagFor(n) {
  const key = Engine.levelKey(n, S.side);
  const base = Engine.baseline(key, prevStatsFor(S.cur.info));
  return Engine.diagnose(n.metrics, base, Config.cur);
}
function opKey(n) { return n.level === 0 ? 'ROOT|' + S.side : n.id; }
const badge = j => `<span class="badge ${['강점', '취약', '보통', '참고'].includes(j) ? j : 'none'}">${j === '강점' ? '▲ ' : j === '취약' ? '▼ ' : j === '보통' ? '● ' : ''}${U.esc(j)}</span>`;

function renderDiag() {
  const n = curNode();
  const d = diagFor(n);
  const ops = S.cur.opinions[opKey(n)] || {};
  const panels = [];
  const lvName = n.level === 0 ? '훈련 전체' : LEVELS[n.level];
  const baseText = d.hasBase ? `비교 기준: 이전 훈련 ${d.baseN}건의 ${lvName} 평균 (=100)` : '비교 기준 없음 (이전 훈련 없음) — 지표 수치만 표시';

  panels.push({
    id: 'radar', title: '6대 기능 종합지수', sub: baseText,
    chart: d.hasBase ? s => {
      const vals = d.functions.map(f => (f.index == null ? null : +f.index.toFixed(0)));
      const mx = Math.max(150, ...vals.filter(v => v != null).map(v => Math.ceil(v / 25) * 25));
      return {
        type: 'radar',
        data: {
          labels: FUNCTIONS, datasets: [
            { label: '이번 훈련', data: vals, borderColor: COL.series[0], backgroundColor: 'rgba(42,120,214,0.15)', borderWidth: 2 * Math.max(1, s * 0.8), pointRadius: 4 * s, pointBackgroundColor: COL.series[0] },
            { label: '이전 평균(100)', data: FUNCTIONS.map(() => 100), borderColor: COL.muted, borderDash: [6, 4], borderWidth: 1.5 * Math.max(1, s * 0.8), pointRadius: 0, fill: false },
          ],
        },
        options: {
          responsive: true, maintainAspectRatio: false, animation: { duration: 200 },
          plugins: { legend: { position: 'top', align: 'end', labels: { font: { size: 12 * s }, boxWidth: 12 * s } }, tooltip: { bodyFont: { size: 12 * s } } },
          scales: { r: { min: 0, max: mx, ticks: { stepSize: 25, font: { size: 10 * s }, backdropColor: 'transparent', color: COL.muted }, grid: { color: COL.grid }, angleLines: { color: COL.grid }, pointLabels: { font: { size: 14 * s, weight: '700' }, color: COL.ink } } },
        },
      };
    } : null,
    html: d.hasBase ? null : () => `<div class="empty" style="flex:1">첫 훈련(또는 이보다 앞선 저장 훈련 없음)이라 비교 기준이 없습니다.<br>지표 수치는 아래 기능 카드에서 확인하세요.</div>`,
    draw: d.hasBase ? null : (ctx, x, y, w, h, s) => Draw.wrap(ctx, '비교 기준 없음 — 첫 훈련은 지표 수치만 표시합니다.', x, y + 20 * s, w, { size: 18 * s, color: COL.ink2 }),
  });

  panels.push({
    id: 'dsum', title: '진단 요약', sub: '규칙 기반 1차 판단 · 최종 판단은 담당자 소견',
    html: () => `<div style="display:flex;flex-direction:column;gap:6px">${d.functions.map(f => `<div class="fn-sum">${badge(f.judge)} ${U.esc(f.summary)}</div>`).join('')}</div>`,
    draw: (ctx, x, y, w, h, s) => {
      const lh = Math.min(h / 6, 120 * s);
      d.functions.forEach((f, i) => {
        const yy = y + i * lh;
        const bw = Draw.badge(ctx, f.judge, x, yy, 13 * s);
        Draw.wrap(ctx, f.summary, x + bw + 10 * s, yy + 1 * s, w - bw - 10 * s, { size: 14 * s, maxLines: Math.max(1, Math.floor(lh / (14 * s * 1.45)) - 0) });
      });
    },
  });

  for (const f of d.functions) {
    const ind = id => Config.ind(id);
    const rowsHtml = f.indicators.map(r => {
      const i = ind(r.id);
      const ch = r.improve == null ? '' : `<span class="${r.improve >= 0 ? 'up' : 'down'}">${r.change >= 0 ? '+' : ''}${r.change.toFixed(1)}%</span>`;
      return `<tr title="${U.esc(r.sentence)}"><td>${U.esc(r.name)}${i.dir === 'low' ? ' <span class="hint">↓좋음</span>' : ''}</td><td>${U.fmtInd(i, r.value)}</td><td>${r.base == null ? '–' : U.fmtInd(i, r.base)}</td><td>${ch}</td><td>${badge(r.judge)}</td></tr>`;
    }).join('');
    panels.push({
      id: 'fn_' + f.fn, title: f.fn, cls: 'fn-card',
      headExtra: `${badge(f.judge)}<span class="fn-idx">종합지수 <b>${f.index == null ? '–' : f.index.toFixed(0)}</b></span>`,
      html: () => `
        <table class="itable"><thead><tr><th>지표</th><th>이번</th><th>이전 평균</th><th>변화</th><th>판정</th></tr></thead><tbody>${rowsHtml}</tbody></table>
        <div class="fn-sum">${U.esc(f.summary)}</div>
        <ul class="fn-sent">${f.indicators.filter(r => r.change != null).map(r => `<li>${U.esc(r.sentence)}</li>`).join('')}</ul>
        <div class="op-label">담당자 소견 (최종 판단)</div>
        <textarea data-fn="${f.fn}" placeholder="${f.fn} 기능에 대한 소견을 입력하세요">${U.esc(ops[f.fn] || '')}</textarea>`,
      draw: (ctx, x, y, w, h, s) => {
        let yy = y;
        Draw.badge(ctx, f.judge, x, yy, 14 * s);
        Draw.text(ctx, `종합지수 ${f.index == null ? '–' : f.index.toFixed(0)}`, x + w, yy + 2 * s, { size: 16 * s, weight: 700, align: 'right' });
        yy += 34 * s;
        const cx = [x, x + w * 0.46, x + w * 0.62, x + w * 0.78, x + w * 0.88];
        ['지표', '이번', '이전 평균', '변화', '판정'].forEach((t, i) => Draw.text(ctx, t, cx[i], yy, { size: 12 * s, color: COL.ink2, weight: 600 }));
        yy += 20 * s;
        for (const r of f.indicators) {
          const i = ind(r.id);
          Draw.text(ctx, r.name, cx[0], yy, { size: 13 * s, maxW: w * 0.45 });
          Draw.text(ctx, U.fmtInd(i, r.value), cx[1], yy, { size: 13 * s });
          Draw.text(ctx, r.base == null ? '–' : U.fmtInd(i, r.base), cx[2], yy, { size: 13 * s, color: COL.ink2 });
          if (r.change != null) Draw.text(ctx, `${r.change >= 0 ? '+' : ''}${r.change.toFixed(1)}%`, cx[3], yy, { size: 13 * s, color: r.improve >= 0 ? COL.good : COL.bad });
          Draw.text(ctx, r.judge, cx[4], yy, { size: 13 * s, weight: 700, color: r.judge === '강점' ? COL.good : r.judge === '취약' ? COL.bad : COL.mid, maxW: w * 0.12 });
          yy += 22 * s;
        }
        yy += 8 * s;
        ctx.fillStyle = '#f4f4f1'; ctx.fillRect(x, yy - 4 * s, w, 2);
        yy = Draw.wrap(ctx, f.summary, x, yy + 4 * s, w, { size: 14 * s, maxLines: 4 });
        const op = (S.cur.opinions[opKey(n)] || {})[f.fn];
        if (op && op.trim() && yy < y + h - 30 * s) {
          yy += 6 * s;
          Draw.text(ctx, '담당자 소견', x, yy, { size: 12 * s, weight: 700, color: COL.ink2 });
          Draw.wrap(ctx, op, x, yy + 18 * s, w, { size: 14 * s, maxLines: Math.max(1, Math.floor((y + h - yy - 18 * s) / (14 * s * 1.45))) });
        }
      },
      after: el => {
        $('textarea', el).addEventListener('input', e => {
          const k = opKey(n);
          S.cur.opinions[k] = Object.assign({}, S.cur.opinions[k], { [f.fn]: e.target.value });
          markDirty();
        });
      },
    });
  }

  const c = $('#tab-diag');
  c.innerHTML = '';
  mountPanels(c, panels);
  // 배치: 상단(레이더+요약) / 기능 카드 3열
  const top = document.createElement('div'); top.className = 'diag-top';
  const grid = document.createElement('div'); grid.className = 'fn-grid';
  const note = document.createElement('div'); note.className = 'note';
  note.innerHTML = `<b>규칙 기반 1차 판단</b> · 같은 제대 수준에서 이전 훈련 평균 대비 변화율이 +${Config.cur.판정기준N}% 이상이면 강점, −${Config.cur.판정기준N}% 이하면 취약입니다. 근거 수치를 확인하고 <b>담당자 소견</b>으로 최종 판단을 기록하세요.`;
  $$('.panel', c).forEach(p => (p.dataset.pid === 'radar' || p.dataset.pid === 'dsum' ? top : grid).appendChild(p));
  c.append(note, top, grid);
  note.style.marginBottom = '12px';
  S.charts.forEach(ch => ch.resize());
}

/* ---------- ③ 훈련 비교 ---------- */
function cmpCandidates() {
  const list = S.trainings.map(t => ({ key: t.폴더, info: t, stats: t.levelStats }));
  const curKey = S.cur.info.폴더 || CUR_KEY;
  const i = list.findIndex(t => t.key === curKey);
  const curItem = { key: curKey, info: S.cur.info, stats: S.cur.model.levelStats, model: S.cur.model, isCur: true };
  if (i >= 0) list[i] = curItem; else list.push(curItem);
  return list.sort((a, b) => String(a.info.시작일).localeCompare(String(b.info.시작일)));
}

async function renderCompare() {
  const c = $('#tab-cmp');
  const cands = cmpCandidates();
  const curKey = S.cur.info.폴더 || CUR_KEY;
  if (!S.cmp.sel) S.cmp.sel = cands.slice(-5).map(t => t.key);
  S.cmp.sel = S.cmp.sel.filter(k => cands.some(t => t.key === k)).slice(0, 5);
  const sel = cands.filter(t => S.cmp.sel.includes(t.key));
  const lv = +S.cmp.level;
  if (lv > 0 && S.cmp.side === '전체') S.cmp.side = '훈련부대';

  // 특정 제대 선택 시 모델 필요
  if (S.cmp.unit) {
    U.busy('비교 데이터 불러오는 중…');
    for (const t of sel) if (!t.model) t.model = S.modelCache[t.key] || (S.modelCache[t.key] = await Repo.loadModel(t.key).catch(() => null));
    U.busy();
  }
  const unitOpts = new Set();
  if (lv > 0) {
    for (const t of sel) {
      const m = t.model || S.modelCache[t.key];
      if (m) Object.values(m.nodes).forEach(nn => { if (nn.level === lv && nn.side === S.cmp.side) unitOpts.add(nn.path.join(' > ')); });
    }
  }
  const valueOf = t => {
    if (lv === 0) return t.stats && t.stats['ROOT|' + S.cmp.side];
    if (!S.cmp.unit) return t.stats && t.stats[S.cmp.side + '|' + lv];
    const m = t.model;
    if (!m) return null;
    const nn = Object.values(m.nodes).find(x => x.level === lv && x.side === S.cmp.side && x.path.join(' > ') === S.cmp.unit);
    return nn ? nn.metrics : null;
  };
  const vals = sel.map(valueOf);
  const tName = t => t.info.훈련명 + (t.isCur && !t.info.폴더 ? '(미저장)' : '');
  const target = lv === 0 ? `훈련 전체 · ${S.cmp.side}` : S.cmp.unit ? `${S.cmp.side} ${S.cmp.unit}` : `${S.cmp.side} ${LEVELS[lv]} 평균`;

  destroyCharts();
  c.innerHTML = `<div class="cmp-ctl">
    <div class="group"><h4>비교 훈련 (2~5건, 시간순)</h4>${cands.map(t => {
      const on = S.cmp.sel.includes(t.key);
      return `<label><input type="checkbox" data-ck="${U.esc(t.key)}" ${on ? 'checked' : ''} ${!on && S.cmp.sel.length >= 5 ? 'disabled' : ''}> ${U.esc(tName(t))} <span class="hint">${U.esc(t.info.시작일 || '')}</span>${t.key === curKey ? ' <span class="hint">· 현재</span>' : ''}</label>`;
    }).join('')}</div>
    <div class="group"><h4>제대 수준</h4><select id="cmpLevel">${LEVELS.map((l, i) => `<option value="${i}" ${i === lv ? 'selected' : ''}>${i === 0 ? '훈련 전체' : l}</option>`).join('')}</select>
      <h4 style="margin-top:8px">진영</h4><select id="cmpSide">${SIDES.filter(s => lv === 0 || s !== '전체').map(s => `<option ${s === S.cmp.side ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
    <div class="group"><h4>대상</h4><select id="cmpUnit" ${lv === 0 ? 'disabled' : ''}><option value="">${lv === 0 ? '훈련 전체' : LEVELS[lv] + ' 평균 (수준 전체)'}</option>${[...unitOpts].sort().map(u => `<option ${u === S.cmp.unit ? 'selected' : ''}>${U.esc(u)}</option>`).join('')}</select>
      <p class="hint" style="max-width:260px">수준 평균: 각 훈련에서 해당 수준 제대들의 지표 평균. 특정 제대를 고르면 같은 이름의 제대끼리 비교합니다.</p></div>
  </div><div id="cmpPanels" class="grid"></div>`;

  $$('[data-ck]', c).forEach(el => el.onchange = () => {
    const k = el.dataset.ck;
    S.cmp.sel = el.checked ? [...S.cmp.sel, k] : S.cmp.sel.filter(x => x !== k);
    renderCompare();
  });
  $('#cmpLevel').onchange = e => { S.cmp.level = +e.target.value; S.cmp.unit = ''; renderCompare(); };
  $('#cmpSide').onchange = e => { S.cmp.side = e.target.value; S.cmp.unit = ''; renderCompare(); };
  $('#cmpUnit').onchange = e => { S.cmp.unit = e.target.value; renderCompare(); };

  if (sel.length < 2) {
    $('#cmpPanels').outerHTML = '<div class="empty">비교할 훈련을 2건 이상 선택하세요. (저장된 훈련이 1건뿐이면 다른 훈련을 먼저 분석·저장하세요)</div>';
    S.panels = [];
    return;
  }

  const inds = Config.cur.indicators;
  // 레이더: 선택 훈련 평균 = 100
  const avg = {};
  inds.forEach(i => { avg[i.id] = U.mean(vals.map(v => v && v[i.id])); });
  const fnIdx = vals.map(v => FUNCTIONS.map(fn => {
    if (!v) return null;
    const xs = inds.filter(i => i.fn === fn && i.judged !== false).map(i => Engine.indIndex(i, v[i.id], avg[i.id])).filter(x => x != null);
    return xs.length ? +U.mean(xs).toFixed(0) : null;
  }));
  const panels = [];
  panels.push({
    id: 'cmp_radar', title: '6대 기능 레이더 (겹쳐 보기)', sub: `${target} · 선택 훈련 평균 = 100`,
    chart: s => ({
      type: 'radar',
      data: {
        labels: FUNCTIONS, datasets: sel.map((t, i) => ({
          label: tName(t), data: fnIdx[i], borderColor: COL.series[i], backgroundColor: COL.series[i] + '1f',
          borderWidth: 2 * Math.max(1, s * 0.8), pointRadius: 3.5 * s, pointBackgroundColor: COL.series[i],
        })),
      },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 200 },
        plugins: { legend: { position: 'top', labels: { font: { size: 12 * s }, boxWidth: 12 * s } } },
        scales: { r: { min: 0, suggestedMax: 150, ticks: { stepSize: 25, font: { size: 10 * s }, backdropColor: 'transparent', color: COL.muted }, grid: { color: COL.grid }, angleLines: { color: COL.grid }, pointLabels: { font: { size: 14 * s, weight: '700' }, color: COL.ink } } },
      },
    }),
  });

  const indSel = `<select id="cmpInd">${inds.map(i => `<option value="${i.id}" ${i.id === S.cmp.ind ? 'selected' : ''}>${U.esc(i.fn)} · ${U.esc(i.name)}</option>`).join('')}</select>`;
  const ti = Config.ind(S.cmp.ind) || inds[0];
  panels.push({
    id: 'cmp_trend', title: '지표 추이', sub: `${target} · ${ti.name} (${ti.dir === 'low' ? '낮을수록' : '높을수록'} 좋음)`, headExtra: indSel,
    chart: s => ({
      type: 'line',
      data: { labels: sel.map(tName), datasets: [{ label: ti.name, data: vals.map(v => (v && v[ti.id] != null ? (ti.fmt === 'pct' ? +(v[ti.id] * 100).toFixed(1) : +v[ti.id].toFixed(3)) : null)), borderColor: COL.series[0], backgroundColor: COL.series[0], borderWidth: 2.5 * Math.max(1, s * 0.8), pointRadius: 5 * s, spanGaps: true }] },
      options: Charts.opts(s, {
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: c2 => `${ti.name} ${c2.raw}${ti.fmt === 'pct' ? '%' : ti.unit}` } } },
        scales: { y: { beginAtZero: false, ticks: { callback: v => v + (ti.fmt === 'pct' ? '%' : '') } } },
      }),
    }),
    after: el => { $('#cmpInd', el).onchange = e => { S.cmp.ind = e.target.value; renderCompare(); }; },
  });

  const rows = inds.map(i => ({ i, cells: vals.map((v, k) => {
    const x = v ? v[i.id] : null, p = k > 0 && vals[k - 1] ? vals[k - 1][i.id] : null;
    let arrow = '', good = null;
    if (x != null && p != null && x !== p) { const upv = x > p; good = i.dir === 'low' ? !upv : upv; arrow = upv ? '▲' : '▼'; }
    return { x, arrow, good };
  }) }));
  panels.push({
    id: 'cmp_table', title: '핵심 지표 비교', sub: `${target} · ▲▼ 직전 훈련 대비 (초록=개선, 빨강=악화)`, wide: true,
    html: () => `<div style="overflow:auto"><table class="ctable"><thead><tr><th>기능</th><th>지표</th>${sel.map(t => `<th>${U.esc(tName(t))}<div class="hint">${U.esc(t.info.시작일 || '')}</div></th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr><td class="fn">${r.i.fn}</td><td>${U.esc(r.i.name)}${r.i.judged === false ? ' <span class="hint">(참고)</span>' : ''}</td>${r.cells.map(cl => `<td>${U.fmtInd(r.i, cl.x)} ${cl.arrow ? `<span class="${cl.good ? 'up' : 'down'}">${cl.arrow}</span>` : ''}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`,
    draw: (ctx, x, y, w, h, s) => {
      const nCol = sel.length;
      const c0 = w * 0.12, c1 = w * 0.30, cw = (w - c0 - c1) / nCol;
      const lh = Math.min(30 * s, (h - 30 * s) / (rows.length + 1));
      const fs = Math.min(14 * s, lh * 0.62);
      Draw.text(ctx, '기능', x, y, { size: fs, weight: 700, color: COL.ink2 });
      Draw.text(ctx, '지표', x + c0, y, { size: fs, weight: 700, color: COL.ink2 });
      sel.forEach((t, k) => Draw.text(ctx, tName(t), x + c0 + c1 + cw * (k + 1) - 8, y, { size: fs, weight: 700, align: 'right', maxW: cw - 10 }));
      rows.forEach((r, j) => {
        const yy = y + lh * (j + 1);
        if (j % 2 === 0) { ctx.fillStyle = '#f6f6f3'; ctx.fillRect(x, yy - lh * 0.2, w, lh); }
        Draw.text(ctx, r.i.fn, x, yy, { size: fs * 0.9, color: COL.ink2 });
        Draw.text(ctx, r.i.name, x + c0, yy, { size: fs, maxW: c1 - 8 });
        r.cells.forEach((cl, k) => {
          const rx = x + c0 + c1 + cw * (k + 1) - 8;
          const aw = cl.arrow ? Draw.text(ctx, cl.arrow, rx, yy, { size: fs, align: 'right', color: cl.good ? COL.good : COL.bad }) + 4 : 0;
          Draw.text(ctx, U.fmtInd(r.i, cl.x), rx - aw, yy, { size: fs, align: 'right' });
        });
      });
    },
  });
  mountPanels($('#cmpPanels'), panels);
}

/* ---------- PNG 저장 ---------- */
const PNG_W = 1920, PNG_H = 1080;
function pngCanvas(title) {
  const cv = document.createElement('canvas');
  cv.width = PNG_W; cv.height = PNG_H;
  const ctx = cv.getContext('2d');
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, PNG_W, PNG_H);
  Draw.text(ctx, title, 60, 40, { size: 44, weight: 700 });
  Draw.text(ctx, crumbText(), 60, 100, { size: 24, color: COL.ink2 });
  ctx.fillStyle = COL.line || '#e1e0d9'; ctx.fillRect(60, 140, PNG_W - 120, 2);
  Draw.text(ctx, `사후검토 분석기 · ${U.nowStr()}`, PNG_W - 60, PNG_H - 36, { size: 16, color: COL.muted, align: 'right' });
  return { cv, ctx };
}
function drawPanelBody(p, ctx, x, y, w, h, s) {
  if (p.chart) ctx.drawImage(Charts.renderToCanvas(p.chart(s), Math.round(w), Math.round(h)), x, y);
  else if (p.draw) { ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip(); p.draw(ctx, x, y, w, h, s); ctx.restore(); }
}
function fileBase() { return `${U.safeName(S.cur.info.훈련명)}_${U.safeName(unitLabel(curNode()))}`; }

function savePanelPng(pid) {
  const p = S.panels.find(x => x.id === pid);
  if (!p) return;
  const { cv, ctx } = pngCanvas(p.title + (p.sub && !p.chart ? '' : ''));
  if (p.sub) Draw.text(ctx, p.sub, PNG_W - 60, 108, { size: 20, color: COL.muted, align: 'right', maxW: 900 });
  drawPanelBody(p, ctx, 60, 170, PNG_W - 120, PNG_H - 240, 2);
  U.download(cv.toDataURL('image/png'), `${fileBase()}_${U.safeName(p.title)}.png`);
}

function saveAllPng() {
  const tabName = { result: '결과분석', diag: '6대기능진단', cmp: '훈련비교' }[S.tab];
  const ps = S.panels.filter(p => p.chart || p.draw);
  if (!ps.length) { U.toast('저장할 그래프가 없습니다.', 'err'); return; }
  const { cv, ctx } = pngCanvas({ result: '결과 분석', diag: '6대 기능 진단', cmp: '훈련 비교' }[S.tab]);
  const cols = ps.length <= 2 ? ps.length : ps.length <= 4 ? 2 : 3;
  const rowsN = Math.ceil(ps.length / cols);
  const gx = 24, gy = 20, x0 = 60, y0 = 165, W = PNG_W - 120, H = PNG_H - 165 - 70;
  const cw = (W - gx * (cols - 1)) / cols, chh = (H - gy * (rowsN - 1)) / rowsN;
  const s = Math.max(0.9, Math.min(1.5, chh / 250));
  ps.forEach((p, i) => {
    const x = x0 + (i % cols) * (cw + gx), y = y0 + Math.floor(i / cols) * (chh + gy);
    ctx.strokeStyle = '#e1e0d9'; ctx.lineWidth = 1; Draw.round(ctx, x, y, cw, chh, 8); ctx.stroke();
    Draw.text(ctx, p.title, x + 14, y + 10, { size: 18 * s, weight: 700, maxW: cw - 28 });
    drawPanelBody(p, ctx, x + 10, y + 16 + 22 * s, cw - 20, chh - 26 - 22 * s, s);
  });
  U.download(cv.toDataURL('image/png'), `${fileBase()}_${tabName}_전체.png`);
}

/* ---------- 회의(전체화면) 모드 ---------- */
function enterPresent() {
  S.present = true; S.presentIdx = 0; S.scale = 1.7;
  document.body.classList.add('present');
  $('#present-nav').hidden = false;
  if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => {});
  renderTab();
}
function exitPresent() {
  if (!S.present) return;
  S.present = false; S.scale = 1;
  document.body.classList.remove('present');
  $('#present-nav').hidden = true;
  if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
  renderTab();
}
function showPresentPanel() {
  const els = $$(`#tab-${S.tab} .panel`);
  if (!els.length) return;
  S.presentIdx = Math.max(0, Math.min(els.length - 1, S.presentIdx));
  els.forEach((el, i) => el.classList.toggle('cur', i === S.presentIdx));
  $('#presentPos').textContent = `${S.presentIdx + 1} / ${els.length}`;
  requestAnimationFrame(() => S.charts.forEach(c => c.resize()));
}
function presentStep(d) {
  const n = $$(`#tab-${S.tab} .panel`).length;
  S.presentIdx = (S.presentIdx + d + n) % n;
  showPresentPanel();
}

function bindDash() {
  $('#crumb').addEventListener('click', e => { const a = e.target.closest('[data-node]'); if (a) drill(a.dataset.node); });
  $('#sideSeg').addEventListener('click', e => {
    const b = e.target.closest('[data-side]');
    if (!b || b.disabled) return;
    S.side = b.dataset.side;
    const n = S.cur.model.nodes[S.nodeId];
    if (!n || n.level === 0 || (S.side !== '전체' && n.side !== S.side)) S.nodeId = 'ROOT|' + S.side;
    renderDash();
  });
  $('#tabs').addEventListener('click', e => { const b = e.target.closest('[data-tab]'); if (b) { S.tab = b.dataset.tab; renderDash(); } });
  $('#scr-dash').addEventListener('click', e => { const b = e.target.closest('[data-png]'); if (b) savePanelPng(b.dataset.png); });
  $('#btnShotAll').onclick = saveAllPng;
  $('#btnPresent').onclick = enterPresent;
  $('#btnSave').onclick = saveCurrent;
  document.addEventListener('keydown', e => {
    if (!S.present) return;
    if (['ArrowRight', 'PageDown', ' '].includes(e.key)) { e.preventDefault(); presentStep(1); }
    else if (['ArrowLeft', 'PageUp'].includes(e.key)) { e.preventDefault(); presentStep(-1); }
    else if (e.key === 'Escape') exitPresent();
  });
  document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && S.present) exitPresent(); });
}

async function saveCurrent() {
  if (!S.cur) return;
  if (!(await ensureStore())) return;
  U.busy('저장 중…');
  try {
    const rawToWrite = S.cur.raw; // 처음 저장할 때만 원천 엑셀을 쓴다
    const info = await Repo.save({ info: S.cur.info, model: S.cur.model, opinions: S.cur.opinions, raw: rawToWrite, prevStats: prevStatsFor(S.cur.info) }, (i, n) => U.busy(`저장 중… (${i}/${n} 제대)`));
    S.cur.info = info; S.cur.raw = null; S.dirty = false;
    S.modelCache[info.폴더] = S.cur.model;
    if (S.cmp.sel) S.cmp.sel = S.cmp.sel.map(k => (k === CUR_KEY ? info.폴더 : k));
    await refreshList();
    U.busy();
    renderDash();
    U.toast(`저장했습니다: ${info.폴더}`, 'ok');
  } catch (e) { U.busy(); console.error(e); U.toast('저장 실패: ' + e.message, 'err'); }
}

/* ======================= S4 설정 ======================= */
function renderSettings() {
  const c = Config.cur;
  const body = $('#settingsBody');
  body.innerHTML = `<div class="set-grid">
    <div class="card"><h3>저장</h3><div class="kv">
      <span>현재 방식</span><b>${U.esc(Store.label)}</b>
      <span>저장 방식</span><select id="cfgMode">
        <option value="auto" ${c.저장방식 === 'auto' ? 'selected' : ''}>자동 (지원 시 A, 아니면 B)</option>
        <option value="A" ${c.저장방식 === 'A' ? 'selected' : ''} ${Store.supportsA ? '' : 'disabled'}>A · 폴더 직접 저장${Store.supportsA ? '' : ' (이 브라우저 미지원)'}</option>
        <option value="B" ${c.저장방식 === 'B' ? 'selected' : ''}>B · 브라우저 저장 + 내보내기</option></select>
      <span>저장 폴더</span><span><button class="btn sm" id="btnPickFolder" ${Store.supportsA && Store.mode === 'A' ? '' : 'disabled'}>폴더 지정/변경</button></span>
    </div><p class="hint">A: Chrome/Edge 최신 버전에서 공유 폴더(예: AAR_데이터)에 훈련/여단/대대/중대/소대 폴더로 직접 저장합니다.<br>B: 브라우저 내부에 저장하며, 목록의 [내보내기]/[훈련 불러오기]로 다른 PC와 옮깁니다.</p></div>

    <div class="card"><h3>진단 기준값</h3><div class="kv">
      <span>판정 기준 N (%)</span><input type="number" id="cfgN" min="1" max="100" value="${c.판정기준N}">
      <span>보고 지연 기준 (초)</span><input type="number" id="cfgDelay" min="1" value="${c.보고지연기준초}">
      <span>교전 구분 간격 (분)</span><input type="number" id="cfgGap" min="1" value="${c.교전간격분}">
      <span>초기 피해 구간 (분)</span><input type="number" id="cfgEarly" min="1" value="${c.초기피해구간분}">
      <span>시간대 묶음 단위 (분)</span><input type="number" id="cfgBin" min="5" value="${c.시간대단위분}">
      <span>분석기준버전</span><input id="cfgVer" value="${U.esc(c.분석기준버전)}">
    </div><p class="hint">판정 기준 N은 화면에 바로 반영됩니다. 나머지 값과 지표 정의를 바꾸면 분석기준버전을 올리고, 기록 목록에서 [재분석]해야 저장된 훈련에 반영됩니다.</p></div>

    <div class="card full"><h3>열 매칭</h3><p class="hint">실제 엑셀의 시트 이름·열 이름이 기본 양식과 다르면 실제 이름을 입력하세요. 띄어쓰기와 대소문자는 무시합니다. <span class="err">*</span> 필수 열</p>
      <div class="map-sheets">${Object.entries(SHEET_DEFS).map(([s, def]) => `<div class="map-sheet"><table>
        <tr><td><b>시트: ${s}</b></td><td><input data-sheet="${s}" value="${U.esc(c.sheets[s].sheetName)}"></td></tr>
        ${Object.entries(def.cols).map(([col, cd]) => `<tr><td class="${cd.req ? 'req' : ''}">${U.esc(col)}</td><td><input data-sheet="${s}" data-col="${U.esc(col)}" value="${U.esc(c.sheets[s].columns[col] || col)}"></td></tr>`).join('')}
      </table></div>`).join('')}</div></div>

    <div class="card full"><h3>지표 정의</h3><p class="hint">계산식은 고정이며, 이름·좋아지는 방향·판정 포함 여부를 바꿀 수 있습니다. (참고 지표는 판정·종합지수에서 제외)</p>
      <table class="vtable"><thead><tr><th>기능</th><th>지표 ID</th><th>이름</th><th>단위</th><th>좋아지는 방향</th><th>판정 포함</th><th>원천 시트</th></tr></thead><tbody>
      ${c.indicators.map(i => `<tr><td>${i.fn}</td><td class="hint">${i.id}</td><td><input data-ind="${i.id}" data-k="name" value="${U.esc(i.name)}" style="width:100%"></td><td>${U.esc(i.unit)}</td>
        <td><select data-ind="${i.id}" data-k="dir"><option value="high" ${i.dir === 'high' ? 'selected' : ''}>높을수록</option><option value="low" ${i.dir === 'low' ? 'selected' : ''}>낮을수록</option></select></td>
        <td><input type="checkbox" data-ind="${i.id}" data-k="judged" ${i.judged !== false ? 'checked' : ''}></td><td>${i.src}</td></tr>`).join('')}
      </tbody></table></div>
  </div>`;
  $('#btnPickFolder').onclick = async () => {
    try { await Store.pickFolder(); updateStoreBadge(); await refreshList(); renderSettings(); U.toast('저장 폴더를 지정했습니다.', 'ok'); }
    catch (e) { if (e.name !== 'AbortError') U.toast('폴더 지정 실패: ' + e.message, 'err'); }
  };
}

function collectSettings() {
  const c = JSON.parse(JSON.stringify(Config.cur));
  const num = (id, d) => { const v = Number($(id).value); return isFinite(v) && v > 0 ? v : d; };
  c.저장방식 = $('#cfgMode').value;
  c.판정기준N = num('#cfgN', 15);
  c.보고지연기준초 = num('#cfgDelay', 180);
  c.교전간격분 = num('#cfgGap', 5);
  c.초기피해구간분 = num('#cfgEarly', 30);
  c.시간대단위분 = num('#cfgBin', 30);
  c.분석기준버전 = $('#cfgVer').value.trim() || '1.0';
  $$('[data-sheet]').forEach(el => {
    const s = el.dataset.sheet;
    if (el.dataset.col) c.sheets[s].columns[el.dataset.col] = el.value.trim() || el.dataset.col;
    else c.sheets[s].sheetName = el.value.trim() || s;
  });
  $$('[data-ind]').forEach(el => {
    const i = c.indicators.find(x => x.id === el.dataset.ind);
    if (el.dataset.k === 'judged') i.judged = el.checked; else i[el.dataset.k] = el.value;
  });
  return c;
}

async function applyConfig(c) {
  const modeChanged = c.저장방식 !== Config.cur.저장방식;
  Config.save(c);
  if (modeChanged) { await Store.init(Config.cur.저장방식); updateStoreBadge(); await refreshList(); }
  renderSettings();
}

function bindSettings() {
  $('#btnCfgSave').onclick = async () => { await applyConfig(collectSettings()); U.toast('설정을 저장했습니다.', 'ok'); };
  $('#btnCfgReset').onclick = async () => {
    if (!confirm('모든 설정을 기본값으로 되돌릴까요?')) return;
    const mode = Config.cur.저장방식;
    Config.reset();
    Config.cur.저장방식 = mode;
    await applyConfig(Config.cur);
    U.toast('기본값으로 복원했습니다.');
  };
  $('#btnCfgExport').onclick = () => U.download(new Blob([JSON.stringify(collectSettings(), null, 2)], { type: 'application/json' }), `AAR_설정_v${Config.cur.분석기준버전}.json`);
  $('#btnCfgImport').onclick = () => $('#fileCfg').click();
  $('#fileCfg').onchange = async e => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    try { await applyConfig(JSON.parse(await f.text())); U.toast('설정을 불러왔습니다.', 'ok'); }
    catch (err) { U.toast('설정 파일을 읽지 못했습니다: ' + err.message, 'err'); }
  };
}

/* ======================= 시작 ======================= */
async function init() {
  Config.load();
  Charts.setup();
  await Store.init(Config.cur.저장방식);
  updateStoreBadge();
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-go]');
    if (b && !b.disabled) go(b.dataset.go);
  });
  $('#btnReconnect').onclick = async () => {
    try { await Store.reconnect(); updateStoreBadge(); await refreshList(); go(S.screen); }
    catch (e) { U.toast(e.message, 'err'); }
  };
  bindList(); bindUpload(); bindDash(); bindSettings();
  window.addEventListener('beforeunload', e => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });
  await refreshList();
  go('list');
}

/* 자동 점검용 진입점 (화면 조작 없이 기능 확인) */
window.AAR = { S, Engine, Store, Repo, Config, U, setUploadFile, runParse, runAnalyze, saveCurrent, openTraining, renderCompare, savePanelPng, saveAllPng, enterPresent, exitPresent, presentStep, go, drill, refreshList };

init().catch(e => { console.error(e); U.toast('초기화 실패: ' + e.message, 'err'); });
