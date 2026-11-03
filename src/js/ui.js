'use strict';
/* 화면: 파일 넣기 → 자동 분석 → 시각화 → 브리핑 자료(.pptx) 다운로드. 분석한 훈련은 다음 비교 기준으로 쓰도록 브라우저에 요약만 남긴다. */

const $ = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

const S = { history: [], cur: null, R: null, charts: [] };
const HIST_KEY = 'history.v2';

/* ---------- 이전 훈련 기록 (비교 기준용 요약) ---------- */
const byDate = (a, b) => String(a.시작일).localeCompare(String(b.시작일));
async function loadHistory() {
  try { S.history = (await U.idbGet('meta', HIST_KEY)) || []; } catch (e) { S.history = []; }
  if (!Cloud.on()) return;
  // 공유 저장소 기록을 우선하고, 이 브라우저에만 있던 기록은 올려서 합친다
  try {
    const remote = await Cloud.list();
    const ids = new Set(remote.map(h => h.id));
    const localOnly = S.history.filter(h => !ids.has(h.id));
    await Cloud.save(localOnly);
    S.history = remote.concat(localOnly).sort(byDate);
    Cloud.state = 'ok';
    try { await U.idbPut('meta', HIST_KEY, S.history); } catch (e) { /* 무시 */ }
  } catch (e) { console.error(e); Cloud.state = 'err'; }
}
async function saveHistory(changed) {
  try { await U.idbPut('meta', HIST_KEY, S.history); } catch (e) { /* 저장 불가 환경: 이번 세션에서만 유지 */ }
  if (!Cloud.on() || !changed || !changed.length) return;
  try { await Cloud.save(changed); Cloud.state = 'ok'; }
  catch (e) { console.error(e); Cloud.state = 'err'; U.toast('공유 저장소(Supabase)에 저장하지 못했습니다. 이 브라우저에만 저장했습니다.', 'err'); }
}
function remember(info, model) {
  const h = { id: info.id, 훈련명: info.훈련명, 시작일: info.시작일, 종료일: info.종료일, 훈련부대: info.훈련부대, 원천파일명: info.원천파일명, levelStats: model.levelStats, 분석일: U.nowStr() };
  S.history = S.history.filter(x => x.id !== info.id);
  S.history.push(h);
  S.history.sort(byDate);
  return h;
}

function cloudNote() {
  if (Cloud.state === 'ok') return '<span class="hint">☁ 공유 저장소(Supabase)와 동기화됨</span>';
  if (Cloud.state === 'err') return '<span class="hint">⚠ 공유 저장소(Supabase)에 연결하지 못해 이 브라우저 기록만 씁니다</span>';
  return '';
}

function renderHistory() {
  const box = $('#histBox');
  if (!S.history.length) { box.innerHTML = cloudNote() ? `<div class="hist">${cloudNote()}</div>` : ''; return; }
  box.innerHTML = `<div class="hist"><b>비교 기준으로 쓰는 이전 훈련 ${S.history.length}건</b> · <a id="histClear">기록 지우기</a> ${cloudNote()}
    <ul>${S.history.slice().reverse().map(h => `<li>${U.esc(h.훈련명)} <span class="hint">(${U.esc(h.시작일)})</span></li>`).join('')}</ul></div>`;
  $('#histClear').onclick = async () => {
    const shared = Cloud.state === 'ok' ? ' 공유 저장소(Supabase)의 기록도 함께 지워져 다른 사용자에게도 사라집니다.' : '';
    if (!confirm(`이전 훈련 기록을 모두 지울까요?${shared} 이후 분석은 대항군과의 피아 비교로 판정합니다.`)) return;
    if (Cloud.state === 'ok') {
      try { await Cloud.clear(); } catch (e) { console.error(e); U.toast('공유 저장소 기록을 지우지 못했습니다.', 'err'); return; }
    }
    S.history = []; await saveHistory(); renderHistory();
  };
}

/* ---------- 파일 → 분석 ---------- */
function infoFromName(name) {
  const base = name.replace(/\.xlsx?$/i, '').replace(/^샘플\d+_/, '');
  const d = base.match(/^(\d{4}-\d{2}(?:-\d{2})?)_(.+)$/);
  return { 훈련명: (d ? d[2] : base).trim() || '이름없는 훈련', 원천파일명: name };
}

async function analyzeFile(file) {
  const info = infoFromName(file.name);
  const parsed = Engine.parse(Engine.readWorkbook(await file.arrayBuffer()), Config.cur, null);
  if (!Object.values(parsed.tables).some(t => t.length)) {
    throw new Error(`${file.name}: 분석할 수 있는 시트가 없습니다. (필요 시트: ${Object.keys(SHEET_DEFS).join(', ')})`);
  }
  const model = Engine.analyze(parsed.tables, Config.cur, info);
  info.시작일 = U.dateStr(model.tStart);
  info.종료일 = U.dateStr(model.tEnd);
  info.id = info.훈련명 + '|' + info.시작일;
  info.훈련부대 = (model.nodes['ROOT|훈련부대'] ? model.nodes['ROOT|훈련부대'].children : []).map(id => model.nodes[id].name).join(', ');
  info.제외행수 = Object.values(parsed.errorRows).reduce((s, x) => s + x.size, 0);
  info.누락 = parsed.missingSheets.concat(parsed.missingCols.map(m => `${m.sheet}(${m.cols.join(', ')} 열 없음)`));
  info.행수 = {};
  for (const k of Object.keys(SHEET_DEFS)) info.행수[k] = parsed.tables[k].length;
  return { info, model };
}

async function handleFiles(fileList) {
  const files = Array.from(fileList).filter(f => /\.xlsx?$/i.test(f.name));
  if (!files.length) { U.toast('엑셀 파일(.xlsx)을 넣어 주세요.', 'err'); return; }
  const done = [];
  for (let i = 0; i < files.length; i++) {
    U.busy(files.length > 1 ? `분석 중… (${i + 1}/${files.length}) ${files[i].name}` : '분석 중…');
    await U.nextFrame();
    try { done.push(await analyzeFile(files[i])); }
    catch (e) { console.error(e); U.toast(e.message || `${files[i].name}: 읽지 못했습니다.`, 'err'); }
  }
  U.busy();
  if (!done.length) return;
  // 날짜순으로 기록에 넣고 가장 최근 훈련을 보여 준다
  done.sort((a, b) => String(a.info.시작일).localeCompare(String(b.info.시작일)));
  const changed = done.map(r => remember(r.info, r.model));
  await saveHistory(changed);
  show(done[done.length - 1]);
  U.toast(done.length > 1 ? `${done.length}개 훈련을 분석했습니다. 가장 최근 훈련을 표시합니다.` : '분석 완료', 'ok');
}

/* ---------- 결과 화면 ---------- */
function destroyCharts() { S.charts.forEach(c => c.destroy()); S.charts = []; }
const badge = j => `<span class="badge ${['강점', '취약', '보통', '참고', '우수', '미흡'].includes(j) ? j : 'none'}">${j === '강점' || j === '우수' ? '▲ ' : j === '취약' || j === '미흡' ? '▼ ' : ''}${U.esc(j)}</span>`;
const stripFn = s => s.replace(/^\[[^\]]+\]\s*/, '');

function show(cur) {
  S.cur = cur;
  S.R = Insight.build(cur, S.history);
  $('#home').hidden = true;
  $('#dash').hidden = false;
  $('#btnReport').hidden = false;
  renderDash();
  window.scrollTo(0, 0);
}

function goHome() {
  destroyCharts();
  S.cur = null; S.R = null;
  $('#dash').hidden = true; $('#dash').innerHTML = '';
  $('#home').hidden = false;
  $('#btnReport').hidden = true;
  renderHistory();
}

function renderDash() {
  const R = S.R, { info, root, opfor, diag } = R;
  const m = root.metrics;
  const nStrong = diag.functions.filter(f => f.judge === '강점').length;
  const nWeak = diag.functions.filter(f => f.judge === '취약').length;
  const notices = [];
  if (info.제외행수) notices.push(`형식 오류 ${info.제외행수.toLocaleString()}행은 자동으로 제외했습니다.`);
  if (info.누락.length) notices.push(`누락 자료: ${info.누락.join(', ')} → 관련 지표는 "자료 없음"으로 표시합니다.`);
  if (R.baseKind === 'opfor') notices.push('이전 훈련 기록이 없어 대항군 지표와 비교(피아 비교)해 판정했습니다. 이전 훈련 파일을 함께 넣으면 훈련 간 비교로 판정합니다.');

  const kpi = (k, v, d) => `<div class="kpi"><div class="k">${k}</div><div class="v">${v}</div><div class="d">${d || ''}</div></div>`;
  const vs = b => (opfor ? `<small>대 ${b}</small>` : '');

  const fnTiles = diag.functions.map(f => `
    <div class="fn ${f.judge}">
      <div class="h"><span class="name">${f.fn}</span>${badge(f.judge)}<span class="idx">${f.index == null ? '–' : f.index.toFixed(0)}</span></div>
      <div class="why">${U.esc(stripFn(f.summary).replace(/^종합지수 \d+ \([^)]*\) → [^.]+\.\s*/, ''))}</div>
      <details><summary>지표 ${f.indicators.length}개 보기</summary>
        <table class="itable"><thead><tr><th>지표</th><th>이번</th><th>${U.esc(R.baseName)}</th><th>판정</th></tr></thead><tbody>
        ${f.indicators.map(r => { const it = Config.ind(r.id); return `<tr><td>${U.esc(r.name)}${it.dir === 'low' ? ' <span class="hint">↓</span>' : ''}</td><td>${U.fmtInd(it, r.value)}</td><td>${r.base == null ? '–' : U.fmtInd(it, r.base)}</td><td>${badge(r.judge)}</td></tr>`; }).join('')}
        </tbody></table></details>
    </div>`).join('');

  const odrList = (list, kind) => list.length ? list.map(x => `
    <div class="odr"><div class="t">${U.esc(x.title)} ${x.soft ? '<span class="soft">(기준 미달 · 상대 비교)</span>' : ''}</div>
      <dl><dt>관찰</dt><dd>${U.esc(x.obs)}</dd><dt>논의</dt><dd>${U.esc(x.disc)}</dd><dt>건의</dt><dd>${U.esc(x.rec)}</dd></dl></div>`).join('')
    : `<p class="hint">${kind === 's' ? '유지할 점으로 판정된 지표가 없습니다.' : '개선할 점으로 판정된 지표가 없습니다.'}</p>`;

  const unitRows = R.units.map(u => {
    const um = u.node.metrics;
    return `<tr><td><b>${U.esc(u.name)}</b></td><td>${U.fmt(um.fire_hit, 'pct')}</td><td>${U.fmt(um.prot_survival, 'pct')}</td><td>${U.fmt(um.int_predetect, 'pct')}</td>
      <td>${U.fmt(um.c2_report_delay, 'pct')}</td><td>${u.node.personnel}</td><td><b>${u.score == null ? '–' : u.score.toFixed(0)}</b></td><td>${badge(u.judge)}</td><td>${U.esc(u.weakFns.join(', ') || '–')}</td></tr>`;
  }).join('');

  const chartCard = (id, title, sub, cls) => `<div class="card"><h3>${title}<span class="sub">${sub || ''}</span></h3><div class="chart-box ${cls || ''}"><canvas data-chart="${id}"></canvas></div></div>`;

  $('#dash').innerHTML = `
    <div class="hero">
      <div>
        <div class="eyebrow">사후검토(AAR) 분석 결과</div>
        <h1>${U.esc(info.훈련명)}</h1>
        <div class="meta"><span>📅 ${U.esc(info.시작일)}${info.종료일 !== info.시작일 ? ' ~ ' + U.esc(info.종료일) : ''}</span><span>🪖 ${U.esc(R.sideName)}</span><span>📄 ${U.esc(info.원천파일명 || '')}</span><span>⚖ 비교 기준: ${U.esc(R.baseText)}</span></div>
      </div>
      <div class="tally">
        <div class="t good"><b>${nStrong}</b><span>강점 기능</span></div>
        <div class="t bad"><b>${nWeak}</b><span>취약 기능</span></div>
        ${R.combat ? `<div class="t"><b style="font-size:18px;line-height:30px">${R.combat}</b><span>피아 교전</span></div>` : ''}
      </div>
    </div>
    ${notices.length ? `<div class="notice">${notices.map(U.esc).join('<br>')}</div>` : ''}
    <div class="summary"><b>종합 판단</b> · ${U.esc(R.overall)}</div>

    <div class="kpis">
      ${kpi('교전 수', U.fmt(root.engagements, 'int') + vs(U.fmt(opfor && opfor.engagements, 'int')), `사격 ${root.shots.toLocaleString()}회`)}
      ${kpi('명중률', U.fmt(m.fire_hit, 'pct') + vs(U.fmt(opfor && opfor.metrics.fire_hit, 'pct')), `명중 ${root.hits.toLocaleString()}회`)}
      ${kpi('피해 유발', U.fmt(root.dmgHits, 'int') + vs(U.fmt(opfor && opfor.dmgHits, 'int')), '상대에게 준 피해(건)')}
      ${kpi('생존율', U.fmt(m.prot_survival, 'pct') + vs(U.fmt(opfor && opfor.metrics.prot_survival, 'pct')), `사망 ${root.cas.사망} · 중상 ${root.cas.중상} · 경상 ${root.cas.경상}`)}
      ${kpi('손실교환비', R.exch == null ? '–' : R.exch.toFixed(2), '대항군 사망 ÷ 아군 사망')}
      ${kpi('인원', U.fmt(root.personnel, 'int'), Object.entries(R.org).filter(([, v]) => v).map(([k, v]) => `${k} ${v}`).join(' · '))}
    </div>

    <div class="sec"><h2>전투수행기능 진단</h2><span class="sub">${U.esc(R.baseName)} = 100 · 개선율 ±${Config.cur.판정기준N}% 기준</span></div>
    <div class="grid-r">
      ${diag.hasBase ? chartCard('radar', '6대 기능 종합지수', '', 'tall') : '<div class="card"><h3>6대 기능 종합지수</h3><p class="hint">비교 기준이 없어 지수를 계산하지 않았습니다.</p></div>'}
      <div class="fn-grid">${fnTiles}</div>
    </div>

    <div class="sec"><h2>유지할 점 · 개선할 점</h2><span class="sub">관찰 → 논의 → 건의</span></div>
    <div class="grid2">
      <div class="card odr-col s"><h3>${badge('강점')} 유지할 점 (Sustain)</h3>${odrList(R.sustain, 's')}</div>
      <div class="card odr-col i"><h3>${badge('취약')} 개선할 점 (Improve)</h3>${odrList(R.improve, 'i')}</div>
    </div>

    <div class="sec"><h2>피아 교전 결과</h2></div>
    <div class="grid2">
      ${chartCard('versus', '핵심 지표 비교', opfor ? `${R.sideName} 대 대항군` : '')}
      ${root.tsDmg.length > 1 ? chartCard('timeline', '시간대별 피해·교전', `${R.model.binMin}분 단위${R.peak ? ` · 피해 최다 ${R.peak.from}~${R.peak.to}` : ''}`) : ''}
    </div>

    ${R.units.length ? `
    <div class="sec"><h2>제대별 분석</h2><span class="sub">${R.lvName} · 동일 제대 평균 = 100</span></div>
    <div class="grid2">
      ${chartCard('unitScore', `${R.lvName}별 종합지수`, '초록 우수 · 빨강 미흡')}
      ${chartCard('casualty', `${R.lvName}별 피해 현황`, '명')}
    </div>
    <div class="card scroll-x" style="margin-top:12px"><table class="utable"><thead><tr><th>제대</th><th>명중률</th><th>생존율</th><th>선탐지</th><th>보고지연</th><th>인원</th><th>종합지수</th><th>판정</th><th>취약 기능</th></tr></thead><tbody>${unitRows}</tbody></table></div>` : ''}

    <div class="sec"><h2>화기 · 훈련 간 추이</h2></div>
    <div class="grid2">
      ${chartCard('weapons', '화기별 명중률', '막대 명중률 · 선 사격 수')}
      ${R.trend.length > 1 ? chartCard('trend', '훈련 간 추이', R.trend.map(t => t.name).join(' → ')) : '<div class="card"><h3>훈련 간 추이</h3><p class="hint">이전 훈련 파일을 함께 넣으면 훈련별 명중률·생존율·선탐지 추이를 보여 줍니다.</p></div>'}
    </div>

    <div class="foot">
      <button class="btn primary" id="btnUpload2">📂 다른 엑셀 데이터 넣기</button>
      <button class="btn" id="btnReport2">브리핑 PPT 받기 (.pptx)</button>
      <span class="hint">파일을 화면 아무 곳에나 끌어 놓아도 새로 분석합니다</span>
    </div>`;

  destroyCharts();
  $$('canvas[data-chart]').forEach(cv => S.charts.push(new Chart(cv, Charts.defs[cv.dataset.chart](R, 1))));
  $('#btnReport2').onclick = downloadReport;
  $('#btnUpload2').onclick = pickFile;
}

const pickFile = () => $('#fileXlsx').click();


async function downloadReport() {
  if (!S.R) return;
  U.busy('브리핑 자료(PPT) 만드는 중…');
  await U.nextFrame();
  try {
    const blob = await Deck.build(S.R);
    U.download(blob, Deck.fileName(S.R));
    U.toast('브리핑 자료(PPT)를 다운로드했습니다.', 'ok');
  } catch (e) { console.error(e); U.toast('PPT 생성 실패: ' + e.message, 'err'); }
  U.busy();
}

/* ---------- 시작 ---------- */
function bindDrop() {
  const drop = $('#drop'), cover = $('#dropCover');
  $('#fileXlsx').addEventListener('change', e => { if (e.target.files.length) handleFiles(e.target.files); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, () => drop.classList.remove('over')));
  // 결과 화면에서도 아무 곳에나 끌어 놓으면 새로 분석
  let depth = 0;
  const hasFiles = e => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  window.addEventListener('dragenter', e => { if (!hasFiles(e)) return; e.preventDefault(); depth++; if (!$('#dash').hidden) cover.hidden = false; });
  window.addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; cover.hidden = true; } });
  window.addEventListener('drop', e => {
    if (!hasFiles(e)) return;
    e.preventDefault(); depth = 0; cover.hidden = true;
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });
}

async function init() {
  Config.load();
  Charts.setup();
  bindDrop();
  $('.brand').onclick = goHome;
  $('#btnReport').onclick = downloadReport;
  $('#btnUpload').onclick = pickFile;
  $('#ctaUpload').onclick = pickFile;
  $('#sampleLink').onclick = e => {
    e.preventDefault();
    try { U.download(Sample.build(), Sample.fileName); } catch (err) { console.error(err); U.toast('샘플 생성 실패: ' + err.message, 'err'); }
  };
  // 배포 사이트에 함께 올린 샘플 엑셀(samples/)을 불러와 바로 분석한다
  const DEMO = ['샘플1_2026-06_백두1호훈련.xlsx', '샘플2_2026-08_백두2호훈련.xlsx', '샘플3_2026-10_백두3호훈련.xlsx'];
  document.querySelectorAll('[data-demo]').forEach(b => b.onclick = async () => {
    const names = b.dataset.demo === 'all' ? DEMO : [DEMO[+b.dataset.demo]];
    try {
      U.busy('샘플 불러오는 중…');
      const files = await Promise.all(names.map(async n => {
        const r = await fetch('samples/' + encodeURIComponent(n));
        if (!r.ok) throw new Error(n + ' (' + r.status + ')');
        return new File([await r.blob()], n);
      }));
      U.busy();
      await handleFiles(files);
    } catch (err) {
      U.busy(); console.error(err);
      U.toast('샘플을 불러오지 못했습니다. 아래 링크로 샘플을 받아 직접 넣어 주세요. (' + err.message + ')', 'err');
    }
  });
  await loadHistory();
  renderHistory();
}

/* 자동 점검용 진입점 */
window.AAR = { S, Engine, Insight, Deck, Sample, Config, U, Cloud, loadHistory, renderHistory, handleFiles, analyzeFile, show, goHome, downloadReport, pickFile };

init().catch(e => { console.error(e); U.toast('초기화 실패: ' + e.message, 'err'); });
