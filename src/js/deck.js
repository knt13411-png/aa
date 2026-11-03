'use strict';
/* 사후검토(AAR) 결과 브리핑 (.pptx, 16:9)
   구성: 표지 → 목차 → 보고 요지 → 훈련 개요 → 피아 교전 결과 → 시간대별·훈련 간 추이 → 6대 기능 종합 → 기능별 지표
        → 제대별 분석 → 유지할 점 → 개선할 점 → 건의 및 발전방향 → 붙임(지표 산출 기준)
   한국군 브리핑의 □ ○ - 위계와 미 육군 AAR(TC 25-20)의 유지할 점/개선할 점·관찰-논의-건의 형식을 따른다. */

const Deck = {};
const NAVY = '1F3864', GOOD = '0A7A0A', BAD = 'C03030', GREY = '7F7F7F', LIGHT = 'F2F5FA';

Deck.fileName = R => `${U.safeName(R.info.훈련명)}_사후검토결과_브리핑_${U.nowStr().slice(0, 10)}.pptx`;

Deck.build = async function (R) {
  const { info, root, opfor, diag } = R;
  const title = `${info.훈련명} 사후검토(AAR) 결과`;
  const d = new Pptx.Deck(title);
  const W = d.W;
  const ind = id => Config.ind(id);
  const jColor = j => (j === '강점' || j === '우수' ? GOOD : j === '취약' || j === '미흡' ? BAD : null);
  const jc = j => ({ t: j, b: !!jColor(j), color: jColor(j), align: 'ctr' });
  const strip = s => s.replace(/^\[[^\]]+\]\s*/, '').replace(/^종합지수 \d+ \([^)]*\) → [^.]+\.\s*/, '') || '–';
  // 차트 이미지: 슬라이드에서 1인치 ≈ 160px, 글자 배율 1.7 (약 9pt)
  const chart = (def, wIn, hIn) => Charts.png(Charts.defs[def](R, 1.7), Math.round(wIn * 160), Math.round(hIn * 160));

  let page = 0;
  const sections = ['훈련 개요', '훈련 결과', '전투수행기능별 분석', '제대별 분석', '관찰 · 건의'];
  /* 본문 슬라이드 공통 머리·바닥 */
  const slide = (heading, sec) => {
    const s = d.slide();
    page++;
    s.rect(0, 0, W, 0.12, { fill: NAVY });
    s.rect(0.45, 0.42, 0.09, 0.5, { fill: NAVY });
    s.text(0.62, 0.32, 9.6, 0.7, { t: heading, b: true, sz: 24, color: NAVY }, { anchor: 'ctr' });
    if (sec) s.text(9.9, 0.4, 3, 0.5, { t: sec, sz: 11, color: GREY, align: 'r' }, { anchor: 'ctr' });
    s.rect(0.45, 1.05, W - 0.9, 0.025, { fill: NAVY });
    s.text(0.45, 7.05, 8, 0.32, { t: `${title} 브리핑`, sz: 9, color: GREY }, { anchor: 'ctr' });
    s.text(W - 1.45, 7.05, 1, 0.32, { t: String(page), sz: 10, color: GREY, align: 'r' }, { anchor: 'ctr' });
    return s;
  };
  const box = (s, x, y, w, h, head, color) => {
    s.rect(x, y, w, h, { fill: 'FFFFFF', line: 'BFBFBF', lw: 0.75 });
    s.rect(x, y, w, 0.42, { fill: color || NAVY });
    s.text(x + 0.1, y, w - 0.2, 0.42, { t: head, b: true, sz: 13, color: 'FFFFFF' }, { anchor: 'ctr' });
  };

  /* ---------- 1. 표지 ---------- */
  {
    const s = d.slide();
    page++;
    s.rect(0, 0, W, 7.5, { fill: NAVY });
    s.rect(0, 4.55, W, 0.05, { fill: 'C9A227' });
    s.text(0.9, 1.6, 11.5, 0.5, { t: '사후검토(After Action Review) 결과 브리핑', sz: 16, color: 'C9D4E4' });
    s.text(0.9, 2.2, 11.5, 1.4, { t: info.훈련명, b: true, sz: 44, color: 'FFFFFF' }, { anchor: 'ctr' });
    s.text(0.9, 3.55, 11.5, 0.7, { t: `훈련부대 ${info.훈련부대 || R.sideName}  ·  ${info.시작일}${info.종료일 !== info.시작일 ? ' ~ ' + info.종료일 : ''}`, sz: 16, color: 'E6ECF5' });
    s.text(0.9, 4.9, 11.5, 0.5, { t: `보고일자 ${U.nowStr().slice(0, 10)}`, sz: 14, color: 'C9D4E4' });
    s.text(0.9, 5.35, 11.5, 0.5, { t: `비교 기준: ${R.baseText}`, sz: 12, color: 'A9BBD3' });
    s.text(0.9, 6.6, 11.5, 0.4, { t: '본 자료는 훈련 데이터를 자동 분석한 초안이며, 담당자 검토 후 활용', sz: 10, color: '8FA3BF' });
  }

  /* ---------- 2. 목차 ---------- */
  {
    const s = slide('목 차');
    sections.forEach((t, i) => {
      const y = 1.6 + i * 0.95;
      s.rect(3.2, y, 0.75, 0.62, { fill: NAVY, round: 12000 });
      s.text(3.2, y, 0.75, 0.62, { t: String(i + 1), b: true, sz: 20, color: 'FFFFFF', align: 'ctr' }, { anchor: 'ctr' });
      s.text(4.15, y, 6, 0.62, { t, b: true, sz: 20, color: '1A1A1A' }, { anchor: 'ctr' });
      s.rect(4.15, y + 0.68, 6, 0.012, { fill: 'D9D9D9' });
    });
  }

  /* ---------- 3. 보고 요지 ---------- */
  {
    const s = slide('보고 요지', '종합');
    s.text(0.45, 1.3, W - 0.9, 1.45, [{ runs: [{ t: '종합 판단  ', b: true, color: NAVY }, R.overall] }], { sz: 15, fill: LIGHT, border: 'C5D3EA', inset: 0.18, anchor: 'ctr', lineSp: 125 });
    const nS = diag.functions.filter(f => f.judge === '강점').length, nW = diag.functions.filter(f => f.judge === '취약').length;
    const tally = [['강점 기능', String(nS), GOOD], ['취약 기능', String(nW), BAD], ['피아 교전', R.combat || '–', NAVY],
      ['손실교환비', R.exch == null ? '–' : R.exch.toFixed(2), NAVY], ['명중률', U.fmt(root.metrics.fire_hit, 'pct'), NAVY], ['생존율', U.fmt(root.metrics.prot_survival, 'pct'), NAVY]];
    const tw = (W - 0.9 - 0.2 * 5) / 6;
    tally.forEach(([k, v, c], i) => {
      const x = 0.45 + i * (tw + 0.2);
      s.rect(x, 2.95, tw, 1.05, { fill: 'FFFFFF', line: 'BFBFBF', lw: 0.75 });
      s.text(x, 3.0, tw, 0.32, { t: k, sz: 11, color: GREY, align: 'ctr' });
      s.text(x, 3.3, tw, 0.62, { t: v, b: true, sz: 24, color: c, align: 'ctr' }, { anchor: 'ctr' });
    });
    const half = (W - 0.9 - 0.3) / 2;
    const list = arr => (arr.length ? arr.slice(0, 4).map(x => ({ lvl: 1, runs: [{ t: x.fn + ' · ', b: true }, x.ind.name] })) : [{ t: '해당 없음', color: GREY }]);
    box(s, 0.45, 4.2, half, 2.65, '유지할 점 (Sustain)', GOOD);
    s.text(0.55, 4.7, half - 0.2, 2.1, list(R.sustain), { sz: 14 });
    box(s, 0.45 + half + 0.3, 4.2, half, 2.65, '개선할 점 (Improve)', BAD);
    s.text(0.55 + half + 0.3, 4.7, half - 0.2, 2.1, list(R.improve), { sz: 14 });
  }

  /* ---------- 4. 훈련 개요 ---------- */
  {
    const s = slide('1. 훈련 개요', sections[0]);
    const orgTxt = Object.entries(R.org).filter(([, v]) => v).map(([k, v]) => `${k} ${v}개`).join(' · ');
    const paras = [
      { lvl: 0, runs: [{ t: '목 적 : ', b: true }, '과학화 훈련 데이터를 근거로 6대 전투수행기능별 성과를 진단하고 유지할 점·개선할 점 도출'] },
      { lvl: 0, runs: [{ t: '기 간 : ', b: true }, `${info.시작일} ~ ${info.종료일} (교전 기록 ${U.timeStr(R.model.tStart)} ~ ${U.timeStr(R.model.tEnd)})`] },
      { lvl: 0, runs: [{ t: '참가부대', b: true }] },
      { lvl: 1, t: `${R.sideName}: ${orgTxt}, 인원 ${root.personnel.toLocaleString()}명` },
    ];
    if (opfor) paras.push({ lvl: 1, t: `대항군: 인원 ${opfor.personnel.toLocaleString()}명` });
    paras.push({ lvl: 0, runs: [{ t: '분석 방법', b: true }] });
    paras.push({ lvl: 1, t: `비교 기준: ${R.baseText}` });
    paras.push({ lvl: 1, t: `판정: 비교 기준 대비 개선율 +${Config.cur.판정기준N}% 이상 강점, −${Config.cur.판정기준N}% 이하 취약` });
    if (info.제외행수) paras.push({ lvl: 1, t: `형식 오류 ${info.제외행수.toLocaleString()}행은 분석에서 제외` });
    if (info.누락 && info.누락.length) paras.push({ lvl: 1, t: `누락 자료: ${info.누락.join(', ')} → 관련 지표 "자료 없음"` });
    s.text(0.45, 1.35, 7.4, 5.5, paras, { sz: 15, after: 8, lineSp: 115 });
    s.text(8.2, 1.35, 4.7, 0.4, { t: '분석 자료', b: true, sz: 13, color: NAVY });
    s.table(8.2, 1.8, 4.7, [['시트', '기록 수']].concat(Object.keys(SHEET_DEFS).map(k => [k, { t: ((info.행수 && info.행수[k]) || 0).toLocaleString() + '건', align: 'r' }])),
      { widths: [1.6, 1.2], rowH: 0.4, sz: 12 });
    s.text(8.2, 4.75, 4.7, 0.6, { t: `원천 파일: ${info.원천파일명 || '–'}`, sz: 10, color: GREY });
  }

  /* ---------- 5. 피아 교전 결과 ---------- */
  {
    const s = slide('2. 훈련 결과 — 피아 교전 결과', sections[1]);
    const vsRow = (name, f) => [{ t: name, b: true }, { t: f(root), align: 'ctr' }].concat(opfor ? [{ t: f(opfor), align: 'ctr' }] : []);
    const rows = [
      ['구분', R.sideName].concat(opfor ? ['대항군'] : []),
      vsRow('교전 수', n => n.engagements.toLocaleString()),
      vsRow('사격 / 명중', n => `${n.shots.toLocaleString()} / ${n.hits.toLocaleString()}`),
      vsRow('명중률', n => U.fmt(n.metrics.fire_hit, 'pct')),
      vsRow('피해 유발(건)', n => n.dmgHits.toLocaleString()),
      vsRow('인원(명)', n => n.personnel.toLocaleString()),
      vsRow('사망/중상/경상', n => `${n.cas.사망} / ${n.cas.중상} / ${n.cas.경상}`),
      vsRow('생존율', n => U.fmt(n.metrics.prot_survival, 'pct')),
    ];
    if (R.exch != null) rows.push([{ t: '손실교환비', b: true }, { t: R.exch.toFixed(2), span: opfor ? 2 : 1, align: 'ctr', b: true }]);
    s.table(0.45, 1.35, 5.2, rows, { widths: opfor ? [1.6, 1.3, 1.3] : [1.6, 2], rowH: 0.48, sz: 12, firstColFill: 'F2F2F2' });
    s.image(await chart('versus', 7.1, 4.3), 5.9, 1.35, 7.0, 4.24);
    s.text(5.9, 5.7, 7.0, 1.2, [{ lvl: 0, t: R.combat ? `${R.combat}: 피해 유발 ${root.dmgHits}건 대 ${opfor ? opfor.dmgHits : '–'}건` : '피아 비교 자료 없음' },
      { lvl: 1, t: '↓ 표시 지표는 낮을수록 좋음', sz: 11, color: GREY }], { sz: 14 });
  }

  /* ---------- 6. 시간대별 추이 / 훈련 간 추이 ---------- */
  if (root.tsDmg.length > 1) {
    const s = slide('2. 훈련 결과 — 시간대별 피해·교전 추이', sections[1]);
    s.image(await chart('timeline', 12.4, 4.6), 0.45, 1.3, 12.4, 4.6);
    const notes = [];
    if (R.peak) notes.push({ lvl: 0, t: `${R.sideName} 피해는 ${R.peak.from}~${R.peak.to} 시간대에 가장 많이(${R.peak.n}건) 발생` });
    notes.push({ lvl: 1, t: `${R.model.binMin}분 단위 집계 · 점선은 교전 개시 수`, sz: 12, color: GREY });
    s.text(0.45, 6.0, 12.4, 0.95, notes, { sz: 15 });
  }
  if (R.trend.length > 1) {
    const s = slide('2. 훈련 결과 — 이전 훈련 대비 추이', sections[1]);
    s.image(await chart('trend', 8.2, 5.2), 0.45, 1.35, 8.2, 5.2);
    const first = R.trend[0], last = R.trend[R.trend.length - 1];
    const rows = [['지표', first.name, last.name, '변화']];
    ['fire_hit', 'prot_survival', 'int_predetect', 'prot_uncovered', 'man_move_dmg'].forEach(id => {
      const a = first.m[id], b = last.m[id];
      if (a == null || b == null) return;
      const good = ind(id).dir === 'low' ? b <= a : b >= a;
      rows.push([ind(id).name, { t: U.fmt(a, 'pct'), align: 'ctr' }, { t: U.fmt(b, 'pct'), align: 'ctr' }, { t: `${b - a >= 0 ? '+' : ''}${((b - a) * 100).toFixed(1)}%p`, align: 'ctr', color: good ? GOOD : BAD }]);
    });
    s.table(8.85, 1.35, 4.05, rows, { widths: [1.7, 1, 1, 0.9], rowH: 0.5, sz: 11 });
  }

  /* ---------- 7. 6대 기능 종합 ---------- */
  {
    const s = slide('3. 전투수행기능별 분석 — 종합 판정', sections[2]);
    let x0 = 0.45, tw = W - 0.9;
    if (diag.hasBase) { s.image(await chart('radar', 5.0, 4.6), 0.35, 1.35, 5.0, 4.6); x0 = 5.55; tw = W - 0.45 - x0; }
    s.table(x0, 1.35, tw, [['기능', '종합지수', '판정', '판단 근거']].concat(diag.functions.map(f => [
      { t: f.fn, b: true, align: 'ctr' }, { t: f.index == null ? '–' : f.index.toFixed(0), align: 'ctr' }, jc(f.judge), strip(f.summary),
    ])), { widths: [1.25, 0.9, 0.8, 4.4], rowH: 0.72, sz: 11 });
    s.text(0.45, 6.35, W - 0.9, 0.6, { t: `종합지수 = 판정 지표별 지수의 평균 (${R.baseText} = 100) · 개선율 ±${Config.cur.판정기준N}% 기준`, sz: 11, color: GREY });
  }

  /* ---------- 8~9. 기능별 지표 (3개 기능씩) ---------- */
  for (let k = 0; k < 2; k++) {
    const s = slide(`3. 전투수행기능별 분석 — 기능별 지표 (${k + 1}/2)`, sections[2]);
    const fns = diag.functions.slice(k * 3, k * 3 + 3);
    const cw = (W - 0.9 - 0.4) / 3;
    fns.forEach((f, i) => {
      const x = 0.45 + i * (cw + 0.2);
      s.rect(x, 1.3, cw, 0.6, { fill: jColor(f.judge) || NAVY });
      s.text(x + 0.12, 1.3, cw - 0.24, 0.6, [{ runs: [{ t: f.fn, b: true, sz: 17 }, { t: `   ${f.judge} · 지수 ${f.index == null ? '–' : f.index.toFixed(0)}`, sz: 13 }] }], { color: 'FFFFFF', anchor: 'ctr' });
      s.table(x, 2.05, cw, [['지표', '이번', R.baseName, '판정']].concat(f.indicators.map(r => {
        const it = ind(r.id);
        return [r.name + (it.dir === 'low' ? ' ↓' : ''), { t: U.fmtInd(it, r.value), align: 'ctr' }, { t: r.base == null ? '–' : U.fmtInd(it, r.base), align: 'ctr' }, jc(r.judge)];
      })), { widths: [1.9, 0.95, 0.95, 0.7], rowH: 0.55, sz: 10.5, head: '595959' });
      const notes = f.indicators.filter(r => r.change != null && r.judge !== '참고')
        .map(r => ({ lvl: 1, t: r.sentence.replace(/^\[[^\]]+\]\s*/, '').replace(/\s*→.*$/, ''), sz: 10.5, color: '404040' }));
      s.text(x, 4.45, cw, 2.45, [{ t: '판단', b: true, sz: 12, color: NAVY }, { t: strip(f.summary), sz: 12 }].concat(notes),
        { color: '262626', fill: LIGHT, inset: 0.12, lineSp: 115, after: 4 });
    });
  }

  /* ---------- 10~11. 제대별 분석 ---------- */
  if (R.units.length) {
    const s = slide(`4. 제대별 분석 — ${R.lvName}별 종합 판정`, sections[3]);
    let tx = 0.45, tw = W - 0.9;
    if (R.units.length > 1) {
      s.image(await chart('unitScore', 5.6, 4.6), 0.45, 1.35, 5.6, 4.6);
      tx = 6.3; tw = W - 0.45 - tx;
    }
    const rows = [['제대', '명중률', '생존율', '선탐지', '종합', '판정', '취약 기능']].concat(R.units.map(u => {
      const m = u.node.metrics;
      return [{ t: u.name, b: true }, { t: U.fmt(m.fire_hit, 'pct'), align: 'ctr' }, { t: U.fmt(m.prot_survival, 'pct'), align: 'ctr' }, { t: U.fmt(m.int_predetect, 'pct'), align: 'ctr' },
        { t: u.score == null ? '–' : u.score.toFixed(0), align: 'ctr', b: true }, jc(u.judge), u.weakFns.join(', ') || '–'];
    }));
    const rowH = Math.min(0.62, 4.6 / rows.length);
    s.table(tx, 1.35, tw, rows, { widths: [1.5, 0.9, 0.9, 0.9, 0.7, 0.7, 1.7], rowH, sz: rows.length > 10 ? 9 : 11 });
    const best = R.units.filter(u => u.judge === '우수'), worst = R.units.filter(u => u.judge === '미흡');
    s.text(0.45, 6.2, W - 0.9, 0.8, [
      { runs: [{ t: '우수 제대 : ', b: true, color: GOOD }, best.length ? best.map(u => `${u.name}(${u.score.toFixed(0)})`).join(', ') : '해당 없음',
        { t: '      미흡 제대 : ', b: true, color: BAD }, worst.length ? worst.map(u => `${u.name}(${u.score.toFixed(0)}${u.weakFns.length ? ', ' + u.weakFns.join('·') : ''})`).join(', ') : '해당 없음'] },
      { t: `동일 수준 ${R.lvName} 평균 = 100 기준`, sz: 11, color: GREY },
    ], { sz: 14 });

    if (R.units.length > 1) {
      const s2 = slide(`4. 제대별 분석 — ${R.lvName}별 피해 현황`, sections[3]);
      s2.image(await chart('casualty', 8.4, 5.3), 0.45, 1.3, 8.4, 5.3);
      const surv = Insight.rankBy(R.units, 'prot_survival');
      s2.text(9.1, 1.4, 3.8, 5.3, [
        { lvl: 0, t: '생존율 상위', b: true },
        ...surv.slice(0, 2).map(u => ({ lvl: 1, t: `${u.name} ${U.fmt(u.node.metrics.prot_survival, 'pct')}` })),
        { lvl: 0, t: '생존율 하위', b: true },
        ...surv.slice(-2).reverse().map(u => ({ lvl: 1, t: `${u.name} ${U.fmt(u.node.metrics.prot_survival, 'pct')}` })),
        { lvl: 0, t: '피해 합계', b: true },
        { lvl: 1, t: `사망 ${root.cas.사망} · 중상 ${root.cas.중상} · 경상 ${root.cas.경상}` },
      ], { sz: 14, after: 6 });
    }
  }

  /* ---------- 12~13. 유지할 점 / 개선할 점 (관찰 · 논의 · 건의) ---------- */
  const odrSlides = (list, label, color) => {
    const chunks = [];
    for (let i = 0; i < list.length; i += 3) chunks.push(list.slice(i, i + 3));
    if (!chunks.length) chunks.push([]);
    chunks.forEach((chunk, ci) => {
      const s = slide(`5. 주요 관찰 사항 — ${label}${chunks.length > 1 ? ` (${ci + 1}/${chunks.length})` : ''}`, sections[4]);
      if (!chunk.length) { s.text(0.45, 1.5, W - 0.9, 1, { t: '해당 사항 없음', sz: 16, color: GREY }); return; }
      const ch = 5.55 / 3 - 0.12;
      chunk.forEach((x, i) => {
        const y = 1.3 + i * (ch + 0.15);
        s.rect(0.45, y, W - 0.9, ch, { fill: 'FFFFFF', line: 'BFBFBF', lw: 0.75 });
        s.rect(0.45, y, 0.1, ch, { fill: color });
        s.text(0.7, y + 0.06, 3.1, ch - 0.12, [
          { t: `${ci * 3 + i + 1}. ${x.fn}`, b: true, sz: 13, color },
          { t: x.ind.name, b: true, sz: 15 },
          x.soft ? { t: '(판정 기준 미달 · 상대 비교)', sz: 10, color: GREY } : '',
        ].filter(Boolean), { anchor: 'ctr' });
        s.text(3.9, y + 0.06, W - 0.45 - 3.9 - 0.1, ch - 0.12, [
          { runs: [{ t: '관찰  ', b: true, color: NAVY }, x.obs] },
          { runs: [{ t: '논의  ', b: true, color: NAVY }, x.disc] },
          { runs: [{ t: '건의  ', b: true, color: NAVY }, x.rec] },
        ], { sz: 11.5, after: 3, anchor: 'ctr' });
      });
    });
  };
  odrSlides(R.sustain, '유지할 점 (Sustain)', GOOD);
  odrSlides(R.improve, '개선할 점 (Improve)', BAD);

  /* ---------- 14. 건의 및 발전방향 ---------- */
  {
    const s = slide('5. 건의 및 발전방향', sections[4]);
    const recRows = R.improve.slice(0, 6).map((x, i) => [{ t: String(i + 1), align: 'ctr' }, { t: x.fn, align: 'ctr' }, x.rec, x.target.join(', ') || R.sideName]);
    R.sustain.slice(0, Math.max(0, 7 - recRows.length)).slice(0, 2).forEach(x => recRows.push([{ t: String(recRows.length + 1), align: 'ctr' }, { t: x.fn, align: 'ctr' }, x.rec, '전 제대']));
    let y = 1.35;
    if (recRows.length) {
      const rows = [['번호', '기능', '건의 내용', '중점 대상']].concat(recRows);
      const rowH = Math.min(0.55, 3.9 / rows.length);
      s.table(0.45, y, W - 0.9, rows, { widths: [0.6, 1.3, 7.4, 2.2], rowH, sz: 11 });
      y += rows.length * rowH + 0.35;
    }
    const plan = [{ lvl: 0, t: '차기 훈련 중점', b: true }];
    if (R.weakF.length) plan.push({ lvl: 1, t: `${R.weakF.join('·')} 기능 보완을 차기 훈련의 중점 과제로 선정하고 동일 지표로 개선 여부 재측정` });
    const wu = R.units.filter(u => u.judge === '미흡');
    if (wu.length) plan.push({ lvl: 1, t: `${wu.map(u => u.name).join(', ')}은(는) 예비훈련 단계에서 취약 기능 중심 보충 교육` });
    plan.push({ lvl: 1, t: '본 분석 결과를 훈련 기록으로 축적하여 차기 훈련의 비교 기준으로 활용' });
    s.text(0.45, Math.min(y, 5.2), W - 0.9, 6.95 - Math.min(y, 5.2), plan, { sz: 14, after: 6 });
  }

  /* ---------- 15. 붙임 ---------- */
  {
    const s = slide('붙임 — 지표 산출 기준');
    const rows = [['기능', '지표', '산출 방법', '좋은 방향']].concat(Config.cur.indicators.map(i => [
      { t: i.fn, align: 'ctr' }, i.name + (i.judged === false ? ' (참고)' : ''), (IND_TEXT[i.id] || {}).how || '', { t: i.dir === 'low' ? '낮을수록' : '높을수록', align: 'ctr' },
    ]));
    s.table(0.45, 1.25, W - 0.9, rows, { widths: [1.3, 2.6, 6.2, 1.1], rowH: 0.322, sz: 9.5 });
    s.text(0.45, 6.75, W - 0.9, 0.3, { t: `교전 수: 최하위 제대 사격 간격 ${Config.cur.교전간격분}분 초과 시 별개 교전 · 보고 지연 기준 ${Config.cur.보고지연기준초}초 · 초기 피해 구간 ${Config.cur.초기피해구간분}분`, sz: 9, color: GREY });
  }

  return d.build();
};
