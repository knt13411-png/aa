'use strict';
/* 사후검토(AAR) 결과 보고서 (.docx)
   구성: 보고 요지 → 1.훈련 개요 → 2.훈련 결과 → 3.전투수행기능별 분석 → 4.제대별 분석
        → 5.주요 관찰 사항(관찰·논의·건의) → 6.건의 및 발전방향 → 붙임(지표 산출 기준)
   한국군 결과보고서의 □ ○ - 위계와, 미 육군 AAR(TC 25-20)의 유지할 점/개선할 점·관찰-논의-건의 형식을 따른다. */

const Report = {};

Report.fileName = R => `${U.safeName(R.info.훈련명)}_사후검토결과보고_${U.nowStr().slice(0, 10)}.docx`;

Report.build = async function (R) {
  const { info, root, opfor, diag } = R;
  const title = `${info.훈련명} 사후검토(AAR) 결과 보고`;
  const d = new Docx.Builder(title);
  const judgeColor = j => (j === '강점' || j === '우수' ? '0A7A0A' : j === '취약' || j === '미흡' ? 'C03030' : null);
  const jc = j => ({ t: j, b: j === '강점' || j === '취약' || j === '우수' || j === '미흡', color: judgeColor(j), align: 'center' });
  const ind = id => Config.ind(id);
  const IMG_W = 1500, IMG_H = 700, S = 2.1;

  // 차트 이미지
  const img = {};
  img.versus = await Charts.png(Charts.defs.versus(R, S), IMG_W, IMG_H);
  if (root.tsDmg.length > 1) img.timeline = await Charts.png(Charts.defs.timeline(R, S), IMG_W, IMG_H);
  if (diag.hasBase) img.radar = await Charts.png(Charts.defs.radar(R, S), 1200, 900);
  if (R.units.length > 1) {
    img.units = await Charts.png(Charts.defs.unitScore(R, S), IMG_W, Math.max(420, 110 + R.units.length * 70));
    img.cas = await Charts.png(Charts.defs.casualty(R, S), IMG_W, IMG_H);
  }
  if (R.trend.length > 1) img.trend = await Charts.png(Charts.defs.trend(R, S), IMG_W, IMG_H);

  /* ---------- 표지 머리 ---------- */
  d.title1(title);
  d.table([
    [{ t: '보고일자', b: true }, U.nowStr().slice(0, 10), { t: '훈련기간', b: true }, `${info.시작일} ~ ${info.종료일}`],
    [{ t: '훈련부대', b: true }, info.훈련부대 || '–', { t: '비교기준', b: true }, R.baseKind === 'prev' ? `이전 훈련 ${R.base._n}건 평균` : R.baseKind === 'opfor' ? '대항군(피아 비교)' : '없음'],
    [{ t: '분석자료', b: true }, { t: info.원천파일명 || '–', span: 3 }],
  ], { header: false, widths: [1.1, 2.4, 1.1, 2.4], firstColFill: 'EDEDED' });

  // 보고 요지
  const top = (arr, n) => arr.slice(0, n).map(x => x.title).join(', ') || '해당 없음';
  const boxP = (label, text) => Docx.para([{ t: `${label}  `, b: true, color: '1F3864' }, text], { after: 80, align: 'left' });
  d.box([
    Docx.para({ t: '보 고 요 지', b: true, sz: 24, color: '1F3864' }, { align: 'center', after: 120 }),
    boxP('□ 종합', R.overall),
    boxP('□ 유지할 점', top(R.sustain, 3)),
    boxP('□ 개선할 점', top(R.improve, 3)),
    boxP('□ 핵심 건의', R.improve.slice(0, 2).map(x => x.rec).join(' / ') || '현 수준 유지'),
  ]);

  /* ---------- 1. 훈련 개요 ---------- */
  d.h1('1. 훈련 개요');
  d.item(0, [{ t: '목 적 : ', b: true }, '과학화 훈련 데이터를 근거로 전투수행기능(6대 기능)별 성과를 진단하고, 유지할 점·개선할 점과 발전방향을 도출']);
  d.item(0, [{ t: '기 간 : ', b: true }, `${info.시작일} ~ ${info.종료일} (교전 기록 ${U.timeStr(R.model.tStart)} ~ ${U.timeStr(R.model.tEnd)})`]);
  d.item(0, [{ t: '참가부대', b: true }]);
  const orgTxt = Object.entries(R.org).filter(([, v]) => v).map(([k, v]) => `${k} ${v}개`).join(' · ');
  d.item(1, `${R.sideName}: ${orgTxt}, 인원 ${root.personnel.toLocaleString()}명`);
  if (opfor) d.item(1, `대항군: 인원 ${opfor.personnel.toLocaleString()}명`);
  d.item(0, [{ t: '분석 방법', b: true }]);
  const rows = Object.keys(SHEET_DEFS).map(k => `${k} ${(info.행수 && info.행수[k] || 0).toLocaleString()}건`).join(', ');
  d.item(1, `원천 기록: ${rows}`);
  if (info.제외행수) d.item(1, `형식 오류 ${info.제외행수.toLocaleString()}행은 분석에서 제외`);
  if (info.누락 && info.누락.length) d.item(1, `누락 자료: ${info.누락.join(', ')} → 관련 지표 "자료 없음" 처리`);
  d.item(1, `비교 기준: ${R.baseText}`);
  d.item(1, `판정 기준: 비교 기준 대비 개선율 +${Config.cur.판정기준N}% 이상 "강점", −${Config.cur.판정기준N}% 이하 "취약", 그 사이 "보통"`);

  /* ---------- 2. 훈련 결과 ---------- */
  d.h1('2. 훈련 결과');
  d.item(0, [{ t: '종합 판단', b: true }]);
  d.item(1, R.overall);
  d.item(0, [{ t: '피아 교전 결과', b: true }]);
  d.tcap('피아 교전 결과');
  const vsRow = (name, f, fo) => [{ t: name, b: true }, { t: f(root), align: 'center' }].concat(opfor ? [{ t: fo ? fo(opfor) : f(opfor), align: 'center' }] : []);
  d.table([
    ['구분', R.sideName].concat(opfor ? ['대항군'] : []),
    vsRow('교전 수', n => n.engagements.toLocaleString()),
    vsRow('사격 수 / 명중 수', n => `${n.shots.toLocaleString()} / ${n.hits.toLocaleString()}`),
    vsRow('명중률', n => U.fmt(n.metrics.fire_hit, 'pct')),
    vsRow('피해 유발(건)', n => n.dmgHits.toLocaleString()),
    vsRow('인원(명)', n => n.personnel.toLocaleString()),
    vsRow('사망 / 중상 / 경상', n => `${n.cas.사망} / ${n.cas.중상} / ${n.cas.경상}`),
    vsRow('생존율', n => U.fmt(n.metrics.prot_survival, 'pct')),
  ].concat(R.exch != null ? [[{ t: '손실교환비', b: true }, { t: `${R.exch.toFixed(2)} (대항군 사망 ÷ ${R.sideName} 사망)`, span: opfor ? 2 : 1, align: 'center' }]] : []),
  { widths: opfor ? [2, 2, 2] : [2, 3], firstColFill: 'F2F2F2' });
  d.figure(img.versus, '피아 핵심 지표 비교', 15.5);
  if (img.timeline) {
    d.item(0, [{ t: '시간대별 추이', b: true }]);
    if (R.peak) d.item(1, `${R.sideName} 피해는 ${R.peak.from}~${R.peak.to} 시간대에 가장 많이(${R.peak.n}건) 발생함.`);
    d.figure(img.timeline, `시간대별 피해·교전 추이 (${R.model.binMin}분 단위)`, 15.5);
  }
  if (img.trend) {
    d.item(0, [{ t: '이전 훈련 대비 추이', b: true }]);
    const first = R.trend[0], last = R.trend[R.trend.length - 1];
    ['fire_hit', 'prot_survival', 'int_predetect'].forEach(id => {
      const a = first.m[id], b = last.m[id];
      if (a != null && b != null) d.item(1, `${ind(id).name}: ${first.name} ${U.fmt(a, 'pct')} → ${last.name} ${U.fmt(b, 'pct')} (${b - a >= 0 ? '+' : ''}${((b - a) * 100).toFixed(1)}%p)`);
    });
    d.figure(img.trend, '훈련 간 주요 지표 추이', 15.5);
  }

  /* ---------- 3. 전투수행기능별 분석 ---------- */
  d.h1('3. 전투수행기능별 분석');
  d.item(0, [{ t: '기능별 종합 판정', b: true }, ` (비교 기준: ${R.baseText})`]);
  d.tcap('6대 전투수행기능 종합 판정');
  d.table([['기능', '종합지수', '판정', '판단 근거']].concat(diag.functions.map(f => [
    { t: f.fn, b: true, align: 'center' }, { t: f.index == null ? '–' : f.index.toFixed(0), align: 'center' }, jc(f.judge),
    f.summary.replace(/^\[[^\]]+\]\s*/, '').replace(/^종합지수 \d+ \([^)]*\) → [^.]+\.\s*/, '') || '–',
  ])), { widths: [1.3, 1, 0.9, 5.8] });
  if (img.radar) d.figure(img.radar, `6대 기능 종합지수 (${R.baseName}=100)`, 11);

  const fnLetters = ['가', '나', '다', '라', '마', '바'];
  diag.functions.forEach((f, i) => {
    d.h2(`${fnLetters[i]}. ${f.fn}`);
    d.table([['지표', '이번 훈련', R.baseName, '변화', '판정']].concat(f.indicators.map(r => {
      const it = ind(r.id);
      return [
        `${r.name}${it.dir === 'low' ? ' (↓)' : ''}`, { t: U.fmtInd(it, r.value), align: 'center' }, { t: r.base == null ? '–' : U.fmtInd(it, r.base), align: 'center' },
        { t: r.change == null ? '–' : Insight.diffText(it, r), align: 'center', color: r.improve == null ? null : r.improve >= 0 ? '0A7A0A' : 'C03030' }, jc(r.judge),
      ];
    })), { widths: [3.2, 1.5, 1.5, 1.3, 1] });
    d.item(1, f.summary.replace(/^\[[^\]]+\]\s*/, ''));
  });
  d.note('※ (↓) 표시 지표는 값이 낮을수록 좋음. 종합지수는 판정 지표별 지수(비교 기준=100)의 평균.');

  /* ---------- 4. 제대별 분석 ---------- */
  if (R.units.length) {
    d.h1(`4. 제대별 분석 (${R.lvName})`);
    d.item(0, `동일 수준 ${R.lvName} 평균(=100) 대비 각 제대의 기능별 지수를 평균하여 종합지수를 산출함.`);
    d.tcap(`${R.lvName}별 주요 지표 및 종합 판정`);
    d.table([['제대', '명중률', '생존율', '선탐지', '보고지연', '종합지수', '판정', '취약 기능']].concat(R.units.map(u => {
      const m = u.node.metrics;
      return [{ t: u.name, b: true }, { t: U.fmt(m.fire_hit, 'pct'), align: 'center' }, { t: U.fmt(m.prot_survival, 'pct'), align: 'center' },
        { t: U.fmt(m.int_predetect, 'pct'), align: 'center' }, { t: U.fmt(m.c2_report_delay, 'pct'), align: 'center' },
        { t: u.score == null ? '–' : u.score.toFixed(0), align: 'center' }, jc(u.judge), u.weakFns.join(', ') || '–'];
    })), { widths: [1.8, 1.1, 1.1, 1.1, 1.1, 1.1, 0.9, 2.2], fontSz: 18 });
    if (img.units) d.figure(img.units, `${R.lvName}별 종합지수 (동일 제대 평균=100)`, 15);
    const best = R.units.filter(u => u.judge === '우수'), worst = R.units.filter(u => u.judge === '미흡');
    d.item(0, [{ t: '우수 제대 : ', b: true }, best.length ? best.map(u => `${u.name}(${u.score.toFixed(0)}${u.strongFns.length ? ', 강점 ' + u.strongFns.join('·') : ''})`).join(', ') : '해당 없음']);
    d.item(0, [{ t: '미흡 제대 : ', b: true }, worst.length ? worst.map(u => `${u.name}(${u.score.toFixed(0)}${u.weakFns.length ? ', 취약 ' + u.weakFns.join('·') : ''})`).join(', ') : '해당 없음']);
    if (img.cas) d.figure(img.cas, `${R.lvName}별 피해 현황 (명)`, 15.5);
  }

  /* ---------- 5. 주요 관찰 사항 ---------- */
  d.h1(`${R.units.length ? 5 : 4}. 주요 관찰 사항 (관찰 · 논의 · 건의)`);
  const odr = (list, label) => {
    d.h2(label);
    if (!list.length) { d.item(1, '해당 사항 없음'); return; }
    list.forEach((x, i) => {
      d.p([{ t: `${i + 1}) ${x.title}`, b: true }, x.soft ? { t: '  (판정 기준 미달 · 상대 비교)', sz: 18, color: '7F7F7F' } : ''], { keepNext: true, left: 200, before: 120, after: 40 });
      d.item(1, [{ t: '관찰 : ', b: true }, x.obs]);
      d.item(1, [{ t: '논의 : ', b: true }, x.disc]);
      d.item(1, [{ t: '건의 : ', b: true }, x.rec]);
    });
  };
  odr(R.sustain, '가. 유지할 점 (Sustain)');
  odr(R.improve, '나. 개선할 점 (Improve)');

  /* ---------- 6. 건의 및 발전방향 ---------- */
  d.h1(`${R.units.length ? 6 : 5}. 건의 및 발전방향`);
  const recRows = R.improve.map((x, i) => [{ t: String(i + 1), align: 'center' }, { t: x.fn, align: 'center' }, x.rec, x.target.join(', ') || R.sideName]);
  R.sustain.slice(0, 2).forEach(x => recRows.push([{ t: String(recRows.length + 1), align: 'center' }, { t: x.fn, align: 'center' }, x.rec, '전 제대']));
  if (recRows.length) {
    d.tcap('건의 사항');
    d.table([['번호', '기능', '건의 내용', '중점 대상']].concat(recRows), { widths: [0.7, 1.3, 5.5, 1.8] });
  }
  d.item(0, [{ t: '차기 훈련 중점', b: true }]);
  if (R.weakF.length) d.item(1, `${R.weakF.join('·')} 기능 보완을 차기 훈련의 중점 과제로 선정하고, 동일 지표로 개선 여부를 재측정`);
  const wu = R.units.filter(u => u.judge === '미흡');
  if (wu.length) d.item(1, `${wu.map(u => u.name).join(', ')}은(는) 예비훈련 단계에서 취약 기능 중심의 보충 교육 실시`);
  d.item(1, '본 분석 결과를 훈련 기록으로 축적하여 차기 훈련의 비교 기준으로 활용');

  d.p('', { after: 200 });
  d.note('※ 본 보고서는 사후검토 분석기가 훈련 데이터를 자동 분석하여 작성한 초안이며, 판정은 규칙 기반 1차 판단임. 담당자 검토 후 최종 판단을 기재할 것.');
  d.p({ t: '끝.', b: true }, { align: 'right' });

  /* ---------- 붙임 ---------- */
  d.pageBreak();
  d.p({ t: '붙임  지표 산출 기준', b: true, sz: 24 }, { after: 160 });
  d.table([['기능', '지표', '산출 방법', '좋은 방향', '자료']].concat(Config.cur.indicators.map(i => [
    { t: i.fn, align: 'center' }, i.name + (i.judged === false ? ' (참고)' : ''), (IND_TEXT[i.id] || {}).how || '', { t: i.dir === 'low' ? '낮을수록' : '높을수록', align: 'center' }, { t: i.src, align: 'center' },
  ])), { widths: [1.3, 2.4, 3.8, 1.1, 1.1], fontSz: 18 });
  d.note(`※ 교전 수는 최하위 제대별 사격 간격이 ${Config.cur.교전간격분}분을 넘으면 별개 교전으로 셈. 보고 지연 기준 ${Config.cur.보고지연기준초}초, 초기 피해 구간 ${Config.cur.초기피해구간분}분.`);

  return d.build();
};
