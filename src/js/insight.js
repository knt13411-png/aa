'use strict';
/* 분석 결과 → 보고용 정리: 비교 기준 선택, 6대 기능 판정, 제대별 우열, 유지할 점·개선할 점(관찰-논의-건의) */

/* 지표별 산출 방법(붙임용)과 개선 건의 문안 */
const IND_TEXT = {
  c2_order_sec: { how: '명령 응답시각 − 하달시각의 평균', rec: '단편명령 양식·전파 순서를 표준화하고, 지휘소연습(CPX)으로 명령 하달 절차를 반복 숙달' },
  c2_report_delay: { how: `보고 소요시간이 기준(설정값)을 넘은 보고 ÷ 전체 보고`, rec: '상황·접적·피해보고의 시기와 양식을 부대 SOP에 명시하고, 지연 제대 대상 통신 운용 절차를 재교육' },
  c2_max_gap: { how: '제대 첫 사격~마지막 사격 사이 가장 긴 무사격 구간', rec: '교전 공백 구간의 상황 판단·후속 조치 지연 원인을 분석하고, 예비대 투입·전투력 재편성 결심 시점을 사전 계획' },
  int_id_sec: { how: '식별시각 − 탐지시각의 평균', rec: '감시장비 운용과 피아식별 교육을 강화하고, 탐지-식별-보고 연계 훈련을 반복' },
  int_predetect: { how: '상대보다 먼저 탐지한 건 ÷ 전체 탐지', rec: '감시·정찰 자산 배치 계획을 보완하고, 관측소(OP)·정찰조 운용과 기도비닉을 강화' },
  man_speed: { how: '이동 중인 위치 기록의 평균 속도', rec: '기동로를 사전 정찰하고, 기동 대형·통제 수단과 기동 간 지휘통제 절차를 숙달' },
  man_move_dmg: { how: '이동 중 피해 ÷ 이동 여부가 기록된 피해', rec: '사격과 기동(엄호) 절차를 숙달하고, 노출 구간에 연막·제압사격 계획을 반영' },
  fire_hit: { how: '명중 ÷ 사격', rec: '화기별 사격술 예비훈련과 전투사격을 숙달하고, 사거리·표적 특성에 맞는 화기 선택을 교육' },
  fire_dmg_per_shot: { how: '피해 판정이 난 사격 ÷ 전체 사격', rec: '표적 우선순위와 화력 집중 절차를 정립하고, 사격 지정·분배 통제를 숙달' },
  fire_indirect: { how: '간접화기 피해 ÷ 전체 피해 (참고 지표)', rec: '관측자-사격지휘소 연계 등 간접화력 요청·유도 절차를 숙달' },
  prot_survival: { how: '(인원 − 사망) ÷ 인원, 개체별 가장 심한 피해 기준', rec: '진지 강화·분산·엄폐 등 부대 방호 대책을 보완하고 전투손실 최소화 절차를 정립' },
  prot_uncovered: { how: '비엄폐 상태 피해 ÷ 엄폐 여부가 기록된 피해', rec: '정지 시 즉각 엄폐를 습관화하고, 진지 편성과 엄폐물 활용 교육을 강화' },
  prot_early: { how: '교전 개시 후 초기 구간(설정값) 내 피해 ÷ 전체 피해', rec: '접적 전 경계와 전투준비태세를 강화하고, 초기 교전 즉각조치 훈련을 반복' },
  sus_ammo: { how: '탄약 보급 수량 ÷ 탄약 소모량(없으면 사격 수)', rec: '탄약 소모를 예측해 사전 추진보급 계획을 세우고, 보급 요청 절차를 간소화' },
  sus_evac_min: { how: '후송 완료시각 − 요청시각의 평균', rec: '구호소 위치·후송로·후송수단을 사전 계획하고 응급처치를 숙달' },
  sus_resup_min: { how: '재보급 완료시각 − 요청시각의 평균', rec: '재보급 지점·시기를 사전 판단하고, 전투근무지원 제대와의 협조 체계를 강화' },
};

const Insight = {};

Insight.unitName = n => n.path.join(' ');

/* 지표 값이 좋은 순으로 정렬 (dir 반영) */
Insight.rankBy = function (rows, id) {
  const ind = Config.ind(id);
  return rows.filter(u => u.node.metrics[id] != null)
    .sort((a, b) => (ind.dir === 'low' ? a.node.metrics[id] - b.node.metrics[id] : b.node.metrics[id] - a.node.metrics[id]));
};

Insight.diffText = function (ind, r) {
  if (r.change == null) return '';
  return ind.fmt === 'pct'
    ? `${r.value - r.base >= 0 ? '+' : ''}${((r.value - r.base) * 100).toFixed(1)}%p`
    : `${r.change >= 0 ? '+' : ''}${r.change.toFixed(1)}%`;
};

/* cur: {info, model}, history: [{id, 훈련명, 시작일, levelStats}] */
Insight.build = function (cur, history) {
  const { info, model } = cur;
  const nodes = model.nodes;
  const cfg = Config.cur;
  const side = nodes['ROOT|훈련부대'] ? '훈련부대' : '전체';
  const root = nodes['ROOT|' + side];
  const opfor = side === '훈련부대' ? nodes['ROOT|대항군'] || null : null;

  // 비교 기준: 이전 훈련 평균 → 없으면 대항군(피아 비교)
  const prev = history.filter(h => h.id !== info.id && String(h.시작일 || '') < String(info.시작일 || ''))
    .sort((a, b) => String(a.시작일).localeCompare(String(b.시작일)));
  let base = Engine.baseline('ROOT|' + side, prev.map(h => h.levelStats));
  let baseKind = null, baseName = '이전 평균', baseText = '비교 기준 없음';
  if (base) { baseKind = 'prev'; baseText = `이전 훈련 ${base._n}건(${prev.map(h => h.훈련명).join(', ')}) 평균`; }
  else if (opfor) { base = Object.assign({ _n: 1 }, opfor.metrics); baseKind = 'opfor'; baseName = '대항군'; baseText = '대항군 지표 (피아 비교 · 이전 훈련 자료 없음)'; }
  const diag = Engine.diagnose(root.metrics, base, cfg, baseName);

  // 제대별: 훈련부대 대대(없으면 여단/중대) 수준, 같은 수준 평균 대비
  const unitSide = side === '전체' ? null : side;
  let lv = 0;
  for (const l of [2, 1, 3, 4]) if (Object.values(nodes).some(n => n.level === l && (!unitSide || n.side === unitSide))) { lv = l; break; }
  const unitNodes = Object.values(nodes).filter(n => n.level === lv && (!unitSide || n.side === unitSide))
    .sort((a, b) => Insight.unitName(a).localeCompare(Insight.unitName(b), 'ko', { numeric: true }));
  const lvAvg = unitNodes.length ? model.levelStats[unitNodes[0].side + '|' + lv] : null;
  const units = unitNodes.map(node => {
    const d = Engine.diagnose(node.metrics, lvAvg ? Object.assign({ _n: 1 }, lvAvg) : null, cfg, '동일 제대 평균');
    const idx = d.functions.map(f => f.index).filter(x => x != null);
    const score = idx.length ? U.mean(idx) : null;
    const N = cfg.판정기준N;
    const judge = score == null ? '자료 없음' : score >= 100 + N / 2 ? '우수' : score <= 100 - N / 2 ? '미흡' : '보통';
    const weakFns = d.functions.filter(f => f.judge === '취약').map(f => f.fn);
    const strongFns = d.functions.filter(f => f.judge === '강점').map(f => f.fn);
    return { node, name: Insight.unitName(node), diag: d, score, judge, weakFns, strongFns };
  });

  // 유지할 점 / 개선할 점 (관찰-논의-건의)
  const items = [];
  for (const f of diag.functions) for (const r of f.indicators) if (r.judged && r.improve != null) items.push({ fn: f.fn, r });
  let sustain = items.filter(x => x.r.judge === '강점').sort((a, b) => b.r.improve - a.r.improve);
  let improve = items.filter(x => x.r.judge === '취약').sort((a, b) => a.r.improve - b.r.improve);
  // 강점/취약이 하나도 없으면 상대적으로 가장 나은/못한 지표를 2개씩 제시
  if (!sustain.length) sustain = items.filter(x => x.r.improve > 0).sort((a, b) => b.r.improve - a.r.improve).slice(0, 2).map(x => Object.assign({ soft: true }, x));
  if (!improve.length) improve = items.filter(x => x.r.improve < 0).sort((a, b) => a.r.improve - b.r.improve).slice(0, 2).map(x => Object.assign({ soft: true }, x));

  const odr = (x, kind) => {
    const ind = Config.ind(x.r.id);
    const ranked = Insight.rankBy(units, x.r.id);
    const fmtU = u => `${u.name}(${U.fmtInd(ind, u.node.metrics[x.r.id])})`;
    const isTime = ind.unit === '초' || ind.unit === '분';
    const verb = kind === 'sustain' ? (ind.dir === 'low' ? (isTime ? '단축' : '감소') : '향상') : (ind.dir === 'low' ? (isTime ? '지연' : '증가') : '저하');
    const obs = `${ind.name}이(가) ${U.fmtInd(ind, x.r.value)}로 ${baseName} ${U.fmtInd(ind, x.r.base)} 대비 ${Insight.diffText(ind, x.r)} ${verb}되었음.`;
    let disc = '';
    if (ranked.length >= 2) {
      const top = ranked.slice(0, 2), bot = ranked.slice(-2).reverse();
      disc = kind === 'sustain'
        ? `제대별로는 ${top.map(fmtU).join(', ')}이(가) 가장 우수하여 성과를 견인한 것으로 판단됨.`
        : `제대별로는 ${bot.map(fmtU).join(', ')}이(가) 가장 저조하여 전체 지표를 낮춘 주요 요인으로 판단됨.`;
      const lo = ranked[ranked.length - 1].node.metrics[x.r.id], hi = ranked[0].node.metrics[x.r.id];
      disc += ` (제대 간 편차 ${ind.fmt === 'pct' ? (Math.abs(hi - lo) * 100).toFixed(1) + '%p' : U.fmtInd(ind, Math.abs(hi - lo))})`;
    } else disc = `${x.fn} 기능 ${x.r.judge === '강점' || x.r.judge === '취약' ? x.r.judge : '변화'}의 직접 요인으로 판단됨.`;
    const rec = kind === 'sustain'
      ? `${ranked.length ? ranked[0].name + '의 ' : ''}${ind.name} 관련 수행 방법을 사례로 정리하여 부대 SOP에 반영하고 타 제대에 전파`
      : (IND_TEXT[x.r.id] || {}).rec || '원인 분석 후 보완 훈련 계획 수립';
    const target = kind === 'improve' && ranked.length >= 2 ? ranked.slice(-2).map(u => u.name) : [];
    return Object.assign({}, x, { kind, ind, title: `${x.fn} · ${ind.name}`, obs, disc, rec, target });
  };
  sustain = sustain.slice(0, 5).map(x => odr(x, 'sustain'));
  improve = improve.map(x => odr(x, 'improve'));

  // 피아 교전 결과
  const exch = opfor && root.cas.사망 > 0 ? opfor.cas.사망 / root.cas.사망 : null; // 손실교환비 = 적 사망 ÷ 아군 사망
  let combat = '';
  if (opfor) {
    const r = root.dmgHits, o = opfor.dmgHits;
    const ratio = o > 0 ? r / o : null;
    combat = ratio == null ? '' : ratio >= 1.2 ? '교전 우세' : ratio <= 0.83 ? '교전 열세' : '대등';
  }
  const strongF = diag.functions.filter(f => f.judge === '강점').map(f => f.fn);
  const weakF = diag.functions.filter(f => f.judge === '취약').map(f => f.fn);
  const sideName = side === '훈련부대' ? (info.훈련부대 || '훈련부대') : '참가부대';
  let overall = '';
  if (opfor) {
    overall += `피아 교전 결과 ${U.josa(sideName, '은', '는')} 피해 유발 ${root.dmgHits.toLocaleString()}건 대 ${opfor.dmgHits.toLocaleString()}건, ` +
      `명중률 ${U.fmt(root.metrics.fire_hit, 'pct')} 대 ${U.fmt(opfor.metrics.fire_hit, 'pct')}${exch != null ? `, 손실교환비 ${exch.toFixed(2)}` : ''}로 ${combat ? `${combat}였음` : '우열 판정 불가'}. `;
  }
  if (diag.hasBase) {
    overall += `${baseName} 대비 전투수행기능 진단 결과 ` +
      (strongF.length ? `${strongF.join('·')} 기능은 강점, ` : '') +
      (weakF.length ? `${weakF.join('·')} 기능은 취약으로 판단됨.` : strongF.length ? '그 외 기능은 보통 수준임.' : '6대 기능 모두 보통 수준임.');
  } else overall += '비교 기준이 없어 기능별 판정은 하지 않고 지표 수치만 제시함.';
  const weakUnits = units.filter(u => u.judge === '미흡');
  if (weakUnits.length) overall += ` 제대별로는 ${weakUnits.map(u => u.name).join(', ')}의 수행 수준이 동일 제대 평균보다 낮아 집중 보완이 필요함.`;

  // 피해 집중 시간대
  let peak = null;
  if (root.tsDmg && root.tsDmg.length) {
    let bi = 0;
    root.tsDmg.forEach((v, i) => { if (v > root.tsDmg[bi]) bi = i; });
    if (root.tsDmg[bi] > 0) {
      const t = model.t0 + bi * model.binMin * 60000;
      peak = { from: U.timeStr(t), to: U.timeStr(t + model.binMin * 60000), n: root.tsDmg[bi] };
    }
  }

  // 훈련 간 추이 (이전 최대 4건 + 이번)
  const trend = prev.slice(-4).map(h => ({ name: h.훈련명, date: h.시작일, m: h.levelStats && h.levelStats['ROOT|' + side] }))
    .concat([{ name: info.훈련명, date: info.시작일, m: root.metrics, cur: true }])
    .filter(t => t.m);

  const countLevel = l => Object.values(nodes).filter(n => n.level === l && (!unitSide || n.side === unitSide)).length;
  return {
    info, model, side, sideName, root, opfor, base, baseKind, baseName, baseText, diag,
    lv, lvName: LEVELS[lv] || '제대', units, sustain, improve, exch, combat, overall, strongF, weakF, peak, trend,
    org: { 여단: countLevel(1), 대대: countLevel(2), 중대: countLevel(3), 소대: countLevel(4) },
  };
};
