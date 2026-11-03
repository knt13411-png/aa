'use strict';
/* 분석 결과 엑셀(.xlsx) 내보내기: 요약 · 기능 진단 · 지표 상세 · 제대별 판정 · 전 제대 지표 · 유지/개선 · 시간대별 · 화기별 · 추이 · 지표 기준 */

const Results = {};

Results.fileName = R => `${U.safeName(R.info.훈련명)}_사후검토결과_${U.nowStr().slice(0, 10)}.xlsx`;

const ZF = { pct: '0.0%', num1: '0.0', num2: '0.00', num3: '0.000', int: '#,##0' };
const zOf = ind => ZF[ind.fmt] || '0.0';
/* 셀 값: 일반 값 또는 {v, z(표시 형식)} */
const cv = (v, z) => (v == null || (typeof v === 'number' && !isFinite(v)) ? '' : z ? { v, z } : v);

/* rows: [[셀]] → 시트. widths: 열 너비(글자 수) */
function sheetOf(rows, widths) {
  const plain = rows.map(r => r.map(c => (c && typeof c === 'object' ? c.v : c)));
  const ws = XLSX.utils.aoa_to_sheet(plain);
  rows.forEach((r, ri) => r.forEach((c, ci) => {
    if (c && typeof c === 'object' && c.z) {
      const cell = ws[XLSX.utils.encode_cell({ r: ri, c: ci })];
      if (cell) cell.z = c.z;
    }
  }));
  const n = Math.max(...rows.map(r => r.length));
  ws['!cols'] = Array.from({ length: n }, (_, i) => ({ wch: (widths && widths[i]) || 12 }));
  return ws;
}

Results.build = function (R) {
  const { info, root, opfor, diag, model } = R;
  const inds = Config.cur.indicators;
  const wb = XLSX.utils.book_new();
  const add = (name, rows, widths) => XLSX.utils.book_append_sheet(wb, sheetOf(rows, widths), name);

  /* 1. 요약 */
  const sum = [
    ['사후검토(AAR) 분석 결과'], [],
    ['훈련명', info.훈련명], ['훈련기간', `${info.시작일} ~ ${info.종료일}`], ['훈련부대', info.훈련부대 || R.sideName],
    ['비교 기준', R.baseText], ['판정 기준', `개선율 +${Config.cur.판정기준N}% 이상 강점, −${Config.cur.판정기준N}% 이하 취약`],
    ['원천 파일', info.원천파일명 || ''], ['분석 일시', model.분석일시], ['제외 행(형식 오류)', info.제외행수 || 0],
    ['누락 자료', (info.누락 || []).join(', ') || '없음'], [],
    ['종합 판단', R.overall], [],
    ['핵심 지표', R.sideName].concat(opfor ? ['대항군'] : []),
  ];
  const kv = (name, f, z) => sum.push([name, cv(f(root), z)].concat(opfor ? [cv(f(opfor), z)] : []));
  kv('교전 수', n => n.engagements, ZF.int);
  kv('사격 수', n => n.shots, ZF.int);
  kv('명중 수', n => n.hits, ZF.int);
  kv('명중률', n => n.metrics.fire_hit, ZF.pct);
  kv('피해 유발(건)', n => n.dmgHits, ZF.int);
  kv('인원(명)', n => n.personnel, ZF.int);
  kv('사망', n => n.cas.사망, ZF.int);
  kv('중상', n => n.cas.중상, ZF.int);
  kv('경상', n => n.cas.경상, ZF.int);
  kv('생존율', n => n.metrics.prot_survival, ZF.pct);
  if (R.exch != null) sum.push(['손실교환비(대항군 사망 ÷ 아군 사망)', cv(R.exch, ZF.num2)]);
  if (R.combat) sum.push(['피아 교전 판정', R.combat]);
  sum.push([], ['기능별 판정', '강점 ' + R.strongF.length + '개', '취약 ' + R.weakF.length + '개']);
  add('요약', sum, [30, 60, 16]);

  /* 2. 6대 기능 진단 */
  add('6대기능 진단', [['기능', `종합지수(${R.baseName}=100)`, '판정', '판단 근거']].concat(diag.functions.map(f => [
    f.fn, cv(f.index, ZF.num1), f.judge, f.summary.replace(/^\[[^\]]+\]\s*/, ''),
  ])), [14, 20, 10, 90]);

  /* 3. 지표 상세 */
  const det = [['기능', '지표', '단위', '좋은 방향', '이번 훈련', R.baseName, '변화율', '개선율', '판정', '판정 포함', '설명']];
  for (const f of diag.functions) for (const r of f.indicators) {
    const it = Config.ind(r.id);
    det.push([f.fn, r.name, it.fmt === 'pct' ? '%' : it.unit || '', it.dir === 'low' ? '낮을수록' : '높을수록', cv(r.value, zOf(it)), cv(r.base, zOf(it)),
      cv(r.change == null ? null : r.change / 100, ZF.pct), cv(r.improve == null ? null : r.improve / 100, ZF.pct), r.judge, r.judged ? '예' : '참고',
      r.sentence.replace(/^\[[^\]]+\]\s*/, '')]);
  }
  add('지표 상세', det, [12, 24, 7, 10, 11, 11, 10, 10, 9, 9, 80]);

  /* 4. 제대별 판정 */
  if (R.units.length) {
    add(`${R.lvName}별 판정`, [['제대', '인원', '종합지수', '판정'].concat(FUNCTIONS.map(f => f + ' 지수'), ['강점 기능', '취약 기능'])].concat(R.units.map(u => [
      u.name, u.node.personnel, cv(u.score, ZF.num1), u.judge,
    ].concat(u.diag.functions.map(f => cv(f.index, ZF.num1)), [u.strongFns.join(', '), u.weakFns.join(', ')]))), [18, 8, 10, 8, 11, 11, 11, 11, 11, 13, 20, 20]);
  }

  /* 5. 전 제대 지표 (훈련 전체·진영·여단~소대) */
  const nodes = Object.values(model.nodes).sort((a, b) =>
    (a.level === 0 ? -1 : 0) - (b.level === 0 ? -1 : 0) ||
    (a.side === b.side ? 0 : a.side === '훈련부대' ? -1 : 1) ||
    a.path.join(' ').localeCompare(b.path.join(' '), 'ko', { numeric: true }));
  const all = [['진영', '수준', '제대', '사격', '명중', '교전', '피해 유발', '인원', '사망', '중상', '경상'].concat(inds.map(i => i.name + (i.fmt === 'pct' ? '' : i.unit ? `(${i.unit})` : '')))];
  for (const n of nodes) {
    all.push([n.level === 0 ? (n.side === '전체' ? '전체' : n.side) : n.side, n.level === 0 ? '훈련 전체' : LEVELS[n.level], n.level === 0 ? (n.side === '전체' ? '훈련 전체' : n.side + ' 전체') : n.path.join(' '),
      n.shots, n.hits, n.engagements, n.dmgHits, n.personnel, n.cas.사망, n.cas.중상, n.cas.경상].concat(inds.map(i => cv(n.metrics[i.id], zOf(i)))));
  }
  add('전 제대 지표', all, [9, 9, 22, 8, 8, 7, 9, 7, 6, 6, 6].concat(inds.map(() => 13)));

  /* 6. 유지할 점 · 개선할 점 */
  const odr = [['구분', '기능', '지표', '관찰', '논의', '건의', '중점 대상', '비고']];
  R.sustain.forEach(x => odr.push(['유지할 점', x.fn, x.ind.name, x.obs, x.disc, x.rec, '전 제대', x.soft ? '판정 기준 미달 · 상대 비교' : '']));
  R.improve.forEach(x => odr.push(['개선할 점', x.fn, x.ind.name, x.obs, x.disc, x.rec, x.target.join(', ') || R.sideName, x.soft ? '판정 기준 미달 · 상대 비교' : '']));
  add('유지·개선', odr, [10, 12, 22, 55, 65, 60, 24, 22]);

  /* 7. 시간대별 */
  const ts = [['시각', `${R.sideName} 교전 개시`, `${R.sideName} 피해`].concat(opfor ? ['대항군 교전 개시', '대항군 피해'] : [])];
  root.tsDmg.forEach((_, i) => ts.push([U.timeStr(model.t0 + i * model.binMin * 60000), root.tsEng[i], root.tsDmg[i]].concat(opfor ? [opfor.tsEng[i], opfor.tsDmg[i]] : [])));
  add('시간대별', ts, [8, 18, 16, 16, 14]);

  /* 8. 화기별 */
  const wpn = [['진영', '화기', '사격', '명중', '명중률', '피해 기여']];
  for (const n of [root, opfor].filter(Boolean)) {
    Object.entries(n.weapons).sort((a, b) => b[1].shots - a[1].shots)
      .forEach(([w, v]) => wpn.push([n.side === '전체' ? '전체' : n.side, w, v.shots, v.hits, cv(v.shots ? v.hits / v.shots : null, ZF.pct), v.dmg]));
  }
  add('화기별', wpn, [10, 18, 9, 9, 9, 10]);

  /* 9. 훈련 간 추이 */
  if (R.trend.length > 1) {
    add('훈련 간 추이', [['지표'].concat(R.trend.map(t => `${t.name}(${t.date})`))].concat(inds.map(i => [i.name].concat(R.trend.map(t => cv(t.m[i.id], zOf(i)))))),
      [24].concat(R.trend.map(() => 20)));
  }

  /* 10. 지표 기준 */
  add('지표 기준', [['기능', '지표', '산출 방법', '좋은 방향', '자료 시트', '판정 포함']].concat(inds.map(i => [
    i.fn, i.name, (IND_TEXT[i.id] || {}).how || '', i.dir === 'low' ? '낮을수록' : '높을수록', i.src, i.judged === false ? '참고' : '예',
  ])).concat([[], ['설정값', `교전 간격 ${Config.cur.교전간격분}분 · 보고 지연 기준 ${Config.cur.보고지연기준초}초 · 초기 피해 구간 ${Config.cur.초기피해구간분}분 · 시간대 ${Config.cur.시간대단위분}분`]]),
  [12, 24, 50, 10, 10, 9]);

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};
