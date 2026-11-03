'use strict';
/* 분석 엔진: 엑셀 파싱/검증 → 제대 트리 집계 → 지표 계산 → 6대 기능 진단 */

const Engine = {};

/* ---------- 1. 파싱 및 검증 ---------- */
Engine.readWorkbook = function (buf) {
  return XLSX.read(buf, { type: 'array', cellDates: false, dense: true, cellFormula: false, cellHTML: false, cellText: false });
};

/* 결과: { tables: {시트: [{열: 값}]}, errors: [{sheet,row,col,value,reason}], errorRows: {시트: Set(행번호)},
           missingSheets: [], missingCols: [{sheet, cols}], counts } */
Engine.parse = function (wb, cfg, baseDayMs) {
  const res = { tables: {}, errors: [], errorRows: {}, missingSheets: [], missingCols: [], counts: {}, sheetNames: wb.SheetNames.slice() };
  const normSheets = {};
  wb.SheetNames.forEach(n => { normSheets[U.norm(n)] = n; });
  for (const [key, def] of Object.entries(SHEET_DEFS)) {
    const map = cfg.sheets[key];
    const actual = normSheets[U.norm(map.sheetName)] || normSheets[U.norm(key)];
    res.tables[key] = [];
    res.errorRows[key] = new Set();
    if (!actual) { res.missingSheets.push(key); continue; }
    const ws = wb.Sheets[actual];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    if (!aoa.length) { res.missingSheets.push(key); continue; }
    const header = aoa[0].map(h => U.norm(h));
    const idx = {};
    const miss = [];
    for (const [col, cd] of Object.entries(def.cols)) {
      const want = map.columns[col] || col;
      let i = header.indexOf(U.norm(want));
      if (i < 0) i = header.indexOf(U.norm(col));
      if (i < 0) i = header.indexOf(U.norm(col.replace(/\(.*\)/, ''))); // '거리(m)' ↔ '거리'
      if (i >= 0) idx[col] = i; else if (cd.req) miss.push(col);
    }
    if (miss.length) { res.missingCols.push({ sheet: key, actual, cols: miss }); continue; }
    const rows = res.tables[key];
    const cols = Object.entries(def.cols).filter(([c]) => idx[c] !== undefined);
    for (let r = 1; r < aoa.length; r++) {
      const a = aoa[r];
      const o = {};
      let bad = false;
      for (const [c, cd] of cols) {
        const raw = a[idx[c]];
        let v;
        if (cd.type === 'time') {
          v = U.parseTime(raw, baseDayMs);
          if (Number.isNaN(v)) { bad = true; res.errors.push({ sheet: key, row: r + 1, col: c, value: raw, reason: '날짜/시각 형식 오류' }); v = null; }
          else if (v === null && cd.req) { bad = true; res.errors.push({ sheet: key, row: r + 1, col: c, value: '', reason: '필수 값 없음' }); }
        } else if (cd.type === 'num') {
          v = U.parseNum(raw);
          if (Number.isNaN(v)) { bad = true; res.errors.push({ sheet: key, row: r + 1, col: c, value: raw, reason: '숫자 형식 오류' }); v = null; }
        } else {
          v = raw == null ? '' : String(raw).trim();
          if (!v && cd.req) { bad = true; res.errors.push({ sheet: key, row: r + 1, col: c, value: '', reason: '필수 값 없음' }); }
        }
        o[c] = v;
      }
      if (bad) res.errorRows[key].add(r + 1);
      else rows.push(o);
    }
    res.counts[key] = aoa.length - 1;
  }
  return res;
};

/* ---------- 2. 집계 ---------- */
const SEV = { 경상: 1, 중상: 2, 사망: 3 };
const SEV_NAME = ['없음', '경상', '중상', '사망'];
const isYes = v => /^(y|yes|예|o|1|true|이동|엄폐)$/i.test(String(v || '').trim());
const isNo = v => /^(n|no|아니오|x|0|false|정지|비엄폐)$/i.test(String(v || '').trim());

function newNode(id, side, level, name, path, parent) {
  return {
    id, side, level, name, path, parent, children: [],
    shots: 0, hits: 0, dmgHits: 0, indirectDmgHits: 0, engagements: 0,
    weapons: {}, personnel: 0, cas: { 사망: 0, 중상: 0, 경상: 0 },
    dmgRows: 0, dmgUncovered: 0, dmgCoverKnown: 0, dmgMoving: 0, dmgMoveKnown: 0, dmgEarly: 0,
    tsEng: null, tsDmg: null,
    det: 0, detIdSum: 0, detIdN: 0, detPre: 0, detPreKnown: 0,
    moveSpeedSum: 0, moveN: 0,
    orders: 0, orderSecSum: 0, orderN: 0, reports: 0, reportN: 0, reportDelayed: 0,
    maxGapMin: null, firstTime: null,
    ammoSupply: 0, ammoUse: 0, evacSum: 0, evacN: 0, resupSum: 0, resupN: 0,
    metrics: {},
  };
}

Engine.analyze = function (tables, cfg, info) {
  const nodes = {};
  const ROOT = s => 'ROOT|' + s;
  for (const s of SIDES) nodes[ROOT(s)] = newNode(ROOT(s), s, 0, info.훈련명 || '훈련', [], null);

  const chainCache = new Map();
  /* 기록 1건이 속하는 노드 id 목록 (전체 루트, 진영 루트, 여단, 대대, 중대, 소대) */
  function chain(side, names) {
    side = String(side || '').trim();
    if (side !== '훈련부대' && side !== '대항군') side = /대항|opfor|적/i.test(side) ? '대항군' : '훈련부대';
    const key = side + '|' + names.join('|');
    let c = chainCache.get(key);
    if (c) return c;
    c = [ROOT('전체'), ROOT(side)];
    const path = [];
    let parent = ROOT(side);
    for (let i = 0; i < names.length && i < 4; i++) {
      const nm = String(names[i] == null ? '' : names[i]).trim();
      if (!nm) break;
      path.push(nm);
      const id = side + '|' + path.join('|');
      if (!nodes[id]) {
        nodes[id] = newNode(id, side, i + 1, nm, path.slice(), parent);
        nodes[id].tsEng = new Array(nBins).fill(0);
        nodes[id].tsDmg = new Array(nBins).fill(0);
        nodes[parent].children.push(id);
        if (i === 0) nodes[ROOT('전체')].children.push(id);
      }
      c.push(id);
      parent = id;
    }
    chainCache.set(key, c);
    return c;
  }
  const unitChain = r => chain(r.진영, [r.여단, r.대대, r.중대, r.소대]);

  // 전체 시간 범위
  let tMin = Infinity, tMax = -Infinity;
  const touch = t => { if (t != null) { if (t < tMin) tMin = t; if (t > tMax) tMax = t; } };
  tables.교전기록.forEach(r => touch(r.시각));
  tables.피해기록.forEach(r => touch(r.시각));
  if (!isFinite(tMin)) { tables.위치기록.forEach(r => touch(r.시각)); }
  if (!isFinite(tMin)) { tMin = U.dayMs(info.시작일) || 0; tMax = tMin + 3600000; }
  const binMs = cfg.시간대단위분 * 60000;
  const t0 = Math.floor(tMin / binMs) * binMs;
  const nBins = Math.min(2000, Math.max(1, Math.floor((tMax - t0) / binMs) + 1));
  const bin = t => Math.max(0, Math.min(nBins - 1, Math.floor((t - t0) / binMs)));
  for (const s of SIDES) { nodes[ROOT(s)].tsEng = new Array(nBins).fill(0); nodes[ROOT(s)].tsDmg = new Array(nBins).fill(0); }

  const entities = new Map(); // 개체ID → {chain, sev}
  const addEnt = (id, ch) => { if (!id) return; if (!entities.has(id)) entities.set(id, { chain: ch, sev: 0 }); };
  const fireTimes = {}; // 노드 → 사격 시각
  const leafFire = new Map(); // 최하위 노드 → 사격 시각 (교전 구분)
  const dmgTimes = {};

  // 교전기록
  for (const r of tables.교전기록) {
    const ch = unitChain(r);
    const hit = /명중|hit/i.test(r.결과);
    const sev = SEV[r.피해판정] || 0;
    const indirect = /간접/.test(r.화기구분);
    const w = r.화기종류 || '기타';
    for (const id of ch) {
      const n = nodes[id];
      n.shots++;
      if (hit) n.hits++;
      if (sev) { n.dmgHits++; if (indirect) n.indirectDmgHits++; }
      const ws = n.weapons[w] || (n.weapons[w] = { shots: 0, hits: 0, dmg: 0 });
      ws.shots++; if (hit) ws.hits++; if (sev) ws.dmg++;
      (fireTimes[id] || (fireTimes[id] = [])).push(r.시각);
    }
    const leaf = ch[ch.length - 1];
    (leafFire.get(leaf) || leafFire.set(leaf, []).get(leaf)).push(r.시각);
    addEnt(r.사수ID, ch);
  }

  // 교전 구분: 최하위 제대별로 사격 간격이 기준을 넘으면 새 교전
  const gapMs = cfg.교전간격분 * 60000;
  const leafChain = {};
  for (const c of chainCache.values()) leafChain[c[c.length - 1]] = c;
  for (const [leaf, times] of leafFire) {
    times.sort((a, b) => a - b);
    let last = -Infinity;
    for (const t of times) {
      if (t - last > gapMs) {
        for (const id of leafChain[leaf]) { nodes[id].engagements++; nodes[id].tsEng[bin(t)]++; }
      }
      last = t;
    }
  }
  // 최장 교전 공백: 제대의 첫 사격~마지막 사격 사이 최대 무사격 구간
  for (const id in fireTimes) {
    const ts = fireTimes[id].sort((a, b) => a - b);
    let g = 0;
    for (let i = 1; i < ts.length; i++) if (ts[i] - ts[i - 1] > g) g = ts[i] - ts[i - 1];
    nodes[id].maxGapMin = ts.length > 1 ? g / 60000 : null;
    nodes[id].firstTime = ts[0];
  }

  // 피해기록
  for (const r of tables.피해기록) {
    const sev = SEV[r.피해판정] || 0;
    const ch = unitChain(r);
    addEnt(r.개체ID, ch);
    if (!sev) continue;
    const e = entities.get(r.개체ID);
    if (e && sev > e.sev) e.sev = sev;
    const uncovered = r.엄폐여부 ? isNo(r.엄폐여부) : null;
    const moving = r.이동여부 ? isYes(r.이동여부) : null;
    for (const id of ch) {
      const n = nodes[id];
      n.dmgRows++;
      if (uncovered !== null) { n.dmgCoverKnown++; if (uncovered) n.dmgUncovered++; }
      if (moving !== null) { n.dmgMoveKnown++; if (moving) n.dmgMoving++; }
      n.tsDmg[bin(r.시각)]++;
      (dmgTimes[id] || (dmgTimes[id] = [])).push(r.시각);
    }
  }
  // 초기 피해 집중도: 제대 교전 개시(첫 사격 또는 첫 피해) 후 N분 내 피해
  const earlyMs = cfg.초기피해구간분 * 60000;
  for (const id in dmgTimes) {
    const n = nodes[id];
    let first = n.firstTime != null ? n.firstTime : Infinity;
    for (const t of dmgTimes[id]) if (t < first) first = t;
    n.dmgEarly = dmgTimes[id].filter(t => t <= first + earlyMs).length;
  }

  // 위치기록: 인원 파악 + 이동속도
  for (const r of tables.위치기록) {
    const ch = unitChain(r);
    addEnt(r.개체ID, ch);
    if (isYes(r.이동여부) && r['속도(km/h)'] != null) {
      for (const id of ch) { nodes[id].moveSpeedSum += r['속도(km/h)']; nodes[id].moveN++; }
    }
  }
  // 인원·피해 상태 (개체별 가장 심한 피해 기준)
  for (const e of entities.values()) {
    for (const id of e.chain) {
      const n = nodes[id];
      n.personnel++;
      if (e.sev) n.cas[SEV_NAME[e.sev]]++;
    }
  }

  // 탐지기록
  for (const r of tables.탐지기록) {
    const ch = unitChain(r);
    const idSec = r.식별시각 != null && r.탐지시각 != null ? (r.식별시각 - r.탐지시각) / 1000 : null;
    for (const id of ch) {
      const n = nodes[id];
      n.det++;
      if (idSec != null && idSec >= 0) { n.detIdSum += idSec; n.detIdN++; }
      n.detPreKnown++;
      if (r.상대탐지시각 == null || r.탐지시각 < r.상대탐지시각) n.detPre++;
    }
  }

  // 지휘통신: 발신·수신 제대 양쪽(상위 제대 포함)에 귀속
  const splitPath = s => String(s || '').split(/[\/>\-,]+/).map(x => x.trim()).filter(Boolean);
  for (const r of tables.지휘통신) {
    const ids = new Set([...chain(r.진영, splitPath(r.발신제대)), ...chain(r.진영, splitPath(r.수신제대))]);
    const sec = r.응답시각 != null ? (r.응답시각 - r.시각) / 1000 : null;
    const isOrder = /명령|지시|order/i.test(r.유형);
    for (const id of ids) {
      const n = nodes[id];
      if (isOrder) { n.orders++; if (sec != null && sec >= 0) { n.orderSecSum += sec; n.orderN++; } }
      else { n.reports++; if (sec != null && sec >= 0) { n.reportN++; if (sec > cfg.보고지연기준초) n.reportDelayed++; } }
    }
  }

  // 지속지원
  for (const r of tables.지속지원) {
    const ch = unitChain(r);
    const min = r.완료시각 != null ? (r.완료시각 - r.요청시각) / 60000 : null;
    for (const id of ch) {
      const n = nodes[id];
      if (r.탄약소모량 != null) n.ammoUse += r.탄약소모량;
      if (/탄약/.test(r.유형)) {
        n.ammoSupply += r.수량 || 0;
        if (min != null && min >= 0) { n.resupSum += min; n.resupN++; }
      } else if (/후송/.test(r.유형)) {
        if (min != null && min >= 0) { n.evacSum += min; n.evacN++; }
      }
    }
  }

  for (const id in nodes) Engine.computeMetrics(nodes[id], tables);

  const model = {
    분석기준버전: cfg.분석기준버전,
    분석일시: U.nowStr(),
    params: { 판정기준N: cfg.판정기준N, 보고지연기준초: cfg.보고지연기준초, 교전간격분: cfg.교전간격분, 초기피해구간분: cfg.초기피해구간분, 시간대단위분: cfg.시간대단위분 },
    t0, binMin: cfg.시간대단위분, nBins,
    tStart: tMin, tEnd: tMax,
    nodes,
  };
  // 빈 진영 루트 제거
  for (const s of ['훈련부대', '대항군']) if (!nodes[ROOT(s)].children.length && !nodes[ROOT(s)].shots) delete nodes[ROOT(s)];
  model.levelStats = Engine.levelStats(model);
  return Engine.sortChildren(model);
};

/* 하위 제대를 이름순(1대대, 2대대, …)으로 정렬 */
Engine.sortChildren = function (model) {
  if (!model) return model;
  const nodes = model.nodes;
  const cmp = (a, b) => (nodes[a].side === nodes[b].side ? 0 : nodes[a].side === '훈련부대' ? -1 : 1) ||
    String(nodes[a].name).localeCompare(String(nodes[b].name), 'ko', { numeric: true });
  for (const n of Object.values(nodes)) n.children = n.children.filter(id => nodes[id]).sort(cmp);
  return model;
};

Engine.computeMetrics = function (n, tables) {
  const r = (a, b) => (b > 0 ? a / b : null);
  const m = n.metrics;
  m.c2_order_sec = r(n.orderSecSum, n.orderN);
  m.c2_report_delay = r(n.reportDelayed, n.reportN);
  m.c2_max_gap = n.maxGapMin;
  m.int_id_sec = r(n.detIdSum, n.detIdN);
  m.int_predetect = r(n.detPre, n.detPreKnown);
  m.man_speed = r(n.moveSpeedSum, n.moveN);
  m.man_move_dmg = r(n.dmgMoving, n.dmgMoveKnown);
  m.fire_hit = r(n.hits, n.shots);
  m.fire_dmg_per_shot = r(n.dmgHits, n.shots);
  m.fire_indirect = r(n.indirectDmgHits, n.dmgHits);
  m.prot_survival = n.personnel > 0 ? (n.personnel - n.cas.사망) / n.personnel : null;
  m.prot_uncovered = r(n.dmgUncovered, n.dmgCoverKnown);
  m.prot_early = r(n.dmgEarly, n.dmgRows);
  // 탄약 소모량이 지속지원 시트에 없으면 교전기록의 사격 수로 대체
  m.sus_ammo = r(n.ammoSupply, n.ammoUse > 0 ? n.ammoUse : n.shots);
  m.sus_evac_min = r(n.evacSum, n.evacN);
  m.sus_resup_min = r(n.resupSum, n.resupN);
};

/* 비교 기준용: 진영·제대 수준별 지표 평균 (루트는 진영 필터별 값 그대로) */
Engine.levelStats = function (model) {
  const out = {};
  const groups = {};
  for (const n of Object.values(model.nodes)) {
    if (n.level === 0) { out['ROOT|' + n.side] = Object.assign({}, n.metrics); continue; }
    (groups[n.side + '|' + n.level] || (groups[n.side + '|' + n.level] = [])).push(n);
  }
  for (const [k, arr] of Object.entries(groups)) {
    const o = {};
    for (const ind of DEFAULT_INDICATORS) o[ind.id] = U.mean(arr.map(n => n.metrics[ind.id]));
    out[k] = o;
  }
  return out;
};
Engine.levelKey = (node, sideFilter) => (node.level === 0 ? 'ROOT|' + (sideFilter || node.side) : node.side + '|' + node.level);

/* ---------- 3. 6대 기능 진단 ---------- */
/* prevStats: [levelStats(이전 훈련들)] */
Engine.baseline = function (key, prevStats) {
  const vals = prevStats.map(s => s[key]).filter(Boolean);
  if (!vals.length) return null;
  const base = {};
  for (const ind of DEFAULT_INDICATORS) base[ind.id] = U.mean(vals.map(v => v[ind.id]));
  base._n = vals.length;
  return base;
};

Engine.indIndex = function (ind, v, b) {
  if (v == null || b == null || !isFinite(v) || !isFinite(b) || b === 0) return null;
  let x = ind.dir === 'low' ? (v === 0 ? 200 : (100 * b) / v) : (100 * v) / b;
  return Math.max(0, Math.min(200, x));
};

Engine.diagnose = function (metrics, base, cfg) {
  const N = cfg.판정기준N;
  const result = { hasBase: !!base, baseN: base ? base._n : 0, functions: [] };
  for (const fn of FUNCTIONS) {
    const inds = cfg.indicators.filter(i => i.fn === fn);
    const rows = inds.map(ind => {
      const v = metrics[ind.id];
      const b = base ? base[ind.id] : null;
      const row = { id: ind.id, name: ind.name, value: v, base: b, judged: ind.judged !== false, change: null, improve: null, index: null, judge: '' };
      if (!base) row.judge = '비교 기준 없음';
      else if (v == null) row.judge = '자료 없음';
      else if (b == null || b === 0) row.judge = '비교 불가';
      else {
        row.change = ((v - b) / Math.abs(b)) * 100;
        row.improve = ind.dir === 'low' ? -row.change : row.change;
        row.index = Engine.indIndex(ind, v, b);
        row.judge = !row.judged ? '참고' : row.improve >= N ? '강점' : row.improve <= -N ? '취약' : '보통';
      }
      row.sentence = Engine.indSentence(fn, ind, row);
      return row;
    });
    const idxs = rows.filter(r => r.judged && r.index != null).map(r => r.index);
    const index = idxs.length ? U.mean(idxs) : null;
    const judge = !base ? '비교 기준 없음' : index == null ? '자료 없음' : index >= 100 + N ? '강점' : index <= 100 - N ? '취약' : '보통';
    result.functions.push({ fn, index, judge, indicators: rows, summary: Engine.fnSentence(fn, index, judge, rows) });
  }
  return result;
};

Engine.indSentence = function (fn, ind, r) {
  const val = U.fmtInd(ind, r.value);
  if (r.base == null || r.change == null) return `[${fn}] ${ind.name} ${val}` + (r.judge === '비교 기준 없음' ? ' (비교 기준 없음)' : '');
  const diff = ind.fmt === 'pct'
    ? `${r.value - r.base >= 0 ? '+' : ''}${((r.value - r.base) * 100).toFixed(1)}%p`
    : `${r.change >= 0 ? '+' : ''}${r.change.toFixed(1)}%`;
  let s = `[${fn}] ${ind.name} ${val} (이전 평균 ${U.fmtInd(ind, r.base)}, ${diff})`;
  if (r.judge === '참고') return s + ' → 참고 지표';
  s += ` → ${r.judge}.`;
  if (r.judge === '강점') s += ` ${ind.name} ${ind.dir === 'low' ? '단축/감소' : '향상'}이 주요 성과 요인으로 판단됨`;
  if (r.judge === '취약') s += ` ${ind.name} ${ind.dir === 'low' ? '증가' : '저하'}가 주요 원인으로 판단됨`;
  return s;
};

Engine.fnSentence = function (fn, index, judge, rows) {
  if (judge === '비교 기준 없음') return `[${fn}] 이전 훈련이 없어 지표 수치만 표시합니다.`;
  if (index == null) return `[${fn}] 판정에 필요한 자료가 없습니다.`;
  const judged = rows.filter(r => r.judged && r.improve != null);
  const best = judged.slice().sort((a, b) => b.improve - a.improve)[0];
  const worst = judged.slice().sort((a, b) => a.improve - b.improve)[0];
  const fmtImp = r => `${r.improve >= 0 ? '개선' : '악화'} ${Math.abs(r.improve).toFixed(0)}%`;
  let s = `[${fn}] 종합지수 ${index.toFixed(0)} (이전 평균=100) → ${judge}.`;
  if (judge === '강점') s += ` ${best.name}(${fmtImp(best)})이(가) 주요 성과 요인으로 판단됨.`;
  else if (judge === '취약') s += ` ${worst.name}(${fmtImp(worst)})이(가) 주요 원인으로 판단됨.`;
  else s += ' 이전 훈련과 비슷한 수준.';
  const others = judged.filter(r => (judge !== '강점' && r.judge === '강점' && r !== best) || (judge !== '취약' && r.judge === '취약' && r !== worst));
  if (judge === '보통') {
    const st = judged.filter(r => r.judge === '강점'), wk = judged.filter(r => r.judge === '취약');
    if (st.length) s += ` 강점: ${st.map(r => r.name).join(', ')}.`;
    if (wk.length) s += ` 취약: ${wk.map(r => r.name).join(', ')}.`;
  } else if (others.length) {
    s += ` 단, ${others.map(r => `${r.name} ${r.judge}`).join(', ')}.`;
  }
  return s;
};
