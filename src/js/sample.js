'use strict';
/* 입력 양식 샘플(.xlsx): 작성안내 시트 + 분석기가 읽는 6개 시트(SHEET_DEFS 그대로)에 가상 훈련 데이터를 채운다.
   난수는 고정 시드라 언제 만들어도 같은 파일이 나온다. 모두 가상 데이터. */

const Sample = {};
Sample.fileName = '사후검토_입력양식_샘플.xlsx';

/* 열 설명: [형식, 설명, 예시] */
const COL_GUIDE = {
  시각: ['날짜+시각', '기록 시각', '2026-11-01 06:05:00'],
  진영: ['글자', '"훈련부대" 또는 "대항군"', '훈련부대'],
  여단: ['글자', '소속 여단', '1여단'], 대대: ['글자', '소속 대대 (없으면 비움)', '1대대'],
  중대: ['글자', '소속 중대 (없으면 비움)', '1중대'], 소대: ['글자', '소속 소대 (없으면 비움)', '1소대'],
  사수ID: ['글자', '사격한 개체(인원·장비) ID', 'T1-1-1-01'],
  화기종류: ['글자', '화기 이름', '소총'], 화기구분: ['글자', '"직사" 또는 "간접"', '직사'],
  표적ID: ['글자', '표적 개체 ID', 'O1-1-1-03'], 표적진영: ['글자', '표적의 진영', '대항군'],
  결과: ['글자', '"명중" 또는 "빗나감"', '명중'], 피해판정: ['글자', '"사망" / "중상" / "경상" / "없음"', '경상'],
  '거리(m)': ['숫자', '사격 거리(m)', '320'],
  탐지수단: ['글자', '육안·열상·레이더·UAV 등', '열상'],
  탐지시각: ['날짜+시각', '표적을 처음 탐지한 시각', '2026-11-01 06:40:10'],
  식별시각: ['날짜+시각', '표적을 식별(피아 확인)한 시각', '2026-11-01 06:41:05'],
  상대탐지시각: ['날짜+시각', '상대가 우리를 탐지한 시각 (비우면 선탐지로 봄)', '2026-11-01 06:42:00'],
  개체ID: ['글자', '인원·장비 ID', 'T1-1-1-01'], X좌표: ['숫자', '위치 X', '1520'], Y좌표: ['숫자', '위치 Y', '830'],
  '속도(km/h)': ['숫자', '이동 속도', '8.5'], 이동여부: ['글자', '"Y"(이동) 또는 "N"(정지)', 'Y'],
  피해원인: ['글자', '피해를 준 화기·원인', '기관총'], 엄폐여부: ['글자', '"Y"(엄폐) 또는 "N"(비엄폐)', 'N'],
  발신제대: ['글자', '보낸 제대, 여단/대대/중대/소대를 "/"로 연결', '1여단/1대대'],
  수신제대: ['글자', '받은 제대, 같은 형식', '1여단/1대대/1중대'],
  유형: ['글자', '지휘통신: "명령" 또는 "보고" / 지속지원: "탄약보급" 또는 "후송"', '명령'],
  관련메시지ID: ['글자', '명령 ID (보고는 응답한 명령 ID)', 'M00012'],
  응답시각: ['날짜+시각', '수신 제대가 응답한 시각', '2026-11-01 07:02:30'],
  요청시각: ['날짜+시각', '지원 요청 시각', '2026-11-01 08:10:00'], 완료시각: ['날짜+시각', '지원 완료 시각', '2026-11-01 08:52:00'],
  수량: ['숫자', '보급 수량(탄약보급) 또는 인원 수(후송)', '400'], 탄약소모량: ['숫자', '해당 시점까지 탄약 소모량', '520'],
};

Sample.build = function () {
  let seed = 20261101;
  const rnd = () => { // mulberry32
    seed = (seed + 0x6D2B79F5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const uni = (a, b) => a + rnd() * (b - a);
  const pick = arr => arr[Math.floor(rnd() * arr.length)];
  const t0 = new Date(2026, 10, 1, 6, 0, 0);
  const DUR = 240; // 분
  const at = min => new Date(t0.getTime() + Math.round(min * 60) * 1000);
  const addS = (d, s) => new Date(d.getTime() + Math.round(s) * 1000);

  // 편성: 훈련부대 1여단 2개 대대(2대대는 숙련도 낮게), 대항군 1개 대대
  const units = [];
  for (const bn of ['1대대', '2대대']) for (const co of ['1중대', '2중대']) for (const pl of ['1소대', '2소대'])
    units.push({ side: '훈련부대', bde: '1여단', bn, co, pl, skill: bn === '1대대' ? 1.15 : 0.85, pre: 'T' });
  for (const co of ['1중대', '2중대']) for (const pl of ['1소대', '2소대'])
    units.push({ side: '대항군', bde: '대항군여단', bn: '1대대', co, pl, skill: 1.0, pre: 'O' });
  for (const u of units) {
    const k = `${u.pre}${u.bn[0]}-${u.co[0]}-${u.pl[0]}`;
    u.ents = Array.from({ length: 9 }, (_, i) => `${k}-${String(i + 1).padStart(2, '0')}`);
    u.x = u.side === '훈련부대' ? uni(0, 2000) : uni(6000, 8000); u.y = uni(0, 3000);
  }
  const uc = u => [u.side, u.bde, u.bn, u.co, u.pl];
  const sev = {}; // 개체 → 가장 심한 피해
  const SEVN = { 경상: 1, 중상: 2, 사망: 3 };
  const enemyOf = u => units.filter(x => x.side !== u.side);
  const W = [['소총', '직사', 0.55, 1.0, [50, 400]], ['기관총', '직사', 0.25, 0.85, [100, 800]], ['대전차', '직사', 0.08, 1.2, [200, 1500]], ['박격포', '간접', 0.12, 0.9, [500, 3000]]];
  const pickW = () => { let r = rnd(), acc = 0; for (const w of W) { acc += w[2]; if (r <= acc) return w; } return W[0]; };

  const fire = [], dmg = [], det = [], pos = [], c2 = [], sus = [];
  // 교전기록 + 피해기록
  for (const u of units) {
    const bursts = Math.floor(uni(4, 7));
    for (let b = 0; b < bursts; b++) {
      const start = uni(10, DUR - 10);
      const shots = Math.floor(uni(8, 22));
      for (let s = 0; s < shots; s++) {
        const t = at(start + uni(0, 4));
        const w = pickW();
        const tu = pick(enemyOf(u));
        const tgt = pick(tu.ents);
        const hit = rnd() < 0.3 * u.skill * w[3];
        let res = '없음';
        if (hit && (sev[tgt] || 0) < 3 && rnd() < 0.45) {
          const q = rnd();
          res = q < 0.3 ? '사망' : q < 0.6 ? '중상' : '경상';
          if (SEVN[res] > (sev[tgt] || 0)) sev[tgt] = SEVN[res];
          const cover = rnd() < (tu.side === '훈련부대' ? 0.55 / tu.skill : 0.5) ? 'N' : 'Y';
          dmg.push([t, ...uc(tu), tgt, res, w[0], cover, rnd() < 0.4 ? 'Y' : 'N']);
        }
        fire.push([t, ...uc(u), pick(u.ents), w[0], w[1], tgt, tu.side, hit ? '명중' : '빗나감', res, Math.round(uni(w[4][0], w[4][1]))]);
      }
    }
  }
  // 탐지기록
  for (const u of units) {
    for (let i = 0, n = Math.floor(uni(8, 13)); i < n; i++) {
      const td = at(uni(5, DUR - 5));
      const opp = rnd() < 0.2 ? null : addS(td, uni(-120, 180 * u.skill));
      det.push([td, ...uc(u), pick(['육안', '육안', '열상', '레이더', 'UAV']), pick(pick(enemyOf(u)).ents), td, addS(td, uni(20, 120) / u.skill), opp]);
    }
  }
  // 위치기록: 개체별 10분 간격
  for (const u of units) {
    for (const e of u.ents) {
      let x = u.x + uni(-80, 80), y = u.y + uni(-80, 80);
      for (let m = 0; m <= DUR; m += 10) {
        const moving = rnd() < 0.5;
        const v = moving ? +(uni(4, 12) * (u.side === '훈련부대' ? u.skill : 1)).toFixed(1) : 0;
        if (moving) { x += uni(-50, 250) * (u.side === '훈련부대' ? 1 : -1); y += uni(-80, 80); }
        pos.push([at(m), ...uc(u), e, Math.round(x), Math.round(y), v, moving ? 'Y' : 'N']);
      }
    }
  }
  // 지휘통신: 상급 → 하급 명령, 하급 → 상급 보고
  let mid = 0;
  const pairs = [];
  for (const u of units) {
    const bn = `${u.bde}/${u.bn}`, co = `${bn}/${u.co}`, pl = `${co}/${u.pl}`;
    pairs.push([u, co, pl]);
    if (u.pl === '1소대') pairs.push([u, bn, co]);
    if (u.pl === '1소대' && u.co === '1중대') pairs.push([u, u.bde, bn]);
  }
  for (const [u, up, down] of pairs) {
    const ids = [];
    for (let k = 0, n = Math.floor(uni(3, 6)); k < n; k++) {
      const id = 'M' + String(++mid).padStart(5, '0'); ids.push(id);
      const t = at(uni(0, DUR - 15));
      c2.push([t, u.side, up, down, '명령', id, addS(t, uni(30, 160) / u.skill)]);
    }
    for (let k = 0, n = Math.floor(uni(3, 6)); k < n; k++) {
      const t = at(uni(0, DUR - 15));
      c2.push([t, u.side, down, up, '보고', pick(ids), addS(t, uni(40, 260) / u.skill)]);
    }
  }
  // 지속지원: 탄약보급 2~3회, 피해 인원이 있으면 후송
  for (const u of units) {
    let used = 0;
    for (let k = 0, n = Math.floor(uni(2, 4)); k < n; k++) {
      const t = at(uni(30, DUR - 30));
      used += Math.round(uni(150, 400));
      sus.push([t, addS(t, uni(15, 60) * 60 / u.skill), ...uc(u), '탄약보급', Math.round(used * uni(0.6, 1.0)), used]);
    }
    const hurt = u.ents.filter(e => (sev[e] || 0) >= 1 && sev[e] < 3).length;
    for (let k = 0; k < Math.min(hurt, 3); k++) {
      const t = at(uni(40, DUR - 20));
      sus.push([t, addS(t, uni(25, 70) * 60 / u.skill), ...uc(u), '후송', 1, null]);
    }
  }
  const byTime = (a, b) => a[0] - b[0];
  [fire, dmg, det, pos, c2, sus].forEach(t => t.sort(byTime));

  const wb = XLSX.utils.book_new();
  // 작성안내
  const guide = [
    ['사후검토 분석기 입력 양식 (샘플)'],
    ['· 아래 6개 시트를 이 이름 그대로 두고 열 이름도 바꾸지 마세요. 이 "작성안내" 시트는 분석에 쓰지 않으므로 지워도 됩니다.'],
    ['· 시각은 날짜와 시각을 함께 적습니다 (예: 2026-11-01 06:05:00). 엑셀 날짜 셀이나 글자 모두 됩니다.'],
    ['· 필수 열이 비었거나 형식이 틀린 행은 자동으로 빼고 분석합니다. 없는 시트는 관련 지표가 "자료 없음"이 됩니다.'],
    ['· 파일 이름이 훈련명이 됩니다 (예: 2026-11_백두4호훈련.xlsx → 백두4호훈련). 이 샘플의 값은 모두 가상 데이터입니다.'],
    [],
    ['시트', '관련 기능', '열', '형식', '필수', '설명', '예시'],
  ];
  for (const [sheet, def] of Object.entries(SHEET_DEFS)) {
    for (const [col, cd] of Object.entries(def.cols)) {
      const g = COL_GUIDE[col] || ['', '', ''];
      guide.push([sheet, def.fn, col, g[0], cd.req ? '필수' : '', g[1], g[2]]);
    }
  }
  const gws = XLSX.utils.aoa_to_sheet(guide);
  gws['!cols'] = [{ wch: 10 }, { wch: 14 }, { wch: 13 }, { wch: 10 }, { wch: 6 }, { wch: 58 }, { wch: 22 }];
  gws['!merges'] = [0, 1, 2, 3, 4].map(r => ({ s: { r, c: 0 }, e: { r, c: 6 } }));
  XLSX.utils.book_append_sheet(wb, gws, '작성안내');

  const data = { 교전기록: fire, 탐지기록: det, 위치기록: pos, 피해기록: dmg, 지휘통신: c2, 지속지원: sus };
  for (const [sheet, def] of Object.entries(SHEET_DEFS)) {
    const head = Object.keys(def.cols);
    const ws = XLSX.utils.aoa_to_sheet([head].concat(data[sheet]), { dateNF: 'yyyy-mm-dd hh:mm:ss' });
    ws['!cols'] = head.map(h => ({ wch: def.cols[h].type === 'time' ? 20 : Math.max(8, h.length * 2 + 2) }));
    XLSX.utils.book_append_sheet(wb, ws, sheet);
  }
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};
