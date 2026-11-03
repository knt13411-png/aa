'use strict';
/* 기본 설정: 시트/열 정의, 지표 정의, 진단 기준값.
   설정 화면에서 수정한 값은 브라우저(localStorage)에 저장되고 JSON으로 내보내기/불러오기 할 수 있다. */

const FUNCTIONS = ['지휘통제', '정보', '기동', '화력', '방호', '작전지속지원'];
const LEVELS = ['훈련', '여단', '대대', '중대', '소대'];
const SIDES = ['전체', '훈련부대', '대항군'];

const UNIT_COLS = ['진영', '여단', '대대', '중대', '소대'];

/* 시트별 표준 열. type: time | num | text, req: 필수 여부 */
const SHEET_DEFS = {
  교전기록: {
    fn: '화력, 지휘통제',
    cols: {
      시각: { type: 'time', req: true }, 진영: { type: 'text', req: true }, 여단: { type: 'text', req: true },
      대대: { type: 'text' }, 중대: { type: 'text' }, 소대: { type: 'text' },
      사수ID: { type: 'text' }, 화기종류: { type: 'text', req: true }, 화기구분: { type: 'text', req: true },
      표적ID: { type: 'text' }, 표적진영: { type: 'text' }, 결과: { type: 'text', req: true },
      피해판정: { type: 'text', req: true }, '거리(m)': { type: 'num' },
    },
  },
  탐지기록: {
    fn: '정보',
    cols: {
      시각: { type: 'time' }, 진영: { type: 'text', req: true }, 여단: { type: 'text', req: true },
      대대: { type: 'text' }, 중대: { type: 'text' }, 소대: { type: 'text' },
      탐지수단: { type: 'text' }, 표적ID: { type: 'text' },
      탐지시각: { type: 'time', req: true }, 식별시각: { type: 'time' }, 상대탐지시각: { type: 'time' },
    },
  },
  위치기록: {
    fn: '기동',
    cols: {
      시각: { type: 'time', req: true }, 진영: { type: 'text', req: true }, 여단: { type: 'text', req: true },
      대대: { type: 'text' }, 중대: { type: 'text' }, 소대: { type: 'text' },
      개체ID: { type: 'text', req: true }, X좌표: { type: 'num' }, Y좌표: { type: 'num' },
      '속도(km/h)': { type: 'num', req: true }, 이동여부: { type: 'text', req: true },
    },
  },
  피해기록: {
    fn: '방호, 기동',
    cols: {
      시각: { type: 'time', req: true }, 진영: { type: 'text', req: true }, 여단: { type: 'text', req: true },
      대대: { type: 'text' }, 중대: { type: 'text' }, 소대: { type: 'text' },
      개체ID: { type: 'text', req: true }, 피해판정: { type: 'text', req: true }, 피해원인: { type: 'text' },
      엄폐여부: { type: 'text' }, 이동여부: { type: 'text' },
    },
  },
  지휘통신: {
    fn: '지휘통제',
    cols: {
      시각: { type: 'time', req: true }, 진영: { type: 'text', req: true },
      발신제대: { type: 'text', req: true }, 수신제대: { type: 'text', req: true },
      유형: { type: 'text', req: true }, 관련메시지ID: { type: 'text' }, 응답시각: { type: 'time' },
    },
  },
  지속지원: {
    fn: '작전지속지원',
    cols: {
      요청시각: { type: 'time', req: true }, 완료시각: { type: 'time' }, 진영: { type: 'text', req: true },
      여단: { type: 'text', req: true }, 대대: { type: 'text' }, 중대: { type: 'text' }, 소대: { type: 'text' },
      유형: { type: 'text', req: true }, 수량: { type: 'num' }, 탄약소모량: { type: 'num' },
    },
  },
};

/* 지표 정의. 계산식은 id로 engine.js에 고정되어 있고, 이름·방향·판정 포함 여부·표시 형식은 설정으로 바꿀 수 있다.
   dir: high(높을수록 좋음) | low(낮을수록 좋음), judged:false 는 참고 지표, fmt: pct | num1 | num2 */
const DEFAULT_INDICATORS = [
  { id: 'c2_order_sec', fn: '지휘통제', name: '명령 하달 평균 소요시간', unit: '초', dir: 'low', judged: true, fmt: 'num1', src: '지휘통신' },
  { id: 'c2_report_delay', fn: '지휘통제', name: '보고 지연 비율', unit: '%', dir: 'low', judged: true, fmt: 'pct', src: '지휘통신' },
  { id: 'c2_max_gap', fn: '지휘통제', name: '최장 교전 공백 시간', unit: '분', dir: 'low', judged: true, fmt: 'num1', src: '교전기록' },
  { id: 'int_id_sec', fn: '정보', name: '평균 탐지→식별 소요시간', unit: '초', dir: 'low', judged: true, fmt: 'num1', src: '탐지기록' },
  { id: 'int_predetect', fn: '정보', name: '선탐지 비율', unit: '%', dir: 'high', judged: true, fmt: 'pct', src: '탐지기록' },
  { id: 'man_speed', fn: '기동', name: '평균 이동속도', unit: 'km/h', dir: 'high', judged: true, fmt: 'num1', src: '위치기록' },
  { id: 'man_move_dmg', fn: '기동', name: '기동 중 피해 비율', unit: '%', dir: 'low', judged: true, fmt: 'pct', src: '피해기록' },
  { id: 'fire_hit', fn: '화력', name: '명중률', unit: '%', dir: 'high', judged: true, fmt: 'pct', src: '교전기록' },
  { id: 'fire_dmg_per_shot', fn: '화력', name: '사격 1회당 피해 유발 수', unit: '', dir: 'high', judged: true, fmt: 'num3', src: '교전기록' },
  { id: 'fire_indirect', fn: '화력', name: '간접화력 피해 비율', unit: '%', dir: 'high', judged: false, fmt: 'pct', src: '교전기록' },
  { id: 'prot_survival', fn: '방호', name: '생존율', unit: '%', dir: 'high', judged: true, fmt: 'pct', src: '피해기록' },
  { id: 'prot_uncovered', fn: '방호', name: '비엄폐 상태 피해 비율', unit: '%', dir: 'low', judged: true, fmt: 'pct', src: '피해기록' },
  { id: 'prot_early', fn: '방호', name: '초기 피해 집중도', unit: '%', dir: 'low', judged: true, fmt: 'pct', src: '피해기록' },
  { id: 'sus_ammo', fn: '작전지속지원', name: '탄약 보급률', unit: '%', dir: 'high', judged: true, fmt: 'pct', src: '지속지원' },
  { id: 'sus_evac_min', fn: '작전지속지원', name: '평균 후송 소요시간', unit: '분', dir: 'low', judged: true, fmt: 'num1', src: '지속지원' },
  { id: 'sus_resup_min', fn: '작전지속지원', name: '평균 재보급 소요시간', unit: '분', dir: 'low', judged: true, fmt: 'num1', src: '지속지원' },
];

function defaultConfig() {
  const sheets = {};
  for (const [s, def] of Object.entries(SHEET_DEFS)) {
    const cols = {};
    for (const c of Object.keys(def.cols)) cols[c] = c;
    sheets[s] = { sheetName: s, columns: cols };
  }
  return {
    분석기준버전: '1.0',
    판정기준N: 15,            // 이전 훈련 평균 대비 ±N% → 강점/취약
    보고지연기준초: 180,       // 보고 소요시간이 이 값을 넘으면 지연
    교전간격분: 5,            // 같은 소대의 사격 간격이 이 값을 넘으면 별개 교전
    초기피해구간분: 30,        // 교전 개시 후 이 시간 내 피해 → 초기 피해
    시간대단위분: 30,          // 시간대별 추이 묶음 단위
    저장방식: 'auto',          // auto | A | B
    sheets,
    indicators: JSON.parse(JSON.stringify(DEFAULT_INDICATORS)),
  };
}

const Config = {
  KEY: 'aar.config.v1',
  cur: null,
  load() {
    let saved = null;
    try { saved = JSON.parse(localStorage.getItem(this.KEY) || 'null'); } catch (e) { saved = null; }
    this.cur = this.merge(defaultConfig(), saved);
    return this.cur;
  },
  merge(base, saved) {
    if (!saved || typeof saved !== 'object') return base;
    const out = Object.assign({}, base, saved);
    out.sheets = JSON.parse(JSON.stringify(base.sheets));
    if (saved.sheets) {
      for (const [s, v] of Object.entries(saved.sheets)) {
        if (!out.sheets[s]) continue;
        if (v.sheetName) out.sheets[s].sheetName = v.sheetName;
        Object.assign(out.sheets[s].columns, v.columns || {});
      }
    }
    // 지표는 id 기준으로 병합 (계산식이 없는 id는 무시)
    const byId = {};
    (saved.indicators || []).forEach(i => { if (i && i.id) byId[i.id] = i; });
    out.indicators = base.indicators.map(i => Object.assign({}, i, byId[i.id] || {}, { id: i.id }));
    return out;
  },
  save(cfg) {
    this.cur = this.merge(defaultConfig(), cfg);
    try { localStorage.setItem(this.KEY, JSON.stringify(this.cur)); } catch (e) { /* 저장 불가 환경 */ }
  },
  reset() {
    try { localStorage.removeItem(this.KEY); } catch (e) { /* ignore */ }
    this.cur = defaultConfig();
  },
  ind(id) { return this.cur.indicators.find(i => i.id === id); },
};
