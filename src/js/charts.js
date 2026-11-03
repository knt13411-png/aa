'use strict';
/* 차트 공통 설정, 값 라벨 플러그인, 화면·보고서 공용 차트 구성, 보고서용 PNG 렌더링 */

const COL = {
  series: ['#2a5db0', '#d9622b', '#1a9b6c', '#d79a00', '#c75b8f'],
  blue: '#2a5db0', red: '#c03a32',
  사망: '#b8312f', 중상: '#ec835a', 경상: '#f2b632', 무피해: '#cfcec6',
  ink: '#0b0b0b', ink2: '#52514e', muted: '#898781', grid: '#e1e0d9', axis: '#c3c2b7', surface: '#ffffff',
  good: '#0a7a0a', bad: '#c03030', mid: '#8a8984',
};
const FONT = "'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', system-ui, sans-serif";

const Charts = {};

Charts.setup = function () {
  Chart.defaults.font.family = FONT;
  Chart.defaults.color = COL.ink2;
  Chart.register({
    id: 'valueLabels',
    afterDatasetsDraw(chart, args, opts) {
      if (!opts || !opts.enabled) return;
      const ctx = chart.ctx;
      const horiz = chart.options.indexAxis === 'y';
      ctx.save();
      ctx.font = `600 ${opts.size || 12}px ${FONT}`;
      ctx.fillStyle = COL.ink;
      ctx.textAlign = horiz ? 'left' : 'center';
      ctx.textBaseline = horiz ? 'middle' : 'bottom';
      chart.data.datasets.forEach((ds, i) => {
        const meta = chart.getDatasetMeta(i);
        if (meta.hidden || meta.type !== 'bar') return;
        meta.data.forEach((el, j) => {
          const v = ds.data[j];
          if (v == null) return;
          const t = opts.fmt ? opts.fmt(v) : String(v);
          if (horiz) ctx.fillText(t, el.x + 5, el.y); else ctx.fillText(t, el.x, el.y - 3);
        });
      });
      ctx.restore();
    },
  });
  Chart.register({
    id: 'bgFill',
    beforeDraw(chart, args, opts) {
      if (!opts || !opts.color) return;
      const ctx = chart.ctx;
      ctx.save(); ctx.fillStyle = opts.color; ctx.fillRect(0, 0, chart.width, chart.height); ctx.restore();
    },
  });
};

function merge(a, b) {
  if (!b) return a;
  for (const k of Object.keys(b)) {
    const v = b[k];
    if (v && typeof v === 'object' && !Array.isArray(v) && typeof v !== 'function' && a[k] && typeof a[k] === 'object') merge(a[k], v);
    else a[k] = v;
  }
  return a;
}
Charts.merge = merge;

/* 막대/선 공통 옵션. s = 글자 배율(화면 1, 보고서 이미지 1.6) */
Charts.opts = function (s, extra) {
  const f = Math.round(12 * s);
  return merge({
    responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
    layout: { padding: { top: Math.round(14 * s), right: Math.round(8 * s) } },
    datasets: { bar: { maxBarThickness: Math.round(56 * s) } },
    plugins: {
      legend: { position: 'top', align: 'end', labels: { font: { size: f }, boxWidth: f, boxHeight: f, color: COL.ink2 } },
      tooltip: { titleFont: { size: f }, bodyFont: { size: f }, footerFont: { size: f } },
    },
    scales: {
      x: { grid: { display: false }, border: { color: COL.axis }, ticks: { font: { size: f }, color: COL.ink2 } },
      y: { grid: { color: COL.grid }, border: { display: false }, ticks: { font: { size: f }, color: COL.muted }, beginAtZero: true },
    },
  }, extra);
};

Charts.bar = (label, data, color, extra) => Object.assign({
  type: 'bar', label, data, backgroundColor: color,
  borderRadius: 4, borderSkipped: 'start', categoryPercentage: 0.7, barPercentage: 0.9,
}, extra || {});

const pct1 = v => (v == null ? null : +(v * 100).toFixed(1));
const JUDGE_COL = j => (j === '강점' || j === '우수' ? COL.good : j === '취약' || j === '미흡' ? COL.bad : COL.mid);

/* ---------- 화면·보고서 공용 차트 구성 (R = Insight.build 결과) ---------- */
Charts.defs = {
  radar(R, s) {
    const vals = R.diag.functions.map(f => (f.index == null ? null : +f.index.toFixed(0)));
    const mx = Math.max(150, ...vals.filter(v => v != null).map(v => Math.ceil(v / 25) * 25));
    return {
      type: 'radar',
      data: {
        labels: FUNCTIONS, datasets: [
          { label: R.sideName, data: vals, borderColor: COL.blue, backgroundColor: 'rgba(42,93,176,0.16)', borderWidth: 2.5 * s, pointRadius: 4 * s, pointBackgroundColor: vals.map((v, i) => JUDGE_COL(R.diag.functions[i].judge)) },
          { label: `${R.baseName} (=100)`, data: FUNCTIONS.map(() => 100), borderColor: COL.muted, borderDash: [6 * s, 4 * s], borderWidth: 1.5 * s, pointRadius: 0, fill: false },
        ],
      },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
        plugins: { legend: { position: 'top', labels: { font: { size: 12 * s }, boxWidth: 12 * s } }, tooltip: { bodyFont: { size: 12 * s } } },
        scales: { r: { min: 0, max: mx, ticks: { stepSize: 25, font: { size: 10 * s }, backdropColor: 'transparent', color: COL.muted }, grid: { color: COL.grid }, angleLines: { color: COL.grid }, pointLabels: { font: { size: 14 * s, weight: '700' }, color: COL.ink } } },
      },
    };
  },

  /* 피아 핵심 비율 지표 */
  versus(R, s) {
    const ids = ['fire_hit', 'prot_survival', 'int_predetect', 'prot_uncovered', 'man_move_dmg'];
    const labels = ids.map(id => Config.ind(id).name + (Config.ind(id).dir === 'low' ? ' ↓' : ''));
    const ds = [Charts.bar(R.sideName, ids.map(id => pct1(R.root.metrics[id])), COL.blue)];
    if (R.opfor) ds.push(Charts.bar('대항군', ids.map(id => pct1(R.opfor.metrics[id])), COL.red));
    return {
      type: 'bar', data: { labels, datasets: ds },
      options: Charts.opts(s, {
        plugins: { valueLabels: { enabled: true, size: 11 * s, fmt: v => v + '%' }, subtitle: { display: true, text: '↓ 표시는 낮을수록 좋음', align: 'start', font: { size: 11 * s }, color: COL.muted } },
        scales: { y: { max: 100, ticks: { callback: v => v + '%' } } },
      }),
    };
  },

  /* 시간대별 피해·교전 추이 */
  timeline(R, s) {
    const m = R.model;
    const labels = R.root.tsDmg.map((_, i) => U.timeStr(m.t0 + i * m.binMin * 60000));
    const line = (label, data, color, extra) => Object.assign({ label, data, borderColor: color, backgroundColor: color, borderWidth: 2.2 * s, pointRadius: 2.5 * s, tension: 0.3 }, extra || {});
    const ds = [line(`${R.sideName} 피해`, R.root.tsDmg, COL.blue)];
    if (R.opfor) ds.push(line('대항군 피해', R.opfor.tsDmg, COL.red));
    ds.push(line('교전 개시', R.root.tsEng, COL.muted, { borderDash: [5 * s, 4 * s], pointRadius: 0 }));
    return {
      type: 'line', data: { labels, datasets: ds },
      options: Charts.opts(s, { interaction: { mode: 'index', intersect: false }, plugins: { tooltip: { mode: 'index', intersect: false } } }),
    };
  },

  /* 제대별 종합지수 (동일 제대 평균=100) */
  unitScore(R, s) {
    const us = R.units.filter(u => u.score != null);
    return {
      type: 'bar',
      data: { labels: us.map(u => u.name), datasets: [Charts.bar('종합지수', us.map(u => +u.score.toFixed(0)), us.map(u => JUDGE_COL(u.judge)))] },
      options: Charts.opts(s, {
        indexAxis: 'y',
        plugins: { legend: { display: false }, valueLabels: { enabled: true, size: 12 * s }, tooltip: { callbacks: { label: c => `종합지수 ${c.raw} · ${us[c.dataIndex].judge}` } } },
        scales: {
          x: { grid: { color: COL.grid }, min: 0, max: Math.max(130, Math.ceil((Math.max(0, ...us.map(u => u.score)) + 20) / 10) * 10), ticks: { font: { size: 11 * s } } },
          y: { grid: { display: false }, ticks: { font: { size: 12 * s }, color: COL.ink } },
        },
      }),
    };
  },

  /* 제대별 피해 현황 */
  casualty(R, s) {
    const us = R.units.map(u => u.node);
    const alive = c => Math.max(0, c.personnel - c.cas.사망 - c.cas.중상 - c.cas.경상);
    return {
      type: 'bar',
      data: {
        labels: us.map(Insight.unitName), datasets: [['사망', c => c.cas.사망], ['중상', c => c.cas.중상], ['경상', c => c.cas.경상], ['무피해', alive]]
          .map(([k, fn]) => Object.assign(Charts.bar(k, us.map(fn), COL[k]), { borderRadius: 0, borderColor: '#fff', borderWidth: { top: 1.5 }, borderSkipped: false })),
      },
      options: Charts.opts(s, {
        scales: { x: { stacked: true }, y: { stacked: true } },
        plugins: { tooltip: { mode: 'index', callbacks: { footer: it => `인원 ${us[it[0].dataIndex].personnel}명 · 생존율 ${U.fmt(us[it[0].dataIndex].metrics.prot_survival, 'pct')}` } } },
      }),
    };
  },

  /* 화기별 명중률·사격 수 */
  weapons(R, s) {
    const w = R.root.weapons;
    const names = Object.keys(w).sort((a, b) => w[b].shots - w[a].shots).slice(0, 10);
    return {
      type: 'bar',
      data: {
        labels: names, datasets: [
          Charts.bar('명중률(%)', names.map(k => pct1(w[k].shots ? w[k].hits / w[k].shots : null)), COL.blue, { yAxisID: 'y' }),
          { type: 'line', label: '사격 수', data: names.map(k => w[k].shots), borderColor: COL.series[3], backgroundColor: COL.series[3], borderWidth: 2 * s, pointRadius: 3.5 * s, yAxisID: 'y2' },
        ],
      },
      options: Charts.opts(s, {
        plugins: { valueLabels: { enabled: true, size: 11 * s, fmt: v => v + '%' } },
        scales: {
          y: { ticks: { callback: v => v + '%' } },
          y2: { position: 'right', beginAtZero: true, grid: { display: false }, border: { display: false }, ticks: { font: { size: 11 * s }, color: COL.muted } },
        },
      }),
    };
  },

  /* 훈련 간 추이 */
  trend(R, s) {
    const ids = ['fire_hit', 'prot_survival', 'int_predetect'];
    return {
      type: 'line',
      data: {
        labels: R.trend.map(t => t.name), datasets: ids.map((id, i) => ({
          label: Config.ind(id).name, data: R.trend.map(t => pct1(t.m[id])), borderColor: COL.series[i], backgroundColor: COL.series[i],
          borderWidth: 2.5 * s, pointRadius: R.trend.map(t => (t.cur ? 6 : 4) * s), spanGaps: true,
        })),
      },
      options: Charts.opts(s, { scales: { y: { beginAtZero: false, ticks: { callback: v => v + '%' } } } }),
    };
  },
};

/* config → 지정 크기 canvas (애니메이션 없이 즉시 그림) */
Charts.renderToCanvas = function (config, w, h) {
  const holder = document.createElement('div');
  holder.style.cssText = `position:fixed;left:-99999px;top:0;width:${w}px;height:${h}px;`;
  const cv = document.createElement('canvas');
  cv.width = w; cv.height = h;
  holder.appendChild(cv);
  document.body.appendChild(holder);
  config.options = config.options || {};
  config.options.responsive = false;
  config.options.animation = false;
  config.options.devicePixelRatio = 1;
  config.options.plugins = config.options.plugins || {};
  config.options.plugins.bgFill = { color: '#ffffff' };
  const ch = new Chart(cv, config);
  const out = document.createElement('canvas');
  out.width = w; out.height = h;
  out.getContext('2d').drawImage(cv, 0, 0);
  ch.destroy();
  holder.remove();
  return out;
};

/* 보고서용 PNG 바이트 */
Charts.png = async function (config, w, h) {
  const cv = Charts.renderToCanvas(config, w, h);
  const blob = await new Promise(res => cv.toBlob(res, 'image/png'));
  return { data: new Uint8Array(await blob.arrayBuffer()), w, h };
};
