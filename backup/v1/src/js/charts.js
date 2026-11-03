'use strict';
/* 차트 공통 설정, 값 라벨 플러그인, PNG(1920×1080) 렌더링 */

const COL = {
  series: ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4'],
  사망: '#b8312f', 중상: '#ec835a', 경상: '#f2b632', 무피해: '#cfcec6',
  ink: '#0b0b0b', ink2: '#52514e', muted: '#898781', grid: '#e1e0d9', axis: '#c3c2b7', surface: '#fcfcfb',
  good: '#0a7a0a', bad: '#c03030', mid: '#6b6a65',
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
      ctx.save();
      ctx.font = `600 ${opts.size || 12}px ${FONT}`;
      ctx.fillStyle = COL.ink;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      chart.data.datasets.forEach((ds, i) => {
        const meta = chart.getDatasetMeta(i);
        if (meta.hidden || meta.type !== 'bar') return;
        meta.data.forEach((el, j) => {
          const v = ds.data[j];
          if (v == null) return;
          ctx.fillText(opts.fmt ? opts.fmt(v) : String(v), el.x, el.y - 3);
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

/* 막대/선 공통 옵션. s = 글자 배율(화면 1, 회의 1.7, PNG 2) */
Charts.opts = function (s, extra) {
  const f = Math.round(12 * s);
  return merge({
    responsive: true, maintainAspectRatio: false, animation: { duration: 200 },
    layout: { padding: { top: Math.round(14 * s), right: 8 } },
    datasets: { bar: { maxBarThickness: Math.round(64 * s) } },
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
  borderRadius: { topLeft: 4, topRight: 4 }, borderSkipped: 'bottom', categoryPercentage: 0.7, barPercentage: 0.9,
}, extra || {});

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

/* ---------- 캔버스 글자 도우미 (PNG 저장용) ---------- */
const Draw = {};
Draw.text = function (ctx, t, x, y, o) {
  o = o || {};
  ctx.font = `${o.weight || 400} ${o.size || 16}px ${FONT}`;
  ctx.fillStyle = o.color || COL.ink;
  ctx.textAlign = o.align || 'left';
  ctx.textBaseline = o.base || 'top';
  let s = String(t);
  if (o.maxW) while (s.length > 1 && ctx.measureText(s).width > o.maxW) s = s.slice(0, -2) + '…';
  ctx.fillText(s, x, y);
  return ctx.measureText(s).width;
};
/* 줄바꿈 출력, 마지막 y 반환 */
Draw.wrap = function (ctx, t, x, y, maxW, o) {
  o = o || {};
  const size = o.size || 16, lh = o.lh || Math.round(size * 1.45), maxLines = o.maxLines || 99;
  ctx.font = `${o.weight || 400} ${size}px ${FONT}`;
  ctx.fillStyle = o.color || COL.ink;
  ctx.textAlign = 'left'; ctx.textBaseline = 'top';
  let line = '', n = 0;
  const chars = Array.from(String(t));
  for (let i = 0; i < chars.length; i++) {
    const test = line + chars[i];
    if (chars[i] === '\n' || (ctx.measureText(test).width > maxW && line)) {
      if (n === maxLines - 1) { ctx.fillText(line.slice(0, -1) + '…', x, y); return y + lh; }
      ctx.fillText(line, x, y); y += lh; n++;
      line = chars[i] === '\n' ? '' : chars[i];
    } else line = test;
  }
  if (line) { ctx.fillText(line, x, y); y += lh; }
  return y;
};
Draw.badge = function (ctx, judge, x, y, size) {
  const map = { 강점: ['#e3f3e3', COL.good, '▲ 강점'], 취약: ['#fbe5e3', COL.bad, '▼ 취약'], 보통: ['#ecebe6', COL.mid, '● 보통'], 참고: ['#ecebe6', COL.mid, '참고'] };
  const [bg, fg, label] = map[judge] || ['#f0efea', COL.muted, judge || '–'];
  ctx.font = `700 ${size}px ${FONT}`;
  const w = ctx.measureText(label).width + size;
  ctx.fillStyle = bg;
  Draw.round(ctx, x, y, w, size * 1.6, size * 0.8); ctx.fill();
  ctx.fillStyle = fg; ctx.textBaseline = 'middle'; ctx.textAlign = 'left';
  ctx.fillText(label, x + size / 2, y + size * 0.8);
  return w;
};
Draw.round = function (ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath();
};
