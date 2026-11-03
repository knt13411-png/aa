'use strict';
/* 공통 유틸: 시각 변환, 숫자 형식, HTML 이스케이프, 다운로드, ZIP(무압축) 읽기/쓰기, IndexedDB */

const U = {};

/* 시각은 모두 "현지 벽시계 시각을 UTC로 표현한 ms"로 다룬다. (시간대 변환 오차 방지) */
U.excelSerialToMs = s => Math.round((s - 25569) * 86400000);
U.msToExcelSerial = ms => ms / 86400000 + 25569;

U.parseTime = function (v, baseDayMs) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') {
    if (!isFinite(v)) return NaN;
    if (v < 1 && baseDayMs != null) return baseDayMs + Math.round(v * 86400000); // 시간만 있는 셀
    return U.excelSerialToMs(v);
  }
  if (v instanceof Date) return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds());
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
  m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})\.?$/);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m && baseDayMs != null) return baseDayMs + ((+m[1]) * 3600 + (+m[2]) * 60 + (+(m[3] || 0))) * 1000;
  return NaN;
};

U.parseNum = function (v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : NaN;
  const n = Number(String(v).replace(/,/g, '').trim());
  return isFinite(n) ? n : NaN;
};

U.dateStr = ms => {
  if (ms == null || isNaN(ms)) return '';
  const d = new Date(ms);
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
};
U.timeStr = ms => {
  const d = new Date(ms);
  return String(d.getUTCHours()).padStart(2, '0') + ':' + String(d.getUTCMinutes()).padStart(2, '0');
};
U.nowStr = () => {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') + ' ' +
    String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
};
/* yyyy-mm-dd 입력값 → ms(UTC 표현) */
U.dayMs = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : null; };

U.fmt = function (v, fmt) {
  if (v === null || v === undefined || (typeof v === 'number' && !isFinite(v))) return '–';
  if (fmt === 'pct') return (v * 100).toFixed(1) + '%';
  if (fmt === 'num3') return v.toFixed(3);
  if (fmt === 'num2') return v.toFixed(2);
  if (fmt === 'int') return Math.round(v).toLocaleString('ko-KR');
  return v.toFixed(1);
};
U.fmtInd = (ind, v) => {
  const s = U.fmt(v, ind.fmt);
  return s === '–' || ind.fmt === 'pct' || !ind.unit ? s : s + ind.unit;
};

U.esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
U.safeName = s => String(s || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim() || '이름없음';
U.uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
U.mean = arr => { const a = arr.filter(x => x != null && isFinite(x)); return a.length ? a.reduce((s, x) => s + x, 0) / a.length : null; };
U.norm = s => String(s == null ? '' : s).replace(/[\s_]/g, '').toLowerCase();

U.download = function (blobOrUrl, filename) {
  const a = document.createElement('a');
  const url = typeof blobOrUrl === 'string' ? blobOrUrl : URL.createObjectURL(blobOrUrl);
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  if (typeof blobOrUrl !== 'string') setTimeout(() => URL.revokeObjectURL(url), 2000);
};

U.toast = function (msg, kind) {
  const el = document.createElement('div');
  el.className = 'toast ' + (kind || '');
  el.textContent = msg;
  document.getElementById('toasts').appendChild(el);
  setTimeout(() => el.classList.add('out'), 3200);
  setTimeout(() => el.remove(), 3700);
};

U.busy = function (msg) {
  const el = document.getElementById('busy');
  if (msg) { el.querySelector('span').textContent = msg; el.hidden = false; } else el.hidden = true;
};
U.nextFrame = () => new Promise(r => setTimeout(r, 30));

/* ---------- ZIP (무압축 저장 방식, UTF-8 파일명) ---------- */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
U.crc32 = function (buf) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
};
U.zip = function (files) { // files: [{path, data: Uint8Array}]
  const enc = new TextEncoder();
  const parts = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.path);
    const data = f.data;
    const crc = U.crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, 0, true); lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true);
    lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true);
    parts.push(new Uint8Array(lh.buffer), name, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true);
    ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const cdSize = central.reduce((s, p) => s + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: 'application/zip' });
};
U.unzip = function (buf) { // 이 도구가 만든 무압축 ZIP만 읽는다
  const u8 = new Uint8Array(buf), dv = new DataView(buf);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('ZIP 형식이 아닙니다.');
  const count = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const dec = new TextDecoder();
  const out = [];
  for (let i = 0; i < count; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('ZIP 목록이 손상되었습니다.');
    const method = dv.getUint16(p + 10, true);
    const size = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const lho = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
    if (method !== 0) throw new Error('압축된 ZIP은 지원하지 않습니다. 이 도구에서 내보낸 파일을 사용하세요.');
    const ds = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
    if (!name.endsWith('/')) out.push({ path: name, data: u8.slice(ds, ds + size) });
    p += 46 + nlen + xlen + clen;
  }
  return out;
};

/* ---------- IndexedDB ---------- */
U.idb = function () {
  if (U._db) return U._db;
  U._db = new Promise((res, rej) => {
    const rq = indexedDB.open('AAR_분석기', 1);
    rq.onupgradeneeded = () => { rq.result.createObjectStore('files'); rq.result.createObjectStore('meta'); };
    rq.onsuccess = () => res(rq.result);
    rq.onerror = () => rej(rq.error);
  });
  return U._db;
};
U.idbReq = async function (store, mode, fn) {
  const db = await U.idb();
  return new Promise((res, rej) => {
    const tx = db.transaction(store, mode);
    const rq = fn(tx.objectStore(store));
    tx.oncomplete = () => res(rq ? rq.result : undefined);
    tx.onerror = () => rej(tx.error);
  });
};
U.idbGet = (store, key) => U.idbReq(store, 'readonly', s => s.get(key));
U.idbPut = (store, key, val) => U.idbReq(store, 'readwrite', s => s.put(val, key));
U.idbDel = (store, key) => U.idbReq(store, 'readwrite', s => s.delete(key));
U.idbKeys = store => U.idbReq(store, 'readonly', s => s.getAllKeys());
