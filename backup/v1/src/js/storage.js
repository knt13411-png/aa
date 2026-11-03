'use strict';
/* 저장소: A(폴더 직접 저장, File System Access API) / B(브라우저 IndexedDB + ZIP 내보내기)
   두 방식 모두 "경로 문자열(/ 구분)" 단위로 읽고 쓴다. */

const Store = {
  mode: 'B',
  root: null,          // A 방식의 폴더 핸들
  needsReconnect: false,
  supportsA: typeof window.showDirectoryPicker === 'function',

  async init(pref) {
    this.root = null; this.needsReconnect = false;
    const wantA = pref === 'A' || (pref !== 'B' && this.supportsA);
    this.mode = wantA && this.supportsA ? 'A' : 'B';
    if (this.mode === 'A') {
      let h = null;
      try { h = await U.idbGet('meta', 'rootHandle'); } catch (e) { h = null; }
      if (h) {
        let p = 'prompt';
        try { p = await h.queryPermission({ mode: 'readwrite' }); } catch (e) { p = 'prompt'; }
        if (p === 'granted') this.root = h; else { this.root = null; this.pending = h; this.needsReconnect = true; }
      }
    }
  },
  get ready() { return this.mode === 'B' || !!this.root; },
  get label() {
    if (this.mode === 'B') return '저장 방식 B · 브라우저 저장';
    return this.root ? `저장 방식 A · 폴더: ${this.root.name}` : '저장 방식 A · 폴더 미지정';
  },
  async pickFolder() {
    const h = await window.showDirectoryPicker({ id: 'aar-data', mode: 'readwrite' });
    await this.useHandle(h);
  },
  async useHandle(h) {
    this.root = h; this.mode = 'A'; this.needsReconnect = false; this.pending = null;
    try { await U.idbPut('meta', 'rootHandle', h); } catch (e) { /* 핸들 저장 불가 시 매번 지정 */ }
  },
  async reconnect() {
    const h = this.pending;
    if (!h) return this.pickFolder();
    const p = await h.requestPermission({ mode: 'readwrite' });
    if (p === 'granted') { this.root = h; this.needsReconnect = false; this.pending = null; }
    else throw new Error('폴더 접근 권한이 거부되었습니다.');
  },
  check() { if (!this.ready) throw new Error('저장 폴더가 지정되지 않았습니다. [설정]에서 저장 폴더를 지정하세요.'); },

  /* ---- A 방식 내부 ---- */
  async _dir(segs, create) {
    let d = this.root;
    for (const s of segs) d = await d.getDirectoryHandle(s, { create });
    return d;
  },
  _split(path) { const segs = path.split('/').filter(Boolean); return { dir: segs.slice(0, -1), name: segs[segs.length - 1] }; },

  async write(path, data) { // data: string | Uint8Array | ArrayBuffer
    this.check();
    if (this.mode === 'A') {
      const { dir, name } = this._split(path);
      const d = await this._dir(dir, true);
      const fh = await d.getFileHandle(name, { create: true });
      const w = await fh.createWritable();
      await w.write(data);
      await w.close();
    } else {
      const val = typeof data === 'string' ? { t: 'text', d: data } : { t: 'bin', d: new Uint8Array(data).slice() };
      await U.idbPut('files', path, Object.assign(val, { m: Date.now() }));
    }
  },
  writeJSON(path, obj) { return this.write(path, JSON.stringify(obj, null, 1)); },
  async readRaw(path) {
    this.check();
    if (this.mode === 'A') {
      try {
        const { dir, name } = this._split(path);
        const d = await this._dir(dir, false);
        const f = await (await d.getFileHandle(name)).getFile();
        return f;
      } catch (e) { return null; }
    }
    const v = await U.idbGet('files', path);
    if (!v) return null;
    return v.t === 'text' ? new Blob([v.d], { type: 'application/json' }) : new Blob([v.d]);
  },
  async readText(path) { const f = await this.readRaw(path); return f ? f.text() : null; },
  async readJSON(path) { const t = await this.readText(path); if (t == null) return null; try { return JSON.parse(t); } catch (e) { return null; } },
  async readBin(path) { const f = await this.readRaw(path); return f ? f.arrayBuffer() : null; },

  async listDirs(path) {
    this.check();
    const out = [];
    if (this.mode === 'A') {
      let d;
      try { d = await this._dir(path.split('/').filter(Boolean), false); } catch (e) { return out; }
      for await (const [name, h] of d.entries()) if (h.kind === 'directory') out.push(name);
    } else {
      const pre = path ? path.replace(/\/$/, '') + '/' : '';
      const set = new Set();
      for (const k of await U.idbKeys('files')) {
        if (!k.startsWith(pre)) continue;
        const rest = k.slice(pre.length);
        const i = rest.indexOf('/');
        if (i > 0) set.add(rest.slice(0, i));
      }
      out.push(...set);
    }
    return out.sort();
  },
  /* path 아래 모든 파일 경로(상대) */
  async listFiles(path) {
    this.check();
    const out = [];
    if (this.mode === 'A') {
      const walk = async (d, pre) => {
        for await (const [name, h] of d.entries()) {
          if (h.kind === 'directory') await walk(h, pre + name + '/');
          else out.push(pre + name);
        }
      };
      try { await walk(await this._dir(path.split('/').filter(Boolean), false), ''); } catch (e) { /* 없음 */ }
    } else {
      const pre = path.replace(/\/$/, '') + '/';
      for (const k of await U.idbKeys('files')) if (k.startsWith(pre)) out.push(k.slice(pre.length));
    }
    return out.sort();
  },
  async exists(path) { return (await this.readRaw(path)) !== null; },
  async removeDir(path) {
    this.check();
    if (this.mode === 'A') {
      const { dir, name } = this._split(path);
      const d = await this._dir(dir, false);
      await d.removeEntry(name, { recursive: true });
    } else {
      const pre = path.replace(/\/$/, '') + '/';
      for (const k of await U.idbKeys('files')) if (k.startsWith(pre)) await U.idbDel('files', k);
    }
  },
};
