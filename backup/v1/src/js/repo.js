'use strict';
/* 훈련 단위 계층 저장: 훈련/여단/대대/중대/소대 폴더 구조 */

const Repo = {};

Repo.folderName = info => `${(info.시작일 || '').slice(0, 7) || '날짜미상'}_${U.safeName(info.훈련명)}`;

/* 제대 노드 → 훈련 폴더 기준 상대 경로. 같은 여단명이 두 진영에 모두 있으면 '(진영)'을 붙인다. */
Repo.nodePath = function (model, node) {
  if (node.level === 0) return '';
  const bdeSides = {};
  for (const n of Object.values(model.nodes)) if (n.level === 1) (bdeSides[n.name] || (bdeSides[n.name] = new Set())).add(n.side);
  const segs = node.path.map(U.safeName);
  if (bdeSides[node.path[0]] && bdeSides[node.path[0]].size > 1) segs[0] += `(${node.side})`;
  return segs.join('/');
};

Repo.hasOpinion = o => !!o && Object.values(o).some(v => v && String(v).trim());

/* state: {info, model, opinions, raw(ArrayBuffer|null)} → 저장 후 info 반환 */
Repo.save = async function (state, onProgress) {
  const info = state.info;
  if (!info.id) info.id = U.uid();
  let folder = info.폴더;
  if (!folder) {
    const base = Repo.folderName(info);
    folder = base;
    for (let i = 2; ; i++) {
      const other = await Store.readJSON(folder + '/_훈련정보.json');
      if (!other || other.id === info.id) break;
      folder = `${base}_${i}`;
    }
  }
  const model = state.model;
  const now = U.nowStr();
  info.폴더 = folder;
  info.생성일 = info.생성일 || now;
  info.수정일 = now;
  info.분석기준버전 = model.분석기준버전;
  info.levelStats = model.levelStats;
  info.tStart = model.tStart; info.tEnd = model.tEnd;
  info.소견제대 = Object.keys(state.opinions).filter(k => Repo.hasOpinion(state.opinions[k]));
  info.제대수 = Object.values(model.nodes).filter(n => n.level > 0).length;

  if (state.raw) await Store.write(folder + '/_원천/원천데이터.xlsx', new Uint8Array(state.raw));
  await Store.writeJSON(folder + '/_분석_전체.json', model);

  // 루트(진영 필터별) 소견은 훈련 폴더에 저장
  const rootOp = {};
  for (const s of SIDES) if (state.opinions['ROOT|' + s]) rootOp[s] = state.opinions['ROOT|' + s];
  await Store.writeJSON(folder + '/_소견.json', { 제대: '훈련 전체', 소견: rootOp, 수정일: now });

  const units = Object.values(model.nodes).filter(n => n.level > 0);
  const prevStats = state.prevStats || [];
  let i = 0;
  for (const n of units) {
    const p = folder + '/' + Repo.nodePath(model, n);
    const base = Engine.baseline(Engine.levelKey(n), prevStats);
    await Store.writeJSON(p + '/_분석.json', {
      분석기준버전: model.분석기준버전, 분석일시: model.분석일시,
      제대: n.name, 제대경로: n.path.join(' > '), 진영: n.side, 수준: LEVELS[n.level],
      지표: n.metrics,
      집계: { 사격수: n.shots, 명중수: n.hits, 교전수: n.engagements, 피해유발: n.dmgHits, 인원: n.personnel, 피해: n.cas },
      진단: Engine.diagnose(n.metrics, base, Config.cur).functions.map(f => ({ 기능: f.fn, 종합지수: f.index, 판정: f.judge, 진단문: f.summary })),
    });
    await Store.writeJSON(p + '/_소견.json', { 제대: n.path.join(' > '), 진영: n.side, 소견: state.opinions[n.id] || {}, 수정일: now });
    if (onProgress && ++i % 10 === 0) onProgress(i, units.length);
  }
  await Store.writeJSON(folder + '/_훈련정보.json', info);
  return info;
};

Repo.list = async function () {
  const out = [];
  for (const d of await Store.listDirs('')) {
    const info = await Store.readJSON(d + '/_훈련정보.json');
    if (info) { info.폴더 = d; out.push(info); }
  }
  return out.sort((a, b) => String(a.시작일).localeCompare(String(b.시작일)));
};

Repo.load = async function (folder) {
  const info = await Store.readJSON(folder + '/_훈련정보.json');
  if (!info) throw new Error('훈련 정보 파일(_훈련정보.json)이 없습니다.');
  info.폴더 = folder;
  const model = Engine.sortChildren(await Store.readJSON(folder + '/_분석_전체.json'));
  if (!model) throw new Error('분석 결과 파일(_분석_전체.json)이 없습니다. 재분석이 필요합니다.');
  const opinions = {};
  const rootOp = await Store.readJSON(folder + '/_소견.json');
  if (rootOp && rootOp.소견) for (const [s, o] of Object.entries(rootOp.소견)) opinions['ROOT|' + s] = o;
  const units = Object.values(model.nodes).filter(n => n.level > 0);
  await Promise.all(units.map(async n => {
    const o = await Store.readJSON(folder + '/' + Repo.nodePath(model, n) + '/_소견.json');
    if (o && Repo.hasOpinion(o.소견)) opinions[n.id] = o.소견;
  }));
  return { info, model, opinions };
};

Repo.loadModel = async folder => Engine.sortChildren(await Store.readJSON(folder + '/_분석_전체.json'));
Repo.loadRaw = folder => Store.readBin(folder + '/_원천/원천데이터.xlsx');
Repo.remove = folder => Store.removeDir(folder);

Repo.exportZip = async function (folder) {
  const files = [];
  for (const p of await Store.listFiles(folder)) {
    const buf = await Store.readBin(folder + '/' + p);
    files.push({ path: folder + '/' + p, data: new Uint8Array(buf) });
  }
  return U.zip(files);
};

/* 내보낸 ZIP 불러오기. 같은 id의 훈련이 있으면 덮어쓴다. */
Repo.importZip = async function (buf) {
  const files = U.unzip(buf);
  const infoFile = files.find(f => /(^|\/)_훈련정보\.json$/.test(f.path));
  if (!infoFile) throw new Error('훈련 내보내기 파일이 아닙니다 (_훈련정보.json 없음).');
  const prefix = infoFile.path.replace(/_훈련정보\.json$/, '');
  const info = JSON.parse(new TextDecoder().decode(infoFile.data));
  const existing = (await Repo.list()).find(t => t.id === info.id);
  let folder = existing ? existing.폴더 : prefix.replace(/\/$/, '') || Repo.folderName(info);
  if (!existing) {
    const base = folder;
    for (let i = 2; await Store.exists(folder + '/_훈련정보.json'); i++) folder = `${base}_${i}`;
  } else {
    await Repo.remove(folder);
  }
  for (const f of files) {
    if (!f.path.startsWith(prefix)) continue;
    let data = f.data;
    const rel = f.path.slice(prefix.length);
    if (rel === '_훈련정보.json') { info.폴더 = folder; data = JSON.stringify(info, null, 1); }
    else if (/\.json$/.test(rel)) data = new TextDecoder().decode(f.data);
    await Store.write(folder + '/' + rel, data);
  }
  return info;
};
