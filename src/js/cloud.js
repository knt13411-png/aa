'use strict';
/* Supabase 공유 저장소: 분석한 훈련 요약을 aar_history 테이블에 저장/불러오기 (REST, 라이브러리 없이 fetch).
   주소나 키가 비어 있으면 꺼진 상태로 두고 브라우저 저장만 쓴다. 테이블은 supabase/schema.sql 로 만든다. */

const Cloud = {
  url: 'https://wpqwigdmbwnirdxqsnut.supabase.co',   // 예: https://xxxx.supabase.co
  key: 'sb_publishable_HhBvpmItqsYAJ6ReKSYnhQ_Y0GOrm5u',   // Project API keys 의 anon / publishable 키 (공개용)
  state: 'off', // off | ok | err
};

Cloud.on = () => !!(Cloud.url && Cloud.key);

Cloud.req = async function (path, opt) {
  const o = opt || {};
  const res = await fetch(`${Cloud.url.replace(/\/$/, '')}/rest/v1/${path}`, {
    method: o.method || 'GET',
    headers: Object.assign({ apikey: Cloud.key, Authorization: `Bearer ${Cloud.key}`, 'Content-Type': 'application/json' }, o.headers || {}),
    body: o.body ? JSON.stringify(o.body) : undefined,
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${text.slice(0, 200)}`);
  return text ? JSON.parse(text) : null; // return=minimal 이면 본문이 비어 있다
};

const toRow = h => ({ id: h.id, name: h.훈련명, start_date: h.시작일, end_date: h.종료일, units: h.훈련부대 || null, source_file: h.원천파일명 || null, level_stats: h.levelStats, analyzed_at: h.분석일 });
const fromRow = r => ({ id: r.id, 훈련명: r.name, 시작일: r.start_date, 종료일: r.end_date, 훈련부대: r.units, 원천파일명: r.source_file, levelStats: r.level_stats, 분석일: r.analyzed_at });

Cloud.list = async () => (await Cloud.req('aar_history?select=*&order=start_date.asc')).map(fromRow);
Cloud.save = async list => {
  if (list.length) await Cloud.req('aar_history?on_conflict=id', { method: 'POST', body: list.map(toRow), headers: { Prefer: 'resolution=merge-duplicates,return=minimal' } });
};
Cloud.clear = () => Cloud.req('aar_history?id=not.is.null', { method: 'DELETE', headers: { Prefer: 'return=minimal' } });
