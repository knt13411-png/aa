-- 사후검토 분석기: 분석한 훈련 요약(비교 기준용)을 저장하는 테이블
-- Supabase 대시보드 → SQL Editor 에 붙여 넣고 Run 한 번 실행하면 된다.

create table if not exists public.aar_history (
  id          text primary key,          -- 훈련명|시작일
  name        text not null,             -- 훈련명
  start_date  text,                      -- 시작일 (YYYY-MM-DD)
  end_date    text,                      -- 종료일
  units       text,                      -- 훈련부대
  source_file text,                      -- 원천 엑셀 파일명
  level_stats jsonb not null,            -- 제대별 집계 지표 (분석 결과)
  analyzed_at text,                      -- 분석일 (앱 기준 시각)
  updated_at  timestamptz not null default now()
);

alter table public.aar_history enable row level security;

-- 사이트(공개 키)로 읽기/쓰기/삭제 허용. 사이트 주소를 아는 사람은 누구나 기록을 보고 고칠 수 있다.
drop policy if exists "aar read"   on public.aar_history;
drop policy if exists "aar insert" on public.aar_history;
drop policy if exists "aar update" on public.aar_history;
drop policy if exists "aar delete" on public.aar_history;
create policy "aar read"   on public.aar_history for select to anon, authenticated using (true);
create policy "aar insert" on public.aar_history for insert to anon, authenticated with check (true);
create policy "aar update" on public.aar_history for update to anon, authenticated using (true) with check (true);
create policy "aar delete" on public.aar_history for delete to anon, authenticated using (true);

grant select, insert, update, delete on public.aar_history to anon, authenticated;
