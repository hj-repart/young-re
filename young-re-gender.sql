-- 근로자 성별 칸 추가 (남 M / 여 F). 앱 파일을 올리기 전에 먼저 실행.
alter table public.workers
  add column if not exists gender text check (gender in ('M','F'));
