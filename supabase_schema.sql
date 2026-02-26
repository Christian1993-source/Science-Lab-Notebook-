create table if not exists public.lab_reports (
  id text primary key,
  teacher_email text not null,
  title text not null,
  student_name text not null,
  experiment_date text not null,
  status text not null check (status in ('Draft', 'Submitted')),
  payload jsonb not null,
  updated_at timestamptz not null default now(),
  submitted_at timestamptz
);

create index if not exists idx_lab_reports_status on public.lab_reports(status);
create index if not exists idx_lab_reports_updated_at on public.lab_reports(updated_at desc);
