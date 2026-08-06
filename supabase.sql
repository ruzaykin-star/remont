-- Run this once in your Supabase project's SQL editor (Project -> SQL Editor -> New query).

create table if not exists remont_data (
  id text primary key,
  data jsonb not null,
  updated_by text,
  updated_at timestamptz not null default now()
);

alter table remont_data enable row level security;

-- The app has no login screen, so it connects with the public "anon" key.
-- These policies let anyone holding that key (i.e. anyone who opens the deployed
-- site) read and write the single shared row. That's fine for a private family
-- tool whose URL isn't shared publicly, but it is not real per-user security —
-- see the README for how to lock it down further if that ever matters.
create policy "anon can read remont_data"
  on remont_data for select
  using (true);

create policy "anon can insert remont_data"
  on remont_data for insert
  with check (true);

create policy "anon can update remont_data"
  on remont_data for update
  using (true)
  with check (true);

-- Enable realtime so all open browser tabs see each other's changes live.
alter publication supabase_realtime add table remont_data;
