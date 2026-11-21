-- Data Book's shared persistence layer. Apply to the connected Supabase project.
create table public.app_data (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  module text not null check (length(module) between 1 and 80),
  record_key text not null check (length(record_key) between 1 and 256),
  title text not null default '',
  payload jsonb not null default '{}'::jsonb,
  file_path text unique,
  file_name text,
  file_size bigint check (file_size is null or file_size >= 0),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, module, record_key)
);

create index app_data_user_updated_idx on public.app_data (user_id, updated_at desc);

alter table public.app_data enable row level security;
revoke all on public.app_data from anon;
grant select, insert, update, delete on public.app_data to authenticated;

create policy app_data_select_own on public.app_data
  for select to authenticated
  using ((select auth.uid()) = user_id);

create policy app_data_insert_own on public.app_data
  for insert to authenticated
  with check ((select auth.uid()) = user_id);

create policy app_data_update_own on public.app_data
  for update to authenticated
  using ((select auth.uid()) = user_id)
  with check ((select auth.uid()) = user_id);

create policy app_data_delete_own on public.app_data
  for delete to authenticated
  using ((select auth.uid()) = user_id);

insert into storage.buckets (id, name, public)
values ('data-book-files', 'data-book-files', false);

create policy data_book_files_select_own on storage.objects
  for select to authenticated
  using (
    bucket_id = 'data-book-files'
    and (storage.foldername(name))[1] = (select auth.uid())::text
  );

create policy data_book_files_insert_own on storage.objects
  for insert to authenticated
  with check (
    bucket_id = 'data-book-files'
    and (storage.foldername(name))[1] = (select auth.uid())::text
  );

create policy data_book_files_update_own on storage.objects
  for update to authenticated
  using (
    bucket_id = 'data-book-files'
    and (storage.foldername(name))[1] = (select auth.uid())::text
  )
  with check (
    bucket_id = 'data-book-files'
    and (storage.foldername(name))[1] = (select auth.uid())::text
  );

create policy data_book_files_delete_own on storage.objects
  for delete to authenticated
  using (
    bucket_id = 'data-book-files'
    and (storage.foldername(name))[1] = (select auth.uid())::text
  );
