-- Ejecutá esto en Supabase: Project → SQL Editor → New query → pegar y "Run"
-- Podés correrlo las veces que quieras sin que rompa nada (es "idempotente").

create table if not exists app_data (
  id text primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

alter table app_data enable row level security;

-- Borramos cualquier policy previa (de versiones anteriores, con o sin login)
drop policy if exists "Lectura pública" on app_data;
drop policy if exists "Escritura pública" on app_data;
drop policy if exists "Actualización pública" on app_data;
drop policy if exists "Lectura para usuarios logueados" on app_data;
drop policy if exists "Escritura para usuarios logueados" on app_data;
drop policy if exists "Actualización para usuarios logueados" on app_data;

-- Acceso público: cualquiera con la URL del sitio puede leer/escribir (sin login).
create policy "Lectura pública" on app_data
  for select using (true);

create policy "Escritura pública" on app_data
  for insert with check (true);

create policy "Actualización pública" on app_data
  for update using (true);

-- Habilitar Realtime para esta tabla (para que sincronice en vivo entre pestañas/dispositivos).
-- Este chequeo evita el error "already member of publication" si ya estaba habilitado.
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and tablename = 'app_data'
  ) then
    alter publication supabase_realtime add table app_data;
  end if;
end $$;
