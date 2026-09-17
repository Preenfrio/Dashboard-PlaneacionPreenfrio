-- ================================================================
-- Planeador de Preenfrio - Esquema inicial consolidado para Supabase
-- Excel = fuente documental; Supabase = fuente operativa persistente
-- ================================================================

create extension if not exists pgcrypto;

create type public.estado_sincronizacion as enum (
  'NUEVO', 'SIN_CAMBIOS', 'MODIFICADO', 'RETIRADO_ORIGEN',
  'REACTIVADO', 'AMBIGUO', 'CAMBIO_RETROACTIVO'
);

create type public.tipo_asignacion as enum (
  'AUTOMATICA', 'MANUAL', 'DIRECTO_SIN_PREENFRIO'
);

create type public.estado_importacion as enum (
  'INICIADA', 'VALIDADA', 'APLICADA', 'ERROR'
);

create table public.importaciones (
  id uuid primary key default gen_random_uuid(),
  nombre_archivo text not null,
  hash_archivo text not null,
  semanas integer[] not null default '{}',
  filas_leidas integer not null default 0,
  filas_validas integer not null default 0,
  filas_nuevas integer not null default 0,
  filas_modificadas integer not null default 0,
  filas_retiradas integer not null default 0,
  filas_ambiguas integer not null default 0,
  estado public.estado_importacion not null default 'INICIADA',
  detalle_error text,
  importado_por text,
  creado_en timestamptz not null default now(),
  aplicado_en timestamptz,
  unique (hash_archivo)
);

create table public.procesos (
  id uuid primary key default gen_random_uuid(),
  identidad_normalizada text not null unique,
  activo_fuente boolean not null default true,
  estado_sincronizacion public.estado_sincronizacion not null default 'NUEVO',
  primera_importacion_id uuid references public.importaciones(id),
  ultima_importacion_id uuid references public.importaciones(id),
  creado_en timestamptz not null default now(),
  actualizado_en timestamptz not null default now()
);

create table public.procesos_fuente (
  proceso_id uuid primary key references public.procesos(id) on delete cascade,
  llave_base_consulta text,
  ocurrencia_consulta integer,
  llave_embarque_consulta text,
  ov text,
  descripcion text,
  sku text not null,
  lote text,
  finca text not null,
  destino text not null,
  cliente text,
  cantidad numeric(14,3),
  pallets numeric(12,3) not null check (pallets > 0),
  fecha_proceso timestamp without time zone not null,
  semana integer not null,
  fecha_cita timestamp without time zone not null,
  dias_transito numeric(8,3),
  horas_transito numeric(10,3),
  nivel_critico integer,
  archivo_origen text not null,
  hash_contenido text not null,
  datos_extra jsonb not null default '{}'::jsonb,
  actualizado_en timestamptz not null default now()
);

create table public.procesos_planificacion (
  proceso_id uuid primary key references public.procesos(id) on delete cascade,
  preenfrio_sugerido text,
  preenfrio_manual text,
  orden_sugerido integer,
  orden_manual integer,
  tipo_asignacion public.tipo_asignacion not null default 'AUTOMATICA',
  bloqueado_manual boolean not null default false,
  observacion_manual text,
  fecha_ingreso timestamp without time zone,
  fecha_despacho timestamp without time zone,
  llegada_estimada timestamp without time zone,
  cumplimiento text check (cumplimiento in ('CUMPLE','NO CUMPLE') or cumplimiento is null),
  holgura_horas numeric(12,3),
  distancia_km numeric(12,3),
  modificado_por text,
  modificado_en timestamptz,
  actualizado_en timestamptz not null default now(),
  check (preenfrio_manual is null or preenfrio_manual in ('NELLY','FORTALEZA','DIRECTO SIN PREENFRIO')),
  check (preenfrio_sugerido is null or preenfrio_sugerido in ('NELLY','FORTALEZA'))
);

create table public.procesos_versiones (
  id bigint generated always as identity primary key,
  proceso_id uuid not null references public.procesos(id) on delete cascade,
  importacion_id uuid references public.importaciones(id),
  version_numero integer not null,
  tipo_cambio public.estado_sincronizacion not null,
  datos_anteriores jsonb,
  datos_nuevos jsonb not null,
  creado_en timestamptz not null default now(),
  unique (proceso_id, version_numero)
);

create table public.conciliaciones_pendientes (
  id uuid primary key default gen_random_uuid(),
  importacion_id uuid not null references public.importaciones(id) on delete cascade,
  proceso_candidato_id uuid references public.procesos(id),
  registro_nuevo jsonb not null,
  puntaje numeric(8,3),
  estado text not null default 'PENDIENTE' check (estado in ('PENDIENTE','RELACIONADO','NUEVO','IGNORADO')),
  resuelto_por text,
  resuelto_en timestamptz,
  creado_en timestamptz not null default now()
);

create table public.ciclos_preenfrio (
  id uuid primary key default gen_random_uuid(),
  preenfrio text not null check (preenfrio in ('NELLY','FORTALEZA')),
  numero_ciclo integer not null,
  inicio timestamp without time zone not null,
  fin timestamp without time zone not null,
  capacidad numeric(8,3) not null,
  tarimas_usadas numeric(8,3) not null,
  tarimas_libres numeric(8,3) not null,
  actualizado_en timestamptz not null default now(),
  check (fin > inicio),
  unique (preenfrio, inicio, numero_ciclo)
);

create table public.ciclos_procesos (
  ciclo_id uuid not null references public.ciclos_preenfrio(id) on delete cascade,
  proceso_id uuid not null references public.procesos(id) on delete cascade,
  tarimas numeric(8,3) not null check (tarimas > 0),
  primary key (ciclo_id, proceso_id)
);

create table public.historial_cambios (
  id bigint generated always as identity primary key,
  proceso_id uuid not null references public.procesos(id) on delete cascade,
  campo text not null,
  valor_anterior jsonb,
  valor_nuevo jsonb,
  origen text not null check (origen in ('IMPORTACION','DASHBOARD','SISTEMA')),
  usuario text,
  creado_en timestamptz not null default now()
);

create table public.catalogo_destinos (
  destino text primary key,
  cliente text,
  horas_transito numeric(10,3),
  ciclo_horas_nelly numeric(10,3),
  ciclo_horas_fortaleza numeric(10,3),
  criticidad integer,
  actualizado_en timestamptz not null default now()
);

create table public.ubicaciones (
  nombre text primary key,
  tipo text not null check (tipo in ('FINCA','PREENFRIO')),
  latitud numeric(11,8) not null,
  longitud numeric(11,8) not null,
  actualizado_en timestamptz not null default now()
);

create index idx_fuente_semana_fecha on public.procesos_fuente(semana, fecha_proceso);
create index idx_fuente_finca_destino on public.procesos_fuente(finca, destino);
create index idx_fuente_sku_lote on public.procesos_fuente(sku, lote);
create index idx_fuente_ov on public.procesos_fuente(ov) where ov is not null;
create index idx_fuente_llave_embarque on public.procesos_fuente(llave_embarque_consulta) where llave_embarque_consulta is not null;
create index idx_procesos_activos on public.procesos(activo_fuente, estado_sincronizacion);
create index idx_ciclos_inicio on public.ciclos_preenfrio(preenfrio, inicio);
create index idx_historial_proceso_fecha on public.historial_cambios(proceso_id, creado_en desc);

create or replace function public.touch_actualizado_en()
returns trigger language plpgsql as $$
begin
  new.actualizado_en = now();
  return new;
end;
$$;

create trigger trg_procesos_touch before update on public.procesos
for each row execute function public.touch_actualizado_en();
create trigger trg_fuente_touch before update on public.procesos_fuente
for each row execute function public.touch_actualizado_en();
create trigger trg_plan_touch before update on public.procesos_planificacion
for each row execute function public.touch_actualizado_en();

create or replace view public.v_planeacion_efectiva as
select
  p.id as proceso_id,
  p.identidad_normalizada,
  p.activo_fuente,
  p.estado_sincronizacion,
  f.ov, f.descripcion, f.sku, f.lote, f.finca, f.destino, f.cliente,
  f.cantidad, f.pallets, f.fecha_proceso, f.semana, f.fecha_cita,
  f.horas_transito, f.nivel_critico,
  pl.preenfrio_sugerido,
  pl.preenfrio_manual,
  case
    when pl.tipo_asignacion = 'DIRECTO_SIN_PREENFRIO' then 'DIRECTO SIN PREENFRIO'
    else coalesce(pl.preenfrio_manual, pl.preenfrio_sugerido)
  end as preenfrio_efectivo,
  pl.orden_sugerido,
  pl.orden_manual,
  coalesce(pl.orden_manual, pl.orden_sugerido) as orden_efectivo,
  pl.tipo_asignacion,
  pl.bloqueado_manual,
  pl.fecha_ingreso, pl.fecha_despacho, pl.llegada_estimada,
  pl.cumplimiento, pl.holgura_horas, pl.distancia_km,
  pl.modificado_por, pl.modificado_en
from public.procesos p
join public.procesos_fuente f on f.proceso_id = p.id
left join public.procesos_planificacion pl on pl.proceso_id = p.id;

-- Seguridad: no exponer escritura directa desde el navegador.
alter table public.importaciones enable row level security;
alter table public.procesos enable row level security;
alter table public.procesos_fuente enable row level security;
alter table public.procesos_planificacion enable row level security;
alter table public.procesos_versiones enable row level security;
alter table public.conciliaciones_pendientes enable row level security;
alter table public.ciclos_preenfrio enable row level security;
alter table public.ciclos_procesos enable row level security;
alter table public.historial_cambios enable row level security;
alter table public.catalogo_destinos enable row level security;
alter table public.ubicaciones enable row level security;

revoke all on all tables in schema public from anon, authenticated;
grant select on public.v_planeacion_efectiva to authenticated;
