-- setup_db.sql — Schema inicial para el Intelligence Bridge (Capa 1)
-- Ejecutar como: psql -U postgres -d rpf_intelligence -f setup_db.sql
-- (o reemplaza rpf_intelligence con el nombre de tu base de datos)

-- ── Tabla de log de archivos recibidos ────────────────────────────────────────
CREATE TABLE IF NOT EXISTS rpf_file_log (
    id           SERIAL PRIMARY KEY,
    filepath     TEXT NOT NULL,
    -- Nombre del archivo extraído automáticamente del path
    filename     TEXT GENERATED ALWAYS AS (
                     split_part(filepath, '/', -1)
                 ) STORED,
    received_at  TIMESTAMPTZ DEFAULT NOW(),
    processed    BOOLEAN DEFAULT FALSE,
    -- Clasificación automática por carpeta origen
    file_type    VARCHAR(20),   -- 'scada', 'evento', 'mapeo', 'resultado'
    size_bytes   BIGINT,
    -- Metadatos opcionales (rellenados por n8n en Capa 2)
    semestre     VARCHAR(20),   -- ej: '2024-1'
    evento       VARCHAR(100),  -- ej: 'Evento_001'
    fecha_evento DATE,          -- fecha del evento eléctrico
    notes        TEXT           -- observaciones o errores de procesamiento
);

CREATE INDEX IF NOT EXISTS idx_rpf_file_log_processed
    ON rpf_file_log(processed);
CREATE INDEX IF NOT EXISTS idx_rpf_file_log_received
    ON rpf_file_log(received_at DESC);
CREATE INDEX IF NOT EXISTS idx_rpf_file_log_type
    ON rpf_file_log(file_type);

-- ── Tabla de sync summary (registro de cada ejecución del daemon) ─────────────
CREATE TABLE IF NOT EXISTS rpf_sync_log (
    id               SERIAL PRIMARY KEY,
    sync_at          TIMESTAMPTZ DEFAULT NOW(),
    duration_secs    FLOAT,
    files_downloaded INT DEFAULT 0,
    files_skipped    INT DEFAULT 0,
    errors           INT DEFAULT 0,
    bytes_downloaded BIGINT DEFAULT 0,
    dry_run          BOOLEAN DEFAULT FALSE,
    stats_json       JSONB       -- detalle por carpeta
);

-- Vista rápida: archivos pendientes de procesar
CREATE OR REPLACE VIEW v_rpf_pending_files AS
SELECT id, filename, file_type, size_bytes, received_at, filepath
FROM rpf_file_log
WHERE processed = FALSE
ORDER BY received_at DESC;

-- Vista rápida: resumen de syncs recientes
CREATE OR REPLACE VIEW v_rpf_sync_history AS
SELECT
    sync_at,
    duration_secs,
    files_downloaded,
    files_skipped,
    errors,
    ROUND(bytes_downloaded / 1048576.0, 2) AS mb_downloaded
FROM rpf_sync_log
ORDER BY sync_at DESC
LIMIT 50;

-- ── Tabla de KPIs COBEE (Capa 2) ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS rpf_kpi_cobee (
    id               SERIAL PRIMARY KEY,
    semestre         VARCHAR(20),
    evento           VARCHAR(50),
    fecha_evento     DATE,
    unidad           VARCHAR(20),
    p_max_mw         FLOAT,
    p_0_mw           FLOAT,
    p_35_mw          FLOAT,
    r_inicial_mw     FLOAT,
    r_inicial_pct    FLOAT,
    p_entregada_mw   FLOAT,
    p_entregada_pct  FLOAT,
    aporta_rpf       VARCHAR(20),
    droop_inf_pct    FLOAT,
    droop_calc_pct   FLOAT,
    f_0_hz           FLOAT,
    f_min_hz         FLOAT,
    f_35_hz          FLOAT,
    t_0              TIME,
    t_min            TIME,
    t_35             TIME,
    source_file      TEXT,
    extracted_at     TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_kpi_cobee_evento   ON rpf_kpi_cobee(evento);
CREATE INDEX IF NOT EXISTS idx_kpi_cobee_semestre ON rpf_kpi_cobee(semestre);
CREATE INDEX IF NOT EXISTS idx_kpi_cobee_unidad   ON rpf_kpi_cobee(unidad);

-- Mensaje de confirmación
DO $$
BEGIN
    RAISE NOTICE 'Schema RPF Intelligence creado correctamente.';
    RAISE NOTICE 'Tablas: rpf_file_log, rpf_sync_log, rpf_kpi_cobee';
    RAISE NOTICE 'Vistas: v_rpf_pending_files, v_rpf_sync_history';
END $$;
