-- ================================================================
-- SCRIPT DE CONFIGURACIÓN Y PERMISOS SUPABASE PARA CLASESINC AI
-- Ejecutar en el SQL Editor de Supabase (https://app.supabase.com)
-- ================================================================

-- 1. Habilitar la extensión de vectores (pgvector)
CREATE EXTENSION IF NOT EXISTS vector;

-- 2. Eliminar restricción de clave foránea estricta de usuarios si existe
ALTER TABLE public.documents DROP CONSTRAINT IF EXISTS documents_user_id_fkey;

-- 3. Crear la tabla de documentos si no existe
CREATE TABLE IF NOT EXISTS public.documents (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id TEXT NOT NULL DEFAULT '00000000-0000-0000-0000-000000000000',
    title TEXT NOT NULL,
    file_type TEXT NOT NULL,
    file_url TEXT,
    summary_markdown TEXT,
    status TEXT NOT NULL DEFAULT 'processing',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 4. Crear la tabla de fragmentos (chunks) con vector de dimensión 768
CREATE TABLE IF NOT EXISTS public.document_chunks (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    document_id UUID NOT NULL REFERENCES public.documents(id) ON DELETE CASCADE,
    content TEXT NOT NULL,
    embedding vector(768),
    page_number INT DEFAULT 1,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 5. Crear la función RPC para búsqueda semántica vectorial
CREATE OR REPLACE FUNCTION match_document_chunks(
    query_embedding vector(768),
    match_count INT DEFAULT 5,
    filter_document_id UUID DEFAULT NULL
)
RETURNS TABLE (
    id UUID,
    document_id UUID,
    content TEXT,
    page_number INT,
    similarity FLOAT
)
LANGUAGE plpgsql
AS $$
BEGIN
    RETURN QUERY
    SELECT
        dc.id,
        dc.document_id,
        dc.content,
        dc.page_number,
        1 - (dc.embedding <=> query_embedding) AS similarity
    FROM public.document_chunks dc
    WHERE (filter_document_id IS NULL OR dc.document_id = filter_document_id)
    ORDER BY dc.embedding <=> query_embedding
    LIMIT match_count;
END;
$$;

-- 6. Deshabilitar RLS en todas las tablas para permitir lectura/escritura abierta mediante Anon Key
ALTER TABLE public.documents DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.document_chunks DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.profiles DISABLE ROW LEVEL SECURITY;

-- 7. Configurar Bucket de Storage 'class-files'
INSERT INTO storage.buckets (id, name, public) 
VALUES ('class-files', 'class-files', true)
ON CONFLICT (id) DO NOTHING;
