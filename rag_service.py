# -*- coding: utf-8 -*-
"""
Módulo RAG y Servicio de Vectores (VectorService) para ClaseSinc AI.
Maneja el fragmentado de textos largos (chunking), generación de embeddings de 768 dimensiones
con la API de Gemini y búsqueda semántica por similitud de cosenos en Supabase.
"""
import os
import math
from typing import List, Dict, Any, Optional
from dotenv import load_dotenv
from google import genai
from supabase_client import get_supabase_client

load_dotenv()


def chunk_text(text: str, chunk_size: int = 800, overlap: int = 100) -> List[Dict[str, Any]]:
    """
    Divide un texto largo en fragmentos (chunks) de entre 500 y 1000 caracteres (por defecto 800)
    con una superposición (overlap) de 100 caracteres.
    """
    if not text or not text.strip():
        return []
        
    chunks = []
    start = 0
    text_length = len(text)
    chunk_index = 1
    
    # Estimación simple de página (asumiendo ~2000 caracteres por página)
    chars_per_page = 2000
    
    while start < text_length:
        end = min(start + chunk_size, text_length)
        
        # Ajustar fin al último espacio o punto para no cortar palabras a la mitad
        if end < text_length:
            last_space = text.rfind(" ", start, end)
            if last_space != -1 and last_space > start + (chunk_size // 2):
                end = last_space
                
        chunk_content = text[start:end].strip()
        if chunk_content:
            page_num = max(1, math.ceil((start + 1) / chars_per_page))
            chunks.append({
                "chunk_index": chunk_index,
                "content": chunk_content,
                "page_number": page_num,
                "start_char": start,
                "end_char": end
            })
            chunk_index += 1
            
        start = end - overlap if end < text_length else text_length
        if start >= text_length:
            break
            
    return chunks


def generate_embedding(text: str, api_key: Optional[str] = None) -> List[float]:
    """
    Genera un vector de embedding de dimensión 768 utilizando la API de Gemini.
    """
    key = api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        raise ValueError("No se proporcionó GEMINI_API_KEY para generar el embedding.")
        
    client = genai.Client(api_key=key)
    res = client.models.embed_content(
        model="models/gemini-embedding-001",
        contents=text,
        config={"output_dimensionality": 768}
    )
    
    vector = res.embeddings[0].values
    # Asegurar dimensión exacta de 768
    if len(vector) > 768:
        return list(vector[:768])
    return list(vector)


def insert_document_chunks(
    document_id: str,
    chunks: List[Dict[str, Any]],
    user_id: str = "default_user",
    api_key: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Calcula los embeddings para cada fragmento e inserta los registros en la tabla 'document_chunks'.
    """
    if not chunks:
        return []
        
    client = get_supabase_client()
    rows_to_insert = []
    
    for chunk in chunks:
        embedding_vector = generate_embedding(chunk["content"], api_key=api_key)
        row = {
            "document_id": document_id,
            "content": chunk["content"],
            "embedding": embedding_vector,
            "page_number": chunk.get("page_number", 1)
        }
        rows_to_insert.append(row)
        
    inserted_rows = []
    try:
        res = client.table("document_chunks").insert(rows_to_insert).execute()
        if res.data:
            inserted_rows = res.data
        else:
            inserted_rows = rows_to_insert
    except Exception as e:
        print(f"[RAGService Warning] Falló inserción de chunks en Supabase DB: {e}")
        inserted_rows = rows_to_insert
        
    return inserted_rows


def _cosine_similarity(vec_a: List[float], vec_b: List[float]) -> float:
    """Calcula la similitud de cosenos entre dos vectores de igual dimensión."""
    dot_product = sum(a * b for a, b in zip(vec_a, vec_b))
    norm_a = math.sqrt(sum(a * a for a in vec_a))
    norm_b = math.sqrt(sum(b * b for b in vec_b))
    if norm_a == 0 or norm_b == 0:
        return 0.0
    return dot_product / (norm_a * norm_b)


def search_similar_chunks(
    document_id: str,
    query_text: str,
    top_k: int = 5,
    api_key: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Busca los fragmentos más semejantes al query_text dentro de un documento.
    Paso 1: Transforma query_text en un vector de embedding (dim=768).
    Paso 2: Realiza la consulta por similitud vectorial en Supabase (o fallback local).
    """
    query_vector = generate_embedding(query_text, api_key=api_key)
    client = get_supabase_client()
    
    # Método A: Intentar llamada RPC nativa a Supabase
    try:
        rpc_res = client.rpc("match_document_chunks", {
            "query_embedding": query_vector,
            "match_count": top_k,
            "filter_document_id": document_id
        }).execute()
        if rpc_res.data:
            return rpc_res.data
    except Exception as e:
        print(f"[RAGService Notice] RPC nativo match_document_chunks no disponible: {e}. Usando fallback local...")

    # Método B: Fallback recuperando fragmentos del documento y calculando cosenos localmente
    try:
        res = client.table("document_chunks").select("*").eq("document_id", document_id).execute()
        chunks_data = res.data or []
        
        scored_chunks = []
        for row in chunks_data:
            emb = row.get("embedding")
            if isinstance(emb, list) and len(emb) > 0:
                score = _cosine_similarity(query_vector, emb)
                scored_chunks.append({**row, "similarity": score})
                
        scored_chunks.sort(key=lambda x: x.get("similarity", 0), reverse=True)
        return scored_chunks[:top_k]
    except Exception as e:
        print(f"[RAGService Error] Error durante la búsqueda por similitud: {e}")
        return []
