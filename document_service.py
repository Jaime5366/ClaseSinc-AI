# -*- coding: utf-8 -*-
"""
Módulo de Servicio de Documentos para ClaseSinc AI.
Maneja el registro de documentos en Supabase DB y el almacenamiento de archivos en Supabase Storage.
"""
import uuid
from typing import Optional, Dict, Any
from supabase_client import get_supabase_client

BUCKET_NAME = "class-files"
DEFAULT_USER_ID = "00000000-0000-0000-0000-000000000000"


def create_document(title: str, file_type: str, user_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Inserta un nuevo documento en la tabla 'documents' con status = 'processing'.
    Retorna el diccionario con la información del documento creado (incluyendo su ID).
    """
    client = get_supabase_client()
    doc_id = str(uuid.uuid4())
    actual_user_id = user_id if (user_id and "-" in user_id and len(user_id) == 36) else DEFAULT_USER_ID
    
    data = {
        "id": doc_id,
        "title": title,
        "file_type": file_type,
        "user_id": actual_user_id,
        "status": "processing"
    }
    
    try:
        res = client.table("documents").insert(data).execute()
        if res.data and len(res.data) > 0:
            return res.data[0]
        return data
    except Exception as e:
        print(f"[DocumentService Warning] Falló inserción en BD Supabase: {e}")
        return data


def upload_document_file(document_id: str, file_bytes: bytes, file_extension: str, user_id: Optional[str] = None) -> str:
    """
    Sube el archivo al bucket de Supabase Storage 'class-files' en la ruta {user_id}/{document_id}.ext
    y actualiza la columna 'file_url' en la tabla 'documents'.
    Retorna la URL resultante del archivo.
    """
    client = get_supabase_client()
    clean_ext = file_extension.lstrip(".")
    actual_user_id = user_id if (user_id and "-" in user_id and len(user_id) == 36) else DEFAULT_USER_ID
    storage_path = f"{actual_user_id}/{document_id}.{clean_ext}"
    
    file_url = ""
    try:
        client.storage.from_(BUCKET_NAME).upload(
            path=storage_path,
            file=file_bytes,
            file_options={"upsert": "true"}
        )
        file_url = client.storage.from_(BUCKET_NAME).get_public_url(storage_path)
    except Exception as e:
        print(f"[DocumentService Warning] Falló carga a Storage: {e}")
        file_url = f"storage://{BUCKET_NAME}/{storage_path}"
    
    try:
        client.table("documents").update({"file_url": file_url}).eq("id", document_id).execute()
    except Exception as e:
        print(f"[DocumentService Warning] Falló actualización de file_url en BD: {e}")
        
    return file_url


def update_document_summary(document_id: str, summary_markdown: str, status: str = "completed") -> Dict[str, Any]:
    """
    Actualiza la fila del documento en 'documents' guardando 'summary_markdown' y el 'status' ('completed' o 'failed').
    """
    client = get_supabase_client()
    update_payload = {
        "summary_markdown": summary_markdown,
        "status": status
    }
    
    try:
        res = client.table("documents").update(update_payload).eq("id", document_id).execute()
        if res.data and len(res.data) > 0:
            return res.data[0]
        return update_payload
    except Exception as e:
        print(f"[DocumentService Warning] Falló actualización de resumen en BD: {e}")
        return update_payload
