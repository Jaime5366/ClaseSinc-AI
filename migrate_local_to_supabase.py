# -*- coding: utf-8 -*-
"""
Script de Migración masiva de Clases y Lecturas locales a Supabase DB + RAG.
Lee todas las clases y lecturas guardadas localmente en 'saved_classes/' y 'saved_readings/',
las registra en la tabla 'documents' de Supabase y genera sus vectores de 768 dimensiones en 'document_chunks'.
"""
import os
import json
import glob
from document_service import create_document, update_document_summary
from rag_service import chunk_text, insert_document_chunks


def migrate_all_local_data():
    print("=" * 65)
    print("INICIANDO MIGRACIÓN DE DATOS LOCALES A SUPABASE BD + RAG VECTORS")
    print("=" * 65)
    
    total_docs = 0
    total_chunks = 0
    errors = 0

    # 1. Migrar Clases Guardadas
    classes_dir = "saved_classes"
    if os.path.exists(classes_dir):
        print(f"\n[1/2] Procesando clases guardadas en '{classes_dir}'...")
        subject_folders = [f for f in os.listdir(classes_dir) if os.path.isdir(os.path.join(classes_dir, f))]
        
        for subject in subject_folders:
            folder_path = os.path.join(classes_dir, subject)
            json_files = glob.glob(os.path.join(folder_path, "*.json"))
            
            for json_path in json_files:
                try:
                    with open(json_path, "r", encoding="utf-8", errors="ignore") as f:
                        data = json.load(f)
                        
                    doc_title = f"[{subject}] {data.get('name', 'Clase sin título')}"
                    summary_text = data.get("summary", "")
                    docs_extracted = data.get("docs_extracted", "")
                    
                    if not summary_text:
                        continue
                        
                    # Crear registro de documento
                    doc = create_document(title=doc_title, file_type="class")
                    doc_id = doc["id"]
                    
                    # Actualizar resumen
                    update_document_summary(doc_id, summary_text, status="completed")
                    
                    # Generar fragmentos y vectores
                    full_content = f"{doc_title}\n\n{summary_text}\n\n{docs_extracted}"
                    chunks = chunk_text(full_content, chunk_size=800, overlap=100)
                    if chunks:
                        insert_document_chunks(doc_id, chunks)
                        total_chunks += len(chunks)
                        
                    total_docs += 1
                    print(f"  [OK] Clase migrada: '{doc_title}' ({len(chunks)} chunks)")
                except Exception as e:
                    errors += 1
                    print(f"  [ERROR] Falló migración de '{json_path}': {e}")

    # 2. Migrar Lecturas Guardadas
    readings_dir = "saved_readings"
    if os.path.exists(readings_dir):
        print(f"\n[2/2] Procesando lecturas guardadas en '{readings_dir}'...")
        subject_folders = [f for f in os.listdir(readings_dir) if os.path.isdir(os.path.join(readings_dir, f))]
        
        for subject in subject_folders:
            folder_path = os.path.join(readings_dir, subject)
            json_files = glob.glob(os.path.join(folder_path, "*.json"))
            
            for json_path in json_files:
                try:
                    with open(json_path, "r", encoding="utf-8", errors="ignore") as f:
                        data = json.load(f)
                        
                    pages = data.get("pages", [])
                    if isinstance(pages, list) and pages:
                        summary_text = "\n\n".join([p.get("content", "").strip() for p in pages if isinstance(p, dict) and p.get("content")])
                    else:
                        summary_text = data.get("summary", "") or data.get("content", "")

                    if not summary_text or not summary_text.strip():
                        continue
                        
                    # Crear registro de documento
                    doc = create_document(title=doc_title, file_type="reading")
                    doc_id = doc["id"]
                    
                    # Actualizar resumen
                    update_document_summary(doc_id, summary_text, status="completed")
                    
                    # Generar fragmentos y vectores
                    chunks = chunk_text(summary_text, chunk_size=800, overlap=100)
                    if chunks:
                        insert_document_chunks(doc_id, chunks)
                        total_chunks += len(chunks)
                        
                    total_docs += 1
                    print(f"  [OK] Lectura migrada: '{doc_title}' ({len(chunks)} chunks)")
                except Exception as e:
                    errors += 1
                    print(f"  [ERROR] Falló migración de '{json_path}': {e}")

    print("\n" + "=" * 65)
    print(f"MIGRACIÓN COMPLETADA SATISFACTORIAMENTE")
    print(f"  - Total Documentos Migrados: {total_docs}")
    print(f"  - Total Chunks Vectorizados: {total_chunks}")
    print(f"  - Total Errores: {errors}")
    print("=" * 65)

if __name__ == "__main__":
    migrate_all_local_data()
