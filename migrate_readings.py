from pathlib import Path
import json
from document_service import create_document, update_document_summary
from rag_service import chunk_text, insert_document_chunks

def main():
    readings_dir = Path("saved_readings")
    if not readings_dir.exists():
        print("No existe la carpeta saved_readings.")
        return

    total = 0
    errors = 0
    print("Iniciando migración Unicode pathlib de todas las lecturas...")

    for subject_dir in readings_dir.iterdir():
        if not subject_dir.is_dir():
            continue
        subject_name = subject_dir.name
        
        for json_file in subject_dir.glob("*.json"):
            try:
                with open(json_file, "r", encoding="utf-8", errors="ignore") as f:
                    data = json.load(f)
                    
                name = data.get("name", "").strip() or json_file.stem
                pages = data.get("pages", [])
                
                if isinstance(pages, list) and len(pages) > 0:
                    summary_text = "\n\n".join([p.get("content", "").strip() for p in pages if isinstance(p, dict) and p.get("content")])
                else:
                    summary_text = data.get("summary", "") or data.get("content", "")

                if not summary_text or not summary_text.strip():
                    print(f"  [SKIP] Sin texto: '{json_file.name}'")
                    continue

                doc_title = f"[{subject_name}] {name}"
                doc = create_document(title=doc_title, file_type="reading")
                doc_id = doc["id"]
                
                update_document_summary(doc_id, summary_text, status="completed")
                
                chunks = chunk_text(summary_text, chunk_size=800, overlap=100)
                if chunks:
                    insert_document_chunks(doc_id, chunks)
                    
                total += 1
                print(f"  [OK] #{total} Lectura: '{doc_title}' ({len(chunks)} chunks)")
            except Exception as e:
                errors += 1
                print(f"  [ERROR] '{json_file.name}': {e}")

    print("=" * 60)
    print(f"MIGRACIÓN FINALIZADA SATISFACTORIAMENTE")
    print(f"Total lecturas procesadas: {total}, Errores: {errors}")
    print("=" * 60)

if __name__ == "__main__":
    main()
