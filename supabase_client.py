# -*- coding: utf-8 -*-
import os
from dotenv import load_dotenv
from supabase import create_client, Client

load_dotenv()

SUPABASE_URL = os.getenv('SUPABASE_URL')
SUPABASE_ANON_KEY = os.getenv('SUPABASE_ANON_KEY')

_supabase_client: Client | None = None

def get_supabase_client() -> Client:
    global _supabase_client
    if _supabase_client is not None:
        return _supabase_client
    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        raise RuntimeError(
            'Configuración incompleta: Asegúrate de definir SUPABASE_URL y SUPABASE_ANON_KEY en el archivo .env.'
        )
    _supabase_client = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)
    return _supabase_client

try:
    supabase = get_supabase_client()
except Exception:
    supabase = None
