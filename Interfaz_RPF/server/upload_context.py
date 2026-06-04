#!/usr/bin/env python3
"""
upload_context.py — Capa 4: Sube rpf_context.md al knowledge base de Open WebUI

Flujo:
  1. Login con email/password → JWT token
  2. Sube rpf_context.md como archivo
  3. Si el knowledge base "RPF COBEE" no existe, lo crea
  4. Reemplaza el archivo anterior por el nuevo (contexto actualizado)

Uso:
  python3 upload_context.py
  python3 upload_context.py --context /ruta/custom.md
"""

import os
import sys
import argparse
import logging
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

WEBUI_URL   = os.getenv('WEBUI_URL', 'http://localhost:3000')
WEBUI_EMAIL = os.getenv('WEBUI_EMAIL', '')
WEBUI_PASS  = os.getenv('WEBUI_PASSWORD', '')
KB_NAME     = 'RPF COBEE'
KB_DESC     = 'Contexto de Regulación Primaria de Frecuencia — unidades generadoras COBEE, SIN Bolivia'
DEFAULT_CTX = Path(__file__).parent / 'context' / 'rpf_context.md'


# ─── API helpers ─────────────────────────────────────────────────────────────

def login(session: requests.Session) -> str:
    resp = session.post(f'{WEBUI_URL}/api/v1/auths/signin',
                        json={'email': WEBUI_EMAIL, 'password': WEBUI_PASS},
                        timeout=10)
    if resp.status_code != 200:
        raise RuntimeError(f'Login fallido ({resp.status_code}): {resp.text[:200]}')
    token = resp.json().get('token')
    if not token:
        raise RuntimeError(f'No se obtuvo token: {resp.json()}')
    log.info('Login OK')
    return token


def upload_file(session: requests.Session, ctx_path: Path) -> str:
    """Sube el archivo Markdown y devuelve su file_id."""
    with open(ctx_path, 'rb') as f:
        resp = session.post(
            f'{WEBUI_URL}/api/v1/files/',
            files={'file': (ctx_path.name, f, 'text/markdown')},
            timeout=30,
        )
    if resp.status_code not in (200, 201):
        raise RuntimeError(f'Upload fallido ({resp.status_code}): {resp.text[:200]}')
    file_id = resp.json().get('id')
    log.info(f'Archivo subido — file_id: {file_id}')
    return file_id


def get_or_create_kb(session: requests.Session) -> str:
    """Devuelve el knowledge_id del KB 'RPF COBEE', creándolo si no existe."""
    resp = session.get(f'{WEBUI_URL}/api/v1/knowledge/', timeout=10)
    if resp.status_code != 200:
        raise RuntimeError(f'Error listando knowledge bases: {resp.text[:200]}')

    for kb in resp.json():
        if kb.get('name') == KB_NAME:
            log.info(f'Knowledge base existente — id: {kb["id"]}')
            return kb['id']

    # Crear nuevo
    resp = session.post(f'{WEBUI_URL}/api/v1/knowledge/create',
                        json={'name': KB_NAME, 'description': KB_DESC},
                        timeout=10)
    if resp.status_code not in (200, 201):
        raise RuntimeError(f'Error creando knowledge base: {resp.text[:200]}')
    kb_id = resp.json().get('id')
    log.info(f'Knowledge base creado — id: {kb_id}')
    return kb_id


def remove_old_files(session: requests.Session, kb_id: str):
    """Elimina archivos anteriores del knowledge base para reemplazarlos."""
    resp = session.get(f'{WEBUI_URL}/api/v1/knowledge/{kb_id}', timeout=10)
    if resp.status_code != 200:
        return
    files = resp.json().get('files', [])
    for f in files:
        fid = f.get('id') or (f.get('file', {}).get('id'))
        if fid:
            session.delete(f'{WEBUI_URL}/api/v1/knowledge/{kb_id}/file/delete',
                           json={'file_id': fid}, timeout=10)
            log.info(f'Archivo anterior eliminado: {fid}')


def add_file_to_kb(session: requests.Session, kb_id: str, file_id: str):
    resp = session.post(f'{WEBUI_URL}/api/v1/knowledge/{kb_id}/file/add',
                        json={'file_id': file_id}, timeout=30)
    if resp.status_code not in (200, 201):
        raise RuntimeError(f'Error añadiendo archivo al KB: {resp.text[:200]}')
    log.info(f'Archivo añadido al knowledge base "{KB_NAME}"')


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Sube contexto RPF a Open WebUI knowledge base')
    parser.add_argument('--context', type=Path, default=DEFAULT_CTX,
                        help='Ruta al archivo Markdown de contexto')
    args = parser.parse_args()

    if not args.context.exists():
        log.error(f'Archivo de contexto no encontrado: {args.context}')
        log.error('Ejecuta primero: python3 generate_context.py')
        sys.exit(1)

    if not WEBUI_EMAIL or not WEBUI_PASS:
        log.error('Faltan WEBUI_EMAIL o WEBUI_PASSWORD en el .env')
        sys.exit(1)

    session = requests.Session()

    token = login(session)
    session.headers.update({'Authorization': f'Bearer {token}'})

    file_id = upload_file(session, args.context)
    kb_id   = get_or_create_kb(session)
    remove_old_files(session, kb_id)
    add_file_to_kb(session, kb_id, file_id)

    size_kb = args.context.stat().st_size / 1024
    log.info(f'Capa 4 OK — "{KB_NAME}" actualizado ({size_kb:.1f} KB, file_id={file_id})')


if __name__ == '__main__':
    main()
