#!/usr/bin/env python3
"""
sync_vault_to_kb.py — Sincroniza vault Obsidian → Open WebUI Knowledge Base

Lee todos los archivos .md del vault y los sube a una KB dedicada en Open WebUI,
reemplazando los wikilinks [[nota]] con el contenido real de esa nota para
maximizar el contexto disponible al modelo durante el RAG.

Uso:
  python3 sync_vault_to_kb.py
  python3 sync_vault_to_kb.py --vault-dir /ruta/vault --kb-name "RPF Obsidian"
  python3 sync_vault_to_kb.py --expand-links   # expande [[wikilinks]] con contenido real
"""

import os
import re
import time
import uuid
import logging
import argparse
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

WEBUI_URL   = os.getenv('OPEN_WEBUI_URL',   'http://localhost:3000')
WEBUI_EMAIL = os.getenv('OPEN_WEBUI_EMAIL', 'admin@cobee.com')
WEBUI_PASS  = os.getenv('OPEN_WEBUI_PASS',  'admin')
DEFAULT_VAULT = Path(__file__).parent / 'obsidian_vault'
KB_NAME     = 'RPF Obsidian COBEE'
KB_DESC     = 'Base de conocimiento estructurada con notas interconectadas sobre RPF COBEE'


def get_session():
    s = requests.Session()
    r = s.post(f'{WEBUI_URL}/api/v1/auths/signin',
               json={'email': WEBUI_EMAIL, 'password': WEBUI_PASS}, timeout=15)
    r.raise_for_status()
    token = r.json().get('token')
    s.headers.update({'Authorization': f'Bearer {token}'})
    return s


def get_or_create_kb(session, name: str, description: str) -> str:
    r = session.get(f'{WEBUI_URL}/api/v1/knowledge/', timeout=10)
    items = r.json().get('items', []) if isinstance(r.json(), dict) else r.json()
    for kb in items:
        if isinstance(kb, dict) and kb.get('name') == name:
            return kb['id']
    r2 = session.post(f'{WEBUI_URL}/api/v1/knowledge/create',
                      json={'name': name, 'description': description}, timeout=10)
    r2.raise_for_status()
    return r2.json()['id']


def clear_kb_files(session, kb_id: str):
    """Elimina todos los archivos de la KB y sus registros físicos."""
    r = session.get(f'{WEBUI_URL}/api/v1/knowledge/{kb_id}', timeout=10)
    files = r.json().get('files') or []
    for f in files:
        fid = f.get('id') or f.get('file', {}).get('id')
        if fid:
            session.delete(f'{WEBUI_URL}/api/v1/knowledge/{kb_id}/file/delete',
                           json={'file_id': fid}, timeout=10)
            session.delete(f'{WEBUI_URL}/api/v1/files/{fid}', timeout=10)
    if files:
        log.info(f"  {len(files)} archivos anteriores eliminados")


def expand_wikilinks(content: str, vault: Path) -> str:
    """Expande [[wikilinks]] con un resumen de la nota enlazada."""
    def replace_link(m):
        raw = m.group(1)
        # Soporte para [[nota|alias]] → usa 'nota' como nombre de archivo
        parts = raw.split('|')
        note_name = parts[0].strip()
        display   = parts[-1].strip()

        # Buscar el archivo en el vault
        candidates = list(vault.rglob(f'{note_name}.md'))
        if not candidates:
            return display  # si no existe, devolver solo el texto

        note_path = candidates[0]
        try:
            note_content = note_path.read_text(encoding='utf-8')
            # Extraer solo el primer párrafo significativo (después del frontmatter)
            lines = note_content.split('\n')
            in_frontmatter = False
            summary_lines = []
            for line in lines:
                if line.strip() == '---':
                    in_frontmatter = not in_frontmatter
                    continue
                if not in_frontmatter and line.strip() and not line.startswith('#'):
                    summary_lines.append(line.strip())
                    if len(summary_lines) >= 2:
                        break
            summary = ' '.join(summary_lines)[:200]
            return f'{display} ({summary})' if summary else display
        except Exception:
            return display

    return re.sub(r'\[\[([^\]]+)\]\]', replace_link, content)


def upload_note(session, kb_id: str, note_path: Path, vault: Path,
                expand_links: bool = False) -> bool:
    """Sube una nota markdown a la KB."""
    try:
        content = note_path.read_text(encoding='utf-8')
    except Exception as e:
        log.warning(f"No se pudo leer {note_path}: {e}")
        return False

    if expand_links:
        content = expand_wikilinks(content, vault)

    # Añadir UID único para evitar detección de contenido duplicado
    content = content + f'\n\n<!-- vault-sync:{uuid.uuid4().hex} -->\n'

    # Nombre relativo al vault como nombre del archivo
    rel_name = str(note_path.relative_to(vault)).replace('\\', '/').replace('/', '_')
    if not rel_name.endswith('.md'):
        rel_name += '.md'

    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md',
                                     encoding='utf-8', delete=False) as tf:
        tf.write(content)
        tmp = tf.name

    try:
        with open(tmp, 'rb') as f:
            r = session.post(f'{WEBUI_URL}/api/v1/files/',
                             files={'file': (rel_name, f, 'text/markdown')},
                             timeout=90)
        if r.status_code not in (200, 201):
            log.warning(f"Upload fallido '{rel_name}': {r.text[:80]}")
            return False
        file_id = r.json().get('id')

        for attempt in range(3):
            try:
                r2 = session.post(f'{WEBUI_URL}/api/v1/knowledge/{kb_id}/file/add',
                                  json={'file_id': file_id}, timeout=120)
                if r2.status_code in (200, 201):
                    return True
                log.warning(f"  KB add intento {attempt+1} ({r2.status_code}): {r2.text[:60]}")
            except Exception as e:
                log.warning(f"  KB add intento {attempt+1} timeout: {e}")
            if attempt < 2:
                time.sleep(5)
        return False
    finally:
        os.unlink(tmp)


def main():
    parser = argparse.ArgumentParser(description='Sincroniza vault Obsidian con Open WebUI')
    parser.add_argument('--vault-dir', default=str(DEFAULT_VAULT))
    parser.add_argument('--kb-name', default=KB_NAME)
    parser.add_argument('--expand-links', action='store_true',
                        help='Expandir [[wikilinks]] con resumen del contenido enlazado')
    parser.add_argument('--dry-run', action='store_true',
                        help='Solo listar archivos sin subir')
    args = parser.parse_args()

    vault = Path(args.vault_dir)
    if not vault.exists():
        log.error(f"Vault no encontrado: {vault}")
        log.error("Ejecuta primero: python3 generate_obsidian_vault.py")
        return

    notes = sorted(vault.rglob('*.md'))
    log.info(f"Vault: {vault}")
    log.info(f"Notas encontradas: {len(notes)}")

    if args.dry_run:
        for n in notes:
            log.info(f"  {n.relative_to(vault)}")
        return

    log.info("Conectando a Open WebUI...")
    session = get_session()
    log.info("Conectado OK")

    kb_id = get_or_create_kb(session, args.kb_name, KB_DESC)
    log.info(f"KB: {args.kb_name} (id: {kb_id})")
    clear_kb_files(session, kb_id)

    ok = 0
    for note in notes:
        rel = note.relative_to(vault)
        if upload_note(session, kb_id, note, vault, args.expand_links):
            ok += 1
            log.info(f"  ✓ {rel}")
        else:
            log.warning(f"  ✗ {rel}")

    log.info(f"\n✓ Sincronización completada: {ok}/{len(notes)} notas subidas a '{args.kb_name}'")


if __name__ == '__main__':
    main()
