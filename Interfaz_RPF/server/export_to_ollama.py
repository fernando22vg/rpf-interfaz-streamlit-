#!/usr/bin/env python3
"""
export_to_ollama.py — Exporta el modelo fusionado a GGUF y actualiza Ollama

Convierte el modelo HuggingFace fusionado (rpf_cobee_merged/) a formato GGUF
y recrea el modelo 'rpf-cobee' en Ollama con los pesos entrenados.

Uso:
  python3 export_to_ollama.py
  python3 export_to_ollama.py --quant q4_k_m   # cuantización (default: q4_k_m)
"""

import os
import sys
import subprocess
import logging
import argparse
import shutil
from pathlib import Path

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

TRAINING_DIR   = Path(__file__).parent / 'training'
MERGED_DIR     = TRAINING_DIR / 'rpf_cobee_merged'
GGUF_DIR       = TRAINING_DIR / 'rpf_cobee_gguf'
MODELFILE_PATH = Path(__file__).parent / 'Modelfile.rpf'

SYSTEM_PROMPT = """Eres COBEE-AI, asistente técnico especializado en Regulación Primaria de Frecuencia (RPF) del Sistema Interconectado Nacional de Bolivia.

Tienes conocimiento profundo sobre:
- Las unidades generadoras de COBEE: BOT01-03, HUA01-02, ZON, CHU01-02, CAH01-02, SAI, CUT01-05, HAR01-02, SRO01-02, TIQ
- Análisis de droop, nadir de frecuencia, reserva girante y cumplimiento RPF
- Normativa CDM del CNDC boliviano (droop 4-6%, evaluación semestral)
- Diagnóstico de problemas en reguladores de velocidad
- Simulaciones RMS en DIgSILENT PowerFactory
- Corrección de modelos de generadores para condiciones iniciales RPF

Responde siempre en español con precisión técnica. Cuando tengas datos específicos de la base de datos, úsalos. Si no tienes datos suficientes, indica qué información adicional necesitarías."""


def check_llama_cpp():
    """Verifica si llama.cpp está disponible para conversión a GGUF."""
    # Buscar en ubicaciones comunes
    candidates = [
        Path('/usr/local/bin/llama-quantize'),
        Path('/usr/bin/llama-quantize'),
        Path.home() / 'llama.cpp' / 'llama-quantize',
        Path.home() / 'llama.cpp' / 'build' / 'bin' / 'llama-quantize',
    ]
    for p in candidates:
        if p.exists():
            return p.parent
    return None


def install_llama_cpp():
    """Instala llama.cpp desde pip (versión Python con soporte GGUF)."""
    log.info("Instalando llama-cpp-python para conversión GGUF...")
    try:
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install', '--quiet',
            'llama-cpp-python', 'gguf'
        ])
        return True
    except subprocess.CalledProcessError:
        return False


def convert_to_gguf(merged_dir: Path, gguf_dir: Path, quant: str) -> Path | None:
    """Convierte modelo HuggingFace a GGUF usando llama.cpp o transformers."""
    gguf_dir.mkdir(parents=True, exist_ok=True)
    gguf_path = gguf_dir / f'rpf_cobee_{quant}.gguf'

    # Método 1: script convert_hf_to_gguf.py de llama.cpp
    convert_script = None
    for candidate in [
        Path.home() / 'llama.cpp' / 'convert_hf_to_gguf.py',
        Path('/opt/llama.cpp/convert_hf_to_gguf.py'),
    ]:
        if candidate.exists():
            convert_script = candidate
            break

    if convert_script:
        log.info(f"Convirtiendo con llama.cpp: {convert_script}")
        cmd = [sys.executable, str(convert_script),
               str(merged_dir), '--outfile', str(gguf_path),
               '--outtype', 'f16']
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            log.info(f"Conversión exitosa: {gguf_path}")
            return gguf_path
        else:
            log.warning(f"convert_hf_to_gguf.py falló: {result.stderr[:200]}")

    # Método 2: llama-cpp-python con exportación directa
    try:
        log.info("Intentando conversión con librería gguf...")
        from transformers import AutoModelForCausalLM, AutoTokenizer
        import torch

        log.info("Cargando modelo fusionado...")
        tokenizer = AutoTokenizer.from_pretrained(str(merged_dir))
        model = AutoModelForCausalLM.from_pretrained(
            str(merged_dir), torch_dtype=torch.float16, device_map='cpu')

        # Guardar en formato safetensors para luego cuantizar
        log.info("Modelo cargado. Guarda en formato intermedio...")
        model.save_pretrained(str(gguf_dir / 'hf_model'))
        tokenizer.save_pretrained(str(gguf_dir / 'hf_model'))
        log.info(f"Modelo HF guardado en: {gguf_dir / 'hf_model'}")
        log.info("Necesitas instalar llama.cpp para el paso GGUF final.")
        log.info("Ver instrucciones al final del script.")
        return None
    except Exception as e:
        log.error(f"Conversión falló: {e}")
        return None


def quantize_gguf(gguf_f16: Path, quant: str, llama_bin_dir: Path) -> Path | None:
    """Cuantiza el GGUF F16 al tipo especificado."""
    quantize_bin = llama_bin_dir / 'llama-quantize'
    if not quantize_bin.exists():
        return None

    gguf_q = gguf_f16.parent / f'rpf_cobee_{quant}.gguf'
    cmd = [str(quantize_bin), str(gguf_f16), str(gguf_q), quant.upper()]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        log.info(f"Cuantización {quant} exitosa: {gguf_q}")
        return gguf_q
    else:
        log.warning(f"Cuantización falló: {result.stderr[:200]}")
        return None


def create_modelfile(gguf_path: Path) -> Path:
    """Crea un Modelfile de Ollama apuntando al GGUF entrenado."""
    modelfile_content = f"""FROM {gguf_path}

PARAMETER temperature 0.3
PARAMETER top_p 0.9
PARAMETER top_k 40
PARAMETER repeat_penalty 1.1
PARAMETER num_ctx 4096
PARAMETER num_gpu 999
PARAMETER num_thread 4

SYSTEM \"\"\"{SYSTEM_PROMPT}\"\"\"
"""
    mf_path = gguf_path.parent / 'Modelfile.trained'
    mf_path.write_text(modelfile_content, encoding='utf-8')
    log.info(f"Modelfile creado: {mf_path}")
    return mf_path


def update_ollama(modelfile_path: Path):
    """Actualiza el modelo rpf-cobee en Ollama."""
    log.info("Actualizando modelo rpf-cobee en Ollama...")
    result = subprocess.run(
        ['ollama', 'create', 'rpf-cobee', '-f', str(modelfile_path)],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        log.info("✓ Modelo rpf-cobee actualizado en Ollama con pesos entrenados")
        # Prueba rápida
        test = subprocess.run(
            ['ollama', 'run', 'rpf-cobee',
             '¿Cuál es tu especialidad? Responde en máximo 2 líneas.'],
            capture_output=True, text=True, timeout=60
        )
        if test.returncode == 0:
            log.info(f"Prueba: {test.stdout.strip()}")
    else:
        log.error(f"Error actualizando Ollama: {result.stderr[:300]}")


def main():
    parser = argparse.ArgumentParser(description='Exporta modelo entrenado a Ollama')
    parser.add_argument('--quant', default='q4_k_m',
                        choices=['q4_k_m', 'q5_k_m', 'q8_0', 'f16'],
                        help='Tipo de cuantización GGUF (default: q4_k_m)')
    parser.add_argument('--merged-dir', default=str(MERGED_DIR),
                        help='Directorio del modelo fusionado')
    args = parser.parse_args()

    merged_dir = Path(args.merged_dir)

    if not merged_dir.exists():
        log.error(f"Modelo fusionado no encontrado: {merged_dir}")
        log.error("Ejecuta primero: python3 train_qlora.py")
        sys.exit(1)

    log.info("=" * 50)
    log.info("EXPORTACIÓN MODELO COBEE-AI → OLLAMA")
    log.info("=" * 50)

    # Intentar conversión a GGUF
    gguf_path = convert_to_gguf(merged_dir, GGUF_DIR, args.quant)

    if gguf_path and gguf_path.exists():
        modelfile = create_modelfile(gguf_path)
        update_ollama(modelfile)
    else:
        log.info("\n" + "=" * 50)
        log.info("INSTRUCCIONES MANUALES PARA CONVERSIÓN GGUF")
        log.info("=" * 50)
        log.info("""
El modelo fusionado está en: """ + str(merged_dir) + """

Para convertir a GGUF manualmente:

1. Instalar llama.cpp:
   git clone https://github.com/ggerganov/llama.cpp ~/llama.cpp
   cd ~/llama.cpp && make -j$(nproc)
   pip install -r requirements/requirements-convert_hf_to_gguf.txt

2. Convertir a GGUF:
   python3 ~/llama.cpp/convert_hf_to_gguf.py \\
     """ + str(merged_dir) + """ \\
     --outfile """ + str(GGUF_DIR / f'rpf_cobee_f16.gguf') + """ \\
     --outtype f16

3. Cuantizar (reduce tamaño de 6GB a ~2GB):
   ~/llama.cpp/llama-quantize \\
     """ + str(GGUF_DIR / 'rpf_cobee_f16.gguf') + """ \\
     """ + str(GGUF_DIR / f'rpf_cobee_{args.quant}.gguf') + """ \\
     """ + args.quant.upper() + """

4. Actualizar Ollama:
   ollama create rpf-cobee -f """ + str(GGUF_DIR / 'Modelfile.trained') + """
""")


if __name__ == '__main__':
    main()
