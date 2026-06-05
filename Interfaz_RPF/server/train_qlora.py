#!/usr/bin/env python3
"""
train_qlora.py — Fine-tuning QLoRA 4-bit del modelo COBEE-AI

Entrena qwen2.5:3b con QLoRA sobre el dataset Q&A de RPF COBEE.
Optimizado para GTX 950M (4 GB VRAM) con Open WebUI detenido.

IMPORTANTE: Detener Open WebUI antes de ejecutar:
  docker stop open-webui

Uso:
  python3 train_qlora.py
  python3 train_qlora.py --epochs 3 --batch-size 1
  python3 train_qlora.py --check-gpu   # solo verificar hardware

Salida:
  training/rpf_cobee_lora/   ← adaptadores LoRA
  training/rpf_cobee_merged/ ← modelo completo fusionado
"""

import os
import sys
import json
import logging
import argparse
from pathlib import Path

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

TRAINING_DIR  = Path(__file__).parent / 'training'
DATASET_FILE  = TRAINING_DIR / 'qa_dataset.jsonl'
LORA_OUT_DIR  = TRAINING_DIR / 'rpf_cobee_lora'
MERGED_OUT_DIR = TRAINING_DIR / 'rpf_cobee_merged'

# Modelo base — mismo que usa Ollama pero desde HuggingFace
HF_MODEL = 'Qwen/Qwen2.5-3B-Instruct'

# Parámetros QLoRA conservadores para 4 GB VRAM
LORA_R         = 16
LORA_ALPHA     = 32
LORA_DROPOUT   = 0.05
TARGET_MODULES = ['q_proj', 'k_proj', 'v_proj', 'o_proj',
                  'gate_proj', 'up_proj', 'down_proj']

MAX_SEQ_LEN    = 512    # reducido al mínimo para CPU con poca RAM
GRAD_ACCUM     = 4      # reducido para CPU


def check_gpu():
    """Verifica GPU disponible y VRAM."""
    try:
        import torch
        if torch.cuda.is_available():
            name = torch.cuda.get_device_name(0)
            vram = torch.cuda.get_device_properties(0).total_memory / 1e9
            log.info(f"GPU: {name} — {vram:.1f} GB VRAM")
            if vram < 3.5:
                log.warning("VRAM < 3.5 GB — entrenamiento puede fallar por OOM")
                log.warning("Considera usar --cpu-only si falla")
            return True, vram
        else:
            log.warning("No se detectó GPU CUDA — se usará CPU (lento)")
            return False, 0.0
    except ImportError:
        log.error("PyTorch no instalado")
        return False, 0.0


def install_dependencies():
    """Instala dependencias si faltan."""
    deps = {
        'transformers': 'transformers>=4.40',
        'peft': 'peft>=0.10',
        'trl': 'trl>=0.8',
        'bitsandbytes': 'bitsandbytes>=0.43',
        'accelerate': 'accelerate>=0.28',
        'datasets': 'datasets>=2.18',
    }
    missing = []
    for mod, pkg in deps.items():
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)

    if missing:
        log.info(f"Instalando dependencias faltantes: {missing}")
        import subprocess
        # Intentar primero sin --break-system-packages, luego con él
        result = subprocess.run(
            [sys.executable, '-m', 'pip', 'install', '--quiet'] + missing,
            capture_output=True)
        if result.returncode != 0:
            log.info("pip normal falló, intentando con --break-system-packages...")
            subprocess.check_call(
                [sys.executable, '-m', 'pip', 'install', '--quiet',
                 '--break-system-packages'] + missing)
        log.info("Dependencias instaladas OK")


def load_dataset_hf(dataset_file: Path):
    """Carga el JSONL y lo convierte a HuggingFace Dataset."""
    from datasets import Dataset

    records = []
    with open(dataset_file, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    records.append(json.loads(line))
                except Exception:
                    pass

    log.info(f"Dataset: {len(records)} pares Q&A cargados")
    return Dataset.from_list(records)


def format_chat(sample, tokenizer):
    """Formatea un par Q&A al formato de chat del modelo."""
    return tokenizer.apply_chat_template(
        sample['messages'],
        tokenize=False,
        add_generation_prompt=False,
    )


def train(args):
    """Ejecuta el fine-tuning QLoRA."""
    install_dependencies()

    import torch
    from transformers import (AutoModelForCausalLM, AutoTokenizer,
                              BitsAndBytesConfig)
    from peft import LoraConfig, get_peft_model, TaskType
    import trl as _trl
    trl_version = tuple(int(x) for x in _trl.__version__.split('.')[:2])
    if trl_version >= (0, 9):
        from trl import SFTTrainer, SFTConfig as TrainingArguments
    else:
        from trl import SFTTrainer
        from transformers import TrainingArguments

    has_gpu, vram = check_gpu()
    use_4bit = has_gpu and not args.cpu_only

    # ── Configuración de cuantización ─────────────────────────────────────────
    bnb_config = None
    if use_4bit:
        log.info("Usando cuantización 4-bit NF4 (QLoRA)")
        bnb_config = BitsAndBytesConfig(
            load_in_4bit=True,
            bnb_4bit_quant_type='nf4',
            bnb_4bit_compute_dtype=torch.float16,
            bnb_4bit_use_double_quant=True,
        )
    else:
        log.info("Usando CPU / float32 (sin cuantización)")

    # ── Cargar tokenizer y modelo ──────────────────────────────────────────────
    log.info(f"Cargando modelo base: {HF_MODEL}")
    log.info("(Primera vez descarga ~6 GB desde HuggingFace — puede tardar)")

    tokenizer = AutoTokenizer.from_pretrained(HF_MODEL, trust_remote_code=True)
    if tokenizer.pad_token is None:
        tokenizer.pad_token = tokenizer.eos_token

    model_kwargs = {
        'trust_remote_code': True,
        'device_map': 'auto' if has_gpu and not args.cpu_only else 'cpu',
    }
    if bnb_config:
        model_kwargs['quantization_config'] = bnb_config
    else:
        model_kwargs['torch_dtype'] = torch.float16  # float16 usa la mitad de RAM que float32

    model = AutoModelForCausalLM.from_pretrained(HF_MODEL, **model_kwargs)
    model.config.use_cache = False  # necesario para entrenamiento con gradient checkpointing

    # ── Configurar LoRA ────────────────────────────────────────────────────────
    log.info(f"Aplicando LoRA: r={LORA_R}, alpha={LORA_ALPHA}")
    lora_config = LoraConfig(
        r=LORA_R,
        lora_alpha=LORA_ALPHA,
        lora_dropout=LORA_DROPOUT,
        target_modules=TARGET_MODULES,
        task_type=TaskType.CAUSAL_LM,
        bias='none',
    )
    model = get_peft_model(model, lora_config)
    model.print_trainable_parameters()

    # ── Cargar y preparar dataset ──────────────────────────────────────────────
    if not DATASET_FILE.exists():
        log.error(f"Dataset no encontrado: {DATASET_FILE}")
        log.error("Ejecuta primero: python3 create_qa_dataset.py && python3 expand_dataset.py")
        sys.exit(1)

    dataset = load_dataset_hf(DATASET_FILE)

    def tokenize(sample):
        text = format_chat(sample, tokenizer)
        return {'text': text}

    dataset = dataset.map(tokenize, remove_columns=['messages'])

    # Split train/eval 90/10
    split = dataset.train_test_split(test_size=0.1, seed=42)
    train_ds = split['train']
    eval_ds  = split['test']
    log.info(f"Train: {len(train_ds)} | Eval: {len(eval_ds)}")

    # ── Argumentos de entrenamiento ────────────────────────────────────────────
    LORA_OUT_DIR.mkdir(parents=True, exist_ok=True)

    # Ajustar batch size según VRAM disponible
    batch_size = args.batch_size
    if use_4bit and vram < 4.5:
        batch_size = 1
        log.info("VRAM limitada — batch_size forzado a 1")

    # ── Argumentos de entrenamiento (SFTConfig en TRL>=0.9, TrainingArguments antes)
    common_kwargs = dict(
        output_dir=str(LORA_OUT_DIR),
        num_train_epochs=args.epochs,
        per_device_train_batch_size=batch_size,
        per_device_eval_batch_size=1,
        gradient_accumulation_steps=GRAD_ACCUM,
        gradient_checkpointing=True,
        optim='paged_adamw_8bit' if use_4bit else 'adamw_torch',
        learning_rate=2e-4,
        lr_scheduler_type='cosine',
        warmup_steps=10,
        logging_steps=10,
        eval_strategy='epoch',
        save_strategy='epoch',
        load_best_model_at_end=True,
        fp16=use_4bit,
        bf16=False,
        max_grad_norm=0.3,
        report_to='none',
        dataloader_pin_memory=False,
    )
    if trl_version >= (0, 9):
        # SFTConfig acepta max_seq_length, packing y dataset_text_field
        training_args = TrainingArguments(
            **common_kwargs,
            max_seq_length=MAX_SEQ_LEN,
            packing=False,
            dataset_text_field='text',
        )
    else:
        from transformers import TrainingArguments as _TA
        training_args = _TA(**common_kwargs)

    # ── Entrenamiento ──────────────────────────────────────────────────────────
    sft_kwargs = dict(
        model=model,
        args=training_args,
        train_dataset=train_ds,
        eval_dataset=eval_ds,
    )
    if trl_version >= (0, 9):
        sft_kwargs['processing_class'] = tokenizer
    else:
        sft_kwargs['tokenizer'] = tokenizer
        sft_kwargs['dataset_text_field'] = 'text'
        sft_kwargs['max_seq_length'] = MAX_SEQ_LEN
        sft_kwargs['packing'] = False

    trainer = SFTTrainer(**sft_kwargs)

    log.info("=" * 50)
    log.info("INICIANDO ENTRENAMIENTO")
    log.info(f"Épocas: {args.epochs} | Batch: {batch_size} | Grad accum: {GRAD_ACCUM}")
    log.info(f"Effective batch size: {batch_size * GRAD_ACCUM}")
    log.info("=" * 50)

    trainer.train()
    trainer.save_model(str(LORA_OUT_DIR))
    tokenizer.save_pretrained(str(LORA_OUT_DIR))
    log.info(f"Adaptadores LoRA guardados en: {LORA_OUT_DIR}")


def merge_and_save(args):
    """Fusiona adaptadores LoRA con el modelo base."""
    install_dependencies()

    import torch
    from transformers import AutoModelForCausalLM, AutoTokenizer
    from peft import PeftModel

    if not LORA_OUT_DIR.exists():
        log.error(f"Adaptadores no encontrados: {LORA_OUT_DIR}")
        log.error("Ejecuta primero el entrenamiento")
        sys.exit(1)

    log.info("Cargando modelo base para fusión...")
    tokenizer = AutoTokenizer.from_pretrained(HF_MODEL, trust_remote_code=True)
    base_model = AutoModelForCausalLM.from_pretrained(
        HF_MODEL, torch_dtype=torch.float16,
        device_map='cpu', trust_remote_code=True)

    log.info("Fusionando adaptadores LoRA...")
    model = PeftModel.from_pretrained(base_model, str(LORA_OUT_DIR))
    model = model.merge_and_unload()

    MERGED_OUT_DIR.mkdir(parents=True, exist_ok=True)
    model.save_pretrained(str(MERGED_OUT_DIR))
    tokenizer.save_pretrained(str(MERGED_OUT_DIR))
    log.info(f"Modelo fusionado guardado en: {MERGED_OUT_DIR}")
    log.info("Siguiente paso: python3 export_to_ollama.py")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Fine-tuning QLoRA para COBEE-AI')
    parser.add_argument('--epochs', type=int, default=3, help='Épocas de entrenamiento')
    parser.add_argument('--batch-size', type=int, default=1, help='Batch size por dispositivo')
    parser.add_argument('--cpu-only', action='store_true', help='Forzar CPU (sin GPU)')
    parser.add_argument('--check-gpu', action='store_true', help='Solo verificar GPU y salir')
    parser.add_argument('--merge-only', action='store_true',
                        help='Solo fusionar LoRA existente (sin entrenar)')
    args = parser.parse_args()

    if args.check_gpu:
        install_dependencies()
        has_gpu, vram = check_gpu()
        sys.exit(0 if has_gpu else 1)

    if args.merge_only:
        merge_and_save(args)
        return

    train(args)
    log.info("Entrenamiento completado. Fusionando modelo...")
    merge_and_save(args)
    log.info("\n✓ Proceso completo. Ejecuta: python3 export_to_ollama.py")


if __name__ == '__main__':
    main()
