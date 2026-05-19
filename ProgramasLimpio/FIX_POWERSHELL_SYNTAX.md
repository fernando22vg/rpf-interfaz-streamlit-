🔧 SOLUCIÓN: PowerShell no entiende "&&"
═══════════════════════════════════════════════════════════════════════════════

ERROR:
  El token '&&' no es un separador de instrucciones válido en esta versión.

CAUSA:
  PowerShell usa diferente sintaxis que CMD. El "&&" es de CMD.
  En PowerShell se usa ";" para separar comandos.

SOLUCIÓN: EJECUTAR CADA COMANDO SEPARADO (30 segundos)
═════════════════════════════════════════════════════════════════════════════════

Ejecuta UNO POR UNO (no todos juntos):

┌─────────────────────────────────────────────────────────────────────────┐
│                                                                         │
│  COMANDO 1:                                                             │
│  git pull origin main --allow-unrelated-histories                       │
│                                                                         │
│  (Espera a que termine. Verás: "Already up to date" o cambios)         │
│                                                                         │
│  COMANDO 2:                                                             │
│  git add -A                                                             │
│                                                                         │
│  (Sin output visible, es normal)                                        │
│                                                                         │
│  COMANDO 3:                                                             │
│  git push -u origin main                                               │
│                                                                         │
│  (Verás progreso de upload)                                            │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘


COPIA Y PEGA ESTOS COMANDOS UNO A UNO
═════════════════════════════════════════════════════════════════════════════════

COMANDO 1 (copiar y pegar):
──────────────────────────────────────────────────────────────────────────
git pull origin main --allow-unrelated-histories
──────────────────────────────────────────────────────────────────────────

Presiona ENTER. Espera a que termine.


COMANDO 2 (copiar y pegar):
──────────────────────────────────────────────────────────────────────────
git add -A
──────────────────────────────────────────────────────────────────────────

Presiona ENTER. (No verás output, es normal)


COMANDO 3 (copiar y pegar):
──────────────────────────────────────────────────────────────────────────
git push -u origin main
──────────────────────────────────────────────────────────────────────────

Presiona ENTER. Espera a que termine.


═════════════════════════════════════════════════════════════════════════════


SI QUIERES HACERLO EN UNA SOLA LÍNEA (Sintaxis PowerShell):
═════════════════════════════════════════════════════════════════════════════

En PowerShell el separador es ";" (punto y coma):

git pull origin main --allow-unrelated-histories ; git add -A ; git push -u origin main

(Pero recomiendo hacer uno por uno para ver dónde hay error si lo hay)


═════════════════════════════════════════════════════════════════════════════


ALTERNATIVA: USA CMD EN LUGAR DE POWERSHELL
═════════════════════════════════════════════════════════════════════════════

Si prefieres usar CMD (que sí entiende &&):

1. Cierra PowerShell
2. Abre CMD (tecla Windows → escribe "cmd" → Enter)
3. Ejecuta:
   cd c:\Programas\ Python\ProgramasLimpio
   git pull origin main --allow-unrelated-histories && git add -A && git push -u origin main

(Ahí el && funciona correctamente)


═════════════════════════════════════════════════════════════════════════════


RESUMEN
═════════════════════════════════════════════════════════════════════════════

OPCIÓN A (PowerShell - una línea):
  git pull origin main --allow-unrelated-histories ; git add -A ; git push -u origin main

OPCIÓN B (PowerShell - 3 líneas):
  git pull origin main --allow-unrelated-histories
  git add -A
  git push -u origin main

OPCIÓN C (CMD - una línea):
  git pull origin main --allow-unrelated-histories && git add -A && git push -u origin main


===== YO RECOMIENDO OPCIÓN B (Una por una en PowerShell) =====
Así ves qué pasa en cada paso.


═════════════════════════════════════════════════════════════════════════════
