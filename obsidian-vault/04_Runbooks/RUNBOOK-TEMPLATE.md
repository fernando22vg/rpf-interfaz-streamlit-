# RUNBOOK: {{NOMBRE}}

## Resumen
{{Qué problema/objetivo cubre este runbook}}

## Requisitos
- {{dependencias / versión de Python / venv}}
- {{rutas necesarias}}
- {{archivos de entrada requeridos}}

## Comando(s) para ejecutar
- `{{comando 1}}`
- `{{comando 2}}`

## Qué esperar (señales de éxito)
- {{ej: genera archivos X}}
- {{ej: no hay traceback}}
- {{ej: imprime resumen con ...}}

## Qué hacer si falla
1. **Capturar evidencia**
   - Log stdout/stderr:
     - {{ruta/log}}
   - Estado actual:
     - {{qué archivos cambió}}
2. **Diagnóstico rápido (hipótesis)**
   - Hipótesis 1: {{...}}
   - Hipótesis 2: {{...}}
3. **Pruebas de validación**
   - Prueba A: {{...}}
   - Prueba B: {{...}}
4. **Fix (paso a paso)**
   - Paso 1:
   - Paso 2:
5. **Rollback**
   - {{cómo revertir}}

## Versión / cambios
- Fecha: {{YYYY-MM-DD}}
- Cambios: {{...}}
