# Claude x VS Code + Obsidian (Vault para contexto)

Este vault está diseñado para ayudarte a programar con **Claude Code (VS Code)** usando Obsidian como **fuente de contexto**:
- decisiones (por qué cambiaste algo)
- runbooks (cómo reproducir)
- criterios de éxito
- “prompt packs” reutilizables para que Claude code actúe con consistencia

## 1) Estructura recomendada

- `00_Meta/`  
  Índices y convenciones globales.
- `01_Projects/<project>/`  
  Estado por proyecto (roadmap, notas grandes, enlaces a tickets).
- `02_Notes/`  
  Investigación, ideas, scratchpads.
- `03_Decisions/`  
  ADRs (Architecture Decision Records): decisiones técnicas y su justificación.
- `04_Runbooks/`  
  Cómo ejecutar scripts, reproducir errores y revisar logs.
- `05_Tickets/`  
  “Tickets” por cambio/feature: objetivo, archivos relevantes, riesgos, criterios de éxito.
- `99_Prompts/Claude/`  
  Plantillas para prompts de Claude Code.

## 2) Flujo de trabajo (rápido)

1. Abres o creas un ticket en `05_Tickets/` para el cambio.
2. Pegas/adjuntas en el ticket:
   - archivos relevantes (rutas)
   - snippet corto (solo lo necesario)
   - criterios de éxito
   - restricciones (no romper X, mantener Y, etc.)
3. En Claude Code:
   - pides “planificar y luego implementar” o “solo diagnóstico”
   - respondes con logs si falla
4. Actualizas el ticket con:
   - qué cambió
   - por qué
   - resultado / links a PR o commits (si aplica)

## 3) Convención de links a código

Usa enlaces internos tipo:
- `[[ProgramasLimpio/CargaCondIniciales_PF.py]]` (para recordarte el archivo)
- pega snippets cortos (20–80 líneas) y deja el resto en el repo.

> Consejo: **no** intentes meter todo el código en Obsidian. El vault es contexto y orquestación, no un repositorio.

## 4) Cómo “adaptar” a Claude (extension en VS Code)

Claude Code suele beneficiarse de una salida estructurada. Por eso, cada ticket y prompt pack usa una plantilla con:
- **Objetivo**
- **Contexto (archivos + snippet)**
- **Tareas**
- **Criterios de éxito**
- **Riesgos y rollback**
- **Qué necesito que Claude pregunte / confirme**

## 5) Checklist de instalación manual de Obsidian

Obsidian debes instalarlo tú (el sistema no puede instalar apps por ti), pero:
- Crea un vault nuevo apuntando a esta carpeta: `obsidian-vault/`
- En VS Code abre esa carpeta como referencia (opcional)
- Ya puedes usar las plantillas del vault
