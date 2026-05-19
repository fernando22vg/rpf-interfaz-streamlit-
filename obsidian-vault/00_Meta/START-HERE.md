# START HERE (Claude Code + Obsidian)

Este vault sirve para que **Claude Code (VS Code)** tenga **contexto consistente** y tú no tengas que “explicar todo de nuevo”.

## Flujo recomendado (cada cambio)
1. Crea un ticket en `05_Tickets/`.
2. Pega/adjunta en el ticket:
   - **Objetivo**
   - **Archivos relevantes** (rutas)
   - **Snippet corto** (20–80 líneas como máximo)
   - **Criterios de éxito**
   - **Restricciones** / “no romper X”
   - **Cómo reproducir** (runbook si aplica)
3. En Claude Code pide:
   - **Planificar** primero
   - Luego **implementar**
   - Si falla: **diagnosticar con logs**
4. Vuelve al ticket y completa:
   - qué cambió
   - por qué
   - resultados (logs / archivos generados)

## Convenciones
- Usa enlaces internos con doble corchete: `[[ruta/archivo.py]]`
- En cada ticket, incluye siempre:
  - “**Qué espero que pase**”
  - “**Qué no debe cambiar**”
- No pegues el repo completo: pega solo lo necesario.

## Estructura que vas a usar
- `04_Runbooks/` → cómo ejecutar y depurar
- `03_Decisions/` → decisiones técnicas (ADR)
- `05_Tickets/` → tareas/cambios concretos
- `99_Prompts/Claude/` → plantillas listas para usar

> Consejo: si algo te toma más de 10 minutos explicar “lo que ya hicimos”, conviértelo en un Runbook o decisión y reutilízalo.
