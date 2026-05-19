# Ejecución Remota de DIgSILENT desde Streamlit Cloud

## Situación actual

Los botones que invocan scripts de DIgSILENT/PowerFactory están **deshabilitados en Streamlit Cloud** porque el servidor remoto (Linux) no tiene DIgSILENT instalado. Solo funcionan cuando la app corre localmente en una PC con DIgSILENT instalado.

## Posible implementación futura: Agente local con túnel

Es posible conectar Streamlit Cloud con una instalación local de DIgSILENT mediante un agente HTTP y un túnel de red.

### Arquitectura

```
Streamlit Cloud          Túnel (ngrok)           PC con DIgSILENT
      ☁️          ──HTTP──>  🔗 URL pública  ──>  🖥️ Agente local
  botón click                                      ejecuta PF scripts
                                                   devuelve logs/resultados
```

### Componentes necesarios

| Componente | Descripción |
|---|---|
| **Agente local** | Servidor FastAPI (~50 líneas) que recibe comandos y lanza `subprocess` hacia DIgSILENT |
| **ngrok** | Crea una URL pública gratuita que apunta al agente local |
| **Modificación en la app** | Los botones hacen `requests.post(ngrok_url, ...)` en vez de `subprocess.Popen` directamente |
| **Streamlit secret** | La URL de ngrok se guarda en los secrets de Streamlit Cloud |

### Flujo de ejecución

1. El operador enciende la PC con DIgSILENT y lanza el agente local
2. ngrok genera una URL pública (ej. `https://abc123.ngrok.io`)
3. El operador actualiza el secret `AGENT_URL` en Streamlit Cloud
4. Cualquier usuario en la nube puede lanzar simulaciones; el agente las ejecuta localmente y devuelve resultados

### Limitación principal

La PC con DIgSILENT debe estar **encendida y con el agente activo** cada vez que se quieran usar las funciones de simulación desde la nube. Si se apaga, esas funciones no responden.

### Cuándo tiene sentido implementarlo

- Hay un equipo de usuarios que necesita lanzar simulaciones desde la nube sin instalar nada localmente
- Existe una PC dedicada (servidor) con DIgSILENT que puede estar encendida de forma continua
- Se quiere centralizar la ejecución en una sola licencia de DIgSILENT

### Cuándo NO es necesario

- El usuario que corre simulaciones ya tiene DIgSILENT instalado y puede correr la app localmente con `streamlit run interfaz_analisis_RPF.py`
- La nube se usa solo para visualización y análisis de resultados ya generados
