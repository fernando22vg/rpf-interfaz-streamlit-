# TODO - Rediseño UI `interfaz_analisis_RPF.py`

## Paso 1: Preparación
- [ ] Inspeccionar `interfaz_analisis_RPF.py` para localizar puntos de UI a refactorizar: sidebar, headers, emojis, tabs, st.metric.
- [x] Inspeccionar dónde se generan curvas reales vs simulación en Bloques 3 y 5 (pendiente de ajustes visuales).


## Paso 2: Nuevo framework UI
- [ ] Añadir tokens `TOKENS`, helper `icon()` y `inject_css()` (CSS único) dentro de `interfaz_analisis_RPF.py`.
- [ ] Limpiar mención/visual de COBEE y emojis (fase previa al rediseño).

- [ ] Implementar funciones UI: `render_top_bar`, `render_stepper`, `render_side_nav`, `render_unit_context`, `render_block_head`, `render_kpi_strip`, `render_controls`, `render_comparative_table`.


## Paso 3: Reemplazo layout (sin tocar lógica de negocio)
- [ ] Sustituir sidebar nativa por layout con top bar + stepper + side nav + unidad sticky.
- [ ] Eliminar header duplicado del Bloque 6.

## Paso 4: Limpieza de UI textual
- [ ] Quitar todos los emojis visibles.
- [ ] Reemplazar cualquier aparición visible de "COBEE" por "Sistema" o eliminarla.

## Paso 5: Visual y componentes
- [ ] Estilo tabs con CSS (pill activa + subrayado).
- [ ] Reemplazar `st.metric` (Bloques 3/4/5) por `render_kpi_strip()`.

## Paso 6: Restricciones visuales de gráficas
- [ ] Confirmar que Bloque 3 muestra solo curvas reales (sin simulación).
- [ ] Confirmar que Bloque 5 muestra comparativa Real vs Simulación.
- [ ] Verificar que las trazas de curvas de datos (frecuencia/potencia) no usan `dash`.
- [ ] Si aparece `dash` en curvas de datos reales/simuladas, corregir (sin tocar referencias).

## Paso 7: Validación
- [ ] Ejecutar `streamlit run ...interfaz_analisis_RPF.py` y validar checklist.

