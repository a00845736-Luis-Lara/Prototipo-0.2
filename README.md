# Tablero Ejecutivo CD — Consumo Masivo
link al sheets: https://docs.google.com/spreadsheets/d/19c9CcFXaSLsj-E0y2qw0LQNUxh9X-xTR/edit?gid=1360834295#gid=1360834295

link al dashboard:
https://dashboardpy-hm4axwuddscomayxalotua.streamlit.app/ link
Dashboard interactivo en Python (Streamlit + Plotly) que lee directamente
`Tablero_Ejecutivo_CD.xlsx` y lo muestra como un tablero de control, siguiendo
el diseño de referencia (KPIs arriba, paneles por área, alertas, resumen ejecutivo).

## Archivos
- `dashboard.py` → la app.
- `Tablero_Ejecutivo_CD.xlsx` → tu Excel de datos (edítalo libremente).
- `requirements.txt` → librerías necesarias.

## Cómo correrlo en VS Code (paso a paso)

1. **Abre la carpeta** `dashboard_cd` en VS Code (File → Open Folder).

2. **Crea un entorno virtual** (recomendado, evita conflictos con otros
   proyectos de Python). Abre la terminal integrada de VS Code
   (``Ctrl + ñ`` o `Terminal → New Terminal`) y ejecuta:

   ```bash
   python -m venv venv
   ```

   Actívalo:
   - **Windows (PowerShell):** `venv\Scripts\Activate.ps1`
   - **Windows (cmd):** `venv\Scripts\activate.bat`
   - **Mac/Linux:** `source venv/bin/activate`

   VS Code normalmente detecta el entorno y te pregunta si quieres usarlo
   como intérprete del proyecto — dile que sí (o selecciónalo con
   `Ctrl+Shift+P` → "Python: Select Interpreter").

3. **Instala las dependencias:**

   ```bash
   pip install -r requirements.txt
   ```

4. **Corre la app:**

   ```bash
   streamlit run dashboard.py
   ```

   Se abrirá automáticamente en tu navegador (normalmente
   `http://localhost:8501`). Si no se abre solo, copia esa URL en el navegador.

5. **Para detenerla:** vuelve a la terminal y presiona `Ctrl + C`.

## Cómo actualizar los datos

- Abre `Tablero_Ejecutivo_CD.xlsx` en Excel y edita solo las celdas amarillas
  (Valor Hoy, Meta, Acum. Mes), tal como ya lo hacías.
- Guarda el Excel.
- En el navegador, haz clic en el botón **"🔄 Actualizar datos"** de la
  esquina superior derecha del tablero (o simplemente recarga la página).

No necesitas tocar el código para que los cambios del Excel se reflejen.

## Qué SÍ replica del mockup
- Fila superior de 10 KPIs principales con semáforo (verde/amarillo/rojo),
  meta y acumulado del mes.
- 4 paneles (Servicio al Cliente, Productividad, Inventarios, Costos) con
  tabla de indicadores, estado y mini-gráfica de tendencia.
- 5 tarjetas de área operativa (Recepción, Picking, Andenes, Transporte,
  E-commerce) con sus métricas del día y su tabla de indicadores.
- Donut de utilización de andenes/muelles.
- Top 5 Alertas y Resumen Ejecutivo, tal como están en el Excel.
- Tema oscuro tipo "sala de control", igual de espíritu al mockup.

## Qué es aproximado (por ahora)
- **Las mini-gráficas de tendencia ("7d" y "30 días") son ilustrativas.**
  Tu Excel actual solo guarda el valor de HOY, la meta y el acumulado del
  mes — no un histórico diario. Para que esas curvas sean 100% reales,
  lo más limpio es agregar una hoja nueva llamada `Historico` con una fila
  por día y columna por indicador (fecha, indicador, valor). Si quieres,
  te ayudo a armar esa hoja y a conectar el dashboard a datos históricos
  reales — es un cambio pequeño sobre este mismo código.
- El **"Mapa Operativo del CD"** (el layout de estantería/zonas) no se
  incluyó todavía porque no hay datos de zonas en el Excel; se puede
  agregar como una imagen estática del layout o como un plano interactivo
  si me compartes la disposición del CD (recepción, picking, muelles, etc.).

## Siguientes pasos sugeridos
- Agregar hoja `Historico` para tendencias reales.
- Agregar filtro real de "CD" si en el futuro manejan más de un centro de
  distribución (hoy el Excel es de un solo CD, por eso el selector está fijo
  en "Todos").
- Publicarlo en la nube (Streamlit Community Cloud, gratis) para que tu papá
  y su equipo lo vean desde un link sin instalar nada — puedo ayudarte con
  eso cuando quieras.
