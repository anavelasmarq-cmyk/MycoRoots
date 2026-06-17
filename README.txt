══════════════════════════════════════════════════════════════
  MycoRoots — HyphaPod Generator
  TFG Ingeniería Agroalimentaria · Universidad de Córdoba
  Autora: Ana Velasco Márquez
══════════════════════════════════════════════════════════════

ESTRUCTURA DE CARPETAS
──────────────────────
MycoRoots/
├── README.txt              ← este archivo
├── app.py                  ← interfaz web (Streamlit) ★ recomendada
├── requirements.txt        ← dependencias Python
├── scripts/
│   ├── hyphapod_generator_v4.py   ← generador (también ejecutable solo)
│   ├── hyphapod_plano_v5.py       ← generador de plano 2D (PDF + DXF)
│   ├── hyphapod_glb.py            ← generador de modelo 3D (.glb)
│   └── catalogo_especies_v5.xlsx  ← base de datos de especies
└── diseños/                ← archivos generados por el script de consola
    ├── HyphaPod_v3_NxM.xlsx
    ├── HyphaPod_Plano_v5.pdf
    ├── HyphaPod_Plano_v5.dxf
    └── HyphaPod_3D.glb

REQUISITOS
──────────
- Python 3.8 o superior
- Paquetes: streamlit, openpyxl, reportlab, ezdxf, matplotlib, pandas
  Instalar con:  pip install -r requirements.txt

═══════════════════════════════════════════════
  OPCIÓN A — Interfaz web (recomendada)
═══════════════════════════════════════════════
1. Abre una terminal (cmd o PowerShell) en la carpeta MycoRoots/.
2. Ejecuta:
     streamlit run app.py
3. Se abrirá el navegador automáticamente en http://localhost:8501
4. Configura la parcela en el panel izquierdo y pulsa "Generar diseño".
5. Descarga los archivos directamente desde la pestaña "Descargas".

═══════════════════════════════════════════════
  OPCIÓN B — Script de consola (Spyder/terminal)
═══════════════════════════════════════════════
1. Abre Spyder (o cualquier IDE/consola de Python).
2. Abre el archivo: scripts/hyphapod_generator_v4.py
3. Ejecútalo (F5 en Spyder).
4. Sigue las instrucciones en pantalla.
5. Los archivos generados aparecerán en la carpeta diseños/.

NOTAS
──────
- El catálogo (catalogo_especies_v5.xlsx) NO debe moverse de scripts/.
- La carpeta MycoRoots puede estar en cualquier ubicación del ordenador.
- En la interfaz web, los archivos se descargan directamente al navegador
  (no se guarda nada en diseños/ cuando se usa app.py).

══════════════════════════════════════════════════════════════
