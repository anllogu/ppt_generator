# Generador de Success Story CRM

Este proyecto es una aplicación de escritorio en Python que permite generar presentaciones PowerPoint (.pptx) de historias de éxito para CRM, a partir de una plantilla y datos personalizados.

## Características

- Interfaz gráfica intuitiva (Tkinter)
- Permite seleccionar una plantilla PowerPoint
- Personalización de cliente, sector, tecnología, país, título y subtítulo
- Secciones editables: Client Challenge, Business Value, Solution, Technical Information
- Exporta la presentación final en formato .pptx

## Requisitos

- Python 3.8 o superior
- macOS, Windows o Linux

## Instalación

1. Clona este repositorio o descarga el archivo `ppt_generator.py`.
2. Crea un entorno virtual con [uv](https://github.com/astral-sh/uv):

   ```
   uv venv .venv
   source .venv/bin/activate
   ```

3. Instala las dependencias ejecutando:

   ```
   uv pip install -r requirements.txt
   ```

## Uso

1. Ejecuta el script:

   ```
   python ppt_generator.py
   ```

2. Selecciona una plantilla PowerPoint (.pptx).
3. Completa los campos y secciones según tu caso de éxito.
4. Haz clic en "Generar PowerPoint" para guardar la presentación.

## Notas

- La plantilla debe tener cuadros de texto con textos de ejemplo similares a los que busca el script (por ejemplo: "Success Story", "CLIENT CHALLENGE", etc.).
- El script funciona en macOS, Windows y Linux.
