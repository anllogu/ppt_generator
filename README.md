# Generador de Success Story CRM

Este proyecto es un script en Python que permite generar presentaciones PowerPoint (.pptx) de historias de éxito para CRM, a partir de una plantilla y un archivo YAML con los datos personalizados.

## Características

- Reemplazo automático de textos en la plantilla PowerPoint usando llaves (`{campo}`)
- Soporte para listas tipo viñetas desde YAML
- Uso sencillo desde la línea de comandos

## Requisitos

- Python 3.8 o superior
- macOS, Windows o Linux

## Instalación

1. Clona este repositorio o descarga el archivo `ppt_generator.py`.
2. Crea un entorno virtual con [uv](https://github.com/astral-sh/uv):

   ```sh
   uv venv .venv
   source .venv/bin/activate
   ```

3. Instala las dependencias ejecutando:

   ```sh
   uv pip install -r requirements.txt
   ```

## Uso

Ejecuta el script desde la terminal con:

```sh
python ppt_generator.py <plantilla.pptx> <contenido.yaml> <salida.pptx>
```

- `<plantilla.pptx>`: Ruta al archivo de plantilla PowerPoint.
- `<contenido.yaml>`: Archivo YAML con los datos a reemplazar.
- `<salida.pptx>`: Nombre o ruta del archivo PowerPoint de salida.

### Ejemplo de uso

```sh
python ppt_generator.py plantilla.pptx datos.yaml resultado.pptx
```

### Ejemplo de archivo YAML

```yaml
titulo: "Historia de Éxito"
cliente: "Empresa XYZ"
desafio:
  - "Reducción de costos"
  - "Mejora de procesos"
solucion: |
  - Implementación de CRM
  - Integración con sistemas existentes
```

## Notas

- Los campos en la plantilla deben estar entre llaves, por ejemplo: `{cliente}`
- Las listas en YAML se mostrarán como viñetas en la presentación.
- El script funciona en macOS, Windows y Linux.
