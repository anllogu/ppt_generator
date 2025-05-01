import sys
import os
import yaml
import logging
from pptx import Presentation


def replace_text_in_shape(shape, replacements):
    # Procesa cuadros de texto normales
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for key, value in replacements.items():
                marcador = f'{{{key}}}'
                # Si el valor es una lista, lo convertimos en viñetas si el marcador está contenido en el texto del párrafo (aunque no sea lo único)
                full_text = "".join(run.text for run in paragraph.runs).strip()
                if isinstance(value, list) and marcador in full_text:
                    logging.info(f"Insertando viñetas para '{key}': {value}")
                    paragraph.clear()
                    for v in value:
                        p = shape.text_frame.add_paragraph()
                        p.text = f"• {v}"
                        p.level = 0
                else:
                    # Reemplazo robusto: reconstruye el texto completo del párrafo y reemplaza el marcador aunque esté partido en varios runs
                    if marcador in full_text:
                        logging.info(f"Reemplazando marcador '{marcador}' por '{value}' (runs partidos)")
                        new_text = full_text.replace(marcador, str(value))
                        # Borra todos los runs y crea uno nuevo con el texto reemplazado
                        for run in paragraph.runs:
                            run.text = ""
                        if paragraph.runs:
                            paragraph.runs[0].text = new_text
                        else:
                            paragraph.add_run().text = new_text
                    else:
                        # Reemplazo en runs individuales (caso simple)
                        for run in paragraph.runs:
                            if marcador in run.text:
                                logging.info(f"Reemplazando marcador '{marcador}' por '{value}'")
                                run.text = run.text.replace(marcador, str(value))
    # Procesa tablas si las hay
    if shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                replace_text_in_table_cell(cell, replacements)

def replace_text_in_table_cell(cell, replacements):
    for paragraph in cell.text_frame.paragraphs:
        for key, value in replacements.items():
            marcador = f'{{{key}}}'
            # Solo reemplazo simple en tablas (no viñetas)
            full_text = "".join(run.text for run in paragraph.runs)
            if marcador in full_text:
                logging.info(f"Reemplazando marcador '{marcador}' por '{value}' en celda de tabla")
                new_text = full_text.replace(marcador, str(value))
                for run in paragraph.runs:
                    run.text = ""
                if paragraph.runs:
                    paragraph.runs[0].text = new_text
                else:
                    paragraph.add_run().text = new_text
            else:
                for run in paragraph.runs:
                    if marcador in run.text:
                        logging.info(f"Reemplazando marcador '{marcador}' por '{value}' en celda de tabla")
                        run.text = run.text.replace(marcador, str(value))


def replace_text_in_presentation(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)


def parse_yaml(yaml_path):
    logging.info(f"Leyendo archivo YAML: {yaml_path}")
    with open(yaml_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    # Convierte strings multilínea con guion a listas para viñetas
    for k, v in data.items():
        if isinstance(v, str) and v.strip().startswith('- '):
            data[k] = [line[2:].strip() for line in v.strip().split('\n') if line.strip().startswith('- ')]
    logging.info(f"Datos leídos del YAML: {data}")
    return data


def main():
    logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')

    if len(sys.argv) != 4:
        print("Uso: python ppt_generator.py <plantilla.pptx> <contenido.yaml> <salida.pptx>")
        sys.exit(1)

    template_path = sys.argv[1]
    yaml_path = sys.argv[2]
    output_path = sys.argv[3]

    if not os.path.exists(template_path):
        logging.error(f"No se encuentra la plantilla {template_path}")
        print(f"Error: No se encuentra la plantilla {template_path}")
        sys.exit(1)
    if not os.path.exists(yaml_path):
        logging.error(f"No se encuentra el archivo YAML {yaml_path}")
        print(f"Error: No se encuentra el archivo YAML {yaml_path}")
        sys.exit(1)

    replacements = parse_yaml(yaml_path)
    logging.info(f"Abriendo plantilla: {template_path}")
    prs = Presentation(template_path)
    logging.info("Reemplazando textos en la presentación...")
    replace_text_in_presentation(prs, replacements)
    logging.info(f"Guardando presentación en: {output_path}")
    prs.save(output_path)
    print(f"Presentación generada correctamente: {output_path}")


if __name__ == "__main__":
    main()
