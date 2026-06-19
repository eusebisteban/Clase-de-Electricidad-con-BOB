import zipfile
import xml.etree.ElementTree as ET
import os
import re
import sys

# Configurar codificación UTF-8
sys.stdout.reconfigure(encoding='utf-8')

def extract_text_from_pptx(pptx_file):
    """Extrae todo el texto de un archivo PowerPoint"""
    slide_data = []
    try:
        with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
            # Obtener lista de slides
            slides = [name for name in zip_ref.namelist()
                     if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
            slides.sort()
            
            print(f"\n{'='*60}")
            print(f"ARCHIVO: {pptx_file}")
            print(f"{'='*60}")
            print(f"Número total de slides: {len(slides)}\n")
            
            for i, slide_name in enumerate(slides, 1):
                print(f"\n--- SLIDE {i} ---")
                
                # Leer el contenido XML del slide
                slide_xml = zip_ref.read(slide_name)
                
                # Parsear XML
                root = ET.fromstring(slide_xml)
                
                # Buscar todos los elementos de texto
                texts = []
                for elem in root.iter():
                    if elem.tag.endswith('}t'):  # Elementos de texto
                        if elem.text:
                            texts.append(elem.text)
                
                if texts:
                    slide_content = '\n'.join(texts)
                    print(slide_content)
                    slide_data.append({
                        'slide_num': i,
                        'content': slide_content,
                        'text_count': len(texts)
                    })
                else:
                    print("(Sin texto o solo imágenes)")
                    slide_data.append({
                        'slide_num': i,
                        'content': '',
                        'text_count': 0
                    })
            
            return slide_data
                    
    except Exception as e:
        print(f"Error procesando {pptx_file}: {str(e)}")
        return []

# Archivos a analizar
archivos = [
    'S12_campo_magnetico.pptx',
    'S13_campo_magnetico_por_corriente.pptx',
    'S14_Inducción.pptx',
    'S15_Inductancia.pptx'
]

print("ANÁLISIS DE PRESENTACIONES - ELECTRICIDAD Y MAGNETISMO")
print("="*60)

for archivo in archivos:
    if os.path.exists(archivo):
        extract_text_from_pptx(archivo)
    else:
        print(f"\nArchivo no encontrado: {archivo}")

print("\n" + "="*60)
print("ANÁLISIS COMPLETADO")
print("="*60)

# Made with Bob
