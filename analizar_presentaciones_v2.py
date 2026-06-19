import zipfile
import xml.etree.ElementTree as ET
import os
import json

def extract_text_from_pptx(pptx_file):
    """Extrae todo el texto de un archivo PowerPoint"""
    slide_data = []
    try:
        with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
            # Obtener lista de slides
            slides = [name for name in zip_ref.namelist() 
                     if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
            slides.sort()
            
            for i, slide_name in enumerate(slides, 1):
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
                
                slide_data.append({
                    'slide_num': i,
                    'texts': texts,
                    'text_count': len(texts)
                })
            
            return {
                'filename': pptx_file,
                'total_slides': len(slides),
                'slides': slide_data
            }
                    
    except Exception as e:
        return {
            'filename': pptx_file,
            'error': str(e),
            'total_slides': 0,
            'slides': []
        }

# Archivos a analizar
archivos = [
    'S12_campo_magnetico.pptx',
    'S13_campo_magnetico_por_corriente.pptx',
    'S14_Inducción.pptx',
    'S15_Inductancia.pptx'
]

resultados = []

for archivo in archivos:
    if os.path.exists(archivo):
        resultado = extract_text_from_pptx(archivo)
        resultados.append(resultado)

# Guardar resultados en JSON
with open('analisis_presentaciones.json', 'w', encoding='utf-8') as f:
    json.dump(resultados, f, ensure_ascii=False, indent=2)

# Crear reporte en texto
with open('analisis_presentaciones.txt', 'w', encoding='utf-8') as f:
    f.write("ANÁLISIS DE PRESENTACIONES - ELECTRICIDAD Y MAGNETISMO\n")
    f.write("="*80 + "\n\n")
    
    for resultado in resultados:
        f.write(f"\n{'='*80}\n")
        f.write(f"ARCHIVO: {resultado['filename']}\n")
        f.write(f"{'='*80}\n")
        f.write(f"Número total de slides: {resultado['total_slides']}\n\n")
        
        if 'error' in resultado:
            f.write(f"ERROR: {resultado['error']}\n")
            continue
        
        for slide in resultado['slides']:
            f.write(f"\n--- SLIDE {slide['slide_num']} ---\n")
            f.write(f"Elementos de texto: {slide['text_count']}\n\n")
            
            if slide['texts']:
                for text in slide['texts']:
                    f.write(f"{text}\n")
            else:
                f.write("(Sin texto o solo imágenes)\n")
    
    f.write(f"\n{'='*80}\n")
    f.write("ANÁLISIS COMPLETADO\n")
    f.write(f"{'='*80}\n")

print("Análisis completado. Revisa los archivos:")
print("- analisis_presentaciones.json")
print("- analisis_presentaciones.txt")

# Made with Bob
