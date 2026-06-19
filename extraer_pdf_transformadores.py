import sys
import os

try:
    # Intentar con pdfplumber primero
    import pdfplumber
    
    pdf_path = 'Transformadores.pdf'
    
    if not os.path.exists(pdf_path):
        print(f"Error: No se encuentra el archivo {pdf_path}")
        sys.exit(1)
    
    print("Extrayendo contenido de Transformadores.pdf...")
    print("="*80)
    
    with pdfplumber.open(pdf_path) as pdf:
        print(f"Número de páginas: {len(pdf.pages)}\n")
        
        full_text = ""
        for i, page in enumerate(pdf.pages, 1):
            print(f"\n{'='*80}")
            print(f"PÁGINA {i}")
            print('='*80)
            text = page.extract_text()
            if text:
                print(text)
                full_text += f"\n\n--- PÁGINA {i} ---\n\n{text}"
            else:
                print("(Página sin texto extraíble - posiblemente solo imágenes)")
    
    # Guardar en archivo
    with open('transformadores_extraido.txt', 'w', encoding='utf-8') as f:
        f.write("CONTENIDO EXTRAÍDO DE TRANSFORMADORES.PDF\n")
        f.write("="*80 + "\n")
        f.write(full_text)
    
    print("\n" + "="*80)
    print("Contenido guardado en: transformadores_extraido.txt")
    print("="*80)

except ImportError:
    print("pdfplumber no está instalado. Intentando instalar...")
    os.system("pip install pdfplumber")
    print("\nPor favor, ejecuta el script nuevamente.")

# Made with Bob
