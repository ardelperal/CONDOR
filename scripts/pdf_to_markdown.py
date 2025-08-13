#!/usr/bin/env python3
"""
Script para convertir archivos PDF a formato Markdown
Utiliza PyMuPDF (fitz) para extraer texto de PDFs
"""

import fitz  # PyMuPDF
import os
import sys
import re
from pathlib import Path

def clean_text(text):
    """Limpia y formatea el texto extraído del PDF"""
    # Eliminar saltos de línea excesivos
    text = re.sub(r'\n\s*\n\s*\n', '\n\n', text)
    
    # Eliminar espacios al inicio y final de líneas
    lines = [line.strip() for line in text.split('\n')]
    
    # Reconstruir el texto
    cleaned_text = '\n'.join(lines)
    
    return cleaned_text

def extract_text_from_pdf(pdf_path):
    """Extrae texto de un archivo PDF"""
    try:
        doc = fitz.open(pdf_path)
        text = ""
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            page_text = page.get_text()
            text += f"\n\n<!-- Página {page_num + 1} -->\n\n"
            text += page_text
        
        doc.close()
        return clean_text(text)
    
    except Exception as e:
        print(f"Error al procesar {pdf_path}: {str(e)}")
        return None

def pdf_to_markdown(pdf_path, output_path=None):
    """Convierte un PDF a Markdown"""
    if not os.path.exists(pdf_path):
        print(f"Error: El archivo {pdf_path} no existe")
        return False
    
    # Generar nombre de salida si no se proporciona
    if output_path is None:
        pdf_name = Path(pdf_path).stem
        output_path = f"{pdf_name}.md"
    
    print(f"Convirtiendo {pdf_path} a {output_path}...")
    
    # Extraer texto del PDF
    text = extract_text_from_pdf(pdf_path)
    
    if text is None:
        return False
    
    # Crear encabezado Markdown
    pdf_name = Path(pdf_path).stem
    markdown_content = f"# {pdf_name}\n\n"
    markdown_content += f"*Convertido automáticamente desde PDF*\n\n"
    markdown_content += "---\n\n"
    markdown_content += text
    
    # Guardar archivo Markdown
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        print(f"✓ Conversión completada: {output_path}")
        return True
    
    except Exception as e:
        print(f"Error al guardar {output_path}: {str(e)}")
        return False

def main():
    """Función principal"""
    if len(sys.argv) < 2:
        print("Uso: python pdf_to_markdown.py <archivo.pdf> [salida.md]")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    success = pdf_to_markdown(pdf_path, output_path)
    
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()