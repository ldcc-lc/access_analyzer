"""Lee reportes de operadores de Sintesis y conserva su salida original en Excel."""

import re
import os
from pathlib import Path


# Limpia el texto del reporte y extrae los campos de cada operador con sus roles.
def procesar_texto_reporte(texto):
    
    corte_final = re.search(r'(CAJERO SUPERVISOR O ADM\.|Totales\s*:|FIN DEL REPORTE)', texto, flags=re.IGNORECASE)
    if corte_final:
        texto = texto[:corte_final.start()]

    texto = texto.replace('|', ' ')
    texto = re.sub(r'\n+', ' ', texto)
    
    
    basura_previa = [
        r'SINTESIS S\.A\.', r'cnrep\d+', r'REPORTE DE O[PF]ERADORES EN LINEA',
        r'CTA\.\s*UNICA', r'CSTA\.\s*UNICA', r'Fecha:\s*[\d=/-]+', r'Hora\s*:\s*[\d/]*',
        r'Pag\s*:\s*\d*', r'Inst\.:\s*\d+', r'Agen\.:\s*\d+',
        r'\bCIU\.', r'\bENTIDAD\b', r'\bAG\b', r'\bCOD\b', r'NRO\.TELF\.', 
        r'\bNOMBRE\b', r'\bROL\b', r'\bEST\b', r'\b0000\b', r'\b00\b',
        r'DPTO:\s*\d+', r'CIUDAD:\s*\d+',
        r'-{3,}',                       
        r'\d{2,3}[=-]?/\d{2}/\d{0,4}',  
        r'\by\s+Pag\s*:\s*\d+',         
        r'\by\s+CTA\.\s*UNICA',        
        r'\by\s+(?=\d)',                
        r'_{3,}'                        
    ]
    
    for b in basura_previa:
        texto = re.sub(b, ' ', texto, flags=re.IGNORECASE)
        
    texto = re.sub(r'\s+', ' ', texto) 

    patron_registro = r'\b(\d{3})\s+(880|\d{1,2})\s+(880|\d{1,2})\s+(\d{1,6})\s+(.*?)(?=\s+\d{3}\s+(?:880\s+\d{1,2}|\d{1,2}\s+880)\s+\d{1,6}|\s*$)'
    
    registros = re.findall(patron_registro, texto)
    datos_procesados = []

    roles_posibles = [
        'PEPE CONS.HISTORI', 'CREDITOS BDP', 'CONTABILIDAD BDP', 'CONTABILIDAD BDF', 
        'ADM. NACIONAL', 'OPERADOR SFL', 'SUP. AGENCIA', 
        'CARTERA BDP', 'CREDITOS', 'CONS.HISTORI'
    ]

    for ciu, val1, val2, cod, resto in registros:
        entidad = '880'
        ag = val2 if val1 == '880' else val1

       
        cta_unica = ''
        
      
        cta_match = re.search(r'\b(\d{7,})\b', resto)
        if cta_match:
            cta_unica = cta_match.group(1)
            resto = resto.replace(cta_unica, '', 1)
        
      
        if not cta_unica:
            cta_match = re.search(r'\b(\d{6,}[A-Za-z]+\d*)\b', resto)
            if cta_match:
                cta_unica = cta_match.group(1)
                resto = resto.replace(cta_unica, '', 1)
        
       
        if not cta_unica:
            cta_match = re.search(r'\b((?=\w*[a-z])(?=\w*[A-Z])\w+|\d{3,}[A-Za-z]{1,}|[A-Za-z]{1,}\d+)\b', resto)
            if cta_match:
                cta_unica = cta_match.group(1)
                resto = resto.replace(cta_unica, '', 1)
        
        # 2. Extraer ROL (antes del ESTADO para no confundir V del nombre)
        rol = ''
        for r in roles_posibles:
            if r in resto:
                rol = r
                resto = resto.replace(r, '', 1)
                break
        
        # 3. Extraer NRO.TELF
        nro_telf_match = re.search(r'\b(\d{7,8})\b', resto)
        nro_telf = nro_telf_match.group(1) if nro_telf_match else ''
        if nro_telf:
            resto = resto.replace(nro_telf, '', 1)
                
        # 4. Limpiar fechas y datos adicionales
        resto = re.sub(r'\b\d{2,4}/\d{2}/\d{2,4}\b', '', resto) 
        resto = re.sub(r'\b\d{2}/\d{2}/\d{2}\b', '', resto)
        
        # 5. Limpiar DPTO y CIUDAD (con o sin espacio después de :)
        resto = re.sub(r'\s*DPTO\s*:\s*\d+', '', resto, flags=re.IGNORECASE)
        resto = re.sub(r'\s*CIUDAD\s*:\s*\d+', '', resto, flags=re.IGNORECASE)
        
        # 6. Extraer ESTADO (X o V) - Usar el ÚLTIMO match para evitar confundir con V del nombre
        est_matches = list(re.finditer(r'\b(X|V)\b', resto))
        est = ''
        if est_matches:
            est = est_matches[-1].group(1)  # Tomar el último
            resto = resto[:est_matches[-1].start()] + resto[est_matches[-1].end():]
        
        nombre = re.sub(r'\s+', ' ', resto).strip()
        
        datos_procesados.append({
            'CIU': ciu,
            'ENTIDAD': entidad,
            'AG': ag,
            'COD': cod,
            'NRO.TELF.': nro_telf,
            'NOMBRE': nombre,
            'ROL': rol,
            'CTA.UNICA': cta_unica,
            'EST': est
        })

    return datos_procesados


# Extrae texto de todas las paginas PDF y lo convierte en una tabla de operadores.
def procesar_sintesis_pdf(ruta_pdf):
    import pandas as pd
    import pdfplumber

    """Función para procesar PDFs de Síntesis y retornar DataFrame"""
    texto_bruto = ""
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for page in pdf.pages:
                texto_bruto += page.extract_text() or ""
    except FileNotFoundError:
        raise Exception(f"No se encontró el archivo '{ruta_pdf}'")
    except Exception as e:
        raise Exception(f"Error al leer el PDF: {e}")
    
    # Procesar el texto del PDF
    datos = procesar_texto_reporte(texto_bruto)
    df = pd.DataFrame(datos)
    df = df[df['NOMBRE'] != '']
    
    return df


# Procesa el PDF o TXT con las reglas originales y guarda los datos en Excel.
def procesar_archivo(ruta_archivo, ruta_excel_salida=None):
    import pandas as pd

    """
    Función genérica para procesar archivos PDF o TXT y convertir a Excel.
    
    Args:
        ruta_archivo (str): Ruta al archivo PDF o TXT
        ruta_excel_salida (str, optional): Ruta de salida del Excel. 
                                          Si no se especifica, se guarda en la misma carpeta con extensión .xlsx
    
    Returns:
        pd.DataFrame: DataFrame con los datos procesados
    """
    ruta_archivo = Path(ruta_archivo)
    
    
    if not ruta_archivo.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")
    
    extension = ruta_archivo.suffix.lower()
    texto_bruto = ""
    

    if extension == '.pdf':
        try:
            import pdfplumber
            with pdfplumber.open(str(ruta_archivo)) as pdf:
                for page in pdf.pages:
                    texto_bruto += page.extract_text() or ""
        except Exception as e:
            raise Exception(f"Error al leer el PDF: {e}")
    
    elif extension == '.txt':
        try:
            with open(ruta_archivo, 'r', encoding='utf-8') as f:
                texto_bruto = f.read()
        except UnicodeDecodeError:
            # Intentar con otra codificación si la primera falla
            try:
                with open(ruta_archivo, 'r', encoding='latin-1') as f:
                    texto_bruto = f.read()
            except Exception as e:
                raise Exception(f"Error al leer el TXT: {e}")
        except Exception as e:
            raise Exception(f"Error al leer el TXT: {e}")
    
    else:
        raise ValueError(f"Formato de archivo no soportado: {extension}. Use .pdf o .txt")
    
   
    datos = procesar_texto_reporte(texto_bruto)
    df = pd.DataFrame(datos)
    df = df[df['NOMBRE'] != '']
    
 
    if ruta_excel_salida is None:
        ruta_excel_salida = ruta_archivo.with_suffix('.xlsx')
    
    # Guardar a Excel
    ruta_excel_salida = Path(ruta_excel_salida)
    try:
        df.to_excel(str(ruta_excel_salida), index=False, sheet_name='Operadores')
        print(f"✓ Archivo Excel guardado en: {ruta_excel_salida}")
    except Exception as e:
        raise Exception(f"Error al guardar el Excel: {e}")
    
    return df


def convertir_pdf_sintesis_a_excel(ruta_pdf, ruta_excel_salida):
    """Aplica el procesamiento original de Síntesis y devuelve ruta y registros."""
    ruta_excel_salida = Path(ruta_excel_salida)
    datos = procesar_archivo(ruta_pdf, ruta_excel_salida)
    return ruta_excel_salida, len(datos)
