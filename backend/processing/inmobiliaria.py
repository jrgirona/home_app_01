import os
import shutil
import glob
import re
import traceback
from datetime import datetime
import openpyxl
import pdfplumber
import pytesseract
from pdf2image import convert_from_path

def log_error(base_dir, message, receipt_text=""):
    log_path = os.path.join(base_dir, "error.log")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    clean_text = receipt_text.replace('\n', ' ').strip()
    with open(log_path, "a", encoding="utf-8") as f:
        # Table format: Date | Error Message | Notes (Text of receipt)
        f.write(f"{timestamp}\t{message}\t{clean_text}\n")

def perform_ocr(pdf_path):
    """Fallback OCR using pytesseract if the PDF is scanned."""
    try:
        images = convert_from_path(pdf_path)
        text = ""
        for img in images:
            text += pytesseract.image_to_string(img, lang='spa') + "\n"
        return text
    except Exception as e:
        return f"OCR Failed: {str(e)}"

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        pass
    
    # If no text found via pdfplumber (scanned document), use OCR fallback
    if not text.strip():
        text = perform_ocr(pdf_path)
        
    return text

def process_inmobiliaria(year, data_dir, results_dir):
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    target_filename = f"{year} Agua y luz.xlsx"
    target_path = os.path.join(results_dir, target_filename)
    
    # 1. Setup Excel template
    if not os.path.exists(target_path):
        ref_path = os.path.join(base_dir, "references", "2025 Agua y luz.xlsx")
        if not os.path.exists(ref_path):
            log_error(base_dir, f"Missing template file: {ref_path}")
            raise FileNotFoundError(f"Missing template file: {ref_path}")
        
        os.makedirs(results_dir, exist_ok=True)
        shutil.copy(ref_path, target_path)
        
        wb = openpyxl.load_workbook(target_path)
        
        if 'Luz' in wb.sheetnames:
            ws_luz = wb['Luz']
            # Clear data rows 2 to 13
            for row in range(2, 14):
                for col in range(1, 4):
                    ws_luz.cell(row=row, column=col).value = None
            
        if 'Agua' in wb.sheetnames:
            ws_agua = wb['Agua']
            # Clear data rows 2 to 7
            for row in range(2, 8):
                for col in range(1, 5):
                    ws_agua.cell(row=row, column=col).value = None
            
        wb.save(target_path)
        wb.close()
    
    wb = openpyxl.load_workbook(target_path)
    
    # --- PROCESS AGUA ---
    agua_pdfs = glob.glob(os.path.join(data_dir, "Comunidad Las Colinas Cuota*.pdf"))
    if 'Agua' in wb.sheetnames:
        ws_agua = wb['Agua']
        
        # Find next empty row between 2 and 7
        agua_next_row = 2
        while agua_next_row <= 7 and ws_agua.cell(row=agua_next_row, column=2).value:
            agua_next_row += 1
            
        # We need to keep track of newly added rows to calculate "Total desde el nuevo recibo"
        first_new_agua_row = agua_next_row
        
        for pdf_path in agua_pdfs:
            if agua_next_row > 7:
                break # Reached max capacity for the year
                
            text = extract_text_from_pdf(pdf_path)
            lines = text.split('\n')
            found_agua = False
            for line in lines:
                if 'Concepto:' in line and 'AGUA' in line:
                    found_agua = True
                    try:
                        recibo_match = re.search(r'\(lect\.\s*\d{2}/(\d{2})/(\d{4})\)', line)
                        recibo = f"{recibo_match.group(1)}-{recibo_match.group(2)}" if recibo_match else f"Unknown-{year}"
                        vol_match = re.search(r'\)\s*(\d+)-\d+', line)
                        vol = vol_match.group(1) if vol_match else "0"
                        importe_match = re.search(r':\s*(\d+[,.]\d+)$', line.strip())
                        if importe_match:
                            importe = float(importe_match.group(1).replace(',', '.'))
                        else:
                            importe = 0.0
                            
                        # Clear old 'Nuevo recibo'
                        for r in range(2, agua_next_row):
                            if ws_agua.cell(row=r, column=1).value == "Nuevo recibo":
                                ws_agua.cell(row=r, column=1).value = None
                                
                        ws_agua.cell(row=agua_next_row, column=1).value = "Nuevo recibo"
                        ws_agua.cell(row=agua_next_row, column=2).value = recibo
                        ws_agua.cell(row=agua_next_row, column=3).value = int(vol)
                        ws_agua.cell(row=agua_next_row, column=4).value = importe
                        
                        agua_next_row += 1
                    except Exception as e:
                        log_error(base_dir, f"Failed to parse AGUA data in {os.path.basename(pdf_path)}: {e}", text)
            
            if not found_agua:
                # Normal for odd months, just skip
                pass

        # Calculate Agua summaries (Rows 8, 9, 10, Label in Col 3, Value in Col 4)
        total_nuevo = 0.0
        if first_new_agua_row < agua_next_row:
            total_nuevo = sum(ws_agua.cell(row=r, column=4).value or 0.0 for r in range(first_new_agua_row, agua_next_row))
            
        gran_total = sum(ws_agua.cell(row=r, column=4).value or 0.0 for r in range(2, 8))
        
        last_receipts = [ws_agua.cell(row=r, column=4).value or 0.0 for r in range(2, 8) if ws_agua.cell(row=r, column=4).value]
        promedio = sum(last_receipts[-3:]) / len(last_receipts[-3:]) if last_receipts else 0.0
        
        ws_agua.cell(row=8, column=3).value = "Total desde el nuevo recibo"
        ws_agua.cell(row=8, column=4).value = total_nuevo
        ws_agua.cell(row=9, column=3).value = "Gran total"
        ws_agua.cell(row=9, column=4).value = gran_total
        ws_agua.cell(row=10, column=3).value = "Promedio (3 últimos meses)"
        ws_agua.cell(row=10, column=4).value = promedio


    # --- PROCESS LUZ ---
    luz_pdfs = glob.glob(os.path.join(data_dir, "Electricidad*.pdf"))
    if 'Luz' in wb.sheetnames:
        ws_luz = wb['Luz']
        
        # Find next empty row between 2 and 13
        luz_next_row = 2
        while luz_next_row <= 13 and ws_luz.cell(row=luz_next_row, column=2).value:
            luz_next_row += 1
            
        first_new_luz_row = luz_next_row
        
        for pdf_path in luz_pdfs:
            if luz_next_row > 13:
                break
                
            filename = os.path.basename(pdf_path)
            recibo_match = re.search(r'(\d{2}-\d{4})', filename)
            recibo = recibo_match.group(1) if recibo_match else f"Unknown-{year}"
            text = extract_text_from_pdf(pdf_path)
            
            importe = 0.0
            found_importe = False
            lines = text.split('\n')
            for line in lines:
                if 'IMPORTE TOTAL: EUROS' in line:
                    importe_match = re.search(r'\*{1,}\s*(\d+[,.]\d+)', line)
                    if importe_match:
                        importe = float(importe_match.group(1).replace(',', '.'))
                        found_importe = True
                        break
            if not found_importe:
                for line in lines:
                    if 'IMPORTE TOTAL' in line:
                        importe_match = re.search(r'(\d+[,.]\d+)$', line.strip())
                        if importe_match:
                            importe = float(importe_match.group(1).replace(',', '.'))
                            found_importe = True
                            break
                            
            if not found_importe:
                 log_error(base_dir, f"Failed to find IMPORTE TOTAL in {filename}", text)
                 
            # Clear old 'Nuevo recibo'
            for r in range(2, luz_next_row):
                if ws_luz.cell(row=r, column=1).value == "Nuevo recibo":
                    ws_luz.cell(row=r, column=1).value = None
                    
            ws_luz.cell(row=luz_next_row, column=1).value = "Nuevo recibo"
            ws_luz.cell(row=luz_next_row, column=2).value = recibo
            ws_luz.cell(row=luz_next_row, column=3).value = importe
            
            luz_next_row += 1

        # Calculate Luz summaries (Rows 14, 15, 16, Label in Col 2, Value in Col 3)
        total_nuevo = 0.0
        if first_new_luz_row < luz_next_row:
            total_nuevo = sum(ws_luz.cell(row=r, column=3).value or 0.0 for r in range(first_new_luz_row, luz_next_row))
            
        gran_total = sum(ws_luz.cell(row=r, column=3).value or 0.0 for r in range(2, 14))
        
        last_receipts = [ws_luz.cell(row=r, column=3).value or 0.0 for r in range(2, 14) if ws_luz.cell(row=r, column=3).value]
        promedio = sum(last_receipts[-3:]) / len(last_receipts[-3:]) if last_receipts else 0.0
        
        ws_luz.cell(row=14, column=2).value = "Total desde el nuevo recibo"
        ws_luz.cell(row=14, column=3).value = total_nuevo
        ws_luz.cell(row=15, column=2).value = "Gran total"
        ws_luz.cell(row=15, column=3).value = gran_total
        ws_luz.cell(row=16, column=2).value = "Promedio (3 últimos meses)"
        ws_luz.cell(row=16, column=3).value = promedio

    wb.save(target_path)
    wb.close()
    
    return target_path
