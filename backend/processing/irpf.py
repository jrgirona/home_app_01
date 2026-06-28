import os
import pandas as pd
import pdfplumber
import glob

def process_irpf(year, data_dir, results_dir):
    excel_pattern = os.path.join(data_dir, f"{year}_BBVA_movements.xlsx")
    excel_files = glob.glob(excel_pattern)
    if not excel_files:
        raise FileNotFoundError(f"Could not find {year}_BBVA_movements.xlsx in {data_dir}")
    
    excel_file = excel_files[0]
    
    # 1. Load the table starting on row 5 (header is row 4 zero-indexed), column B (which is index 1)
    # The actual layout might vary, so we load using skiprows=4.
    df = pd.read_excel(excel_file, skiprows=4)
    
    # 2. Filter & Transform
    target_conceptos = ["Adeudo comunidad propietarios edf.las colinas", "Adeudo geo alternativa sl"]
    df_filtered = df[df['Concepto'].isin(target_conceptos)].copy()
    
    keep_cols = ['F.Valor', 'Concepto', 'Importe', 'Observaciones']
    # Check if these columns exist, if not, adjust names slightly
    actual_cols = df_filtered.columns.tolist()
    for col in keep_cols:
        if col not in actual_cols:
            raise KeyError(f"Column '{col}' not found in Excel file. Found: {actual_cols}")
            
    df_filtered = df_filtered[keep_cols]
    
    # Reverse sign of Importe
    df_filtered['Importe'] = -df_filtered['Importe']
    
    # 3. PDF Extraction
    pdf_pattern = os.path.join(data_dir, f"FACTURA KN45-{year}*.pdf")
    pdf_files = glob.glob(pdf_pattern)
    
    pdf_date = ""
    pdf_importe = 0.0
    pdf_observaciones = ""
    
    if pdf_files:
        pdf_file = pdf_files[0]
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
                
            lines = text.split("\n")
            for i, line in enumerate(lines):
                if "FACTURAR A" in line:
                    # Attempt to find date before this
                    parts = line.split("FACTURAR A")
                    pdf_date = parts[0].strip()
                if "TOTAL A PAGAR" in line:
                    # Attempt to extract figure
                    words = line.split()
                    for word in reversed(words):
                        try:
                            # Replace comma with dot for float conversion if needed
                            clean_word = word.replace('€', '').replace('.', '').replace(',', '.').strip()
                            pdf_importe = float(clean_word)
                            break
                        except ValueError:
                            continue
                if "DESCRIPCIÓN" in line:
                    # Look for Honorarios
                    for j in range(i+1, min(i+10, len(lines))):
                        if "Honorarios" in lines[j]:
                            pdf_observaciones = lines[j].strip()
                            break
                            
    # 4. Final Assembly
    new_row = {
        'F.Valor': pdf_date,
        'Concepto': "Gastos de arrendamiento",
        'Importe': pdf_importe,
        'Observaciones': pdf_observaciones
    }
    
    df_filtered = pd.concat([df_filtered, pd.DataFrame([new_row])], ignore_index=True)
    
    # Totalizer row
    total_importe = df_filtered['Importe'].sum()
    total_row = {
        'F.Valor': 'TOTAL',
        'Concepto': '',
        'Importe': total_importe,
        'Observaciones': ''
    }
    df_filtered = pd.concat([df_filtered, pd.DataFrame([total_row])], ignore_index=True)
    
    output_file = os.path.join(results_dir, f"{year}_IRPF_Result.xlsx")
    df_filtered.to_excel(output_file, index=False)
    
    return output_file
