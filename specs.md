# MatildeMartin App Specifications

## 1. Overview
The MatildeMartin app is a local web-based tool designed to process financial data and generate Excel reports for two main categories: **IRPF** and **Inmobiliaria**.

## 2. Architecture & Tech Stack
* **Frontend Framework:** React (Next.js or Vite)
  * Provides a highly responsive, modern, and premium user interface.
* **Backend Framework:** Python (FastAPI or Flask)
  * Required due to the need for advanced dataframes and PDF parsing.
  * Handles secure access to the local file system (reading from `Data&Docs` and writing to `Results`).
* **Styling:** Vanilla CSS with a focus on rich aesthetics, sleek dark mode, glassmorphism, and micro-animations.
* **Data Processing:** Python libraries (`pandas` for dataframes, `pdfplumber`/`PyMuPDF` for PDF extraction, and `openpyxl` for Excel generation).

## 3. Directory Structure
* **Root:** `C:\Users\Ramon\Documents\Projects\OwnFinances\MatildeMartin`
* **Input Data:** `Data&Docs/<Year>/` - Contains the source files for the given year.
* **Outputs:** 
  * `Results/IRPF/<Year>/` - Destination for generated IRPF Excel reports.
  * `Results/Inmobiliaria/<Year>/` - Destination for generated Inmobiliaria Excel reports.

## 4. User Interface Features
* **Year Selector:** A dropdown menu allowing the user to select the financial year. The options will dynamically span from 2025 up to the current calendar year.
* **Report Type Selector:** Interactive buttons or a dropdown to select the desired output report type: **IRPF** or **Inmobiliaria**.
* **Generate Button:** A prominent call-to-action to trigger the processing of the files.
* **Status Indicators:** Real-time feedback showing processing status (Loading, Success, Error).

## 5. Workflows
1. The user opens the web app in their browser.
2. They select a `Year` (e.g., 2025, 2026).
3. They select the `Report Type` (IRPF or Inmobiliaria).
4. Upon clicking "Generate", the app's backend reads the relevant documents from `Data&Docs/<Selected Year>`.
5. The backend processes the data and generates a structured Excel file.
6. The resulting Excel file is saved locally to `Results/<Report Type>/`.
7. The UI updates to notify the user of the successful generation.

## 6. Business Logic Details

### IRPF Report Rules
1. **Source Data (Excel):** Read `202X_BBVA_movements.xlsx` from the `Data&Docs/<Year>` folder (where 202X is the selected year).
   - Load the table starting on row 5, column B into a **pandas DataFrame**.
   - Fields expected: `F.Valor`, `Fecha`, `Concepto`, `Movimiento`, `Importe`, `Divisa`, `Disponible`, `Divisa`, `Observaciones`.
2. **Filtering & Transformation:**
   - Filter rows where `Concepto` is either "Adeudo comunidad propietarios edf.las colinas" or "Adeudo geo alternativa sl".
   - Keep only columns: `F.Valor`, `Concepto`, `Importe`, and `Observaciones`.
   - Reverse the sign of `Importe` to make amounts positive.
3. **PDF Extraction (Real Estate Expenses):**
   - Read PDF: `FACTURA KN45-2025 GESTION DE ARRENDAMIENTO - MATILDE MARTIN 29.pdf` (Note: year should be dynamic if applicable).
   - Extract the date prior to "FACTURAR A" (format dd/mm/yyyy) and map to `F.Valor`.
   - Map "Gastos de arrendamiento" to `Concepto`.
   - Extract the figure with "TOTAL A PAGAR" and map to `Importe`.
   - Extract the text under "DESCRIPCIÓN" beginning with "Honorarios" and map to `Observaciones`.
4. **Final Assembly:**
   - Append the row extracted from the PDF to the DataFrame.
   - Add a totalizer row at the very bottom that sums all figures in the `Importe` column.
   - Export this final DataFrame to an Excel file in `Results/IRPF/`.

### Inmobiliaria Report Rules
1. **Target File:** Fill the Excel file `202X Agua y luz.xlsx` (where 202X is the selected year).
2. **Template Behavior:**
   - If the file does NOT exist in `Results/Inmobiliaria/`, copy `references/2025 Agua y luz.xlsx` as a template. Remove values in rows 2-13 (Luz tab) and 2-7 (Agua tab).
   - If it exists, append to the existing data.
3. **Agua Tab Logic:**
   - Columns: `Status`, `Recibo`, `Vol`, `Importe`.
   - Read PDFs in `Data&Docs/202X` starting with "Comunidad Las Colinas Cuota". Not all will contain "AGUA".
   - Search for a line starting with "Concepto:" and ending with "AGUA", for example: `AGUA (lect. 02/12/2025) 4352-4332:20m3 :38,8`.
   - Extract `Month-Year` (e.g., 12-2025) to `Recibo`.
   - Extract the trailing figure (e.g., 38,8) to `Importe` formatting it with a dot and 2 decimals (e.g., 38.00).
   - Extract the first figure (e.g., 4352) to `Vol`.
4. **Luz Tab Logic:**
   - Columns: `Status`, `Recibo`, `Importe`.
   - Read PDFs in `Data&Docs/202X` starting with "Electricidad".
   - Extract the month and year from the filename (e.g., `XX-202Y`) and map to `Recibo`.
   - Search for variations of "IMPORTE TOTAL: EUROS" (ignoring any missing whitespace caused by OCR) followed by asterisks, extract the amount (e.g., 51,98), and map to `Importe` (e.g., 51.98).
5. **Status Column Logic (Both Tabs):**
   - The newest receipt gets the string "Nuevo recibo" in the `Status` column.
   - If appending to an existing file, remove "Nuevo recibo" from the previous row and place it on the new appended row.
6. **Summary Calculations:**
   - *Note: Empty or corrupted cells are safely ignored (treated as 0.0) during calculations to prevent crashes.*
   - Luz (Rows 14-16) and Agua (Rows 8-10) display summary values:
     - `Total desde el nuevo recibo`: Sum of newly added receipts during this processing session.
     - `Gran total`: Sum of all receipts for the year.
     - `Promedio (3 últimos meses)`: Average of the last up to 3 non-empty receipts.
7. **OCR & Error Handling:**
   - Use Tesseract OCR (with Spanish language support) for PDFs that lack directly extractable text.
   - If any receipts cannot be read or fields are missing, write the issues to a dedicated error log file in a tabular format (Date/Time | Error Message | Dump of Receipt Text).
