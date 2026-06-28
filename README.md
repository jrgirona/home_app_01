# MatildeMartin Web App

This is a local financial processing tool designed to read local PDFs and Excel files and generate highly formatted "IRPF" and "Inmobiliaria" Excel reports. It uses a modern React frontend and a Python FastAPI backend for advanced data extraction (OCR + DataFrames).

## Prerequisites
1. **Python 3.9+** installed on your system.
2. **Node.js & npm** installed on your system.
3. **Tesseract OCR** installed on your machine (with the Spanish language pack) and available in your system PATH.

## How to Open the App

To run the application locally, you need to start both the Python backend server and the React frontend development server.

### 1. Start the Backend
Open a terminal (PowerShell or Command Prompt) and run the following commands:
```powershell
# Navigate to the backend folder
cd backend

# (Optional but recommended) Create and activate a virtual environment
python -m venv venv
.\venv\Scripts\activate

# Install the required Python dependencies
pip install -r requirements.txt

# Start the FastAPI server
uvicorn main:app --reload
```
*The backend will now be running on `http://127.0.0.1:8000`.*

### 2. Start the Frontend
Open a **new, separate** terminal window and run the following commands:
```powershell
# Navigate to the frontend folder
cd frontend

# Install the React dependencies (only needed the very first time)
npm install

# Start the Vite development server
npm run dev
```

### 3. Open in Browser
Once both servers are running, open your web browser and go to:
**`http://localhost:5173`**

You will see the MatildeMartin interface where you can select the year and the type of report you wish to generate!
