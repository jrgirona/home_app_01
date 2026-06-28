from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os
import traceback

from processing.irpf import process_irpf
from processing.inmobiliaria import process_inmobiliaria

app = FastAPI()

# Allow frontend to communicate with backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class GenerateRequest(BaseModel):
    year: int
    report_type: str

@app.post("/api/generate")
async def generate_report(request: GenerateRequest):
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, "Data&Docs", str(request.year))
    results_dir = os.path.join(base_dir, "Results", request.report_type, str(request.year))
    
    if not os.path.exists(data_dir):
        raise HTTPException(status_code=400, detail=f"Data directory {data_dir} does not exist.")
    
    os.makedirs(results_dir, exist_ok=True)
    
    try:
        if request.report_type == "IRPF":
            result_file = process_irpf(str(request.year), data_dir, results_dir)
            return {"status": "success", "message": f"IRPF report generated successfully at {result_file}"}
        elif request.report_type == "Inmobiliaria":
            result_file = process_inmobiliaria(str(request.year), data_dir, results_dir)
            return {"status": "success", "message": f"Inmobiliaria report generated successfully at {result_file}"}
        else:
            raise HTTPException(status_code=400, detail="Invalid report type.")
    except Exception as e:
        error_msg = traceback.format_exc()
        # Log to error.log
        with open(os.path.join(base_dir, "error.log"), "a", encoding="utf-8") as f:
            f.write(f"Error generating {request.report_type} for {request.year}:\n{error_msg}\n")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
