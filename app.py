import os
import uuid
import shutil
import subprocess
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

TMP_DIR = "/tmp/kichikpdf"
os.makedirs(TMP_DIR, exist_ok=True)

@app.get("/")
def health():
    return {"status": "KichikPDF Server ishlayapti"}

@app.post("/convert")
async def convert_file(file: UploadFile = File(...), target: str = Form(...)):
    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    in_path = os.path.join(job_dir, file.filename)
    with open(in_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    base_name = os.path.splitext(file.filename)[0]

    try:
        # Office (DOCX, XLSX, PPTX) -> PDF
        if target == "pdf":
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", job_dir, in_path],
                check=True,
                timeout=60
            )
            out_path = os.path.join(job_dir, base_name + ".pdf")
            if not os.path.exists(out_path):
                raise HTTPException(status_code=500, detail="PDF yaratilmadi")
            return FileResponse(out_path, filename=base_name + ".pdf", media_type="application/pdf")

        # PDF -> DOCX (Word)
        elif target == "docx":
            out_path = os.path.join(job_dir, base_name + ".docx")
            cv = Converter(in_path)
            cv.convert(out_path)
            cv.close()
            if not os.path.exists(out_path):
                raise HTTPException(status_code=500, detail="DOCX yaratilmadi")
            return FileResponse(
                out_path,
                filename=base_name + ".docx",
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            raise HTTPException(status_code=400, detail="Noto'g'ri format tanlandi")

    except subprocess.TimeoutExpired:
        raise HTTPException(status_code=504, detail="Fayl juda katta, vaqt tugadi")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
