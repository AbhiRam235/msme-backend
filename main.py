# backend/main.py

import os
import uuid
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv
import uvicorn

# Internal module imports
from generation import generate_dpr_package

# Load environment variables (.env must include GEMINI_API_KEY or OPENAI_API_KEY)
load_dotenv()

# ---------------- APP CONFIG ----------------
app = FastAPI(
    title="AI DPR Architect - Backend",
    description="Backend service that generates Detailed Project Reports (DPRs) using AI and financial modeling.",
    version="1.0.0"
)

# Enable CORS (allow all origins for now; restrict in production)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],           # Replace with your frontend domain in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------- REQUEST MODEL ----------------
class ProjectRequest(BaseModel):
    title: str
    short_description: str
    location: str | None = None
    capacity: int | None = None
    currency: str = "INR"
    additional: dict | None = None


# ---------------- ROUTES ----------------
@app.get("/")
def root():
    """Health check / root endpoint."""
    return {"message": "✅ AI DPR Architect Backend is running!"}


@app.post("/generate")
def generate_project_dpr(req: ProjectRequest):
    """
    Main endpoint: Accepts project brief and returns generated file paths.
    It calls the DPR generator which produces DOCX, PDF, and Excel reports.
    """
    try:
        result = generate_dpr_package(req.dict())
        return {
            "status": "success",
            "message": "DPR generated successfully.",
            "data": result
        }
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"❌ Error generating DPR: {str(e)}"
        )


# ---------------- MAIN ENTRY ----------------
if __name__ == "__main__":
    # Host and port configuration (can be overridden via .env)
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", 8000))

    uvicorn.run(
        "main:app",
        host=host,
        port=port,
        reload=True
    )
