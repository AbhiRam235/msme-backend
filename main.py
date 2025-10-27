# backend/main.py
import os
import io
import uuid
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from generation import generate_dpr_package
from dotenv import load_dotenv

load_dotenv()  # expects OPENAI_API_KEY in .env if using OpenAI

app = FastAPI(title="AI DPR Architect - Backend")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # change in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class ProjectRequest(BaseModel):
    title: str
    short_description: str
    location: str | None = None
    capacity: int | None = None
    currency: str = "INR"
    additional: dict | None = None

@app.post("/generate")
def generate_project_dpr(req: ProjectRequest):
    """
    Main endpoint: Accepts project brief and returns links to generated files
    (or base64 content). Here we save files to disk and return filenames.
    """
    try:
        # generate files; returns dict with file paths
        result = generate_dpr_package(req.dict())
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
