"""
pptx-autofill-app — PPTX 템플릿 자동 채우기 웹앱 (v2: 검토+검증)
"""
import asyncio
import os
import json
import uuid
import shutil
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

from fastapi import FastAPI, Request, Form, BackgroundTasks, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import Optional

from app.pptx_engine import (
    unpack, repack, analyze_template, analyze_template_summary,
    apply_replacements, get_markitdown_text,
)
from app.writer import generate_content_map

app = FastAPI(title="PPTX AutoFill", version="2.0")
templates = Jinja2Templates(directory="app/templates")
app.mount("/static", StaticFiles(directory="app/static"), name="static")

jobs: dict = {}

BASE_DIR = Path(__file__).resolve().parent.parent
SAMPLE_DIR = BASE_DIR / "sample_templates"
UPLOAD_DIR = BASE_DIR / "uploads"
WORKSPACE = BASE_DIR / "workspace"


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/api/templates")
async def list_templates():
    result = []
    for folder, source in [(SAMPLE_DIR, "sample"), (UPLOAD_DIR, "upload")]:
        if not folder.exists():
            continue
        for f in sorted(folder.glob("*.pptx")):
            try:
                info = analyze_template(str(f))
                result.append({
                    "name": f.stem,
                    "filename": f.name,
                    "source": source,
                    "slide_count": info["slide_count"],
                    "shapes": sum(len(s["shapes"]) for s in info["slides"]),
                    "tables": sum(len(s["tables"]) for s in info["slides"]),
                })
            except Exception:
                result.append({
                    "name": f.stem, "filename": f.name, "source": source,
                    "slide_count": "?", "shapes": "?", "tables": "?",
                })
    return result


@app.post("/api/upload-template")
async def upload_template(file: UploadFile = File(...)):
    UPLOAD_DIR.mkdir(exist_ok=True)
    safe_name = file.filename.replace(" ", "_")
    dest = UPLOAD_DIR / safe_name
    with open(dest, "wb") as f:
        content = await file.read()
        f.write(content)
    return {"status": "ok", "filename": safe_name}


@app.delete("/api/templates/{source}/{filename}")
async def delete_template(source: str, filename: str):
    if source != "upload":
        return JSONResponse({"error": "샘플 템플릿은 삭제할 수 없습니다"}, 400)
    path = UPLOAD_DIR / filename
    if path.exists():
        path.unlink()
    return {"status": "ok"}


def extract_file_text(file_path: str) -> str:
    """업로드된 참고 파일에서 텍스트 추출 (markitdown 활용)"""
    try:
        from markitdown import MarkItDown
        md = MarkItDown()
        result = md.convert(file_path)
        return result.text_content[:3000]  # 토큰 절약 (3000자 상한)
    except Exception:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()[:5000]
        except Exception:
            return ""


@app.post("/api/start")
async def start_generation(
    background_tasks: BackgroundTasks,
    template_source: str = Form(...),
    template_filename: str = Form(...),
    topic: str = Form(...),
    extra_info: str = Form(""),
    ref_files: list[UploadFile] = File([]),
):
    job_id = str(uuid.uuid4())[:8]

    if template_source == "sample":
        template_path = SAMPLE_DIR / template_filename
    else:
        template_path = UPLOAD_DIR / template_filename

    if not template_path.exists():
        return JSONResponse({"error": "템플릿 파일을 찾을 수 없습니다"}, 400)

    safe_topic = "".join(c for c in topic if c.isalnum() or c in "_ -").strip().replace(" ", "_")[:30]
    project_dir = WORKSPACE / f"{safe_topic}_{datetime.now().strftime('%Y%m%d_%H%M')}"
    project_dir.mkdir(parents=True, exist_ok=True)

    work_template = project_dir / "template.pptx"
    shutil.copy2(template_path, work_template)

    # 참고 파일 저장 + 텍스트 추출
    ref_texts = []
    for ref in ref_files:
        if ref and ref.filename and ref.size > 0:
            ext = os.path.splitext(ref.filename)[1] or ".txt"
            ref_path = os.path.join(str(project_dir), f"ref_{ref.filename}")
            with open(ref_path, "wb") as f:
                content = await ref.read()
                f.write(content)
            text = extract_file_text(ref_path)
            if text.strip():
                ref_texts.append(f"[파일: {ref.filename}]\n{text}")

    ref_context = "\n\n".join(ref_texts) if ref_texts else ""

    jobs[job_id] = {
        "status": "running",
        "phase": "준비 중...",
        "progress": 0,
        "project_dir": str(project_dir),
        "template_path": str(work_template),
        "topic": topic,
        "extra_info": extra_info,
        "ref_context": ref_context,
        "content_map": None,
        "analysis": None,
        "original_text": None,
        "output_file": None,
        "error": None,
    }

    background_tasks.add_task(run_pipeline, job_id)
    return {"job_id": job_id}


async def run_pipeline(job_id: str):
    """전체 파이프라인: 분석 → AI 생성 → XML 적용 → 검증"""
    job = jobs[job_id]
    template_path = job["template_path"]
    project_dir = job["project_dir"]

    try:
        # Step 1: 템플릿 분석
        job["phase"] = "템플릿 구조 분석 중..."
        job["progress"] = 10
        await asyncio.sleep(0.1)

        analysis = analyze_template(template_path)
        job["analysis"] = analysis
        job["original_text"] = get_markitdown_text(template_path)
        job["progress"] = 20

        # Step 2: AI 콘텐츠 생성
        job["phase"] = "AI가 콘텐츠를 생성하고 있습니다..."
        job["progress"] = 25
        await asyncio.sleep(0.1)

        # 참고 파일 내용이 있으면 extra_info에 합산
        full_extra = job["extra_info"]
        if job.get("ref_context"):
            full_extra += f"\n\n## 참고 자료 (업로드 파일 내용)\n{job['ref_context']}"

        content_map = await generate_content_map(
            analysis, job["original_text"], job["topic"], full_extra
        )
        job["content_map"] = content_map
        job["progress"] = 60

        # Step 3: XML 교체
        job["phase"] = "템플릿에 콘텐츠를 적용하고 있습니다..."
        job["progress"] = 65
        await asyncio.sleep(0.1)

        unpacked_dir = os.path.join(project_dir, "unpacked")
        unpack(template_path, unpacked_dir)

        int_map = {int(k): v for k, v in content_map.items()}
        apply_replacements(unpacked_dir, int_map)
        job["progress"] = 85

        # Step 4: 재패킹
        job["phase"] = "PPTX 파일을 생성하고 있습니다..."
        job["progress"] = 90
        await asyncio.sleep(0.1)

        safe_topic = "".join(c for c in job["topic"] if c.isalnum() or c in "_.가-힣 -")[:30]
        output_path = os.path.join(project_dir, f"{safe_topic}_완성.pptx")
        repack(unpacked_dir, template_path, output_path)
        job["output_file"] = output_path

        # Step 5: 검증
        job["phase"] = "결과를 검증하고 있습니다..."
        job["progress"] = 95
        job["result_text"] = get_markitdown_text(output_path)

        job["phase"] = "완료!"
        job["progress"] = 100
        job["status"] = "done"

    except Exception as e:
        job["status"] = "error"
        job["phase"] = f"오류 발생: {str(e)}"
        job["error"] = str(e)


@app.get("/api/status/{job_id}")
async def get_status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return JSONResponse({"error": "작업을 찾을 수 없습니다"}, 404)
    return {
        "status": job["status"],
        "phase": job["phase"],
        "progress": job["progress"],
        "error": job.get("error"),
    }



@app.get("/api/verify/{job_id}")
async def verify(job_id: str):
    """검증: 원본↔결과 텍스트 비교"""
    job = jobs.get(job_id)
    if not job:
        return JSONResponse({"error": "작업을 찾을 수 없습니다"}, 404)
    return {
        "original": job.get("original_text", ""),
        "result": job.get("result_text", ""),
        "slide_count_match": job["analysis"]["slide_count"] if job.get("analysis") else 0,
    }


@app.get("/api/download/{job_id}")
async def download(job_id: str):
    job = jobs.get(job_id)
    if not job or not job.get("output_file"):
        return JSONResponse({"error": "파일을 찾을 수 없습니다"}, 404)
    return FileResponse(
        job["output_file"],
        filename=os.path.basename(job["output_file"]),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
