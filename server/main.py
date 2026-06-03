from __future__ import annotations

import asyncio
import json
import mimetypes
import os
import shutil
import subprocess
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import BackgroundTasks, FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field


SERVER_DIR = Path(__file__).resolve().parent
REPO_ROOT = SERVER_DIR.parent
JOBS_DIR = SERVER_DIR / "jobs"
RESULT_JSON = "sitpo_project.json"
PPTX_FILENAME = "sitpo_project.pptx"


class JobRequest(BaseModel):
  grade: str = Field(..., min_length=1)
  subject: str = Field(..., min_length=1)
  unit: str = Field(..., min_length=1)
  topic: str = Field(..., min_length=1)
  style: str = Field(..., min_length=1)
  slideCount: int = Field(..., ge=1, le=30)


def now_iso() -> str:
  return datetime.now(timezone.utc).isoformat()


def job_dir(job_id: str) -> Path:
  return JOBS_DIR / job_id


def job_state_path(job_id: str) -> Path:
  return job_dir(job_id) / "job.json"


def safe_filename(filename: str) -> str:
  candidate = Path(filename).name
  if candidate != filename or not candidate:
    raise HTTPException(status_code=400, detail="Invalid filename")
  return candidate


def load_job(job_id: str) -> dict[str, Any]:
  path = job_state_path(job_id)
  if not path.exists():
    raise HTTPException(status_code=404, detail="Job not found")
  return json.loads(path.read_text(encoding="utf-8"))


def write_job(job: dict[str, Any]) -> None:
  path = job_state_path(job["id"])
  path.parent.mkdir(parents=True, exist_ok=True)
  tmp_path = path.with_suffix(".tmp")
  tmp_path.write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")
  tmp_path.replace(path)


def append_log(job: dict[str, Any], message: str) -> None:
  timestamp = datetime.now(timezone.utc).strftime("%H:%M:%S")
  job["logs"].append(f"[{timestamp}] {message}")
  job["updatedAt"] = now_iso()
  write_job(job)


def list_files(job_id: str) -> list[dict[str, Any]]:
  root = job_dir(job_id)
  files: list[dict[str, Any]] = []
  for path in sorted(root.iterdir() if root.exists() else []):
    if path.is_file() and path.name != "job.json":
      mime_type, _ = mimetypes.guess_type(path.name)
      files.append(
        {
          "filename": path.name,
          "url": f"/jobs/{job_id}/files/{path.name}",
          "sizeBytes": path.stat().st_size,
          "mimeType": mime_type or "application/octet-stream",
        }
      )
  return files


def fallback_project_shape(job: dict[str, Any]) -> dict[str, Any]:
  request = job["request"]
  return {
    "id": job["id"],
    "title": f"{request['grade']} {request['subject']} - {request['topic']}",
    "grade": request["grade"],
    "subject": request["subject"],
    "unit": request["unit"],
    "topic": request["topic"],
    "style": request["style"],
    "currentStep": "research",
    "createdAt": job["createdAt"],
    "slides": [],
    "assets": [],
    "diagrams": [],
    "qa": [],
  }


def build_codex_prompt(job: dict[str, Any]) -> str:
  request = job["request"]
  return f"""
You are generating a real SITPO elementary slide project. Work in the current directory only.

Input:
- grade: {request["grade"]}
- subject: {request["subject"]}
- unit: {request["unit"]}
- topic: {request["topic"]}
- style: {request["style"]}
- slideCount: {request["slideCount"]}

Required outputs:
1. Create `{RESULT_JSON}` as UTF-8 JSON matching this schema:
{{
  "id": "{job["id"]}",
  "title": string,
  "grade": string,
  "subject": string,
  "unit": string,
  "topic": string,
  "style": string,
  "currentStep": "download",
  "createdAt": string,
  "slides": [
    {{
      "slideNo": number,
      "title": string,
      "learningGoal": string,
      "mainMessage": string,
      "visibleText": string[],
      "studentActivity": string,
      "imagePlan": string,
      "diagramPlan": string,
      "teacherNote": string,
      "imageFile": optional string
    }}
  ],
  "assets": [
    {{
      "id": string,
      "name": string,
      "slides": number[],
      "kind": "배경" | "투명 PNG" | "아이콘 세트" | "캐릭터/오브젝트",
      "prompt": string,
      "status": "완료",
      "fileName": optional string
    }}
  ],
  "diagrams": [
    {{
      "id": string,
      "title": string,
      "type": "process_flow" | "comparison_table" | "cycle" | "relationship_map" | "classification",
      "slides": number[],
      "nodes": string[],
      "layout": string,
      "qaRule": string,
      "status": "완료"
    }}
  ],
  "qa": [{{"id": string, "label": string, "detail": string, "passed": boolean}}]
}}
2. Generate exactly {request["slideCount"]} custom slides for the user's requested grade, subject, unit, topic, and style. Do not copy or reuse any sample project content.
3. Use native image generation for at least one representative visual asset. Save the real generated image file in this directory as PNG/JPEG/WebP and reference the filename from `assets[].fileName` and at least one related slide `imageFile`.
4. If native image generation is unavailable, fail clearly: write `image_generation_unavailable.txt` explaining the unavailable native tool and do not create fake, copied, downloaded, SVG-only, base64-placeholder, or synthesized substitute images.
5. Do not download images from the web. Do not create text-only placeholders for visual assets.
6. Keep all visible text Korean, age-appropriate, concise, and classroom-ready.
"""


def validate_result(job: dict[str, Any], project: dict[str, Any]) -> None:
  request = job["request"]
  slides = project.get("slides")
  if not isinstance(slides, list) or len(slides) != request["slideCount"]:
    raise ValueError(f"Codex result must contain exactly {request['slideCount']} slides.")

  image_refs: set[str] = set()
  for slide in slides:
    if isinstance(slide, dict) and isinstance(slide.get("imageFile"), str):
      image_refs.add(Path(slide["imageFile"]).name)

  for asset in project.get("assets", []):
    if isinstance(asset, dict):
      for key in ("fileName", "filename", "path", "generatedImage"):
        if isinstance(asset.get(key), str):
          image_refs.add(Path(asset[key]).name)

  real_images = []
  for filename in image_refs:
    path = job_dir(job["id"]) / filename
    if path.exists() and path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"} and path.stat().st_size > 0:
      real_images.append(path)

  if not real_images:
    unavailable = job_dir(job["id"]) / "image_generation_unavailable.txt"
    if unavailable.exists():
      raise ValueError(unavailable.read_text(encoding="utf-8").strip())
    raise ValueError("Codex did not produce a real generated image file. Native image generation is required.")


def run_codex(job: dict[str, Any]) -> None:
  codex_bin = os.environ.get("SITPO_CODEX_BIN", "codex")
  if not shutil.which(codex_bin):
    raise RuntimeError(f"Codex CLI not found: {codex_bin}")

  prompt = build_codex_prompt(job)
  prompt_path = job_dir(job["id"]) / "codex_prompt.md"
  prompt_path.write_text(prompt, encoding="utf-8")

  command = [codex_bin, "exec", "--skip-git-repo-check", prompt]
  append_log(job, "Codex CLI 작업을 시작했습니다. ChatGPT OAuth 로그인이 필요할 수 있습니다.")
  completed = subprocess.run(
    command,
    cwd=job_dir(job["id"]),
    text=True,
    capture_output=True,
    timeout=int(os.environ.get("SITPO_CODEX_TIMEOUT_SECONDS", "1800")),
    check=False,
  )
  (job_dir(job["id"]) / "codex_stdout.log").write_text(completed.stdout, encoding="utf-8")
  (job_dir(job["id"]) / "codex_stderr.log").write_text(completed.stderr, encoding="utf-8")

  if completed.returncode != 0:
    raise RuntimeError(f"Codex CLI failed with exit code {completed.returncode}.")


def build_pptx(job: dict[str, Any]) -> None:
  script = SERVER_DIR / "build-pptx.mjs"
  input_path = job_dir(job["id"]) / RESULT_JSON
  output_path = job_dir(job["id"]) / PPTX_FILENAME
  completed = subprocess.run(
    ["node", str(script), str(input_path), str(output_path)],
    cwd=REPO_ROOT,
    text=True,
    capture_output=True,
    timeout=180,
    check=False,
  )
  (job_dir(job["id"]) / "build_pptx_stdout.log").write_text(completed.stdout, encoding="utf-8")
  (job_dir(job["id"]) / "build_pptx_stderr.log").write_text(completed.stderr, encoding="utf-8")
  if completed.returncode != 0:
    raise RuntimeError(f"PPTX build failed with exit code {completed.returncode}.")


def run_job(job_id: str) -> None:
  job = load_job(job_id)
  job["status"] = "running"
  job["updatedAt"] = now_iso()
  write_job(job)

  try:
    append_log(job, "요청을 저장했고 Codex 생성 작업을 준비합니다.")
    run_codex(job)
    append_log(job, "Codex 결과 JSON을 검증합니다.")

    result_path = job_dir(job_id) / RESULT_JSON
    if not result_path.exists():
      raise RuntimeError(f"Codex did not create {RESULT_JSON}.")

    project = json.loads(result_path.read_text(encoding="utf-8"))
    validate_result(job, project)
    append_log(job, "실제 생성 이미지 파일을 확인했습니다.")

    build_pptx(job)
    append_log(job, "서버 PPTX 파일을 생성했습니다.")

    job["result"] = project
    job["status"] = "succeeded"
    job["error"] = None
  except Exception as error:
    job["result"] = fallback_project_shape(job)
    job["status"] = "failed"
    job["error"] = str(error)
    append_log(job, f"작업 실패: {error}")
  finally:
    job["files"] = list_files(job_id)
    job["updatedAt"] = now_iso()
    write_job(job)


app = FastAPI(title="SITPO Job API")

allowed_origins = [
  "http://localhost:5173",
  "http://127.0.0.1:5173",
  "http://localhost:4173",
  "http://127.0.0.1:4173",
]
extra_origins = [origin.strip() for origin in os.environ.get("SITPO_CORS_ORIGINS", "").split(",") if origin.strip()]
app.add_middleware(
  CORSMiddleware,
  allow_origins=allowed_origins + extra_origins,
  allow_origin_regex=os.environ.get("SITPO_CORS_REGEX", r"https://.*\.vercel\.app"),
  allow_credentials=False,
  allow_methods=["*"],
  allow_headers=["*"],
)


@app.get("/health")
async def health() -> dict[str, str]:
  return {"status": "ok"}


@app.post("/api/jobs")
async def create_job(request: JobRequest, background_tasks: BackgroundTasks) -> dict[str, Any]:
  JOBS_DIR.mkdir(parents=True, exist_ok=True)
  job_id = f"sitpo-{uuid.uuid4().hex[:12]}"
  created_at = now_iso()
  job = {
    "id": job_id,
    "status": "queued",
    "request": request.model_dump() if hasattr(request, "model_dump") else request.dict(),
    "createdAt": created_at,
    "updatedAt": created_at,
    "logs": ["[대기] 작업이 생성되었습니다. Codex 작업은 수 분 이상 걸릴 수 있습니다."],
    "files": [],
    "result": None,
    "error": None,
  }
  write_job(job)
  background_tasks.add_task(run_job, job_id)
  return job


@app.get("/api/jobs/{job_id}")
async def get_job(job_id: str) -> dict[str, Any]:
  job = load_job(job_id)
  job["files"] = list_files(job_id)
  write_job(job)
  return job


@app.get("/api/jobs/{job_id}/files/{filename}")
async def get_job_file(job_id: str, filename: str) -> FileResponse:
  job = load_job(job_id)
  clean_filename = safe_filename(filename)
  path = job_dir(job["id"]) / clean_filename
  if not path.exists() or not path.is_file():
    raise HTTPException(status_code=404, detail="File not found")
  return FileResponse(path, filename=clean_filename)


if __name__ == "__main__":
  import uvicorn

  uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), reload=True)
