from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

from PIL import Image, ImageChops


def _slug(value: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9_-]+", "-", value.strip().lower()).strip("-")
    return value or "asset"


def _sheet_file(project: dict[str, Any]) -> str:
    sheet = project.get("assetSheet")
    if isinstance(sheet, dict) and isinstance(sheet.get("fileName"), str):
        return Path(sheet["fileName"]).name
    for asset in project.get("assets", []):
        if isinstance(asset, dict):
            for key in ("sourceSheet", "sheetFile", "assetSheet", "sourceFile"):
                if isinstance(asset.get(key), str):
                    return Path(asset[key]).name
    raise ValueError("assetSheet.fileName or assets[].sourceSheet is required")


def _box_from_asset(asset: dict[str, Any], width: int, height: int) -> tuple[int, int, int, int]:
    box = asset.get("cropBoxPct") or asset.get("cropBox") or asset.get("box")
    if isinstance(box, dict):
        x = float(box.get("x", 0))
        y = float(box.get("y", 0))
        w = float(box.get("w", box.get("width", 0)))
        h = float(box.get("h", box.get("height", 0)))
    elif isinstance(box, list) and len(box) >= 4:
        x, y, w, h = [float(v) for v in box[:4]]
    else:
        raise ValueError(f"Asset {asset.get('id') or asset.get('name')} is missing cropBoxPct")

    # Percent boxes are preferred. Accept either 0..1 or 0..100.
    if max(abs(x), abs(y), abs(w), abs(h)) <= 1.5:
        x, y, w, h = x * 100, y * 100, w * 100, h * 100
    if max(abs(x), abs(y), abs(w), abs(h)) <= 100:
        left = round(width * x / 100)
        top = round(height * y / 100)
        right = round(width * (x + w) / 100)
        bottom = round(height * (y + h) / 100)
    else:
        left = round(x)
        top = round(y)
        right = round(x + w)
        bottom = round(y + h)

    left = max(0, min(width - 1, left))
    top = max(0, min(height - 1, top))
    right = max(left + 1, min(width, right))
    bottom = max(top + 1, min(height, bottom))
    return left, top, right, bottom


def _remove_background(crop: Image.Image) -> Image.Image:
    rgba = crop.convert("RGBA")
    # Estimate solid sheet background from corners.
    w, h = rgba.size
    corners = [rgba.getpixel((0, 0)), rgba.getpixel((w - 1, 0)), rgba.getpixel((0, h - 1)), rgba.getpixel((w - 1, h - 1))]
    bg = tuple(round(sum(pixel[i] for pixel in corners) / len(corners)) for i in range(3))

    px = rgba.load()
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            dist = abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2])
            # Keep soft shadows partially, remove near-background pixels.
            if dist < 42:
                px[x, y] = (r, g, b, 0)
            elif dist < 82:
                px[x, y] = (r, g, b, min(a, 120))

    alpha = rgba.getchannel("A")
    bbox = alpha.getbbox()
    if bbox:
        rgba = rgba.crop(bbox)
    return rgba


def crop_assets(project_path: Path) -> dict[str, Any]:
    job_dir = project_path.parent
    project = json.loads(project_path.read_text(encoding="utf-8"))
    sheet_name = _sheet_file(project)
    sheet_path = job_dir / sheet_name
    if not sheet_path.exists():
        raise FileNotFoundError(f"Asset sheet image not found: {sheet_name}")

    sheet_img = Image.open(sheet_path).convert("RGB")
    width, height = sheet_img.size
    cutouts: list[str] = []

    for index, asset in enumerate(project.get("assets", []), start=1):
        if not isinstance(asset, dict):
            continue
        box = _box_from_asset(asset, width, height)
        crop = sheet_img.crop(box)
        cutout = _remove_background(crop)
        target_name = Path(str(asset.get("fileName") or f"{_slug(asset.get('id') or asset.get('name') or f'asset-{index}')}.png")).name
        if target_name == sheet_name:
            target_name = f"{_slug(asset.get('id') or asset.get('name') or f'asset-{index}')}.png"
        if not target_name.lower().endswith(".png"):
            target_name = f"{Path(target_name).stem}.png"
        cutout.save(job_dir / target_name)
        asset["fileName"] = target_name
        asset["sourceSheet"] = sheet_name
        asset["status"] = "완료"
        cutouts.append(target_name)

    qa = project.setdefault("qa", [])
    qa.append({
        "id": "qa-asset-sheet-crop",
        "label": "에셋 시트 분리 확인",
        "detail": f"에셋 시트 {sheet_name}에서 {len(cutouts)}개 PNG 에셋을 잘라냈음.",
        "passed": len(cutouts) > 0,
    })

    project_path.write_text(json.dumps(project, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"assetSheet": sheet_name, "cutouts": cutouts}


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python server/crop_asset_sheet.py <sitpo_project.json>", file=sys.stderr)
        raise SystemExit(2)
    result = crop_assets(Path(sys.argv[1]).resolve())
    print(json.dumps(result, ensure_ascii=False))
