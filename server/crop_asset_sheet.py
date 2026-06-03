from __future__ import annotations

import json
import math
import re
import sys
from pathlib import Path
from typing import Any

from PIL import Image


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


def _sheet_text(project: dict[str, Any]) -> str:
    sheet = project.get("assetSheet")
    chunks: list[str] = []
    if isinstance(sheet, dict):
        chunks.extend(str(sheet.get(key, "")) for key in ("layout", "prompt"))
    chunks.extend(str(project.get(key, "")) for key in ("title", "topic", "style"))
    return "\n".join(chunks)


def _infer_grid(project: dict[str, Any], asset_count: int) -> tuple[int, int, str]:
    text = _sheet_text(project)
    patterns = [
        # Korean: 4열 x 3행, 4 열 3 행
        (r"(\d+)\s*열\s*(?:x|×|by|,|\s)+\s*(\d+)\s*행", "ko_cols_rows"),
        # English: 4 columns by 3 rows / 4 cols x 3 rows
        (r"(\d+)\s*(?:columns?|cols?)\s*(?:x|×|by|,|\s)+\s*(\d+)\s*(?:rows?)", "en_cols_rows"),
        # English: 3 rows by 4 columns
        (r"(\d+)\s*(?:rows?)\s*(?:x|×|by|,|\s)+\s*(\d+)\s*(?:columns?|cols?)", "en_rows_cols"),
        # 3 by 3 grid
        (r"(\d+)\s*(?:by|x|×)\s*(\d+)\s*(?:grid|격자)", "grid_pair"),
    ]
    for pattern, kind in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue
        a, b = int(match.group(1)), int(match.group(2))
        if kind == "en_rows_cols":
            rows, cols = a, b
        else:
            cols, rows = a, b
        if cols > 0 and rows > 0 and cols * rows >= asset_count:
            return cols, rows, kind

    # Fallback: prefer a wide grid for slide asset sheets.
    cols = max(1, math.ceil(math.sqrt(asset_count)))
    rows = max(1, math.ceil(asset_count / cols))
    return cols, rows, "auto_sqrt"




def _count_items_in_row_text(text: str) -> int:
    text = re.sub(r"\([^)]*\)", "", text)
    parts = [part.strip() for part in re.split(r",|、|，", text) if part.strip()]
    return len(parts)


def _infer_row_counts(project: dict[str, Any], rows: int, cols: int, asset_count: int) -> tuple[list[int], str]:
    text = _sheet_text(project)
    counts: list[int] = []

    # Korean layout often says: "1행: A, B. 2행: C, D."
    for row_index in range(1, rows + 1):
        match = re.search(rf"{row_index}\s*행\s*[:：]\s*([^\.\n]+)", text)
        if match:
            counts.append(_count_items_in_row_text(match.group(1)))

    if len(counts) == rows and sum(counts) == asset_count and all(c > 0 for c in counts):
        return counts, "row_text_ko"

    # English layout often says: "Top row: A, B. Middle row: C, D. Bottom row: E."
    labels = ["top", "middle", "bottom", "fourth", "fifth"]
    counts = []
    for label in labels[:rows]:
        match = re.search(rf"{label}\s+row\s*[:：]\s*([^\.\n]+)", text, flags=re.IGNORECASE)
        if match:
            counts.append(_count_items_in_row_text(match.group(1)))
    if len(counts) == rows and sum(counts) == asset_count and all(c > 0 for c in counts):
        return counts, "row_text_en"

    # Fallback: fill each row left-to-right up to the global column count.
    remaining = asset_count
    counts = []
    for _ in range(rows):
        count = min(cols, remaining)
        if count > 0:
            counts.append(count)
            remaining -= count
    while len(counts) < rows:
        counts.append(0)
    return counts, "grid_fill"


def _row_cell_box(
    asset_index: int,
    row_counts: list[int],
    rows: int,
    width: int,
    height: int,
) -> tuple[int, int, int, int]:
    zero = asset_index - 1
    acc = 0
    row = 0
    col = 0
    row_count = row_counts[0] if row_counts else 1
    for row_i, count in enumerate(row_counts):
        if zero < acc + count:
            row = row_i
            col = zero - acc
            row_count = max(1, count)
            break
        acc += count
    top = round(height * row / rows)
    bottom = round(height * (row + 1) / rows)
    left = round(width * col / row_count)
    right = round(width * (col + 1) / row_count)
    # Keep only a small gutter. Do not shave much from the bottom: many generated
    # cutouts sit visually on the lower cell boundary, and bottom shaving makes
    # feet, notes, and shape shadows look clipped in the final PPTX.
    pad_x = max(2, round((right - left) * 0.02))
    pad_top = max(2, round((bottom - top) * 0.018))
    pad_bottom = max(0, round((bottom - top) * 0.004))
    return _clamp_box((left + pad_x, top + pad_top, right - pad_x, bottom - pad_bottom), width, height)

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

    return _clamp_box((left, top, right, bottom), width, height)


def _clamp_box(box: tuple[int, int, int, int], width: int, height: int) -> tuple[int, int, int, int]:
    left, top, right, bottom = box
    left = max(0, min(width - 1, left))
    top = max(0, min(height - 1, top))
    right = max(left + 1, min(width, right))
    bottom = max(top + 1, min(height, bottom))
    return left, top, right, bottom


def _background_rgb(img: Image.Image) -> tuple[int, int, int]:
    rgb = img.convert("RGB")
    w, h = rgb.size
    sample_points = [
        (0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1),
        (w // 2, 0), (w // 2, h - 1), (0, h // 2), (w - 1, h // 2),
    ]
    pixels = [rgb.getpixel(point) for point in sample_points]
    return tuple(round(sum(pixel[i] for pixel in pixels) / len(pixels)) for i in range(3))


def _cell_box(index: int, cols: int, rows: int, width: int, height: int) -> tuple[int, int, int, int]:
    zero = index - 1
    col = zero % cols
    row = zero // cols
    left = round(width * col / cols)
    right = round(width * (col + 1) / cols)
    top = round(height * row / rows)
    bottom = round(height * (row + 1) / rows)
    # Keep a little gutter out of each cell to avoid neighboring assets.
    pad_x = max(2, round((right - left) * 0.035))
    pad_y = max(2, round((bottom - top) * 0.035))
    return _clamp_box((left + pad_x, top + pad_y, right - pad_x, bottom - pad_y), width, height)


def _content_box_in_cell(
    img: Image.Image,
    cell: tuple[int, int, int, int],
    bg: tuple[int, int, int],
) -> tuple[int, int, int, int]:
    rgb = img.convert("RGB")
    left, top, right, bottom = cell
    px = rgb.load()
    xs: list[int] = []
    ys: list[int] = []
    # Two-threshold pass: include colored/ink edges, not the neutral sheet background.
    threshold = 48
    for y in range(top, bottom):
        for x in range(left, right):
            r, g, b = px[x, y]
            dist = abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2])
            chroma = max(r, g, b) - min(r, g, b)
            if dist > threshold or chroma > 32:
                xs.append(x)
                ys.append(y)

    cell_w = right - left
    cell_h = bottom - top
    min_area = cell_w * cell_h * 0.006
    if len(xs) < min_area:
        # If detection fails for a pale object, use the safe cell crop rather than Codex coordinates.
        return cell

    content = (min(xs), min(ys), max(xs) + 1, max(ys) + 1)
    pad_x = max(10, round(cell_w * 0.055))
    pad_top = max(8, round(cell_h * 0.05))
    pad_bottom = max(14, round(cell_h * 0.085))
    # Expand only inside the assigned row/cell. Crossing the cell boundary can pull in
    # thin shadows or object fragments from the neighboring row, which was the source
    # of the “weird crop” issue. Bottom gets extra room because PPTX fitting and
    # transparentization make lower edges look clipped when the crop is too tight.
    return (
        max(left, content[0] - pad_x),
        max(top, content[1] - pad_top),
        min(right, content[2] + pad_x),
        min(bottom, content[3] + pad_bottom),
    )




def _remove_thin_stray_components(rgba: Image.Image) -> Image.Image:
    w, h = rgba.size
    alpha = rgba.getchannel("A")
    mask = alpha.load()
    visited = bytearray(w * h)
    px = rgba.load()

    def idx(x: int, y: int) -> int:
        return y * w + x

    for sy in range(h):
        for sx in range(w):
            start = idx(sx, sy)
            if visited[start] or mask[sx, sy] == 0:
                continue
            stack = [(sx, sy)]
            visited[start] = 1
            points: list[tuple[int, int]] = []
            min_x = max_x = sx
            min_y = max_y = sy
            while stack:
                x, y = stack.pop()
                points.append((x, y))
                min_x = min(min_x, x); max_x = max(max_x, x)
                min_y = min(min_y, y); max_y = max(max_y, y)
                for nx in (x - 1, x, x + 1):
                    for ny in (y - 1, y, y + 1):
                        if nx < 0 or ny < 0 or nx >= w or ny >= h:
                            continue
                        ni = idx(nx, ny)
                        if visited[ni] or mask[nx, ny] == 0:
                            continue
                        visited[ni] = 1
                        stack.append((nx, ny))
            comp_w = max_x - min_x + 1
            comp_h = max_y - min_y + 1
            area = len(points)
            # Remove only thin, isolated horizontal leftovers. This avoids deleting real
            # small details like dotted triangle strokes or mirror center lines.
            is_thin_horizontal = comp_h <= 24 and comp_w >= max(34, comp_h * 4)
            is_top_stray = min_y < h * 0.14 and comp_h <= 30 and comp_w >= 42 and area < 6500
            is_tiny_dash = area < 260 and comp_h <= 8 and comp_w >= 18
            if is_thin_horizontal or is_top_stray or is_tiny_dash:
                for x, y in points:
                    px[x, y] = (px[x, y][0], px[x, y][1], px[x, y][2], 0)
    return rgba

def _add_transparent_padding(rgba: Image.Image) -> Image.Image:
    w, h = rgba.size
    # Preserve a visible safety moat around cutouts. Without this, alpha bbox is
    # exactly flush to the image edge and PowerPoint scaling/cropping can make
    # lower strokes, shadows, and feet look chopped off.
    pad_x = max(14, round(w * 0.06))
    pad_top = max(12, round(h * 0.045))
    pad_bottom = max(20, round(h * 0.10))
    canvas = Image.new("RGBA", (w + pad_x * 2, h + pad_top + pad_bottom), (255, 255, 255, 0))
    canvas.alpha_composite(rgba, (pad_x, pad_top))
    return canvas


def _remove_background(crop: Image.Image, bg: tuple[int, int, int] | None = None) -> Image.Image:
    rgba = crop.convert("RGBA")
    w, h = rgba.size
    if bg is None:
        corners = [rgba.getpixel((0, 0)), rgba.getpixel((w - 1, 0)), rgba.getpixel((0, h - 1)), rgba.getpixel((w - 1, h - 1))]
        bg = tuple(round(sum(pixel[i] for pixel in corners) / len(corners)) for i in range(3))

    px = rgba.load()
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            dist = abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2])
            chroma = max(r, g, b) - min(r, g, b)
            if dist < 38 and chroma < 30:
                px[x, y] = (r, g, b, 0)
            elif dist < 72 and chroma < 35:
                px[x, y] = (r, g, b, min(a, 115))

    rgba = _remove_thin_stray_components(rgba)
    alpha = rgba.getchannel("A")
    bbox = alpha.getbbox()
    if bbox:
        rgba = _add_transparent_padding(rgba.crop(bbox))
    return rgba


def _target_name(asset: dict[str, Any], index: int, sheet_name: str) -> str:
    target_name = Path(str(asset.get("fileName") or f"{_slug(asset.get('id') or asset.get('name') or f'asset-{index}')}.png")).name
    if target_name == sheet_name:
        target_name = f"{_slug(asset.get('id') or asset.get('name') or f'asset-{index}')}.png"
    if not target_name.lower().endswith(".png"):
        target_name = f"{Path(target_name).stem}.png"
    return target_name


def crop_assets(project_path: Path) -> dict[str, Any]:
    job_dir = project_path.parent
    project = json.loads(project_path.read_text(encoding="utf-8"))
    assets = [asset for asset in project.get("assets", []) if isinstance(asset, dict)]
    sheet_name = _sheet_file(project)
    sheet_path = job_dir / sheet_name
    if not sheet_path.exists():
        raise FileNotFoundError(f"Asset sheet image not found: {sheet_name}")

    sheet_img = Image.open(sheet_path).convert("RGB")
    width, height = sheet_img.size
    bg = _background_rgb(sheet_img)
    cols, rows, grid_source = _infer_grid(project, len(assets))
    row_counts, row_source = _infer_row_counts(project, rows, cols, len(assets))
    cutouts: list[str] = []
    boxes: list[dict[str, Any]] = []

    for index, asset in enumerate(assets, start=1):
        if index <= sum(row_counts):
            cell = _row_cell_box(index, row_counts, rows, width, height)
            box = _content_box_in_cell(sheet_img, cell, bg)
            crop_mode = "row_grid_content"
        else:
            box = _box_from_asset(asset, width, height)
            crop_mode = "fallback_codex_box"

        crop = sheet_img.crop(box)
        cutout = _remove_background(crop, bg)
        target_name = _target_name(asset, index, sheet_name)
        cutout.save(job_dir / target_name)
        asset["fileName"] = target_name
        asset["sourceSheet"] = sheet_name
        asset["status"] = "완료"
        asset["cropMode"] = crop_mode
        asset["actualCropBoxPx"] = {"left": box[0], "top": box[1], "right": box[2], "bottom": box[3]}
        cutouts.append(target_name)
        boxes.append({"id": asset.get("id"), "fileName": target_name, "box": asset["actualCropBoxPx"], "mode": crop_mode})

    qa = project.setdefault("qa", [])
    qa = [item for item in qa if not (isinstance(item, dict) and item.get("id") == "qa-asset-sheet-crop")]
    qa.append({
        "id": "qa-asset-sheet-crop",
        "label": "에셋 시트 분리 확인",
        "detail": f"{cols}x{rows} 격자({grid_source}, rows={row_counts}, {row_source}) 기준으로 {len(cutouts)}개 PNG 에셋을 잘라냈음.",
        "passed": len(cutouts) == len(assets) and len(cutouts) > 0,
    })
    project["qa"] = qa

    project_path.write_text(json.dumps(project, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"assetSheet": sheet_name, "grid": {"cols": cols, "rows": rows, "source": grid_source, "rowCounts": row_counts, "rowSource": row_source}, "cutouts": cutouts, "boxes": boxes}


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python server/crop_asset_sheet.py <sitpo_project.json>", file=sys.stderr)
        raise SystemExit(2)
    result = crop_assets(Path(sys.argv[1]).resolve())
    print(json.dumps(result, ensure_ascii=False))
