#!/usr/bin/env python3
"""Auto-crop & Alpha-extract RGBA Transparent Components from Pure-White Sprite Sheets.

Features:
  - Adaptive row/column valley detection for any grid (1x2, 2x2, 2x3, 2x4, 3x3, 3x4)
  - Two precision alpha extraction modes:
      * solid     : exterior BFS flood-fill (protects pure-white interior icons 100%!)
      * soft_fade : white-background anti-premultiply alpha recovery (for soft fading arrows/glows)
  - Connected-component outlier speck removal (eliminates stray 1-3px dust inflating bounding boxes)
  - Optional derived tint variants (e.g. "tint": "white" -> *_white.png for dark/Hero cards)
  - Automatic checkerboard alpha preview sheet (*_alpha_preview.png) for instant visual QA
"""

from __future__ import annotations

import argparse
import json
from collections import deque
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
from PIL import Image, ImageDraw


def _find_splits(proj: np.ndarray, total_len: int, parts: int) -> List[int]:
    if parts <= 1:
        return [0, total_len]
    splits = [0]
    for k in range(1, parts):
        est = int(total_len * k / parts)
        win = max(12, int(total_len * 0.085))
        lo = max(splits[-1] + 8, est - win)
        hi = min(total_len - 8, est + win)
        if lo >= hi:
            splits.append(est)
        else:
            # Prefer exact zero-run center if a flat white valley exists
            sub = proj[lo:hi]
            zeros = np.where(sub == 0)[0]
            if len(zeros) > 0:
                splits.append(lo + int(np.median(zeros)))
            else:
                splits.append(lo + int(np.argmin(sub)))
    splits.append(total_len)
    return splits


def _remove_isolated_specks(alpha: np.ndarray, min_area: int = 24) -> np.ndarray:
    """Zero out tiny disconnected alpha specks that are far smaller than real graphic elements."""
    mask = alpha > 16
    h, w = mask.shape
    visited = np.zeros((h, w), dtype=bool)
    cleaned = alpha.copy()

    # Fast scan over foreground coordinates only
    ys, xs = np.where(mask)
    if len(ys) == 0:
        return cleaned

    comps: List[Tuple[int, List[Tuple[int, int]]]] = []
    for y0, x0 in zip(ys.tolist(), xs.tolist()):
        if visited[y0, x0]:
            continue
        q = deque([(y0, x0)])
        visited[y0, x0] = True
        pts = [(y0, x0)]
        while q:
            cy, cx = q.popleft()
            for ny, nx in ((cy - 1, cx), (cy + 1, cx), (cy, cx - 1), (cy, cx + 1)):
                if 0 <= ny < h and 0 <= nx < w and mask[ny, nx] and not visited[ny, nx]:
                    visited[ny, nx] = True
                    q.append((ny, nx))
                    pts.append((ny, nx))
        comps.append((len(pts), pts))

    if not comps:
        return cleaned
    max_area = max(a for a, _ in comps)
    # Keep components >= min_area AND at least 0.3% of largest component (protects multi-part icons!)
    thresh = max(min_area, int(max_area * 0.0025))
    for area, pts in comps:
        if area < thresh:
            for py, px in pts:
                cleaned[py, px] = 0
    return cleaned


def extract_cell_rgba(cell_rgba: np.ndarray, cut_mode: str = "solid", white_tol: int = 12, pad: int = 4) -> Image.Image:
    h, w, _ = cell_rgba.shape
    rgb = cell_rgba[:, :, :3].astype(np.int16)
    dist_white = np.max(255 - rgb, axis=2)
    bg_seed = dist_white <= white_tol

    visited = np.zeros((h, w), dtype=bool)
    q: deque[Tuple[int, int]] = deque()
    for x in range(w):
        for y in (0, h - 1):
            if bg_seed[y, x] and not visited[y, x]:
                visited[y, x] = True
                q.append((y, x))
    for y in range(h):
        for x in (0, w - 1):
            if bg_seed[y, x] and not visited[y, x]:
                visited[y, x] = True
                q.append((y, x))

    while q:
        cy, cx = q.popleft()
        for ny, nx in ((cy - 1, cx), (cy + 1, cx), (cy, cx - 1), (cy, cx + 1)):
            if 0 <= ny < h and 0 <= nx < w and not visited[ny, nx] and bg_seed[ny, nx]:
                visited[ny, nx] = True
                q.append((ny, nx))

    out = cell_rgba.copy()
    out[visited, 3] = 0

    if cut_mode == "soft_fade":
        # Un-premultiply exterior white-fading pixels: C = a*F + (1-a)*255 => F = (C - 255*(1-a))/a
        fade_thresh = 92.0
        fade_zone = (~visited) & (dist_white < fade_thresh)
        out_a = out[:, :, 3].astype(np.float32) / 255.0
        out_a[fade_zone] = np.clip(dist_white[fade_zone].astype(np.float32) / fade_thresh, 0.0, 1.0)
        safe_a = np.maximum(out_a, 0.05)
        for ch in range(3):
            c = out[:, :, ch].astype(np.float32)
            f = np.clip((c - 255.0 * (1.0 - out_a)) / safe_a, 0, 255)
            out[fade_zone, ch] = f[fade_zone].astype(np.uint8)
        out[fade_zone, 3] = np.clip(out_a[fade_zone] * 255.0, 0, 255).astype(np.uint8)
    else:
        # 1-pixel exterior anti-alias defringe against white
        pad_v = np.pad(visited, 1, constant_values=True)
        near_bg = (~visited) & (
            pad_v[:-2, 1:-1] | pad_v[2:, 1:-1] | pad_v[1:-1, :-2] | pad_v[1:-1, 2:]
        )
        edge = near_bg & (dist_white < 60)
        out[edge, 3] = np.clip(dist_white[edge].astype(np.float32) / 60.0 * 255.0, 0, 255).astype(np.uint8)

    out[:, :, 3] = _remove_isolated_specks(out[:, :, 3])

    ys, xs = np.where(out[:, :, 3] > 10)
    if len(ys) == 0:
        return Image.fromarray(out)
    y0, y1 = max(0, int(ys.min()) - pad), min(h, int(ys.max()) + pad + 1)
    x0, x1 = max(0, int(xs.min()) - pad), min(w, int(xs.max()) + pad + 1)
    return Image.fromarray(out[y0:y1, x0:x1])


def make_tinted_variant(im: Image.Image, tint: str) -> Image.Image:
    arr = np.array(im.convert("RGBA"))
    t = tint.strip().lower()
    if t == "white":
        r, g, b = 255, 255, 255
    elif t.startswith("#") and len(t) == 7:
        r, g, b = int(t[1:3], 16), int(t[3:5], 16), int(t[5:7], 16)
    else:
        return im
    arr[:, :, 0] = r
    arr[:, :, 1] = g
    arr[:, :, 2] = b
    return Image.fromarray(arr)


def build_alpha_preview(items: List[Tuple[str, Image.Image]], rows: int, cols: int, out_path: Path) -> None:
    cell_w, cell_h, pad = 340, 340, 28
    pw = cols * cell_w + pad * 2
    ph = rows * cell_h + pad * 2
    sheet = Image.new("RGB", (pw, ph), (242, 243, 247))
    draw = ImageDraw.Draw(sheet)
    sq = 16
    for y in range(0, ph, sq):
        for x in range(0, pw, sq):
            if ((x // sq) + (y // sq)) % 2 == 0:
                draw.rectangle([x, y, x + sq - 1, y + sq - 1], fill=(224, 226, 232))

    for i, (name, im) in enumerate(items):
        r, c = divmod(i, cols)
        bx, by = pad + c * cell_w, pad + r * cell_h
        bw, bh = cell_w - 28, cell_h - 46
        scale = min(bw / max(1, im.width), bh / max(1, im.height))
        nw, nh = max(1, int(im.width * scale)), max(1, int(im.height * scale))
        im_r = im.resize((nw, nh), Image.LANCZOS)
        sheet.paste(im_r, (bx + (cell_w - nw) // 2, by + (bh - nh) // 2), im_r)
        # Subtle label bar at bottom of cell
        draw.rectangle([bx + 8, by + cell_h - 32, bx + cell_w - 8, by + cell_h - 8], fill=(30, 33, 40))
        draw.text((bx + 16, by + cell_h - 26), f"{name} ({im.width}x{im.height})", fill=(255, 255, 255))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out_path, "PNG")


def extract_batch(
    sheet_png: Path,
    batch: Dict[str, Any],
    comp_dir: Path,
    preview_out: Path,
) -> List[Dict[str, Any]]:
    im = Image.open(sheet_png).convert("RGBA")
    arr = np.array(im)
    H, W, _ = arr.shape
    rows, cols = int(batch["rows"]), int(batch["cols"])
    items = list(batch["items"])

    nonwhite = arr[:, :, :3].min(axis=2) < 245
    row_splits = _find_splits(nonwhite.sum(axis=1), H, rows)

    records: List[Dict[str, Any]] = []
    preview_items: List[Tuple[str, Image.Image]] = []
    comp_dir.mkdir(parents=True, exist_ok=True)

    for idx, it in enumerate(items):
        r_i, c_i = divmod(idx, cols)
        if r_i >= rows:
            break
        ry0, ry1 = row_splits[r_i], row_splits[r_i + 1]
        col_splits = _find_splits(nonwhite[ry0:ry1, :].sum(axis=0), W, cols)
        cx0, cx1 = col_splits[c_i], col_splits[c_i + 1]

        cell = arr[ry0:ry1, cx0:cx1].copy()
        name = str(it.get("name") or f"{batch['id']}_{idx+1:02d}")
        cut_mode = str(it.get("cut_mode") or "solid").lower()
        out_im = extract_cell_rgba(cell, cut_mode=cut_mode)
        out_file = comp_dir / f"{name}.png"
        out_im.save(out_file, "PNG")
        preview_items.append((name, out_im))

        rec: Dict[str, Any] = {
            "name": name,
            "batch": batch["id"],
            "cut_mode": cut_mode,
            "file": str(out_file),
            "size": [out_im.width, out_im.height],
        }
        tint = str(it.get("tint") or "").strip()
        if tint:
            tinted = make_tinted_variant(out_im, tint)
            suffix = "white" if tint.lower() == "white" else tint.lstrip("#")
            t_file = comp_dir / f"{name}_{suffix}.png"
            tinted.save(t_file, "PNG")
            rec["tinted_file"] = str(t_file)
            print(f"  [tint] {t_file.name}")

        records.append(rec)
        print(f"[ok] {out_file.name} ({out_im.width}x{out_im.height}, mode={cut_mode})")

    build_alpha_preview(preview_items, rows, cols, preview_out)
    print(f"[preview] {preview_out}")
    return records


def main() -> None:
    ap = argparse.ArgumentParser(description="Auto-crop & alpha-extract RGBA components from sprite sheets")
    ap.add_argument("--session", required=True)
    ap.add_argument("--manifest", default="", help="Manifest JSON (defaults to sheets/manifest_resolved.json)")
    ap.add_argument("--only", default="", help="Comma-separated batch IDs")
    args = ap.parse_args()

    session = Path(args.session).resolve()
    sheets_dir = session / "sheets"
    man_path = Path(args.manifest).resolve() if args.manifest else sheets_dir / "manifest_resolved.json"
    if not man_path.is_file():
        raise SystemExit(f"Manifest not found: {man_path}")

    manifest = json.loads(man_path.read_text(encoding="utf-8"))
    batches = manifest.get("batches") or []
    if not batches:
        raise SystemExit("Manifest has no resolved 'batches'. Run generate_components.py first.")

    only_set = {x.strip() for x in args.only.split(",") if x.strip()} if args.only else None
    comp_dir = session / "components"
    all_records: List[Dict[str, Any]] = []

    for b in batches:
        bid = b["id"]
        if only_set and bid not in only_set:
            continue
        sheet_png = sheets_dir / f"{bid}_raw.png"
        if not sheet_png.is_file():
            raise SystemExit(f"Missing sprite sheet: {sheet_png}")
        preview_out = sheets_dir / f"{bid}_alpha_preview.png"
        all_records.extend(extract_batch(sheet_png, b, comp_dir, preview_out))

    idx_path = comp_dir / "components_index.json"
    idx_path.write_text(json.dumps({"components": all_records}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[done] Extracted {len(all_records)} RGBA components → {comp_dir}")


if __name__ == "__main__":
    main()
