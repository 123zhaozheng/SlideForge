#!/usr/bin/env python3
"""Seamlessly erase the sample title text on an enterprise baseboard image
so a 100% native editable PowerPoint title textbox can be overlaid cleanly.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import numpy as np
from PIL import Image


def clean_baseboard_title(
    src: Path,
    out: Path,
    box_norm: tuple[float, float, float, float] = (0.042, 0.030, 0.56, 0.110),
) -> Path:
    """Replace box_norm=(x0, y0, x1, y1) with per-column sampled background just above y0."""
    im = Image.open(src).convert("RGB")
    arr = np.array(im)
    h, w, _ = arr.shape

    x0 = max(0, int(w * box_norm[0]))
    y0 = max(1, int(h * box_norm[1]))
    x1 = min(w, int(w * box_norm[2]))
    y1 = min(h, int(h * box_norm[3]))

    # Sample strip just above the title box (handles both solid colors and horizontal gradients)
    samp_y0 = max(0, int(h * max(0.008, box_norm[1] - 0.018)))
    samp_y1 = max(samp_y0 + 1, y0 - 2)
    col_bg = np.median(arr[samp_y0:samp_y1, x0:x1], axis=0).astype(np.uint8)
    arr[y0:y1, x0:x1] = col_bg[None, :, :]

    out.parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(arr).save(out, "PNG")
    print(f"[ok] Clean baseboard saved → {out} (erased x:{box_norm[0]:.2f}-{box_norm[2]:.2f}, y:{box_norm[1]:.2f}-{box_norm[3]:.2f})")
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Erase sample top-left title from baseboard PNG")
    ap.add_argument("--input", required=True, help="Input baseboard PNG")
    ap.add_argument("--out", required=True, help="Output clean baseboard PNG")
    ap.add_argument(
        "--box",
        default="0.042,0.030,0.56,0.110",
        help="Normalized erase box x0,y0,x1,y1 (default fits standard top-left title above header rule)",
    )
    args = ap.parse_args()
    vals = tuple(float(v.strip()) for v in args.box.split(","))
    if len(vals) != 4:
        raise SystemExit("--box must have 4 comma-separated floats: x0,y0,x1,y1")
    clean_baseboard_title(Path(args.input), Path(args.out), vals)  # type: ignore[arg-type]


if __name__ == "__main__":
    main()
