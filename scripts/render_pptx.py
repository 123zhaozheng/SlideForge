#!/usr/bin/env python3
"""Render all slides of a .pptx back to high-res PNGs for visual QA self-check.

Zero pywin32 required on Windows: invokes native PowerShell PowerPoint COM (or WPS COM)
protected by a cross-process lock so parallel agents never collide,
falling back cleanly to LibreOffice + PyMuPDF on macOS/Linux.
"""

from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, List


@contextmanager
def _com_render_lock(timeout_s: float = 180.0) -> Iterator[None]:
    """Cross-process exclusive file lock so parallel agents serialize PowerPoint COM safely."""
    lock_path = Path(tempfile.gettempdir()) / "slideforge_com_render.lock"
    start = time.monotonic()
    fd = None
    while True:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_RDWR)
            os.write(fd, f"{os.getpid()}".encode("ascii"))
            break
        except FileExistsError:
            try:
                if lock_path.exists() and (time.time() - lock_path.stat().st_mtime) > 120:
                    lock_path.unlink(missing_ok=True)
                    continue
            except OSError:
                pass
            if (time.monotonic() - start) > timeout_s:
                raise RuntimeError("Timed out waiting for PowerPoint COM render lock")
            time.sleep(0.5)
    try:
        yield
    finally:
        if fd is not None:
            try:
                os.close(fd)
            except OSError:
                pass
        try:
            lock_path.unlink(missing_ok=True)
        except OSError:
            pass


def render_windows_com(pptx_path: Path, out_dir: Path, width: int = 1920, height: int = 1080) -> List[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    ps_script = f"""
$ErrorActionPreference = "Stop"
$pptx = '{str(pptx_path.resolve()).replace("'", "''")}'
$outDir = '{str(out_dir.resolve()).replace("'", "''")}'
$rendered = $false
foreach ($progId in @("PowerPoint.Application", "KWPP.Application")) {{
  $app = $null
  $pres = $null
  try {{
    $app = New-Object -ComObject $progId
    $pres = $app.Presentations.Open($pptx, $true, $false, $false)
    $count = $pres.Slides.Count
    for ($i = 1; $i -le $count; $i++) {{
      $name = "slide-" + $i.ToString("00") + ".png"
      $dst = Join-Path $outDir $name
      $pres.Slides.Item($i).Export($dst, "PNG", {width}, {height})
      Write-Output "EXPORTED:$dst"
    }}
    $pres.Close()
    if ($app.Presentations.Count -eq 0) {{ $app.Quit() }}
    $rendered = $true
    break
  }} catch {{
    if ($pres) {{ try {{ $pres.Close() }} catch {{}} }}
    if ($app) {{ try {{ if ($app.Presentations.Count -eq 0) {{ $app.Quit() }} }} catch {{}} }}
  }}
}}
if (-not $rendered) {{
  Write-Error "No PowerPoint or WPS COM backend available"
  exit 1
}}
"""
    with _com_render_lock():
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=180,
        )
        time.sleep(0.4)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip() or r.stdout.strip() or "COM render failed")
    paths: List[Path] = []
    for ln in r.stdout.splitlines():
        if ln.startswith("EXPORTED:"):
            p = Path(ln.split("EXPORTED:", 1)[1].strip())
            if p.is_file():
                paths.append(p)
    return paths


def render_libreoffice(pptx_path: Path, out_dir: Path) -> List[Path]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("Neither PowerPoint/WPS COM nor LibreOffice is available")
    out_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path)],
        check=True,
        capture_output=True,
        timeout=180,
    )
    pdf_path = out_dir / (pptx_path.stem + ".pdf")
    try:
        import fitz  # PyMuPDF
    except ImportError as e:
        raise RuntimeError("PyMuPDF (pymupdf) required for PDF->PNG fallback") from e
    doc = fitz.open(str(pdf_path))
    paths: List[Path] = []
    for idx, page in enumerate(doc, 1):
        pix = page.get_pixmap(dpi=180)
        dst = out_dir / f"slide-{idx:02d}.png"
        pix.save(str(dst))
        paths.append(dst)
    doc.close()
    pdf_path.unlink(missing_ok=True)
    return paths


def main() -> None:
    ap = argparse.ArgumentParser(description="Render PPTX slides to PNGs for visual QA")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--width", type=int, default=1920)
    ap.add_argument("--height", type=int, default=1080)
    args = ap.parse_args()

    pptx_path = Path(args.pptx).resolve()
    out_dir = Path(args.out_dir).resolve()
    if not pptx_path.is_file():
        raise SystemExit(f"PPTX not found: {pptx_path}")

    paths: List[Path] = []
    if platform.system() == "Windows":
        try:
            paths = render_windows_com(pptx_path, out_dir, args.width, args.height)
        except Exception as e:
            print(f"[warn] Windows COM unavailable ({e}), trying LibreOffice...")
            paths = render_libreoffice(pptx_path, out_dir)
    else:
        paths = render_libreoffice(pptx_path, out_dir)

    for p in paths:
        print(f"[ok] rendered -> {p}")
    print(f"[done] {len(paths)} slide(s) rendered")


if __name__ == "__main__":
    main()
