#!/usr/bin/env python3
"""Multi-Batch Component Sprite Sheet Generator for SlideForge.

Key capabilities:
  - Supports explicit multi-batch manifests OR automatic chunking of large component lists
    (caps each call at <= 6-8 components so every icon/card stays sharp and well-separated).
  - Enforces 100% pure flat white (#FFFFFF), zero text, zero gridlines, zero checkerboard.
  - Locks cross-batch visual consistency via shared style_ref + palette_lock + optional chain_batch1.
  - Built-in transient retry (429/502/503/504/timeout), incremental skip, and concurrency control.
"""

from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
from urllib.parse import urlparse

import requests

SKILL_ROOT = Path(__file__).resolve().parent.parent
MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
TRANSIENT_MARKERS = (
    "429", "502", "503", "504", "524",
    "timeout", "timed out", "connection aborted", "remotedisconnected",
    "temporarily_unavailable", "service unavailable", "internal_error",
    "server_error", "upstream", "gateway",
)


def load_env() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        load_dotenv = None
    if load_dotenv:
        for p in (
            SKILL_ROOT / ".env",
            Path(os.environ.get("SLIDEFORGE_ENV", "")),
            Path(os.environ.get("IMAGE2_PPT_ENV", "")),
        ):
            if p and str(p) and p.is_file():
                load_dotenv(p, override=False)


def guess_mime(path: Path) -> str:
    return mimetypes.guess_type(str(path))[0] or "image/png"


def is_transient_error(msg: str) -> bool:
    low = (msg or "").lower()
    return any(m in low for m in TRANSIENT_MARKERS)


class ImageClient:
    def __init__(
        self,
        api_key: str,
        base_url: str,
        model: str = "gpt-image-1",
        quality: str = "high",
        size: str = "3840x2160",
        resolution: str = "4k",
        aspect_ratio: str = "16:9",
        timeout: int = 600,
        trust_env: bool = False,
    ):
        self.api_key = api_key
        base = base_url.rstrip("/")
        if not base.endswith("/v1"):
            base += "/v1"
        self.base_url = base
        self.model = model
        self.quality = quality
        self.size = size
        self.resolution = resolution
        self.aspect_ratio = aspect_ratio
        self.timeout = timeout
        self.session = requests.Session()
        self.session.trust_env = trust_env

    def _download(self, url: str) -> bytes:
        parsed = urlparse(url)
        if parsed.scheme not in ("http", "https"):
            raise ValueError(f"Unsupported URL scheme: {url[:80]}")
        r = self.session.get(url, stream=True, timeout=self.timeout)
        r.raise_for_status()
        chunks, total = [], 0
        for chunk in r.iter_content(chunk_size=16384):
            total += len(chunk)
            if total > MAX_DOWNLOAD_BYTES:
                raise ValueError("Image download exceeded 50MB limit")
            chunks.append(chunk)
        return b"".join(chunks)

    def _extract_bytes(self, payload: Dict[str, Any]) -> bytes:
        items = payload.get("data") or []
        if not items:
            raise RuntimeError(f"Empty image response: {json.dumps(payload)[:400]}")
        it = items[0]
        if it.get("b64_json"):
            return base64.b64decode(it["b64_json"])
        if it.get("url"):
            return self._download(it["url"])
        raise RuntimeError(f"Unrecognized image payload: {str(it)[:300]}")

    def generate(self, prompt: str, out: Path) -> Path:
        r = self.session.post(
            f"{self.base_url}/images/generations",
            headers={"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"},
            json={
                "model": self.model,
                "prompt": prompt,
                "n": 1,
                "size": self.size,
                "quality": self.quality,
                "resolution": self.resolution,
                "aspect_ratio": self.aspect_ratio,
            },
            timeout=(15, self.timeout),
        )
        if r.status_code >= 400:
            raise RuntimeError(f"generations {r.status_code}: {r.text[:600]}")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(self._extract_bytes(r.json()))
        return out

    def edit(self, prompt: str, images: Sequence[Path], out: Path) -> Path:
        if not images:
            return self.generate(prompt, out)
        files: List[Tuple[str, Tuple[str, Any, str]]] = []
        field = "image" if len(images) == 1 else "image[]"
        try:
            for p in images:
                files.append((field, (p.name, open(p, "rb"), guess_mime(p))))
            r = self.session.post(
                f"{self.base_url}/images/edits",
                headers={"Authorization": f"Bearer {self.api_key}"},
                data={
                    "model": self.model,
                    "prompt": prompt,
                    "n": "1",
                    "quality": self.quality,
                    "size": self.size,
                    "resolution": self.resolution,
                    "aspect_ratio": self.aspect_ratio,
                    "input_fidelity": "high",
                    "output_format": "png",
                },
                files=files,
                timeout=(15, self.timeout),
            )
        finally:
            for _, tup in files:
                try:
                    tup[1].close()
                except Exception:
                    pass
        if r.status_code >= 400:
            raise RuntimeError(f"edits {r.status_code}: {r.text[:600]}")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(self._extract_bytes(r.json()))
        return out


def _env_first(*names: str, default: str = "") -> str:
    for n in names:
        v = os.environ.get(n, "").strip()
        if v:
            return v
    return default


def resolve_client(args: Optional[argparse.Namespace] = None) -> ImageClient:
    load_env()
    key = _env_first("OPENAI_API_KEY", "SLIDEFORGE_API_KEY")
    if not key:
        raise SystemExit(
            "OPENAI_API_KEY is not set. Copy .env.example to .env and configure your API key."
        )
    base = _env_first("OPENAI_BASE_URL", "SLIDEFORGE_BASE_URL", default="https://api.openai.com/v1")
    model = (getattr(args, "model", None) or _env_first("SLIDEFORGE_MODEL", "IMAGE2_PPT_MODEL", default="gpt-image-1")).strip()
    quality = (getattr(args, "quality", None) or _env_first("SLIDEFORGE_QUALITY", "IMAGE2_PPT_QUALITY", default="high")).strip()
    size = (getattr(args, "size", None) or _env_first("SLIDEFORGE_SIZE", "IMAGE2_PPT_SIZE", default="3840x2160")).strip()
    resolution = (getattr(args, "resolution", None) or _env_first("SLIDEFORGE_RESOLUTION", "IMAGE2_PPT_RESOLUTION", default="4k")).strip()
    aspect = (getattr(args, "aspect", None) or _env_first("SLIDEFORGE_ASPECT", "IMAGE2_PPT_ASPECT", default="16:9")).strip()
    timeout = int(getattr(args, "timeout", 600) or 600)
    trust_env = _env_first("SLIDEFORGE_TRUST_ENV", "IMAGE2_PPT_TRUST_ENV", default="false").lower() in ("1", "true", "yes")
    return ImageClient(key, base, model, quality, size, resolution, aspect, timeout, trust_env)


def choose_grid(count: int) -> Tuple[int, int]:
    """Choose the cleanest landscape grid (rows, cols) for <= 8 components."""
    if count <= 2:
        return (1, max(1, count))
    if count <= 4:
        return (2, 2)
    if count <= 6:
        return (2, 3)
    return (2, 4)


def normalize_manifest(raw: Dict[str, Any], max_per_batch: int = 8) -> Dict[str, Any]:
    """Normalize either explicit 'batches' or flat 'components' list into explicit batches."""
    max_per_batch = max(2, min(8, int(raw.get("max_per_batch") or max_per_batch)))
    out = dict(raw)
    if raw.get("batches"):
        norm_batches = []
        for idx, b in enumerate(raw["batches"], 1):
            items = list(b.get("items") or [])
            if not items:
                continue
            r = int(b.get("rows") or 0)
            c = int(b.get("cols") or 0)
            if not (r and c):
                r, c = choose_grid(len(items))
            norm_batches.append({
                "id": str(b.get("id") or f"batch_{idx:02d}"),
                "rows": r,
                "cols": c,
                "style_note": str(b.get("style_note") or ""),
                "items": items,
            })
        out["batches"] = norm_batches
        return out

    flat = list(raw.get("components") or [])
    if not flat:
        raise ValueError("Manifest must contain 'batches' or 'components'")

    large = [it for it in flat if str(it.get("size_class", "")).lower() == "large"]
    regular = [it for it in flat if str(it.get("size_class", "")).lower() != "large"]

    batches: List[Dict[str, Any]] = []
    b_idx = 1
    for group, cap in ((large, min(4, max_per_batch)), (regular, max_per_batch)):
        for i in range(0, len(group), cap):
            chunk = group[i : i + cap]
            r, c = choose_grid(len(chunk))
            batches.append({
                "id": f"batch_{b_idx:02d}",
                "rows": r,
                "cols": c,
                "style_note": "",
                "items": chunk,
            })
            b_idx += 1
    out["batches"] = batches
    return out


def build_batch_prompt(manifest: Dict[str, Any], batch: Dict[str, Any], has_ref: bool, has_chain_ref: bool) -> str:
    rows, cols = int(batch["rows"]), int(batch["cols"])
    items = batch["items"]
    palette = manifest.get("palette_lock") or "executive consulting clean flat-vector aesthetic"
    style_desc = manifest.get("style_brief") or ""
    batch_note = batch.get("style_note") or ""

    ref_header = ""
    if has_ref and has_chain_ref:
        ref_header = (
            "Use Image 1 STRICTLY as the brand color palette and presentation aesthetic reference, "
            "and Image 2 STRICTLY as the sibling component sheet so line weight, corner radius, lighting, "
            "and icon proportions stay 100% consistent across batches. Do NOT copy text or full slide layouts.\n\n"
        )
    elif has_ref:
        ref_header = (
            "Use the attached reference image STRICTLY as a color palette and visual style reference "
            "(matching its color stops, stroke precision, and modern executive flat-vector aesthetic). "
            "Do NOT copy its slide layout or any text.\n\n"
        )

    lines = [
        ref_header
        + "Create a clean 16:9 UI VISUAL COMPONENT SPRITE SHEET on a 100% PURE SOLID FLAT WHITE (#FFFFFF) background.",
        f"Brand Palette & Aesthetic Lock: {palette}. {style_desc} {batch_note}".strip(),
        "",
        f"Arrange EXACTLY {len(items)} isolated, clean, text-free graphic components in a spacious, perfectly separated "
        f"{rows}-row × {cols}-column grid with LARGE pure white (#FFFFFF) empty margins around EVERY component "
        "(at least 10% empty white gap between any two components so an automated script can crop every slot cleanly):",
        "",
    ]

    for idx, it in enumerate(items):
        r_i, c_i = divmod(idx, cols)
        name = it.get("name") or f"comp_{idx+1:02d}"
        desc = it.get("prompt") or it.get("description") or name
        lines.append(f"{idx+1}. [Row {r_i+1}, Col {c_i+1}] ({name}): {desc}")

    empty_slots = rows * cols - len(items)
    if empty_slots > 0:
        lines.append(
            f"(Leave the remaining {empty_slots} trailing grid cell(s) completely EMPTY pure flat white #FFFFFF.)"
        )

    lines.extend([
        "",
        "【CRITICAL TECHNICAL RULES FOR AUTOMATED CUTTING — STRICTLY ENFORCE】",
        "- Background MUST be 100% uniform pure solid flat white (#FFFFFF) everywhere outside the components.",
        "- NO off-white tint, NO checkerboard transparency pattern, NO borders, NO dividing grid lines, NO outer frame, NO stray drop shadows bleeding far outside components.",
        "- ZERO text, ZERO Chinese characters, ZERO English letters, ZERO numbers (no '01', no captions, no labels) anywhere on the entire image! Pure graphic visual components ONLY!",
        "- Every component must be completely self-contained, sharp, modern consulting flat-vector quality, centered neatly inside its grid cell with generous pure white (#FFFFFF) breathing room.",
    ])
    return "\n".join(lines)


def _resolve_ref_path(raw_path: str, session: Path, manifest_path: Path) -> Optional[Path]:
    if not raw_path:
        return None
    p = Path(raw_path)
    if p.is_absolute() and p.is_file():
        return p
    for base in (session, manifest_path.parent, Path.cwd()):
        cand = base / raw_path
        if cand.is_file():
            return cand
    return None


def generate_one_batch(
    client: ImageClient,
    session: Path,
    manifest: Dict[str, Any],
    batch: Dict[str, Any],
    refs: List[Path],
    has_chain_ref: bool,
    force: bool,
    retries: int,
) -> Tuple[str, str, str]:
    bid = batch["id"]
    sheets_dir = session / "sheets"
    sheets_dir.mkdir(parents=True, exist_ok=True)
    out_png = sheets_dir / f"{bid}_raw.png"
    prompt_txt = sheets_dir / f"{bid}_prompt.txt"

    prompt = build_batch_prompt(manifest, batch, has_ref=bool(refs), has_chain_ref=has_chain_ref)
    prompt_txt.write_text(prompt, encoding="utf-8")

    if out_png.is_file() and not force:
        return bid, "skip", str(out_png)

    last_err = ""
    for attempt in range(1, retries + 1):
        try:
            t0 = time.time()
            if refs:
                client.edit(prompt, refs, out_png)
            else:
                client.generate(prompt, out_png)
            print(f"[ok] {bid} ({len(batch['items'])} components, {int(time.time()-t0)}s) -> {out_png}")
            return bid, "ok", str(out_png)
        except Exception as e:
            last_err = str(e)
            if attempt < retries and is_transient_error(last_err):
                wait = min(10 * attempt, 30)
                print(f"[retry] {bid} attempt {attempt} failed ({last_err[:120]}), retrying in {wait}s...")
                time.sleep(wait)
                continue
            return bid, "error", last_err
    return bid, "error", last_err


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate multi-batch pure-white component sprite sheets for SlideForge")
    ap.add_argument("--session", required=True, help="Session output directory")
    ap.add_argument("--manifest", required=True, help="Path to components_manifest.json")
    ap.add_argument("--only", default="", help="Comma-separated batch IDs to run, e.g. batch_01,batch_02")
    ap.add_argument("--max-per-batch", type=int, default=8, help="Max components per batch when auto-chunking")
    ap.add_argument("--concurrency", type=int, default=2)
    ap.add_argument("--retries", type=int, default=4)
    ap.add_argument("--model", default=None)
    ap.add_argument("--size", default=None)
    ap.add_argument("--timeout", type=int, default=600)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--dry-run", action="store_true", help="Resolve batches and write prompts without calling API")
    args = ap.parse_args()

    session = Path(args.session).resolve()
    manifest_path = Path(args.manifest).resolve()
    if not manifest_path.is_file():
        raise SystemExit(f"Manifest not found: {manifest_path}")

    raw = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest = normalize_manifest(raw, max_per_batch=args.max_per_batch)

    sheets_dir = session / "sheets"
    sheets_dir.mkdir(parents=True, exist_ok=True)
    resolved_path = sheets_dir / "manifest_resolved.json"
    resolved_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    only_set = {x.strip() for x in args.only.split(",") if x.strip()} if args.only else None
    batches = [b for b in manifest["batches"] if (not only_set or b["id"] in only_set)]
    if not batches:
        raise SystemExit("No matching batches to generate.")

    style_ref = _resolve_ref_path(str(manifest.get("style_ref") or ""), session, manifest_path)
    base_refs: List[Path] = [style_ref] if style_ref else []
    chain_batch1 = bool(manifest.get("chain_batch1", False))

    if args.dry_run:
        for idx, b in enumerate(batches):
            prompt = build_batch_prompt(manifest, b, has_ref=bool(base_refs), has_chain_ref=(chain_batch1 and idx > 0))
            (sheets_dir / f"{b['id']}_prompt.txt").write_text(prompt, encoding="utf-8")
            print(f"[dry-run] {b['id']}: {b['rows']}x{b['cols']} ({len(b['items'])} items) -> {b['id']}_prompt.txt")
        print(f"[dry-run] Resolved manifest saved -> {resolved_path}")
        return

    client = resolve_client(args)
    print(
        f"[info] model={client.model} batches={len(batches)} "
        f"style_ref={style_ref.name if style_ref else 'none'} chain_batch1={chain_batch1}"
    )

    results: List[Tuple[str, str, str]] = []
    remaining = list(batches)

    chain_ref: Optional[Path] = None
    if chain_batch1 and remaining:
        first = remaining.pop(0)
        res = generate_one_batch(client, session, manifest, first, base_refs, False, args.force, args.retries)
        results.append(res)
        if res[1] in ("ok", "skip"):
            chain_ref = Path(res[2])
        else:
            print(f"[error] {res[0]} failed: {res[2]}", file=sys.stderr)
            sys.exit(2)

    if remaining:
        refs_for_rest = base_refs + ([chain_ref] if chain_ref and chain_ref.is_file() else [])
        with ThreadPoolExecutor(max_workers=max(1, args.concurrency)) as ex:
            futs = [
                ex.submit(
                    generate_one_batch,
                    client, session, manifest, b, refs_for_rest, bool(chain_ref), args.force, args.retries,
                )
                for b in remaining
            ]
            for f in as_completed(futs):
                results.append(f.result())

    failed = [r for r in results if r[1] == "error"]
    if failed:
        for bid, _, err in failed:
            print(f"[error] {bid}: {err}", file=sys.stderr)
        sys.exit(2)
    print(f"[done] {len(results)} batch sprite sheet(s) ready in {sheets_dir}")


if __name__ == "__main__":
    main()
