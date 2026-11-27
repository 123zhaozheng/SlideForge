#!/usr/bin/env python3
"""Full-slide visual blueprint generator for SlideForge (Mode B: Blueprint-Replica).

Capabilities:
  - --check-env: safely verifies API key/base URL/model readiness without printing secrets
  - /v1/images/edits with optional corporate baseboard master + style reference image(s)
  - Optional real-photo / illustration slot skeletons (transparent orange corner brackets)
    that reserve clean whitespace in the blueprint and save exact inch coordinates for pptx_kit.add_image_slot()
  - /v1/images/generations fallback when no reference image is provided
  - Transient error backoff retry (429/502/503/504/524/timeout) + 16:9 aspect ratio guard
"""

from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import re
import struct
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
from urllib.parse import urlparse

import requests

try:
    from PIL import Image, ImageDraw
except ImportError:
    Image = None  # type: ignore
    ImageDraw = None  # type: ignore

SKILL_ROOT = Path(__file__).resolve().parent.parent
MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
TRANSIENT_MARKERS = (
    "429", "502", "503", "504", "524",
    "timeout", "timed out", "connection aborted", "remotedisconnected",
    "service unavailable", "temporarily unavailable", "internal_error",
    "server_error", "upstream", "gateway",
)
SLOT_CONTAINERS: Dict[str, Tuple[float, float, float, float]] = {
    "右侧": (0.55, 0.16, 0.40, 0.68),
    "right": (0.55, 0.16, 0.40, 0.68),
    "左侧": (0.06, 0.16, 0.40, 0.68),
    "left": (0.06, 0.16, 0.40, 0.68),
    "中央": (0.20, 0.18, 0.60, 0.62),
    "center": (0.20, 0.18, 0.60, 0.62),
    "居中": (0.20, 0.18, 0.60, 0.62),
    "底部": (0.08, 0.62, 0.84, 0.30),
    "bottom": (0.08, 0.62, 0.84, 0.30),
    "右上": (0.58, 0.14, 0.36, 0.40),
    "top-right": (0.58, 0.14, 0.36, 0.40),
    "左上": (0.06, 0.14, 0.36, 0.40),
    "top-left": (0.06, 0.14, 0.36, 0.40),
    "右下": (0.58, 0.54, 0.36, 0.38),
    "bottom-right": (0.58, 0.54, 0.36, 0.38),
    "左下": (0.06, 0.54, 0.36, 0.38),
    "bottom-left": (0.06, 0.54, 0.36, 0.38),
}
_CN_NUM = ["一", "二", "三", "四", "五", "六", "七", "八"]


def load_env() -> Optional[Path]:
    """Load skill-scoped .env only (never walk arbitrary parent directories)."""
    try:
        from dotenv import load_dotenv
    except ImportError:
        load_dotenv = None
    loaded_from: Optional[Path] = None
    if load_dotenv:
        candidates = [
            SKILL_ROOT / ".env",
            Path(os.environ.get("SLIDEFORGE_ENV", "")) if os.environ.get("SLIDEFORGE_ENV") else None,
            Path(os.environ.get("IMAGE2_PPT_ENV", "")) if os.environ.get("IMAGE2_PPT_ENV") else None,
        ]
        for p in candidates:
            if p and p.is_file():
                load_dotenv(p, override=False)
                if loaded_from is None:
                    loaded_from = p
    return loaded_from


def check_env_status() -> int:
    """Print safe environment check summary (never prints raw API keys)."""
    env_file = load_env()
    key = (os.environ.get("OPENAI_API_KEY") or os.environ.get("SLIDEFORGE_API_KEY") or "").strip()
    base = (os.environ.get("OPENAI_BASE_URL") or os.environ.get("SLIDEFORGE_BASE_URL") or "https://api.openai.com/v1").strip()
    model = (os.environ.get("SLIDEFORGE_MODEL") or os.environ.get("IMAGE2_PPT_MODEL") or "gpt-image-1").strip()
    size = (os.environ.get("SLIDEFORGE_SIZE") or os.environ.get("IMAGE2_PPT_SIZE") or "3840x2160").strip()
    quality = (os.environ.get("SLIDEFORGE_QUALITY") or os.environ.get("IMAGE2_PPT_QUALITY") or "high").strip()

    key_ok = bool(key) and not key.lower().startswith("sk-your") and "your-api-key" not in key.lower()
    masked = f"{key[:4]}...{key[-4:]}" if len(key) >= 10 else ("(set)" if key else "(missing)")
    print(f"ENV_STATUS: {'OK' if key_ok else 'MISSING_API_KEY'}")
    print(f"ENV_FILE: {env_file if env_file else 'process-env-only (no .env found)'}")
    print(f"API_KEY: {masked}")
    print(f"BASE_URL: {base}")
    print(f"MODEL: {model} (size={size}, quality={quality})")
    if not key_ok:
        print(
            f"ACTION: Copy {SKILL_ROOT / '.env.example'} to {SKILL_ROOT / '.env'} "
            "and fill in OPENAI_API_KEY, OPENAI_BASE_URL, and SLIDEFORGE_MODEL."
        )
        return 1
    return 0


def guess_mime(path: Path) -> str:
    return mimetypes.guess_type(str(path))[0] or "image/png"


def is_transient_error(msg: str) -> bool:
    low = (msg or "").lower()
    return any(m in low for m in TRANSIENT_MARKERS)


def parse_canvas_size(size: str) -> Tuple[int, int]:
    m = re.fullmatch(r"(\d+)\s*[xX]\s*(\d+)", (size or "").strip())
    return (int(m.group(1)), int(m.group(2))) if m else (3840, 2160)


def read_image_dims(path: Path) -> Tuple[int, int]:
    try:
        if Image is not None:
            with Image.open(path) as im:
                return im.size
    except Exception:
        pass
    try:
        head = path.read_bytes()[:24]
        if head[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack(">II", head[16:24])
    except Exception:
        pass
    return 0, 0


def aspect_ok(w: int, h: int, target_ratio: float = 16 / 9, tol: float = 0.15) -> bool:
    if not (w and h):
        return True
    return abs((w / h) - target_ratio) / target_ratio <= tol


class BlueprintImageClient:
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
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
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


def resolve_client(args: Optional[argparse.Namespace] = None) -> BlueprintImageClient:
    load_env()
    key = (os.environ.get("OPENAI_API_KEY") or os.environ.get("SLIDEFORGE_API_KEY") or "").strip()
    if not key:
        raise SystemExit(
            "OPENAI_API_KEY is not configured. Copy .env.example to .env inside the SlideForge skill folder "
            "or export OPENAI_API_KEY in your environment."
        )
    base = (os.environ.get("OPENAI_BASE_URL") or os.environ.get("SLIDEFORGE_BASE_URL") or "https://api.openai.com/v1").strip()
    model = (getattr(args, "model", None) or os.environ.get("SLIDEFORGE_MODEL") or os.environ.get("IMAGE2_PPT_MODEL") or "gpt-image-1").strip()
    quality = (getattr(args, "quality", None) or os.environ.get("SLIDEFORGE_QUALITY") or os.environ.get("IMAGE2_PPT_QUALITY") or "high").strip()
    size = (getattr(args, "size", None) or os.environ.get("SLIDEFORGE_SIZE") or os.environ.get("IMAGE2_PPT_SIZE") or "3840x2160").strip()
    resolution = (getattr(args, "resolution", None) or os.environ.get("SLIDEFORGE_RESOLUTION") or os.environ.get("IMAGE2_PPT_RESOLUTION") or "4k").strip()
    aspect_ratio = (getattr(args, "aspect_ratio", None) or os.environ.get("SLIDEFORGE_ASPECT") or os.environ.get("IMAGE2_PPT_ASPECT") or "16:9").strip()
    timeout = int(getattr(args, "timeout", 600) or 600)
    trust_raw = (os.environ.get("SLIDEFORGE_TRUST_ENV") or os.environ.get("IMAGE2_PPT_TRUST_ENV") or "false").strip().lower()
    trust_env = trust_raw in ("1", "true", "yes")
    return BlueprintImageClient(
        api_key=key,
        base_url=base,
        model=model,
        quality=quality,
        size=size,
        resolution=resolution,
        aspect_ratio=aspect_ratio,
        timeout=timeout,
        trust_env=trust_env,
    )


def build_ref_manifest(
    refs: Sequence[Path],
    baseboard: Optional[Path],
    skel_paths: Sequence[Path],
    extra_refs: Sequence[Path],
) -> str:
    if not refs:
        return ""
    lines = ["【参考图说明 — 按传入顺序严格区分每张图的角色】"]
    for i, p in enumerate(refs):
        cn = _CN_NUM[i] if i < len(_CN_NUM) else str(i + 1)
        if baseboard and p == baseboard:
            role = "唯一母版底板 — 像素级保留其顶/底装饰、Logo、背景纹理与标题样式，严禁复制底板原标题文字"
        elif p in skel_paths:
            role = "真实图片/插图槽位骨架 — 透明画布橙色角标区域必须留白避让，不画占位框、不保留角标"
        elif p in extra_refs:
            role = "风格与审美参考图 — 仅借鉴配色阶梯、版式呼吸感、微组件工艺与信息密度，严禁抄袭其标题、水印、无关文字或具体数据"
        else:
            role = "参考图"
        lines.append(f"图片{i+1}（第{cn}张，{p.name}）：{role}")
    return "\n".join(lines) + "\n\n"


def _build_skeleton(canvas_w: int, canvas_h: int, rects: Sequence[Tuple[float, float, float, float]], out: Path) -> None:
    if Image is None or ImageDraw is None:
        raise RuntimeError("Pillow is required for slot skeletons: pip install Pillow")
    img = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    L, LW, color = max(24, canvas_w // 100), max(5, canvas_w // 480), (255, 90, 0, 235)
    for x, y, w, h in rects:
        x, y, w, h = int(round(x)), int(round(y)), int(round(w)), int(round(h))
        for seg in (
            [(x, y), (x + L, y)], [(x, y), (x, y + L)],
            [(x + w, y), (x + w - L, y)], [(x + w, y), (x + w, y + L)],
            [(x, y + h), (x + L, y + h)], [(x, y + h), (x, y + h - L)],
            [(x + w, y + h), (x + w - L, y + h)], [(x + w, y + h), (x + w, y + h - L)],
        ):
            draw.line(seg, fill=color, width=LW)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out, "PNG")


def _parse_slot_box(pos: str) -> Tuple[float, float, float, float]:
    if re.fullmatch(r"[\d.]+(\s*,\s*[\d.]+){3}", pos):
        vals = [float(v) for v in re.split(r"\s*,\s*", pos)]
        return vals[0], vals[1], vals[2], vals[3]
    low = pos.lower()
    for k, box in SLOT_CONTAINERS.items():
        if k in pos or k in low:
            return box
    return SLOT_CONTAINERS["右侧"]


def prepare_slots(
    client: BlueprintImageClient,
    session: Path,
    job: Dict[str, Any],
    num: int,
    slide_w_in: float = 13.333,
    slide_h_in: float = 7.5,
) -> Tuple[List[Path], List[Dict[str, Any]], str]:
    """Generate orange corner-bracket skeleton PNGs for real photo / illustration slots."""
    slots = job.get("slots") or []
    if not slots:
        return [], [], ""
    cw, ch = parse_canvas_size(client.size)
    refs_dir = session / "references"
    skels, records, items = [], [], []
    for idx, s in enumerate(slots, 1):
        pos = str(s.get("position") or "右侧").strip()
        bx, by, bw, bh = _parse_slot_box(pos)
        rect_px = (bx * cw, by * ch, bw * cw, bh * ch)
        rect_in = [
            round(bx * slide_w_in, 3),
            round(by * slide_h_in, 3),
            round(bw * slide_w_in, 3),
            round(bh * slide_h_in, 3),
        ]
        sid = str(s.get("id") or f"slot{idx}")
        label = str(s.get("label") or s.get("description") or "").strip()
        skel = refs_dir / f"slide-{num:02d}-skeleton-{sid}.png"
        _build_skeleton(cw, ch, [rect_px], skel)
        skels.append(skel)
        records.append({
            "id": sid,
            "source": str(s.get("source") or ""),
            "label": label,
            "position": pos,
            "rect_norm": [round(bx, 4), round(by, 4), round(bw, 4), round(bh, 4)],
            "rect_in": rect_in,
            "rect_px": [round(v, 1) for v in rect_px],
        })
        desc = f" ({label})" if label else ""
        items.append(f"- {sid}{desc}: {pos}")
    block = (
        "\n\n【真实图片/插图槽位 — 必须留白避让】\n"
        + "\n".join(items)
        + "\n- 橙色角标标记区域为后续原生 PPTX 插入真实照片/产品截图/定制插图的预留槽位，请保持该区域干净留白、背景自然连续，禁止画入文字、占位框或保留橙色角标。"
    )
    return skels, records, block


def load_jobs(session: Path, only: Optional[List[int]]) -> List[Dict[str, Any]]:
    for name in ("blueprint_manifest.json", "prompts_manifest.json"):
        man_path = session / name
        if man_path.is_file():
            jobs = json.loads(man_path.read_text(encoding="utf-8")).get("slides") or []
            break
    else:
        jobs = []
        prompt_dir = session / "blueprint_prompts"
        if not prompt_dir.is_dir():
            prompt_dir = session / "prompts"
        for p in sorted(prompt_dir.glob("slide-*.txt")):
            m = re.search(r"slide-(\d+)", p.stem)
            if m:
                num = int(m.group(1))
                jobs.append({
                    "number": num,
                    "title": p.stem,
                    "prompt_file": str(p.relative_to(session)).replace("\\", "/"),
                    "image_file": f"blueprints/slide-{num:02d}.png",
                    "slots": [],
                })
    if only:
        only_set = set(only)
        jobs = [j for j in jobs if int(j["number"]) in only_set]
    return jobs


def run_one(
    client: BlueprintImageClient,
    session: Path,
    job: Dict[str, Any],
    baseboard: Optional[Path],
    extra_refs: List[Path],
    force: bool,
    retries: int,
    aspect_retries: int,
) -> Tuple[int, str, str, List[Dict[str, Any]]]:
    num = int(job["number"])
    prompt_path = session / job["prompt_file"]
    out_path = session / job.get("image_file", f"blueprints/slide-{num:02d}.png")
    if out_path.is_file() and not force:
        return num, "skip", str(out_path), []
    if not prompt_path.is_file():
        return num, "error", f"missing prompt: {prompt_path}", []

    prompt = prompt_path.read_text(encoding="utf-8")
    skels, slot_records, slot_block = prepare_slots(client, session, job, num)
    refs: List[Path] = ([baseboard] if baseboard else []) + skels + extra_refs
    full_prompt = build_ref_manifest(refs, baseboard, skels, extra_refs) + prompt + slot_block

    last_err = ""
    for ar_try in range(aspect_retries + 1):
        working = full_prompt + ("\n\n【强制16:9横版宽屏】宽度必须明显大于高度。" if ar_try else "")
        for attempt in range(1, retries + 1):
            try:
                t0 = time.time()
                res = client.edit(working, refs, out_path) if refs else client.generate(working, out_path)
                w, h = read_image_dims(res)
                if ar_try < aspect_retries and not aspect_ok(w, h):
                    print(f"[warn] blueprint slide-{num:02d} size {w}x{h} deviates from 16:9, retrying ({ar_try+1}/{aspect_retries})")
                    res.unlink(missing_ok=True)
                    break
                print(f"[ok] blueprint slide-{num:02d} ({w}x{h}, {int(time.time()-t0)}s) -> {res}")
                return num, "ok", str(res), slot_records
            except Exception as e:
                last_err = str(e)
                if attempt < retries and is_transient_error(last_err):
                    wait = min(2 ** attempt, 20)
                    print(f"[retry] blueprint slide-{num:02d} attempt {attempt} failed ({last_err[:120]}), waiting {wait}s")
                    time.sleep(wait)
                    continue
                return num, "error", last_err, []
    return num, "error", last_err or "aspect retries exhausted", []


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate full-slide visual blueprints for SlideForge (Mode B)")
    ap.add_argument("--check-env", action="store_true", help="Verify .env / API key configuration safely and exit")
    ap.add_argument("--session", default="", help="Session directory containing blueprint_manifest.json or prompts/")
    ap.add_argument("--baseboard", default="", help="Optional user template master image path")
    ap.add_argument("--ref", action="append", default=[], help="Optional style/aesthetic reference image(s)")
    ap.add_argument("--only", default="", help="Comma-separated slide numbers to generate first, e.g. 1,2")
    ap.add_argument("--model", default=None)
    ap.add_argument("--quality", default=None)
    ap.add_argument("--size", default=None)
    ap.add_argument("--concurrency", type=int, default=2)
    ap.add_argument("--retries", type=int, default=3)
    ap.add_argument("--aspect-retries", type=int, default=2)
    ap.add_argument("--timeout", type=int, default=600)
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    if args.check_env:
        raise SystemExit(check_env_status())

    if not args.session:
        ap.error("--session is required unless --check-env is passed")

    session = Path(args.session)
    if not session.is_dir():
        raise SystemExit(f"Session not found: {session}")

    baseboard: Optional[Path] = Path(args.baseboard) if args.baseboard else session / "baseboard.png"
    if baseboard and not baseboard.is_file():
        if args.baseboard:
            raise SystemExit(f"Baseboard not found: {baseboard}")
        baseboard = None
        print("[info] No baseboard.png found; generating blueprints in clean-canvas style mode.")

    extra_refs = [Path(p) for p in args.ref]
    for p in extra_refs:
        if not p.is_file():
            raise SystemExit(f"Reference image not found: {p}")

    only = [int(x.strip()) for x in args.only.split(",") if x.strip()] if args.only else None
    jobs = load_jobs(session, only)
    if not jobs:
        raise SystemExit("No blueprint prompts found in session (expected blueprint_manifest.json or prompts/slide-*.txt)")

    client = resolve_client(args)
    (session / "blueprints").mkdir(exist_ok=True)

    meta_path = session / "blueprint_metadata.json"
    meta = json.loads(meta_path.read_text(encoding="utf-8")) if meta_path.is_file() else {"slides": {}}
    meta.update({
        "model": client.model,
        "size": client.size,
        "canvas": list(parse_canvas_size(client.size)),
        "baseboard": str(baseboard) if baseboard else "",
        "style_refs": [str(p) for p in extra_refs],
    })

    print(f"[info] model={client.model} size={client.size} jobs={len(jobs)} refs={int(bool(baseboard))+len(extra_refs)}")
    results = []
    with ThreadPoolExecutor(max_workers=max(1, args.concurrency)) as ex:
        futs = [
            ex.submit(run_one, client, session, j, baseboard, extra_refs, args.force, args.retries, args.aspect_retries)
            for j in jobs
        ]
        for f in as_completed(futs):
            results.append(f.result())

    for num, status, info, slot_records in results:
        entry = meta.setdefault("slides", {}).setdefault(str(num), {})
        entry.update({
            "number": num,
            "status": status,
            "image": info if status in ("ok", "skip") else entry.get("image", ""),
            "error": info if status == "error" else "",
            "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        })
        if slot_records:
            entry["slots"] = slot_records
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    failed = [r for r in results if r[1] == "error"]
    if failed:
        for num, _, err, _ in failed:
            print(f"[error] blueprint slide-{num:02d}: {err}", file=sys.stderr)
        sys.exit(2)
    print("[done] all blueprints ready")


if __name__ == "__main__":
    main()
