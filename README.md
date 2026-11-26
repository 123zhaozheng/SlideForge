# SlideForge

<p align="center">
  <a href="README.zh-CN.md">简体中文</a>
</p>

<h1 align="center">SlideForge</h1>

<p align="center">
  <b>Brief or flat slide in. Native, editable PPTX out.</b><br />
  An Agent-first skill that builds layered PowerPoint decks from AI component sprite sheets and OOXML vector code, verified by real PowerPoint/LibreOffice rendering.
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-3da639" alt="MIT License" /></a>
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-2f7af8" alt="Windows, macOS, and Linux" />
  <img src="https://img.shields.io/badge/Python-%E2%89%A53.10-3776AB?logo=python&logoColor=white" alt="Python 3.10+" />
  <img src="https://img.shields.io/badge/output-Native%20Editable%20PPTX-f97316" alt="Native Editable PPTX" />
  <a href="SKILL.md"><img src="https://img.shields.io/badge/Agent_Skill-SKILL.md-111827" alt="Agent Skill" /></a>
</p>

<p align="center">
  <a href="#features">Core features</a> ·
  <a href="#comparisons">Comparisons</a> ·
  <a href="#architecture">Architecture</a> ·
  <a href="#quick-start">Quick start</a> ·
  <a href="#agent-workflow">Agent workflow</a> ·
  <a href="#contributing">Contributing</a> ·
  <a href="#links">Links</a>
</p>

---

<a id="comparisons"></a>

## Comparisons

<!-- Drop comparison screenshots into docs/assets/ and uncomment below -->
<!--
<p align="center">
  <img src="docs/assets/compare-1.png" alt="Comparison 1: flat AI slide vs native editable PPTX" width="1200" />
</p>
-->

> Comparison screenshots (input flat slide / document vs. native editable PPTX opened in PowerPoint) will be published in [`docs/assets/`](docs/assets/).

<a id="features"></a>

## Core features

### Forward layered reconstruction instead of OCR erasing

Traditional image-to-PPT converters run OCR, erase text regions by inpainting, and crop noisy bitmaps. SlideForge separates responsibilities cleanly across four native layers:

- **Layer 1 — User background**: Keeps any user-supplied corporate slide template (with its existing logo, header rule, and footer) intact on the base layer, or runs on a clean canvas when no template is supplied.
- **Layer 2 — Native OOXML vectors**: Rebuilds cards, dashed Bento frames, linear gradients, exact percentage donut arcs, funnels, tables, and connectors as true PowerPoint shapes.
- **Layer 3 — Isolated transparent AI components**: Sends only icons, badges, illustrations, and soft-fading graphics to an OpenAI-compatible image model on a pure `#FFFFFF` background, then extracts clean RGBA PNGs.
- **Layer 4 — Editable typography**: Renders all text with explicit East Asian (`<a:ea>`), Latin (`<a:latin>`), and Complex Script (`<a:cs>`) font tags so CJK text never falls back to SimSun.

### Multi-batch component sprite sheets with style locking

Packing 12+ icons into one image generation call degrades detail and causes overlap. `generate_components.py` caps each batch at `<= 8` icons (`2x4`) or `<= 4–6` larger cards (`2x2` / `2x3`), automatically splits larger manifests across multiple calls, and binds every batch to the same `style_ref` and `palette_lock` (with optional `chain_batch1` chaining) so stroke weight and palette stay consistent.

### Precision alpha extraction and visual verification

`extract_components.py` slices sprite sheets via row/column projection valleys and supports two alpha modes:

- `solid`: exterior BFS flood-fill with 1px edge defringing, preserving pure-white icons inside colored badges.
- `soft_fade`: white-background un-premultiplication (`C = αF + (1 - α)255`) for smooth fading arrows and gradients.

It also strips isolated outlier specks, generates optional tinted variants (such as `"tint": "white"` for dark hero surfaces), and renders a checkerboard `*_alpha_preview.png` for inspection before PPTX assembly.

### Two production modes

| Mode | Input | Flow | Best for |
| :--- | :--- | :--- | :--- |
| **Mode A · Direct-Design** | Brief/document + optional background template | Skip full-slide image → plan layout → generate component sheets → assemble PPTX → render QA | Fast everyday decks, architecture slides, KPI reports |
| **Mode B · Blueprint-Replica** | Existing flat slide image, or brief → full-slide blueprint first | Use full-slide image as `style_ref` → split & extract components → reconstruct 1:1 in native PPTX → render QA | Executive keynotes, editorial layouts, converting flat AI decks into editable PPTX |

### Real PowerPoint / WPS / LibreOffice render loop

`render_pptx.py` exports every generated `.pptx` back to 1080p PNGs so the Agent can inspect actual typography wrapping, alignment, and margins before returning the file. On Windows it invokes native PowerPoint/WPS COM with zero `pywin32` dependency and a cross-process file lock so parallel sub-agents do not collide; on macOS/Linux it falls back to LibreOffice + PyMuPDF.

<a id="architecture"></a>

## Architecture

```text
brief.md / target_slide.png (+ optional template_bg.png)
                        |
                        v
            Agent planning & manifest
          (components_manifest.json)
                        |
                        v
      generate_components.py (multi-batch, #FFFFFF, 0 text)
                        |
                        v
      extract_components.py (BFS / soft_fade -> RGBA PNG + checkerboard QA)
                        |
                        v
      build_deck.py / build_slide.py (pptx_kit.py OOXML 4-layer assembly)
                        |
                        v
      render_pptx.py (PowerPoint/WPS COM lock or LibreOffice -> PNG QA)
                        |
                        v
               output_editable.pptx
```

### Key OOXML pitfalls handled in `pptx_kit.py`

- **CJK font fallback**: Explicitly writes `<a:latin>`, `<a:ea>`, and `<a:cs>` on every text run.
- **Default shape shadows**: Writes an explicit empty `<a:effectLst/>` on flat shapes so PowerPoint's theme shadow does not bleed onto dividers, pills, or dashed frames.
- **`BLOCK_ARC` angle scaling**: `add_donut_ring()` automatically converts degrees (`deg * 0.6`) to reconcile OOXML's 60,000ths-of-a-degree angle unit with `python-pptx`'s `100,000` adjustment multiplier.
- **Native table cleanup**: Strips default blue banded table style IDs and applies clean hairline row borders.

<a id="quick-start"></a>

## Quick start

### Requirements

- Python `>= 3.10` (recommended: [`uv`](https://docs.astral.sh/uv/))
- An OpenAI-compatible image generation API key (supports `gpt-image-1`, `gpt-image-2.6`, etc.)
- For visual QA rendering: Microsoft PowerPoint or WPS Office on Windows, or LibreOffice + PyMuPDF on macOS/Linux

### Setup

```bash
git clone https://github.com/123zhaozheng/SlideForge.git
cd SlideForge
cp .env.example .env
# Fill in OPENAI_API_KEY and OPENAI_BASE_URL in .env
```

Windows (PowerShell):

```powershell
.\install.ps1
```

macOS / Linux:

```bash
chmod +x install.sh && ./install.sh
```

### Dry-run test (no API call)

```bash
uv run --with requests --with pillow --with python-dotenv \
  python scripts/generate_components.py \
  --session ./outputs/demo \
  --manifest examples/components_manifest.sample.json \
  --dry-run
```

<a id="agent-workflow"></a>

## Agent workflow

SlideForge ships as a standard Agent Skill ([`SKILL.md`](SKILL.md)) compatible with Claude Code, Cursor, Codex, and Antigravity.

1. Place this repository in your Agent's skills directory (or point the Agent directly to `SlideForge/SKILL.md`).
2. Give the Agent a brief, a template background, or an existing flat slide image:
   - **Mode B (recreate a flat slide)**:
     > Read `SlideForge/SKILL.md` and recreate `target_slide.png` into a 100% native editable PPTX, then verify it via `render_pptx.py`.
   - **Mode A (build from a document)**:
     > Read `SlideForge/SKILL.md` (Mode A) and build an 8-slide editable PPTX from `brief.md` using `template_bg.png` as the slide background.
3. For multi-page decks, parallel sub-agents can each build one slide in `slides/slide_XX/build_slide.py` exporting `build_slide(prs)` and merge all slides cleanly into one presentation.

### Human vs Agent responsibilities

| Human | Agent |
| :--- | :--- |
| Supply brief, brand template (optional), or reference slide | Split native vector elements vs. AI graphic components |
| Choose Mode A (fast) or Mode B (blueprint-led) | Run multi-batch sprite generation & verify alpha cutouts |
| Make final wording or layout tweaks directly in PowerPoint | Assemble OOXML PPTX, render back to PNG, and fix alignment/overflow |

<a id="contributing"></a>

## Contributing

```bash
# Syntax check & smoke test
uv run --with requests --with pillow --with numpy --with python-pptx --with python-dotenv \
  python -m py_compile scripts/*.py
```

Pull requests should preserve the `pptx_kit.py` OOXML guarantees (explicit `<a:ea>`, clean `<a:effectLst/>`, aspect-ratio-safe placement) and keep scripts free of brand-specific hardcoding. See [CONTRIBUTING.md](CONTRIBUTING.md) and [SECURITY.md](SECURITY.md).

<a id="links"></a>

## Links

- [Repository](https://github.com/123zhaozheng/SlideForge)
- [Agent skill specification](SKILL.md)
- [Component splitting guide](references/component-splitting.md)
- [Design systems](references/design-systems.md)
- [Layout archetypes](references/layout-archetypes.md)
- [python-pptx OOXML pitfalls](references/pptx-pitfalls.md)
- [LINUX DO - community discussion](https://linux.do/)
- [Issues](https://github.com/123zhaozheng/SlideForge/issues)

## License

Released under the [MIT License](LICENSE). Third-party brand templates, fonts, and user-provided images retain their respective licenses.
