# Contributing to SlideForge

Thank you for considering contributing to **SlideForge**.

## Development setup

1. Fork and clone the repository.
2. Copy `.env.example` to `.env` if you need live image-generation tests (never commit `.env`).
3. Verify syntax and smoke tests locally:

```bash
uv run --with requests --with pillow --with numpy --with python-pptx --with python-dotenv \
  python -m py_compile scripts/*.py
```

## Pull request checklist

- Keep scripts brand-agnostic and path-agnostic (no hardcoded user directories, private relay URLs, or organization-specific assets).
- Preserve `pptx_kit.py` OOXML invariants: explicit `<a:ea>` font tags, explicit `<a:effectLst/>` shadow clearing on flat shapes, and aspect-ratio-safe `add_picture_fit`.
- Update both `README.md` and `README.zh-CN.md` when changing user-facing flags or workflows.
- Ensure CI passes on both Linux and Windows.
