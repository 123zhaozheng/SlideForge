# Security Policy

## Credentials and secrets

- Never commit `.env`, API keys, or private corporate slide templates to the repository.
- SlideForge reads `OPENAI_API_KEY` and `OPENAI_BASE_URL` strictly from local environment variables or a gitignored `.env` file.

## Reporting a vulnerability

If you discover a security issue (such as unsafe URL handling, path traversal, or command injection in helper scripts), please open a private security advisory on GitHub or report it via repository Issues without posting exploit payloads or live credentials.
