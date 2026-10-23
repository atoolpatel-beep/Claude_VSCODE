# CLAUDE.md

Multi-project workspace on Windows 11 (PowerShell + Git Bash).

- Create each new project in its own folder under `projects/<name>/`; keep project-specific config (package.json, tsconfig, pyproject) inside that folder.
- Available: Git, Node 26, npm 11. Python is NOT installed — ask before choosing a Python stack.
- Line endings are LF (enforced via `.gitattributes` / `.editorconfig`), except .cmd/.bat/.ps1.
- Never commit `.env` files or secrets; provide `.env.example` instead.
