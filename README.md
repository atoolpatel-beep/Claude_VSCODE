# Claude_VSCODE workspace

Workspace for new projects. Each project lives in its own folder under `projects/`.

## Toolchain (verified 2026-10-22)

| Tool   | Version  | Status                     |
|--------|----------|----------------------------|
| Git    | 2.55.0   | OK                         |
| Node   | 26.1.0   | OK                         |
| npm    | 11.13.0  | OK                         |
| Python | —        | Not installed (Store stub) |

## Starting a new project

```powershell
# Node / TypeScript
mkdir projects/my-app; cd projects/my-app; npm init -y

# Vite web app
cd projects; npm create vite@latest my-web-app
```

Shared conventions (`.editorconfig`, `.gitattributes`, `.gitignore`, `.vscode/`) apply to every project in this folder.
