# AGENTS.md

## Cursor Cloud specific instructions

### Repository layout

This repo is **not** a Node/monorepo. It contains:

| Artifact | Runnable in Cloud VM? |
|----------|------------------------|
| `ELGIN.py` | Yes — Python FastAPI shop app (primary dev target) |
| `gemini1.bas`, `NEED FIX.bas` | No — SolidWorks VBA on Windows only |

### Elgin CMS (`ELGIN.py`)

**Start the app** (required for CNC polling threads — do not use bare `uvicorn ELGIN:app`):

```bash
python3 ELGIN.py
```

Listens on **0.0.0.0:2926** (`ELGIN_APP_PORT` / `NEXUS_APP_PORT` override). Creates `shop_analytics_pro.db` and `uploads/` in the working directory.

**Lint** (no repo linter configured):

```bash
python3 -m py_compile ELGIN.py
```

**Tests / build**: none in-repo.

**Default admin password**: `cms123` (override with `ELGIN_ADMIN_PASSWORD` or `NEXUS_ADMIN_PASSWORD`).

**Optional runtime** (graceful degradation if missing):

- Haas Q-code (`10.1.10.18/19:5051`) and MTConnect (`:8082`) — machines show **OFFLINE** when unreachable (normal in Cloud Agent VMs).
- `openpyxl`, `PyPDF2`, `pillow`, `pytesseract` — steel-sheet / OCR features; core UI works without them.
- `ELGIN_SSL_CERT` / `ELGIN_SSL_KEY` — HTTPS for iPhone live camera barcode scan only.

**Long-running server**: use a named tmux session (e.g. `elgin-server`) rather than a one-shot background shell.

### CMS XT Export macros

Import `.bas` into SolidWorks on Windows (**Tools → Macro → Edit**), configure job/network paths in the module, then run `Sub main()`. `NEED FIX.bas` can POST job signatures to ELGIN at `http://localhost:2926` when `AUTO_UPLOAD_JOB_SIGNATURE_TO_ELGIN` is enabled.

### Smoke verification (API)

```bash
curl -s -c /tmp/elgin.txt -X POST http://127.0.0.1:2926/api/admin/login \
  -H 'Content-Type: application/json' -d '{"password":"cms123"}'
curl -s -b /tmp/elgin.txt http://127.0.0.1:2926/api/admin/me
```

Browser UI: http://127.0.0.1:2926 — sidebar **LOGIN** with `cms123`, then **FLEET** / **INVENTORY / QA** tabs.
