# CLAUDE.md — AI Assistant Guide for Loxone Commissioning Tool

## Project Overview

A professional desktop commissioning tool for Loxone home automation installers. Technicians use it to systematically test all inputs/outputs on a Loxone Miniserver and generate a signed PDF acceptance report. The project is deployed as a standalone Windows `.exe`.

**Language:** Python 3.9+  
**UI frameworks:** Tkinter (desktop GUI) and Flask (web/mobile alternative)  
**Deployment:** PyInstaller single-file EXE for Windows

---

## Repository Structure

```
/
├── loxone_checklist.py          # Core library — API, parsing, PDF generation (942 lines)
├── loxone_checklist_gui.py      # Desktop GUI application — Tkinter (1371 lines)
├── loxone_webapp.py             # Web UI alternative — Flask (929 lines)
├── LoxoneCommissioning.spec     # PyInstaller build specification
├── build_exe.bat                # One-click Windows EXE builder
├── make_icon.py                 # Icon generation utility (requires Pillow)
├── requirements_loxone.txt      # Runtime dependencies
├── loxone_icon.ico              # Application icon (bundled into EXE)
├── loxone_icon_preview.png      # Icon preview
├── LoxoneCommissioning.exe      # Pre-built Windows executable (not for editing)
└── README.md                    # User-facing documentation
```

---

## Architecture

The project follows a **layered architecture** with strict separation between logic and presentation.

### Layer 1: Core Library (`loxone_checklist.py`)

Stateless, reusable — no UI dependencies. Responsibilities:
- Loxone Miniserver HTTP API communication
- Token-based and Basic authentication
- JSON structure file parsing
- Control type classification
- PDF report generation via ReportLab

Key exports:
- `fetch_structure(host, username, password, use_https, use_token)` — downloads `LoxAPP3.json`
- `load_local_structure(file_path)` — loads offline JSON
- `parse_structure(structure)` — returns `(inputs, outputs, others)` lists
- `generate_pdf(inputs, outputs, others, ms_info, output_path, include_other, company)` — creates PDF
- Constants: `INPUT_TYPES`, `OUTPUT_TYPES`, `CONTROL_CMDS`, `DIMMER_TYPES`, `PRIMARY_STATE`

### Layer 2: GUI (`loxone_checklist_gui.py`)

Tkinter desktop application. Three-tab layout:
1. **Connect** — Live Miniserver or local JSON file, company branding config
2. **Test I/O** — Hierarchical I/O tree, live value reading, command buttons, status tracking, session save/load
3. **Report** — Filtered PDF generation

Uses background threads for long-running operations (connections, polling, PDF generation).

### Layer 3: Web UI (`loxone_webapp.py`)

Flask-based mobile-friendly alternative. Same core library, HTML template embedded in the Python file, modal-based interface.

---

## Key Concepts

### Control Type Classification

Loxone controls are classified into three categories defined in `loxone_checklist.py`:

```python
INPUT_TYPES  = {"DigitalInput", "AnalogInput", "Presence", "Motion", ...}
OUTPUT_TYPES = {"Switch", "Jalousie", "Dimmer", "LightController", ...}
# Everything else → "others" (logical controllers, scenes, etc.)
```

`parse_structure()` flattens the nested `LoxAPP3.json` hierarchy (controls + subControls) into these three lists.

### Test Status Values

Each control is tracked by its UUID. Valid status values:

| Value | Display | Meaning |
|-------|---------|---------|
| `"ok"` | ✓ OK | Tested and working |
| `"nok"` | ✗ Not OK | Tested with fault |
| `"skip"` | → Skip | Deferred/skipped |
| `"ignored"` | ⊘ Ignore | Excluded from checklist |
| (absent) | — | Untested |

### Session File Format

User-created `.json` files for resuming multi-day commissions:

```json
{
  "timestamp": "ISO-8601 datetime",
  "saved_at": "human-readable datetime",
  "project": "project name string",
  "ms_info": { "serialNr": "...", "projectName": "...", "swVersion": "..." },
  "statuses": { "<uuid>": "ok|nok|skip|ignored", ... },
  "notes": { "<uuid>": "fault description text", ... },
  "ignored": ["<uuid>", ...],
  "summary": { "ok": 0, "nok": 0, "skip": 0, "untested": 0 }
}
```

### Runtime Config File (`loxone_tool_config.json`)

Auto-created next to the executable on first run. Stores company branding:

```json
{
  "company_name": "Acme Installations",
  "company_sub": "Smart Home Division"
}
```

---

## Loxone API

All HTTP requests target the Loxone Miniserver:

| Endpoint | Purpose |
|----------|---------|
| `GET /jdev/sys/getkey` | Acquire key for token auth |
| `GET /jdev/sys/gettoken/{hash}/{user}/4/{uuid}/app` | Token acquisition (Gen2) |
| `GET /data/LoxAPP3.json` | Full structure download |
| `GET /jdev/sps/io/{uuid}/{cmd}` | Send control command |
| `GET /jdev/sps/getvalue/{state-uuid}` | Read live state value |

**Auth:** Basic Auth for Gen1, Bearer token preferred for Gen2. Self-signed HTTPS certificates are silently accepted (`verify=False`).

---

## Development Workflow

### Running from Source

```bash
# Install dependencies
pip install -r requirements_loxone.txt

# Desktop GUI
python loxone_checklist_gui.py

# Web UI (open browser at http://localhost:5000)
python loxone_webapp.py

# CLI / library usage
python loxone_checklist.py 192.168.1.100 admin secret
python loxone_checklist.py 192.168.1.100 admin secret --https --token -o report.pdf
python loxone_checklist.py --file LoxAPP3.json  # offline mode
```

### Building the Windows EXE

Run on Windows:
```batch
build_exe.bat
```

This installs PyInstaller + dependencies and outputs `dist/LoxoneCommissioning.exe`.

Manual build:
```bash
pip install pyinstaller requests reportlab pillow
pyinstaller LoxoneCommissioning.spec
```

---

## Code Conventions

### Style
- **snake_case** for functions and variables
- **UPPER_SNAKE_CASE** for module-level constants (`INPUT_TYPES`, `OUTPUT_TYPES`, etc.)
- **Classes** use PascalCase (only used for Tkinter widgets)
- Docstrings on all public functions; inline comments for non-obvious logic

### Threading
Long-running operations in the GUI always run in background threads to keep the UI responsive. Use `self.root.after()` to marshal results back to the main thread — never update Tkinter widgets from a worker thread directly.

### Adding New Control Types
1. Add the type string to `INPUT_TYPES` or `OUTPUT_TYPES` in `loxone_checklist.py`
2. If the control needs custom commands, add an entry to `CONTROL_CMDS`
3. If it has a dimmer-like slider, add to `DIMMER_TYPES`
4. If its primary state key differs from `"active"`, add to `PRIMARY_STATE`

### PDF Generation
`generate_pdf()` in `loxone_checklist.py` uses ReportLab's Platypus (flowable-based layout). Colour scheme: dark blue `#1a3a5c` headers, Loxone green `#6EBD44` accents, status-coloured rows.

### Error Handling
- Network errors during API calls are caught and surfaced to the user via status bar or dialog
- Missing optional features (e.g., token auth failure) fall back gracefully to Basic Auth
- File I/O errors (session load/save, PDF write) show error dialogs, never crash silently

---

## No Test Suite

There is currently no automated test framework. Testing is done manually by running the GUI or Flask server against a real or simulated Miniserver. When adding tests, use `pytest`.

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `requests` | ≥2.31.0 | HTTP API calls to Miniserver |
| `reportlab` | ≥4.0.0 | PDF report generation |
| `urllib3` | ≥2.0.0 | HTTPS/SSL support |
| `Pillow` | optional | Icon generation (`make_icon.py` only) |
| `tkinter` | stdlib | Desktop GUI |
| `flask` | optional | Web UI (`loxone_webapp.py` only) |

---

## Important Notes for AI Assistants

- **Do not modify `LoxoneCommissioning.exe`** — it is a pre-built binary artifact.
- The core library (`loxone_checklist.py`) must remain importable as a module **and** runnable as a CLI script — preserve the `if __name__ == "__main__":` guard.
- The `loxone_checklist_gui.py` and `loxone_webapp.py` are independent UI layers; changes to control type logic belong in `loxone_checklist.py`, not the UI files.
- Self-signed certificate warnings are suppressed intentionally — do not remove `verify=False` from API calls without adding proper cert handling.
- UUID normalization (stripping `{}` braces) is critical for matching Loxone state UUIDs — the code handles both `{uuid}` and `uuid` formats deliberately.
- The session file `"ignored"` list and `statuses` dict both need to stay in sync when a user toggles ignore state.
