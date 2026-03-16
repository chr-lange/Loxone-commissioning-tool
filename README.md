# Loxone I/O Commissioning Tool

A desktop GUI for on-site commissioning of Loxone Miniserver installations.
Connect to any Miniserver, test every physical input and output, and export a signed-off PDF report — all without opening Loxone Config.

---

## Features

- **Live connection** to Loxone Miniserver via HTTP API (no Loxone Config needed on-site)
- **Auto-discovers** all inputs, outputs, and wireless devices (AIR & Tree)
- **Trigger outputs** directly from the tool — lights, blinds, HVAC, dimmers, AV
- **Live input values** displayed in real-time for sensors and digital inputs
- **Test workflow** — mark each item OK / Not OK / Skip with auto-advance to the next item
- **Ignore items** — exclude specific I/O points from the report
- **Session save/load** — resume an interrupted commissioning session (JSON file)
- **PDF report** — colour-coded (green / red / grey), grouped by room and category
- **Partial reports** — filter by room, section, or status before exporting
- **Company branding** — add your company name and sub-title to the GUI header and PDF
- **Offline mode** — load a local `LoxAPP3.json` file without a live Miniserver connection
- **Standalone EXE** — single-file Windows executable, no Python installation required

---



## Quick start

1. Launch the tool (EXE or `py loxone_checklist_gui.py`)
2. **Connect tab** → enter Miniserver IP, username, and password → click **Connect**
3. **Test I/O tab** → select any item in the tree to trigger outputs or read input values
4. Mark each item **OK**, **Not OK**, or **Skip** — the list advances automatically
5. **Report tab** → filter if needed → click **Generate PDF**

---

## Session files

The tool can save and restore a commissioning session:

- **Save session** — writes a `.json` file with all test results so far
- **Load session** — restores results; continue where you left off on a different day or device

---

## Supported Loxone control types

**Outputs tested/triggered:**
`Switch`, `Dimmer`, `Jalousie`, `Gate`, `Garage Door`, `Light Controller V2`, `HVAC`, `Air Conditioner`, `Push Button`, `Analog Output`, `Digital Output`, and more.

**Inputs read live:**
`Digital Sensor`, `Analog Sensor`, `Presence`, `Smoke`, `Wind`, `Brightness`, `Rain`, `NFC/Code Touch`, `Motion Detector`, and all AIR/Tree wireless variants.

---

## Notes

- The tool connects over HTTP (port 80) and accepts self-signed HTTPS certificates without warning — suitable for local LAN use only.
- Config (IP, username, company name) is saved to `loxone_tool_config.json` next to the executable for convenience. Do not commit this file if it contains credentials.
- Tested against Loxone Miniserver Gen 1 and Gen 2 firmware.

---

## License

MIT — free to use, modify, and distribute.
Pull requests welcome.
