#!/usr/bin/env python3
"""
Loxone I/O Checklist — Web App (Android / Mobile friendly)

Run:   py loxone_webapp.py
Then open on any device on your network:  http://<your-pc-ip>:5000
"""

import json
import io
import socket
import sys
import threading
from flask import Flask, request, jsonify, send_file, render_template_string

try:
    import requests as _requests
    from requests.auth import HTTPBasicAuth
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except ImportError:
    pass

# ── Import core logic ─────────────────────────────────────────────────────────
try:
    from loxone_checklist import (
        fetch_structure, load_local_structure,
        parse_structure, generate_pdf,
    )
except ImportError:
    print("ERROR: loxone_checklist.py must be in the same folder.")
    sys.exit(1)

app = Flask(__name__)

# ── Server-side connection store (single-user local tool) ────────────────────
_conn = {}   # host, username, password, use_https

# ── Control commands per Loxone type ─────────────────────────────────────────
# Format: type → list of (label, command_string)
CONTROL_CMDS = {
    "Switch":            [("On","on"),  ("Off","off"),  ("Pulse","pulse")],
    "TimedSwitch":       [("On","on"),  ("Off","off"),  ("Pulse","pulse")],
    "Pushbutton":        [("Press","pulse")],
    "Stairwell":         [("On","on"),  ("Off","off")],
    "Dimmer":            [("On","on"),  ("Off","off"),  ("50 %","50"),  ("100 %","100")],
    "EIBDimmer":         [("On","on"),  ("Off","off"),  ("50 %","50")],
    "ColorpickerV2":     [("On","on"),  ("Off","off")],
    "Colorpicker":       [("On","on"),  ("Off","off")],
    "Jalousie":          [("Up","up"),  ("Down","down"), ("Stop","stop"),
                          ("Full Up","fullup"), ("Full Down","fulldown")],
    "Gate":              [("Open","open"), ("Close","close"), ("Stop","stop")],
    "GarageDoor":        [("Open","open"), ("Close","close"), ("Stop","stop")],
    "UpDownDigital":     [("Up","up"),  ("Down","down"), ("Stop","stop")],
    "LightController":   [("On","on"),  ("Off","off")],
    "LightControllerV2": [("On","on"),  ("Off","off")],
    "CentralLightController": [("On","on"), ("Off","off")],
    "CentralJalousie":   [("Up","up"),  ("Down","down"), ("Stop","stop")],
    "DALI":              [("On","on"),  ("Off","off"),  ("50 %","50")],
    "FanController":     [("On","on"),  ("Off","off"),  ("Speed 1","1"),
                          ("Speed 2","2"), ("Speed 3","3")],
    "IRoomControllerV2": [("Comfort","comfort"), ("Economy","economy"), ("Off","off")],
    "IRoomController":   [("Comfort","comfort"), ("Economy","economy"), ("Off","off")],
    "HeatMixer":         [("On","on"),  ("Off","off")],
    "CoolingMixer":      [("On","on"),  ("Off","off")],
    "Ventilation":       [("On","on"),  ("Off","off"),  ("Auto","auto")],
    "AudioZone":         [("Play","play"), ("Pause","pause"), ("Vol+","volup"),
                          ("Vol−","voldown")],
    "AudioZoneV2":       [("Play","play"), ("Pause","pause")],
    "Alarm":             [("Arm","arm"), ("Disarm","disarm"), ("Quit","quit")],
    "Pool":              [("On","on"),  ("Off","off")],
    "Sauna":             [("On","on"),  ("Off","off")],
    "SolarPump":         [("On","on"),  ("Off","off")],
    "Intercom":          [("Open","open")],
    "MailBox":           [("Reset","reset")],
}


# ── HTML template ─────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta name="theme-color" content="#1a3a5c">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<title>Loxone Checklist</title>
<style>
  :root {
    --green:   #6EBD44;
    --dark:    #1a3a5c;
    --mid:     #2e6da4;
    --light:   #eaf4fb;
    --red:     #c0392b;
    --ok:      #1e8449;
    --bg:      #f0f2f5;
    --card:    #ffffff;
    --border:  #dde1e7;
    --text:    #222222;
    --muted:   #666666;
    --radius:  10px;
    --shadow:  0 2px 8px rgba(0,0,0,.10);
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg);
         color: var(--text); font-size: 15px; padding-bottom: 80px; }

  /* Header */
  header { background: var(--dark); color: #fff; padding: 14px 16px;
           display: flex; align-items: center; gap: 10px;
           position: sticky; top: 0; z-index: 100; box-shadow: var(--shadow); }
  header svg { flex-shrink: 0; }
  header h1 { font-size: 17px; font-weight: 700; }
  header small { font-size: 11px; color: #aac; margin-left: auto; }

  /* Progress bar */
  #progress-bar { height: 4px; background: var(--border); position: sticky;
                  top: 52px; z-index: 99; }
  #progress-fill { height: 100%; background: var(--green); width: 0%;
                   transition: width .3s; }

  /* Main */
  main { max-width: 700px; margin: 0 auto; padding: 14px 12px; }

  /* Cards */
  .card { background: var(--card); border-radius: var(--radius);
          box-shadow: var(--shadow); padding: 16px; margin-bottom: 14px; }
  .card h2 { font-size: 14px; font-weight: 700; color: var(--dark);
             margin-bottom: 12px; display: flex; align-items: center; gap: 6px; }

  /* Form */
  label { display: block; font-size: 12px; color: var(--muted);
          margin-bottom: 3px; margin-top: 10px; }
  label:first-child { margin-top: 0; }
  input[type=text], input[type=password] {
    width: 100%; padding: 10px 12px; border: 1.5px solid var(--border);
    border-radius: 8px; font-size: 15px; outline: none; transition: border-color .2s; }
  input:focus { border-color: var(--mid); }
  .row { display: flex; gap: 8px; }
  .row > * { flex: 1; }

  /* Toggle pills */
  .toggle-row { display: flex; gap: 6px; margin: 10px 0 2px; flex-wrap: wrap; }
  .toggle-btn { padding: 6px 12px; border-radius: 20px; border: 1.5px solid var(--border);
                background: #fff; font-size: 13px; cursor: pointer; transition: all .15s; }
  .toggle-btn.active { background: var(--dark); color: #fff; border-color: var(--dark); }

  /* Buttons */
  .btn { display: block; width: 100%; padding: 13px; border-radius: 8px;
         font-size: 15px; font-weight: 700; border: none; cursor: pointer;
         transition: opacity .15s; text-align: center; }
  .btn:active { opacity: .8; }
  .btn-green  { background: var(--green); color: #fff; }
  .btn-blue   { background: var(--mid);   color: #fff; }
  .btn-outline{ background: #fff; color: var(--dark); border: 1.5px solid var(--border); }
  .btn + .btn { margin-top: 8px; }
  .btn:disabled { opacity: .45; cursor: not-allowed; }

  /* Stats chips */
  .chips { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 4px; }
  .chip { padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
  .chip-in  { background: #eafaf1; color: #1e8449; }
  .chip-out { background: #eaf4fb; color: #1a5276; }
  .chip-oth { background: #fef9e7; color: #7d6608; }

  /* Section headers */
  .section-hdr { display: flex; align-items: center; justify-content: space-between;
                 padding: 10px 14px; border-radius: 8px; margin: 18px 0 6px;
                 font-weight: 700; font-size: 14px; color: #fff; cursor: pointer;
                 user-select: none; }
  .section-hdr .badge { background: rgba(255,255,255,.25); padding: 2px 8px;
                         border-radius: 10px; font-size: 12px; }
  .section-hdr .arrow { font-size: 12px; transition: transform .2s; }
  .section-hdr.collapsed .arrow { transform: rotate(-90deg); }
  .section-in  { background: #1e8449; }
  .section-out { background: #1a5276; }
  .section-oth { background: #7d6608; }

  /* Room group */
  .room-group { margin-bottom: 6px; }
  .room-label { font-size: 12px; font-weight: 700; color: var(--muted);
                text-transform: uppercase; letter-spacing: .5px;
                padding: 6px 0 3px; border-bottom: 1px solid var(--border);
                margin-bottom: 4px; }

  /* Check items */
  .check-item { display: flex; align-items: flex-start; gap: 12px;
                padding: 10px 0; border-bottom: 1px solid #f0f0f0; }
  .check-item:last-child { border-bottom: none; }
  .check-item input[type=checkbox] {
    width: 22px; height: 22px; flex-shrink: 0; cursor: pointer;
    accent-color: var(--green); margin-top: 2px; }
  .check-item.done .item-name { text-decoration: line-through; color: var(--muted); }
  .item-info { flex: 1; min-width: 0; }
  .item-name { font-size: 14px; font-weight: 500; }
  .item-meta { font-size: 12px; color: var(--muted); margin-top: 2px; }

  /* Control buttons */
  .ctrl-row { display: flex; flex-wrap: wrap; gap: 5px; margin-top: 7px; }
  .ctrl-btn {
    padding: 5px 11px; border-radius: 6px; font-size: 12px; font-weight: 600;
    border: 1.5px solid var(--mid); background: #fff; color: var(--mid);
    cursor: pointer; transition: all .12s; white-space: nowrap;
  }
  .ctrl-btn:active { background: var(--mid); color: #fff; }
  .ctrl-btn.sending { opacity: .5; pointer-events: none; }
  .ctrl-btn.ok  { border-color: var(--ok);  color: var(--ok);  background: #eafaf1; }
  .ctrl-btn.err { border-color: var(--red); color: var(--red); background: #fdf0ef; }

  /* Dimmer slider */
  .dimmer-row { display: flex; align-items: center; gap: 8px; margin-top: 6px; }
  .dimmer-row input[type=range] { flex: 1; accent-color: var(--mid); }
  .dimmer-val { font-size: 12px; color: var(--muted); width: 32px; text-align: right; }
  .dimmer-send { padding: 4px 10px; border-radius: 6px; font-size: 12px;
                 border: 1.5px solid var(--mid); background: #fff; color: var(--mid);
                 cursor: pointer; }

  /* Note */
  .item-notes { margin-top: 5px; width: 100%; padding: 5px 8px;
                border: 1px solid var(--border); border-radius: 6px;
                font-size: 13px; display: none; resize: none; height: 60px; }
  .item-notes.visible { display: block; }
  .note-btn { font-size: 11px; color: var(--mid); cursor: pointer;
              background: none; border: none; padding: 2px 0; margin-top: 2px; }

  /* Test mode banner */
  #test-banner { background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px;
                 padding: 8px 14px; font-size: 13px; color: #856404;
                 display: none; margin-bottom: 10px; text-align: center; }
  #test-banner.show { display: block; }

  /* Toast */
  #toast { position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
           background: #333; color: #fff; padding: 10px 20px; border-radius: 24px;
           font-size: 14px; opacity: 0; transition: opacity .3s; pointer-events: none;
           z-index: 999; white-space: nowrap; max-width: 90vw; text-align: center; }
  #toast.show { opacity: 1; }
  #toast.err   { background: var(--red); }
  #toast.ok    { background: var(--ok); }

  /* Spinner */
  #overlay { position: fixed; inset: 0; background: rgba(0,0,0,.45);
             display: none; align-items: center; justify-content: center; z-index: 200; }
  #overlay.show { display: flex; }
  .spinner { width: 48px; height: 48px; border: 5px solid rgba(255,255,255,.3);
             border-top-color: #fff; border-radius: 50%;
             animation: spin .8s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* FAB */
  #fab { position: fixed; bottom: 16px; right: 16px; z-index: 50;
         background: var(--green); color: #fff; border: none; border-radius: 28px;
         padding: 14px 22px; font-size: 15px; font-weight: 700;
         box-shadow: 0 4px 14px rgba(0,0,0,.25); cursor: pointer; display: none; }
  #fab:active { opacity: .85; }
</style>
</head>
<body>

<header>
  <svg width="28" height="28" viewBox="0 0 40 40" fill="none">
    <rect width="40" height="40" rx="8" fill="#6EBD44"/>
    <path d="M10 20 L17 27 L30 13" stroke="white" stroke-width="4"
          stroke-linecap="round" stroke-linejoin="round"/>
  </svg>
  <h1>Loxone Checklist</h1>
  <small id="proj-label"></small>
</header>

<div id="progress-bar"><div id="progress-fill"></div></div>

<main>

<!-- Connection card -->
<div class="card" id="conn-card">
  <h2>
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
         stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M5 12.55a11 11 0 0 1 14.08 0"/>
      <path d="M1.42 9a16 16 0 0 1 21.16 0"/>
      <path d="M8.53 16.11a6 6 0 0 1 6.95 0"/>
      <line x1="12" y1="20" x2="12.01" y2="20"/>
    </svg>
    Miniserver Connection
  </h2>

  <div class="toggle-row" id="src-toggle">
    <button class="toggle-btn active" onclick="setSrc('live',this)">Live Miniserver</button>
    <button class="toggle-btn" onclick="setSrc('file',this)">Paste JSON</button>
  </div>

  <div id="live-fields">
    <label>IP Address / Hostname</label>
    <input type="text" id="host" placeholder="e.g. 192.168.1.100" autocomplete="off">
    <div class="row">
      <div>
        <label>Username</label>
        <input type="text" id="user" placeholder="admin" value="admin">
      </div>
      <div>
        <label>Password</label>
        <input type="password" id="pass" placeholder="••••••">
      </div>
    </div>
    <div class="toggle-row" style="margin-top:8px">
      <button class="toggle-btn" id="btn-https" onclick="toggleFlag('https',this)">HTTPS</button>
      <button class="toggle-btn" id="btn-token" onclick="toggleFlag('token',this)">Token auth</button>
    </div>
  </div>

  <div id="file-fields" style="display:none">
    <label>Paste LoxAPP3.json contents here</label>
    <textarea id="json-paste" rows="5"
      style="width:100%;border:1.5px solid var(--border);border-radius:8px;
             padding:8px;font-size:12px;font-family:monospace;resize:vertical"
      placeholder='{ "msInfo": { ... }, "controls": { ... }, ... }'></textarea>
  </div>

  <button class="btn btn-green" style="margin-top:12px" onclick="doConnect()">
    Connect &amp; Load Structure
  </button>
</div>

<!-- Stats card -->
<div class="card" id="stats-card" style="display:none">
  <h2>Project Overview</h2>
  <div id="proj-info" style="font-size:13px;color:var(--muted);line-height:1.7"></div>
  <div class="chips" id="chips" style="margin-top:8px"></div>
  <div style="margin-top:12px;display:flex;gap:16px;flex-wrap:wrap;align-items:center">
    <label style="margin:0;display:flex;align-items:center;gap:6px;font-size:13px">
      <input type="checkbox" id="chk-other" style="width:16px;height:16px">
      Show other control types
    </label>
    <label style="margin:0;display:flex;align-items:center;gap:6px;font-size:13px;
                  cursor:pointer;color:var(--mid);font-weight:600">
      <input type="checkbox" id="chk-test" style="width:16px;height:16px;accent-color:var(--mid)"
             onchange="toggleTestMode()">
      ⚡ Test Mode (trigger outputs)
    </label>
  </div>
</div>

<!-- Test mode warning -->
<div id="test-banner" class="show" style="display:none">
  ⚡ <b>Test Mode active</b> — buttons send live commands to the Miniserver
</div>

<!-- Checklist -->
<div id="checklist"></div>

</main>

<!-- FAB -->
<button id="fab" onclick="generatePdf()">⬇ Download PDF</button>

<!-- Spinner -->
<div id="overlay"><div class="spinner"></div></div>

<!-- Toast -->
<div id="toast"></div>

<script>
// ── State ─────────────────────────────────────────────────────────────────────
let _data      = null;
let _checked   = {};
let _notes     = {};
let _src       = 'live';
let _flags     = { https: false, token: false };
let _testMode  = false;
// Control commands injected from server
const CTRL     = JSON_CTRL_PLACEHOLDER;

// ── Source / flag toggles ─────────────────────────────────────────────────────
function setSrc(val, btn) {
  _src = val;
  document.querySelectorAll('#src-toggle .toggle-btn')
          .forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('live-fields').style.display = val==='live' ? '' : 'none';
  document.getElementById('file-fields').style.display = val==='file' ? '' : 'none';
}
function toggleFlag(flag, btn) {
  _flags[flag] = !_flags[flag];
  btn.classList.toggle('active', _flags[flag]);
}
function toggleTestMode() {
  _testMode = document.getElementById('chk-test').checked;
  document.getElementById('test-banner').style.display = _testMode ? 'block' : 'none';
  // Show/hide all ctrl-rows
  document.querySelectorAll('.ctrl-row, .dimmer-row')
          .forEach(el => el.style.display = _testMode ? 'flex' : 'none');
}

// ── Connect ───────────────────────────────────────────────────────────────────
async function doConnect() {
  showSpinner(true);
  try {
    let body;
    if (_src === 'live') {
      body = {
        source:   'live',
        host:     document.getElementById('host').value.trim(),
        username: document.getElementById('user').value.trim(),
        password: document.getElementById('pass').value,
        https:    _flags.https,
        token:    _flags.token,
      };
    } else {
      body = { source: 'json', json: document.getElementById('json-paste').value.trim() };
    }
    const res  = await fetch('/api/connect', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');
    _data = data;
    _checked = {};
    _notes   = {};
    renderStats();
    renderChecklist();
    document.getElementById('fab').style.display = 'block';
    toast('Loaded — ' + (data.inputs.length + data.outputs.length) + ' I/O items', 'ok');
  } catch(e) {
    toast('Error: ' + e.message, 'err');
  } finally {
    showSpinner(false);
  }
}

// ── Render stats ──────────────────────────────────────────────────────────────
function renderStats() {
  const mi = _data.ms_info;
  const lines = [
    mi.projectName ? '<b>Project:</b> '    + mi.projectName : '',
    mi.msName      ? '<b>Miniserver:</b> ' + mi.msName      : '',
    mi.serialNr    ? '<b>Serial:</b> '     + mi.serialNr    : '',
    mi.swVersion   ? '<b>Firmware:</b> '   + mi.swVersion   : '',
  ].filter(Boolean).join('<br>');
  document.getElementById('proj-info').innerHTML = lines;
  document.getElementById('proj-label').textContent = mi.projectName || '';
  const chips = document.getElementById('chips');
  chips.innerHTML =
    '<span class="chip chip-in">'  + _data.inputs.length  + ' Inputs</span>'  +
    '<span class="chip chip-out">' + _data.outputs.length + ' Outputs</span>' +
    (_data.others.length ? '<span class="chip chip-oth">' + _data.others.length + ' Other</span>' : '');
  document.getElementById('stats-card').style.display = '';
  document.getElementById('chk-other').onchange = () => renderChecklist();
}

// ── Render checklist ──────────────────────────────────────────────────────────
function renderChecklist() {
  const showOther = document.getElementById('chk-other')?.checked;
  const sections = [
    { items: _data.inputs,  cls: 'section-in',  label: 'INPUTS — Sensors & Digital Inputs', isOutput: false },
    { items: _data.outputs, cls: 'section-out', label: 'OUTPUTS — Actors & Controls',       isOutput: true  },
  ];
  if (showOther && _data.others.length)
    sections.push({ items: _data.others, cls: 'section-oth', label: 'OTHER CONTROLS', isOutput: false });

  const container = document.getElementById('checklist');
  container.innerHTML = '';

  sections.forEach(sec => {
    if (!sec.items.length) return;

    const hdr = document.createElement('div');
    hdr.className = 'section-hdr ' + sec.cls;
    hdr.innerHTML = `<span>${sec.label}</span>
      <span style="display:flex;gap:8px;align-items:center">
        <span class="badge">${sec.items.length}</span>
        <span class="arrow">▼</span>
      </span>`;
    const body = document.createElement('div');
    hdr.onclick = () => {
      hdr.classList.toggle('collapsed');
      body.style.display = hdr.classList.contains('collapsed') ? 'none' : '';
    };
    container.appendChild(hdr);

    // Group by room
    const byRoom = {};
    sec.items.forEach(item => { (byRoom[item.room] = byRoom[item.room]||[]).push(item); });

    Object.keys(byRoom).sort().forEach(room => {
      const grp = document.createElement('div');
      grp.className = 'room-group card';
      grp.innerHTML = '<div class="room-label">' + esc(room) + '</div>';

      byRoom[room].forEach(item => {
        const cmds = sec.isOutput ? (CTRL[item.type] || []) : [];
        const hasDimmer = ['Dimmer','EIBDimmer','DALI'].includes(item.type);

        const row = document.createElement('div');
        row.className = 'check-item' + (_checked[item.uuid] ? ' done' : '');
        row.id = 'row-' + item.uuid;

        // Control buttons HTML
        let ctrlHtml = '';
        if (cmds.length) {
          const btnHtml = cmds.map(([lbl, cmd]) =>
            `<button class="ctrl-btn" data-uuid="${item.uuid}" data-cmd="${cmd}"
               onclick="sendCmd('${item.uuid}','${item.type}','${cmd}',this)">${esc(lbl)}</button>`
          ).join('');
          ctrlHtml += `<div class="ctrl-row" style="display:none">${btnHtml}</div>`;
        }
        if (hasDimmer) {
          ctrlHtml += `
            <div class="dimmer-row" style="display:none">
              <span style="font-size:12px;color:var(--muted)">Dim:</span>
              <input type="range" min="0" max="100" value="50" id="dim-${item.uuid}"
                oninput="document.getElementById('dimval-${item.uuid}').textContent=this.value+'%'">
              <span class="dimmer-val" id="dimval-${item.uuid}">50%</span>
              <button class="dimmer-send"
                onclick="sendDim('${item.uuid}',this)">Set</button>
            </div>`;
        }

        row.innerHTML = `
          <input type="checkbox" ${_checked[item.uuid]?'checked':''}
                 onchange="toggle('${item.uuid}',this)">
          <div class="item-info">
            <div class="item-name">${esc(item.name)}</div>
            <div class="item-meta">${esc(item.type_label)}${item.bus ? ' <span style="background:#e8edf5;color:#2e6da4;border-radius:4px;padding:1px 5px;font-size:10px;font-weight:700">'+esc(item.bus)+'</span>' : ''} &nbsp;·&nbsp; ${esc(item.category)}</div>
            ${ctrlHtml}
            <button class="note-btn" onclick="toggleNote('${item.uuid}')">
              ${_notes[item.uuid] ? '📝 Edit note' : '+ Add note'}
            </button>
            <textarea class="item-notes ${_notes[item.uuid]?'visible':''}"
              id="note-${item.uuid}" placeholder="Notes / observations…"
              onchange="saveNote('${item.uuid}',this.value)">${esc(_notes[item.uuid]||'')}</textarea>
          </div>`;
        grp.appendChild(row);
      });
      body.appendChild(grp);
    });

    container.appendChild(body);
  });
  updateProgress();
  // Restore test mode visibility
  if (_testMode) toggleTestMode();
}

// ── Send command ──────────────────────────────────────────────────────────────
async function sendCmd(uuid, type, cmd, btn) {
  btn.classList.add('sending');
  try {
    const res = await fetch('/api/control', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ uuid, cmd }),
    });
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || 'Failed');
    btn.classList.remove('sending');
    btn.classList.add('ok');
    toast(`✓ ${data.name || uuid}  →  ${cmd}`, 'ok');
    setTimeout(() => btn.classList.remove('ok'), 1800);
  } catch(e) {
    btn.classList.remove('sending');
    btn.classList.add('err');
    toast('Error: ' + e.message, 'err');
    setTimeout(() => btn.classList.remove('err'), 2500);
  }
}
async function sendDim(uuid, btn) {
  const val = document.getElementById('dim-' + uuid)?.value || '50';
  // reuse the same btn styling trick
  await sendCmd(uuid, '', val, btn);
}

// ── Check / Notes ─────────────────────────────────────────────────────────────
function toggle(uuid, cb) {
  _checked[uuid] = cb.checked;
  document.getElementById('row-'+uuid)?.classList.toggle('done', cb.checked);
  updateProgress();
}
function toggleNote(uuid) {
  const ta = document.getElementById('note-'+uuid);
  ta.classList.toggle('visible');
  if (ta.classList.contains('visible')) ta.focus();
}
function saveNote(uuid, val) { _notes[uuid] = val.trim(); }
function updateProgress() {
  if (!_data) return;
  const all  = [..._data.inputs, ..._data.outputs, ..._data.others];
  const done = all.filter(i => _checked[i.uuid]).length;
  document.getElementById('progress-fill').style.width =
    (all.length ? Math.round(done/all.length*100) : 0) + '%';
}

// ── Generate PDF ──────────────────────────────────────────────────────────────
async function generatePdf() {
  if (!_data) return;
  showSpinner(true);
  try {
    const res = await fetch('/api/generate_pdf', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ms_info: _data.ms_info,
        inputs:  _data.inputs,
        outputs: _data.outputs,
        others:  _data.others,
        checked: _checked,
        notes:   _notes,
        include_other: document.getElementById('chk-other')?.checked || false,
      }),
    });
    if (!res.ok) { const e = await res.json(); throw new Error(e.error); }
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url; a.download = 'loxone_checklist.pdf'; a.click();
    URL.revokeObjectURL(url);
    toast('PDF downloaded!', 'ok');
  } catch(e) {
    toast('PDF error: ' + e.message, 'err');
  } finally { showSpinner(false); }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function showSpinner(v) { document.getElementById('overlay').classList.toggle('show',v); }
let _toastTimer;
function toast(msg, type='') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'show' + (type ? ' '+type : '');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.className = '', 2800);
}
function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
                  .replace(/"/g,'&quot;');
}
</script>
</body>
</html>
"""


# ── API routes ────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    # Inject control commands as JSON into the page
    ctrl_json = json.dumps(
        {k: [[lbl, cmd] for lbl, cmd in v] for k, v in CONTROL_CMDS.items()}
    )
    html = HTML.replace("JSON_CTRL_PLACEHOLDER", ctrl_json)
    return render_template_string(html)


@app.route("/api/connect", methods=["POST"])
def api_connect():
    data   = request.get_json(force=True)
    source = data.get("source", "live")
    try:
        if source == "json":
            structure = json.loads(data.get("json", "").strip())
            _conn.clear()   # no live connection for offline JSON
        else:
            host      = data.get("host", "").strip()
            username  = data.get("username", "admin")
            password  = data.get("password", "")
            use_https = data.get("https", False)
            use_token = data.get("token",  False)
            if not host:
                return jsonify(error="No host provided"), 400
            structure = fetch_structure(host, username, password,
                                        use_https=use_https, use_token=use_token)
            # Store credentials for control commands
            _conn.update(host=host, username=username, password=password,
                         use_https=use_https)

        ms_info = structure.get("msInfo", {})
        inputs, outputs, others = parse_structure(structure)
        return jsonify(ms_info=ms_info, inputs=inputs, outputs=outputs, others=others)

    except SystemExit:
        return jsonify(error="Cannot connect to Miniserver. Check IP and credentials."), 502
    except (json.JSONDecodeError, ValueError) as e:
        return jsonify(error=f"Invalid JSON: {e}"), 400
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/api/control", methods=["POST"])
def api_control():
    """Proxy a control command to the Miniserver."""
    if not _conn:
        return jsonify(error="No live Miniserver connection. Reconnect first."), 400

    data = request.get_json(force=True)
    uuid = data.get("uuid", "").strip()
    cmd  = data.get("cmd",  "").strip()

    if not uuid or not cmd:
        return jsonify(error="uuid and cmd are required"), 400

    proto    = "https" if _conn.get("use_https") else "http"
    host     = _conn["host"]
    username = _conn["username"]
    password = _conn["password"]

    # Loxone command endpoint: /jdev/sps/io/<uuid>/<command>
    url = f"{proto}://{host}/jdev/sps/io/{uuid}/{cmd}"

    try:
        r = _requests.get(
            url,
            auth=HTTPBasicAuth(username, password),
            timeout=8,
            verify=False,
        )
        r.raise_for_status()
        resp = r.json()
        # Loxone wraps response in {"LL": {"Code": "200", "value": ...}}
        ll   = resp.get("LL", {})
        code = str(ll.get("Code", "200"))
        if code not in ("200", "0"):
            return jsonify(error=f"Miniserver returned code {code}"), 502
        return jsonify(ok=True, name=ll.get("control", uuid), value=ll.get("value"))

    except _requests.exceptions.ConnectionError:
        return jsonify(error="Cannot reach Miniserver"), 502
    except _requests.exceptions.Timeout:
        return jsonify(error="Miniserver timed out"), 504
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/api/generate_pdf", methods=["POST"])
def api_generate_pdf():
    data          = request.get_json(force=True)
    ms_info       = data.get("ms_info",       {})
    inputs        = data.get("inputs",        [])
    outputs       = data.get("outputs",       [])
    others        = data.get("others",        [])
    checked       = data.get("checked",       {})
    notes         = data.get("notes",         {})
    include_other = data.get("include_other", False)

    def _annotate(items):
        return [{**item, "checked": checked.get(item["uuid"], False),
                          "notes":   notes.get(item["uuid"],   "")}
                for item in items]

    try:
        buf = io.BytesIO()
        _patched_generate_pdf(
            _annotate(inputs), _annotate(outputs), _annotate(others),
            ms_info, buf, include_other=include_other
        )
        buf.seek(0)
        return send_file(buf, mimetype="application/pdf",
                         as_attachment=True, download_name="loxone_checklist.pdf")
    except Exception as e:
        return jsonify(error=str(e)), 500


# ── PDF helper (accepts BytesIO) ──────────────────────────────────────────────

def _patched_generate_pdf(inputs, outputs, others, ms_info,
                           output_path, include_other=False):
    from loxone_checklist import (
        _make_styles, _section_banner, _add_section,
        INPUT_GREEN, INPUT_BG, OUTPUT_BLUE, OUTPUT_BG,
        OTHER_AMBER, OTHER_BG, DARK_BLUE, LOXONE_GREEN, GRID_COLOR,
    )
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph,
        Spacer, HRFlowable, KeepTogether,
    )
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors as _c
    from collections import defaultdict
    from datetime import datetime

    if not hasattr(output_path, "write"):
        generate_pdf(inputs, outputs, others, ms_info,
                     output_path=output_path, include_other=include_other)
        return

    styles   = _make_styles()
    story    = []
    project  = ms_info.get("projectName", "Loxone Project")
    ms_name  = ms_info.get("msName",    "")
    serial   = ms_info.get("serialNr",  "")
    firmware = ms_info.get("swVersion", "")
    generated = datetime.now().strftime("%Y-%m-%d  %H:%M")

    story.append(Paragraph("Loxone I/O Commissioning Checklist", styles["title"]))
    story.append(Spacer(1, 1*mm))
    story.append(HRFlowable(width="100%", thickness=3, color=LOXONE_GREEN))
    story.append(Spacer(1, 2*mm))
    for line in filter(None, [
        f"<b>Project:</b> {project}",
        f"<b>Miniserver:</b> {ms_name}" if ms_name else "",
        f"<b>Serial:</b> {serial}"      if serial  else "",
        f"<b>Firmware:</b> {firmware}"  if firmware else "",
        f"<b>Generated:</b> {generated}",
    ]):
        story.append(Paragraph(line, styles["sub"]))

    story.append(Spacer(1, 4*mm))
    total = len(inputs) + len(outputs) + (len(others) if include_other else 0)
    summary_data = [["Section","Items"],
                    ["Inputs (sensors)", str(len(inputs))],
                    ["Outputs (actors)", str(len(outputs))]]
    if include_other:
        summary_data.append(["Other controls", str(len(others))])
    summary_data.append(["TOTAL", str(total)])
    summary = Table(summary_data, colWidths=[80*mm, 30*mm])
    summary.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,0), DARK_BLUE),
        ("TEXTCOLOR",  (0,0),(-1,0), _c.white),
        ("FONTNAME",   (0,0),(-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0),(-1,-1), 9),
        ("BACKGROUND", (0,-1),(-1,-1), _c.HexColor("#d6eaf8")),
        ("FONTNAME",   (0,-1),(-1,-1), "Helvetica-Bold"),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[_c.white,_c.HexColor("#f5f5f5")]),
        ("GRID",       (0,0),(-1,-1), 0.5, GRID_COLOR),
        ("ALIGN",      (1,0),(1,-1), "CENTER"),
        ("TOPPADDING", (0,0),(-1,-1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
        ("LEFTPADDING",(0,0),(-1,-1), 6),
    ]))
    story.append(summary)
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(
        "<b>Instructions:</b>  Checked items are marked ✓. Notes included where provided.",
        styles["small"]))
    story.append(Spacer(1, 4*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=GRID_COLOR))

    def _add_sec(title, items, hdr_color, row_alt):
        if not items: return
        story.append(Spacer(1, 2*mm))
        story.append(_section_banner(f"{title}  ({len(items)} items)", hdr_color, styles))
        story.append(Spacer(1, 2*mm))
        by_room = defaultdict(list)
        for item in items: by_room[item["room"]].append(item)
        for room, room_items in sorted(by_room.items()):
            block = []
            block.append(Paragraph(f"<b>{room}</b>", styles["room"]))
            col_widths = [7*mm, 62*mm, 38*mm, 22*mm, 40*mm, 11*mm]
            rows = [["","Name","Type","Category","Notes","✓"]]
            for item in room_items:
                tick = "✓" if item.get("checked") else "\u25A1"
                rows.append([tick, item["name"], item["type_label"],
                              item["category"], item.get("notes",""), ""])
            t = Table(rows, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0), hdr_color),
                ("TEXTCOLOR", (0,0),(-1,0), _c.white),
                ("FONTNAME",  (0,0),(-1,0), "Helvetica-Bold"),
                ("FONTSIZE",  (0,0),(-1,-1), 8),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[_c.white, row_alt]),
                ("GRID",      (0,0),(-1,-1), 0.3, GRID_COLOR),
                ("VALIGN",    (0,0),(-1,-1), "MIDDLE"),
                ("TOPPADDING",(0,0),(-1,-1), 3),
                ("BOTTOMPADDING",(0,0),(-1,-1), 3),
                ("LEFTPADDING",(0,0),(-1,-1), 4),
                ("ALIGN",     (0,0),(0,-1), "CENTER"),
                ("FONTSIZE",  (0,1),(0,-1), 11),
            ]))
            block.append(t)
            block.append(Spacer(1, 3*mm))
            story.append(KeepTogether(block))

    _add_sec("INPUTS — Sensors & Digital Inputs",
             inputs, INPUT_GREEN, INPUT_BG)
    _add_sec("OUTPUTS — Actors & Controls",
             outputs, OUTPUT_BLUE, OUTPUT_BG)
    if include_other:
        _add_sec("OTHER CONTROLS", others, OTHER_AMBER, OTHER_BG)

    story.append(Spacer(1, 12*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=_c.grey))
    story.append(Spacer(1, 4*mm))
    ft = Table([["Technician:","_"*28,"Date:","_"*18,"Signature:","_"*22]],
               colWidths=[22*mm,52*mm,12*mm,32*mm,20*mm,38*mm])
    ft.setStyle(TableStyle([("FONTSIZE",(0,0),(-1,-1),9),
                             ("VALIGN",(0,0),(-1,-1),"BOTTOM"),
                             ("TOPPADDING",(0,0),(-1,-1),2),
                             ("BOTTOMPADDING",(0,0),(-1,-1),2)]))
    story.append(ft)

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=15*mm, leftMargin=15*mm,
                            topMargin=18*mm, bottomMargin=18*mm,
                            title="Loxone I/O Commissioning Checklist")
    doc.build(story)


# ── Launch ────────────────────────────────────────────────────────────────────

def _get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def main():
    ip, port = _get_local_ip(), 5000
    print("=" * 52)
    print("  Loxone I/O Checklist  —  Web App")
    print("=" * 52)
    print(f"  Desktop:  http://localhost:{port}")
    print(f"  Android:  http://{ip}:{port}")
    print("  Press  Ctrl+C  to stop")
    print("=" * 52)
    import webbrowser
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    app.run(host="0.0.0.0", port=port, debug=False)


if __name__ == "__main__":
    main()
