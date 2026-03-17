#!/usr/bin/env python3
"""
Loxone Miniserver I/O Commissioning Checklist Generator

Connects to a Loxone Miniserver, reads all configured inputs and outputs,
and generates a printable PDF checklist for on-site commissioning.

Usage:
    python loxone_checklist.py <host> <username> <password> [options]
    python loxone_checklist.py 192.168.1.100 admin secret
    python loxone_checklist.py 192.168.1.100 admin secret --output project.pdf
    python loxone_checklist.py --file LoxAPP3.json  (offline mode)
"""

import sys
import json
import argparse
import hashlib
import hmac
import binascii
from datetime import datetime
from collections import defaultdict

import requests
from requests.auth import HTTPBasicAuth
import urllib3

# Suppress SSL warnings for self-signed certs
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable, KeepTogether, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics


# ─────────────────────────────────────────────
#  Control type classification
# ─────────────────────────────────────────────

INPUT_TYPES = {
    # ── Native Loxone I/O ────────────────────────────────────────────────────
    "InfoOnlyDigital":   "Digital Sensor",
    "InfoOnlyAnalog":    "Analog Sensor",
    "InfoOnlyText":      "Text State",
    "Presence":          "Presence Sensor",
    "Smoke":             "Smoke Detector",
    "WindSpeed":         "Wind Speed Sensor",
    "Brightness":        "Brightness Sensor",
    "RainAlert":         "Rain Sensor",
    "Meter":             "Meter / Counter",
    "Pulse":             "Pulse Input",
    "FrequencyInput":    "Frequency Input",
    "NfcCodeTouch":      "NFC / Code Touch",
    "TextState":         "Text State",
    # ── Energy ───────────────────────────────────────────────────────────────
    "CarCharger":        "Car Charger",
    "Fronius":           "Inverter / Solar",
    "FroniusEco":        "Inverter Eco",
    "Wallbox":           "EV Wallbox",
    "EnergyManager":     "Energy Manager",
    "EnergyManager2":    "Energy Manager v2",
    # ── KNX / EIB ────────────────────────────────────────────────────────────
    "EIBInputStatus":    "KNX Digital Input",
    "EIBAnalogInput":    "KNX Analog Input",
    "EIBPresence":       "KNX Presence Sensor",
    "EIBSmoke":          "KNX Smoke Detector",
    "EIBMotion":         "KNX Motion Sensor",
    "EIBWindowContact":  "KNX Window Contact",
    # ── Modbus ───────────────────────────────────────────────────────────────
    "ModbusInput":       "Modbus Input",
    "ModbusAnalogInput": "Modbus Analog Input",
    # ── DMX / Art-Net ────────────────────────────────────────────────────────
    "DMXInput":          "DMX Input",
    # ── EnOcean ──────────────────────────────────────────────────────────────
    "EnOceanDigital":    "EnOcean Digital",
    "EnOceanAnalog":     "EnOcean Analog",
    "EnOceanPresence":   "EnOcean Presence",
    "EnOceanWindow":     "EnOcean Window",
    # ── 1-Wire ───────────────────────────────────────────────────────────────
    "OneWireInput":      "1-Wire Input",
    "OneWireTemperature":"1-Wire Temperature",
    # ── Modbus TCP / generic virtual ─────────────────────────────────────────
    "VirtualInputDigital": "Virtual Digital Input",
    "VirtualInputAnalog":  "Virtual Analog Input",
    "VirtualInputText":    "Virtual Text Input",
    # ── Loxone AIR ───────────────────────────────────────────────────────────
    "AirInfoOnlyDigital":  "AIR Digital Sensor",
    "AirInfoOnlyAnalog":   "AIR Analog Sensor",
    "AirPresence":         "AIR Presence Sensor",
    "AirSmoke":            "AIR Smoke Detector",
    "AirWindowContact":    "AIR Window Contact",
    "AirMotion":           "AIR Motion Sensor",
    # ── Loxone Tree ──────────────────────────────────────────────────────────
    "TreeInput":           "Tree Digital Input",
    "TreeAnalogInput":     "Tree Analog Input",
}

OUTPUT_TYPES = {
    # ── Native Loxone outputs ─────────────────────────────────────────────────
    "Switch":            "Switch / Relay",
    "TimedSwitch":       "Timed Switch",
    "Pushbutton":        "Pushbutton",
    "Dimmer":            "Dimmer",
    "ColorpickerV2":     "RGB Light",
    "Colorpicker":       "Color Picker",
    "Jalousie":          "Blind / Shutter",
    "Gate":              "Gate / Door",
    "GarageDoor":        "Garage Door",
    "UpDownDigital":     "Up/Down Digital",
    "UpDownAnalog":      "Up/Down Analog",
    "LightController":   "Light Controller",
    "LightControllerV2": "Light Controller V2",
    "DALI":              "DALI Lighting",
    "IRoomControllerV2": "Room Controller",
    "IRoomController":   "Room Controller",
    "HeatMixer":         "Heating Mixer",
    "CoolingMixer":      "Cooling Mixer",
    "SolarPump":         "Solar Pump",
    "Pool":              "Pool Control",
    "Sauna":             "Sauna Control",
    "Ventilation":       "Ventilation",
    "Alarm":             "Alarm",
    "AlarmClock":        "Alarm Clock",
    "AudioZone":         "Audio Zone",
    "AudioZoneV2":       "Audio Zone V2",
    "Intercom":          "Intercom",
    "MailBox":           "Mailbox",
    "Radio":             "Radio Button",
    "Remote":            "Remote Control",
    "ValueSelector":     "Value Selector",
    "Webpage":           "Webpage",
    "FanController":     "Fan Controller",
    "Stairwell":         "Stairwell Light",
    "CentralLightController": "Central Lighting",
    "CentralJalousie":   "Central Blinds",
    "CentralAlarm":      "Central Alarm",
    "CentralAudio":      "Central Audio",
    # ── Generic I/O (direct hardware binding) ────────────────────────────────
    "AnalogOutput":      "Analog Output",
    "DigitalOutput":     "Digital Output",
    # ── KNX / EIB ────────────────────────────────────────────────────────────
    "EIBDimmer":         "KNX Dimmer",
    "EIBSwitch":         "KNX Switch",
    "EIBJalousie":       "KNX Blind / Shutter",
    "EIBLightController":"KNX Light Controller",
    "EIBColorpicker":    "KNX RGB Light",
    "EIBGate":           "KNX Gate",
    "EIBFanCoil":        "KNX Fan Coil",
    "EIBHVACController": "KNX HVAC Controller",
    # ── Modbus ───────────────────────────────────────────────────────────────
    "ModbusOutput":      "Modbus Output",
    "ModbusAnalogOutput":"Modbus Analog Output",
    # ── DMX / Art-Net ────────────────────────────────────────────────────────
    "DMX":               "DMX Channel",
    "DMXOutput":         "DMX Output",
    "ArtNet":            "Art-Net Output",
    # ── EnOcean ──────────────────────────────────────────────────────────────
    "EnOceanActuator":   "EnOcean Actuator",
    "EnOceanDimmer":     "EnOcean Dimmer",
    "EnOceanJalousie":   "EnOcean Blind",
    # ── Virtual / HTTP outputs ────────────────────────────────────────────────
    "VirtualOutputDigital": "Virtual Digital Output",
    "VirtualOutputAnalog":  "Virtual Analog Output",
    "VirtualOutputText":    "Virtual Text Output",
    "HttpOutput":        "HTTP Output",
    # ── Loxone AIR ───────────────────────────────────────────────────────────
    "AirSwitch":         "AIR Switch",
    "AirDimmer":         "AIR Dimmer",
    "AirJalousie":       "AIR Blind / Shutter",
    # ── Loxone Tree ──────────────────────────────────────────────────────────
    "TreeSwitch":        "Tree Switch",
    "TreeDimmer":        "Tree Dimmer",
    "TreeJalousie":      "Tree Blind / Shutter",
}

# ── Bus / protocol tags (shown as badge in the PDF) ──────────────────────────
BUS_TAGS = {
    # KNX/EIB
    "EIBDimmer": "KNX", "EIBSwitch": "KNX", "EIBJalousie": "KNX",
    "EIBLightController": "KNX", "EIBColorpicker": "KNX", "EIBGate": "KNX",
    "EIBFanCoil": "KNX", "EIBHVACController": "KNX",
    "EIBInputStatus": "KNX", "EIBAnalogInput": "KNX", "EIBPresence": "KNX",
    "EIBSmoke": "KNX", "EIBMotion": "KNX", "EIBWindowContact": "KNX",
    # Modbus
    "ModbusInput": "Modbus", "ModbusAnalogInput": "Modbus",
    "ModbusOutput": "Modbus", "ModbusAnalogOutput": "Modbus",
    # DMX / Art-Net
    "DMX": "DMX", "DMXInput": "DMX", "DMXOutput": "DMX", "ArtNet": "Art-Net",
    # EnOcean
    "EnOceanDigital": "EnOcean", "EnOceanAnalog": "EnOcean",
    "EnOceanPresence": "EnOcean", "EnOceanWindow": "EnOcean",
    "EnOceanActuator": "EnOcean", "EnOceanDimmer": "EnOcean",
    "EnOceanJalousie": "EnOcean",
    # 1-Wire
    "OneWireInput": "1-Wire", "OneWireTemperature": "1-Wire",
    # Virtual
    "VirtualInputDigital": "Virtual", "VirtualInputAnalog": "Virtual",
    "VirtualInputText": "Virtual",
    "VirtualOutputDigital": "Virtual", "VirtualOutputAnalog": "Virtual",
    "VirtualOutputText": "Virtual",
    "HttpOutput": "HTTP",
    # DALI
    "DALI": "DALI",
    # AIR
    "NfcCodeTouch": "AIR", "AirSwitch": "AIR", "AirDimmer": "AIR",
    "AirJalousie": "AIR", "AirPresence": "AIR", "AirSmoke": "AIR",
    "AirInfoOnlyDigital": "AIR", "AirInfoOnlyAnalog": "AIR",
    "AirWindowContact": "AIR", "AirMotion": "AIR",
    # Tree
    "TreeSwitch": "Tree", "TreeDimmer": "Tree", "TreeJalousie": "Tree",
    "TreeInput": "Tree", "TreeAnalogInput": "Tree",
}

# Types to skip entirely (internal / virtual / system)
SKIP_TYPES = {
    "Folder", "Virtual", "Sauna", "Calendar",
}

# Physical hardware outputs — directly drive a load on site.
# Anything in OUTPUT_TYPES but NOT here is a logical controller
# and is routed to the "others" bucket instead.
PHYSICAL_OUTPUTS = {
    # ── Relays / switches ─────────────────────────────────────────────────────
    "Switch", "TimedSwitch", "Pushbutton", "Stairwell",
    # ── Dimmers ───────────────────────────────────────────────────────────────
    "Dimmer", "EIBDimmer", "DALI", "AirDimmer", "TreeDimmer",
    # ── Colour / RGB ─────────────────────────────────────────────────────────
    "ColorpickerV2", "Colorpicker", "EIBColorpicker",
    # ── Blinds / shutters / gates ─────────────────────────────────────────────
    "Jalousie", "EIBJalousie", "AirJalousie", "TreeJalousie",
    "Gate", "GarageDoor", "EIBGate",
    "UpDownDigital", "UpDownAnalog",
    # ── Generic I/O ──────────────────────────────────────────────────────────
    "AnalogOutput", "DigitalOutput",
    # ── HVAC hardware ─────────────────────────────────────────────────────────
    "HeatMixer", "CoolingMixer", "FanController",
    "EIBFanCoil", "EIBHVACController",
    "Ventilation",
    # ── Pumps / pools ─────────────────────────────────────────────────────────
    "SolarPump", "Pool",
    # ── Access control ────────────────────────────────────────────────────────
    "Intercom",
    # ── KNX physical ─────────────────────────────────────────────────────────
    "EIBSwitch",
    # ── AIR physical ─────────────────────────────────────────────────────────
    "AirSwitch",
    # ── Tree physical ─────────────────────────────────────────────────────────
    "TreeSwitch",
    # ── Fieldbus ──────────────────────────────────────────────────────────────
    "ModbusOutput", "ModbusAnalogOutput",
    "DMX", "DMXOutput", "ArtNet",
    "EnOceanActuator", "EnOceanDimmer", "EnOceanJalousie",
}

# ── Control commands per type (label, command_string) ────────────────────────
CONTROL_CMDS = {
    "Switch":            [("On","on"),   ("Off","off"),  ("Pulse","pulse")],
    "TimedSwitch":       [("On","on"),   ("Off","off"),  ("Pulse","pulse")],
    "Pushbutton":        [("Press","pulse")],
    "Stairwell":         [("On","on"),   ("Off","off")],
    "Dimmer":            [("On","on"),   ("Off","off"),  ("50%","50"),  ("100%","100")],
    "EIBDimmer":         [("On","on"),   ("Off","off"),  ("50%","50")],
    "ColorpickerV2":     [("On","on"),   ("Off","off")],
    "Colorpicker":       [("On","on"),   ("Off","off")],
    "Jalousie":          [("Up","up"),   ("Down","down"), ("Stop","stop"),
                          ("Full Up","fullup"), ("Full Down","fulldown")],
    "Gate":              [("Open","open"), ("Close","close"), ("Stop","stop")],
    "GarageDoor":        [("Open","open"), ("Close","close"), ("Stop","stop")],
    "UpDownDigital":     [("Up","up"),   ("Down","down"), ("Stop","stop")],
    "LightController":   [("On","on"),   ("Off","off")],
    "LightControllerV2": [("On","on"),   ("Off","off")],
    "CentralLightController": [("On","on"), ("Off","off")],
    "CentralJalousie":   [("Up","up"),   ("Down","down"), ("Stop","stop")],
    "DALI":              [("On","on"),   ("Off","off"),  ("50%","50")],
    "FanController":     [("On","on"),   ("Off","off"),
                          ("Spd 1","1"), ("Spd 2","2"),  ("Spd 3","3")],
    # IRoomControllerV2 uses numeric operating modes:
    # 0=Auto, 1=Comfort, 2=Standby, 3=Economy, 4=Building Protection
    "IRoomControllerV2": [("Auto","0"), ("Comfort","1"),
                          ("Standby","2"), ("Economy","3"), ("Protect","4")],
    "IRoomController":   [("Auto","0"), ("Comfort","1"),
                          ("Standby","2"), ("Economy","3")],
    "HeatMixer":         [("On","on"),   ("Off","off")],
    "CoolingMixer":      [("On","on"),   ("Off","off")],
    "Ventilation":       [("On","on"),   ("Off","off"),  ("Auto","auto")],
    "AudioZone":         [("Play","play"), ("Pause","pause"),
                          ("Vol+","volup"), ("Vol-","voldown")],
    "AudioZoneV2":       [("Play","play"), ("Pause","pause")],
    "Alarm":             [("Arm","arm"), ("Disarm","disarm"), ("Quit","quit")],
    "Pool":              [("On","on"),   ("Off","off")],
    "SolarPump":         [("On","on"),   ("Off","off")],
    "Intercom":          [("Open","open")],
    "MailBox":           [("Reset","reset")],
    "EIBSwitch":         [("On","on"),   ("Off","off"),  ("Pulse","pulse")],
    "EIBJalousie":       [("Up","up"),   ("Down","down"), ("Stop","stop")],
    "EIBGate":           [("Open","open"), ("Close","close"), ("Stop","stop")],
    "ModbusOutput":      [("On","on"),   ("Off","off")],
    "VirtualOutputDigital": [("On","on"), ("Off","off"), ("Pulse","pulse")],
    "AirSwitch":         [("On","on"),   ("Off","off"),  ("Pulse","pulse")],
    "AirDimmer":         [("On","on"),   ("Off","off"),  ("50%","50")],
    "AirJalousie":       [("Up","up"),   ("Down","down"), ("Stop","stop")],
    "TreeSwitch":        [("On","on"),   ("Off","off"),  ("Pulse","pulse")],
    "TreeDimmer":        [("On","on"),   ("Off","off"),  ("50%","50")],
    "TreeJalousie":      [("Up","up"),   ("Down","down"), ("Stop","stop")],
}

# Primary state name to poll for live value display (most important per type)
PRIMARY_STATE = {
    # Outputs
    "Switch":            "active",    "TimedSwitch":  "active",
    "Pushbutton":        "active",    "Stairwell":    "active",
    "Dimmer":            "value",     "EIBDimmer":    "value",
    "ColorpickerV2":     "color",     "Colorpicker":  "color",
    "Jalousie":          "position",  "EIBJalousie":  "position",
    "AirJalousie":       "position",  "TreeJalousie": "position",
    "Gate":              "active",    "GarageDoor":   "active",
    "UpDownDigital":     "active",    "UpDownAnalog": "value",
    "HeatMixer":         "value",     "CoolingMixer": "value",
    "FanController":     "value",     "Ventilation":  "v0",
    "SolarPump":         "active",    "Pool":         "active",
    "DALI":              "value",     "AirDimmer":    "value",
    "TreeDimmer":        "value",     "AirSwitch":    "active",
    "TreeSwitch":        "active",    "EIBSwitch":    "active",
    "ModbusOutput":      "active",    "ModbusAnalogOutput": "value",
    "DMX":               "value",     "EnOceanActuator": "active",
    "IRoomControllerV2": "operatingMode", "IRoomController": "operatingMode",
    "LightController":   "activeScene","LightControllerV2": "activeScene",
    "AudioZone":         "volume",    "AudioZoneV2":  "volume",
    "Alarm":             "armed",
    # Inputs
    "InfoOnlyDigital":   "active",    "InfoOnlyAnalog":    "value",
    "Presence":          "active",    "Smoke":             "active",
    "WindSpeed":         "value",     "Brightness":        "value",
    "RainAlert":         "active",    "Meter":             "value",
    "Pulse":             "value",     "FrequencyInput":    "value",
    "NfcCodeTouch":      "code",
    "EIBInputStatus":    "active",    "EIBAnalogInput":    "value",
    "EIBPresence":       "active",    "EIBSmoke":          "active",
    "EIBMotion":         "active",    "EIBWindowContact":  "active",
    "ModbusInput":       "value",     "ModbusAnalogInput": "value",
    "EnOceanDigital":    "active",    "EnOceanAnalog":     "value",
    "EnOceanPresence":   "active",    "EnOceanWindow":     "active",
    "OneWireTemperature":"value",     "OneWireInput":      "value",
    "AirInfoOnlyDigital":"active",    "AirInfoOnlyAnalog": "value",
    "AirPresence":       "active",    "AirSmoke":          "active",
    "AirWindowContact":  "active",    "AirMotion":         "active",
    "TreeInput":         "active",    "TreeAnalogInput":   "value",
    "CarCharger":        "power",     "Wallbox":           "power",
    "Fronius":           "current",   "EnergyManager":     "current",
}

# Types that support a numeric dim value (0-100 slider)
DIMMER_TYPES = {
    "Dimmer", "EIBDimmer", "DALI", "AirDimmer", "TreeDimmer",
    "ColorpickerV2", "Colorpicker", "VirtualOutputAnalog", "ModbusAnalogOutput",
}


# ─────────────────────────────────────────────
#  Token-based authentication
# ─────────────────────────────────────────────

def _get_token(session, base_url, username, password):
    """Acquire a Loxone API token (Config ≥ 9.0)."""
    # Step 1: get key
    r = session.get(f"{base_url}/jdev/sys/getkey", timeout=10, verify=False)
    r.raise_for_status()
    key_hex = r.json().get("LL", {}).get("value", "")
    key = binascii.unhexlify(key_hex)

    # Step 2: hash password
    pwd_hash = hashlib.sha1(f"{password}:{username}".encode()).hexdigest().upper()
    token_hash = hmac.new(key, pwd_hash.encode(), hashlib.sha1).hexdigest()

    # Step 3: acquire token
    app_uuid = "edfc5f9a-0c4e-4d56-9b09-a1b4a5de4bdc"
    r = session.get(
        f"{base_url}/jdev/sys/gettoken/{token_hash}/{username}/4/{app_uuid}/loxone_checklist",
        timeout=10, verify=False
    )
    r.raise_for_status()
    token = r.json().get("LL", {}).get("value", {}).get("token", "")
    return token


# ─────────────────────────────────────────────
#  Miniserver connection
# ─────────────────────────────────────────────

def fetch_structure(host, username, password, use_https=False, use_token=False):
    """Download LoxAPP3.json from the Miniserver."""
    protocol = "https" if use_https else "http"
    base_url = f"{protocol}://{host}"
    structure_url = f"{base_url}/data/LoxAPP3.json"

    session = requests.Session()

    try:
        if use_token:
            print("  Acquiring token...")
            token = _get_token(session, base_url, username, password)
            if not token:
                print("  Token acquisition failed, falling back to Basic Auth.")
                session.auth = HTTPBasicAuth(username, password)
            else:
                session.headers.update({"Authorization": f"Bearer {token}"})
        else:
            session.auth = HTTPBasicAuth(username, password)

        print(f"  Fetching structure file...")
        r = session.get(structure_url, timeout=20, verify=False)
        r.raise_for_status()
        return r.json()

    except requests.exceptions.ConnectionError:
        print(f"\nERROR: Cannot connect to Miniserver at {host}")
        print("  - Check the IP address / hostname")
        print("  - Make sure the Miniserver is reachable on your network")
        sys.exit(1)
    except requests.exceptions.Timeout:
        print("\nERROR: Connection timed out. Is the Miniserver online?")
        sys.exit(1)
    except requests.exceptions.HTTPError as e:
        code = e.response.status_code
        if code == 401:
            print(f"\nERROR: Authentication failed (401). Check username and password.")
        elif code == 403:
            print(f"\nERROR: Access denied (403). User may lack permissions.")
        else:
            print(f"\nERROR: HTTP {code}")
        sys.exit(1)
    except (json.JSONDecodeError, ValueError):
        print("\nERROR: Could not parse response as JSON. Is the host a Loxone Miniserver?")
        sys.exit(1)


def load_local_structure(file_path):
    """Load LoxAPP3.json from a local file (offline mode)."""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"\nERROR: File not found: {file_path}")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"\nERROR: Could not parse {file_path} as JSON.")
        sys.exit(1)


# ─────────────────────────────────────────────
#  Structure parsing
# ─────────────────────────────────────────────

def parse_structure(structure):
    """Extract and classify all controls from the structure.

    Walks both top-level controls and their subControls so that
    items nested inside compound blocks (LightControllerV2, AudioZoneV2,
    IRoomControllerV2, etc.) are captured individually.
    """
    raw_controls = structure.get("controls", {})
    rooms        = structure.get("rooms", {})
    cats         = structure.get("cats", {})

    room_name = {uuid: d.get("name", "—") for uuid, d in rooms.items()}
    cat_name  = {uuid: d.get("name", "—") for uuid, d in cats.items()}

    inputs, outputs, others = [], [], []
    seen_uuids = set()   # avoid double-counting items that appear at both levels

    def _classify(uuid, ctrl, parent_room_uuid="", parent_cat_uuid=""):
        """Classify one control block and recurse into its subControls."""
        if uuid in seen_uuids:
            return
        seen_uuids.add(uuid)

        ctrl_type = ctrl.get("type", "")

        if ctrl_type in SKIP_TYPES:
            return

        # Inherit room / cat from parent when not set on the sub-control
        room_uuid = ctrl.get("room", "") or parent_room_uuid
        cat_uuid  = ctrl.get("cat",  "") or parent_cat_uuid

        item = {
            "uuid":        uuid,
            "name":        ctrl.get("name", "Unnamed"),
            "type":        ctrl_type,
            "room":        room_name.get(room_uuid, "— No Room —"),
            "category":    cat_name.get(cat_uuid,   "— No Category —"),
            "is_favorite": ctrl.get("isFavorite", False),
            "bus":         BUS_TAGS.get(ctrl_type, ""),
            "states":      ctrl.get("states", {}),   # name → state-UUID
        }

        if ctrl_type in INPUT_TYPES:
            item["type_label"] = INPUT_TYPES[ctrl_type]
            inputs.append(item)
        elif ctrl_type in PHYSICAL_OUTPUTS:
            item["type_label"] = OUTPUT_TYPES.get(ctrl_type, ctrl_type)
            outputs.append(item)
        elif ctrl_type in OUTPUT_TYPES:
            # Logical controller — kept in "others" to avoid cluttering outputs
            item["type_label"] = OUTPUT_TYPES[ctrl_type]
            others.append(item)
        elif ctrl_type:
            # Unknown type — still list it so nothing is silently dropped
            item["type_label"] = ctrl_type
            others.append(item)
        # empty type → skip silently

        # ── Recurse into sub-controls ──────────────────────────────────────
        for sub_uuid, sub_ctrl in ctrl.get("subControls", {}).items():
            # Sub-control UUIDs can be "parentUUID/subUUID" — normalise
            short_uuid = sub_uuid.split("/")[-1] if "/" in sub_uuid else sub_uuid
            _classify(sub_uuid, sub_ctrl, room_uuid, cat_uuid)

    for uuid, ctrl in raw_controls.items():
        _classify(uuid, ctrl)

    # Sort each list by room then name
    key = lambda x: (x["room"].lower(), x["name"].lower())
    inputs.sort(key=key)
    outputs.sort(key=key)
    others.sort(key=key)

    return inputs, outputs, others


# ─────────────────────────────────────────────
#  PDF generation
# ─────────────────────────────────────────────

# Brand colours
LOXONE_GREEN = colors.HexColor("#6EBD44")
DARK_BLUE    = colors.HexColor("#1a3a5c")
MID_BLUE     = colors.HexColor("#2e6da4")
LIGHT_BLUE   = colors.HexColor("#d6eaf8")
INPUT_GREEN  = colors.HexColor("#1e8449")
INPUT_BG     = colors.HexColor("#eafaf1")
OUTPUT_BLUE  = colors.HexColor("#1a5276")
OUTPUT_BG    = colors.HexColor("#eaf4fb")
OTHER_AMBER  = colors.HexColor("#7d6608")
OTHER_BG     = colors.HexColor("#fef9e7")
ROW_ALT      = colors.HexColor("#f5f5f5")
GRID_COLOR   = colors.HexColor("#cccccc")


def _make_styles():
    styles = getSampleStyleSheet()

    title = ParagraphStyle(
        "LoxTitle", parent=styles["Title"],
        fontSize=20, textColor=DARK_BLUE,
        spaceAfter=2, spaceBefore=0,
    )
    subtitle = ParagraphStyle(
        "LoxSub", parent=styles["Normal"],
        fontSize=9, textColor=colors.HexColor("#555555"),
        spaceAfter=2,
    )
    section = ParagraphStyle(
        "LoxSection", parent=styles["Heading1"],
        fontSize=12, textColor=colors.white,
        spaceAfter=4, spaceBefore=10,
        leftIndent=4,
    )
    room_hdr = ParagraphStyle(
        "LoxRoom", parent=styles["Heading2"],
        fontSize=10, textColor=DARK_BLUE,
        spaceAfter=3, spaceBefore=6,
    )
    small = ParagraphStyle(
        "LoxSmall", parent=styles["Normal"],
        fontSize=8,
    )
    footer = ParagraphStyle(
        "LoxFooter", parent=styles["Normal"],
        fontSize=9, textColor=colors.grey,
    )

    return {
        "title":   title,
        "sub":     subtitle,
        "section": section,
        "room":    room_hdr,
        "small":   small,
        "footer":  footer,
        "normal":  styles["Normal"],
    }


STATUS_COLOR = {
    "ok":   colors.HexColor("#1e8449"),
    "nok":  colors.HexColor("#c0392b"),
    "skip": colors.HexColor("#e67e22"),
    None:   colors.HexColor("#aaaaaa"),
}
STATUS_BG = {
    "ok":   colors.HexColor("#eafaf1"),
    "nok":  colors.HexColor("#fdecea"),
    "skip": colors.HexColor("#fef5e7"),
    None:   colors.white,
}
STATUS_LABEL = {
    "ok":   "✓  OK",
    "nok":  "✗  NOK",
    "skip": "→  Skip",
    None:   "",
}


def _section_banner(text, bg_color, styles):
    """A full-width coloured banner acting as a section header."""
    data = [[Paragraph(f"<b>{text}</b>", styles["section"])]]
    t = Table(data, colWidths=[180 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",  (0, 0), (-1, -1), bg_color),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]))
    return t


def _item_table(items, header_color, row_alt_color):
    """Build a ReportLab table for a list of controls.

    Columns: Name | Type [Bus] | Category | Status | Notes / Fault
    Status and Notes are populated from _status / _note keys when present
    (set by the GUI after testing); otherwise left blank for a print-and-fill sheet.
    """
    # Name(72) | Type(42) | Category(24) | Status(20) | Notes(22) = 180 mm
    col_widths = [72*mm, 42*mm, 24*mm, 20*mm, 22*mm]
    header = ["Name", "Type  [Bus]", "Category", "Status", "Notes / Fault"]

    rows   = [header]
    styles_map = []   # (row_index, status) for colouring

    for i, item in enumerate(items, start=1):
        bus    = item.get("bus", "")
        ttype  = item["type_label"] + (f"  [{bus}]" if bus else "")
        status = item.get("_status")   # None if not yet tested
        note   = item.get("_note", "")
        rows.append([
            item["name"],
            ttype,
            item["category"],
            STATUS_LABEL[status],
            note,
        ])
        styles_map.append((i, status))

    t = Table(rows, colWidths=col_widths, repeatRows=1)

    style = [
        # Header
        ("BACKGROUND",    (0, 0), (-1, 0),  header_color),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.white),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  8),
        # Data
        ("FONTSIZE",      (0, 1), (-1, -1), 8),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, row_alt_color]),
        ("GRID",          (0, 0), (-1, -1), 0.3, GRID_COLOR),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        # Status column — centred
        ("ALIGN",         (3, 0), (3, -1),  "CENTER"),
        ("FONTNAME",      (3, 1), (3, -1),  "Helvetica-Bold"),
    ]

    # Per-row status cell colouring
    for row_idx, status in styles_map:
        if status is not None:
            style.append(("BACKGROUND", (3, row_idx), (3, row_idx), STATUS_BG[status]))
            style.append(("TEXTCOLOR",  (3, row_idx), (3, row_idx), STATUS_COLOR[status]))

    t.setStyle(TableStyle(style))
    return t


def _add_section(story, title, items, header_color, row_alt_color, styles, include_empty):
    """Add a titled section grouped by room to the story."""
    if not items and not include_empty:
        return

    story.append(Spacer(1, 2 * mm))
    story.append(_section_banner(f"{title}  ({len(items)} items)", header_color, styles))
    story.append(Spacer(1, 2 * mm))

    if not items:
        story.append(Paragraph("  No items of this type in the project.", styles["small"]))
        story.append(Spacer(1, 4 * mm))
        return

    # Group by room
    by_room = defaultdict(list)
    for item in items:
        by_room[item["room"]].append(item)

    for room, room_items in sorted(by_room.items()):
        block = []
        block.append(Paragraph(f"<b>{room}</b>", styles["room"]))
        block.append(_item_table(room_items, header_color, row_alt_color))
        block.append(Spacer(1, 3 * mm))
        story.append(KeepTogether(block))


def generate_pdf(inputs, outputs, others, ms_info, output_path,
                 include_other=False, company=""):
    """Build and save the PDF checklist."""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title="Loxone I/O Commissioning Checklist",
        author=company or "loxone_checklist.py",
    )

    styles = _make_styles()
    story  = []

    # ── Title block ──────────────────────────────────────────────────────────
    project   = ms_info.get("projectName", "Loxone Project")
    ms_name   = ms_info.get("msName",      "")
    serial    = ms_info.get("serialNr",    "")
    firmware  = ms_info.get("swVersion",   "")
    generated = datetime.now().strftime("%Y-%m-%d  %H:%M")

    if company:
        story.append(Paragraph(company, styles["title"]))
        story.append(Spacer(1, 1 * mm))
        story.append(HRFlowable(width="100%", thickness=2, color=LOXONE_GREEN))
        story.append(Spacer(1, 1 * mm))
        story.append(Paragraph("Loxone I/O Commissioning Checklist", styles["sub"]))
    else:
        story.append(Paragraph("Loxone I/O Commissioning Checklist", styles["title"]))
    story.append(Spacer(1, 1 * mm))
    story.append(HRFlowable(width="100%", thickness=3, color=LOXONE_GREEN))
    story.append(Spacer(1, 2 * mm))

    info_lines = [f"<b>Project:</b> {project}"]
    if ms_name:
        info_lines.append(f"<b>Miniserver:</b> {ms_name}")
    if serial:
        info_lines.append(f"<b>Serial:</b> {serial}")
    if firmware:
        info_lines.append(f"<b>Firmware:</b> {firmware}")
    info_lines.append(f"<b>Generated:</b> {generated}")

    for line in info_lines:
        story.append(Paragraph(line, styles["sub"]))

    story.append(Spacer(1, 4 * mm))

    # ── Summary table ────────────────────────────────────────────────────────
    total = len(inputs) + len(outputs) + (len(others) if include_other else 0)
    summary_data = [
        ["Section",               "Items"],
        ["Inputs  (sensors)",     str(len(inputs))],
        ["Outputs (actors)",      str(len(outputs))],
    ]
    if include_other:
        summary_data.append(["Other controls", str(len(others))])
    summary_data.append(["TOTAL", str(total)])

    summary = Table(summary_data, colWidths=[80 * mm, 30 * mm])
    summary.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0),   DARK_BLUE),
        ("TEXTCOLOR",    (0, 0), (-1, 0),   colors.white),
        ("FONTNAME",     (0, 0), (-1, 0),   "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1),  9),
        ("BACKGROUND",   (0, -1),(-1, -1),  LIGHT_BLUE),
        ("FONTNAME",     (0, -1),(-1, -1),  "Helvetica-Bold"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -2), [colors.white, ROW_ALT]),
        ("GRID",         (0, 0), (-1, -1),  0.5, GRID_COLOR),
        ("ALIGN",        (1, 0), (1, -1),   "CENTER"),
        ("TOPPADDING",   (0, 0), (-1, -1),  4),
        ("BOTTOMPADDING",(0, 0), (-1, -1),  4),
        ("LEFTPADDING",  (0, 0), (-1, -1),  6),
    ]))
    story.append(summary)
    story.append(Spacer(1, 6 * mm))

    # ── Legend ───────────────────────────────────────────────────────────────
    legend = Paragraph(
        "<b>Status key:</b>  "
        "<b><font color='#1e8449'>✓ OK</font></b> — tested and working   "
        "<b><font color='#c0392b'>✗ NOK</font></b> — fault / not working   "
        "<b><font color='#e67e22'>→ Skip</font></b> — deferred / not applicable   "
        "blank — not yet tested",
        styles["small"]
    )
    story.append(legend)
    story.append(Spacer(1, 4 * mm))
    story.append(HRFlowable(width="100%", thickness=1, color=GRID_COLOR))

    # ── Inputs section ───────────────────────────────────────────────────────
    _add_section(story, "INPUTS — Sensors & Digital Inputs",
                 inputs, INPUT_GREEN, INPUT_BG, styles, include_empty=True)

    # ── Outputs section ──────────────────────────────────────────────────────
    _add_section(story, "OUTPUTS — Actors & Controls",
                 outputs, OUTPUT_BLUE, OUTPUT_BG, styles, include_empty=True)

    # ── Other section ────────────────────────────────────────────────────────
    if include_other:
        _add_section(story, "OTHER CONTROLS",
                     others, OTHER_AMBER, OTHER_BG, styles, include_empty=False)

    # ── Sign-off block ───────────────────────────────────────────────────────
    story.append(Spacer(1, 12 * mm))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
    story.append(Spacer(1, 4 * mm))

    signoff_data = [[
        "Technician:",
        "_" * 28,
        "Date:",
        "_" * 18,
        "Signature:",
        "_" * 22,
    ]]
    signoff = Table(signoff_data, colWidths=[22*mm, 52*mm, 12*mm, 32*mm, 20*mm, 38*mm])
    signoff.setStyle(TableStyle([
        ("FONTSIZE",     (0, 0), (-1, -1), 9),
        ("VALIGN",       (0, 0), (-1, -1), "BOTTOM"),
        ("TOPPADDING",   (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 2),
    ]))
    story.append(signoff)

    # ── Build ────────────────────────────────────────────────────────────────
    doc.build(story)
    print(f"\nChecklist saved: {output_path}")
    print(f"  Total pages:  see PDF reader")
    print(f"  Items:        {len(inputs)} inputs  |  {len(outputs)} outputs"
          + (f"  |  {len(others)} other" if include_other else ""))


# ─────────────────────────────────────────────
#  CLI entry point
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        prog="loxone_checklist",
        description="Generate a PDF I/O commissioning checklist from a Loxone Miniserver.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Connect live to a Miniserver
  python loxone_checklist.py 192.168.1.100 admin secret

  # Custom output filename
  python loxone_checklist.py 192.168.1.100 admin secret -o my_project_checklist.pdf

  # Use HTTPS
  python loxone_checklist.py 192.168.1.100 admin secret --https

  # Include all other control types in the PDF
  python loxone_checklist.py 192.168.1.100 admin secret --include-other

  # Offline mode — parse a local LoxAPP3.json file
  python loxone_checklist.py --file LoxAPP3.json
""",
    )

    # Connection arguments
    parser.add_argument("host",     nargs="?", help="Miniserver IP address or hostname")
    parser.add_argument("username", nargs="?", help="Miniserver username (e.g. admin)")
    parser.add_argument("password", nargs="?", help="Miniserver password")

    # Options
    parser.add_argument("-o", "--output", default="loxone_checklist.pdf",
                        metavar="FILE",
                        help="Output PDF filename (default: loxone_checklist.pdf)")
    parser.add_argument("--https",        action="store_true",
                        help="Connect via HTTPS instead of HTTP")
    parser.add_argument("--token",        action="store_true",
                        help="Use token-based authentication (recommended for Gen2)")
    parser.add_argument("--include-other",action="store_true",
                        help="Include unrecognised control types in the PDF")
    parser.add_argument("--file",         metavar="PATH",
                        help="Load structure from a local LoxAPP3.json file (offline mode)")

    args = parser.parse_args()

    # ── Load structure ────────────────────────────────────────────────────────
    if args.file:
        print(f"Loading structure from: {args.file}")
        structure = load_local_structure(args.file)
    else:
        if not all([args.host, args.username, args.password]):
            parser.print_help()
            print("\nERROR: Provide <host> <username> <password>, or --file <LoxAPP3.json>")
            sys.exit(1)
        print(f"Connecting to Miniserver: {args.host}")
        structure = fetch_structure(
            args.host, args.username, args.password,
            use_https=args.https, use_token=args.token
        )

    # ── Parse ─────────────────────────────────────────────────────────────────
    ms_info = structure.get("msInfo", {})
    project = ms_info.get("projectName", "Unknown Project")
    print(f"Project: {project}")

    inputs, outputs, others = parse_structure(structure)
    print(f"Parsed:  {len(inputs)} inputs  |  {len(outputs)} outputs  |  {len(others)} other")

    if len(inputs) + len(outputs) == 0:
        print("\nWARNING: No recognised inputs or outputs found.")
        print("  The project may use control types not in the built-in list.")
        print("  Re-run with --include-other to include all control types.")

    # ── Generate PDF ───────────────────────────────────────────────────────────
    generate_pdf(
        inputs, outputs, others, ms_info,
        output_path=args.output,
        include_other=args.include_other
    )


if __name__ == "__main__":
    main()
