#!/usr/bin/env python3
"""LitSearch Pro v19.4 collaboration server with Windows GUI management."""

import argparse
import faulthandler
import base64
import csv
import ctypes
import gc
import hashlib
import html
import io
import json
import os
import queue
import re
import secrets
import socket
import sqlite3
import sys
import time
import threading
import tkinter as tk
import traceback
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    import pystray
except Exception:
    pystray = None
try:
    import fitz
except Exception:
    fitz = None
try:
    from PIL import Image, ImageDraw
except Exception:
    Image = ImageDraw = None


NIM_ADD = 0x00000000
NIM_DELETE = 0x00000002
NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004
WM_USER = 0x0400
WM_COMMAND = 0x0111
WM_RBUTTONUP = 0x0205
WM_LBUTTONDBLCLK = 0x0203
TPM_RETURNCMD = 0x0100
TPM_NONOTIFY = 0x0080
MF_STRING = 0x0000
MF_BYPOSITION = 0x00000400
CMD_OPEN = 10001
CMD_EXIT = 10002


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("hWnd", ctypes.c_void_p),
        ("uID", ctypes.c_uint),
        ("uFlags", ctypes.c_uint),
        ("uCallbackMessage", ctypes.c_uint),
        ("hIcon", ctypes.c_void_p),
        ("szTip", ctypes.c_wchar * 128),
        ("dwState", ctypes.c_ulong),
        ("dwStateMask", ctypes.c_ulong),
        ("szInfo", ctypes.c_wchar * 256),
        ("uTimeoutOrVersion", ctypes.c_uint),
        ("szInfoTitle", ctypes.c_wchar * 64),
        ("dwInfoFlags", ctypes.c_ulong),
    ]


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class MSG(ctypes.Structure):
    _fields_ = [("hwnd", ctypes.c_void_p), ("message", ctypes.c_uint), ("wParam", ctypes.c_void_p), ("lParam", ctypes.c_void_p), ("time", ctypes.c_ulong), ("pt", POINT)]


WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_longlong, ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [("style", ctypes.c_uint), ("lpfnWndProc", WNDPROC), ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int), ("hInstance", ctypes.c_void_p), ("hIcon", ctypes.c_void_p), ("hCursor", ctypes.c_void_p), ("hbrBackground", ctypes.c_void_p), ("lpszMenuName", ctypes.c_wchar_p), ("lpszClassName", ctypes.c_wchar_p)]


SERVER_VERSION = "22.1.21-generic Campus Safety Workflow Platform"
PDF_CLEANUP_LOCK = threading.Lock()
PDF_CLEANUP_LAST = 0.0
PDF_CLEANUP_INTERVAL_SECONDS = 3600
BLACKLIST_CLEANUP_LOCK = threading.Lock()
BLACKLIST_CLEANUP_LAST = 0.0
BLACKLIST_CLEANUP_INTERVAL_SECONDS = 60
SERVER_NAME = "LitSearchPro Generic Collaboration Server"
UX_ACCENT = "#02529F"
UX_BG = "#F5F7FA"
UX_SURFACE = "#FFFFFF"
UX_BORDER = "#D7DEE8"
UX_TEXT = "#111827"
UX_MUTED = "#5B6472"
UX_HOVER = "#EAF3FF"
APP_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "LitSearchProServer")
DB_FILE = os.path.join(APP_DIR, "collaboration_server.db")
FILES_DIR = os.path.join(APP_DIR, "uploads")
CONFIG_FILE = os.path.join(APP_DIR, "server_config.json")
ARCHIVE_DIR = os.path.join(APP_DIR, "permanent_archives")

LAB_WORKBOOK_SHEETS = (
    ("实验流程", ("步骤序号", "操作内容", "设备/条件", "预计时长", "责任人")),
    ("化学品使用", ("化学品名称", "CAS号", "预计用量", "单位", "危险性", "废弃物处置")),
    ("实验安全预案", ("风险源", "可能后果", "预防措施", "应急处置", "联系人")),
)


def xlsx_column_name(index):
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def create_lab_workbook_bytes():
    """Create the same dependency-free three-sheet workbook used by the desktop client."""
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    workbook_sheets, workbook_rels, worksheet_files = [], [], {}
    for sheet_index, (sheet_name, headers) in enumerate(LAB_WORKBOOK_SHEETS, 1):
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{sheet_index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        workbook_sheets.append(
            f'<sheet name="{html.escape(sheet_name, quote=True)}" sheetId="{sheet_index}" r:id="rId{sheet_index}"/>'
        )
        workbook_rels.append(
            f'<Relationship Id="rId{sheet_index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{sheet_index}.xml"/>'
        )
        cells = []
        for col_index, value in enumerate(headers, 1):
            ref = f"{xlsx_column_name(col_index)}1"
            cells.append(f'<c r="{ref}" t="inlineStr" s="1"><is><t>{html.escape(value)}</t></is></c>')
        worksheet_files[f"xl/worksheets/sheet{sheet_index}.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
            f'<row r="1">{"".join(cells)}</row></sheetData></worksheet>'
        )
    content_types.extend([
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
        "</Types>",
    ])
    workbook_rels.append(
        f'<Relationship Id="rId{len(LAB_WORKBOOK_SHEETS)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "".join(content_types))
        archive.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + "".join(workbook_sheets) + "</sheets></workbook>")
        archive.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(workbook_rels) + "</Relationships>")
        archive.writestr("xl/styles.xml", '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Microsoft YaHei"/></font><font><b/><sz val="11"/><name val="Microsoft YaHei"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCEBFA"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="1" fillId="1" borderId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>')
        for name, content in worksheet_files.items():
            archive.writestr(name, content)
    return output.getvalue()


def parse_lab_workbook_bytes(raw):
    if not raw:
        raise ValueError("请选择实验室预约资料工作簿")
    if len(raw) > 8 * 1024 * 1024:
        raise ValueError("工作簿不能超过 8 MB")
    namespace = {
        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            shared = []
            if "xl/sharedStrings.xml" in archive.namelist():
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                for item in root.findall("m:si", namespace):
                    shared.append("".join(node.text or "" for node in item.iter() if node.tag.endswith("}t")))
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rels = {node.attrib["Id"]: node.attrib["Target"] for node in rels_root}
            sheets = {}
            for sheet in workbook.findall("m:sheets/m:sheet", namespace):
                target = rels[sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]]
                target = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
                root = ET.fromstring(archive.read(target))
                rows = []
                for row in root.findall(".//m:sheetData/m:row", namespace):
                    values = []
                    for cell in row.findall("m:c", namespace):
                        match = re.match(r"[A-Z]+", cell.attrib.get("r", "A1"))
                        if not match:
                            continue
                        col_index = 0
                        for char in match.group(0):
                            col_index = col_index * 26 + ord(char) - 64
                        while len(values) < col_index:
                            values.append("")
                        inline, value_node = cell.find("m:is", namespace), cell.find("m:v", namespace)
                        value = ""
                        if inline is not None:
                            value = "".join(node.text or "" for node in inline.iter() if node.tag.endswith("}t"))
                        elif value_node is not None:
                            value = value_node.text or ""
                            if cell.attrib.get("t") == "s":
                                value = shared[int(value)]
                        values[col_index - 1] = str(value).strip()
                    rows.append(values)
                sheets[sheet.attrib["name"]] = rows
    except (KeyError, IndexError, ET.ParseError, zipfile.BadZipFile, ValueError) as exc:
        raise ValueError("无法读取该 Excel 工作簿，请下载并使用系统模板") from exc
    result = {}
    keys = {"实验流程": "workflow", "化学品使用": "chemicals", "实验安全预案": "safety_plan"}
    for sheet_name, required_headers in LAB_WORKBOOK_SHEETS:
        rows = sheets.get(sheet_name)
        if not rows:
            raise ValueError(f"工作簿缺少“{sheet_name}”工作表")
        headers = [str(x).strip() for x in rows[0]]
        if headers[:len(required_headers)] != list(required_headers):
            raise ValueError(f"“{sheet_name}”表头不正确，请重新下载模板")
        records = []
        for values in rows[1:]:
            padded = list(values) + [""] * (len(required_headers) - len(values))
            record = {header: padded[index].strip() for index, header in enumerate(required_headers)}
            if any(record.values()):
                records.append(record)
        result[keys[sheet_name]] = records
    if not result["workflow"]:
        raise ValueError("“实验流程”工作表至少需要填写一条记录")
    if not result["safety_plan"]:
        raise ValueError("“实验安全预案”工作表至少需要填写一条记录")
    return result


def write_server_log(filename, text):
    try:
        ensure_dirs()
        with open(os.path.join(APP_DIR, filename), "a", encoding="utf-8") as fh:
            fh.write(f"\n[{now()}]\n{text}\n")
    except Exception:
        pass


def install_global_crash_handlers():
    ensure_dirs()
    try:
        crash_path = os.path.join(APP_DIR, "server_native_crash.log")
        fh = open(crash_path, "a", encoding="utf-8")
        faulthandler.enable(fh, all_threads=True)
    except Exception:
        pass

    def excepthook(exc_type, exc_value, exc_tb):
        write_server_log("server_crash.log", "".join(traceback.format_exception(exc_type, exc_value, exc_tb)))

    def threadhook(args):
        write_server_log("server_crash.log", "".join(traceback.format_exception(args.exc_type, args.exc_value, args.exc_traceback)))

    sys.excepthook = excepthook
    if hasattr(threading, "excepthook"):
        threading.excepthook = threadhook


class ServerRoundedButton(tk.Canvas):
    """Small Windows 11 style rounded button used by the server GUI."""

    def __init__(self, parent, text, command, kind="secondary", width=None, height=34):
        self.text = text
        self.command = command
        self.kind = kind
        self.hover = False
        self.pressed = False
        width = width or max(92, len(text) * 13 + 32)
        try:
            parent_bg = parent.cget("bg")
        except Exception:
            parent_bg = UX_SURFACE
        super().__init__(parent, width=width, height=height, highlightthickness=0, bd=0, bg=parent_bg, cursor="hand2")
        self.bind("<Configure>", lambda _e: self._draw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self._draw()

    def _palette(self):
        if self.kind == "primary":
            return ("#014A91" if self.hover else UX_ACCENT, "#01386F" if self.pressed else "#01427E", "white")
        if self.kind == "danger":
            return ("#FDECEC" if not self.hover else "#FAD1D1", "#F8B4B4", "#B42318")
        return ("#F8FAFD" if not self.hover else UX_HOVER, "#D5E1F0", UX_TEXT)

    def _round_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = [
            x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
            x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
            x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def _draw(self):
        self.delete("all")
        w = max(2, self.winfo_width())
        h = max(2, self.winfo_height())
        fill, outline, fg = self._palette()
        offset = 1 if self.pressed else 0
        self._round_rect(1, 1 + offset, w - 1, h - 1 + offset, 10, fill=fill, outline=outline, width=1)
        self.create_text(w / 2, h / 2 + offset, text=self.text, fill=fg, font=("Microsoft YaHei UI", 9, "bold" if self.kind == "primary" else "normal"))

    def _on_enter(self, _event):
        self.hover = True
        self._draw()

    def _on_leave(self, _event):
        self.hover = False
        self.pressed = False
        self._draw()

    def _on_press(self, _event):
        self.pressed = True
        self._draw()

    def _on_release(self, _event):
        should_run = self.pressed
        self.pressed = False
        self._draw()
        if should_run and callable(self.command):
            self.command()


def now():
    return datetime.now().isoformat(timespec="seconds")


def ensure_dirs():
    for path in (APP_DIR, FILES_DIR, ARCHIVE_DIR):
        if os.path.lexists(path) and not os.path.isdir(path):
            backup = f"{path}.file-backup-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            os.replace(path, backup)
        try:
            os.makedirs(path, exist_ok=True)
        except FileExistsError:
            if not os.path.isdir(path):
                backup = f"{path}.file-backup-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                os.replace(path, backup)
                os.makedirs(path, exist_ok=True)
            else:
                raise


def load_config():
    ensure_dirs()
    defaults = {"host": "0.0.0.0", "port": 8765, "public_url": "", "group_code": "research-lab", "data_dir": APP_DIR}
    try:
        with open(CONFIG_FILE, encoding="utf-8") as fh:
            defaults.update(json.load(fh))
    except Exception:
        pass
    return defaults


def save_config(config):
    ensure_dirs()
    with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
        json.dump(config, fh, ensure_ascii=False, indent=2)


def apply_data_dir(config):
    global APP_DIR, DB_FILE, FILES_DIR, CONFIG_FILE, ARCHIVE_DIR
    data_dir = os.path.abspath(os.path.expandvars(config.get("data_dir") or APP_DIR))
    APP_DIR = data_dir
    DB_FILE = os.path.join(APP_DIR, "collaboration_server.db")
    FILES_DIR = os.path.join(APP_DIR, "uploads")
    CONFIG_FILE = os.path.join(APP_DIR, "server_config.json")
    ARCHIVE_DIR = os.path.join(APP_DIR, "permanent_archives")
    ensure_dirs()


def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 160000).hex()
    return salt, digest


def open_db_connection():
    ensure_dirs()
    conn = sqlite3.connect(DB_FILE, timeout=15, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=15000")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA temp_store=MEMORY")
    conn.execute("PRAGMA cache_size=-512")
    conn.execute("PRAGMA wal_autocheckpoint=1000")
    return conn


def connect_db():
    conn = open_db_connection()
    conn.executescript(
        """
        PRAGMA journal_mode=WAL;
        CREATE TABLE IF NOT EXISTS users(
          id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, username TEXT NOT NULL,
          role TEXT DEFAULT '学生', salt TEXT NOT NULL, password_hash TEXT NOT NULL,
          active INTEGER DEFAULT 0, created_at TEXT NOT NULL,
          UNIQUE(group_code, username)
        );
        CREATE TABLE IF NOT EXISTS tokens(token TEXT PRIMARY KEY, user_id INTEGER NOT NULL, created_at TEXT NOT NULL, last_seen TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS changes(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, sender TEXT DEFAULT '', entity_type TEXT NOT NULL, entity_id TEXT DEFAULT '', action TEXT NOT NULL, payload TEXT DEFAULT '{}', created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS messages(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, sender TEXT DEFAULT '', recipient TEXT DEFAULT '', subject TEXT DEFAULT '', body TEXT NOT NULL, created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS files(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, uploader TEXT DEFAULT '', name TEXT NOT NULL, path TEXT NOT NULL, size INTEGER DEFAULT 0, created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS leave_requests(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, requester TEXT DEFAULT '', leave_type TEXT DEFAULT '请假', start_time TEXT DEFAULT '', end_time TEXT DEFAULT '', reason TEXT DEFAULT '', status TEXT DEFAULT '待导师审批', approver TEXT DEFAULT '', approved_at TEXT DEFAULT '', created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS attendance(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, username TEXT DEFAULT '', action TEXT DEFAULT '打卡', ip_address TEXT DEFAULT '', note TEXT DEFAULT '', created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS task_plans(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, creator TEXT DEFAULT '', assignee TEXT DEFAULT '', title TEXT NOT NULL, detail TEXT DEFAULT '', due_date TEXT DEFAULT '', status TEXT DEFAULT '待导师审批', reviewer TEXT DEFAULT '', review_note TEXT DEFAULT '', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS meeting_settings(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL UNIQUE, weekday TEXT DEFAULT '周五', time_text TEXT DEFAULT '15:00', location TEXT DEFAULT '', updated_by TEXT DEFAULT '', updated_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS meeting_reports(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, student TEXT DEFAULT '', name TEXT NOT NULL, path TEXT NOT NULL, size INTEGER DEFAULT 0, status TEXT DEFAULT '已上传', expires_at TEXT DEFAULT '', created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS equipment(id INTEGER PRIMARY KEY, team_name TEXT DEFAULT '', group_code TEXT NOT NULL, owner_teacher TEXT DEFAULT '', name TEXT NOT NULL, brand TEXT DEFAULT '', category TEXT DEFAULT '', model TEXT DEFAULT '', manager1 TEXT DEFAULT '', manager2 TEXT DEFAULT '', current_user TEXT DEFAULT '', approver TEXT DEFAULT '', status TEXT DEFAULT '可用', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS equipment_requests(id INTEGER PRIMARY KEY, equipment_id INTEGER NOT NULL, group_code TEXT NOT NULL, requester TEXT DEFAULT '', request_type TEXT DEFAULT '借用', reason TEXT DEFAULT '', status TEXT DEFAULT '待审批', approver TEXT DEFAULT '', review_note TEXT DEFAULT '', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS equipment_managers(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, username TEXT NOT NULL, created_by TEXT DEFAULT '', created_at TEXT NOT NULL, UNIQUE(group_code, username));
        CREATE TABLE IF NOT EXISTS auth_codes(id INTEGER PRIMARY KEY, code TEXT UNIQUE NOT NULL, group_code TEXT NOT NULL, team_name TEXT DEFAULT '', creator TEXT DEFAULT '', expires_at TEXT NOT NULL, created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS v21_records(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, creator TEXT DEFAULT '', kind TEXT NOT NULL, title TEXT NOT NULL, body TEXT DEFAULT '', tags TEXT DEFAULT '', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS v21_notifications(id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, creator TEXT DEFAULT '', title TEXT NOT NULL, body TEXT DEFAULT '', level TEXT DEFAULT '普通', created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS user_signatures(
          id INTEGER PRIMARY KEY,user_id INTEGER NOT NULL UNIQUE,image_b64 TEXT NOT NULL,
          file_name TEXT DEFAULT '',updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratories(
          id INTEGER PRIMARY KEY,name TEXT NOT NULL,college TEXT NOT NULL,address TEXT NOT NULL,
          lab_type TEXT NOT NULL DEFAULT '公共实验室',team_name TEXT DEFAULT '',group_code TEXT DEFAULT '',
          managers_json TEXT DEFAULT '[]',commitment_text TEXT DEFAULT '',active INTEGER DEFAULT 1,
          created_by TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratory_auth_codes(
          id INTEGER PRIMARY KEY,code TEXT UNIQUE NOT NULL,laboratory_id INTEGER NOT NULL,
          creator TEXT DEFAULT '',expires_at TEXT NOT NULL,created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratory_blacklist(
          id INTEGER PRIMARY KEY,laboratory_id INTEGER NOT NULL,user_id INTEGER NOT NULL,
          blacklist_type TEXT NOT NULL DEFAULT '暂停',reason TEXT NOT NULL,
          starts_at TEXT NOT NULL,ends_at TEXT DEFAULT '',active INTEGER DEFAULT 1,
          created_by TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratory_reservations(
          id INTEGER PRIMARY KEY,laboratory_id INTEGER NOT NULL,requester_id INTEGER NOT NULL,
          requester_group TEXT NOT NULL,requester_teacher TEXT DEFAULT '',companion_user_id INTEGER,
          companion_name TEXT DEFAULT '',companion_role TEXT DEFAULT '',start_time TEXT NOT NULL,end_time TEXT NOT NULL,
          purpose TEXT DEFAULT '',workbook_name TEXT DEFAULT '',workbook_b64 TEXT DEFAULT '',
          workflow_json TEXT DEFAULT '[]',chemicals_json TEXT DEFAULT '[]',safety_plan_json TEXT DEFAULT '[]',
          commitment_snapshot TEXT DEFAULT '',student_signature_b64 TEXT DEFAULT '',
          mentor_signature_b64 TEXT DEFAULT '',manager_signature_b64 TEXT DEFAULT '',
          mentor_status TEXT DEFAULT '待导师审核',mentor_reviewer TEXT DEFAULT '',mentor_note TEXT DEFAULT '',
          manager_status TEXT DEFAULT '待实验室审核',manager_reviewer TEXT DEFAULT '',manager_note TEXT DEFAULT '',
          status TEXT DEFAULT '待导师审核',completion_status TEXT DEFAULT '未开始',
          completion_report TEXT DEFAULT '',hazard_report TEXT DEFAULT '',completion_review_note TEXT DEFAULT '',
          pdf_path TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratory_reservation_participants(
          id INTEGER PRIMARY KEY,reservation_id INTEGER NOT NULL,user_id INTEGER NOT NULL,
          participant_name TEXT DEFAULT '',participant_role TEXT DEFAULT '',teacher_name TEXT DEFAULT '',
          confirmation_status TEXT DEFAULT '待确认',confirmation_note TEXT DEFAULT '',
          signature_b64 TEXT DEFAULT '',confirmed_at TEXT DEFAULT '',
          created_at TEXT NOT NULL,updated_at TEXT NOT NULL,
          UNIQUE(reservation_id,user_id)
        );
        CREATE TABLE IF NOT EXISTS laboratory_reservation_mentor_reviews(
          id INTEGER PRIMARY KEY,reservation_id INTEGER NOT NULL,mentor_user_id INTEGER NOT NULL,
          mentor_name TEXT NOT NULL,status TEXT DEFAULT '待审核',review_note TEXT DEFAULT '',
          signature_b64 TEXT DEFAULT '',reviewed_at TEXT DEFAULT '',
          created_at TEXT NOT NULL,updated_at TEXT NOT NULL,
          UNIQUE(reservation_id,mentor_user_id)
        );
        CREATE TABLE IF NOT EXISTS laboratory_audit_logs(
          id INTEGER PRIMARY KEY,reservation_id INTEGER NOT NULL,action TEXT NOT NULL,
          actor_id INTEGER,actor_name TEXT DEFAULT '',actor_role TEXT DEFAULT '',
          detail TEXT DEFAULT '',created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratory_announcements(
          id INTEGER PRIMARY KEY,laboratory_id INTEGER NOT NULL,title TEXT NOT NULL,
          body TEXT NOT NULL,active INTEGER DEFAULT 1,created_by TEXT DEFAULT '',
          created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratory_channel_proposals(
          id INTEGER PRIMARY KEY,laboratory_id INTEGER NOT NULL,target_state TEXT NOT NULL,
          reason TEXT DEFAULT '',status TEXT DEFAULT '表决中',created_by_id INTEGER,
          created_by TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laboratory_channel_votes(
          id INTEGER PRIMARY KEY,proposal_id INTEGER NOT NULL,manager_user_id INTEGER NOT NULL,
          manager_name TEXT DEFAULT '',decision TEXT NOT NULL,voted_at TEXT NOT NULL,
          UNIQUE(proposal_id,manager_user_id)
        );
        CREATE TABLE IF NOT EXISTS chemical_warehouses(
          id INTEGER PRIMARY KEY,name TEXT NOT NULL,college TEXT NOT NULL,address TEXT NOT NULL,
          managers_json TEXT NOT NULL DEFAULT '[]',commitment_text TEXT DEFAULT '',active INTEGER DEFAULT 1,
          created_by TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS chemicals(
          id INTEGER PRIMARY KEY,warehouse_id INTEGER NOT NULL,name TEXT NOT NULL,unit TEXT NOT NULL DEFAULT 'g',
          owner_teacher TEXT NOT NULL,owner_group TEXT NOT NULL,quantity REAL NOT NULL DEFAULT 0,
          available_per_student REAL NOT NULL DEFAULT 0,auth_code TEXT DEFAULT '',
          created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS chemical_inbound_requests(
          id INTEGER PRIMARY KEY,chemical_id INTEGER NOT NULL,purchaser_id INTEGER NOT NULL,quantity REAL NOT NULL,
          source_note TEXT DEFAULT '',status TEXT DEFAULT '待库管审核',reviewer TEXT DEFAULT '',review_note TEXT DEFAULT '',
          created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS chemical_withdrawals(
          id INTEGER PRIMARY KEY,withdrawal_no TEXT DEFAULT '',chemical_id INTEGER NOT NULL,requester_id INTEGER NOT NULL,
          requester_group TEXT NOT NULL,requester_teacher TEXT DEFAULT '',owner_teacher TEXT NOT NULL,
          quantity REAL NOT NULL,purpose TEXT NOT NULL,commitment_snapshot TEXT DEFAULT '',
          student_signature_b64 TEXT DEFAULT '',mentor_signature_b64 TEXT DEFAULT '',
          owner_signature_b64 TEXT DEFAULT '',manager1_signature_b64 TEXT DEFAULT '',manager2_signature_b64 TEXT DEFAULT '',
          mentor_status TEXT DEFAULT '待导师审核',owner_status TEXT DEFAULT '无需审核',
          manager1_status TEXT DEFAULT '待库管一审核',manager2_status TEXT DEFAULT '待库管二审核',
          status TEXT DEFAULT '待导师审核',review_notes_json TEXT DEFAULT '{}',
          pdf_path TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS chemical_withdrawal_participants(
          id INTEGER PRIMARY KEY,withdrawal_id INTEGER NOT NULL,user_id INTEGER NOT NULL,
          participant_name TEXT DEFAULT '',confirmation_status TEXT DEFAULT '待确认',
          confirmation_note TEXT DEFAULT '',signature_b64 TEXT DEFAULT '',
          confirmed_at TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL,
          UNIQUE(withdrawal_id,user_id)
        );
        CREATE TABLE IF NOT EXISTS chemical_audit_logs(
          id INTEGER PRIMARY KEY,withdrawal_id INTEGER NOT NULL,action TEXT NOT NULL,
          actor_id INTEGER,actor_name TEXT DEFAULT '',detail TEXT DEFAULT '',created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS warehouse_channel_proposals(
          id INTEGER PRIMARY KEY,warehouse_id INTEGER NOT NULL,target_state TEXT NOT NULL,
          reason TEXT DEFAULT '',status TEXT DEFAULT '表决中',created_by_id INTEGER,
          created_by TEXT DEFAULT '',created_at TEXT NOT NULL,updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS warehouse_channel_votes(
          id INTEGER PRIMARY KEY,proposal_id INTEGER NOT NULL,manager_user_id INTEGER NOT NULL,
          manager_name TEXT DEFAULT '',decision TEXT NOT NULL,voted_at TEXT NOT NULL,
          UNIQUE(proposal_id,manager_user_id)
        );
        CREATE TABLE IF NOT EXISTS chemical_inventory_logs(
          id INTEGER PRIMARY KEY,chemical_id INTEGER NOT NULL,action TEXT NOT NULL,quantity REAL NOT NULL,
          balance_after REAL NOT NULL,related_type TEXT DEFAULT '',related_id INTEGER,
          operator TEXT DEFAULT '',note TEXT DEFAULT '',created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_users_group_username ON users(group_code,username);
        CREATE INDEX IF NOT EXISTS idx_users_role_active_group ON users(role,active,group_code);
        CREATE INDEX IF NOT EXISTS idx_tokens_user_last_seen ON tokens(user_id,last_seen);
        CREATE INDEX IF NOT EXISTS idx_equipment_scope_status ON equipment(group_code,team_name,status,name);
        CREATE INDEX IF NOT EXISTS idx_equipment_owner ON equipment(group_code,owner_teacher,status);
        CREATE INDEX IF NOT EXISTS idx_equipment_requests_latest ON equipment_requests(equipment_id,status,id DESC);
        CREATE INDEX IF NOT EXISTS idx_equipment_requests_requester ON equipment_requests(group_code,requester,status,id DESC);
        CREATE INDEX IF NOT EXISTS idx_equipment_requests_approval ON equipment_requests(status,equipment_id,id DESC);
        CREATE INDEX IF NOT EXISTS idx_equipment_managers_lookup ON equipment_managers(group_code,username);
        CREATE INDEX IF NOT EXISTS idx_lab_res_requester ON laboratory_reservations(requester_id,status,id DESC);
        CREATE INDEX IF NOT EXISTS idx_lab_res_requester_lab ON laboratory_reservations(requester_id,laboratory_id);
        CREATE INDEX IF NOT EXISTS idx_lab_res_lab_status ON laboratory_reservations(laboratory_id,status,id DESC);
        CREATE INDEX IF NOT EXISTS idx_lab_participant_user ON laboratory_reservation_participants(user_id,confirmation_status,reservation_id);
        CREATE INDEX IF NOT EXISTS idx_lab_mentor_review_user ON laboratory_reservation_mentor_reviews(mentor_user_id,status,reservation_id);
        CREATE INDEX IF NOT EXISTS idx_laboratories_scope ON laboratories(active,lab_type,team_name,id);
        CREATE INDEX IF NOT EXISTS idx_lab_blacklist_user ON laboratory_blacklist(laboratory_id,user_id,active,ends_at);
        CREATE INDEX IF NOT EXISTS idx_chemical_owner ON chemicals(warehouse_id,owner_teacher,name);
        CREATE INDEX IF NOT EXISTS idx_chemical_withdraw_requester ON chemical_withdrawals(requester_id,status,id DESC);
        CREATE INDEX IF NOT EXISTS idx_chemical_logs ON chemical_inventory_logs(chemical_id,id DESC);
        """
    )
    def add_column(table, name, definition):
        cols = [row["name"] for row in conn.execute(f"PRAGMA table_info({table})")]
        if name not in cols:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")
    for col, definition in (("display_name", "TEXT DEFAULT ''"), ("team_name", "TEXT DEFAULT ''")):
        add_column("users", col, definition)
    add_column("users", "mentor_user_id", "INTEGER")
    add_column("users", "mentor_username", "TEXT DEFAULT ''")
    for col, definition in (("recipient", "TEXT DEFAULT '全体成员'"), ("encrypted", "INTEGER DEFAULT 1"), ("expires_at", "TEXT DEFAULT ''"), ("note", "TEXT DEFAULT ''")):
        add_column("files", col, definition)
    add_column("equipment_requests", "requester_teacher", "TEXT DEFAULT ''")
    add_column("meeting_reports", "expires_at", "TEXT DEFAULT ''")
    add_column("users", "phone", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "reservation_no", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "requester_phone", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "participants_json", "TEXT DEFAULT '[]'")
    add_column("laboratory_reservations", "mentor_reviews_json", "TEXT DEFAULT '[]'")
    add_column("laboratories", "booking_open", "INTEGER DEFAULT 1")
    add_column("laboratory_reservations", "experiment_status", "TEXT DEFAULT '未开始'")
    add_column("laboratory_reservations", "started_at", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "ended_at", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "terminated_by", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "termination_reason", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "pdf_expires_at", "TEXT DEFAULT ''")
    add_column("laboratory_reservations", "certificate_text", "TEXT DEFAULT ''")
    add_column("chemical_withdrawals", "co_collector_id", "INTEGER")
    add_column("chemical_withdrawals", "withdrawal_no", "TEXT DEFAULT ''")
    add_column("chemical_withdrawals", "co_collector_name", "TEXT DEFAULT ''")
    add_column("chemical_withdrawals", "storage_location", "TEXT DEFAULT ''")
    add_column("chemical_withdrawals", "participant_status", "TEXT DEFAULT '待共同领用人确认'")
    add_column("chemical_withdrawals", "disposal_status", "TEXT DEFAULT '未报告'")
    add_column("chemical_withdrawals", "disposal_report", "TEXT DEFAULT ''")
    add_column("chemical_withdrawals", "disposed_at", "TEXT DEFAULT ''")
    add_column("chemical_withdrawals", "pdf_expires_at", "TEXT DEFAULT ''")
    add_column("chemical_warehouses", "service_open", "INTEGER DEFAULT 1")
    add_column("chemicals", "auth_expires_at", "TEXT DEFAULT ''")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_lab_pdf_expiry ON laboratory_reservations(pdf_expires_at) WHERE pdf_path<>''")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_chemical_pdf_expiry ON chemical_withdrawals(pdf_expires_at) WHERE pdf_path<>''")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_chemical_withdrawal_no ON chemical_withdrawals(withdrawal_no)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_users_mentor_user ON users(mentor_user_id,active)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_users_username_global ON users(username,active)")
    for account in conn.execute(
        """SELECT * FROM users
           WHERE mentor_user_id IS NULL AND role NOT IN ('导师','老师','教授','PI','teacher','tutor','mentor','超级管理员','管理员')"""
    ):
        candidates = conn.execute(
            """SELECT * FROM users WHERE active=1
               AND role IN ('导师','老师','教授','PI','teacher','tutor','mentor')
               AND (group_code=? OR display_name=?)
               ORDER BY id""",
            (account["group_code"], account["group_code"]),
        ).fetchall()
        if len(candidates) == 1:
            conn.execute(
                "UPDATE users SET mentor_user_id=?,mentor_username=? WHERE id=?",
                (candidates[0]["id"], candidates[0]["username"], account["id"]),
            )
    for reservation in conn.execute("SELECT * FROM laboratory_reservations ORDER BY id"):
        if not reservation["reservation_no"]:
            conn.execute(
                "UPDATE laboratory_reservations SET reservation_no=? WHERE id=?",
                (f"LAB-LEGACY-{reservation['id']:06d}", reservation["id"]),
            )
        if reservation["companion_user_id"] and not conn.execute(
            "SELECT 1 FROM laboratory_reservation_participants WHERE reservation_id=?",
            (reservation["id"],),
        ).fetchone():
            companion = conn.execute("SELECT * FROM users WHERE id=?", (reservation["companion_user_id"],)).fetchone()
            if companion:
                mentor = conn.execute(
                    "SELECT * FROM users WHERE active=1 AND role='导师' AND group_code=? ORDER BY id LIMIT 1",
                    (companion["group_code"],),
                ).fetchone()
                conn.execute(
                    """INSERT OR IGNORE INTO laboratory_reservation_participants(
                       reservation_id,user_id,participant_name,participant_role,teacher_name,
                       confirmation_status,confirmation_note,signature_b64,confirmed_at,created_at,updated_at
                    ) VALUES(?,?,?,?,?,'已同意','由旧版本预约记录迁移',?,?,?,?)""",
                    (reservation["id"], companion["id"], companion["display_name"] or companion["username"],
                     companion["role"], mentor["display_name"] or mentor["username"] if mentor else "",
                     "", reservation["created_at"], reservation["created_at"], reservation["updated_at"]),
                )
        if not conn.execute(
            "SELECT 1 FROM laboratory_reservation_mentor_reviews WHERE reservation_id=?",
            (reservation["id"],),
        ).fetchone():
            mentor = conn.execute(
                """SELECT * FROM users WHERE active=1 AND role='导师' AND
                   (display_name=? OR username=? OR group_code=?)
                   ORDER BY CASE WHEN display_name=? OR username=? THEN 0 ELSE 1 END,id LIMIT 1""",
                (reservation["requester_teacher"], reservation["requester_teacher"], reservation["requester_group"],
                 reservation["requester_teacher"], reservation["requester_teacher"]),
            ).fetchone()
            if mentor:
                old_status = "已批准" if reservation["mentor_status"] == "已批准" else (
                    "已驳回" if reservation["mentor_status"] == "已驳回" else "待审核"
                )
                conn.execute(
                    """INSERT OR IGNORE INTO laboratory_reservation_mentor_reviews(
                       reservation_id,mentor_user_id,mentor_name,status,review_note,signature_b64,
                       reviewed_at,created_at,updated_at
                    ) VALUES(?,?,?,?,?,?,?,?,?)""",
                    (reservation["id"], mentor["id"], mentor["display_name"] or mentor["username"],
                     old_status, reservation["mentor_note"], reservation["mentor_signature_b64"],
                     reservation["updated_at"] if old_status != "待审核" else "",
                     reservation["created_at"], reservation["updated_at"]),
                )
    for withdrawal in conn.execute("SELECT id,withdrawal_no,created_at FROM chemical_withdrawals ORDER BY id"):
        if not withdrawal["withdrawal_no"]:
            created = str(withdrawal["created_at"] or "")
            digits = re.sub(r"\D", "", created)[:8] or datetime.now().strftime("%Y%m%d")
            conn.execute(
                "UPDATE chemical_withdrawals SET withdrawal_no=? WHERE id=?",
                (f"CHEM-{digits}-{withdrawal['id']:06d}", withdrawal["id"]),
            )
    conn.commit()
    return conn


def local_ipv4_addresses():
    addresses = set()
    try:
        for item in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            addresses.add(item[4][0])
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        addresses.add(s.getsockname()[0])
        s.close()
    except Exception:
        pass
    addresses.add("127.0.0.1")
    return sorted(addresses)


class ApiHandler(BaseHTTPRequestHandler):
    server_version = "LitSearchProGeneric/22.1"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        self.server.app_log("%s - %s" % (self.address_string(), fmt % args))

    def read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def send_json(self, code, data):
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        try:
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            self.close_connection = True

    def send_bytes(self, code, raw, content_type, filename=""):
        try:
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            if filename:
                quoted = urllib.parse.quote(filename)
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quoted}")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            self.close_connection = True

    def error_json(self, code, message):
        self.send_json(code, {"ok": False, "error": message})

    def auth_user(self, conn):
        header = self.headers.get("Authorization", "")
        token = header[7:].strip() if header.lower().startswith("bearer ") else ""
        if not token:
            return None
        row = conn.execute("SELECT users.*,tokens.last_seen AS token_last_seen FROM tokens JOIN users ON users.id=tokens.user_id WHERE tokens.token=? AND users.active=1", (token,)).fetchone()
        if row:
            cutoff = (datetime.now() - timedelta(minutes=5)).isoformat(timespec="seconds")
            if not row["token_last_seen"] or row["token_last_seen"] < cutoff:
                conn.execute("UPDATE tokens SET last_seen=? WHERE token=?", (now(), token))
                conn.commit()
        return row

    def group_from(self, data, user=None):
        return (data.get("group_code") or (user["group_code"] if user else self.server.config.get("group_code", "research-lab"))).strip()

    def normalize_role(self, role):
        text = str(role or "").strip()
        lower = text.lower()
        if any(x in text for x in ("导师", "老师", "教授", "Supervisor", "PI")) or lower in ("teacher", "tutor", "mentor", "supervisor", "pi"):
            return "导师"
        if any(x in text for x in ("超级管理员", "管理员", "管理", "Admin", "SuperAdmin")) or lower in ("admin", "superadmin", "administrator", "root"):
            return "超级管理员"
        if "合作者" in text:
            return "合作者"
        return "学生"

    def is_manager(self, user):
        return self.normalize_role(user["role"]) in ("导师", "超级管理员")

    def is_admin(self, user):
        return self.normalize_role(user["role"]) == "超级管理员"

    def is_equipment_manager(self, conn, user, equipment_row):
        role = self.normalize_role(user["role"])
        equipment_group = equipment_row["equipment_group_code"] if "equipment_group_code" in equipment_row.keys() else equipment_row["group_code"]
        if role == "超级管理员":
            return True
        if role == "导师":
            return str(user["group_code"] or "") == str(equipment_group or "")
        username = user["username"]
        if username and username in (equipment_row["manager1"], equipment_row["manager2"]):
            return True
        if username:
            row = conn.execute("SELECT 1 FROM equipment_managers WHERE group_code=? AND username=?", (equipment_group, username)).fetchone()
            if row:
                return True
        return False

    def is_group_equipment_manager(self, conn, user, group_code):
        role = self.normalize_role(user["role"])
        if role == "超级管理员":
            return True
        if role == "导师":
            return str(user["group_code"] or "") == str(group_code or "")
        username = user["username"]
        if username:
            row = conn.execute("SELECT 1 FROM equipment_managers WHERE group_code=? AND username=?", (group_code, username)).fetchone()
            if row:
                return True
        return False

    def cleanup_expired_files(self, conn):
        cutoff = now()
        rows = conn.execute("SELECT id,path FROM files WHERE expires_at<>'' AND expires_at<?", (cutoff,)).fetchall()
        for row in rows:
            try:
                if os.path.exists(row["path"]):
                    os.remove(row["path"])
            except Exception:
                pass
            conn.execute("DELETE FROM files WHERE id=?", (row["id"],))
        if rows:
            conn.commit()

    def maintain_login_sessions(self, conn, user_id):
        conn.execute(
            """DELETE FROM tokens
               WHERE user_id=? AND rowid NOT IN (
                   SELECT rowid FROM tokens WHERE user_id=?
                   ORDER BY created_at DESC,rowid DESC LIMIT 8
               )""",
            (user_id, user_id),
        )
        current = time.monotonic()
        if current - self.server.last_maintenance_at >= 3600 and self.server.maintenance_lock.acquire(blocking=False):
            try:
                if current - self.server.last_maintenance_at >= 3600:
                    cutoff = (datetime.now() - timedelta(days=30)).isoformat(timespec="seconds")
                    conn.execute("DELETE FROM tokens WHERE COALESCE(NULLIF(last_seen,''),created_at)<?", (cutoff,))
                    conn.commit()
                    conn.execute("PRAGMA wal_checkpoint(PASSIVE)")
                    self.server.last_maintenance_at = current
            finally:
                self.server.maintenance_lock.release()

    def json_list(self, value):
        try:
            parsed = json.loads(value or "[]")
            return parsed if isinstance(parsed, list) else []
        except Exception:
            return []

    def json_dict(self, value):
        try:
            parsed = json.loads(value or "{}")
            return parsed if isinstance(parsed, dict) else {}
        except Exception:
            return {}

    def user_label(self, user):
        return (user["display_name"] or user["username"]) if user else ""

    def can_approve_account(self, reviewer, target):
        if not reviewer or not target:
            return False
        if int(reviewer["id"]) == int(target["id"]):
            return False
        reviewer_role = self.normalize_role(reviewer["role"])
        target_role = self.normalize_role(target["role"])
        if reviewer_role == "超级管理员":
            return True
        if reviewer_role != "导师" or target_role != "学生":
            return False
        reviewer_id = int(reviewer["id"])
        reviewer_username = str(reviewer["username"] or "")
        reviewer_label = self.user_label(reviewer)
        target_mentor_id = target["mentor_user_id"] if "mentor_user_id" in target.keys() else None
        target_mentor_username = str(target["mentor_username"] or "") if "mentor_username" in target.keys() else ""
        if target_mentor_id and int(target_mentor_id) == reviewer_id:
            return True
        if target_mentor_username and target_mentor_username == reviewer_username:
            return True
        return False

    def user_signature(self, conn, user_id):
        row = conn.execute("SELECT image_b64 FROM user_signatures WHERE user_id=?", (user_id,)).fetchone()
        return row["image_b64"] if row else ""

    def team_mentor(self, conn, user):
        if not user:
            return None
        if self.normalize_role(user["role"]) == "导师":
            return user
        if "mentor_user_id" in user.keys() and user["mentor_user_id"]:
            mentor = conn.execute(
                "SELECT * FROM users WHERE id=? AND active=1",
                (user["mentor_user_id"],),
            ).fetchone()
            if mentor and self.normalize_role(mentor["role"]) == "导师":
                return mentor
        if "mentor_username" in user.keys() and user["mentor_username"]:
            mentor = conn.execute(
                "SELECT * FROM users WHERE username=? AND active=1 ORDER BY id LIMIT 1",
                (user["mentor_username"],),
            ).fetchone()
            if mentor and self.normalize_role(mentor["role"]) == "导师":
                return mentor
        candidates = conn.execute(
            """SELECT * FROM users
               WHERE active=1 AND (group_code=? OR (team_name<>'' AND team_name=?))
               ORDER BY CASE WHEN group_code=? THEN 0 ELSE 1 END,id""",
            (user["group_code"], user["team_name"], user["group_code"]),
        ).fetchall()
        for candidate in candidates:
            if self.normalize_role(candidate["role"]) == "导师":
                return candidate
        return None

    def refresh_lab_reservation_state(self, conn, reservation_id):
        reservation = conn.execute("SELECT * FROM laboratory_reservations WHERE id=?", (reservation_id,)).fetchone()
        if not reservation:
            return
        participants = [dict(x) for x in conn.execute(
            "SELECT * FROM laboratory_reservation_participants WHERE reservation_id=? ORDER BY id",
            (reservation_id,),
        )]
        reviews = [dict(x) for x in conn.execute(
            "SELECT * FROM laboratory_reservation_mentor_reviews WHERE reservation_id=? ORDER BY id",
            (reservation_id,),
        )]
        if any(x["confirmation_status"] == "已拒绝" for x in participants):
            status = "同行人已拒绝"
        elif any(x["confirmation_status"] == "待确认" for x in participants):
            status = "待同行人确认"
        elif any(x["status"] == "已驳回" for x in reviews):
            status = "导师已驳回"
        elif any(x["status"] == "待审核" for x in reviews):
            status = "待导师审核"
        elif reservation["manager_status"] == "已批准":
            status = "已批准"
        elif reservation["manager_status"] == "已驳回":
            status = "实验室已驳回"
        else:
            status = "待实验室审核"
        mentor_status = "已批准" if reviews and all(x["status"] == "已批准" for x in reviews) else (
            "已驳回" if any(x["status"] == "已驳回" for x in reviews) else "待导师审核"
        )
        conn.execute(
            """UPDATE laboratory_reservations
               SET status=?,mentor_status=?,participants_json=?,mentor_reviews_json=?,updated_at=?
               WHERE id=?""",
            (status, mentor_status, json.dumps(participants, ensure_ascii=False),
             json.dumps(reviews, ensure_ascii=False), now(), reservation_id),
        )

    def is_lab_manager(self, user, lab):
        if self.is_admin(user):
            return True
        if self.normalize_role(user["role"]) != "导师":
            return False
        if lab["lab_type"] == "团队实验室" and user["team_name"] and user["team_name"] == lab["team_name"]:
            return True
        return user["username"] in self.json_list(lab["managers_json"])

    def is_warehouse_manager(self, user, warehouse):
        if self.is_admin(user):
            return True
        return self.normalize_role(user["role"]) == "导师" and user["username"] in self.json_list(warehouse["managers_json"])

    def current_blacklist(self, conn, lab_id, user_id):
        current = now()
        conn.execute(
            "UPDATE laboratory_blacklist SET active=0,updated_at=? WHERE active=1 AND blacklist_type='暂停' AND ends_at<>'' AND ends_at<=?",
            (current, current),
        )
        return conn.execute(
            "SELECT * FROM laboratory_blacklist WHERE laboratory_id=? AND user_id=? AND active=1 ORDER BY id DESC LIMIT 1",
            (lab_id, user_id),
        ).fetchone()

    def lab_manager_users(self, conn, lab):
        usernames = set(self.json_list(lab["managers_json"]))
        if lab["lab_type"] == "团队实验室" and lab["team_name"]:
            usernames.update(
                x["username"] for x in conn.execute(
                    "SELECT username,role FROM users WHERE active=1 AND team_name=?",
                    (lab["team_name"],),
                ) if self.normalize_role(x["role"]) == "导师"
            )
        if not usernames:
            return []
        marks = ",".join("?" for _ in usernames)
        return [x for x in conn.execute(
            f"SELECT * FROM users WHERE active=1 AND username IN ({marks}) ORDER BY display_name,username",
            tuple(sorted(usernames)),
        ) if self.normalize_role(x["role"]) == "导师"]

    def lab_log(self, conn, reservation_id, user, action, detail=""):
        conn.execute(
            """INSERT INTO laboratory_audit_logs(
               reservation_id,action,actor_id,actor_name,actor_role,detail,created_at
            ) VALUES(?,?,?,?,?,?,?)""",
            (reservation_id, action, user["id"] if user else None,
             self.user_label(user) if user else "系统",
             self.normalize_role(user["role"]) if user else "系统", detail, now()),
        )

    def chemical_log(self, conn, withdrawal_id, user, action, detail=""):
        conn.execute(
            "INSERT INTO chemical_audit_logs(withdrawal_id,action,actor_id,actor_name,detail,created_at) VALUES(?,?,?,?,?,?)",
            (withdrawal_id, action, user["id"] if user else None,
             self.user_label(user) if user else "系统", detail, now()),
        )

    def cleanup_expired_approval_pdfs(self, conn):
        global PDF_CLEANUP_LAST
        monotonic_now = time.monotonic()
        if monotonic_now - PDF_CLEANUP_LAST < PDF_CLEANUP_INTERVAL_SECONDS:
            return
        if not PDF_CLEANUP_LOCK.acquire(blocking=False):
            return
        try:
            monotonic_now = time.monotonic()
            if monotonic_now - PDF_CLEANUP_LAST < PDF_CLEANUP_INTERVAL_SECONDS:
                return
            PDF_CLEANUP_LAST = monotonic_now
            current = now()
            for table in ("laboratory_reservations", "chemical_withdrawals"):
                for row in conn.execute(
                    f"SELECT id,pdf_path FROM {table} WHERE pdf_path<>'' AND pdf_expires_at<>'' AND pdf_expires_at<=?",
                    (current,),
                ):
                    try:
                        if os.path.isfile(row["pdf_path"]):
                            os.remove(row["pdf_path"])
                    except OSError:
                        continue
                    conn.execute(f"UPDATE {table} SET pdf_path='' WHERE id=?", (row["id"],))
        finally:
            PDF_CLEANUP_LOCK.release()

    def cleanup_expired_blacklists(self, conn):
        global BLACKLIST_CLEANUP_LAST
        monotonic_now = time.monotonic()
        if monotonic_now - BLACKLIST_CLEANUP_LAST < BLACKLIST_CLEANUP_INTERVAL_SECONDS:
            return
        if not BLACKLIST_CLEANUP_LOCK.acquire(blocking=False):
            return
        try:
            monotonic_now = time.monotonic()
            if monotonic_now - BLACKLIST_CLEANUP_LAST < BLACKLIST_CLEANUP_INTERVAL_SECONDS:
                return
            BLACKLIST_CLEANUP_LAST = monotonic_now
            current = now()
            conn.execute(
                """UPDATE laboratory_blacklist SET active=0,updated_at=?
                   WHERE active=1 AND blacklist_type='暂停' AND ends_at<>'' AND ends_at<=?""",
                (current, current),
            )
        finally:
            BLACKLIST_CLEANUP_LOCK.release()

    def archive_path(self, kind, record_id, title):
        folder = os.path.join(ARCHIVE_DIR, kind)
        os.makedirs(folder, exist_ok=True)
        safe = "".join(ch if ch.isalnum() or "\u4e00" <= ch <= "\u9fff" else "_" for ch in title)[:80]
        return os.path.join(folder, f"{record_id}_{safe}.pdf")

    def pdf_font(self):
        for path in (
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            r"C:\Windows\Fonts\simsun.ttc",
        ):
            if os.path.exists(path):
                return path
        return None

    def render_pdf(self, title, sections, signatures, target):
        if fitz is None:
            raise RuntimeError("PDF generation component is unavailable")
        doc = fitz.open()
        fontfile = self.pdf_font()
        fontname = "genericfont" if fontfile else "helv"
        page = doc.new_page(width=595, height=842)
        if fontfile:
            page.insert_font(fontname=fontname, fontfile=fontfile)
        y = 38

        def add_text(text, size=10, bold=False, gap=8):
            nonlocal page, y
            text = str(text or "")
            lines = max(1, len(text) // 42 + text.count("\n") + 1)
            height = max(38 if bold else 28, lines * (size + 10))
            if y + height > 785:
                page = doc.new_page(width=595, height=842)
                if fontfile:
                    page.insert_font(fontname=fontname, fontfile=fontfile)
                y = 38
            page.insert_textbox(
                fitz.Rect(42, y, 553, y + height),
                text,
                fontsize=size,
                fontname=fontname,
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_CENTER if bold else fitz.TEXT_ALIGN_LEFT,
            )
            y += height + gap

        add_text(title, 16, True, 14)
        add_text(f"生成时间：{now()}", 9, False, 8)
        for heading, content in sections:
            add_text(heading, 12, True, 3)
            if isinstance(content, (list, tuple)):
                for index, row in enumerate(content, 1):
                    if isinstance(row, dict):
                        add_text(f"{index}. " + "；".join(f"{k}：{v}" for k, v in row.items()), 9, False, 2)
                    else:
                        add_text(f"{index}. {row}", 9, False, 2)
            else:
                add_text(content or "无", 9, False, 6)
        add_text("电子签名", 12, True, 6)
        for label, name, image_b64 in signatures:
            if y + 80 > 785:
                page = doc.new_page(width=595, height=842)
                if fontfile:
                    page.insert_font(fontname=fontname, fontfile=fontfile)
                y = 38
            page.insert_text((42, y + 18), f"{label}：{name}", fontsize=10, fontname=fontname)
            if image_b64:
                try:
                    raw = base64.b64decode(image_b64)
                    page.insert_image(fitz.Rect(250, y, 410, y + 58), stream=raw, keep_proportion=True)
                except Exception:
                    page.insert_text((250, y + 18), "签名图片无法解析", fontsize=9, fontname=fontname)
            else:
                page.insert_text((250, y + 18), "未上传签名", fontsize=9, fontname=fontname)
            y += 68
        metadata = doc.metadata
        metadata["title"] = title
        metadata["subject"] = "永久安全审批归档"
        doc.set_metadata(metadata)
        doc.save(target, garbage=4, deflate=True)
        doc.close()

    def generate_lab_pdf(self, conn, reservation_id):
        row = conn.execute(
            """SELECT r.*,l.name AS lab_name,l.college,l.address,l.lab_type,l.team_name,
                      u.username,u.display_name
               FROM laboratory_reservations r
               JOIN laboratories l ON l.id=r.laboratory_id
               JOIN users u ON u.id=r.requester_id WHERE r.id=?""",
            (reservation_id,),
        ).fetchone()
        if not row:
            return ""
        title = f"{row['college']}{row['lab_name']}预约清单"
        safe_name = "".join(ch if ch.isalnum() or "\u4e00" <= ch <= "\u9fff" else "_" for ch in (row["display_name"] or row["username"]))
        folder = os.path.join(ARCHIVE_DIR, "laboratory");os.makedirs(folder,exist_ok=True)
        target = os.path.join(folder, f"{safe_name}_{row['reservation_no'] or reservation_id}.pdf")
        participants = [dict(x) for x in conn.execute(
            "SELECT * FROM laboratory_reservation_participants WHERE reservation_id=? ORDER BY id",
            (reservation_id,),
        )]
        mentors = [dict(x) for x in conn.execute(
            "SELECT * FROM laboratory_reservation_mentor_reviews WHERE reservation_id=? ORDER BY id",
            (reservation_id,),
        )]

        def esc(value):
            value = str(value or "")
            output = []
            position = 0
            for match in re.finditer(r"[A-Za-z0-9][A-Za-z0-9._:/()+\-]*", value):
                output.append(html.escape(value[position:match.start()]))
                output.append(f'<span class="en">{html.escape(match.group(0))}</span>')
                position = match.end()
            output.append(html.escape(value[position:]))
            return "".join(output).replace("\n", "<br>")

        def data_table(records, empty_text="无"):
            records = records or []
            if not records:
                return f"<p>{esc(empty_text)}</p>"
            headers = []
            for record in records:
                if isinstance(record, dict):
                    for key in record:
                        if key not in headers:
                            headers.append(key)
            if not headers:
                return "<table><tbody>" + "".join(f"<tr><td>{esc(x)}</td></tr>" for x in records) + "</tbody></table>"
            head = "".join(f"<th>{esc(key)}</th>" for key in headers)
            body = "".join(
                "<tr>" + "".join(f"<td>{esc(record.get(key, ''))}</td>" for key in headers) + "</tr>"
                for record in records
            )
            return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"

        signature_overlays = []

        def signature_cell(label, name, image_b64):
            marker = f"SIGMARK{len(signature_overlays):03d}"
            signature_overlays.append((marker, image_b64))
            state = f'<span class="sigmark">{marker}</span>' if image_b64 else "未上传签名"
            return f"<td><b>{esc(label)}</b><br>{esc(name)}<br>{state}</td>"

        participant_text = "、".join(
            f"{x['participant_name']}（{x['participant_role']}，{x['confirmation_status']}）"
            for x in participants
        ) or "无"
        basic_rows = [
            ("预约单编号", row["reservation_no"] or f"LAB-{reservation_id:06d}"),
            ("申请人", row["display_name"] or row["username"]),
            ("联系电话", row["requester_phone"]),
            ("申请人导师", row["requester_teacher"]),
            ("实验室", f"{row['lab_name']}（{row['lab_type']}）"),
            ("实验室地址", row["address"]),
            ("预约时间", f"{row['start_time']} 至 {row['end_time']}"),
            ("同行人员", participant_text),
            ("实验目的", row["purpose"]),
        ]
        basic_table = "<table class='basic'><tbody>" + "".join(
            f"<tr><th>{esc(label)}</th><td>{esc(value)}</td></tr>" for label, value in basic_rows
        ) + "</tbody></table>"
        student_signatures = [
            ("预约学生", row["display_name"] or row["username"], row["student_signature_b64"])
        ] + [
            ("同行人", x["participant_name"], x["signature_b64"]) for x in participants
        ]
        mentor_signatures = [
            ("导师", x["mentor_name"], x["signature_b64"]) for x in mentors
        ]
        sign_rows = ""
        signatures = student_signatures + mentor_signatures
        for index in range(0, len(signatures), 2):
            cells = signatures[index:index + 2]
            sign_rows += "<tr>" + "".join(signature_cell(*item) for item in cells)
            if len(cells) == 1:
                sign_rows += "<td></td>"
            sign_rows += "</tr>"
        mentor_notes = data_table([
            {"导师": x["mentor_name"], "审核状态": x["status"], "审批意见": x["review_note"], "审核时间": x["reviewed_at"]}
            for x in mentors
        ])
        manager_signature = signature_cell(
            "实验室管理员", row["manager_reviewer"], row["manager_signature_b64"]
        )
        simsun = r"C:\Windows\Fonts\simsun.ttc"
        times = r"C:\Windows\Fonts\times.ttf"
        font_css = f"@font-face{{font-family:Song;src:url({simsun.replace(os.sep, '/')});}}" if os.path.exists(simsun) else ""
        if os.path.exists(times):
            font_css += f"@font-face{{font-family:Roman;src:url({times.replace(os.sep, '/')});}}"
        css = font_css + """
        @page { size:A4; margin:18mm 16mm 17mm 16mm; }
        body { font-family:Song; font-size:10.5pt; color:#111; line-height:1.55; }
        .en { font-family:Roman; }
        h1 { text-align:center; font-size:18pt; margin:0 0 12pt 0; }
        h2 { font-size:13pt; margin:12pt 0 6pt 0; }
        h1,h2,th,b,strong { font-weight:normal; text-shadow:0.25pt 0 0 #111,-0.25pt 0 0 #111; }
        .meta { text-align:center; color:#555; margin-bottom:10pt; }
        table { width:100%; border-collapse:collapse; table-layout:fixed; margin:5pt 0 11pt 0; }
        th,td { border:0.7pt solid #333; padding:5pt; vertical-align:top; }
        th { background:#eef3f8; text-align:center; }
        .basic th { width:105pt; white-space:nowrap; }
        .basic td { width:390pt; }
        .page-break { page-break-before:always; }
        .commitment { white-space:normal; text-align:justify; min-height:420pt; }
        .signature-table td { width:50%; height:72pt; text-align:center; }
        .sigmark { color:#ffffff; font-size:1pt; }
        """
        document_html = f"""
        <h1>{esc(title)}</h1>
        <div class="meta">预约单编号：{esc(row["reservation_no"])}　生成时间：{esc(now())}</div>
        <h2>一、预约基本信息</h2>{basic_table}
        <h2>二、实验流程</h2>{data_table(self.json_list(row["workflow_json"]))}
        <h2>三、化学品使用</h2>{data_table(self.json_list(row["chemicals_json"]))}
        <h2>四、实验安全预案</h2>{data_table(self.json_list(row["safety_plan_json"]))}
        <h1 class="page-break">实验室安全承诺书</h1>
        <div class="meta">预约单编号：{esc(row["reservation_no"])}</div>
        <div class="commitment">{esc(row["commitment_snapshot"])}</div>
        <h2>学生、同行人与导师电子签名</h2>
        <table class="signature-table">{sign_rows}</table>
        <h1 class="page-break">审批意见</h1>
        <div class="meta">预约单编号：{esc(row["reservation_no"])}</div>
        <h2>导师审批记录</h2>{mentor_notes}
        <h2>实验室管理员审批记录</h2>
        <table class="basic"><tbody>
          <tr><th>审核状态</th><td>{esc(row["manager_status"])}</td></tr>
          <tr><th>审核人</th><td>{esc(row["manager_reviewer"])}</td></tr>
          <tr><th>审批意见</th><td>{esc(row["manager_note"])}</td></tr>
        </tbody></table>
        <table class="signature-table"><tr>{manager_signature}<td></td></tr></table>
        <h2>电子归档证书</h2>
        <table class="basic"><tbody>
          <tr><th>证书主体</th><td>科研文献与实验室安全管理平台实验室安全审批归档证书</td></tr>
          <tr><th>预约单号</th><td>{esc(row["reservation_no"])}</td></tr>
          <tr><th>签发时间</th><td>{esc(now())}</td></tr>
          <tr><th>完整性标识</th><td>{esc(hashlib.sha256((str(reservation_id)+row["reservation_no"]+row["updated_at"]).encode()).hexdigest())}</td></tr>
        </tbody></table>
        """
        story = fitz.Story(document_html, user_css=css)
        writer = fitz.DocumentWriter(target)
        mediabox = fitz.paper_rect("a4")
        content_rect = mediabox + (45, 45, -45, -45)
        story.write(writer, lambda _number, _filled: (mediabox, content_rect, None))
        writer.close()
        del writer
        del story
        gc.collect()
        document = fitz.open(target)
        marker_positions = {}
        for marker, image_b64 in signature_overlays:
            if not image_b64:
                continue
            for page in document:
                found = page.search_for(marker)
                if not found:
                    continue
                anchor = found[0]
                marker_positions.setdefault(page.number, []).append(anchor)
                raw = base64.b64decode(image_b64)
                width = 110
                page.insert_image(
                    fitz.Rect(anchor.x0, anchor.y1 + 2, anchor.x0 + width, anchor.y1 + 38),
                    stream=raw,
                    keep_proportion=True,
                    overlay=True,
                )
                break
        for page_number, anchors in marker_positions.items():
            page = document[page_number]
            clear_from = max(x.y1 for x in anchors) + (56 if len(anchors) == 1 else 74)
            if clear_from < page.rect.height - 36:
                page.draw_rect(
                    fitz.Rect(38, clear_from, page.rect.width - 38, page.rect.height - 36),
                    color=(1, 1, 1),
                    fill=(1, 1, 1),
                    overlay=True,
                )
        song_font = fitz.Font(fontfile=simsun) if os.path.exists(simsun) else fitz.Font("china-s")
        roman_font = fitz.Font(fontfile=times) if os.path.exists(times) else fitz.Font("tiro")

        def redraw_bold_title(page, title_text):
            matches = page.search_for(title_text)
            if not matches:
                return
            original = matches[0]
            page.draw_rect(
                fitz.Rect(38, original.y0 - 2, page.rect.width - 38, original.y1 + 3),
                color=(1, 1, 1), fill=(1, 1, 1), overlay=True,
            )
            page.insert_font(fontname="songtitle", fontfile=simsun if os.path.exists(simsun) else None)
            page.insert_font(fontname="romantitle", fontfile=times if os.path.exists(times) else None)
            runs = []
            for character in title_text:
                kind = "roman" if ord(character) < 128 else "song"
                if runs and runs[-1][0] == kind:
                    runs[-1] = (kind, runs[-1][1] + character)
                else:
                    runs.append((kind, character))
            size = 18
            total = sum(
                (roman_font if kind == "roman" else song_font).text_length(value, fontsize=size)
                for kind, value in runs
            )
            start_x = (page.rect.width - total) / 2
            baseline = original.y1 - 1
            for offset in (0, 0.32):
                x = start_x + offset
                for kind, value in runs:
                    font = roman_font if kind == "roman" else song_font
                    page.insert_text(
                        (x, baseline + offset), value, fontsize=size,
                        fontname="romantitle" if kind == "roman" else "songtitle",
                        color=(0, 0, 0), overlay=True,
                    )
                    x += font.text_length(value, fontsize=size)

        for page in document:
            for heading in (title, "实验室安全承诺书", "审批意见"):
                if page.search_for(heading):
                    redraw_bold_title(page, heading)
        document.subset_fonts()
        optimized = target + ".optimized.pdf"
        document.save(optimized, garbage=4, deflate=True, clean=True)
        document.close()
        del document
        gc.collect()
        replaced = False
        for _attempt in range(20):
            try:
                os.replace(optimized, target)
                replaced = True
                break
            except PermissionError:
                time.sleep(0.1)
        if not replaced:
            raise PermissionError(f"无法替换优化后的预约 PDF：{target}")
        expires = (datetime.now() + timedelta(days=7)).isoformat(timespec="seconds")
        conn.execute(
            "UPDATE laboratory_reservations SET pdf_path=?,pdf_expires_at=?,certificate_text=?,updated_at=? WHERE id=?",
            (target, expires, "科研文献与实验室安全管理平台实验室安全审批归档证书", now(), reservation_id),
        )
        return target

    def generate_chemical_pdf(self, conn, withdrawal_id):
        row = conn.execute(
            """SELECT w.*,c.name AS chemical_name,c.unit,c.owner_teacher,
                      wh.name AS warehouse_name,wh.college,wh.address,wh.managers_json,
                      u.username,u.display_name
               FROM chemical_withdrawals w JOIN chemicals c ON c.id=w.chemical_id
               JOIN chemical_warehouses wh ON wh.id=c.warehouse_id
               JOIN users u ON u.id=w.requester_id WHERE w.id=?""",
            (withdrawal_id,),
        ).fetchone()
        if not row:
            return ""
        title = f"{row['college']}危险化学品领用申请单"
        withdrawal_no = row["withdrawal_no"] or f"CHEM-{datetime.now():%Y%m%d}-{withdrawal_id:06d}"
        if not row["withdrawal_no"]:
            conn.execute("UPDATE chemical_withdrawals SET withdrawal_no=?,updated_at=? WHERE id=?", (withdrawal_no, now(), withdrawal_id))
        applicant_name = row["display_name"] or row["username"]
        safe_name = "".join(ch if ch.isalnum() or "\u4e00" <= ch <= "\u9fff" else "_" for ch in applicant_name)
        folder = os.path.join(ARCHIVE_DIR, "chemicals"); os.makedirs(folder, exist_ok=True)
        target = os.path.join(folder, f"{safe_name}_{withdrawal_no}.pdf")
        notes = self.json_dict(row["review_notes_json"])
        managers = self.json_list(row["managers_json"])
        manager1 = managers[0] if managers else ""
        manager2 = managers[1] if len(managers) > 1 else ""
        manager_names = {}
        if managers:
            placeholders = ",".join("?" for _ in managers)
            for manager_row in conn.execute(f"SELECT username,display_name FROM users WHERE username IN ({placeholders})", tuple(managers)):
                manager_names[manager_row["username"]] = self.user_label(manager_row)
        manager1_name = manager_names.get(manager1, manager1)
        manager2_name = manager_names.get(manager2, manager2)
        co=conn.execute("SELECT * FROM chemical_withdrawal_participants WHERE withdrawal_id=?",(withdrawal_id,)).fetchone()
        def esc(value):
            value = str(value or "")
            output = []
            position = 0
            for match in re.finditer(r"[A-Za-z0-9][A-Za-z0-9._:/()+\-]*", value):
                output.append(html.escape(value[position:match.start()]))
                output.append(f'<span class="en">{html.escape(match.group(0))}</span>')
                position = match.end()
            output.append(html.escape(value[position:]))
            return "".join(output).replace("\n", "<br>")
        basic_rows = [
            ("领用单号", withdrawal_no),
            ("申请人", applicant_name),
            ("共同领用人", row["co_collector_name"]),
            ("申请人导师", row["requester_teacher"]),
            ("化学品", row["chemical_name"]),
            ("领用数量", f"{row['quantity']} {row['unit']}"),
            ("存放位置", row["storage_location"]),
            ("化学品归属导师", row["owner_teacher"]),
            ("库房", f"{row['warehouse_name']}｜{row['address']}"),
            ("用途", row["purpose"]),
        ]
        basic_table = "<table class='basic'><tbody>" + "".join(
            f"<tr><th>{esc(label)}</th><td>{esc(value)}</td></tr>" for label, value in basic_rows
        ) + "</tbody></table>"
        review_rows = [
            {"审批环节": "共同领用人确认", "审核人": row["co_collector_name"], "状态": row["participant_status"], "审批意见": co["confirmation_note"] if co else "", "审核时间": co["confirmed_at"] if co else ""},
            {"审批环节": "申请人导师审核", "审核人": row["requester_teacher"], "状态": row["mentor_status"], "审批意见": notes.get("申请人导师", ""), "审核时间": ""},
        ]
        if row["owner_status"] != "无需审核":
            review_rows.append({"审批环节": "化学品归属导师审核", "审核人": row["owner_teacher"], "状态": row["owner_status"], "审批意见": notes.get("化学品归属导师", ""), "审核时间": ""})
        review_rows += [
            {"审批环节": "库房管理员一审核", "审核人": manager1_name, "状态": row["manager1_status"], "审批意见": notes.get("库房管理员一", ""), "审核时间": ""},
            {"审批环节": "库房管理员二审核", "审核人": manager2_name, "状态": row["manager2_status"], "审批意见": notes.get("库房管理员二", ""), "审核时间": ""},
            {"审批环节": "出库与处置", "审核人": "", "状态": f"{row['status']}；{row['disposal_status']}", "审批意见": (row["disposal_report"] or "") + (f"；库管确认意见：{notes.get('处置确认')}" if notes.get("处置确认") else ""), "审核时间": row["disposed_at"]},
        ]
        review_table = "<table><thead><tr><th>审批环节</th><th>审核人</th><th>状态</th><th>审批意见/报告</th><th>时间</th></tr></thead><tbody>" + "".join(
            "<tr>" + "".join(f"<td>{esc(record.get(key, ''))}</td>" for key in ("审批环节", "审核人", "状态", "审批意见", "审核时间")) + "</tr>"
            for record in review_rows
        ) + "</tbody></table>"
        signature_overlays = []
        def signature_cell(label, name, image_b64):
            marker = f"CHEMSIG{len(signature_overlays):03d}"
            signature_overlays.append((marker, image_b64))
            state = f'<span class="sigmark">{marker}</span>' if image_b64 else "未上传签名"
            return f"<td><b>{esc(label)}</b><br>{esc(name)}<br>{state}</td>"
        owner_signature = row["owner_signature_b64"] or (row["mentor_signature_b64"] if row["owner_teacher"] == row["requester_teacher"] or row["owner_status"] == "无需审核" else "")
        signatures = [
            ("学生", applicant_name, row["student_signature_b64"]),
            ("共同领用人", row["co_collector_name"], co["signature_b64"] if co else ""),
            ("申请人导师", row["requester_teacher"], row["mentor_signature_b64"]),
            ("化学品归属导师", row["owner_teacher"], owner_signature),
            ("库房管理员一", manager1_name, row["manager1_signature_b64"]),
            ("库房管理员二", manager2_name, row["manager2_signature_b64"]),
        ]
        sign_rows = ""
        for index in range(0, len(signatures), 2):
            cells = signatures[index:index + 2]
            sign_rows += "<tr>" + "".join(signature_cell(*item) for item in cells)
            if len(cells) == 1:
                sign_rows += "<td></td>"
            sign_rows += "</tr>"
        simsun = r"C:\Windows\Fonts\simsun.ttc"
        times = r"C:\Windows\Fonts\times.ttf"
        font_css = f"@font-face{{font-family:Song;src:url({simsun.replace(os.sep, '/')});}}" if os.path.exists(simsun) else ""
        if os.path.exists(times):
            font_css += f"@font-face{{font-family:Roman;src:url({times.replace(os.sep, '/')});}}"
        css = font_css + """
        @page { size:A4; margin:18mm 16mm 17mm 16mm; }
        body { font-family:Song; font-size:10.5pt; color:#111; line-height:1.55; }
        .en { font-family:Roman; }
        h1 { text-align:center; font-size:18pt; margin:0 0 12pt 0; }
        h2 { font-size:13pt; margin:12pt 0 6pt 0; }
        h1,h2,th,b,strong { font-weight:normal; text-shadow:0.25pt 0 0 #111,-0.25pt 0 0 #111; }
        .meta { text-align:center; color:#555; margin-bottom:10pt; }
        table { width:100%; border-collapse:collapse; table-layout:fixed; margin:5pt 0 11pt 0; }
        th,td { border:0.7pt solid #333; padding:5pt; vertical-align:top; }
        th { background:#eef3f8; text-align:center; }
        .basic th { width:105pt; white-space:nowrap; }
        .basic td { width:390pt; }
        .page-break { page-break-before:always; }
        .commitment { white-space:normal; text-align:justify; min-height:420pt; }
        .signature-table td { width:50%; height:72pt; text-align:center; }
        .sigmark { color:#ffffff; font-size:1pt; }
        """
        archive_hash = hashlib.sha256((str(withdrawal_id) + withdrawal_no + str(row["updated_at"])).encode()).hexdigest()
        document_html = f"""
        <h1>{esc(title)}</h1>
        <div class="meta">领用单号：{esc(withdrawal_no)}　生成时间：{esc(now())}</div>
        <h2>一、领用基本信息</h2>{basic_table}
        <h1 class="page-break">危险化学品使用承诺书</h1>
        <div class="meta">领用单号：{esc(withdrawal_no)}</div>
        <div class="commitment">{esc(row["commitment_snapshot"])}</div>
        <h2>学生、共同领用人与审批人电子签名</h2>
        <table class="signature-table">{sign_rows}</table>
        <h1 class="page-break">审批意见</h1>
        <div class="meta">领用单号：{esc(withdrawal_no)}</div>
        <h2>审批记录</h2>{review_table}
        <h2>电子归档证书</h2>
        <table class="basic"><tbody>
          <tr><th>证书主体</th><td>科研文献与实验室安全管理平台危险化学品审批归档证书</td></tr>
          <tr><th>领用单号</th><td>{esc(withdrawal_no)}</td></tr>
          <tr><th>签发时间</th><td>{esc(now())}</td></tr>
          <tr><th>完整性标识</th><td>{esc(archive_hash)}</td></tr>
        </tbody></table>
        """
        story = fitz.Story(document_html, user_css=css)
        writer = fitz.DocumentWriter(target)
        mediabox = fitz.paper_rect("a4")
        content_rect = mediabox + (45, 45, -45, -45)
        story.write(writer, lambda _number, _filled: (mediabox, content_rect, None))
        writer.close(); del writer; del story; gc.collect()
        document = fitz.open(target)
        for marker, image_b64 in signature_overlays:
            if not image_b64:
                continue
            for page in document:
                found = page.search_for(marker)
                if not found:
                    continue
                anchor = found[0]
                page.insert_image(fitz.Rect(anchor.x0, anchor.y1 + 2, anchor.x0 + 110, anchor.y1 + 38), stream=base64.b64decode(image_b64), keep_proportion=True, overlay=True)
                break
        document.subset_fonts()
        optimized = target + ".optimized.pdf"
        document.save(optimized, garbage=4, deflate=True, clean=True)
        document.close(); del document; gc.collect()
        os.replace(optimized, target)
        expires = (datetime.now() + timedelta(days=7)).isoformat(timespec="seconds")
        conn.execute("UPDATE chemical_withdrawals SET pdf_path=?,pdf_expires_at=?,updated_at=? WHERE id=?", (target, expires, now(), withdrawal_id))
        return target

    def laboratory_web_page(self):
        return """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>科研安全轻量门户</title>
<style>
body{font-family:"Microsoft YaHei",sans-serif;background:#f4f7fb;color:#172033;margin:0}.wrap{max-width:1120px;margin:24px auto;padding:0 16px}
.card{background:white;border:1px solid #dce5f0;border-radius:14px;padding:20px;box-shadow:0 8px 28px #2342a514;margin-bottom:14px}.subcard{background:#f8fbff;border:1px solid #dce5f0;border-radius:12px;padding:14px;margin:10px 0}
h1{font-size:24px;color:#014a91}.muted{color:#607087;font-size:13px}.row{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}
label{display:block;font-weight:700;margin:10px 0 5px}input,textarea,select,button{box-sizing:border-box;width:100%;padding:10px;border:1px solid #bdcada;border-radius:8px;font-size:14px}
button{background:#02529f;color:white;border:0;font-weight:700;cursor:pointer;margin-top:8px}button.secondary{background:#eef3f8;color:#014a91;border:1px solid #bed0e6}.tabs{display:flex;flex-wrap:wrap;gap:6px;margin:12px 0;position:sticky;top:0;background:white;padding:8px 0;z-index:5}.tabs button{width:auto;padding:9px 14px}
.panel{display:none}.panel.active{display:block}.item{padding:12px;border:1px solid #dce5f0;border-radius:9px;margin:8px 0}.ok{color:#08783e}.err{color:#b42318}
@media(max-width:720px){.row{grid-template-columns:1fr}}
</style></head><body><div class="wrap"><div class="card">
<h1>科研安全轻量门户</h1><p class="muted">覆盖账号注册与审批、电子签名、实验器材借用、实验室预约及危险化学品领用。</p>
<div class="row"><div><label>账号</label><input id="username"></div><div><label>密码</label><input id="password" type="password"></div></div>
<label>导师姓名/团队代码</label><input id="group"><button onclick="login()">登录</button>
<details><summary>没有账号？提交注册申请</summary><div class="row"><div><label>用户名</label><input id="regUser"></div><div><label>真实姓名</label><input id="regName"></div><div><label>密码</label><input id="regPass" type="password"></div><div><label>导师姓名</label><input id="regAdvisor"></div></div><button onclick="register()">提交学生账号注册</button></details>
<div id="message"></div>
<div class="tabs"><button onclick="tab('signature')">电子签名</button><button onclick="tab('equipment')">实验器材</button><button onclick="tab('laboratory')">实验室预约</button><button onclick="tab('mentor')">导师确认</button><button onclick="tab('chemical')">危险化学品</button><button onclick="tab('accounts')">账号审批</button></div>
<div id="signature" class="panel"><h2>电子签名上传</h2><input id="signatureFile" type="file" accept="image/*"><button onclick="uploadSignature()">上传电子签名</button></div>
<div id="equipment" class="panel"><h2>实验器材借用</h2><button onclick="loadEquipment()">刷新器材</button><div id="equipmentRecords"></div></div>
<div id="laboratory" class="panel active"><h2>实验室预约与结束报告</h2><button onclick="loadLabs()">刷新实验室与记录</button><div id="labList"></div><div class="subcard"><h3>提交预约</h3><p class="muted" id="chosenLab">请先在上方选择实验室。时间请写清楚日期并具体到几点，例如 2026-07-01 09:00。</p><div class=row><div><label>开始时间</label><input id="labStart" placeholder="2026-07-01 09:00"></div><div><label>结束时间</label><input id="labEnd" placeholder="2026-07-01 12:00"></div><div><label>联系电话</label><input id="labPhone"></div><div><label>同行人用户名（逗号分隔，可跨团队）</label><input id="labParticipants"></div></div><label>实验目的</label><textarea id="labPurpose"></textarea><label>实验流程</label><textarea id="labWorkflow" placeholder="每行一个步骤"></textarea><label>实验安全预案</label><textarea id="labSafety" placeholder="每行一个风险和预案"></textarea><label><input id="labCommitmentSigned" type="checkbox" style="width:auto"> 我已阅读并签署实验室安全承诺书，同意使用当前电子签名。</label><button onclick="submitLabReservation()">提交实验室预约</button></div><h3>正在进行/待结束确认</h3><div id="records"></div></div>
<div id="mentor" class="panel"><h2>导师确认入口</h2><p class="muted">导师可在这里处理学生实验室预约审核。批准前请确认已上传电子签名。</p><button onclick="loadMentorReviews()">刷新待导师审核</button><div id="mentorRecords"></div></div>
<div id="chemical" class="panel"><h2>危险化学品领用与处置报告</h2><button onclick="loadChemicals()">刷新化学品与领用记录</button><div id="chemicalList"></div><div class="subcard"><h3>提交领用</h3><p class="muted" id="chosenChemical">请先选择化学品。</p><div class=row><div><label>数量</label><input id="chemQty"></div><div><label>共同领用人用户名</label><input id="chemCo"></div><div><label>领用后存放位置</label><input id="chemStorage"></div></div><label>用途</label><textarea id="chemPurpose"></textarea><button onclick="submitChemical()">提交危化品领用</button></div><h3>领用记录</h3><div id="chemicalRecords"></div></div>
<div id="accounts" class="panel"><h2>账号审批</h2><button onclick="loadAccounts()">刷新待审批账号</button><div id="accountRecords"></div></div>
</div></div><script>
let token="", selectedLabId=null, selectedChemicalId=null;
async function api(path,data){let r=await fetch(path,{method:"POST",headers:{"Content-Type":"application/json",...(token?{"Authorization":"Bearer "+token}:{})},body:JSON.stringify(data||{})});let j=await r.json();if(!r.ok)throw Error(j.error||"请求失败");return j}
function tab(id){document.querySelectorAll('.panel').forEach(x=>x.classList.remove('active'));document.getElementById(id).classList.add('active')}
async function register(){try{await api("/api/register",{username:regUser.value,display_name:regName.value,password:regPass.value,role:"学生",advisor_name:regAdvisor.value});alert("注册申请已提交")}catch(e){alert(e.message)}}
async function login(){try{let j=await api("/api/login",{username:username.value,password:password.value,group_code:group.value});token=j.token;message.innerHTML='<p class="ok">登录成功：'+j.display_name+'（'+j.role+'）</p>';load()}catch(e){message.innerHTML='<p class="err">'+e.message+'</p>'}}
async function uploadSignature(){let f=signatureFile.files[0];if(!f)return alert("请选择图片");let b=await f.arrayBuffer();let s=btoa(String.fromCharCode(...new Uint8Array(b)));await api("/api/signature",{action:"set",file_name:f.name,image_b64:s});alert("签名已上传")}
async function load(){loadLabs();loadEquipment();loadChemicals()}
function rowsFromText(text,key){return (text||"").split(/\\n+/).map(x=>x.trim()).filter(Boolean).map((x,i)=>({序号:i+1,[key]:x}))}
async function idsFromNames(text){let names=(text||"").split(/[,，\\s]+/).map(x=>x.trim()).filter(Boolean),ids=[];for(let name of names){let j=await api("/api/users/lookup",{username:name});ids.push(j.user.id)}return ids}
async function loadLabs(){try{let labs=await api("/api/laboratory/list",{});labList.innerHTML=labs.items.map(x=>`<div class=item><b>${x.name}</b>｜${x.address}｜预约通道:${x.booking_open?"开启":"关闭"}<br>管理员:${(x.manager_names||[]).join("、")}<br>${x.booking_open?`<button class=secondary onclick="selectedLabId=${x.id};chosenLab.textContent='已选择：${x.name}｜${x.address}'">选择此实验室</button>`:`<button class=secondary disabled>预约通道关闭</button>`}<small>${x.commitment_text||""}</small></div>`).join("");let j=await api("/api/laboratory/active-for-web",{});records.innerHTML=j.items.map(x=>`<div class="item"><b>${x.lab_name}</b>｜${x.reservation_no||""}<br>${x.start_time} 至 ${x.end_time}<br>地点：${x.address}<label>实验结束说明</label><textarea id="report_${x.id}"></textarea><label>安全隐患</label><textarea id="hazard_${x.id}"></textarea><button onclick="finish(${x.id})">提交实验结束报告</button></div>`).join("")||'<p>暂无正在进行的预约。</p>'}catch(e){message.innerHTML='<p class="err">'+e.message+'</p>'}}
async function submitLabReservation(){try{if(!selectedLabId)return alert("请先选择实验室");if(!labCommitmentSigned.checked)return alert("请先勾选并签署实验室安全承诺书");let ids=await idsFromNames(labParticipants.value);await api("/api/laboratory/reserve",{laboratory_id:selectedLabId,participant_user_ids:ids,phone:labPhone.value,start_time:labStart.value,end_time:labEnd.value,purpose:labPurpose.value,workbook_name:"网页端填写",workbook_b64:"",workflow:rowsFromText(labWorkflow.value,"实验流程"),chemicals:[],safety_plan:rowsFromText(labSafety.value,"安全预案"),commitment_signed:labCommitmentSigned.checked});alert("预约已提交，请等待同行人与审批。");loadLabs()}catch(e){alert(e.message)}}
async function loadMentorReviews(){try{let j=await api("/api/laboratory/requests",{scope:"mentor"});mentorRecords.innerHTML=j.items.map(x=>`<div class=item><b>${x.lab_name}</b>｜${x.reservation_no}<br>申请人:${x.display_name||x.username}｜${x.start_time} 至 ${x.end_time}<br>目的:${x.purpose||""}<label>审批意见</label><input id="mentor_note_${x.id}"><button onclick="mentorReview(${x.id},'批准')">批准</button><button class=secondary onclick="mentorReview(${x.id},'驳回')">驳回</button></div>`).join("")||"<p>暂无待导师审核。</p>"}catch(e){mentorRecords.innerHTML='<p class=err>'+e.message+'</p>'}}
async function mentorReview(id,decision){try{await api("/api/laboratory/review",{id,scope:"mentor",decision,note:document.getElementById("mentor_note_"+id).value});alert("导师审核已提交");loadMentorReviews()}catch(e){alert(e.message)}}
async function finish(id){try{await api("/api/laboratory/complete-report",{id,completion_report:document.getElementById("report_"+id).value,hazard_report:document.getElementById("hazard_"+id).value});alert("已提交，等待实验室管理员确认结束。");load()}catch(e){alert(e.message)}}
async function loadEquipment(){try{let j=await api("/api/equipment/list",{});equipmentRecords.innerHTML=j.items.map(x=>`<div class=item><b>${x.name}</b>｜${x.brand||""} ${x.model||""}<br>状态:${x.status}｜当前使用人:${x.current_user||"无"}<label>借用原因</label><input id="eq_${x.id}">${x.status==="可用"?`<button onclick="borrow(${x.id})">申请借用</button>`:`<button class=secondary disabled>正在使用，不可借用</button>`}</div>`).join("")}catch(e){equipmentRecords.innerHTML='<p class=err>'+e.message+'</p>'}}
async function borrow(id){try{await api("/api/equipment/request",{equipment_id:id,request_type:"借用",reason:document.getElementById("eq_"+id).value});alert("申请已提交")}catch(e){alert(e.message)}}
async function loadChemicals(){try{let list=await api("/api/chemical/list",{});let rec=await api("/api/chemical/withdrawals",{});chemicalList.innerHTML=list.items.map(x=>`<div class=item><b>${x.name}</b>｜库存:${x.quantity} ${x.unit}｜归属:${x.owner_teacher}<br><button class=secondary onclick="selectedChemicalId=${x.id};chosenChemical.textContent='已选择：${x.name}｜库存 ${x.quantity} ${x.unit}'">选择此化学品</button></div>`).join("");chemicalRecords.innerHTML=rec.items.map(x=>`<div class=item>${x.chemical_name}｜${x.status}｜共同领用人:${x.co_collector_name||""}｜处置:${x.disposal_status||""}${x.can_confirm?`<button onclick="confirmChemical(${x.id},'同意')">同意共同领用</button><button class=secondary onclick="confirmChemical(${x.id},'拒绝')">拒绝共同领用</button>`:""}${x.can_dispose?`<label>处置报告</label><input id="dispose_${x.id}"><button onclick="dispose(${x.id})">报告处理完毕</button>`:""}${x.can_review_disposal?`<button onclick="reviewDisposal(${x.id},'批准')">确认处置完成</button><button class=secondary onclick="reviewDisposal(${x.id},'驳回')">驳回处置报告</button>`:""}</div>`).join("")||"<p>暂无领用记录</p>"}catch(e){chemicalRecords.innerHTML='<p class=err>'+e.message+'</p>'}}
async function submitChemical(){try{if(!selectedChemicalId)return alert("请先选择化学品");await api("/api/chemical/withdraw",{chemical_id:selectedChemicalId,quantity:Number(chemQty.value),purpose:chemPurpose.value,co_collector_username:chemCo.value,storage_location:chemStorage.value});alert("领用申请已提交，等待共同领用人与审批。");loadChemicals()}catch(e){alert(e.message)}}
async function confirmChemical(id,decision){try{await api("/api/chemical/participant-confirm",{id,decision});alert("已处理共同领用确认");loadChemicals()}catch(e){alert(e.message)}}
async function dispose(id){try{await api("/api/chemical/disposal-report",{id,report:document.getElementById("dispose_"+id).value});alert("处置报告已提交");loadChemicals()}catch(e){alert(e.message)}}
async function reviewDisposal(id,decision){let note=prompt("请输入处置确认意见")||"";try{await api("/api/chemical/disposal-review",{id,decision,note});alert("处置确认已处理");loadChemicals()}catch(e){alert(e.message)}}
async function loadAccounts(){try{let j=await api("/api/users/list",{});accountRecords.innerHTML=j.items.filter(x=>!x.active).map(x=>`<div class=item>${x.display_name}｜${x.username}｜${x.role}<button onclick="approve(${x.id})">批准账号</button></div>`).join("")||"<p>暂无待审批账号</p>"}catch(e){accountRecords.innerHTML='<p class=err>'+e.message+'</p>'}}
async function approve(id){try{await api("/api/users/approve",{id,active:1});alert("账号已批准");loadAccounts()}catch(e){alert(e.message)}}
</script></body></html>"""

    def safety_portal_page(self):
        return r"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>科研安全服务门户</title>
<style>
:root{--blue:#02529f;--blue2:#01427f;--bg:#f3f5f8;--panel:#fff;--line:#dfe5ec;--text:#1f2937;--muted:#667085;--red:#c9363e;--green:#15803d;--amber:#b45309}
*{box-sizing:border-box}body{margin:0;font-family:"Microsoft YaHei UI","Microsoft YaHei",sans-serif;background:var(--bg);color:var(--text)}
button,input,textarea,select{font:inherit}.hidden{display:none!important}.shell{min-height:100vh}.topbar{height:64px;background:#fff;border-bottom:1px solid var(--line);display:flex;align-items:center;padding:0 24px;position:sticky;top:0;z-index:20}
.brand{font-size:19px;font-weight:800;color:#111827}.brand small{display:block;font-size:12px;color:var(--muted);font-weight:400}.identity{margin-left:auto;color:var(--muted);font-size:13px}.layout{display:grid;grid-template-columns:230px minmax(0,1fr);min-height:calc(100vh - 64px)}
.sidebar{background:#fff;border-right:1px solid var(--line);padding:18px 12px}.nav{width:100%;border:0;background:transparent;color:#344054;text-align:left;padding:11px 14px;border-radius:8px;margin:2px 0;cursor:pointer;font-weight:700}.nav:hover,.nav.active{background:#eaf3fb;color:var(--blue)}
.content{padding:22px;max-width:1500px;width:100%;min-width:0;margin:0 auto}.page{display:none;min-width:0}.page.active{display:block}.page-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:14px}.page-head h2{margin:0;font-size:22px}.muted{color:var(--muted);font-size:13px}
.card{min-width:0;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:18px;margin-bottom:14px;box-shadow:0 3px 14px #102a4310}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}
label{display:block;font-size:13px;font-weight:700;margin:4px 0 6px}input,textarea,select{width:100%;border:1px solid #cdd5df;border-radius:7px;background:#fff;padding:9px 10px;outline:none}input:focus,textarea:focus,select:focus{border-color:var(--blue);box-shadow:0 0 0 3px #02529f18}textarea{min-height:90px;resize:vertical}
.btn{display:inline-block;width:auto;border:1px solid transparent;border-radius:7px;padding:9px 15px;background:var(--blue);color:#fff;font-weight:700;cursor:pointer;text-decoration:none}.btn:hover{background:var(--blue2)}.btn.secondary{background:#fff;color:var(--blue);border-color:#b9cce0}.btn.danger{background:#fff1f2;color:var(--red);border-color:#f1c4c7}.btn:disabled{cursor:not-allowed;opacity:.55}
.toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:9px}table{width:100%;border-collapse:collapse;min-width:760px}th{background:#f3f6f9;color:#475467;text-align:left;font-size:13px;padding:11px;white-space:nowrap}td{padding:11px;border-top:1px solid #edf0f4;font-size:13px;vertical-align:top}tr:hover td{background:#f8fbff}.status{display:inline-block;border-radius:999px;padding:3px 9px;background:#eef2f6;color:#475467;font-size:12px}.status.ok{background:#eaf8ef;color:var(--green)}.status.warn{background:#fff7e8;color:var(--amber)}.status.bad{background:#fff0f1;color:var(--red)}
.login{max-width:760px;margin:48px auto}.login h1{margin:0 0 5px;color:#0f172a}.notice{padding:10px 12px;border-radius:7px;background:#eef6ff;color:#174a7e;margin:10px 0}.notice.error{background:#fff0f1;color:#a62f35}.empty{padding:28px;text-align:center;color:var(--muted)}
.modal-mask{position:fixed;inset:0;background:#0f172a66;display:flex;align-items:center;justify-content:center;padding:18px;z-index:50}.modal{width:min(920px,96vw);max-height:90vh;overflow:auto;background:#fff;border-radius:12px;border:1px solid var(--line);box-shadow:0 24px 80px #0004}.modal-head,.modal-foot{padding:15px 18px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:10px}.modal-head h3{margin:0}.modal-body{padding:18px}.modal-foot{border-top:1px solid var(--line);border-bottom:0;justify-content:flex-end}
@media(max-width:900px){.layout{grid-template-columns:1fr}.sidebar{display:flex;overflow:auto;border-right:0;border-bottom:1px solid var(--line);padding:8px;position:sticky;top:64px;z-index:15}.nav{white-space:nowrap;width:auto}.content{padding:14px}.grid,.grid3{grid-template-columns:1fr}.identity{display:none}}
</style></head><body>
<div id="loginView" class="login card">
  <h1>科研安全服务门户</h1><p class="muted">网页版与桌面端使用同一账号、审批数据和安全流程。</p>
  <div id="loginMessage"></div>
  <div class="grid"><div><label>账号</label><input id="loginUser"></div><div><label>密码</label><input id="loginPass" type="password"></div></div>
  <label>课题组代码（可不填）</label><input id="loginGroup"><div class="toolbar" style="margin-top:14px"><button class="btn" onclick="login()">登录</button><button class="btn secondary" onclick="openRegister()">注册账号</button></div>
</div>
<div id="appView" class="shell hidden">
  <header class="topbar"><div class="brand">科研安全服务门户<small>实验器材 · 实验室预约 · 危险化学品</small></div><div id="identity" class="identity"></div></header>
  <div class="layout"><aside class="sidebar">
    <button class="nav active" data-page="home" onclick="showPage('home',this)">首页概览</button>
    <button class="nav" data-page="signature" onclick="showPage('signature',this)">电子签名</button>
    <button class="nav" data-page="equipment" onclick="showPage('equipment',this)">实验器材</button>
    <button class="nav" data-page="laboratory" onclick="showPage('laboratory',this)">实验室预约</button>
    <button class="nav manager-nav" data-page="labReview" onclick="showPage('labReview',this)">实验室审批</button>
    <button class="nav" data-page="chemical" onclick="showPage('chemical',this)">危险化学品</button>
    <button class="nav manager-nav" data-page="accounts" onclick="showPage('accounts',this)">账号审批</button>
    <button class="nav" onclick="logout()">退出登录</button>
  </aside><main class="content">
    <section id="home" class="page active"><div class="page-head"><div><h2>首页概览</h2><p class="muted">常用安全服务快捷入口</p></div><button class="btn secondary" onclick="refreshAll()">刷新全部</button></div><div id="homeCards" class="grid3"></div></section>
    <section id="signature" class="page"><div class="page-head"><div><h2>电子签名</h2><p class="muted">审批与承诺书将使用该签名</p></div></div><div class="card"><input id="signatureFile" type="file" accept="image/*"><div class="toolbar" style="margin-top:12px"><button class="btn" onclick="uploadSignature()">上传电子签名</button></div><div id="signatureState" class="notice">尚未查询签名状态。</div></div></section>
    <section id="equipment" class="page"><div class="page-head"><div><h2>实验器材管理</h2><p class="muted">借用、归还、审批和器材维护均与桌面端使用同一套权限</p></div><button class="btn secondary" onclick="loadEquipment()">刷新</button></div>
      <div class="card"><div class="toolbar"><button id="equipmentAddButton" class="btn hidden" onclick="equipmentForm()">新增器材</button><a class="btn secondary" href="/api/equipment/template.csv">下载CSV模板</a><input id="equipmentCsv" type="file" accept=".csv,text/csv"><button id="equipmentImportButton" class="btn secondary hidden" onclick="importEquipmentCsv()">导入CSV</button><button class="btn secondary" onclick="loadMyEquipmentRequests()">我的申请</button><button class="btn secondary" onclick="loadCurrentBorrowed()">我正在借用</button><button id="equipmentReviewButton" class="btn secondary hidden" onclick="loadEquipmentReviews()">待审批申请</button></div></div>
      <div class="card"><h3>器材列表</h3><div id="equipmentTable"></div></div>
      <div id="equipmentAuxCard" class="card hidden"><div class="page-head"><div><h3 id="equipmentAuxTitle" style="margin:0">器材记录</h3></div><button class="btn secondary" onclick="$('equipmentAuxCard').classList.add('hidden')">收起</button></div><div id="equipmentAuxTable"></div></div>
    </section>
    <section id="laboratory" class="page"><div class="page-head"><div><h2>实验室预约</h2><p class="muted">选择实验室后，在一个表单中完成全部信息</p></div><button class="btn secondary" onclick="loadLabs()">刷新</button></div>
      <div class="card"><h3>实验室列表</h3><div id="labTable"></div></div>
      <div class="card"><h3>新建预约</h3><div id="chosenLab" class="notice">请先从实验室列表中选择实验室。</div><div class="grid3"><div><label>开始时间</label><input id="labStart" placeholder="2026-07-01 09:00"></div><div><label>结束时间</label><input id="labEnd" placeholder="2026-07-01 12:00"></div><div><label>联系电话</label><input id="labPhone"></div></div><p class="muted">预约时间必须写清楚日期，并精确到每一天的具体时刻，例如 2026-07-01 09:00 至 2026-07-01 12:00。</p><label>同行人用户名（可填写多位，用逗号分隔）</label><input id="labParticipants"><label>实验目的</label><textarea id="labPurpose"></textarea>
      <label>实验室预约资料工作簿（实验流程、化学品使用、实验安全预案三张表）</label><input id="labWorkbook" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onchange="parseLabWorkbook()"><div class="toolbar" style="margin-top:10px"><a class="btn secondary" href="/api/laboratory/template.xlsx">下载电子表格模板</a><button class="btn secondary" onclick="parseLabWorkbook()">重新读取工作簿</button></div><div id="labWorkbookState" class="notice">尚未上传工作簿。请下载模板，完整填写三张工作表后上传。</div><div id="labWorkbookPreview"></div><label class="notice"><input id="labCommitmentSigned" type="checkbox" style="width:auto;margin-right:8px">我已阅读并签署实验室安全承诺书，同意使用当前电子签名。</label><button class="btn" onclick="submitLabReservation()">提交预约</button></div>
      <div class="card"><h3>我的预约与结束报告</h3><div id="labRecords"></div></div>
    </section>
    <section id="labReview" class="page"><div class="page-head"><div><h2>实验室审批</h2><p class="muted">导师审核与实验室管理员审核分区展示</p></div><button class="btn secondary" onclick="loadLabReviews()">刷新</button></div><div class="card"><h3>导师待审核</h3><div id="mentorReviewTable"></div></div><div class="card"><h3>实验室管理员待审核</h3><div id="managerReviewTable"></div></div></section>
    <section id="chemical" class="page"><div class="page-head"><div><h2>危险化学品</h2><p class="muted">双人领用、导师及双库管审批</p></div><button class="btn secondary" onclick="loadChemicals()">刷新</button></div><div class="card"><h3>可领用化学品</h3><div id="chemicalTable"></div></div><div class="card"><h3>提交领用申请</h3><div id="chosenChemical" class="notice">请先选择化学品。</div><div class="grid3"><div><label>领用数量</label><input id="chemQty"></div><div><label>共同领用人用户名</label><input id="chemCo"></div><div><label>领用后存放位置</label><input id="chemStorage"></div></div><label>用途</label><textarea id="chemPurpose"></textarea><button class="btn" onclick="submitChemical()">提交领用</button></div><div class="card"><h3>领用与审批记录</h3><div id="chemicalRecords"></div></div></section>
    <section id="accounts" class="page"><div class="page-head"><div><h2>账号审批</h2><p class="muted">常规界面显示真实姓名；账号仅用于管理和区分同名用户</p></div><button class="btn secondary" onclick="loadAccounts()">刷新</button></div><div class="card"><div id="accountTable"></div></div></section>
  </main></div>
</div>
<div id="modalRoot"></div>
<script>
let token="",currentUser=null,selectedLab=null,selectedChemical=null,selectedMentor=null,labWorkbookData=null,equipmentItems=[];
const $=id=>document.getElementById(id), esc=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
async function api(path,data){let r=await fetch(path,{method:"POST",headers:{"Content-Type":"application/json",...(token?{"Authorization":"Bearer "+token}:{})},body:JSON.stringify(data||{})});let j=await r.json().catch(()=>({error:"服务器返回格式异常"}));if(!r.ok)throw Error(j.error||"请求失败");return j}
function status(v){let s=String(v||"");let c=/批准|完成|开启|可用|已同意|已入库/.test(s)?"ok":/驳回|拒绝|关闭|停止/.test(s)?"bad":"warn";return `<span class="status ${c}">${esc(s||"未知")}</span>`}
function table(headers,rows,empty="暂无记录"){if(!rows.length)return `<div class="empty">${empty}</div>`;return `<div class="table-wrap"><table><thead><tr>${headers.map(x=>`<th>${x[1]}</th>`).join("")}</tr></thead><tbody>${rows.map(r=>`<tr>${headers.map(x=>`<td>${typeof x[2]==="function"?x[2](r):esc(r[x[0]])}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`}
function saveBase64File(name,b64,mime="application/pdf"){let bytes=Uint8Array.from(atob(b64),c=>c.charCodeAt(0));let url=URL.createObjectURL(new Blob([bytes],{type:mime}));let a=document.createElement("a");a.href=url;a.download=name||"审批资料.pdf";document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)}
async function downloadPdf(entity,id){try{let j=await api("/api/safety/pdf",{entity,id});saveBase64File(j.name,j.content_b64)}catch(e){alert(e.message)}}
function showPage(id,button){document.querySelectorAll(".page").forEach(x=>x.classList.remove("active"));$(id).classList.add("active");document.querySelectorAll(".nav").forEach(x=>x.classList.remove("active"));if(button)button.classList.add("active");if(id==="signature")loadSignature();if(id==="equipment")loadEquipment();if(id==="laboratory")loadLabs();if(id==="labReview")loadLabReviews();if(id==="chemical")loadChemicals();if(id==="accounts")loadAccounts()}
function closeModal(){$("modalRoot").innerHTML=""}
function modal(title,body,foot=""){$("modalRoot").innerHTML=`<div class="modal-mask"><div class="modal"><div class="modal-head"><h3>${title}</h3></div><div class="modal-body">${body}</div><div class="modal-foot">${foot}<button class="btn secondary" onclick="closeModal()">关闭</button></div></div></div>`}
async function login(){try{let j=await api("/api/login",{username:$("loginUser").value,password:$("loginPass").value,group_code:$("loginGroup").value});token=j.token;currentUser=j;$("loginView").classList.add("hidden");$("appView").classList.remove("hidden");$("identity").textContent=`${j.display_name}｜${j.role}｜${j.team_name||"未设置团队"}`;document.querySelectorAll(".manager-nav").forEach(x=>x.classList.toggle("hidden",!["导师","超级管理员"].includes(j.role)));refreshAll()}catch(e){$("loginMessage").innerHTML=`<div class="notice error">${esc(e.message)}</div>`}}
function logout(){token="";currentUser=null;$("appView").classList.add("hidden");$("loginView").classList.remove("hidden")}
function openRegister(){selectedMentor=null;modal("注册账号",`<div class="grid"><div><label>用户名（全校唯一）</label><input id="regUser"></div><div><label>真实姓名</label><input id="regName"></div><div><label>密码</label><input id="regPass" type="password"></div><div><label>导师真实姓名</label><input id="regMentorName"></div></div><div class="toolbar" style="margin-top:12px"><button class="btn secondary" onclick="searchMentors()">查询并选择导师</button></div><div id="selectedMentor" class="notice">尚未选择导师。</div>`,`<button class="btn" onclick="register()">提交注册</button>`)}
async function searchMentors(){try{let j=await api("/api/mentors/search",{keyword:$("regMentorName").value});let h=[["display_name","导师真实姓名"],["username","导师账号（仅用于区分）"],["team_name","所属团队"],["group_code","课题组"],["op","操作",x=>`<button class="btn secondary" onclick='pickMentor(${JSON.stringify(x)})'>选择</button>`]];$("selectedMentor").innerHTML=table(h,j.items,"未找到导师")}catch(e){$("selectedMentor").innerHTML=`<span class="status bad">${esc(e.message)}</span>`}}
function pickMentor(x){selectedMentor=x;$("selectedMentor").textContent=`已选择导师：${x.display_name}｜所属团队：${x.team_name||"未设置"}`}
async function register(){if(!selectedMentor)return alert("请先查询并选择导师");try{await api("/api/register",{username:$("regUser").value,display_name:$("regName").value,password:$("regPass").value,role:"学生",advisor_name:selectedMentor.display_name,mentor_user_id:selectedMentor.id,mentor_username:selectedMentor.username,team_name:selectedMentor.team_name});alert("注册申请已提交");closeModal()}catch(e){alert(e.message)}}
async function refreshAll(){await Promise.allSettled([loadEquipment(),loadLabs(),loadChemicals(),loadSignature()]);$("homeCards").innerHTML=`<div class="card"><b>实验器材</b><p class="muted">借用与归还状态实时同步</p><button class="btn secondary" onclick="showPage('equipment',document.querySelector('[data-page=equipment]'))">进入</button></div><div class="card"><b>实验室预约</b><p class="muted">预约、同行确认及结束报告</p><button class="btn secondary" onclick="showPage('laboratory',document.querySelector('[data-page=laboratory]'))">进入</button></div><div class="card"><b>危险化学品</b><p class="muted">双人领用与处置报告</p><button class="btn secondary" onclick="showPage('chemical',document.querySelector('[data-page=chemical]'))">进入</button></div>`}
async function loadSignature(){if(!token)return;try{let j=await api("/api/signature",{});$("signatureState").textContent=j.signature?`已上传：${j.signature.file_name}｜更新时间：${j.signature.updated_at}`:"尚未上传电子签名。"}catch(e){$("signatureState").textContent=e.message}}
async function uploadSignature(){let f=$("signatureFile").files[0];if(!f)return alert("请选择签名图片");let b=await f.arrayBuffer();let bytes=new Uint8Array(b),binary="";for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));await api("/api/signature",{action:"set",file_name:f.name,image_b64:btoa(binary)});alert("电子签名已上传");loadSignature()}
async function loadEquipment(){if(!token)return;try{let j=await api("/api/equipment/list",{});equipmentItems=j.items;$("equipmentAddButton").classList.toggle("hidden",!j.capabilities?.can_add);$("equipmentImportButton").classList.toggle("hidden",!j.capabilities?.can_add);$("equipmentReviewButton").classList.toggle("hidden",!j.capabilities?.can_review);$("equipmentTable").innerHTML=table([["name","器材"],["brand","品牌"],["category","类别"],["model","型号"],["owner_teacher","归属导师"],["current_user","当前使用人",x=>esc(x.current_user||"无")],["status","状态",x=>status(x.status)],["op","操作",equipmentActions]],j.items,"暂无实验器材")}catch(e){$("equipmentTable").innerHTML=`<div class="notice error">${esc(e.message)}</div>`}}
async function importEquipmentCsv(){let f=$("equipmentCsv").files[0];if(!f)return alert("请选择 CSV 文件");try{let content_b64=await fileToBase64(f);let j=await api("/api/equipment/import-csv",{file_name:f.name,content_b64});alert(`导入完成：成功 ${j.imported} 条，跳过 ${j.skipped} 条`);$("equipmentCsv").value="";loadEquipment()}catch(e){alert(e.message)}}
function equipmentActions(x){let actions=[];if(x.status==="可用")actions.push(`<button class="btn" onclick="borrow(${x.id})">申请借用</button>`);else if(x.can_return)actions.push(`<button class="btn" onclick="requestReturn(${x.id})">申请归还</button>`);else actions.push(`<button class="btn secondary" disabled>正在使用，不可借用</button>`);if(x.can_manage){actions.push(`<button class="btn secondary" onclick="equipmentForm(${x.id})">编辑</button>`);actions.push(`<button class="btn danger" onclick="deleteEquipment(${x.id})">删除</button>`)}return actions.join(" ")}
async function borrow(id){let item=equipmentItems.find(x=>x.id===id);if(item&&item.status!=="可用")return alert("该器材正在使用，不可申请借用。");let reason=prompt("请输入借用原因");if(reason===null)return;try{await api("/api/equipment/request",{equipment_id:id,request_type:"借用",reason});alert("借用申请已提交");loadEquipment();loadMyEquipmentRequests()}catch(e){alert(e.message)}}
async function requestReturn(id){let reason=prompt("请输入归还说明")||"申请归还";try{await api("/api/equipment/request",{equipment_id:id,request_type:"归还",reason});alert("归还申请已提交");loadEquipment();loadCurrentBorrowed()}catch(e){alert(e.message)}}
function equipmentForm(id=0){let x=equipmentItems.find(item=>item.id===id)||{};modal(id?"编辑实验器材":"新增实验器材",`<div class="grid"><div><label>器材名称</label><input id="equipmentName" value="${esc(x.name||"")}"></div><div><label>品牌</label><input id="equipmentBrand" value="${esc(x.brand||"")}"></div><div><label>类别</label><input id="equipmentCategory" value="${esc(x.category||"")}"></div><div><label>型号</label><input id="equipmentModel" value="${esc(x.model||"")}"></div><div><label>器材管理员账号 1（可选）</label><input id="equipmentManager1" value="${esc(x.manager1||"")}"></div><div><label>器材管理员账号 2（可选）</label><input id="equipmentManager2" value="${esc(x.manager2||"")}"></div></div>`,`<button class="btn" onclick="saveEquipment(${id})">保存</button>`)}
async function saveEquipment(id){let payload={id:id||undefined,name:$("equipmentName").value.trim(),brand:$("equipmentBrand").value.trim(),category:$("equipmentCategory").value.trim(),model:$("equipmentModel").value.trim(),manager1:$("equipmentManager1").value.trim(),manager2:$("equipmentManager2").value.trim()};if(!payload.name)return alert("请填写器材名称");try{await api("/api/equipment/upsert",payload);alert(id?"器材已更新":"器材已新增");closeModal();loadEquipment()}catch(e){alert(e.message)}}
async function deleteEquipment(id){let x=equipmentItems.find(item=>item.id===id);if(!confirm(`确定删除实验器材“${x?.name||""}”吗？相关历史申请也会删除。`))return;try{await api("/api/equipment/delete",{id});alert("器材已删除");loadEquipment()}catch(e){alert(e.message)}}
function showEquipmentAux(title,html){$("equipmentAuxTitle").textContent=title;$("equipmentAuxTable").innerHTML=html;$("equipmentAuxCard").classList.remove("hidden")}
async function loadMyEquipmentRequests(){try{let j=await api("/api/equipment/requests",{scope:"mine"});showEquipmentAux("我的器材申请",table([["name","器材"],["request_type","申请类型"],["reason","申请说明"],["status","状态",x=>status(x.status)],["approver_name","审批人"],["review_note","审批意见"],["created_at","提交时间"]],j.items,"暂无器材申请"))}catch(e){alert(e.message)}}
async function loadCurrentBorrowed(){try{let j=await api("/api/equipment/current-borrowed",{});showEquipmentAux("我正在借用",table([["name","器材"],["owner_teacher","归属导师"],["borrowed_at","借用时间"],["equipment_status","状态",x=>status(x.equipment_status)],["op","操作",x=>x.pending_return_id?status("归还待审批"):`<button class="btn" onclick="requestReturn(${x.equipment_id})">申请归还</button>`]],j.items,"当前没有正在借用的器材"))}catch(e){alert(e.message)}}
async function loadEquipmentReviews(){try{let j=await api("/api/equipment/requests",{scope:"approvable"});showEquipmentAux("待审批器材申请",table([["name","器材"],["requester_name","申请人"],["requester_teacher","申请人导师"],["request_type","类型"],["reason","申请说明"],["created_at","提交时间"],["op","操作",x=>`<button class="btn" onclick="reviewEquipment(${x.id},'已批准')">批准</button> <button class="btn danger" onclick="reviewEquipment(${x.id},'已拒绝')">驳回</button>`]],j.items,"暂无可审批的器材申请"))}catch(e){alert(e.message)}}
async function reviewEquipment(id,statusValue){let note=prompt("请输入审批意见")||"";try{await api("/api/equipment/review",{id,status:statusValue,review_note:note});alert("器材审批已提交");loadEquipmentReviews();loadEquipment()}catch(e){alert(e.message)}}
async function userIds(text){let names=String(text||"").split(/[,，\s]+/).filter(Boolean),ids=[];for(let n of names){let j=await api("/api/users/lookup",{username:n});ids.push(j.user.id)}return ids}
async function loadLabs(){if(!token)return;try{let labs=await api("/api/laboratory/list",{});$("labTable").innerHTML=table([["name","实验室"],["college","学院"],["address","地点"],["manager_names","管理员",x=>esc((x.manager_names||[]).join("、"))],["booking_open","预约通道",x=>status(x.booking_open?"开启":"关闭")],["op","操作",x=>x.booking_open?`<button class="btn secondary" onclick='chooseLab(${JSON.stringify(x)})'>选择</button>`:`<button class="btn secondary" disabled>通道关闭</button>`]],labs.items,"暂无实验室");let rec=await api("/api/laboratory/requests",{scope:"records"});$("labRecords").innerHTML=table([["reservation_no","预约单号"],["lab_name","实验室"],["start_time","开始时间"],["end_time","结束时间"],["status","审批状态",x=>status(x.status)],["experiment_status","实验状态",x=>status(x.experiment_status)],["op","操作",x=>labActions(x)]],rec.items,"暂无预约记录")}catch(e){$("labTable").innerHTML=`<div class="notice error">${esc(e.message)}</div>`}}
function labActions(x){let ops=[];if(x.can_confirm){ops.push(`<button class="btn" onclick="confirmLab(${x.id},'同意')">同意同行预约</button> <button class="btn danger" onclick="confirmLab(${x.id},'拒绝')">拒绝同行预约</button>`)}if(x.status==="已批准"&&!["已结束","已强制停止"].includes(x.experiment_status)){ops.push(`<button class="btn secondary" onclick="completionForm(${x.id})">提交结束报告</button>`)}if(x.pdf_path){ops.push(`<button class="btn secondary" onclick="downloadPdf('laboratory',${x.id})">下载PDF</button>`)}return ops.join(" ")||"—"}
function chooseLab(x){selectedLab=x;$("chosenLab").textContent=`已选择：${x.name}｜${x.address}｜管理员：${(x.manager_names||[]).join("、")}`}
async function fileToBase64(file){let buffer=await file.arrayBuffer(),bytes=new Uint8Array(buffer),binary="";for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(binary)}
function previewWorkbookSection(title,rows){if(!rows.length)return `<div class="card"><h4>${title}</h4><div class="empty">未填写（允许为空）</div></div>`;let headers=Object.keys(rows[0]).map(key=>[key,key]);return `<div class="card"><h4>${title}</h4>${table(headers,rows)}</div>`}
async function parseLabWorkbook(){let file=$("labWorkbook").files[0];if(!file){labWorkbookData=null;$("labWorkbookState").textContent="请先选择 .xlsx 工作簿。";$("labWorkbookPreview").innerHTML="";return}if(!file.name.toLowerCase().endsWith(".xlsx"))return alert("仅支持 .xlsx 格式");try{$("labWorkbookState").textContent="正在读取并校验工作簿……";let content_b64=await fileToBase64(file);let j=await api("/api/laboratory/workbook/parse",{file_name:file.name,content_b64});labWorkbookData={name:file.name,b64:content_b64,workflow:j.workflow,chemicals:j.chemicals,safety_plan:j.safety_plan};$("labWorkbookState").innerHTML=`<span class="status ok">已读取</span> ${esc(file.name)}｜实验流程 ${j.workflow.length} 条｜化学品 ${j.chemicals.length} 条｜安全预案 ${j.safety_plan.length} 条`;$("labWorkbookPreview").innerHTML=previewWorkbookSection("实验流程",j.workflow)+previewWorkbookSection("化学品使用",j.chemicals)+previewWorkbookSection("实验安全预案",j.safety_plan)}catch(e){labWorkbookData=null;$("labWorkbookState").innerHTML=`<span class="status bad">读取失败</span> ${esc(e.message)}`;$("labWorkbookPreview").innerHTML=""}}
async function submitLabReservation(){if(!selectedLab)return alert("请先选择实验室");if(!labWorkbookData)return alert("请先上传并成功读取实验室预约资料工作簿");if(!$("labCommitmentSigned").checked)return alert("请先勾选并签署实验室安全承诺书");try{let ids=await userIds($("labParticipants").value);await api("/api/laboratory/reserve",{laboratory_id:selectedLab.id,participant_user_ids:ids,phone:$("labPhone").value,start_time:$("labStart").value,end_time:$("labEnd").value,purpose:$("labPurpose").value,workbook_name:labWorkbookData.name,workbook_b64:labWorkbookData.b64,workflow:labWorkbookData.workflow,chemicals:labWorkbookData.chemicals,safety_plan:labWorkbookData.safety_plan,commitment_signed:$("labCommitmentSigned").checked});alert("预约已提交");labWorkbookData=null;$("labWorkbook").value="";$("labCommitmentSigned").checked=false;$("labWorkbookState").textContent="尚未上传工作簿。";$("labWorkbookPreview").innerHTML="";loadLabs()}catch(e){alert(e.message)}}
function completionForm(id){modal("提交实验结束报告",`<label>实验完成情况</label><textarea id="completionText"></textarea><label>安全隐患或异常情况</label><textarea id="hazardText"></textarea>`,`<button class="btn" onclick="submitCompletion(${id})">提交报告</button>`)}
async function submitCompletion(id){try{await api("/api/laboratory/complete-report",{id,completion_report:$("completionText").value,hazard_report:$("hazardText").value});alert("结束报告已提交");closeModal();loadLabs()}catch(e){alert(e.message)}}
async function confirmLab(id,decision){let note=prompt("请输入同行确认说明（可不填）")||"";try{await api("/api/laboratory/participant-confirm",{id,decision,note});alert("同行预约确认已处理");loadLabs()}catch(e){alert(e.message)}}
async function loadLabReviews(){if(!token)return;for(let scope of ["mentor","manager"]){let target=scope==="mentor"?"mentorReviewTable":"managerReviewTable";try{let j=await api("/api/laboratory/requests",{scope});$(target).innerHTML=table([["reservation_no","预约单号"],["display_name","申请人"],["requester_teacher","申请人导师"],["lab_name","实验室"],["start_time","开始时间"],["status","状态",x=>status(x.status)],["op","操作",x=>`<button class="btn" onclick="reviewLab(${x.id},'${scope}','批准')">批准</button> <button class="btn danger" onclick="reviewLab(${x.id},'${scope}','驳回')">驳回</button>`]],j.items,"暂无待审核事项")}catch(e){$(target).innerHTML=`<div class="notice error">${esc(e.message)}</div>`}}}
async function reviewLab(id,scope,decision){let note=prompt("请输入审批意见")||"";try{await api("/api/laboratory/review",{id,scope,decision,note});alert("审批已提交");loadLabReviews()}catch(e){alert(e.message)}}
function chemicalActions(x){let ops=[];if(x.can_confirm){ops.push(`<button class="btn" onclick="confirmChemical(${x.id},'同意')">同意同行领用</button> <button class="btn danger" onclick="confirmChemical(${x.id},'拒绝')">拒绝</button>`)}if(x.can_review){let stage=x.review_stage?`（${esc(x.review_stage)}）`:"";ops.push(`<button class="btn" onclick="reviewChemical(${x.id},'批准')">批准${stage}</button> <button class="btn danger" onclick="reviewChemical(${x.id},'驳回')">驳回${stage}</button>`)}if(x.can_dispose){ops.push(`<button class="btn secondary" onclick="disposeForm(${x.id})">报告处理完毕</button>`)}if(x.can_review_disposal){ops.push(`<button class="btn" onclick="reviewDisposal(${x.id},'批准')">确认处置完成</button> <button class="btn danger" onclick="reviewDisposal(${x.id},'驳回')">驳回处置报告</button>`)}if(x.pdf_path){ops.push(`<button class="btn secondary" onclick="downloadPdf('chemical',${x.id})">下载PDF</button>`)}return ops.join(" ")||"—"}
async function loadChemicals(){if(!token)return;try{let list=await api("/api/chemical/list",{});$("chemicalTable").innerHTML=table([["name","化学品"],["warehouse_name","库房"],["owner_teacher","归属导师"],["quantity","库存",x=>`${esc(x.quantity)} ${esc(x.unit)}`],["available_per_student","单人可领",x=>`${esc(x.available_per_student)} ${esc(x.unit)}`],["service_open","通道",x=>status(x.service_open?"开启":"关闭")],["op","操作",x=>x.service_open?`<button class="btn secondary" onclick='chooseChemical(${JSON.stringify(x)})'>选择</button>`:"—"]],list.items,"暂无已登记化学品");let rec=await api("/api/chemical/withdrawals",{});$("chemicalRecords").innerHTML=table([["withdrawal_no","领用单号"],["chemical_name","化学品"],["display_name","领用人"],["co_collector_name","共同领用人"],["owner_teacher","归属导师"],["quantity","数量",x=>`${esc(x.quantity)} ${esc(x.unit)}`],["status","状态",x=>status(x.status)],["disposal_status","处置",x=>status(x.disposal_status)],["op","操作",x=>chemicalActions(x)]],rec.items,"暂无领用记录")}catch(e){$("chemicalTable").innerHTML=`<div class="notice error">${esc(e.message)}</div>`;$("chemicalRecords").innerHTML=`<div class="notice error">${esc(e.message)}</div>`}}
function chooseChemical(x){selectedChemical=x;$("chosenChemical").textContent=`已选择：${x.name}｜库存 ${x.quantity} ${x.unit}｜归属导师：${x.owner_teacher}`}
async function submitChemical(){if(!selectedChemical)return alert("请先选择化学品");try{let j=await api("/api/chemical/withdraw",{chemical_id:selectedChemical.id,quantity:Number($("chemQty").value),purpose:$("chemPurpose").value,co_collector_username:$("chemCo").value,storage_location:$("chemStorage").value});alert("领用申请已提交，领用单号："+(j.withdrawal_no||j.id));loadChemicals()}catch(e){alert(e.message)}}
async function confirmChemical(id,decision){try{await api("/api/chemical/participant-confirm",{id,decision,note:"网页端确认"});alert("已处理共同领用确认");loadChemicals()}catch(e){alert(e.message)}}
async function reviewChemical(id,decision){let note=prompt("请输入审批意见")||"";try{await api("/api/chemical/withdraw-review",{id,decision,note});alert("审批已提交");loadChemicals()}catch(e){alert(e.message)}}
function disposeForm(id){modal("危险化学品处置报告",`<label>请说明化学品及废弃物是否按要求处理完毕</label><textarea id="disposeText"></textarea>`,`<button class="btn" onclick="submitDispose(${id})">提交报告</button>`)}
async function submitDispose(id){try{await api("/api/chemical/disposal-report",{id,report:$("disposeText").value});alert("处置报告已提交");closeModal();loadChemicals()}catch(e){alert(e.message)}}
async function reviewDisposal(id,decision){let note=prompt("请输入处置确认意见")||"";try{await api("/api/chemical/disposal-review",{id,decision,note});alert("处置确认已处理");loadChemicals()}catch(e){alert(e.message)}}
async function loadAccounts(){if(!token)return;try{let j=await api("/api/users/list",{});let rows=j.items.filter(x=>!x.active);$("accountTable").innerHTML=table([["display_name","真实姓名"],["advisor_name","导师姓名"],["team_name","团队"],["role","角色"],["created_at","注册时间"],["op","操作",x=>`<button class="btn" onclick="approve(${x.id},1)">批准</button> <button class="btn danger" onclick="approve(${x.id},0)">停用</button>`]],rows,"暂无待审批账号")}catch(e){$("accountTable").innerHTML=`<div class="notice error">${esc(e.message)}</div>`}}
async function approve(id,active){try{await api("/api/users/approve",{id,active});alert(active?"账号已批准":"账号已停用");loadAccounts()}catch(e){alert(e.message)}}
</script></body></html>"""

    def do_GET(self):
        clean_path = self.path.split("?", 1)[0].rstrip("/")
        if clean_path == "/api/health":
            self.send_json(200, {"ok": True, "name": SERVER_NAME, "version": SERVER_VERSION, "time": now(), "public_url": self.server.config.get("public_url", "")})
        elif clean_path == "/api/laboratory/template.xlsx":
            self.send_bytes(
                200,
                create_lab_workbook_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "实验室预约资料模板.xlsx",
            )
        elif clean_path == "/api/equipment/template.csv":
            rows = [
                ["器材名称", "品牌", "类别", "型号", "器材管理员账号1", "器材管理员账号2"],
                ["示例：旋涂仪", "Laurell", "工艺设备", "WS-650", "", ""],
            ]
            stream = io.StringIO()
            writer = csv.writer(stream)
            writer.writerows(rows)
            raw = ("\ufeff" + stream.getvalue()).encode("utf-8")
            self.send_bytes(200, raw, "text/csv; charset=utf-8", "实验器材导入模板.csv")
        elif clean_path in ("","/lab", "/laboratory","/portal"):
            raw = self.safety_portal_page().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        else:
            self.error_json(404, "unknown endpoint")

    def do_POST(self):
        conn = open_db_connection()
        equipment_lock = None
        try:
            data = self.read_json()
            if self.path in ("/api/equipment/request", "/api/equipment/review"):
                equipment_lock = self.server.equipment_write_lock
                equipment_lock.acquire()
            if self.path == "/api/mentors/search":
                keyword = str(data.get("keyword") or "").strip()
                if not keyword:
                    return self.error_json(400, "请输入导师真实姓名")
                pattern = f"%{keyword}%"
                rows = [
                    dict(x)
                    for x in conn.execute(
                        """SELECT id,username,display_name,team_name,group_code
                           FROM users
                           WHERE active=1 AND role IN ('导师','老师','教授','PI','teacher','tutor','mentor')
                             AND (display_name LIKE ? OR username LIKE ? OR team_name LIKE ?)
                           ORDER BY CASE WHEN display_name=? THEN 0 ELSE 1 END,display_name,team_name,username
                           LIMIT 100""",
                        (pattern, pattern, pattern, keyword),
                    )
                ]
                self.send_json(200, {"ok": True, "items": rows})
                return
            if self.path == "/api/register":
                username = (data.get("username") or "").strip()
                password = data.get("password") or ""
                role = self.normalize_role(data.get("role") or "学生")
                display_name = (data.get("display_name") or username).strip()
                team_name = (data.get("team_name") or "").strip()
                advisor_name = (data.get("advisor_name") or data.get("group_code") or "").strip()
                mentor_user_id = int(data.get("mentor_user_id") or 0)
                mentor_username = str(data.get("mentor_username") or "").strip()
                if role == "超级管理员":
                    return self.error_json(403, "超级管理员只能在服务器管理端创建，不能从客户端注册。")
                if not username or not password:
                    return self.error_json(400, "username and password required")
                if not display_name:
                    return self.error_json(400, "注册必须填写姓名")
                username_owner = conn.execute(
                    "SELECT id,active FROM users WHERE username=? ORDER BY id LIMIT 1",
                    (username,),
                ).fetchone()
                if role == "导师":
                    if not team_name:
                        return self.error_json(400, "导师注册必须填写团队代码")
                    group = username
                    mentor_user_id = 0
                    mentor_username = ""
                else:
                    teacher = None
                    if mentor_user_id:
                        teacher = conn.execute(
                            "SELECT * FROM users WHERE id=? AND active=1",
                            (mentor_user_id,),
                        ).fetchone()
                    elif mentor_username:
                        teacher = conn.execute(
                            "SELECT * FROM users WHERE username=? AND active=1 ORDER BY id LIMIT 1",
                            (mentor_username,),
                        ).fetchone()
                    elif advisor_name:
                        matches = conn.execute(
                            """SELECT * FROM users WHERE active=1 AND display_name=?
                               AND role IN ('导师','老师','教授','PI','teacher','tutor','mentor')
                               ORDER BY id""",
                            (advisor_name,),
                        ).fetchall()
                        if len(matches) == 1:
                            teacher = matches[0]
                        elif len(matches) > 1:
                            return self.error_json(409, "存在多位同名导师，请点击查询导师并按账号选择。")
                    if not teacher:
                        return self.error_json(404, "未找到所选导师，请重新查询并选择。")
                    if self.normalize_role(teacher["role"]) != "导师":
                        return self.error_json(400, "所选账号不是导师账号")
                    mentor_user_id = teacher["id"]
                    mentor_username = teacher["username"]
                    advisor_name = self.user_label(teacher)
                    group = teacher["group_code"] or teacher["username"]
                    team_name = teacher["team_name"] or team_name
                existing = conn.execute("SELECT id,active FROM users WHERE group_code=? AND username=?", (group, username)).fetchone()
                if username_owner and (not existing or username_owner["id"] != existing["id"]):
                    return self.error_json(409, "该用户名已被使用，请更换唯一用户名。")
                if existing and existing["active"]:
                    return self.error_json(409, "账号已存在并已批准，请直接登录。")
                salt, digest = hash_password(password)
                if existing:
                    conn.execute(
                        """UPDATE users SET role=?,display_name=?,team_name=?,mentor_user_id=?,
                           mentor_username=?,salt=?,password_hash=?,active=0 WHERE id=?""",
                        (role, display_name, team_name, mentor_user_id or None, mentor_username, salt, digest, existing["id"]),
                    )
                else:
                    conn.execute(
                        """INSERT INTO users(group_code,username,role,display_name,team_name,mentor_user_id,
                           mentor_username,salt,password_hash,active,created_at)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (group, username, role, display_name, team_name, mentor_user_id or None,
                         mentor_username, salt, digest, 0, now()),
                    )
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True, "message": "注册申请已提交，请等待导师或超级管理员批准。"})
                return

            if self.path == "/api/login":
                group = self.group_from(data)
                username = (data.get("username") or "").strip()
                row = conn.execute("SELECT * FROM users WHERE username=? AND active=1 ORDER BY CASE WHEN group_code=? THEN 0 ELSE 1 END,id LIMIT 1", (username, group)).fetchone()
                if not row:
                    return self.error_json(403, "账号不存在、未批准或已停用")
                _salt, digest = hash_password(data.get("password") or "", row["salt"])
                if digest != row["password_hash"]:
                    return self.error_json(403, "用户名或密码错误")
                token = secrets.token_urlsafe(32)
                conn.execute("INSERT INTO tokens(token,user_id,created_at,last_seen) VALUES(?,?,?,?)", (token, row["id"], now(), now()))
                self.maintain_login_sessions(conn, row["id"])
                conn.commit()
                self.send_json(200, {"ok": True, "token": token, "role": self.normalize_role(row["role"]), "username": username, "display_name": row["display_name"], "advisor_name": row["group_code"], "group_code": row["group_code"], "team_name": row["team_name"]})
                return

            user = self.auth_user(conn)
            if not user:
                return self.error_json(401, "login required")
            self.cleanup_expired_approval_pdfs(conn)
            group = self.group_from(data, user)
            if group != user["group_code"] and not self.is_admin(user):
                return self.error_json(403, "group mismatch")

            if self.path == "/api/signature":
                if data.get("action") == "set":
                    image_b64 = str(data.get("image_b64") or "")
                    if not image_b64:
                        return self.error_json(400, "signature image required")
                    try:
                        raw = base64.b64decode(image_b64)
                    except Exception:
                        return self.error_json(400, "invalid signature image")
                    if len(raw) > 2 * 1024 * 1024:
                        return self.error_json(413, "signature image too large")
                    conn.execute(
                        "INSERT INTO user_signatures(user_id,image_b64,file_name,updated_at) VALUES(?,?,?,?) ON CONFLICT(user_id) DO UPDATE SET image_b64=excluded.image_b64,file_name=excluded.file_name,updated_at=excluded.updated_at",
                        (user["id"], image_b64, os.path.basename(data.get("file_name") or "signature.png"), now()),
                    )
                    conn.commit()
                row = conn.execute("SELECT file_name,updated_at,image_b64 FROM user_signatures WHERE user_id=?", (user["id"],)).fetchone()
                self.send_json(200, {"ok": True, "signature": dict(row) if row else None})
                return

            if self.path == "/api/laboratory/workbook/parse":
                file_name = os.path.basename(str(data.get("file_name") or "")).strip()
                if not file_name.lower().endswith(".xlsx"):
                    return self.error_json(400, "仅支持 .xlsx 格式的实验室预约资料")
                try:
                    raw = base64.b64decode(str(data.get("content_b64") or ""), validate=True)
                    parsed = parse_lab_workbook_bytes(raw)
                except (ValueError, base64.binascii.Error) as exc:
                    return self.error_json(400, str(exc) or "工作簿内容无效")
                self.send_json(200, {"ok": True, "file_name": file_name, **parsed})
                return

            if self.path == "/api/team/members":
                team = data.get("team_name") or user["team_name"]
                if not team:
                    return self.error_json(400, "team required")
                rows = [
                    dict(x)
                    for x in conn.execute(
                        "SELECT id,username,display_name,role,group_code,team_name FROM users WHERE team_name=? AND active=1 AND role IN ('学生','导师') ORDER BY role DESC,display_name,username",
                        (team,),
                    )
                ]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/users/lookup":
                key = str(data.get("username") or data.get("keyword") or "").strip()
                if not key:
                    return self.error_json(400, "请输入要查询的用户名")
                row = conn.execute(
                    """SELECT id,username,display_name,role,group_code,team_name
                       FROM users WHERE active=1 AND username=? ORDER BY id LIMIT 1""",
                    (key,),
                ).fetchone()
                if not row:
                    return self.error_json(404, "未找到已批准的用户")
                self.send_json(200, {"ok": True, "user": dict(row)})
                return

            if self.path == "/api/laboratory/upsert":
                if not self.is_admin(user):
                    return self.error_json(403, "only super admin can create laboratories")
                lab_type = data.get("lab_type") or "公共实验室"
                if lab_type not in ("公共实验室", "团队实验室"):
                    return self.error_json(400, "invalid laboratory type")
                managers = list(dict.fromkeys(str(x) for x in data.get("managers", []) if x))[:20]
                if not managers:
                    return self.error_json(400, "请至少指定一位实验室管理员")
                for username in managers:
                    tutor = conn.execute("SELECT role FROM users WHERE username=? AND active=1", (username,)).fetchone()
                    if not tutor or self.normalize_role(tutor["role"]) != "导师":
                        return self.error_json(400, f"实验室管理员必须是导师：{username}")
                values = (
                    str(data.get("name") or "").strip(),
                    str(data.get("college") or "").strip(),
                    str(data.get("address") or "").strip(),
                    lab_type,
                    str(data.get("team_name") or "").strip(),
                    str(data.get("group_code") or "").strip(),
                    json.dumps(managers, ensure_ascii=False),
                    str(data.get("commitment_text") or ""),
                    1 if data.get("active", True) else 0,
                    self.user_label(user),
                    now(),
                )
                if not all(values[:3]):
                    return self.error_json(400, "name, college and address required")
                if lab_type == "团队实验室" and not values[4]:
                    return self.error_json(400, "team laboratory requires team_name")
                item_id = int(data.get("id") or 0)
                if item_id:
                    conn.execute(
                        "UPDATE laboratories SET name=?,college=?,address=?,lab_type=?,team_name=?,group_code=?,managers_json=?,commitment_text=?,active=?,updated_at=? WHERE id=?",
                        values[:9] + (values[10], item_id),
                    )
                else:
                    item_id = conn.execute(
                        "INSERT INTO laboratories(name,college,address,lab_type,team_name,group_code,managers_json,commitment_text,active,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                        values + (values[10],),
                    ).lastrowid
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True, "id": item_id})
                return

            if self.path == "/api/laboratory/delete":
                if not self.is_admin(user):
                    return self.error_json(403, "only super admin can delete laboratories")
                lab_id = int(data.get("id") or 0)
                row = conn.execute("SELECT * FROM laboratories WHERE id=?", (lab_id,)).fetchone()
                if not row:
                    return self.error_json(404, "laboratory not found")
                active = conn.execute(
                    """SELECT reservation_no FROM laboratory_reservations
                       WHERE laboratory_id=? AND status NOT LIKE '%驳回%'
                       AND experiment_status IN ('未开始','正在进行','待结束确认')
                       ORDER BY id DESC LIMIT 1""",
                    (lab_id,),
                ).fetchone()
                if active:
                    return self.error_json(409, f"该实验室仍有未结束预约：{active['reservation_no']}")
                conn.execute("UPDATE laboratories SET active=0,updated_at=? WHERE id=?", (now(), lab_id))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/laboratory/managers":
                if not self.is_admin(user):
                    return self.error_json(403, "only super admin can update laboratory managers")
                lab_id = int(data.get("laboratory_id") or data.get("id") or 0)
                lab = conn.execute("SELECT * FROM laboratories WHERE id=?", (lab_id,)).fetchone()
                if not lab:
                    return self.error_json(404, "laboratory not found")
                managers = list(dict.fromkeys(str(x).strip() for x in data.get("managers", []) if str(x).strip()))[:20]
                if not managers:
                    return self.error_json(400, "请至少保留一位实验室管理员")
                for username in managers:
                    tutor = conn.execute("SELECT role FROM users WHERE username=? AND active=1", (username,)).fetchone()
                    if not tutor or self.normalize_role(tutor["role"]) != "导师":
                        return self.error_json(400, f"实验室管理员必须是导师：{username}")
                conn.execute("UPDATE laboratories SET managers_json=?,updated_at=? WHERE id=?", (json.dumps(managers, ensure_ascii=False), now(), lab_id))
                conn.commit()
                self.send_json(200, {"ok": True, "managers": managers})
                return

            if self.path == "/api/laboratory/list":
                auth_code = str(data.get("auth_code") or "").strip()
                auth_lab_id = 0
                if auth_code:
                    auth = conn.execute("SELECT laboratory_id FROM laboratory_auth_codes WHERE code=? AND expires_at>?", (auth_code, now())).fetchone()
                    auth_lab_id = auth["laboratory_id"] if auth else 0
                if self.is_admin(user):
                    rows = [dict(x) for x in conn.execute("SELECT * FROM laboratories ORDER BY active DESC,college,name")]
                else:
                    rows = [dict(x) for x in conn.execute("SELECT * FROM laboratories WHERE active=1 ORDER BY lab_type,college,name")]
                self.cleanup_expired_blacklists(conn)
                blocked_ids = {
                    x["laboratory_id"]
                    for x in conn.execute(
                        "SELECT laboratory_id FROM laboratory_blacklist WHERE user_id=? AND active=1",
                        (user["id"],),
                    )
                }
                manager_usernames = set()
                team_names = set()
                for item in rows:
                    manager_usernames.update(self.json_list(item.get("managers_json", "[]")))
                    if item["lab_type"] == "团队实验室" and item["team_name"]:
                        team_names.add(item["team_name"])
                manager_map = {}
                if manager_usernames:
                    placeholders = ",".join("?" for _ in manager_usernames)
                    for manager in conn.execute(
                        f"""SELECT username,display_name,role FROM users
                            WHERE username IN ({placeholders}) AND active=1""",
                        tuple(manager_usernames),
                    ):
                        if self.normalize_role(manager["role"]) == "导师":
                            manager_map[manager["username"]] = self.user_label(manager)
                team_manager_map = {}
                if team_names:
                    placeholders = ",".join("?" for _ in team_names)
                    for manager in conn.execute(
                        f"""SELECT username,display_name,team_name,role FROM users
                            WHERE team_name IN ({placeholders}) AND active=1
                            ORDER BY display_name,username""",
                        tuple(team_names),
                    ):
                        if self.normalize_role(manager["role"]) != "导师":
                            continue
                        team_manager_map.setdefault(manager["team_name"], []).append({
                            "username": manager["username"],
                            "display_name": self.user_label(manager),
                        })
                for item in rows:
                    item["managers"] = self.json_list(item.pop("managers_json", "[]"))
                    item["manager_details"] = [
                        {
                            "username": username,
                            "display_name": manager_map.get(username, username),
                        }
                        for username in item["managers"]
                    ]
                    if item["lab_type"] == "团队实验室" and item["team_name"]:
                        known = {x["username"] for x in item["manager_details"]}
                        for manager in team_manager_map.get(item["team_name"], []):
                            if manager["username"] not in known:
                                item["manager_details"].append(manager)
                    item["manager_names"] = [x["display_name"] for x in item["manager_details"]]
                    item["blacklisted"] = item["id"] in blocked_ids
                conn.commit()
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/laboratory/authcode":
                lab = conn.execute("SELECT * FROM laboratories WHERE id=?", (data.get("laboratory_id"),)).fetchone()
                if not lab:
                    return self.error_json(404, "laboratory not found")
                if not self.is_lab_manager(user, lab):
                    return self.error_json(403, "您没有该实验室的管理权限")
                code = secrets.token_urlsafe(12)
                expires = (datetime.now() + timedelta(days=max(1, min(int(data.get("days") or 7), 90)))).isoformat(timespec="seconds")
                conn.execute("INSERT INTO laboratory_auth_codes(code,laboratory_id,creator,expires_at,created_at) VALUES(?,?,?,?,?)", (code, lab["id"], self.user_label(user), expires, now()))
                conn.commit()
                self.send_json(200, {"ok": True, "code": code, "expires_at": expires})
                return

            if self.path == "/api/laboratory/commitment":
                lab = conn.execute("SELECT * FROM laboratories WHERE id=?", (data.get("laboratory_id"),)).fetchone()
                if not lab or not self.is_lab_manager(user, lab):
                    return self.error_json(403, "only laboratory manager can update commitment")
                conn.execute("UPDATE laboratories SET commitment_text=?,updated_at=? WHERE id=?", (data.get("commitment_text") or "", now(), lab["id"]))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/laboratory/reserve":
                lab = conn.execute("SELECT * FROM laboratories WHERE id=? AND active=1", (data.get("laboratory_id"),)).fetchone()
                if not lab:
                    return self.error_json(404, "laboratory not found")
                if not lab["booking_open"]:
                    return self.error_json(409, "该实验室预约通道当前已关闭")
                occupied = conn.execute(
                    """SELECT reservation_no FROM laboratory_reservations
                       WHERE laboratory_id=? AND requester_id=? AND status='已批准'
                       AND experiment_status IN ('未开始','正在进行','待结束确认')
                       ORDER BY id DESC LIMIT 1""",
                    (lab["id"], user["id"]),
                ).fetchone()
                if occupied:
                    return self.error_json(409, f"您在该实验室上一预约尚未结束：{occupied['reservation_no']}")
                if self.current_blacklist(conn, lab["id"], user["id"]):
                    return self.error_json(403, "当前账号处于该实验室预约黑名单中")
                participant_ids = []
                for raw_id in data.get("participant_user_ids") or []:
                    try:
                        participant_id = int(raw_id)
                    except (TypeError, ValueError):
                        continue
                    if participant_id != user["id"] and participant_id not in participant_ids:
                        participant_ids.append(participant_id)
                participants = []
                if participant_ids:
                    marks = ",".join("?" for _ in participant_ids)
                    participants = list(conn.execute(
                        f"SELECT * FROM users WHERE active=1 AND id IN ({marks}) ORDER BY id",
                        participant_ids,
                    ))
                if len(participants) != len(participant_ids):
                    return self.error_json(400, "同行人名单中存在无效账号")
                for participant in participants:
                    if self.current_blacklist(conn, lab["id"], participant["id"]):
                        return self.error_json(403, f"同行人 {self.user_label(participant)} 当前无该实验室预约或同行权限")
                if lab["lab_type"] == "公共实验室" and not participants:
                    return self.error_json(400, "公共实验室不得单人进入，请至少选择一位同行学生或导师")
                signature = self.user_signature(conn, user["id"])
                if not signature:
                    return self.error_json(400, "请先上传电子签名")
                if not data.get("commitment_signed"):
                    return self.error_json(400, "请先勾选并签署实验室安全承诺书")
                workbook_b64 = str(data.get("workbook_b64") or "")
                workbook_name = os.path.basename(str(data.get("workbook_name") or "")).strip()
                if workbook_b64:
                    if not workbook_name.lower().endswith(".xlsx"):
                        return self.error_json(400, "实验室预约资料必须为 .xlsx 工作簿")
                    try:
                        workbook_raw = base64.b64decode(workbook_b64, validate=True)
                        parsed_workbook = parse_lab_workbook_bytes(workbook_raw)
                    except (ValueError, base64.binascii.Error) as exc:
                        return self.error_json(400, str(exc) or "工作簿内容无效")
                    workflow = parsed_workbook["workflow"]
                    chemicals = parsed_workbook["chemicals"]
                    safety_plan = parsed_workbook["safety_plan"]
                else:
                    workflow = data.get("workflow") or []
                    chemicals = data.get("chemicals") or []
                    safety_plan = data.get("safety_plan") or []
                if not workflow or not safety_plan:
                    return self.error_json(400, "实验流程和实验安全预案不能为空")
                teacher = self.team_mentor(conn, user)
                requester_teacher = self.user_label(teacher)
                phone = str(data.get("phone") or user["phone"] or "").strip()
                if not phone:
                    return self.error_json(400, "请填写预约人电话号码")
                conn.execute("UPDATE users SET phone=? WHERE id=?", (phone, user["id"]))
                created = now()
                item_id = conn.execute(
                    """INSERT INTO laboratory_reservations(
                       laboratory_id,requester_id,requester_group,requester_teacher,companion_user_id,companion_name,companion_role,
                       start_time,end_time,purpose,workbook_name,workbook_b64,workflow_json,chemicals_json,safety_plan_json,
                       commitment_snapshot,student_signature_b64,reservation_no,requester_phone,status,created_at,updated_at
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        lab["id"], user["id"], user["group_code"], requester_teacher,
                        participants[0]["id"] if participants else None,
                        "、".join(self.user_label(x) for x in participants),
                        "、".join(self.normalize_role(x["role"]) for x in participants),
                        data.get("start_time") or "", data.get("end_time") or "", data.get("purpose") or "",
                        workbook_name, workbook_b64,
                        json.dumps(workflow, ensure_ascii=False), json.dumps(chemicals, ensure_ascii=False), json.dumps(safety_plan, ensure_ascii=False),
                        lab["commitment_text"], signature, "", phone,
                        "待同行人确认" if participants else "待导师审核", created, created,
                    ),
                ).lastrowid
                reservation_no = f"LAB-{datetime.now():%Y%m%d}-{item_id:06d}"
                conn.execute("UPDATE laboratory_reservations SET reservation_no=? WHERE id=?", (reservation_no, item_id))
                involved_users = [user] + participants
                for participant in participants:
                    mentor = self.team_mentor(conn, participant)
                    conn.execute(
                        """INSERT INTO laboratory_reservation_participants(
                           reservation_id,user_id,participant_name,participant_role,teacher_name,
                           confirmation_status,created_at,updated_at
                        ) VALUES(?,?,?,?,?,'待确认',?,?)""",
                        (item_id, participant["id"], self.user_label(participant),
                         self.normalize_role(participant["role"]), self.user_label(mentor), created, created),
                    )
                mentors = {}
                for involved in involved_users:
                    mentor = self.team_mentor(conn, involved)
                    if mentor:
                        mentors[mentor["id"]] = mentor
                for mentor in mentors.values():
                    conn.execute(
                        """INSERT INTO laboratory_reservation_mentor_reviews(
                           reservation_id,mentor_user_id,mentor_name,status,created_at,updated_at
                        ) VALUES(?,?,?,'待审核',?,?)""",
                        (item_id, mentor["id"], self.user_label(mentor), created, created),
                    )
                self.refresh_lab_reservation_state(conn, item_id)
                self.lab_log(conn, item_id, user, "提交预约", f"预约单号：{reservation_no}")
                conn.commit()
                self.send_json(200, {"ok": True, "id": item_id, "reservation_no": reservation_no})
                return

            if self.path == "/api/laboratory/requests":
                scope = str(data.get("scope") or "records")
                base = """SELECT r.*,l.name AS lab_name,l.college,l.address,l.lab_type,l.team_name,
                          u.username,u.display_name FROM laboratory_reservations r
                          JOIN laboratories l ON l.id=r.laboratory_id JOIN users u ON u.id=r.requester_id"""
                if self.is_admin(user):
                    rows = [dict(x) for x in conn.execute(base + " ORDER BY r.id DESC LIMIT 1000")]
                elif self.normalize_role(user["role"]) == "导师":
                    if scope == "mentor":
                        rows = [dict(x) for x in conn.execute(
                            base + """ JOIN laboratory_reservation_mentor_reviews mr ON mr.reservation_id=r.id
                                      WHERE mr.mentor_user_id=? AND mr.status='待审核'
                                      AND r.status='待导师审核' ORDER BY r.id DESC LIMIT 800""",
                            (user["id"],),
                        )]
                    elif scope == "manager":
                        rows = [dict(x) for x in conn.execute(
                            base + """ WHERE (l.team_name=? OR instr(l.managers_json,?)>0)
                                      AND (r.status='待实验室审核' OR r.completion_status='待结束确认')
                                      ORDER BY r.id DESC LIMIT 800""",
                            (user["team_name"], f'"{user["username"]}"'),
                        )]
                    else:
                        rows = [dict(x) for x in conn.execute(
                            base + """ WHERE r.requester_id=? OR
                                      EXISTS(SELECT 1 FROM laboratory_reservation_participants p
                                             WHERE p.reservation_id=r.id AND p.user_id=?)
                                      ORDER BY r.id DESC LIMIT 800""",
                            (user["id"], user["id"]),
                        )]
                else:
                    rows = [dict(x) for x in conn.execute(
                        base + """ WHERE r.requester_id=? OR
                                  EXISTS(SELECT 1 FROM laboratory_reservation_participants p
                                         WHERE p.reservation_id=r.id AND p.user_id=?)
                                  ORDER BY r.id DESC LIMIT 500""",
                        (user["id"], user["id"]),
                    )]
                for item in rows:
                    item["participants"] = [dict(x) for x in conn.execute(
                        "SELECT * FROM laboratory_reservation_participants WHERE reservation_id=? ORDER BY id",
                        (item["id"],),
                    )]
                    item["mentor_reviews"] = [dict(x) for x in conn.execute(
                        "SELECT * FROM laboratory_reservation_mentor_reviews WHERE reservation_id=? ORDER BY id",
                        (item["id"],),
                    )]
                    item["can_confirm"] = any(
                        x["user_id"] == user["id"] and x["confirmation_status"] == "待确认"
                        for x in item["participants"]
                    )
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/laboratory/participant-confirm":
                participant = conn.execute(
                    """SELECT p.*,r.status FROM laboratory_reservation_participants p
                       JOIN laboratory_reservations r ON r.id=p.reservation_id
                       WHERE p.reservation_id=? AND p.user_id=?""",
                    (data.get("id"), user["id"]),
                ).fetchone()
                if not participant:
                    return self.error_json(403, "当前账号不是该预约的同行人")
                decision = data.get("decision")
                if decision not in ("同意", "拒绝"):
                    return self.error_json(400, "invalid decision")
                signature = self.user_signature(conn, user["id"]) if decision == "同意" else ""
                if decision == "同意" and not signature:
                    return self.error_json(400, "同意同行前请先上传电子签名")
                conn.execute(
                    """UPDATE laboratory_reservation_participants
                       SET confirmation_status=?,confirmation_note=?,signature_b64=?,confirmed_at=?,updated_at=?
                       WHERE reservation_id=? AND user_id=?""",
                    ("已同意" if decision == "同意" else "已拒绝", data.get("note") or "",
                     signature, now(), now(), data.get("id"), user["id"]),
                )
                self.lab_log(conn, int(data.get("id")), user, "同行人确认", f"{decision}；{data.get('note') or ''}")
                self.refresh_lab_reservation_state(conn, int(data.get("id")))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/laboratory/review":
                reservation = conn.execute(
                    """SELECT r.*,l.managers_json,l.lab_type,l.team_name FROM laboratory_reservations r
                       JOIN laboratories l ON l.id=r.laboratory_id WHERE r.id=?""",
                    (data.get("id"),),
                ).fetchone()
                if not reservation:
                    return self.error_json(404, "reservation not found")
                decision = data.get("decision")
                if decision not in ("批准", "驳回"):
                    return self.error_json(400, "invalid decision")
                reviewer = self.user_label(user)
                signature = self.user_signature(conn, user["id"])
                if decision == "批准" and not signature:
                    return self.error_json(400, "批准前请先上传电子签名")
                review_scope = str(data.get("scope") or "")
                if review_scope == "mentor":
                    task = conn.execute(
                        """SELECT * FROM laboratory_reservation_mentor_reviews
                           WHERE reservation_id=? AND mentor_user_id=?""",
                        (reservation["id"], user["id"]),
                    ).fetchone()
                    if not task and not self.is_admin(user):
                        return self.error_json(403, "当前导师没有该预约的导师审核任务")
                    if reservation["status"] == "待同行人确认":
                        return self.error_json(409, "仍有同行人尚未确认")
                    if task:
                        conn.execute(
                            """UPDATE laboratory_reservation_mentor_reviews
                               SET status=?,review_note=?,signature_b64=?,reviewed_at=?,updated_at=?
                               WHERE id=?""",
                            ("已批准" if decision == "批准" else "已驳回",
                             data.get("note") or "", signature, now(), now(), task["id"]),
                        )
                    self.refresh_lab_reservation_state(conn, reservation["id"])
                    self.lab_log(conn, reservation["id"], user, "导师审核", f"{decision}；{data.get('note') or ''}")
                elif review_scope == "manager":
                    lab = conn.execute("SELECT * FROM laboratories WHERE id=?", (reservation["laboratory_id"],)).fetchone()
                    if not self.is_lab_manager(user, lab):
                        return self.error_json(403, "only laboratory manager can review")
                    if reservation["completion_status"] == "待结束确认":
                        completion_status = "已结束" if decision == "批准" else "需补充"
                        experiment_status = "已结束" if decision == "批准" else "正在进行"
                        conn.execute(
                            """UPDATE laboratory_reservations
                               SET completion_status=?,experiment_status=?,ended_at=?,
                                   completion_review_note=?,updated_at=? WHERE id=?""",
                            (completion_status, experiment_status, now() if decision == "批准" else "",
                             data.get("note") or "", now(), reservation["id"]),
                        )
                        self.lab_log(
                            conn, reservation["id"], user, "审核实验结束",
                            f"{decision}：{data.get('note') or ''}",
                        )
                        conn.commit()
                        self.send_json(200, {"ok": True, "review_type": "completion"})
                        return
                    pending = conn.execute(
                        "SELECT COUNT(*) AS n FROM laboratory_reservation_mentor_reviews WHERE reservation_id=? AND status<>'已批准'",
                        (reservation["id"],),
                    ).fetchone()["n"]
                    if pending:
                        return self.error_json(409, "所有相关导师批准后才能进行实验室管理员审核")
                    status = "已批准" if decision == "批准" else "实验室已驳回"
                    conn.execute(
                        """UPDATE laboratory_reservations SET manager_status=?,manager_reviewer=?,manager_note=?,
                           manager_signature_b64=?,status=?,experiment_status=?,started_at=?,updated_at=? WHERE id=?""",
                        ("已批准" if decision == "批准" else "已驳回", reviewer, data.get("note") or "", signature, status,
                         "正在进行" if decision == "批准" else "已拒绝", now() if decision == "批准" else "", now(), reservation["id"]),
                    )
                    self.lab_log(conn, reservation["id"], user, "实验室管理员审核", f"{decision}；{data.get('note') or ''}")
                    if decision == "批准":
                        try:
                            self.generate_lab_pdf(conn, reservation["id"])
                        except Exception as exc:
                            expires = (datetime.now() + timedelta(days=7)).isoformat(timespec="seconds")
                            conn.execute("UPDATE laboratory_reservations SET pdf_expires_at=?,updated_at=? WHERE id=?", (expires, now(), reservation["id"]))
                            self.lab_log(conn, reservation["id"], user, "PDF生成失败", str(exc))
                else:
                    return self.error_json(400, "请明确使用导师审核或实验室管理员审核入口")
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/laboratory/active-for-web":
                rows = [
                    dict(x)
                    for x in conn.execute(
                        """SELECT r.id,r.reservation_no,r.start_time,r.end_time,r.completion_status,r.experiment_status,l.name AS lab_name,l.address
                           FROM laboratory_reservations r JOIN laboratories l ON l.id=r.laboratory_id
                           WHERE (r.requester_id=? OR EXISTS(SELECT 1 FROM laboratory_reservation_participants p WHERE p.reservation_id=r.id AND p.user_id=?))
                           AND r.status='已批准' AND r.experiment_status IN ('未开始','正在进行','待结束确认')
                           ORDER BY r.start_time DESC""",
                        (user["id"],user["id"]),
                    )
                ]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/laboratory/complete-report":
                row = conn.execute("SELECT * FROM laboratory_reservations WHERE id=?", (data.get("id"),)).fetchone()
                if not row:
                    return self.error_json(404, "reservation not found")
                if row["requester_id"] != user["id"]:
                    return self.error_json(403, "only requester can report completion")
                if row["status"] != "已批准":
                    return self.error_json(409, "reservation is not approved")
                conn.execute(
                    "UPDATE laboratory_reservations SET completion_status='待结束确认',experiment_status='待结束确认',completion_report=?,hazard_report=?,updated_at=? WHERE id=?",
                    (data.get("completion_report") or "", data.get("hazard_report") or "", now(), row["id"]),
                )
                self.lab_log(conn, row["id"], user, "提交实验结束报告", data.get("completion_report") or "")
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/laboratory/complete-review":
                row = conn.execute(
                    "SELECT r.*,l.* FROM laboratory_reservations r JOIN laboratories l ON l.id=r.laboratory_id WHERE r.id=?",
                    (data.get("id"),),
                ).fetchone()
                if not row or not self.is_lab_manager(user, row):
                    return self.error_json(403, "only laboratory manager can close reservation")
                if row["completion_status"] != "待结束确认":
                    return self.error_json(409, "completion report not pending")
                decision = data.get("decision")
                if decision not in ("批准", "驳回"):
                    return self.error_json(400, "invalid decision")
                status = "已结束" if decision == "批准" else "需补充"
                conn.execute(
                    "UPDATE laboratory_reservations SET completion_status=?,experiment_status=?,ended_at=?,completion_review_note=?,updated_at=? WHERE id=?",
                    (status, "已结束" if decision == "批准" else "正在进行", now() if decision == "批准" else "", data.get("note") or "", now(), row["id"]),
                )
                self.lab_log(conn, row["id"], user, "审核实验结束", f"{decision}；{data.get('note') or ''}")
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/laboratory/force-stop":
                row = conn.execute(
                    "SELECT r.*,l.* FROM laboratory_reservations r JOIN laboratories l ON l.id=r.laboratory_id WHERE r.id=?",
                    (data.get("id"),),
                ).fetchone()
                if not row or not self.is_lab_manager(user, row):
                    return self.error_json(403, "only laboratory manager can stop an experiment")
                if row["experiment_status"] not in ("未开始", "正在进行", "待结束确认"):
                    return self.error_json(409, "该实验当前不可强制停止")
                reason = str(data.get("reason") or "").strip()
                if not reason:
                    return self.error_json(400, "强制停止必须填写原因")
                conn.execute(
                    """UPDATE laboratory_reservations SET experiment_status='已强制停止',
                       completion_status='已结束',ended_at=?,terminated_by=?,termination_reason=?,updated_at=? WHERE id=?""",
                    (now(), self.user_label(user), reason, now(), row["id"]),
                )
                self.lab_log(conn, row["id"], user, "强制停止实验", reason)
                conn.commit();self.send_json(200, {"ok": True});return

            if self.path == "/api/laboratory/audit-log":
                reservation_no = str(data.get("reservation_no") or "").strip()
                conditions=[];params=[]
                if reservation_no:
                    conditions.append("r.reservation_no LIKE ?");params.append(f"%{reservation_no}%")
                if data.get("laboratory_id"):
                    conditions.append("r.laboratory_id=?");params.append(data.get("laboratory_id"))
                base = """SELECT r.*,l.name AS lab_name,u.display_name,u.username
                          FROM laboratory_reservations r JOIN laboratories l ON l.id=r.laboratory_id
                          JOIN users u ON u.id=r.requester_id"""
                if not self.is_admin(user):
                    if self.normalize_role(user["role"]) == "导师":
                        conditions.append("""(r.requester_id=? OR EXISTS(
                          SELECT 1 FROM laboratory_reservation_mentor_reviews mr WHERE mr.reservation_id=r.id AND mr.mentor_user_id=?
                        ) OR l.team_name=? OR instr(l.managers_json,?)>0)""")
                        params.extend((user["id"],user["id"],user["team_name"],f'"{user["username"]}"'))
                    else:
                        conditions.append("""(r.requester_id=? OR EXISTS(
                          SELECT 1 FROM laboratory_reservation_participants p WHERE p.reservation_id=r.id AND p.user_id=?
                        ))""");params.extend((user["id"],user["id"]))
                sql=base+(" WHERE "+" AND ".join(conditions) if conditions else "")+" ORDER BY r.id DESC LIMIT 1000"
                rows=[dict(x) for x in conn.execute(sql,params)]
                for item in rows:
                    item["logs"]=[dict(x) for x in conn.execute("SELECT * FROM laboratory_audit_logs WHERE reservation_id=? ORDER BY id",(item["id"],))]
                    block=conn.execute("SELECT * FROM laboratory_blacklist WHERE user_id=? AND laboratory_id=? AND active=1 ORDER BY id DESC LIMIT 1",(item["requester_id"],item["laboratory_id"])).fetchone()
                    item["blacklist"]=dict(block) if block else None
                self.cleanup_expired_approval_pdfs(conn);conn.commit()
                self.send_json(200,{"ok":True,"items":rows});return

            if self.path == "/api/laboratory/announcements":
                lab_id=int(data.get("laboratory_id") or 0)
                if data.get("action")=="publish":
                    lab=conn.execute("SELECT * FROM laboratories WHERE id=?",(lab_id,)).fetchone()
                    if not lab or not self.is_lab_manager(user,lab):return self.error_json(403,"您没有该实验室的管理权限")
                    conn.execute("INSERT INTO laboratory_announcements(laboratory_id,title,body,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?)",(lab_id,data.get("title") or "",data.get("body") or "",self.user_label(user),now(),now()));conn.commit()
                rows=[dict(x) for x in conn.execute("SELECT * FROM laboratory_announcements WHERE active=1 AND (?=0 OR laboratory_id=?) ORDER BY id DESC LIMIT 200",(lab_id,lab_id))]
                self.send_json(200,{"ok":True,"items":rows});return

            if self.path == "/api/laboratory/channel":
                lab=conn.execute("SELECT * FROM laboratories WHERE id=?",(data.get("laboratory_id"),)).fetchone()
                if not lab or not self.is_lab_manager(user,lab):return self.error_json(403,"您没有该实验室的管理权限")
                if data.get("action")=="propose":
                    target=data.get("target_state")
                    if target not in ("开启","关闭"):return self.error_json(400,"invalid target state")
                    proposal_id=conn.execute("INSERT INTO laboratory_channel_proposals(laboratory_id,target_state,reason,created_by_id,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(lab["id"],target,data.get("reason") or "",user["id"],self.user_label(user),now(),now())).lastrowid
                    conn.execute("INSERT INTO laboratory_channel_votes(proposal_id,manager_user_id,manager_name,decision,voted_at) VALUES(?,?,?,?,?)",(proposal_id,user["id"],self.user_label(user),"同意",now()))
                elif data.get("action")=="vote":
                    proposal_id=int(data.get("proposal_id") or 0)
                    conn.execute("INSERT OR REPLACE INTO laboratory_channel_votes(proposal_id,manager_user_id,manager_name,decision,voted_at) VALUES(?,?,?,?,?)",(proposal_id,user["id"],self.user_label(user),data.get("decision") or "同意",now()))
                    proposal=conn.execute("SELECT * FROM laboratory_channel_proposals WHERE id=?",(proposal_id,)).fetchone()
                    managers=self.lab_manager_users(conn,lab)
                    votes=list(conn.execute("SELECT * FROM laboratory_channel_votes WHERE proposal_id=?",(proposal_id,)))
                    if any(x["decision"]=="反对" for x in votes):
                        conn.execute("UPDATE laboratory_channel_proposals SET status='已否决',updated_at=? WHERE id=?",(now(),proposal_id))
                    elif managers and {x["manager_user_id"] for x in votes if x["decision"]=="同意"} >= {x["id"] for x in managers}:
                        opening=1 if proposal["target_state"]=="开启" else 0
                        conn.execute("UPDATE laboratories SET booking_open=?,updated_at=? WHERE id=?",(opening,now(),lab["id"]))
                        conn.execute("UPDATE laboratory_channel_proposals SET status='已通过',updated_at=? WHERE id=?",(now(),proposal_id))
                proposals=[dict(x) for x in conn.execute("SELECT * FROM laboratory_channel_proposals WHERE laboratory_id=? ORDER BY id DESC LIMIT 100",(lab["id"],))]
                for p in proposals:p["votes"]=[dict(x) for x in conn.execute("SELECT * FROM laboratory_channel_votes WHERE proposal_id=?",(p["id"],))]
                conn.commit();self.send_json(200,{"ok":True,"booking_open":bool(lab["booking_open"]),"items":proposals});return

            if self.path == "/api/laboratory/blacklist-candidates":
                lab = conn.execute("SELECT * FROM laboratories WHERE id=?", (data.get("laboratory_id"),)).fetchone()
                if not lab or not self.is_lab_manager(user, lab):
                    return self.error_json(403, "only laboratory manager can manage blacklist")
                mentor_name = str(data.get("mentor_name") or "").strip()
                mentors = []
                for row in conn.execute(
                    """SELECT id,username,display_name,role,group_code,team_name FROM users
                       WHERE active=1 AND (?='' OR display_name LIKE ? OR username LIKE ?)
                       ORDER BY display_name,username LIMIT 500""",
                    (mentor_name, f"%{mentor_name}%", f"%{mentor_name}%"),
                ):
                    if self.normalize_role(row["role"]) == "导师":
                        mentors.append(dict(row))
                students = []
                if mentor_name:
                    mentor_ids = [x for x in mentors if mentor_name in (x["display_name"], x["username"])]
                    selected = mentor_ids[0] if mentor_ids else (mentors[0] if mentors else None)
                    if selected:
                        for row in conn.execute(
                            """SELECT * FROM users WHERE active=1 AND role<>'导师' AND group_code=?
                               ORDER BY display_name,username""",
                            (selected["group_code"],),
                        ):
                            item = dict(row)
                            block = self.current_blacklist(conn, lab["id"], row["id"])
                            item["blacklist"] = dict(block) if block else None
                            item["mentor_name"] = self.user_label(selected)
                            students.append(item)
                else:
                    self.cleanup_expired_blacklists(conn)
                    for row in conn.execute(
                        """SELECT u.*,b.id AS blacklist_id,b.blacklist_type,b.reason,b.starts_at,b.ends_at,
                                  b.created_by,b.created_at AS blacklist_created_at
                           FROM laboratory_blacklist b JOIN users u ON u.id=b.user_id
                           WHERE b.laboratory_id=? AND b.active=1 AND u.active=1
                           ORDER BY b.id DESC""",
                        (lab["id"],),
                    ):
                        item = dict(row)
                        item["blacklist"] = {
                            "id": item.pop("blacklist_id"),
                            "blacklist_type": item.pop("blacklist_type"),
                            "reason": item.pop("reason"),
                            "starts_at": item.pop("starts_at"),
                            "ends_at": item.pop("ends_at"),
                            "created_by": item.pop("created_by"),
                            "created_at": item.pop("blacklist_created_at"),
                            "active": 1,
                        }
                        mentor = self.team_mentor(conn, row)
                        item["mentor_name"] = self.user_label(mentor)
                        students.append(item)
                conn.commit()
                self.send_json(200, {"ok": True, "mentors": mentors, "students": students})
                return

            if self.path == "/api/laboratory/permissions":
                self.cleanup_expired_blacklists(conn)
                scope = str(data.get("scope") or "self")
                labs = [dict(x) for x in conn.execute(
                    "SELECT id,name,college,address,lab_type,team_name,booking_open FROM laboratories WHERE active=1 ORDER BY college,name"
                )]
                if scope == "self":
                    items = []
                    for lab in labs:
                        block = self.current_blacklist(conn, lab["id"], user["id"])
                        items.append({
                            **lab,
                            "student_id": user["id"],
                            "student_name": self.user_label(user),
                            "username": user["username"],
                            "mentor_name": self.user_label(self.team_mentor(conn, user)),
                            "allowed": bool(lab["booking_open"]) and not bool(block),
                            "permission_status": "预约通道关闭" if not lab["booking_open"] else ("黑名单限制" if block else "可预约"),
                            "blacklist": dict(block) if block else None,
                        })
                    conn.commit()
                    self.send_json(200, {"ok": True, "labs": labs, "items": items})
                    return
                if scope == "mentor":
                    if self.normalize_role(user["role"]) != "导师" and not self.is_admin(user):
                        return self.error_json(403, "only tutor can view student laboratory permissions")
                    lab_id = int(data.get("laboratory_id") or 0)
                    selected_lab = next((x for x in labs if x["id"] == lab_id), None)
                    if not selected_lab:
                        self.send_json(200, {"ok": True, "labs": labs, "items": []})
                        return
                    if self.is_admin(user):
                        student_rows = conn.execute(
                            "SELECT * FROM users WHERE active=1 AND role<>'导师' ORDER BY display_name,username LIMIT 5000"
                        )
                    else:
                        student_rows = conn.execute(
                            """SELECT * FROM users WHERE active=1 AND role<>'导师'
                               AND (group_code=? OR (team_name<>'' AND team_name=?))
                               ORDER BY display_name,username LIMIT 5000""",
                            (user["group_code"], user["team_name"]),
                        )
                    items = []
                    for student in student_rows:
                        if self.normalize_role(student["role"]) != "学生":
                            continue
                        mentor = self.team_mentor(conn, student)
                        if not self.is_admin(user) and mentor and mentor["id"] != user["id"]:
                            continue
                        block = self.current_blacklist(conn, lab_id, student["id"])
                        items.append({
                            **selected_lab,
                            "student_id": student["id"],
                            "student_name": self.user_label(student),
                            "username": student["username"],
                            "team_name": student["team_name"],
                            "mentor_name": self.user_label(mentor),
                            "allowed": bool(selected_lab["booking_open"]) and not bool(block),
                            "permission_status": "预约通道关闭" if not selected_lab["booking_open"] else ("黑名单限制" if block else "可预约"),
                            "blacklist": dict(block) if block else None,
                        })
                    conn.commit()
                    self.send_json(200, {"ok": True, "labs": labs, "items": items})
                    return
                return self.error_json(400, "invalid permission scope")

            if self.path == "/api/laboratory/blacklist":
                lab = conn.execute("SELECT * FROM laboratories WHERE id=?", (data.get("laboratory_id"),)).fetchone()
                if not lab or not self.is_lab_manager(user, lab):
                    return self.error_json(403, "only laboratory manager can manage blacklist")
                target_id = int(data.get("user_id") or 0)
                if data.get("action") == "remove":
                    conn.execute("UPDATE laboratory_blacklist SET active=0,updated_at=? WHERE laboratory_id=? AND user_id=?", (now(), lab["id"], target_id))
                else:
                    kind = data.get("blacklist_type") or "暂停"
                    if kind == "永久":
                        ends_at = ""
                    else:
                        days = max(1, min(int(data.get("days") or 0), 3650))
                        ends_at = (datetime.now() + timedelta(days=days)).isoformat(timespec="seconds")
                    conn.execute(
                        "INSERT INTO laboratory_blacklist(laboratory_id,user_id,blacklist_type,reason,starts_at,ends_at,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                        (lab["id"], target_id, kind, data.get("reason") or "", now(), ends_at, self.user_label(user), now(), now()),
                    )
                    target=conn.execute("SELECT * FROM users WHERE id=?",(target_id,)).fetchone()
                    conn.execute("INSERT INTO messages(group_code,sender,recipient,subject,body,created_at) VALUES(?,?,?,?,?,?)",(target["group_code"] if target else "",self.user_label(user),target["username"] if target else "","实验室预约黑名单通知",f"您已被加入 {lab['name']} 预约黑名单。类型：{kind}；原因：{data.get('reason') or ''}；自动解除：{ends_at or '永久'}",now()))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/safety/pdf":
                entity = data.get("entity")
                table = "laboratory_reservations" if entity == "laboratory" else "chemical_withdrawals"
                row = conn.execute(f"SELECT pdf_path FROM {table} WHERE id=?", (data.get("id"),)).fetchone()
                if not row or not row["pdf_path"] or not os.path.isfile(row["pdf_path"]):
                    return self.error_json(404, "archived PDF not found")
                if entity == "laboratory":
                    full = conn.execute("SELECT r.*,l.* FROM laboratory_reservations r JOIN laboratories l ON l.id=r.laboratory_id WHERE r.id=?", (data.get("id"),)).fetchone()
                    is_participant = conn.execute(
                        "SELECT 1 FROM laboratory_reservation_participants WHERE reservation_id=? AND user_id=?",
                        (data.get("id"), user["id"]),
                    ).fetchone()
                    is_mentor = conn.execute(
                        "SELECT 1 FROM laboratory_reservation_mentor_reviews WHERE reservation_id=? AND mentor_user_id=?",
                        (data.get("id"), user["id"]),
                    ).fetchone()
                    if not (self.is_admin(user) or full["requester_id"] == user["id"] or is_participant or is_mentor or self.is_lab_manager(user, full)):
                        return self.error_json(403, "no PDF permission")
                else:
                    full = conn.execute("SELECT w.*,c.*,wh.* FROM chemical_withdrawals w JOIN chemicals c ON c.id=w.chemical_id JOIN chemical_warehouses wh ON wh.id=c.warehouse_id WHERE w.id=?", (data.get("id"),)).fetchone()
                    if not (self.is_admin(user) or user["id"] in (full["requester_id"],full["co_collector_id"]) or self.is_warehouse_manager(user, full) or user["display_name"] in (full["requester_teacher"], full["owner_teacher"])):
                        return self.error_json(403, "no PDF permission")
                raw = open(row["pdf_path"], "rb").read()
                self.send_json(200, {"ok": True, "name": os.path.basename(row["pdf_path"]), "content_b64": base64.b64encode(raw).decode("ascii")})
                return

            if self.path == "/api/safety/pdf-delete":
                entity = data.get("entity")
                if entity == "laboratory":
                    row = conn.execute("SELECT r.id,r.pdf_path,l.* FROM laboratory_reservations r JOIN laboratories l ON l.id=r.laboratory_id WHERE r.id=?", (data.get("id"),)).fetchone()
                    allowed = row and (self.is_admin(user) or self.is_lab_manager(user, row))
                    table = "laboratory_reservations"
                else:
                    row = conn.execute("SELECT x.id,x.pdf_path,w.* FROM chemical_withdrawals x JOIN chemicals c ON c.id=x.chemical_id JOIN chemical_warehouses w ON w.id=c.warehouse_id WHERE x.id=?", (data.get("id"),)).fetchone()
                    allowed = row and (self.is_admin(user) or self.is_warehouse_manager(user, row))
                    table = "chemical_withdrawals"
                if not allowed:
                    return self.error_json(403, "only responsible manager or super admin can delete permanent archive")
                if row["pdf_path"] and os.path.isfile(row["pdf_path"]):
                    os.remove(row["pdf_path"])
                conn.execute(f"UPDATE {table} SET pdf_path='',updated_at=? WHERE id=?", (now(), row["id"]))
                conn.commit()
                self.send_json(200, {"ok": True, "warning": "永久审批资料已清理，请确认已完成线下归档备份。"})
                return

            if self.path == "/api/warehouse/upsert":
                if not self.is_admin(user):
                    return self.error_json(403, "only super admin can create warehouse")
                managers = list(dict.fromkeys(str(x) for x in data.get("managers", []) if x))
                if len(managers) != 2:
                    return self.error_json(400, "warehouse requires exactly two tutor managers")
                for username in managers:
                    row = conn.execute("SELECT role FROM users WHERE username=? AND active=1", (username,)).fetchone()
                    if not row or self.normalize_role(row["role"]) != "导师":
                        return self.error_json(400, f"warehouse manager must be tutor: {username}")
                item_id = int(data.get("id") or 0)
                values = (data.get("name") or "", data.get("college") or "", data.get("address") or "", json.dumps(managers, ensure_ascii=False), data.get("commitment_text") or "", 1 if data.get("active", True) else 0, self.user_label(user), now())
                if not all(values[:3]):
                    return self.error_json(400, "name, college and address required")
                if not item_id:
                    existing = conn.execute(
                        """SELECT id FROM chemical_warehouses
                           WHERE name=? AND college=? AND address=?
                           ORDER BY active DESC,id DESC LIMIT 1""",
                        values[:3],
                    ).fetchone()
                    if existing:
                        item_id = existing["id"]
                if item_id:
                    conn.execute("UPDATE chemical_warehouses SET name=?,college=?,address=?,managers_json=?,commitment_text=?,active=?,updated_at=? WHERE id=?", values[:6] + (values[7], item_id))
                else:
                    item_id = conn.execute("INSERT INTO chemical_warehouses(name,college,address,managers_json,commitment_text,active,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?)", values + (values[7],)).lastrowid
                conn.commit()
                self.send_json(200, {"ok": True, "id": item_id})
                return

            if self.path == "/api/warehouse/delete":
                if not self.is_admin(user):
                    return self.error_json(403, "only super admin can delete warehouse")
                warehouse_id = int(data.get("id") or 0)
                row = conn.execute("SELECT * FROM chemical_warehouses WHERE id=?", (warehouse_id,)).fetchone()
                if not row:
                    return self.error_json(404, "warehouse not found")
                pending = conn.execute(
                    """SELECT 1 FROM chemical_withdrawals x JOIN chemicals c ON c.id=x.chemical_id
                       WHERE c.warehouse_id=? AND x.status NOT LIKE '%驳回%' AND x.status<>'已批准并出库'
                       LIMIT 1""",
                    (warehouse_id,),
                ).fetchone()
                if pending:
                    return self.error_json(409, "该库房仍有未完成领用审批，暂不能删除")
                conn.execute("UPDATE chemical_warehouses SET active=0,updated_at=? WHERE id=?", (now(), warehouse_id))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/warehouse/managers":
                if not self.is_admin(user):
                    return self.error_json(403, "only super admin can update warehouse managers")
                warehouse_id = int(data.get("warehouse_id") or data.get("id") or 0)
                warehouse = conn.execute("SELECT * FROM chemical_warehouses WHERE id=?", (warehouse_id,)).fetchone()
                if not warehouse:
                    return self.error_json(404, "warehouse not found")
                managers = list(dict.fromkeys(str(x).strip() for x in data.get("managers", []) if str(x).strip()))[:20]
                if len(managers) < 2:
                    return self.error_json(400, "危险化学品库房至少需要保留两位库房管理员")
                for username in managers:
                    row = conn.execute("SELECT role FROM users WHERE username=? AND active=1", (username,)).fetchone()
                    if not row or self.normalize_role(row["role"]) != "导师":
                        return self.error_json(400, f"库房管理员必须是导师：{username}")
                conn.execute("UPDATE chemical_warehouses SET managers_json=?,updated_at=? WHERE id=?", (json.dumps(managers, ensure_ascii=False), now(), warehouse_id))
                conn.commit()
                self.send_json(200, {"ok": True, "managers": managers})
                return

            if self.path == "/api/warehouse/list":
                include_inactive = bool(data.get("include_inactive")) and self.is_admin(user)
                rows = [dict(x) for x in conn.execute(
                    "SELECT * FROM chemical_warehouses WHERE active=1 OR ?=1 ORDER BY college,name",
                    (1 if include_inactive else 0,),
                )]
                for row in rows:
                    row["managers"] = self.json_list(row.pop("managers_json", "[]"))
                    row["manager_details"] = []
                    for username in row["managers"]:
                        manager = conn.execute(
                            "SELECT username,display_name,role FROM users WHERE username=? AND active=1 ORDER BY id LIMIT 1",
                            (username,),
                        ).fetchone()
                        if manager and self.normalize_role(manager["role"]) != "导师":
                            manager = None
                        row["manager_details"].append({
                            "username": username,
                            "display_name": self.user_label(manager) if manager else username,
                        })
                    row["manager_names"] = [x["display_name"] for x in row["manager_details"]]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/warehouse/commitment":
                warehouse = conn.execute("SELECT * FROM chemical_warehouses WHERE id=?", (data.get("warehouse_id"),)).fetchone()
                if not warehouse or not self.is_warehouse_manager(user, warehouse):
                    return self.error_json(403, "only warehouse manager can update commitment")
                conn.execute("UPDATE chemical_warehouses SET commitment_text=?,updated_at=? WHERE id=?", (data.get("commitment_text") or "", now(), warehouse["id"]))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/warehouse/channel":
                warehouse=conn.execute("SELECT * FROM chemical_warehouses WHERE id=?",(data.get("warehouse_id"),)).fetchone()
                if not warehouse or not self.is_warehouse_manager(user,warehouse):return self.error_json(403,"您没有该危险化学品库房的管理权限")
                managers=[]
                for username in self.json_list(warehouse["managers_json"]):
                    row=conn.execute("SELECT * FROM users WHERE username=? AND active=1",(username,)).fetchone()
                    if row:managers.append(row)
                if data.get("action")=="propose":
                    target=data.get("target_state")
                    if target not in ("开启","关闭"):return self.error_json(400,"invalid target state")
                    proposal_id=conn.execute("INSERT INTO warehouse_channel_proposals(warehouse_id,target_state,reason,created_by_id,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(warehouse["id"],target,data.get("reason") or "",user["id"],self.user_label(user),now(),now())).lastrowid
                    conn.execute("INSERT INTO warehouse_channel_votes(proposal_id,manager_user_id,manager_name,decision,voted_at) VALUES(?,?,?,?,?)",(proposal_id,user["id"],self.user_label(user),"同意",now()))
                elif data.get("action")=="vote":
                    pid=int(data.get("proposal_id") or 0);conn.execute("INSERT OR REPLACE INTO warehouse_channel_votes(proposal_id,manager_user_id,manager_name,decision,voted_at) VALUES(?,?,?,?,?)",(pid,user["id"],self.user_label(user),data.get("decision") or "同意",now()))
                    proposal=conn.execute("SELECT * FROM warehouse_channel_proposals WHERE id=?",(pid,)).fetchone();votes=list(conn.execute("SELECT * FROM warehouse_channel_votes WHERE proposal_id=?",(pid,)))
                    if any(x["decision"]=="反对" for x in votes):conn.execute("UPDATE warehouse_channel_proposals SET status='已否决',updated_at=? WHERE id=?",(now(),pid))
                    elif managers and {x["manager_user_id"] for x in votes if x["decision"]=="同意"} >= {x["id"] for x in managers}:
                        conn.execute("UPDATE chemical_warehouses SET service_open=?,updated_at=? WHERE id=?",(1 if proposal["target_state"]=="开启" else 0,now(),warehouse["id"]));conn.execute("UPDATE warehouse_channel_proposals SET status='已通过',updated_at=? WHERE id=?",(now(),pid))
                items=[dict(x) for x in conn.execute("SELECT * FROM warehouse_channel_proposals WHERE warehouse_id=? ORDER BY id DESC",(warehouse["id"],))]
                for x in items:x["votes"]=[dict(v) for v in conn.execute("SELECT * FROM warehouse_channel_votes WHERE proposal_id=?",(x["id"],))]
                conn.commit();self.send_json(200,{"ok":True,"service_open":bool(warehouse["service_open"]),"items":items});return

            if self.path == "/api/chemical/upsert":
                if self.normalize_role(user["role"]) != "导师" and not self.is_admin(user):
                    return self.error_json(403, "only tutor can register purchased chemicals")
                item_id = int(data.get("id") or 0)
                owner_teacher = data.get("owner_teacher") or self.user_label(user)
                owner_group = data.get("owner_group") or user["group_code"]
                if not self.is_admin(user) and owner_teacher != self.user_label(user):
                    return self.error_json(403, "tutor can only register own chemicals")
                auth_code = data.get("auth_code") or ""
                auth_expires_at = data.get("auth_expires_at") or ""
                values = (
                    int(data.get("warehouse_id") or 0), data.get("name") or "", data.get("unit") or "g",
                    owner_teacher, owner_group, float(data.get("available_per_student") or 0), auth_code, auth_expires_at, now(),
                )
                if item_id:
                    conn.execute("UPDATE chemicals SET warehouse_id=?,name=?,unit=?,owner_teacher=?,owner_group=?,available_per_student=?,auth_code=?,auth_expires_at=?,updated_at=? WHERE id=?", values + (item_id,))
                else:
                    item_id = conn.execute("INSERT INTO chemicals(warehouse_id,name,unit,owner_teacher,owner_group,quantity,available_per_student,auth_code,auth_expires_at,created_at,updated_at) VALUES(?,?,?,?,?,0,?,?,?,?,?)", values[:5] + values[5:8] + (values[8], values[8])).lastrowid
                conn.commit()
                self.send_json(200, {"ok": True, "id": item_id, "auth_code": auth_code})
                return

            if self.path == "/api/chemical/authcode":
                chemical = conn.execute("SELECT * FROM chemicals WHERE id=?", (data.get("chemical_id"),)).fetchone()
                if not chemical:
                    return self.error_json(404, "chemical not found")
                label = self.user_label(user)
                if not self.is_admin(user) and (self.normalize_role(user["role"]) != "导师" or chemical["owner_teacher"] != label):
                    return self.error_json(403, "only chemical owner tutor can generate auth code")
                try:
                    days = max(1, min(int(data.get("days") or 7), 90))
                except Exception:
                    days = 7
                code = secrets.token_urlsafe(10)
                expires = (datetime.now() + timedelta(days=days)).isoformat(timespec="seconds")
                conn.execute("UPDATE chemicals SET auth_code=?,auth_expires_at=?,updated_at=? WHERE id=?", (code, expires, now(), chemical["id"]))
                conn.commit()
                self.send_json(200, {"ok": True, "code": code, "expires_at": expires})
                return

            if self.path == "/api/chemical/list":
                keyword = str(data.get("keyword") or "").strip()
                auth_code = str(data.get("auth_code") or "").strip()
                warehouse_id = int(data.get("warehouse_id") or 0)
                limit = max(1, min(int(data.get("limit") or 500), 1000))
                offset = max(0, int(data.get("offset") or 0))
                pattern = f"%{keyword}%"
                where = """w.active=1 AND (?='' OR c.name LIKE ? OR c.owner_teacher LIKE ?
                           OR w.name LIKE ? OR w.college LIKE ? OR w.address LIKE ?)"""
                params = [keyword, pattern, pattern, pattern, pattern, pattern]
                if warehouse_id:
                    where += " AND w.id=?"
                    params.append(warehouse_id)
                role = self.normalize_role(user["role"])
                if role == "学生":
                    teacher = self.team_mentor(conn, user)
                    owner = self.user_label(teacher)
                    if auth_code:
                        where += " AND (c.owner_teacher=? OR (c.auth_code=? AND c.auth_expires_at>?))"
                        params.extend([owner, auth_code, now()])
                    else:
                        where += " AND c.owner_teacher=?"
                        params.append(owner)
                elif not self.is_admin(user):
                    managed = [x["id"] for x in conn.execute("SELECT * FROM chemical_warehouses WHERE active=1") if self.is_warehouse_manager(user, x)]
                    labels = {self.user_label(user)} if role == "导师" else set()
                    if role == "导师":
                        if user["team_name"]:
                            labels.update(
                                self.user_label(x) for x in conn.execute(
                                    "SELECT * FROM users WHERE active=1 AND team_name=?",
                                    (user["team_name"],),
                                ) if self.normalize_role(x["role"]) == "导师"
                            )
                        labels.update(
                            self.user_label(x) for x in conn.execute(
                                "SELECT * FROM users WHERE active=1 AND group_code=?",
                                (user["group_code"],),
                            ) if self.normalize_role(x["role"]) == "导师"
                        )
                    if managed and labels:
                        wh_placeholders = ",".join("?" for _ in managed)
                        owner_placeholders = ",".join("?" for _ in labels)
                        where += f" AND (w.id IN ({wh_placeholders}) OR c.owner_teacher IN ({owner_placeholders}))"
                        params.extend(managed)
                        params.extend(sorted(labels))
                    elif managed:
                        wh_placeholders = ",".join("?" for _ in managed)
                        where += f" AND w.id IN ({wh_placeholders})"
                        params.extend(managed)
                    elif labels:
                        owner_placeholders = ",".join("?" for _ in labels)
                        where += f" AND c.owner_teacher IN ({owner_placeholders})"
                        params.extend(sorted(labels))
                    else:
                        return self.error_json(403, "no chemical list permission")
                rows = [dict(x) for x in conn.execute(
                    f"""SELECT c.*,w.name AS warehouse_name,w.college,w.address,w.service_open
                        FROM chemicals c JOIN chemical_warehouses w ON w.id=c.warehouse_id
                        WHERE {where} ORDER BY w.name,c.name LIMIT ? OFFSET ?""",
                    tuple(params + [limit, offset]),
                )]
                total = conn.execute(
                    f"""SELECT COUNT(*) AS n FROM chemicals c
                        JOIN chemical_warehouses w ON w.id=c.warehouse_id WHERE {where}""",
                    tuple(params),
                ).fetchone()["n"]
                self.send_json(200, {"ok": True, "items": rows, "total": total, "limit": limit, "offset": offset})
                return

            if self.path == "/api/chemical/stock-summary":
                warehouse_filter = ""
                params = []
                owner_detail_filter = ""
                owner_detail_params = []
                if not self.is_admin(user):
                    managed = [x["id"] for x in conn.execute("SELECT * FROM chemical_warehouses WHERE active=1") if self.is_warehouse_manager(user, x)]
                    if managed:
                        warehouse_filter = f" AND w.id IN ({','.join('?' for _ in managed)})"
                        params.extend(managed)
                    elif self.normalize_role(user["role"]) == "导师":
                        labels = {self.user_label(user)}
                        if user["team_name"]:
                            labels.update(
                                self.user_label(x) for x in conn.execute(
                                    "SELECT * FROM users WHERE active=1 AND team_name=?",
                                    (user["team_name"],),
                                ) if self.normalize_role(x["role"]) == "导师"
                            )
                        labels.update(
                            self.user_label(x) for x in conn.execute(
                                "SELECT * FROM users WHERE active=1 AND group_code=?",
                                (user["group_code"],),
                            ) if self.normalize_role(x["role"]) == "导师"
                        )
                        placeholders = ",".join("?" for _ in labels)
                        warehouse_filter = f" AND c.owner_teacher IN ({placeholders})"
                        params.extend(sorted(labels))
                        owner_detail_filter = f" AND owner_teacher IN ({placeholders})"
                        owner_detail_params = sorted(labels)
                    else:
                        return self.error_json(403, "no chemical summary permission")
                rows = [dict(x) for x in conn.execute(
                    f"""SELECT w.id AS warehouse_id,w.name AS warehouse_name,c.name,c.unit,
                               SUM(c.quantity) AS total_quantity,COUNT(*) AS owner_count
                        FROM chemicals c JOIN chemical_warehouses w ON w.id=c.warehouse_id
                        WHERE w.active=1 {warehouse_filter}
                        GROUP BY w.id,w.name,c.name,c.unit ORDER BY w.name,c.name""",
                    tuple(params),
                )]
                for item in rows:
                    item["owners"] = [dict(x) for x in conn.execute(
                        f"""SELECT owner_teacher,owner_group,quantity,available_per_student,auth_expires_at
                           FROM chemicals WHERE warehouse_id=? AND name=? AND unit=? {owner_detail_filter}
                           ORDER BY owner_teacher""",
                        tuple([item["warehouse_id"], item["name"], item["unit"]] + owner_detail_params),
                    )]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/chemical/approval-logs":
                if self.is_admin(user):
                    warehouse_clause = ""
                    params = []
                else:
                    managed = [x["id"] for x in conn.execute("SELECT * FROM chemical_warehouses WHERE active=1") if self.is_warehouse_manager(user, x)]
                    if not managed:
                        return self.error_json(403, "only warehouse managers can view approval logs")
                    warehouse_clause = f" AND w.id IN ({','.join('?' for _ in managed)})"
                    params = managed
                inbound = [dict(x) for x in conn.execute(
                    f"""SELECT '入库审批' AS kind,i.id,c.name AS chemical_name,c.unit,c.owner_teacher,
                               w.name AS warehouse_name,i.quantity,i.status,i.reviewer AS approver,
                               i.review_note,i.created_at,i.updated_at
                        FROM chemical_inbound_requests i JOIN chemicals c ON c.id=i.chemical_id
                        JOIN chemical_warehouses w ON w.id=c.warehouse_id
                        WHERE 1=1 {warehouse_clause} ORDER BY i.id DESC LIMIT 800""",
                    tuple(params),
                )]
                inventory = [dict(x) for x in conn.execute(
                    f"""SELECT '库存变动' AS kind,l.id,c.name AS chemical_name,c.unit,c.owner_teacher,
                               w.name AS warehouse_name,l.quantity,l.balance_after,l.action,
                               l.operator AS approver,l.note AS review_note,l.created_at,l.created_at AS updated_at
                        FROM chemical_inventory_logs l JOIN chemicals c ON c.id=l.chemical_id
                        JOIN chemical_warehouses w ON w.id=c.warehouse_id
                        WHERE 1=1 {warehouse_clause} ORDER BY l.id DESC LIMIT 800""",
                    tuple(params),
                )]
                self.send_json(200, {"ok": True, "items": sorted(inbound + inventory, key=lambda x: x.get("updated_at") or "", reverse=True)[:1000]})
                return

            if self.path == "/api/chemical/inbound":
                chemical = conn.execute("SELECT * FROM chemicals WHERE id=?", (data.get("chemical_id"),)).fetchone()
                if not chemical:
                    return self.error_json(404, "chemical not found")
                if self.normalize_role(user["role"]) != "导师" or (chemical["owner_teacher"] != self.user_label(user) and not self.is_admin(user)):
                    return self.error_json(403, "only chemical owner tutor can register purchase")
                quantity = float(data.get("quantity") or 0)
                if quantity <= 0:
                    return self.error_json(400, "quantity must be positive")
                item_id = conn.execute(
                    "INSERT INTO chemical_inbound_requests(chemical_id,purchaser_id,quantity,source_note,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                    (chemical["id"], user["id"], quantity, data.get("source_note") or "", now(), now()),
                ).lastrowid
                conn.commit()
                self.send_json(200, {"ok": True, "id": item_id})
                return

            if self.path == "/api/chemical/inbound-review":
                req = conn.execute(
                    "SELECT i.*,c.warehouse_id,c.quantity,w.managers_json FROM chemical_inbound_requests i JOIN chemicals c ON c.id=i.chemical_id JOIN chemical_warehouses w ON w.id=c.warehouse_id WHERE i.id=?",
                    (data.get("id"),),
                ).fetchone()
                if not req or not self.is_warehouse_manager(user, req):
                    return self.error_json(403, "only warehouse manager can review inbound")
                if req["status"] != "待库管审核":
                    return self.error_json(409, "already reviewed")
                decision = data.get("decision")
                if decision == "批准":
                    # The joined duplicate name is driver-dependent; query current stock explicitly.
                    stock = conn.execute("SELECT quantity FROM chemicals WHERE id=?", (req["chemical_id"],)).fetchone()["quantity"]
                    balance = float(stock) + float(req["quantity"])
                    conn.execute("UPDATE chemicals SET quantity=?,updated_at=? WHERE id=?", (balance, now(), req["chemical_id"]))
                    conn.execute("INSERT INTO chemical_inventory_logs(chemical_id,action,quantity,balance_after,related_type,related_id,operator,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)", (req["chemical_id"], "入库", req["quantity"], balance, "inbound", req["id"], self.user_label(user), data.get("note") or "", now()))
                conn.execute("UPDATE chemical_inbound_requests SET status=?,reviewer=?,review_note=?,updated_at=? WHERE id=?", ("已入库" if decision == "批准" else "已驳回", self.user_label(user), data.get("note") or "", now(), req["id"]))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/chemical/inbounds":
                base = """SELECT i.*,c.name AS chemical_name,c.unit,c.owner_teacher,w.name AS warehouse_name,w.managers_json,
                          u.username,u.display_name FROM chemical_inbound_requests i
                          JOIN chemicals c ON c.id=i.chemical_id JOIN chemical_warehouses w ON w.id=c.warehouse_id
                          JOIN users u ON u.id=i.purchaser_id"""
                if self.is_admin(user):
                    rows = [dict(x) for x in conn.execute(base + " ORDER BY i.id DESC LIMIT 1000")]
                elif self.normalize_role(user["role"]) == "导师":
                    label = self.user_label(user)
                    rows = [dict(x) for x in conn.execute(base + " WHERE c.owner_teacher=? OR instr(w.managers_json,?)>0 ORDER BY i.id DESC LIMIT 800", (label, f'"{user["username"]}"'))]
                else:
                    rows = []
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/chemical/withdraw":
                chemical = conn.execute("SELECT c.*,w.commitment_text,w.service_open FROM chemicals c JOIN chemical_warehouses w ON w.id=c.warehouse_id WHERE c.id=?", (data.get("chemical_id"),)).fetchone()
                if not chemical:
                    return self.error_json(404, "chemical not found")
                if not chemical["service_open"]:
                    return self.error_json(409, "危险化学品库房领用通道当前已关闭")
                quantity = float(data.get("quantity") or 0)
                if quantity <= 0 or quantity > float(chemical["quantity"]):
                    return self.error_json(409, "insufficient chemical inventory")
                if chemical["available_per_student"] > 0 and quantity > float(chemical["available_per_student"]):
                    return self.error_json(409, "requested quantity exceeds per-student allowance")
                signature = self.user_signature(conn, user["id"])
                if not signature:
                    return self.error_json(400, "please upload electronic signature first")
                teacher = self.team_mentor(conn, user)
                requester_teacher = self.user_label(teacher)
                cross = chemical["owner_teacher"] != requester_teacher
                co_username=str(data.get("co_collector_username") or "").strip()
                co=conn.execute("SELECT * FROM users WHERE username=? AND active=1",(co_username,)).fetchone()
                if not co or co["id"]==user["id"]:
                    return self.error_json(400,"请选择另一位有效账号作为共同领用人")
                storage=str(data.get("storage_location") or "").strip()
                if not storage:return self.error_json(400,"请填写领用后的存放位置")
                created_at = now()
                item_id = conn.execute(
                    """INSERT INTO chemical_withdrawals(
                       chemical_id,requester_id,requester_group,requester_teacher,owner_teacher,quantity,purpose,
                       commitment_snapshot,student_signature_b64,owner_status,co_collector_id,co_collector_name,
                       storage_location,participant_status,status,created_at,updated_at
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (chemical["id"], user["id"], user["group_code"], requester_teacher, chemical["owner_teacher"], quantity, data.get("purpose") or "", chemical["commitment_text"], signature, "待归属导师审核" if cross else "无需审核",co["id"],self.user_label(co),storage,"待确认","待共同领用人确认",now(),now()),
                ).lastrowid
                withdrawal_no = f"CHEM-{datetime.now():%Y%m%d}-{item_id:06d}"
                conn.execute("UPDATE chemical_withdrawals SET withdrawal_no=?,created_at=?,updated_at=? WHERE id=?", (withdrawal_no, created_at, created_at, item_id))
                conn.execute("""INSERT INTO chemical_withdrawal_participants(
                  withdrawal_id,user_id,participant_name,created_at,updated_at
                ) VALUES(?,?,?,?,?)""",(item_id,co["id"],self.user_label(co),created_at,created_at))
                self.chemical_log(conn,item_id,user,"提交领用申请",f"领用单号：{withdrawal_no}；共同领用人：{self.user_label(co)}；存放位置：{storage}")
                conn.commit()
                self.send_json(200, {"ok": True, "id": item_id, "withdrawal_no": withdrawal_no})
                return

            if self.path == "/api/chemical/withdrawals":
                base = """SELECT x.*,c.name AS chemical_name,c.unit,c.quantity AS stock,c.auth_code,
                          w.name AS warehouse_name,w.address,w.managers_json,u.username,u.display_name
                          FROM chemical_withdrawals x JOIN chemicals c ON c.id=x.chemical_id
                          JOIN chemical_warehouses w ON w.id=c.warehouse_id JOIN users u ON u.id=x.requester_id"""
                if self.is_admin(user):
                    rows = [dict(x) for x in conn.execute(base + " ORDER BY x.id DESC LIMIT 1000")]
                elif self.normalize_role(user["role"]) == "导师":
                    label = self.user_label(user)
                    rows = [dict(x) for x in conn.execute(base + " WHERE x.requester_teacher=? OR x.owner_teacher=? OR instr(w.managers_json,?)>0 ORDER BY x.id DESC LIMIT 800", (label, label, f'"{user["username"]}"'))]
                else:
                    rows = [dict(x) for x in conn.execute(base + " WHERE x.requester_id=? OR x.co_collector_id=? ORDER BY x.id DESC LIMIT 500", (user["id"],user["id"]))]
                for item in rows:
                    managers = self.json_list(item.get("managers_json"))
                    role = self.normalize_role(user["role"])
                    label = self.user_label(user)
                    is_admin = self.is_admin(user)
                    item["can_confirm"]=item.get("co_collector_id")==user["id"] and item.get("participant_status")=="待确认"
                    item["can_review"] = False
                    item["review_stage"] = ""
                    if item.get("mentor_status") == "待导师审核" and item.get("participant_status") == "已同意":
                        item["can_review"] = is_admin or (role == "导师" and label == item.get("requester_teacher"))
                        item["review_stage"] = "申请人导师"
                    elif item.get("owner_status") == "待归属导师审核":
                        item["can_review"] = is_admin or (role == "导师" and label == item.get("owner_teacher"))
                        item["review_stage"] = "化学品归属导师"
                    elif item.get("status") == "待双库管审核":
                        if is_admin:
                            item["can_review"] = item.get("manager1_status") != "已批准" or item.get("manager2_status") != "已批准"
                            item["review_stage"] = "库房管理员"
                        elif user["username"] in managers:
                            idx = managers.index(user["username"])
                            status_field = "manager1_status" if idx == 0 else "manager2_status"
                            item["can_review"] = item.get(status_field) != "已批准"
                            item["review_stage"] = "库房管理员一" if idx == 0 else "库房管理员二"
                    final_disposal_states = ("已报告处理完毕", "已确认处理完毕")
                    item["can_dispose"] = item.get("status") == "已批准并出库" and user["id"] in (item.get("requester_id"), item.get("co_collector_id")) and item.get("disposal_status") not in final_disposal_states and item.get("disposal_status") != "待库管确认"
                    item["can_review_disposal"] = item.get("status") == "已批准并出库" and item.get("disposal_status") == "待库管确认" and (is_admin or user["username"] in managers)
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/chemical/participant-confirm":
                row=conn.execute("SELECT * FROM chemical_withdrawals WHERE id=? AND co_collector_id=?",(data.get("id"),user["id"])).fetchone()
                if not row:return self.error_json(403,"当前账号不是该申请的共同领用人")
                decision=data.get("decision")
                if decision not in ("同意","拒绝"):return self.error_json(400,"invalid decision")
                signature=self.user_signature(conn,user["id"]) if decision=="同意" else ""
                if decision=="同意" and not signature:return self.error_json(400,"同意共同领用前请上传电子签名")
                state="已同意" if decision=="同意" else "已拒绝"
                conn.execute("UPDATE chemical_withdrawal_participants SET confirmation_status=?,confirmation_note=?,signature_b64=?,confirmed_at=?,updated_at=? WHERE withdrawal_id=? AND user_id=?",(state,data.get("note") or "",signature,now(),now(),row["id"],user["id"]))
                conn.execute("UPDATE chemical_withdrawals SET participant_status=?,status=?,updated_at=? WHERE id=?",(state,"待导师审核" if decision=="同意" else "共同领用人已拒绝",now(),row["id"]))
                self.chemical_log(conn,row["id"],user,"共同领用人确认",decision);conn.commit();self.send_json(200,{"ok":True});return

            if self.path == "/api/chemical/withdraw-review":
                with self.server.equipment_write_lock:
                    row = conn.execute(
                        """SELECT x.*,c.quantity AS stock,c.owner_teacher,c.unit,w.managers_json,w.id AS warehouse_id
                           FROM chemical_withdrawals x JOIN chemicals c ON c.id=x.chemical_id
                           JOIN chemical_warehouses w ON w.id=c.warehouse_id WHERE x.id=?""",
                        (data.get("id"),),
                    ).fetchone()
                    if not row:
                        return self.error_json(404, "withdrawal not found")
                    decision = data.get("decision")
                    if decision not in ("批准", "驳回"):
                        return self.error_json(400, "invalid decision")
                    label = self.user_label(user)
                    signature = self.user_signature(conn, user["id"])
                    if decision == "批准" and not signature:
                        return self.error_json(400, "approve requires electronic signature")
                    notes = self.json_dict(row["review_notes_json"])
                    if row["mentor_status"] == "待导师审核":
                        if row["participant_status"]!="已同意":return self.error_json(409,"共同领用人尚未同意")
                        if not (self.is_admin(user) or (self.normalize_role(user["role"]) == "导师" and label == row["requester_teacher"])):
                            return self.error_json(403, "only requester mentor can review")
                        notes["申请人导师"] = data.get("note") or ""
                        conn.execute("UPDATE chemical_withdrawals SET mentor_status=?,mentor_signature_b64=?,status=?,review_notes_json=?,updated_at=? WHERE id=?", ("已批准" if decision == "批准" else "已驳回", signature, "待归属导师审核" if decision == "批准" and row["owner_status"] != "无需审核" else ("待双库管审核" if decision == "批准" else "导师已驳回"), json.dumps(notes, ensure_ascii=False), now(), row["id"]))
                    elif row["owner_status"] == "待归属导师审核":
                        if not (self.is_admin(user) or (self.normalize_role(user["role"]) == "导师" and label == row["owner_teacher"])):
                            return self.error_json(403, "only chemical owner tutor can review")
                        notes["化学品归属导师"] = data.get("note") or ""
                        conn.execute("UPDATE chemical_withdrawals SET owner_status=?,owner_signature_b64=?,status=?,review_notes_json=?,updated_at=? WHERE id=?", ("已批准" if decision == "批准" else "已驳回", signature, "待双库管审核" if decision == "批准" else "归属导师已驳回", json.dumps(notes, ensure_ascii=False), now(), row["id"]))
                    elif row["status"] == "待双库管审核":
                        managers = self.json_list(row["managers_json"])
                        if user["username"] not in managers and not self.is_admin(user):
                            return self.error_json(403, "only designated warehouse managers can review")
                        index = managers.index(user["username"]) if user["username"] in managers else (0 if row["manager1_status"] != "已批准" else 1)
                        field = "manager1" if index == 0 else "manager2"
                        status_field = field + "_status"
                        signature_field = field + "_signature_b64"
                        if row[status_field] == "已批准":
                            return self.error_json(409, "this manager has already approved")
                        notes["库房管理员一" if index == 0 else "库房管理员二"] = data.get("note") or ""
                        new_status = "已批准" if decision == "批准" else "已驳回"
                        conn.execute(f"UPDATE chemical_withdrawals SET {status_field}=?,{signature_field}=?,review_notes_json=?,updated_at=? WHERE id=?", (new_status, signature, json.dumps(notes, ensure_ascii=False), now(), row["id"]))
                        if decision == "驳回":
                            conn.execute("UPDATE chemical_withdrawals SET status='库管已驳回',updated_at=? WHERE id=?", (now(), row["id"]))
                        else:
                            refreshed = conn.execute("SELECT manager1_status,manager2_status FROM chemical_withdrawals WHERE id=?", (row["id"],)).fetchone()
                            if refreshed["manager1_status"] == "已批准" and refreshed["manager2_status"] == "已批准":
                                stock = conn.execute("SELECT quantity FROM chemicals WHERE id=?", (row["chemical_id"],)).fetchone()["quantity"]
                                if float(stock) < float(row["quantity"]):
                                    return self.error_json(409, "inventory changed and is now insufficient")
                                balance = float(stock) - float(row["quantity"])
                                conn.execute("UPDATE chemicals SET quantity=?,updated_at=? WHERE id=?", (balance, now(), row["chemical_id"]))
                                conn.execute("INSERT INTO chemical_inventory_logs(chemical_id,action,quantity,balance_after,related_type,related_id,operator,note,created_at) VALUES(?,?,?,?,?,?,?,?,?)", (row["chemical_id"], "出库", -float(row["quantity"]), balance, "withdrawal", row["id"], label, f"领用人导师：{row['requester_teacher']}；化学品归属：{row['owner_teacher']}", now()))
                                conn.execute("UPDATE chemical_withdrawals SET status='已批准并出库',updated_at=? WHERE id=?", (now(), row["id"]))
                                self.chemical_log(conn,row["id"],user,"危险化学品出库",f"存放位置：{row['storage_location']}")
                                try:
                                    self.generate_chemical_pdf(conn, row["id"])
                                except Exception as exc:
                                    expires = (datetime.now() + timedelta(days=7)).isoformat(timespec="seconds")
                                    conn.execute("UPDATE chemical_withdrawals SET pdf_expires_at=?,updated_at=? WHERE id=?", (expires, now(), row["id"]))
                                    self.chemical_log(conn, row["id"], user, "PDF生成失败", str(exc))
                    else:
                        return self.error_json(409, "当前领用记录不处于可审批阶段")
                    conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/chemical/disposal-report":
                row=conn.execute("SELECT * FROM chemical_withdrawals WHERE id=?",(data.get("id"),)).fetchone()
                if not row or user["id"] not in (row["requester_id"],row["co_collector_id"]):return self.error_json(403,"only collectors can report disposal")
                if row["status"]!="已批准并出库":return self.error_json(409,"该领用记录尚未完成出库")
                if row["disposal_status"] in ("待库管确认", "已报告处理完毕", "已确认处理完毕"):
                    return self.error_json(409,"该领用记录的处置报告已经提交或确认")
                report=str(data.get("report") or "").strip()
                if not report:return self.error_json(400,"请说明危险化学品及废弃物是否按要求处理完毕")
                conn.execute("UPDATE chemical_withdrawals SET disposal_status='待库管确认',disposal_report=?,disposed_at=?,updated_at=? WHERE id=?",(report,now(),now(),row["id"]))
                self.chemical_log(conn,row["id"],user,"提交化学品处置报告",report);conn.commit();self.send_json(200,{"ok":True});return

            if self.path == "/api/chemical/disposal-review":
                row = conn.execute(
                    """SELECT x.*,w.managers_json FROM chemical_withdrawals x
                       JOIN chemicals c ON c.id=x.chemical_id
                       JOIN chemical_warehouses w ON w.id=c.warehouse_id WHERE x.id=?""",
                    (data.get("id"),),
                ).fetchone()
                if not row:
                    return self.error_json(404, "withdrawal not found")
                managers = self.json_list(row["managers_json"])
                if not (self.is_admin(user) or user["username"] in managers):
                    return self.error_json(403, "only warehouse managers can review disposal")
                if row["status"] != "已批准并出库":
                    return self.error_json(409, "该领用记录尚未完成出库")
                if row["disposal_status"] != "待库管确认":
                    return self.error_json(409, "当前处置报告不处于待确认阶段")
                decision = data.get("decision")
                if decision not in ("批准", "驳回"):
                    return self.error_json(400, "invalid decision")
                notes = self.json_dict(row["review_notes_json"])
                notes["处置确认"] = data.get("note") or ""
                new_status = "已确认处理完毕" if decision == "批准" else "处置报告已驳回"
                conn.execute("UPDATE chemical_withdrawals SET disposal_status=?,review_notes_json=?,updated_at=? WHERE id=?", (new_status, json.dumps(notes, ensure_ascii=False), now(), row["id"]))
                self.chemical_log(conn, row["id"], user, "确认化学品处置" if decision == "批准" else "驳回化学品处置报告", data.get("note") or "")
                try:
                    self.generate_chemical_pdf(conn, row["id"])
                except Exception as exc:
                    self.chemical_log(conn, row["id"], user, "处置确认后PDF更新失败", str(exc))
                conn.commit()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/chemical/history":
                if self.normalize_role(user["role"]) == "学生":
                    rows = [dict(x) for x in conn.execute("SELECT l.*,c.name,c.unit,c.owner_teacher FROM chemical_inventory_logs l JOIN chemicals c ON c.id=l.chemical_id WHERE l.related_type='withdrawal' AND l.related_id IN (SELECT id FROM chemical_withdrawals WHERE requester_id=?) ORDER BY l.id DESC", (user["id"],))]
                elif self.normalize_role(user["role"]) == "导师":
                    label = self.user_label(user)
                    rows = [dict(x) for x in conn.execute("SELECT l.*,c.name,c.unit,c.owner_teacher FROM chemical_inventory_logs l JOIN chemicals c ON c.id=l.chemical_id WHERE c.owner_teacher=? OR l.related_id IN (SELECT id FROM chemical_withdrawals WHERE requester_teacher=?) ORDER BY l.id DESC", (label, label))]
                else:
                    rows = [dict(x) for x in conn.execute("SELECT l.*,c.name,c.unit,c.owner_teacher FROM chemical_inventory_logs l JOIN chemicals c ON c.id=l.chemical_id ORDER BY l.id DESC LIMIT 2000")]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/messages/send":
                conn.execute("INSERT INTO messages(group_code,sender,recipient,subject,body,created_at) VALUES(?,?,?,?,?,?)", (group, user["username"], data.get("recipient", ""), data.get("subject", ""), data.get("body", ""), data.get("created_at") or now()))
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/messages/list":
                since = data.get("since") or ""
                try:limit = max(20, min(300, int(data.get("limit", 160))))
                except Exception:limit = 160
                rows = [dict(x) for x in conn.execute("SELECT id,sender,recipient,subject,body,created_at FROM messages WHERE group_code=? AND created_at>? AND (recipient IN ('', '全体成员', ?) OR sender=?) ORDER BY id DESC LIMIT ?", (group, since, user["username"], user["username"], limit))]
                rows.reverse()
                self.send_json(200, {"ok": True, "items": rows, "time": now()})
                return

            if self.path == "/api/v21/records/upsert":
                title=(data.get("title") or "").strip()
                if not title:return self.error_json(400,"title required")
                now_ts=now()
                conn.execute("INSERT INTO v21_records(group_code,creator,kind,title,body,tags,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)",(group,user["username"],data.get("kind","记录"),title,data.get("body",""),data.get("tags",""),now_ts,now_ts))
                conn.commit(); self.server.refresh_async(); self.send_json(200,{"ok":True}); return
            if self.path == "/api/v21/records/list":
                rows=[dict(x) for x in conn.execute("SELECT * FROM v21_records WHERE group_code=? ORDER BY id DESC LIMIT 300",(group,))]
                self.send_json(200,{"ok":True,"items":rows}); return
            if self.path == "/api/v21/notify":
                if not self.is_manager(user):return self.error_json(403,"only manager can notify")
                conn.execute("INSERT INTO v21_notifications(group_code,creator,title,body,level,created_at) VALUES(?,?,?,?,?,?)",(group,user["username"],data.get("title",""),data.get("body",""),data.get("level","普通"),now()))
                conn.commit(); self.server.refresh_async(); self.send_json(200,{"ok":True}); return
            if self.path == "/api/sync/push":
                accepted = 0
                for item in data.get("changes", []):
                    conn.execute("INSERT INTO changes(group_code,sender,entity_type,entity_id,action,payload,created_at) VALUES(?,?,?,?,?,?,?)", (group, user["username"], item.get("entity_type", ""), str(item.get("entity_id", "")), item.get("action", ""), item.get("payload", "{}"), now()))
                    accepted += 1
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True, "accepted": accepted})
                return

            if self.path == "/api/sync/pull":
                since = data.get("since") or ""
                messages = [dict(x) for x in conn.execute("SELECT id,sender,recipient,subject,body,created_at FROM messages WHERE group_code=? AND created_at>? AND (recipient IN ('', '全体成员', ?) OR sender=?) ORDER BY id", (group, since, user["username"], user["username"]))]
                tasks = []
                for row in conn.execute("SELECT payload,created_at FROM changes WHERE group_code=? AND entity_type='research_task' AND created_at>? ORDER BY id", (group, since)):
                    try:
                        payload = json.loads(row["payload"])
                    except Exception:
                        payload = {}
                    if payload:
                        payload.setdefault("created_at", row["created_at"])
                        tasks.append(payload)
                self.send_json(200, {"ok": True, "messages": messages, "tasks": tasks, "time": now()})
                return

            if self.path == "/api/dashboard":
                self.cleanup_expired_files(conn)
                announcements = [dict(x) for x in conn.execute("SELECT sender,subject,body,created_at FROM messages WHERE group_code=? AND subject LIKE '课题组公告:%' ORDER BY id DESC LIMIT 5", (group,))]
                tasks = [dict(x) for x in conn.execute("SELECT title,assignee,status,due_date,review_note,updated_at FROM task_plans WHERE group_code=? AND (assignee IN ('', '全体学生', ?) OR creator=?) ORDER BY id DESC LIMIT 8", (group, user["username"], user["username"]))]
                pending_students = 0
                pending_leaves = 0
                pending_equipment = 0
                team_equipment_pending = 0
                borrowed_count = 0
                if self.is_manager(user):
                    if self.is_admin(user):
                        pending_students = conn.execute("SELECT COUNT(*) FROM users WHERE active=0").fetchone()[0]
                        pending_leaves = conn.execute("SELECT COUNT(*) FROM leave_requests WHERE status='待导师审批'").fetchone()[0]
                        pending_equipment = conn.execute("SELECT COUNT(*) FROM equipment_requests WHERE status='待审批'").fetchone()[0]
                        team_equipment_pending = pending_equipment
                    else:
                        pending_students = conn.execute("SELECT COUNT(*) FROM users WHERE group_code=? AND active=0", (group,)).fetchone()[0]
                        pending_leaves = conn.execute("SELECT COUNT(*) FROM leave_requests WHERE group_code=? AND status='待导师审批'", (group,)).fetchone()[0]
                        pending_equipment = conn.execute("SELECT COUNT(*) FROM equipment_requests r JOIN equipment e ON e.id=r.equipment_id WHERE e.group_code=? AND r.status='待审批'", (group,)).fetchone()[0]
                        team_equipment_pending = pending_equipment
                else:
                    borrowed_count = conn.execute("SELECT COUNT(*) FROM equipment_requests WHERE group_code=? AND requester=? AND status='已批准' AND request_type='借用'", (group, user["username"])).fetchone()[0]
                    pending_equipment = conn.execute("SELECT COUNT(*) FROM equipment_requests r JOIN equipment e ON e.id=r.equipment_id WHERE r.status='待审批' AND EXISTS(SELECT 1 FROM equipment_managers m WHERE m.group_code=e.group_code AND m.username=?)", (user["username"],)).fetchone()[0]
                meeting = conn.execute("SELECT weekday,time_text,location,updated_at FROM meeting_settings WHERE group_code=?", (group,)).fetchone()
                borrowed = [dict(x) for x in conn.execute("SELECT name,brand,model,current_user,approver,status FROM equipment WHERE (group_code=? OR team_name=?) AND status<>'可用' ORDER BY updated_at DESC LIMIT 8", (group, user["team_name"] or ""))]
                self.send_json(200, {"ok": True, "announcements": announcements, "tasks": tasks, "pending_students": pending_students, "pending_leaves": pending_leaves, "pending_equipment": pending_equipment, "team_equipment_pending": team_equipment_pending, "borrowed_count": borrowed_count, "meeting": dict(meeting) if meeting else {}, "borrowed": borrowed})
                return

            if self.path == "/api/leave/submit":
                conn.execute("INSERT INTO leave_requests(group_code,requester,leave_type,start_time,end_time,reason,status,created_at) VALUES(?,?,?,?,?,?,?,?)", (group, user["username"], data.get("leave_type", "请假"), data.get("start_time", ""), data.get("end_time", ""), data.get("reason", ""), "待导师审批", data.get("created_at") or now()))
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/leave/list":
                if self.is_manager(user):
                    rows = [dict(x) for x in conn.execute("SELECT l.*,COALESCE(NULLIF(u.display_name,''),l.requester) AS requester_name FROM leave_requests l LEFT JOIN users u ON u.group_code=l.group_code AND u.username=l.requester WHERE l.group_code=? ORDER BY l.id DESC LIMIT 200", (group,))]
                else:
                    rows = [dict(x) for x in conn.execute("SELECT l.*,COALESCE(NULLIF(u.display_name,''),l.requester) AS requester_name FROM leave_requests l LEFT JOIN users u ON u.group_code=l.group_code AND u.username=l.requester WHERE l.group_code=? AND l.requester=? ORDER BY l.id DESC LIMIT 200", (group, user["username"]))]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/leave/approve":
                if not self.is_manager(user):
                    return self.error_json(403, "only supervisor or admin can approve")
                conn.execute("UPDATE leave_requests SET status=?,approver=?,approved_at=? WHERE id=? AND group_code=?", (data.get("status", "已批准"), user["username"], now(), data.get("id"), group))
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/attendance/checkin":
                ip = self.client_address[0]
                conn.execute("INSERT INTO attendance(group_code,username,action,ip_address,note,created_at) VALUES(?,?,?,?,?,?)", (group, user["username"], data.get("action", "打卡"), ip, data.get("note", ""), data.get("created_at") or now()))
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True, "ip_address": ip})
                return

            if self.path == "/api/attendance/list":
                if not self.is_manager(user):
                    return self.error_json(403, "only supervisor or admin can view attendance records")
                rows = [dict(x) for x in conn.execute("SELECT a.id,a.username,COALESCE(NULLIF(u.display_name,''),a.username) AS display_name,a.action,a.ip_address,a.note,a.created_at FROM attendance a LEFT JOIN users u ON u.group_code=a.group_code AND u.username=a.username WHERE a.group_code=? ORDER BY a.id DESC LIMIT 500", (group,))]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/users/list":
                select_users = """SELECT u.id,u.group_code,u.username,u.display_name,u.role,u.team_name,u.active,
                                  u.created_at,u.mentor_user_id,u.mentor_username,
                                  COALESCE(NULLIF(m.display_name,''),NULLIF(m.username,''),u.group_code) AS advisor_name
                                  FROM users u LEFT JOIN users m ON m.id=u.mentor_user_id"""
                if self.is_admin(user):
                    rows = [dict(x) for x in conn.execute(select_users + " ORDER BY u.team_name,advisor_name,u.active,u.id DESC")]
                elif self.is_manager(user):
                    candidates = [x for x in conn.execute(
                        select_users + " WHERE u.role NOT IN ('导师','老师','教授','PI','teacher','tutor','mentor','超级管理员','管理员') AND (u.mentor_user_id=? OR u.mentor_username=?) ORDER BY u.active,u.id DESC",
                        (user["id"], user["username"]),
                    )]
                    rows = [dict(x) for x in candidates if self.can_approve_account(user, x)]
                else:
                    rows = []
                self.send_json(200, {"ok": True, "items": rows, "capabilities": {"can_approve_accounts": bool(self.is_admin(user) or self.normalize_role(user["role"]) == "导师")}})
                return

            if self.path == "/api/teams/list":
                rows = [dict(x) for x in conn.execute("SELECT DISTINCT team_name FROM users WHERE team_name<>'' ORDER BY team_name")]
                self.send_json(200, {"ok": True, "items": [x["team_name"] for x in rows]})
                return

            if self.path == "/api/users/approve":
                if not self.is_manager(user):
                    return self.error_json(403, "学生账号不能审批账号")
                target = conn.execute(
                    """SELECT u.id,u.group_code,u.username,u.display_name,u.role,u.team_name,u.active,
                              u.created_at,u.mentor_user_id,u.mentor_username,
                              COALESCE(NULLIF(m.display_name,''),NULLIF(m.username,''),u.group_code) AS advisor_name
                       FROM users u LEFT JOIN users m ON m.id=u.mentor_user_id
                       WHERE u.id=?""",
                    (data.get("id"),),
                ).fetchone()
                if not self.can_approve_account(user, target):
                    return self.error_json(403, "导师只能审批明确选择自己为导师的学生账号，不能审批本人或其它角色账号")
                active = 1 if str(data.get("active", 1)) in ("1", "true", "True", "启用", "批准") else 0
                conn.execute("UPDATE users SET active=? WHERE id=?", (active, data.get("id")))
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/users/password":
                target_id = data.get("id") or user["id"]
                new_password = data.get("new_password") or ""
                old_password = data.get("old_password") or ""
                if not new_password:
                    return self.error_json(400, "new password required")
                target = conn.execute("SELECT * FROM users WHERE id=?", (target_id,)).fetchone()
                if not target:
                    return self.error_json(404, "account not found")
                if int(target_id) != int(user["id"]) and not self.is_admin(user):
                    return self.error_json(403, "only super admin can reset others password")
                if int(target_id) == int(user["id"]):
                    _salt, digest = hash_password(old_password, target["salt"])
                    if digest != target["password_hash"]:
                        return self.error_json(403, "old password incorrect")
                salt, digest = hash_password(new_password)
                conn.execute("UPDATE users SET salt=?,password_hash=? WHERE id=?", (salt, digest, target_id))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/users/update":
                if not self.is_admin(user):
                    return self.error_json(403, "only super admin can update accounts")
                target_id = data.get("id")
                fields = []; params = []
                for key in ("username", "display_name", "role", "team_name", "group_code", "active"):
                    if key in data:
                        fields.append(f"{key}=?"); params.append(data.get(key))
                if not fields:
                    return self.error_json(400, "no fields")
                params.append(target_id)
                conn.execute("UPDATE users SET " + ",".join(fields) + " WHERE id=?", params)
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/users/self_delete":
                if self.is_admin(user):
                    return self.error_json(403, "超级管理员不能在客户端注销")
                conn.execute("UPDATE users SET active=0 WHERE id=?", (user["id"],))
                conn.execute("DELETE FROM tokens WHERE user_id=?", (user["id"],))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/users/delete":
                if not self.is_manager(user):
                    return self.error_json(403, "only supervisor or admin can delete accounts")
                target = conn.execute("SELECT id,username,role FROM users WHERE id=? " + ("" if self.is_admin(user) else "AND group_code=?"), (data.get("id"),) if self.is_admin(user) else (data.get("id"), group)).fetchone()
                if not target:
                    return self.error_json(404, "account not found")
                if self.normalize_role(target["role"]) == "超级管理员" and not self.is_admin(user):
                    return self.error_json(403, "不能删除超级管理员账号")
                if self.normalize_role(target["role"]) == "导师" and not self.is_admin(user):
                    return self.error_json(403, "不能删除导师或管理员账号")
                conn.execute("DELETE FROM tokens WHERE user_id=?", (target["id"],))
                conn.execute("DELETE FROM users WHERE id=?", (target["id"],))
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True})
                return

            if self.path == "/api/files/upload":
                self.cleanup_expired_files(conn)
                raw = base64.b64decode(data.get("content_b64", ""))
                safe_name = os.path.basename(data.get("name") or f"upload_{int(datetime.now().timestamp())}")
                folder = os.path.join(FILES_DIR, group)
                os.makedirs(folder, exist_ok=True)
                path = os.path.join(folder, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_name}")
                with open(path, "wb") as fh:
                    fh.write(raw)
                expires_at = (datetime.now() + timedelta(days=7)).isoformat(timespec="seconds")
                cur = conn.execute("INSERT INTO files(group_code,uploader,name,path,size,recipient,encrypted,expires_at,note,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)", (group, user["username"], safe_name, path, len(raw), data.get("recipient", "全体成员"), 1 if data.get("encrypted", True) else 0, expires_at, data.get("note", ""), now()))
                conn.commit()
                self.server.refresh_async()
                self.send_json(200, {"ok": True, "id": cur.lastrowid, "size": len(raw), "expires_at": expires_at})
                return

            if self.path == "/api/files/list":
                self.cleanup_expired_files(conn)
                team_groups=[group]
                if self.is_manager(user) and user["team_name"]:
                    team_groups=[x[0] for x in conn.execute("SELECT DISTINCT group_code FROM users WHERE team_name=? AND role='导师'", (user["team_name"],)).fetchall()] or [group]
                placeholders=",".join("?" for _ in team_groups)
                params=[*team_groups, user["username"], user["username"]]
                rows = [dict(x) for x in conn.execute(f"SELECT id,uploader,recipient,name,size,encrypted,expires_at,note,created_at FROM files WHERE group_code IN ({placeholders}) AND (recipient IN ('', '全体成员', '导师', ?) OR uploader=?) ORDER BY id DESC LIMIT 200", params)]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/files/download":
                self.cleanup_expired_files(conn)
                row = conn.execute("SELECT id,uploader,recipient,name,path,size,encrypted,expires_at,created_at,group_code FROM files WHERE id=?", (data.get("id"),)).fetchone()
                if not row:
                    return self.error_json(404, "file not found")
                same_team_file = False
                if self.is_manager(user) and user["team_name"]:
                    owner = conn.execute("SELECT team_name FROM users WHERE group_code=? AND role='导师' ORDER BY active DESC,id LIMIT 1", (row["group_code"],)).fetchone()
                    same_team_file = bool(owner and owner["team_name"] == user["team_name"] and row["recipient"] in ("全体成员", "导师", ""))
                same_group_file = row["group_code"] == group and (row["recipient"] in ("", "全体成员", "导师", user["username"]) or row["uploader"] == user["username"])
                if not same_group_file and not same_team_file and not self.is_admin(user):
                    return self.error_json(403, "no permission to download this file")
                with open(row["path"], "rb") as fh:
                    raw = fh.read()
                self.send_json(200, {"ok": True, "id": row["id"], "name": row["name"], "uploader": row["uploader"], "recipient": row["recipient"], "size": len(raw), "created_at": row["created_at"], "expires_at": row["expires_at"], "encrypted": bool(row["encrypted"]) or str(row["name"]).endswith(".lspenc"), "content_b64": base64.b64encode(raw).decode("ascii")})
                return

            if self.path == "/api/files/delete":
                row = conn.execute("SELECT id,uploader,path FROM files WHERE group_code=? AND id=?", (group, data.get("id"))).fetchone()
                if not row:return self.error_json(404, "file not found")
                if row["uploader"] != user["username"] and not self.is_manager(user):return self.error_json(403, "no permission to delete this file")
                try:
                    if os.path.exists(row["path"]):os.remove(row["path"])
                except Exception:pass
                conn.execute("DELETE FROM files WHERE id=?", (row["id"],)); conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/taskplan/submit":
                title = (data.get("title") or "").strip()
                if not title:return self.error_json(400, "title required")
                assignee = data.get("assignee", user["username"] if not self.is_manager(user) else "全体学生")
                status = "导师已发布" if self.is_manager(user) else "待导师审批"
                cur = conn.execute("INSERT INTO task_plans(group_code,creator,assignee,title,detail,due_date,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?)", (group, user["username"], assignee, title, data.get("detail",""), data.get("due_date",""), status, now(), now()))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True, "id": cur.lastrowid, "status": status})
                return

            if self.path == "/api/taskplan/list":
                if self.is_manager(user):
                    rows=[dict(x) for x in conn.execute("SELECT * FROM task_plans WHERE group_code=? ORDER BY id DESC LIMIT 300",(group,))]
                else:
                    rows=[dict(x) for x in conn.execute("SELECT * FROM task_plans WHERE group_code=? AND (creator=? OR assignee IN (?, '全体学生', '')) ORDER BY id DESC LIMIT 300",(group,user["username"],user["username"]))]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/taskplan/review":
                if not self.is_manager(user):return self.error_json(403, "only supervisor or admin can review task plans")
                conn.execute("UPDATE task_plans SET status=?,reviewer=?,review_note=?,updated_at=? WHERE id=? AND group_code=?", (data.get("status","已批准"), user["username"], data.get("review_note",""), now(), data.get("id"), group))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/meeting/set":
                if not self.is_manager(user):return self.error_json(403, "only supervisor or admin can set meeting")
                conn.execute("INSERT INTO meeting_settings(group_code,weekday,time_text,location,updated_by,updated_at) VALUES(?,?,?,?,?,?) ON CONFLICT(group_code) DO UPDATE SET weekday=excluded.weekday,time_text=excluded.time_text,location=excluded.location,updated_by=excluded.updated_by,updated_at=excluded.updated_at", (group, data.get("weekday","周五"), data.get("time_text","15:00"), data.get("location",""), user["username"], now()))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/meeting/get":
                row=conn.execute("SELECT * FROM meeting_settings WHERE group_code=?",(group,)).fetchone()
                reports=[dict(x) for x in conn.execute("SELECT r.id,r.student,COALESCE(NULLIF(u.display_name,''),r.student) AS student_name,r.name,r.size,r.status,r.expires_at,r.created_at FROM meeting_reports r LEFT JOIN users u ON u.group_code=r.group_code AND u.username=r.student WHERE r.group_code=? AND r.status<>'已撤回' ORDER BY r.id DESC LIMIT 100",(group,))]
                self.send_json(200, {"ok": True, "meeting": dict(row) if row else {}, "reports": reports})
                return

            if self.path == "/api/meeting/report/upload":
                raw=base64.b64decode(data.get("content_b64","")); safe_name=os.path.basename(data.get("name") or f"meeting_report_{int(datetime.now().timestamp())}.lspenc")
                folder=os.path.join(FILES_DIR, group, "meeting_reports"); os.makedirs(folder,exist_ok=True)
                path=os.path.join(folder, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{user['username']}_{safe_name}")
                with open(path,"wb") as fh:fh.write(raw)
                expires_at=(datetime.now()+timedelta(days=7)).isoformat(timespec="seconds")
                cur=conn.execute("INSERT INTO meeting_reports(group_code,student,name,path,size,status,expires_at,created_at) VALUES(?,?,?,?,?,?,?,?)",(group,user["username"],safe_name,path,len(raw),"已上传",expires_at,now()))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True, "id": cur.lastrowid, "expires_at": expires_at})
                return

            if self.path == "/api/meeting/report/withdraw":
                row=conn.execute("SELECT id,student,path FROM meeting_reports WHERE group_code=? AND id=?",(group,data.get("id"))).fetchone()
                if not row:return self.error_json(404,"report not found")
                if row["student"]!=user["username"] and not self.is_manager(user):return self.error_json(403,"no permission")
                conn.execute("UPDATE meeting_reports SET status='已撤回' WHERE id=?",(row["id"],))
                try:
                    if row["path"] and os.path.exists(row["path"]):os.remove(row["path"])
                except Exception:pass
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/meeting/report/download":
                row=conn.execute("SELECT id,student,name,path,size,status,created_at FROM meeting_reports WHERE group_code=? AND id=?",(group,data.get("id"))).fetchone()
                if not row:return self.error_json(404,"report not found")
                if row["student"]!=user["username"] and not self.is_manager(user):return self.error_json(403,"no permission")
                with open(row["path"],"rb") as fh:raw=fh.read()
                self.send_json(200, {"ok": True, "id": row["id"], "student": row["student"], "name": row["name"], "size": len(raw), "status": row["status"], "created_at": row["created_at"], "encrypted": True, "content_b64": base64.b64encode(raw).decode("ascii")})
                return

            if self.path == "/api/equipment/borrowed-overview":
                if not self.is_manager(user):return self.error_json(403,"only supervisor or admin can view borrowed equipment overview")
                base="""SELECT e.id AS equipment_id,e.name,e.brand,e.category,e.model,e.owner_teacher,e.group_code AS equipment_group_code,e.team_name AS equipment_team_name,e.approver,e.updated_at,
                r.id AS request_id,r.requester,r.requester_teacher,r.group_code AS borrower_group_code,r.created_at AS borrowed_at,
                COALESCE(NULLIF(u.display_name,''),r.requester) AS borrower_name,
                CASE WHEN r.group_code=e.group_code THEN 0 ELSE 1 END AS cross_team
                FROM equipment e
                JOIN equipment_requests r ON r.id=(SELECT rr.id FROM equipment_requests rr WHERE rr.equipment_id=e.id AND rr.status='已批准' ORDER BY rr.id DESC LIMIT 1)
                LEFT JOIN users u ON u.group_code=r.group_code AND u.username=r.requester
                WHERE e.status='使用中' AND r.request_type='借用'"""
                if self.is_admin(user):
                    rows=[dict(x) for x in conn.execute(base+" ORDER BY r.requester_teacher,borrower_name,e.owner_teacher,e.name LIMIT 1000")]
                else:
                    rows=[dict(x) for x in conn.execute(base+" AND r.group_code=? ORDER BY borrower_name,e.owner_teacher,e.name LIMIT 500",(user["group_code"],))]
                self.send_json(200,{"ok":True,"items":rows})
                return

            if self.path == "/api/equipment/owner-overview":
                if not self.is_manager(user):return self.error_json(403,"only supervisor or admin can view owned equipment overview")
                base="""SELECT e.id AS equipment_id,e.name,e.brand,e.category,e.model,e.owner_teacher,e.group_code AS equipment_group_code,
                e.team_name AS equipment_team_name,e.current_user,e.approver AS equipment_approver,e.status AS equipment_status,e.updated_at,
                r.id AS request_id,r.requester,r.requester_teacher,r.group_code AS requester_group_code,r.request_type,r.status AS request_status,
                r.approver AS request_approver,r.review_note,r.created_at AS request_created_at,r.updated_at AS request_updated_at,
                COALESCE(NULLIF(u.display_name,''),r.requester,'') AS requester_name,
                CASE WHEN r.id IS NOT NULL AND r.group_code<>e.group_code THEN 1 ELSE 0 END AS cross_team,
                (SELECT COUNT(*) FROM equipment_requests p WHERE p.equipment_id=e.id AND p.status='待审批') AS pending_count
                FROM equipment e
                LEFT JOIN equipment_requests r ON r.id=(SELECT rr.id FROM equipment_requests rr WHERE rr.equipment_id=e.id ORDER BY rr.id DESC LIMIT 1)
                LEFT JOIN users u ON u.group_code=r.group_code AND u.username=r.requester"""
                if self.is_admin(user):
                    rows=[dict(x) for x in conn.execute(base+" ORDER BY e.owner_teacher,e.status,e.name LIMIT 1000")]
                else:
                    rows=[dict(x) for x in conn.execute(base+" WHERE e.group_code=? ORDER BY e.status,e.name LIMIT 500",(user["group_code"],))]
                self.send_json(200,{"ok":True,"items":rows})
                return

            if self.path == "/api/equipment/force-release":
                if not self.is_admin(user):return self.error_json(403,"only super admin can force release equipment")
                row=conn.execute("SELECT * FROM equipment WHERE id=?",(data.get("id"),)).fetchone()
                if not row:return self.error_json(404,"equipment not found")
                admin_name=user["display_name"] or user["username"]
                latest=conn.execute("SELECT id,request_type FROM equipment_requests WHERE equipment_id=? AND status='已批准' ORDER BY id DESC LIMIT 1",(row["id"],)).fetchone()
                if latest and latest["request_type"]=="借用":
                    conn.execute("UPDATE equipment_requests SET status='已终止',approver=?,review_note=?,updated_at=? WHERE id=?",(admin_name,data.get("note") or "超级管理员强制释放器材使用权",now(),latest["id"]))
                conn.execute("UPDATE equipment_requests SET status='已拒绝',approver=?,review_note=?,updated_at=? WHERE equipment_id=? AND request_type='归还' AND status='待审批'",(admin_name,"器材已由超级管理员直接释放",now(),row["id"]))
                conn.execute("UPDATE equipment SET current_user='',approver=?,status='可用',updated_at=? WHERE id=?",(admin_name,now(),row["id"]))
                conn.commit(); self.server.refresh_async(); self.send_json(200,{"ok":True})
                return

            if self.path == "/api/admin/approvals":
                if not self.is_admin(user):return self.error_json(403,"only super admin can manage all approval states")
                items=[]
                for x in conn.execute("SELECT id,username,display_name,role,group_code,active,created_at FROM users ORDER BY id DESC LIMIT 500"):
                    items.append({"entity_type":"账号","id":x["id"],"subject":x["display_name"] or x["username"],"applicant":x["username"],"group_code":x["group_code"],"status":"已批准" if x["active"] else "停用/待批准","created_at":x["created_at"]})
                for x in conn.execute("SELECT id,requester,leave_type,reason,status,group_code,created_at FROM leave_requests ORDER BY id DESC LIMIT 500"):
                    items.append({"entity_type":"请假","id":x["id"],"subject":x["leave_type"]+"｜"+(x["reason"] or ""), "applicant":x["requester"],"group_code":x["group_code"],"status":x["status"],"created_at":x["created_at"]})
                for x in conn.execute("SELECT id,title,assignee,status,group_code,created_at FROM task_plans ORDER BY id DESC LIMIT 500"):
                    items.append({"entity_type":"任务计划","id":x["id"],"subject":x["title"],"applicant":x["assignee"],"group_code":x["group_code"],"status":x["status"],"created_at":x["created_at"]})
                for x in conn.execute("SELECT id,name,student,status,group_code,created_at FROM meeting_reports ORDER BY id DESC LIMIT 500"):
                    items.append({"entity_type":"组会报告","id":x["id"],"subject":x["name"],"applicant":x["student"],"group_code":x["group_code"],"status":x["status"],"created_at":x["created_at"]})
                for x in conn.execute("SELECT r.id,e.name,r.requester,r.request_type,r.status,r.group_code,r.created_at FROM equipment_requests r JOIN equipment e ON e.id=r.equipment_id ORDER BY r.id DESC LIMIT 800"):
                    items.append({"entity_type":"器材申请","id":x["id"],"subject":x["name"]+"｜"+x["request_type"],"applicant":x["requester"],"group_code":x["group_code"],"status":x["status"],"created_at":x["created_at"]})
                items.sort(key=lambda x:(x.get("created_at") or "",x["id"]),reverse=True)
                self.send_json(200,{"ok":True,"items":items[:1500]})
                return

            if self.path == "/api/admin/approval-status":
                if not self.is_admin(user):return self.error_json(403,"only super admin can modify all approval states")
                entity=str(data.get("entity_type") or ""); item_id=data.get("id"); status=str(data.get("status") or "").strip()
                admin_name=user["display_name"] or user["username"]; note=data.get("review_note") or "超级管理员修正审批状态"
                allowed={
                    "账号":("已批准","停用/待批准"),
                    "请假":("待导师审批","已批准","已驳回","已撤销"),
                    "任务计划":("待导师审批","已批准","已驳回","进行中","已完成"),
                    "组会报告":("已上传","已查看","已通过","需修改","已撤回"),
                    "器材申请":("待审批","已批准","已拒绝","已终止"),
                }
                if entity not in allowed or status not in allowed[entity]:return self.error_json(400,"invalid approval status")
                if entity=="账号":
                    conn.execute("UPDATE users SET active=? WHERE id=?",(1 if status=="已批准" else 0,item_id))
                elif entity=="请假":
                    conn.execute("UPDATE leave_requests SET status=?,approver=?,approved_at=? WHERE id=?",(status,admin_name,now(),item_id))
                elif entity=="任务计划":
                    conn.execute("UPDATE task_plans SET status=?,reviewer=?,review_note=?,updated_at=? WHERE id=?",(status,admin_name,note,now(),item_id))
                elif entity=="组会报告":
                    conn.execute("UPDATE meeting_reports SET status=? WHERE id=?",(status,item_id))
                else:
                    req=conn.execute("SELECT equipment_id FROM equipment_requests WHERE id=?",(item_id,)).fetchone()
                    if not req:return self.error_json(404,"equipment request not found")
                    conn.execute("UPDATE equipment_requests SET status=?,approver=?,review_note=?,updated_at=? WHERE id=?",(status,admin_name,note,now(),item_id))
                    latest=conn.execute("SELECT r.request_type,r.requester,r.group_code,r.requester_teacher,e.owner_teacher FROM equipment_requests r JOIN equipment e ON e.id=r.equipment_id WHERE r.equipment_id=? AND r.status='已批准' ORDER BY r.id DESC LIMIT 1",(req["equipment_id"],)).fetchone()
                    if latest and latest["request_type"]=="借用":
                        borrower=conn.execute("SELECT display_name,username FROM users WHERE group_code=? AND username=?",(latest["group_code"],latest["requester"])).fetchone()
                        borrower_name=(borrower["display_name"] or borrower["username"]) if borrower else latest["requester"]
                        if latest["requester_teacher"] and latest["requester_teacher"]!=latest["owner_teacher"]:borrower_name+=f"（导师：{latest['requester_teacher']}）"
                        conn.execute("UPDATE equipment SET current_user=?,approver=?,status='使用中',updated_at=? WHERE id=?",(borrower_name,admin_name,now(),req["equipment_id"]))
                    else:
                        conn.execute("UPDATE equipment SET current_user='',approver=?,status='可用',updated_at=? WHERE id=?",(admin_name,now(),req["equipment_id"]))
                conn.commit(); self.server.refresh_async(); self.send_json(200,{"ok":True})
                return

            if self.path == "/api/equipment/list":
                keyword=f"%{data.get('keyword','')}%"
                auth_code=(data.get("auth_code") or "").strip()
                auth_groups=[]
                if auth_code:
                    code_row=conn.execute("SELECT group_code FROM auth_codes WHERE code=? AND expires_at>?",(auth_code,now())).fetchone()
                    if code_row:auth_groups.append(code_row["group_code"])
                user_team=user["team_name"] or ""
                if self.is_admin(user):
                    rows=[dict(x) for x in conn.execute("SELECT * FROM equipment WHERE name LIKE ? OR brand LIKE ? OR category LIKE ? OR model LIKE ? OR owner_teacher LIKE ? ORDER BY team_name,group_code,status,name LIMIT 800",(keyword,keyword,keyword,keyword,keyword))]
                else:
                    params=[user_team, group, *auth_groups, user["group_code"], user["username"], keyword, keyword, keyword, keyword, keyword]
                    placeholders=",".join("?" for _ in auth_groups) or "''"
                    rows=[dict(x) for x in conn.execute(f"""SELECT * FROM equipment e
                    WHERE (e.team_name=? OR e.group_code=? OR e.group_code IN ({placeholders})
                      OR e.id IN (
                        SELECT r.equipment_id FROM equipment_requests r
                        WHERE r.group_code=? AND r.requester=? AND r.status='已批准' AND r.request_type='借用'
                          AND r.id=(SELECT rr.id FROM equipment_requests rr WHERE rr.equipment_id=r.equipment_id AND rr.status='已批准' ORDER BY rr.id DESC LIMIT 1)
                      ))
                    AND (e.name LIKE ? OR e.brand LIKE ? OR e.category LIKE ? OR e.model LIKE ? OR e.owner_teacher LIKE ?)
                    ORDER BY e.status,e.name LIMIT 500""",params)]
                for item in rows:
                    latest=conn.execute("SELECT requester,request_type FROM equipment_requests WHERE equipment_id=? AND status='已批准' ORDER BY id DESC LIMIT 1",(item["id"],)).fetchone()
                    item["current_borrower_username"]=(latest["requester"] if latest and latest["request_type"]=="借用" else "")
                    item["can_manage"] = bool(self.is_equipment_manager(conn, user, item))
                    item["can_delete"] = item["can_manage"]
                    item["can_return"] = bool(
                        item["status"] == "使用中"
                        and item["current_borrower_username"] == user["username"]
                    )
                self.send_json(200, {
                    "ok": True,
                    "items": rows,
                    "capabilities": {
                        "can_add": bool(self.is_group_equipment_manager(conn, user, group)),
                        "can_review": bool(self.is_group_equipment_manager(conn, user, group)),
                        "is_super_admin": bool(self.is_admin(user)),
                    },
                })
                return

            if self.path == "/api/equipment/current-borrowed":
                rows=[dict(x) for x in conn.execute("""
                SELECT e.id AS equipment_id,e.name,e.brand,e.category,e.model,e.owner_teacher,
                       e.group_code AS equipment_group_code,e.team_name AS equipment_team_name,
                       e.current_user,e.approver,e.status AS equipment_status,e.updated_at,
                       r.id AS borrow_request_id,r.created_at AS borrowed_at,r.requester_teacher,
                       (SELECT rr.id FROM equipment_requests rr
                        WHERE rr.equipment_id=e.id AND rr.group_code=? AND rr.requester=?
                          AND rr.request_type='归还' AND rr.status='待审批'
                        ORDER BY rr.id DESC LIMIT 1) AS pending_return_id
                FROM equipment e
                JOIN equipment_requests r ON r.id=(
                    SELECT latest.id FROM equipment_requests latest
                    WHERE latest.equipment_id=e.id AND latest.status='已批准'
                    ORDER BY latest.id DESC LIMIT 1
                )
                WHERE e.status='使用中' AND r.request_type='借用'
                  AND r.group_code=? AND r.requester=?
                ORDER BY r.id DESC
                """,(user["group_code"],user["username"],user["group_code"],user["username"]))]
                self.send_json(200,{"ok":True,"items":rows})
                return

            if self.path == "/api/equipment/upsert":
                if data.get("id"):
                    existing=conn.execute("SELECT * FROM equipment WHERE id=?",(data.get("id"),)).fetchone()
                    if not existing:return self.error_json(404,"equipment not found")
                    if not self.is_equipment_manager(conn,user,existing):return self.error_json(403,"only supervisor or equipment manager can edit equipment")
                    target_group=existing["group_code"]
                    target_team=existing["team_name"]
                    owner=existing["owner_teacher"]
                else:
                    target_group=group
                    if not self.is_group_equipment_manager(conn,user,target_group):return self.error_json(403,"only supervisor or equipment manager can edit equipment")
                    teacher=conn.execute("SELECT display_name,username,team_name FROM users WHERE group_code=? AND role='导师' ORDER BY active DESC,id LIMIT 1",(target_group,)).fetchone()
                    owner=(teacher["display_name"] or teacher["username"]) if teacher else (user["display_name"] or user["username"])
                    target_team=(teacher["team_name"] if teacher and teacher["team_name"] else user["team_name"]) or data.get("team_name","")
                if data.get("id"):
                    conn.execute("UPDATE equipment SET name=?,brand=?,category=?,model=?,manager1=?,manager2=?,owner_teacher=?,team_name=?,updated_at=? WHERE id=? AND (group_code=? OR ?=1)", (data.get("name",""), data.get("brand",""), data.get("category",""), data.get("model",""), data.get("manager1",""), data.get("manager2",""), owner, target_team, now(), data.get("id"), target_group, 1 if self.is_admin(user) else 0))
                    eid=data.get("id")
                else:
                    values=(target_team, target_group, owner, data.get("name",""), data.get("brand",""), data.get("category",""), data.get("model",""), data.get("manager1",""), data.get("manager2",""), now(), now())
                    cur=conn.execute("INSERT INTO equipment(team_name,group_code,owner_teacher,name,brand,category,model,manager1,manager2,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)", values)
                    eid=cur.lastrowid
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True, "id": eid})
                return

            if self.path == "/api/equipment/import-csv":
                if not self.is_group_equipment_manager(conn, user, group):
                    return self.error_json(403, "only supervisor or equipment manager can import equipment")
                content_b64 = str(data.get("content_b64") or "")
                if not content_b64:
                    return self.error_json(400, "missing CSV content")
                try:
                    raw = base64.b64decode(content_b64)
                except Exception:
                    return self.error_json(400, "CSV 文件内容无法读取")
                text = ""
                for encoding in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
                    try:
                        text = raw.decode(encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                if not text:
                    return self.error_json(400, "CSV 文件编码无法识别，请使用网页模板重新保存后导入")
                teacher=conn.execute("SELECT display_name,username,team_name FROM users WHERE group_code=? AND role='导师' ORDER BY active DESC,id LIMIT 1",(group,)).fetchone()
                owner=(teacher["display_name"] or teacher["username"]) if teacher else (user["display_name"] or user["username"])
                target_team=(teacher["team_name"] if teacher and teacher["team_name"] else user["team_name"]) or ""
                reader = csv.DictReader(io.StringIO(text))
                if not reader.fieldnames:
                    return self.error_json(400, "CSV 文件缺少表头，请先下载模板后填写")
                reader.fieldnames = [str(x or "").strip().lstrip("\ufeff") for x in reader.fieldnames]
                aliases = {
                    "name": ("器材名称", "实验器材", "设备名称", "名称", "器材参数", "name", "equipment_name"),
                    "brand": ("品牌", "厂家", "brand"),
                    "category": ("类别", "分类", "种类", "category"),
                    "model": ("型号", "规格型号", "model"),
                    "manager1": ("器材管理员账号1", "学生管理员1", "管理员账号1", "manager1"),
                    "manager2": ("器材管理员账号2", "学生管理员2", "管理员账号2", "manager2"),
                }
                imported = 0
                skipped = 0
                for row in reader:
                    def pick(key):
                        for alias in aliases[key]:
                            if alias in row and str(row.get(alias) or "").strip():
                                return str(row.get(alias) or "").strip()
                        return ""
                    if not any(str(v or "").strip() for v in row.values()):
                        continue
                    name = pick("name")
                    if not name:
                        skipped += 1
                        continue
                    payload = {key: pick(key) for key in aliases}
                    existing = conn.execute(
                        """SELECT id FROM equipment WHERE group_code=? AND team_name=? AND name=? AND model=? ORDER BY id LIMIT 1""",
                        (group, target_team, name, payload["model"]),
                    ).fetchone()
                    if existing:
                        conn.execute(
                            "UPDATE equipment SET brand=?,category=?,manager1=?,manager2=?,owner_teacher=?,updated_at=? WHERE id=?",
                            (payload["brand"], payload["category"], payload["manager1"], payload["manager2"], owner, now(), existing["id"]),
                        )
                    else:
                        conn.execute(
                            "INSERT INTO equipment(team_name,group_code,owner_teacher,name,brand,category,model,manager1,manager2,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (target_team, group, owner, name, payload["brand"], payload["category"], payload["model"], payload["manager1"], payload["manager2"], now(), now()),
                        )
                    imported += 1
                conn.commit(); self.server.refresh_async()
                if imported == 0 and skipped == 0:
                    return self.error_json(400, "CSV 文件没有可导入的数据行")
                self.send_json(200, {"ok": True, "imported": imported, "skipped": skipped})
                return

            if self.path == "/api/equipment/delete":
                row=conn.execute("SELECT * FROM equipment WHERE id=?",(data.get("id"),)).fetchone()
                if not row:return self.error_json(404,"equipment not found")
                if not self.is_equipment_manager(conn,user,row):return self.error_json(403,"only supervisor or equipment manager can delete equipment")
                pending=conn.execute("SELECT COUNT(*) FROM equipment_requests WHERE equipment_id=? AND status='待审批'",(row["id"],)).fetchone()[0]
                if pending:return self.error_json(409,"equipment has pending requests")
                conn.execute("DELETE FROM equipment_requests WHERE equipment_id=?",(row["id"],))
                conn.execute("DELETE FROM equipment WHERE id=?",(row["id"],))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/equipment/request":
                row=conn.execute("SELECT * FROM equipment WHERE id=?",(data.get("equipment_id"),)).fetchone()
                if not row:return self.error_json(404,"equipment not found")
                request_type=data.get("request_type",'借用')
                if request_type not in ("借用","归还"):
                    return self.error_json(400,"invalid equipment request type")
                if request_type=='归还':
                    latest=conn.execute("SELECT requester,request_type FROM equipment_requests WHERE equipment_id=? AND status='已批准' ORDER BY id DESC LIMIT 1",(row["id"],)).fetchone()
                    if row["status"]!='使用中' or not latest or latest["request_type"]!='借用':
                        return self.error_json(409,"equipment is not currently borrowed")
                    if latest["requester"]!=user["username"]:
                        return self.error_json(403,"only the current borrower can request return")
                    duplicate=conn.execute("SELECT 1 FROM equipment_requests WHERE equipment_id=? AND requester=? AND request_type='归还' AND status='待审批'",(row["id"],user["username"])).fetchone()
                    if duplicate:return self.error_json(409,"return request already pending")
                elif row["status"]!='可用':
                    return self.error_json(409,"equipment is not available")
                elif conn.execute("SELECT 1 FROM equipment_requests WHERE equipment_id=? AND requester=? AND request_type='借用' AND status='待审批'",(row["id"],user["username"])).fetchone():
                    return self.error_json(409,"borrow request already pending")
                teacher=conn.execute("SELECT display_name,username FROM users WHERE group_code=? AND role='导师' ORDER BY active DESC,id LIMIT 1",(group,)).fetchone()
                requester_teacher=(teacher["display_name"] or teacher["username"]) if teacher else ""
                cur=conn.execute("INSERT INTO equipment_requests(equipment_id,group_code,requester,requester_teacher,request_type,reason,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",(row["id"],group,user["username"],requester_teacher,request_type,data.get("reason",""),'待审批',now(),now()))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True, "id": cur.lastrowid})
                return

            if self.path == "/api/equipment/requests":
                base = "SELECT r.*,COALESCE(NULLIF(u.display_name,''),r.requester) AS requester_name,COALESCE(NULLIF(a.display_name,''),r.approver) AS approver_name,e.name,e.brand,e.category,e.model,e.owner_teacher,e.group_code AS equipment_group_code,e.team_name AS equipment_team_name FROM equipment_requests r JOIN equipment e ON e.id=r.equipment_id LEFT JOIN users u ON u.group_code=r.group_code AND u.username=r.requester LEFT JOIN users a ON a.group_code=e.group_code AND a.username=r.approver"
                scope=str(data.get("scope") or "audit").strip().lower()
                if scope=="mine":
                    rows=[dict(x) for x in conn.execute(base + " WHERE r.requester=? ORDER BY r.id DESC LIMIT 300",(user["username"],))]
                elif scope=="approvable":
                    if self.is_admin(user):
                        rows=[dict(x) for x in conn.execute(base + " WHERE r.status='待审批' ORDER BY r.id DESC LIMIT 500")]
                    elif self.normalize_role(user["role"])=="导师":
                        rows=[dict(x) for x in conn.execute(base + " WHERE r.status='待审批' AND e.group_code=? ORDER BY r.id DESC LIMIT 300",(user["group_code"],))]
                    else:
                        rows=[dict(x) for x in conn.execute(base + " WHERE r.status='待审批' AND (e.manager1=? OR e.manager2=? OR EXISTS(SELECT 1 FROM equipment_managers m WHERE m.group_code=e.group_code AND m.username=?)) ORDER BY r.id DESC LIMIT 300",(user["username"],user["username"],user["username"]))]
                elif self.is_admin(user):
                    rows=[dict(x) for x in conn.execute(base + " ORDER BY r.id DESC LIMIT 500")]
                elif self.normalize_role(user["role"]) == "导师":
                    rows=[dict(x) for x in conn.execute(base + " WHERE e.group_code=? ORDER BY r.id DESC LIMIT 300", (user["group_code"],))]
                else:
                    rows=[dict(x) for x in conn.execute(base + " WHERE r.requester=? OR e.manager1=? OR e.manager2=? OR EXISTS(SELECT 1 FROM equipment_managers m WHERE m.group_code=e.group_code AND m.username=?) ORDER BY r.id DESC LIMIT 300", (user["username"],user["username"],user["username"],user["username"]))]
                self.send_json(200, {"ok": True, "items": rows})
                return

            if self.path == "/api/equipment/review":
                req=conn.execute("SELECT r.*,e.group_code AS equipment_group_code,e.team_name AS equipment_team_name,e.owner_teacher,e.manager1,e.manager2,e.name,e.current_user,e.approver,e.status AS equipment_status FROM equipment_requests r JOIN equipment e ON e.id=r.equipment_id WHERE r.id=?",(data.get("id"),)).fetchone()
                if not req:return self.error_json(404,"request not found")
                if not self.is_equipment_manager(conn,user,req):return self.error_json(403,"no permission")
                status=data.get("status","已批准"); note=data.get("review_note","")
                if status=="已驳回":status="已拒绝"
                if req["status"]!="待审批":return self.error_json(409,"request has already been reviewed")
                if status not in ("已批准","已拒绝"):return self.error_json(400,"invalid review status")
                if status=="已批准" and req["request_type"]=="借用" and req["equipment_status"]!="可用":
                    return self.error_json(409,"equipment is not available")
                if status=="已批准" and req["request_type"]=="归还":
                    latest=conn.execute("SELECT requester,request_type FROM equipment_requests WHERE equipment_id=? AND status='已批准' AND id<? ORDER BY id DESC LIMIT 1",(req["equipment_id"],req["id"])).fetchone()
                    if req["equipment_status"]!="使用中" or not latest or latest["request_type"]!="借用" or latest["requester"]!=req["requester"]:
                        return self.error_json(409,"return request no longer matches the current borrower")
                approver_name=user["display_name"] or user["username"]
                conn.execute("UPDATE equipment_requests SET status=?,approver=?,review_note=?,updated_at=? WHERE id=?",(status,approver_name,note,now(),data.get("id")))
                if status=="已批准" and req["request_type"]=="借用":
                    requester=conn.execute("SELECT display_name,username FROM users WHERE group_code=? AND username=?",(req["group_code"],req["requester"])).fetchone()
                    requester_name=(requester["display_name"] or requester["username"]) if requester else req["requester"]
                    if req["requester_teacher"] and req["requester_teacher"] != req["owner_teacher"]:
                        requester_name=f"{requester_name}（导师：{req['requester_teacher']}）"
                    conn.execute("UPDATE equipment SET current_user=?,approver=?,status='使用中',updated_at=? WHERE id=?",(requester_name,approver_name,now(),req["equipment_id"]))
                    conn.execute("UPDATE equipment_requests SET status='已拒绝',approver=?,review_note='该器材已由其他申请获批',updated_at=? WHERE equipment_id=? AND request_type='借用' AND status='待审批' AND id<>?",(approver_name,now(),req["equipment_id"],req["id"]))
                if status=="已批准" and req["request_type"]=="归还":
                    conn.execute("UPDATE equipment SET current_user='',approver=?,status='可用',updated_at=? WHERE id=?",(approver_name,now(),req["equipment_id"]))
                conn.commit(); self.server.refresh_async(); self.send_json(200, {"ok": True})
                return

            if self.path == "/api/equipment/managers":
                if data.get("action") == "set":
                    if not self.is_manager(user):return self.error_json(403,"only supervisor or admin can set equipment managers")
                    usernames=[str(x).strip() for x in data.get("usernames",[]) if str(x).strip()]
                    conn.execute("DELETE FROM equipment_managers WHERE group_code=?",(group,))
                    for name in usernames:
                        row=conn.execute("SELECT id FROM users WHERE group_code=? AND username=? AND active=1 AND role='学生'",(group,name)).fetchone()
                        if row:
                            conn.execute("INSERT OR IGNORE INTO equipment_managers(group_code,username,created_by,created_at) VALUES(?,?,?,?)",(group,name,user["username"],now()))
                    conn.commit(); self.server.refresh_async()
                managers=[dict(x) for x in conn.execute("SELECT m.username AS username,COALESCE(NULLIF(u.display_name,''),m.username) AS display_name FROM equipment_managers m LEFT JOIN users u ON u.group_code=m.group_code AND u.username=m.username WHERE m.group_code=? ORDER BY display_name,m.username",(group,))]
                students=[dict(x) for x in conn.execute("SELECT username,COALESCE(NULLIF(display_name,''),username) AS display_name FROM users WHERE group_code=? AND active=1 AND role='学生' ORDER BY display_name,username",(group,))]
                self.send_json(200, {"ok": True, "managers": managers, "students": students})
                return

            if self.path == "/api/equipment/authcode":
                if not self.is_group_equipment_manager(conn,user,group):return self.error_json(403,"only supervisor or equipment manager can create auth code")
                teacher=conn.execute("SELECT team_name FROM users WHERE group_code=? AND role='导师' ORDER BY active DESC,id LIMIT 1",(group,)).fetchone()
                team_name=(teacher["team_name"] if teacher and teacher["team_name"] else user["team_name"])
                code=secrets.token_urlsafe(8); expires=(datetime.now()+timedelta(days=7)).isoformat(timespec="seconds")
                conn.execute("INSERT INTO auth_codes(code,group_code,team_name,creator,expires_at,created_at) VALUES(?,?,?,?,?,?)",(code,group,team_name,user["username"],expires,now()))
                conn.commit(); self.send_json(200, {"ok": True, "code": code, "expires_at": expires})
                return

            self.error_json(404, "unknown endpoint")
        except sqlite3.IntegrityError as exc:
            self.error_json(409, str(exc))
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            self.close_connection = True
        except Exception as exc:
            self.error_json(500, str(exc))
        finally:
            if equipment_lock is not None:
                equipment_lock.release()
            conn.close()


class ManagedHTTPServer(ThreadingHTTPServer):
    daemon_threads = True
    request_queue_size = 256

    def __init__(self, address, handler, app):
        self.app = app
        self.config = app.config
        self.db = app.db
        configured = int(self.config.get("max_concurrent_requests", 128) or 128)
        self.max_concurrent_requests = max(16, min(configured, 512))
        self.equipment_write_lock = threading.RLock()
        self.maintenance_lock = threading.Lock()
        self.last_maintenance_at = 0.0
        super().__init__(address, handler)
        self._work_queue = queue.Queue(maxsize=self.request_queue_size)
        self._workers = []
        previous_stack_size = threading.stack_size()
        try:
            threading.stack_size(512 * 1024)
            for index in range(self.max_concurrent_requests):
                worker = threading.Thread(target=self._worker_loop, daemon=True, name=f"LitSearchProHTTP-{index + 1}")
                worker.start()
                self._workers.append(worker)
        finally:
            threading.stack_size(previous_stack_size)

    def process_request(self, request, client_address):
        try:
            self._work_queue.put((request, client_address), timeout=10)
        except queue.Full:
            try:
                request.sendall(b"HTTP/1.1 503 Service Unavailable\r\nContent-Type: application/json\r\nConnection: close\r\nContent-Length: 50\r\n\r\n{\"ok\":false,\"error\":\"server is busy, retry later\"}")
            except Exception:
                pass
            self.shutdown_request(request)

    def _worker_loop(self):
        while True:
            item = self._work_queue.get()
            try:
                if item is None:
                    return
                request, client_address = item
                try:
                    self.finish_request(request, client_address)
                except Exception:
                    self.handle_error(request, client_address)
                finally:
                    self.shutdown_request(request)
            finally:
                self._work_queue.task_done()

    def server_close(self):
        super().server_close()
        for _worker in self._workers:
            self._work_queue.put(None)
        for worker in self._workers:
            worker.join(timeout=2)

    def app_log(self, text):
        self.app.log(text)

    def handle_error(self, request, client_address):
        exc_type, exc, _tb = sys.exc_info()
        if isinstance(exc, (BrokenPipeError, ConnectionAbortedError, ConnectionResetError, TimeoutError)):
            return
        super().handle_error(request, client_address)

    def refresh_async(self):
        self.app.refresh_requested.set()


class ServerApp:
    def __init__(self, root):
        ensure_dirs()
        self.root = root
        self.root.report_callback_exception = self.handle_exception
        self.config = load_config()
        apply_data_dir(self.config)
        self.db = connect_db()
        self.httpd = None
        self.thread = None
        self.tray_icon = None
        self.tray_popup = None
        self.tray_events = queue.Queue()
        self.refresh_requested = threading.Event()
        self.native_tray = None
        self.native_tray_hwnd = None
        self.native_tray_wndproc = None
        self.refresh_job = None
        self.sensitive_tabs = {"账号管理", "请假审批", "打卡记录"}
        self.sensitive_unlocked_until = {}
        self.sensitive_lock_minutes = int(self.config.get("sensitive_lock_minutes", 5) or 5)
        self.sensitive_lock_job = None
        self.status = tk.StringVar(value="服务未启动")
        self.public_url = tk.StringVar(value=self.config.get("public_url", ""))
        self.port = tk.IntVar(value=int(self.config.get("port", 8765)))
        self.group_code = tk.StringVar(value=self.config.get("group_code", "research-lab"))
        self.data_dir = tk.StringVar(value=self.config.get("data_dir", APP_DIR))
        self.root.title(f"LitSearchPro 通用协作服务器管理端 v{SERVER_VERSION}")
        self.root.geometry("1080x720")
        self.root.minsize(920, 600)
        self.root.configure(bg=UX_BG)
        self.configure_style()
        self.build()
        self.first_run_admin()
        self.start_server()
        self.setup_tray_icon()
        self.poll_tray_events()
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_background)

    def handle_exception(self, exc_type, exc_value, exc_tb):
        detail = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        write_server_log("server_crash.log", detail)
        try:
            self.log(f"GUI exception captured: {exc_value}")
        except Exception:
            pass

    def configure_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook", background=UX_BG, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 8), background="#EEF3F9", foreground=UX_TEXT)
        style.map("TNotebook.Tab", background=[("selected", UX_SURFACE)], foreground=[("selected", UX_ACCENT)])
        style.configure("Treeview", background=UX_SURFACE, fieldbackground=UX_SURFACE, foreground=UX_TEXT, rowheight=28, borderwidth=0)
        style.configure("Treeview.Heading", background="#EEF3F9", foreground=UX_TEXT, relief="flat", font=("Microsoft YaHei UI", 9, "bold"))
        style.map("Treeview", background=[("selected", UX_HOVER)], foreground=[("selected", UX_ACCENT)])
        style.configure("TCombobox", fieldbackground=UX_SURFACE, background=UX_SURFACE, foreground=UX_TEXT, padding=5)
        style.configure("TSpinbox", fieldbackground=UX_SURFACE, background=UX_SURFACE, foreground=UX_TEXT, padding=5)

    def button(self, parent, text, command, kind="secondary"):
        return ServerRoundedButton(parent, text=text, command=command, kind=kind)

    def build(self):
        header = tk.Frame(self.root, bg=UX_ACCENT, height=94)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="科研文献与实验室安全管理平台 协作服务器", bg=UX_ACCENT, fg="white", font=("Microsoft YaHei UI", 17, "bold")).pack(anchor=tk.W, padx=22, pady=(14, 2))
        tk.Label(header, text=f"服务器体验优化版 v{SERVER_VERSION}    双击托盘图标可恢复窗口，右键可打开菜单", bg=UX_ACCENT, fg="#DCEEFF", font=("Microsoft YaHei UI", 9)).pack(anchor=tk.W, padx=24)
        tk.Label(header, textvariable=self.status, bg=UX_ACCENT, fg="#DCEEFF", font=("Microsoft YaHei UI", 10)).pack(anchor=tk.W, padx=24, pady=(2, 0))
        main = tk.Frame(self.root, bg=UX_BG)
        main.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)
        top = tk.Frame(main, bg=UX_SURFACE, highlightthickness=1, highlightbackground=UX_BORDER)
        top.pack(fill=tk.X, pady=(0, 10))
        tk.Label(top, text="端口", bg=UX_SURFACE, fg=UX_TEXT).grid(row=0, column=0, padx=12, pady=10)
        ttk.Spinbox(top, from_=1, to=65535, textvariable=self.port, width=8).grid(row=0, column=1, padx=4)
        tk.Label(top, text="团队代码", bg=UX_SURFACE, fg=UX_TEXT).grid(row=0, column=2, padx=12)
        self.group_combo = ttk.Combobox(top, textvariable=self.group_code, width=16, values=[])
        self.group_combo.grid(row=0, column=3, padx=4, ipady=4)
        tk.Label(top, text="对外访问地址", bg=UX_SURFACE, fg=UX_TEXT).grid(row=0, column=4, padx=12)
        tk.Entry(top, textvariable=self.public_url, bg="#FAFBFD", fg=UX_TEXT, relief=tk.FLAT, highlightthickness=1, highlightbackground=UX_BORDER).grid(row=0, column=5, sticky="ew", padx=4, ipady=6)
        top.columnconfigure(5, weight=1)
        self.button(top, "保存并重启服务", self.restart_server, "primary").grid(row=0, column=6, padx=12)
        tk.Label(top, text="数据目录", bg=UX_SURFACE, fg=UX_TEXT).grid(row=1, column=0, padx=12, pady=(0,10))
        tk.Entry(top, textvariable=self.data_dir, bg="#FAFBFD", fg=UX_TEXT, relief=tk.FLAT, highlightthickness=1, highlightbackground=UX_BORDER).grid(row=1, column=1, columnspan=5, sticky="ew", padx=4, ipady=6)
        self.button(top, "选择...", self.choose_data_dir, "secondary").grid(row=1, column=6, padx=12, pady=(0,10))
        self.nb = ttk.Notebook(main)
        self.nb.pack(fill=tk.BOTH, expand=True)
        self.tabs = {}
        for name in ("概览", "账号管理", "上传文件", "请假审批", "打卡记录", "运行日志"):
            frame = tk.Frame(self.nb, bg=UX_SURFACE)
            self.nb.add(frame, text=name)
            self.tabs[name] = frame
        self.build_overview()
        self.build_table_tabs()
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.root.bind_all("<Key>", self.mark_sensitive_activity, add="+")
        self.root.bind_all("<Button>", self.mark_sensitive_activity, add="+")
        self.check_sensitive_locks()

    def selected_tab_name(self):
        try:
            tab_id = self.nb.select()
            return self.nb.tab(tab_id, "text") if tab_id else ""
        except Exception:
            return ""

    def admin_password_ok(self, password):
        if not password:
            return False
        try:
            rows = self.db.execute("SELECT salt,password_hash FROM users WHERE active=1 AND role IN ('超级管理员','管理员') ORDER BY role DESC,id").fetchall()
            for row in rows:
                _salt, digest = hash_password(password, row["salt"])
                if digest == row["password_hash"]:
                    return True
        except Exception as exc:
            self.log("sensitive auth failed: "+str(exc))
        return False

    def is_sensitive_unlocked(self, name):
        return self.sensitive_unlocked_until.get(name, 0) > time.time()

    def mark_sensitive_activity(self, _event=None):
        name = self.selected_tab_name()
        if name in self.sensitive_tabs and self.is_sensitive_unlocked(name):
            self.sensitive_unlocked_until[name] = time.time() + self.sensitive_lock_minutes * 60

    def unlock_sensitive_tab(self, name):
        password = simpledialog.askstring("敏感页面验证", f"进入“{name}”需要输入服务器管理员密码：", show="*", parent=self.root)
        if self.admin_password_ok(password):
            self.sensitive_unlocked_until[name] = time.time() + self.sensitive_lock_minutes * 60
            self.status.set(f"{name} 已解锁，空闲 {self.sensitive_lock_minutes} 分钟后自动锁定。")
            self.show_sensitive_content(name)
            return True
        messagebox.showwarning("验证失败", "密码不正确，已返回概览页面。", parent=self.root)
        return False

    def show_sensitive_content(self, name):
        overlay = getattr(self, "sensitive_overlays", {}).get(name)
        if overlay:
            try: overlay.place_forget()
            except Exception: pass
        self.mark_sensitive_activity()

    def lock_sensitive_tab(self, name):
        self.sensitive_unlocked_until[name] = 0
        overlay = getattr(self, "sensitive_overlays", {}).get(name)
        if overlay:
            try: overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
            except Exception: pass

    def on_tab_changed(self, _event=None):
        name = self.selected_tab_name()
        if name not in self.sensitive_tabs:
            return
        if self.is_sensitive_unlocked(name):
            self.show_sensitive_content(name)
            return
        self.lock_sensitive_tab(name)
        if not self.unlock_sensitive_tab(name):
            try:self.nb.select(0)
            except Exception:pass

    def check_sensitive_locks(self):
        now_ts = time.time()
        for name in list(self.sensitive_tabs):
            if self.sensitive_unlocked_until.get(name, 0) and self.sensitive_unlocked_until.get(name, 0) <= now_ts:
                self.lock_sensitive_tab(name)
                if self.selected_tab_name() == name:
                    self.status.set(f"{name} 已因空闲超时自动锁定。")
        try:self.sensitive_lock_job = self.root.after(15000, self.check_sensitive_locks)
        except Exception:pass

    def protect_sensitive_tab(self, name):
        frame = self.tabs.get(name)
        if not frame:
            return
        if not hasattr(self, "sensitive_overlays"):
            self.sensitive_overlays = {}
        overlay = tk.Frame(frame, bg=UX_SURFACE, highlightthickness=1, highlightbackground=UX_BORDER)
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        card = tk.Frame(overlay, bg="#F8FAFD", highlightthickness=1, highlightbackground=UX_BORDER)
        card.place(relx=0.5, rely=0.42, anchor=tk.CENTER, width=430, height=190)
        tk.Label(card, text="敏感管理页面已锁定", bg="#F8FAFD", fg=UX_TEXT, font=("Microsoft YaHei UI", 14, "bold")).pack(pady=(26, 8))
        tk.Label(card, text=f"进入“{name}”需要服务器管理员密码。\n解锁后空闲 {self.sensitive_lock_minutes} 分钟会自动锁定。", bg="#F8FAFD", fg=UX_MUTED, justify=tk.CENTER).pack(pady=(0, 18))
        self.button(card, "输入密码解锁", lambda n=name: self.unlock_sensitive_tab(n), "primary").pack()
        self.sensitive_overlays[name] = overlay

    def build_overview(self):
        f = self.tabs["概览"]
        self.overview = tk.Text(f, wrap=tk.WORD, bg=UX_SURFACE, fg=UX_TEXT, relief=tk.FLAT, padx=18, pady=18, font=("Microsoft YaHei UI", 10), spacing3=4)
        self.overview.pack(fill=tk.BOTH, expand=True)

    def make_tree(self, tab, columns):
        frame = self.tabs[tab]
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings")
        for key, title, width in columns:
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor=tk.W)
        tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        return tree

    def build_table_tabs(self):
        self.user_tree = self.make_tree("账号管理", [("id", "ID", 55), ("team_name", "团队代码", 120), ("group_code", "导师姓名", 110), ("username", "用户名", 120), ("display_name", "姓名", 100), ("role", "角色", 90), ("active", "状态", 80), ("created_at", "注册时间", 150)])
        user_actions = tk.Frame(self.tabs["账号管理"], bg=UX_SURFACE)
        user_actions.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.button(user_actions, "批准账号", lambda: self.set_user_active(1), "primary").pack(side=tk.RIGHT)
        self.button(user_actions, "停用账号", lambda: self.set_user_active(0), "secondary").pack(side=tk.RIGHT, padx=8)
        self.button(user_actions, "新增账号", self.add_user_dialog, "secondary").pack(side=tk.RIGHT)
        self.file_tree = self.make_tree("上传文件", [("id", "ID", 55), ("uploader", "上传者", 120), ("name", "文件名", 300), ("size", "大小", 90), ("created_at", "时间", 160)])
        self.button(self.tabs["上传文件"], "打开上传目录", lambda: os.startfile(FILES_DIR), "secondary").pack(anchor=tk.E, padx=12, pady=(0, 10))
        self.leave_tree = self.make_tree("请假审批", [("id", "ID", 55), ("requester", "申请人", 110), ("leave_type", "类型", 80), ("time", "时间", 260), ("status", "状态", 110), ("reason", "原因", 260)])
        leave_actions = tk.Frame(self.tabs["请假审批"], bg=UX_SURFACE)
        leave_actions.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.button(leave_actions, "批准", lambda: self.set_leave_status("已批准"), "primary").pack(side=tk.RIGHT)
        self.button(leave_actions, "驳回", lambda: self.set_leave_status("已驳回"), "secondary").pack(side=tk.RIGHT, padx=8)
        self.attendance_tree = self.make_tree("打卡记录", [("id", "ID", 55), ("username", "用户", 120), ("action", "动作", 80), ("ip_address", "IP地址", 150), ("created_at", "时间", 180)])
        for tab_name in self.sensitive_tabs:
            self.protect_sensitive_tab(tab_name)
        self.log_text = tk.Text(self.tabs["运行日志"], wrap=tk.WORD, bg=UX_SURFACE, fg=UX_TEXT, relief=tk.FLAT, padx=12, pady=12, font=("Consolas", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def first_run_admin(self):
        row = self.db.execute("SELECT COUNT(*) FROM users WHERE role IN ('超级管理员','管理员')").fetchone()[0]
        if row:
            return
        username = simpledialog.askstring("首次设置", "请设置服务器超级管理员用户名：", initialvalue="admin", parent=self.root)
        if not username:
            username = "admin"
        password = simpledialog.askstring("首次设置", "请设置服务器超级管理员密码：", show="*", parent=self.root)
        if not password:
            password = "ChangeMe_2026!"
        salt, digest = hash_password(password)
        self.db.execute("INSERT OR REPLACE INTO users(group_code,username,display_name,role,team_name,salt,password_hash,active,created_at) VALUES(?,?,?,?,?,?,?,?,?)", (self.group_code.get(), username, username, "超级管理员", self.group_code.get(), salt, digest, 1, now()))
        self.db.commit()
        messagebox.showinfo("首次设置完成", "超级管理员账号已创建。请在正式使用前妥善保存密码。", parent=self.root)

    def save_runtime_config(self):
        self.config.update({"host": "0.0.0.0", "port": int(self.port.get()), "public_url": self.public_url.get().strip(), "group_code": self.group_code.get().strip() or "research-lab", "data_dir": self.data_dir.get().strip() or APP_DIR})
        save_config(self.config)

    def choose_data_dir(self):
        path = filedialog.askdirectory(title="选择服务器数据存放目录", initialdir=self.data_dir.get() or APP_DIR)
        if path:self.data_dir.set(path)

    def start_server(self):
        self.save_runtime_config()
        if self.httpd:
            return
        try:
            self.httpd = ManagedHTTPServer((self.config["host"], int(self.config["port"])), ApiHandler, self)
            self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True, name="LitSearchProHTTPServer")
            self.thread.start()
            self.status.set(f"服务运行中：0.0.0.0:{self.config['port']}")
            self.log(f"HTTP server started on 0.0.0.0:{self.config['port']}")
        except OSError as exc:
            self.httpd = None
            self.status.set(f"服务未启动：端口 {self.config['port']} 不可用")
            self.log(f"HTTP server start failed: {exc}")
            try:
                messagebox.showwarning("服务未启动", f"端口 {self.config['port']} 当前不可用，服务器管理端会继续打开。\n\n请更换端口后点击“保存并重启服务”。\n\n错误信息：{exc}", parent=self.root)
            except Exception:
                pass
        except Exception as exc:
            self.httpd = None
            self.status.set("服务未启动：启动异常")
            self.handle_exception(type(exc), exc, exc.__traceback__)
        self.refresh_all()

    def restart_server(self):
        self.stop_server()
        self.start_server()

    def schedule_refresh(self, delay=800):
        if self.refresh_job:
            return
        def run():
            self.refresh_job = None
            self.refresh_all()
        try:self.refresh_job = self.root.after(delay, run)
        except Exception:pass

    def stop_server(self):
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()
            self.httpd = None
        self.status.set("服务已停止")

    def hide_to_background(self):
        self.root.withdraw()
        if not self.tray_icon and not self.native_tray:
            try:
                messagebox.showinfo("后台运行", "协作服务器仍在后台运行。当前环境未启用托盘图标，可从任务栏或进程管理器管理。", parent=self.root)
            except Exception:
                pass

    def setup_tray_icon(self):
        if os.name == "nt":
            self.setup_native_tray_icon()
            return
        if not pystray or not Image:
            self.setup_native_tray_icon()
            return
        image = Image.new("RGB", (64, 64), "#02529F")
        draw = ImageDraw.Draw(image)
        draw.ellipse((8, 8, 56, 56), fill="white")
        draw.text((20, 22), "LSP", fill="#02529F")
        def show(_icon=None, _item=None):
            self.root.after(0, self.show_main_window)
        def quit_app(_icon=None, _item=None):
            self.root.after(0, self.quit_app)
        self.tray_icon = pystray.Icon("LitSearchProServer", image, "LitSearchPro 协作服务器", menu=pystray.Menu(pystray.MenuItem("打开管理端", show), pystray.MenuItem("退出服务器", quit_app)))
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_main_window(self):
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            pass

    def poll_tray_events(self):
        try:
            if self.refresh_requested.is_set():
                self.refresh_requested.clear()
                self.schedule_refresh()
            while True:
                item = self.tray_events.get_nowait()
                action = item[0]
                if action == "show":
                    self.show_main_window()
                elif action == "menu":
                    self.show_tray_menu(item[1], item[2])
        except queue.Empty:
            pass
        except Exception as exc:
            self.handle_exception(type(exc), exc, exc.__traceback__)
        try:
            self.root.after(160, self.poll_tray_events)
        except Exception:
            pass

    def show_tray_menu(self, x, y):
        try:
            if self.tray_popup and self.tray_popup.winfo_exists():
                self.tray_popup.destroy()
        except Exception:
            pass
        popup = tk.Toplevel(self.root)
        self.tray_popup = popup
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg=UX_BORDER)
        popup.geometry(f"190x120+{int(x)}+{int(y)}")
        box = tk.Frame(popup, bg=UX_SURFACE, highlightthickness=1, highlightbackground=UX_BORDER)
        box.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        tk.Label(box, text="协作服务器", bg=UX_SURFACE, fg=UX_MUTED, font=("Microsoft YaHei UI", 9)).pack(anchor=tk.W, padx=14, pady=(10, 5))

        def close_then(action):
            try:
                popup.destroy()
            except Exception:
                pass
            action()

        self.button(box, "打开管理端", lambda: close_then(self.show_main_window), "primary").pack(fill=tk.X, padx=12, pady=(0, 8))
        self.button(box, "退出服务器", lambda: close_then(self.quit_app), "danger").pack(fill=tk.X, padx=12)
        popup.bind("<Escape>", lambda _e: popup.destroy())
        popup.bind("<FocusOut>", lambda _e: popup.destroy())
        popup.after(80, popup.focus_force)

    def setup_native_tray_icon(self):
        if os.name != "nt":
            return
        try:
            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32
            kernel32 = ctypes.windll.kernel32
            kernel32.GetModuleHandleW.argtypes = [ctypes.c_wchar_p]
            kernel32.GetModuleHandleW.restype = ctypes.c_void_p
            user32.DefWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
            user32.DefWindowProcW.restype = ctypes.c_void_p
            user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
            user32.RegisterClassW.restype = ctypes.c_ushort
            user32.CreateWindowExW.argtypes = [ctypes.c_ulong, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_ulong, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
            user32.CreateWindowExW.restype = ctypes.c_void_p
            user32.GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
            user32.GetCursorPos.restype = ctypes.c_int
            user32.LoadImageW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
            user32.LoadImageW.restype = ctypes.c_void_p
            user32.LoadIconW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
            user32.LoadIconW.restype = ctypes.c_void_p
            shell32.Shell_NotifyIconW.argtypes = [ctypes.c_ulong, ctypes.POINTER(NOTIFYICONDATAW)]
            shell32.Shell_NotifyIconW.restype = ctypes.c_int
            hinst = kernel32.GetModuleHandleW(None)
            class_name = "LitSearchProServerTrayWindow"
            def wndproc(hwnd, msg, wparam, lparam):
                if msg == WM_USER + 20:
                    event = int(lparam)
                    if event == WM_LBUTTONDBLCLK:
                        self.tray_events.put(("show",))
                    elif event == WM_RBUTTONUP:
                        pt = POINT(); user32.GetCursorPos(ctypes.byref(pt))
                        self.tray_events.put(("menu", pt.x, pt.y))
                    return 0
                return int(user32.DefWindowProcW(hwnd, msg, wparam, lparam) or 0)
            self.native_tray_wndproc = WNDPROC(wndproc)
            wc = WNDCLASSW()
            wc.lpfnWndProc = self.native_tray_wndproc
            wc.hInstance = hinst
            wc.lpszClassName = class_name
            try:user32.RegisterClassW(ctypes.byref(wc))
            except Exception:pass
            hwnd = user32.CreateWindowExW(0, class_name, "LitSearchProTray", 0, 0, 0, 0, 0, None, None, hinst, None)
            self.native_tray_hwnd = hwnd
            icon_path = os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(__file__)), "generic_logo.ico")
            if os.path.exists(icon_path):
                hicon = user32.LoadImageW(None, icon_path, 1, 0, 0, 0x00000010)
            else:
                hicon = user32.LoadIconW(None, 32512)
            nid = NOTIFYICONDATAW()
            nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
            nid.hWnd = hwnd
            nid.uID = 1001
            nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE
            nid.uCallbackMessage = WM_USER + 20
            nid.hIcon = hicon
            nid.szTip = "LitSearchPro 协作服务器正在运行"
            ok = shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid))
            if ok:
                self.native_tray = nid
                threading.Thread(target=self.native_tray_message_loop, daemon=True).start()
                self.log("Windows 原生托盘图标已启用")
        except Exception as exc:
            self.log("托盘图标启用失败："+str(exc))

    def native_tray_message_loop(self):
        try:
            msg = MSG()
            user32 = ctypes.windll.user32
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        except Exception:
            pass

    def quit_app(self):
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except Exception:
            pass
        try:
            if self.native_tray:
                ctypes.windll.shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self.native_tray))
                self.native_tray = None
            if self.native_tray_hwnd:
                ctypes.windll.user32.DestroyWindow(self.native_tray_hwnd)
                self.native_tray_hwnd = None
        except Exception:
            pass
        self.stop_server()
        self.root.destroy()

    def log(self, text):
        def append():
            self.log_text.insert(tk.END, f"{now()}  {text}\n")
            self.log_text.see(tk.END)
        try:
            self.root.after(0, append)
        except Exception:
            pass

    def refresh_all(self):
        ips = local_ipv4_addresses()
        suggested = self.public_url.get().strip() or f"http://{ips[0]}:{self.port.get()}"
        try:
            teams=[x["team_name"] for x in self.db.execute("SELECT DISTINCT team_name FROM users WHERE team_name<>'' ORDER BY team_name")]
            if hasattr(self, "group_combo"):
                self.group_combo.configure(values=teams)
        except Exception:
            pass
        overview = [
            f"{SERVER_NAME} v{SERVER_VERSION}",
            "",
            f"服务状态：{self.status.get()}",
            f"建议客户端地址：{suggested}",
            "",
            "本机可用 IPv4：",
            *[f"- http://{ip}:{self.port.get()}" for ip in ips],
            "",
            "说明：服务器绑定 0.0.0.0 后会监听本机所有网卡。外网能否访问取决于学校分配 IP、服务器防火墙、路由/NAT 和端口策略；如果学校给的是公网地址，请在客户端填写学校分配的公网 IP 或域名。",
        ]
        self.overview.delete("1.0", tk.END)
        self.overview.insert("1.0", "\n".join(overview))
        self.reload_tree(self.user_tree, "SELECT id,group_code,username,display_name,role,team_name,active,created_at FROM users ORDER BY active,id DESC", lambda x: (x["id"], x["team_name"], x["group_code"], x["username"], x["display_name"], x["role"], "启用" if x["active"] else "待批准/停用", x["created_at"]))
        self.reload_tree(self.file_tree, "SELECT id,uploader,name,size,created_at FROM files ORDER BY id DESC", lambda x: (x["id"], x["uploader"], x["name"], x["size"], x["created_at"]))
        self.reload_tree(self.leave_tree, "SELECT l.*,COALESCE(NULLIF(u.display_name,''),l.requester) AS requester_name FROM leave_requests l LEFT JOIN users u ON u.group_code=l.group_code AND u.username=l.requester ORDER BY l.id DESC", lambda x: (x["id"], x["requester_name"], x["leave_type"], f"{x['start_time']} - {x['end_time']}", x["status"], x["reason"]))
        self.reload_tree(self.attendance_tree, "SELECT a.id,a.username,COALESCE(NULLIF(u.display_name,''),a.username) AS display_name,a.action,a.ip_address,a.created_at FROM attendance a LEFT JOIN users u ON u.group_code=a.group_code AND u.username=a.username ORDER BY a.id DESC", lambda x: (x["id"], x["display_name"], x["action"], x["ip_address"], x["created_at"]))

    def reload_tree(self, tree, sql, formatter):
        tree.delete(*tree.get_children())
        for row in self.db.execute(sql):
            values = formatter(row)
            tree.insert("", tk.END, iid=str(values[0]), values=values)

    def selected_id(self, tree):
        sel = tree.selection()
        return sel[0] if sel else None

    def set_user_active(self, active):
        uid = self.selected_id(self.user_tree)
        if not uid:
            return
        self.db.execute("UPDATE users SET active=? WHERE id=?", (active, uid))
        self.db.commit()
        self.refresh_all()

    def add_user_dialog(self):
        username = simpledialog.askstring("新增账号", "用户名：", parent=self.root)
        if not username:
            return
        display_name = simpledialog.askstring("新增账号", "姓名：", initialvalue=username, parent=self.root) or username
        group_code = simpledialog.askstring("新增账号", "导师姓名（学生归属导师；导师账号填本人姓名）：", initialvalue=self.group_code.get(), parent=self.root) or self.group_code.get()
        role = simpledialog.askstring("新增账号", "角色（学生/导师/超级管理员/合作者）：", initialvalue="学生", parent=self.root) or "学生"
        team_name = simpledialog.askstring("新增账号", "团队代码（导师/超级管理员填写；学生可空，将继承导师团队）：", parent=self.root) or ""
        password = simpledialog.askstring("新增账号", "初始密码：", show="*", parent=self.root) or "123456"
        if role == "导师" and not group_code:
            group_code = display_name
        if role != "导师" and not team_name:
            teacher = self.db.execute("SELECT team_name FROM users WHERE group_code=? AND role='导师' ORDER BY active DESC,id LIMIT 1",(group_code,)).fetchone()
            if teacher:team_name = teacher["team_name"]
        salt, digest = hash_password(password)
        self.db.execute("INSERT INTO users(group_code,username,display_name,role,team_name,salt,password_hash,active,created_at) VALUES(?,?,?,?,?,?,?,?,?)", (group_code, username, display_name, role, team_name, salt, digest, 1, now()))
        self.db.commit()
        self.refresh_all()

    def set_leave_status(self, status):
        lid = self.selected_id(self.leave_tree)
        if not lid:
            return
        self.db.execute("UPDATE leave_requests SET status=?,approved_at=? WHERE id=?", (status, now(), lid))
        self.db.commit()
        self.refresh_all()


def main():
    install_global_crash_handlers()
    parser = argparse.ArgumentParser(description=SERVER_NAME)
    parser.add_argument("--host", default=None)
    parser.add_argument("--port", type=int, default=None)
    args = parser.parse_args()
    cfg = load_config()
    if args.host:
        cfg["host"] = args.host
    if args.port:
        cfg["port"] = args.port
    save_config(cfg)
    root = None
    try:
        root = tk.Tk()
        ServerApp(root)
        root.mainloop()
    except Exception as exc:
        write_server_log("server_crash.log", "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        try:
            if root:
                messagebox.showerror("服务器管理端异常", f"服务器管理端遇到异常，但已记录日志。\n\n日志位置：{os.path.join(APP_DIR, 'server_crash.log')}\n\n{exc}", parent=root)
            else:
                messagebox.showerror("服务器管理端异常", f"服务器管理端启动失败，已记录日志。\n\n日志位置：{os.path.join(APP_DIR, 'server_crash.log')}\n\n{exc}")
        except Exception:
            pass


if __name__ == "__main__":
    main()


