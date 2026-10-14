"""Prevod mesačných Excelov so smenovkou na dáta pre web.

Spustenie:
    python sync.py          sleduje priečinok s Excelmi a po každom uložení aktualizuje web
    python sync.py --once   spracuje všetko raz a skončí
"""
import datetime
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
from pathlib import Path

from openpyxl import load_workbook

# ------------------------- NASTAVENIA -------------------------
# Sem vlož svoju presnú cestu k priečinku s Excelmi:
EXCEL_FOLDER = Path(r"C:\Users\Gastro\Desktop\PLANOVANIE A ODMENY NEW")
PUSH = True                               # True = po zmene sám nahrá web na GitHub
PUSH_EVERY_MINUTES = 60                   # zmeny sa na web nahrávajú najviac raz za toto obdobie
POLL_SECONDS = 20                         # ako často kontroluje priečinok
# --------------------------------------------------------------

BASE = Path(__file__).resolve().parent
SITE = BASE / "docs"
DATA = SITE / "data"
FILE_RE = re.compile(r"^Master_Smenovka_(\d{4})_(\d{2})\.xlsx$", re.I)
MONTHS = ["Január", "Február", "Marec", "Apríl", "Máj", "Jún", "Júl",
          "August", "September", "Október", "November", "December"]


def plain(v):
    """Text bez diakritiky, veľkými písmenami (na porovnávanie)."""
    s = unicodedata.normalize("NFKD", str(v or ""))
    return "".join(c for c in s if not unicodedata.combining(c)).strip().upper()


def find_sheet(wb, *words):
    for ws in wb.worksheets:
        if all(w in plain(ws.title) for w in words):
            return ws
    raise ValueError("V súbore chýba list: " + " ".join(words))


def read_schedule(ws, year, month):
    cols = {}
    for c in range(3, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, datetime.datetime) and v.year == year and v.month == month:
            cols[v.day] = c
    days = sorted(cols)
    shifts = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name or plain(ws.cell(r, 2).value) != "SMENA":
            continue
        s = {}
        for d in days:
            v = ws.cell(r, cols[d]).value
            if v not in (None, ""):
                s[str(d)] = str(v).strip()
        # pracovníci, ktorí nemajú v mesiaci žiadnu smenu (len "x" alebo prázdno), sa vynechajú
        if any(not code.upper().startswith("X") for code in s.values()):
            shifts.append({"n": str(name).strip(), "s": s})
    return days, shifts


def read_rewards(ws, key):
    """Nájde blok odmien podľa nadpisu v stĺpci B (napr. VNUTORNA OBSLUHA)."""
    marker = next((r for r in range(1, ws.max_row + 1)
                   if plain(ws.cell(r, 2).value).startswith(key)), None)
    if marker is None:
        return []
    cat_row = marker + 3
    cats = [(str(ws.cell(cat_row, c).value).strip(), c)
            for c in range(8, ws.max_column + 1) if ws.cell(cat_row, c).value]
    out = []
    r = marker + 5
    while r <= ws.max_row and ws.cell(r, 2).value:
        name = str(ws.cell(r, 2).value).strip()
        num = lambda row, col: round(float(ws.cell(row, col).value or 0), 4)
        total, fulfil = num(r, 5), num(r, 4)
        if total or fulfil:                       # brigádnici s nulami sa nezobrazia
            rank = ws.cell(r, 3).value
            out.append({
                "n": name,
                "rank": rank if isinstance(rank, (int, float)) else None,
                "tot": round(total, 2), "pl": fulfil,
                "c": [[cn.title() if cn.isupper() else cn,
                       num(r, c), num(r, c + 1), num(r, c + 2),
                       round(num(r, c + 3), 2), round(num(r, c + 4), 2)]
                      for cn, c in cats],
            })
        r += 1
    out.sort(key=lambda x: (x["rank"] is None, x["rank"] or 0))
    return out


def read_month(path, year, month):
    with tempfile.TemporaryDirectory() as tmp:
        copy = Path(tmp) / "x.xlsx"
        shutil.copy2(path, copy)
        wb = load_workbook(copy, data_only=True)
    days, shifts = read_schedule(find_sheet(wb, "MASTER"), year, month)
    su = find_sheet(wb, "SUMAR")
    return {
        "id": f"{year}-{month:02d}", "label": f"{MONTHS[month - 1]} {year}",
        "y": year, "m": month - 1, "days": days, "shifts": shifts,
        "vno": read_rewards(su, "VNUTORNA"), "vo": read_rewards(su, "VONKAJSIA"),
    }