"""Kütle Hesap uygulamasını derler.

profil_bilgileri.xlsx -> JSON verisi -> src/template.html içine gömülür.
Çıktılar:
  docs/index.html         (GitHub Pages ile yayınlanan PWA)
  docs/manifest.webmanifest, docs/sw.js, docs/icons/*
  KutleHesap.html         (internetsiz açılabilen tek dosyalık yedek)

Kullanım:  python build.py
"""
import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
from PIL import Image, ImageDraw

ROOT = Path(__file__).parent
SRC = ROOT / "src"
APP = ROOT / "docs"
XLSX = ROOT / "profil_bilgileri.xlsx"

FAMILY_INFO = {
    "IPE": ("IPE", "I profil", "I"),
    "IPN": ("IPN", "I profil", "I"),
    "HEA": ("HEA", "H profil", "I"),
    "HEB": ("HEB", "H profil", "I"),
    "HEM": ("HEM", "H profil", "I"),
    "UPN": ("UPN", "U profil", "U"),
    "UPE": ("UPE", "U profil", "U"),
    "CHS": ("Boru", "Dairesel boru", "CHS"),
    "SHS": ("Kare kutu", "Kare kutu profil", "SHS"),
    "RHS": ("Dikdörtgen kutu", "Dikdörtgen kutu profil", "RHS"),
    "L": ("Köşebent", "Eşit kollu köşebent", "L"),
    "LU": ("Farklı köşebent", "Eşit olmayan köşebent", "LU"),
    "FLAT": ("Lama", "Lama (yassı)", "FLAT"),
}


def parse_header(text):
    m = re.match(r"\s*([^\[]+?)\s*\[(.+)\]\s*$", text)
    return (m.group(1), m.group(2)) if m else (text.strip(), "")


def load_data():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    families = []
    for ws in wb:
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        cols = [parse_header(h) for h in header[1:]]
        items = [list(r) for r in rows[1:] if r[0]]
        label, desc, shape = FAMILY_INFO.get(ws.title, (ws.title, ws.title, "FLAT"))
        families.append({
            "id": ws.title,
            "label": label,
            "desc": desc,
            "shape": shape,
            "cols": [{"k": k, "u": u} for k, u in cols],
            "rows": items,
        })
    return families


def make_icon(size, path, maskable=False):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    bg = (21, 94, 117, 255)
    if maskable:
        d.rectangle([0, 0, size, size], fill=bg)
        pad = size * 0.22
    else:
        d.rounded_rectangle([0, 0, size - 1, size - 1], radius=size * 0.22, fill=bg)
        pad = size * 0.2
    # I kesiti
    w = size - 2 * pad
    fl = w * 0.16
    web = w * 0.16
    x0, y0 = pad, pad
    fg = (255, 255, 255, 255)
    d.rectangle([x0, y0, x0 + w, y0 + fl], fill=fg)
    d.rectangle([x0, y0 + w - fl, x0 + w, y0 + w], fill=fg)
    d.rectangle([x0 + w / 2 - web / 2, y0, x0 + w / 2 + web / 2, y0 + w], fill=fg)
    img.save(path)


def main():
    families = load_data()
    data_json = json.dumps(families, ensure_ascii=False, separators=(",", ":"))
    version = datetime.now().strftime("%Y.%m.%d-%H%M%S")

    html = (SRC / "template.html").read_text(encoding="utf-8")
    html = html.replace("/*__DATA__*/null", data_json).replace("__VERSION__", version)

    APP.mkdir(exist_ok=True)
    (APP / "icons").mkdir(exist_ok=True)
    (APP / "index.html").write_text(html, encoding="utf-8")
    (ROOT / "KutleHesap.html").write_text(html, encoding="utf-8")

    sw = (SRC / "sw.js").read_text(encoding="utf-8").replace("__VERSION__", version)
    (APP / "sw.js").write_text(sw, encoding="utf-8")
    shutil.copy(SRC / "manifest.webmanifest", APP / "manifest.webmanifest")

    make_icon(192, APP / "icons" / "icon-192.png")
    make_icon(512, APP / "icons" / "icon-512.png")
    make_icon(512, APP / "icons" / "icon-maskable-512.png", maskable=True)

    total = sum(len(f["rows"]) for f in families)
    print(f"Sürüm {version}: {len(families)} aile, {total} profil -> docs/ ve KutleHesap.html")


if __name__ == "__main__":
    main()
