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
import sys
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


# Kesit çizimi için her şeklin ihtiyaç duyduğu ölçü sütunları
SHAPE_DIMS = {
    "I": ["h", "b", "tw", "tf"],
    "U": ["h", "b", "tw", "tf"],
    "CHS": ["d", "t"],
    "SHS": ["a", "t"],
    "RHS": ["h", "b", "t"],
    "L": ["a", "t"],
    "LU": ["a", "b", "t"],
    "FLAT": ["b", "t"],
}


def parse_header(text):
    m = re.match(r"\s*([^\[]+?)\s*\[(.+)\]\s*$", str(text or ""))
    return (m.group(1), m.group(2)) if m else (str(text or "").strip(), "")


def is_pos(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0


def load_data():
    """Excel'i okur. Sütunlar sayfadan sayfaya farklı sırada olabilir;
    hepsi başlık adına göre bulunur. Hata varsa liste olarak döner."""
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    families, errors = [], []
    for ws in wb:
        rows = list(ws.iter_rows(values_only=True))
        if ws.title not in FAMILY_INFO:
            errors.append(f"'{ws.title}' sayfası tanımsız (build.py içindeki FAMILY_INFO'ya ekle)")
            continue
        label, desc, shape = FAMILY_INFO[ws.title]
        cols = [parse_header(h) for h in rows[0][1:]]
        keys = [k for k, _ in cols]
        if "G" not in keys:
            errors.append(f"{ws.title}: 'G [kg/m]' sütunu bulunamadı")
            continue
        missing = [k for k in SHAPE_DIMS[shape] if k not in keys]
        if missing:
            errors.append(f"{ws.title}: ölçü sütunu eksik: {', '.join(missing)}")
            continue
        g = keys.index("G") + 1
        need = [g] + [keys.index(k) + 1 for k in SHAPE_DIMS[shape]]
        items = [list(r) for r in rows[1:] if r[0]]
        for r in items:
            bad = [rows[0][i] for i in need if not is_pos(r[i] if i < len(r) else None)]
            if bad:
                errors.append(f"{ws.title} / {r[0]}: geçersiz değer: {', '.join(map(str, bad))}")
        families.append({
            "id": ws.title,
            "label": label,
            "desc": desc,
            "shape": shape,
            "g": g,
            "cols": [{"k": k, "u": u} for k, u in cols],
            "rows": items,
        })
    return families, errors


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
    families, errors = load_data()
    if errors:
        print("Excel'de sorun var, uygulama üretilmedi:")
        for e in errors:
            print("  -", e)
        sys.exit(1)
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
