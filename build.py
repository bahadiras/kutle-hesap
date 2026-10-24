"""Kütle Hesap uygulamasını derler.

profil_bilgileri.xlsx -> JSON verisi -> src/template.html içine gömülür.
Çıktılar:
  docs/index.html         (GitHub Pages ile yayınlanan PWA)
  docs/manifest.webmanifest, docs/sw.js, docs/icons/*
  KutleHesap.html         (internetsiz açılabilen tek dosyalık yedek)

Kullanım:  python build.py
"""
import base64
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


ICON_TOP = (28, 128, 152)     # gradyan üst rengi
ICON_BOTTOM = (10, 52, 70)    # gradyan alt rengi


def icon_mark(S, scale=1.0):
    """Ağırlık (kütle) silueti ve içinde I-profil boşluğu: beyaz alan maskesi.
    Koordinatlar 0..1 birim karede tanımlı; scale, merkeze göre küçültür."""
    def P(x, y):
        return (S * (0.5 + (x - 0.5) * scale), S * (0.5 + (y - 0.5) * scale))

    def L(v):
        return S * v * scale

    m = Image.new("L", (S, S), 0)
    d = ImageDraw.Draw(m)
    # Gövde: köşeleri yuvarlatılmış yamuk (çekirdek + r yarıçaplı genişletme)
    r = 0.05
    core = [P(0.5 - 0.185, 0.45), P(0.5 + 0.185, 0.45), P(0.5 + 0.27, 0.75), P(0.5 - 0.27, 0.75)]
    d.polygon(core, fill=255)
    for i, a in enumerate(core):
        b = core[(i + 1) % len(core)]
        d.line([a, b], fill=255, width=int(L(2 * r)))
        d.ellipse([a[0] - L(r), a[1] - L(r), a[0] + L(r), a[1] + L(r)], fill=255)
    # Tutamak halkası
    cx, cy = P(0.5, 0.315)
    ro, ri = L(0.115), L(0.06)
    d.ellipse([cx - ro, cy - ro, cx + ro, cy + ro], fill=255)
    d.ellipse([cx - ri, cy - ri, cx + ri, cy + ri], fill=0)
    # I-profil boşluğu
    x0, x1, y0, y1 = 0.39, 0.61, 0.505, 0.705
    fl, web = 0.042, 0.048
    d.rectangle([P(x0, y0), P(x1, y0 + fl)], fill=0)
    d.rectangle([P(x0, y1 - fl), P(x1, y1)], fill=0)
    d.rectangle([P(0.5 - web / 2, y0), P(0.5 + web / 2, y1)], fill=0)
    return m


def make_icon(size, path, maskable=False):
    S = 2048  # büyük çizip küçültmek kenarları yumuşatır
    grad = Image.new("RGB", (S, S))
    gd = ImageDraw.Draw(grad)
    for y in range(S):
        t = y / (S - 1)
        gd.line([(0, y), (S, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(ICON_TOP, ICON_BOTTOM)))
    shape = Image.new("L", (S, S), 0)
    if maskable:
        shape.paste(255, (0, 0, S, S))
    else:
        ImageDraw.Draw(shape).rounded_rectangle([0, 0, S - 1, S - 1], radius=S * 0.225, fill=255)
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    img.paste(grad, (0, 0), shape)
    img.paste((255, 255, 255), (0, 0), icon_mark(S, 0.78 if maskable else 1.0))
    img.resize((size, size), Image.LANCZOS).save(path)


def main():
    families, errors = load_data()
    if errors:
        print("Excel'de sorun var, uygulama üretilmedi:")
        for e in errors:
            print("  -", e)
        sys.exit(1)
    data_json = json.dumps(families, ensure_ascii=False, separators=(",", ":"))
    version = datetime.now().strftime("%Y.%m.%d-%H%M%S")

    # Başlık logosu ve sekme ikonu: yedek HTML tek başına da çalışsın diye sayfaya gömülür
    APP.mkdir(exist_ok=True)
    (APP / "icons").mkdir(exist_ok=True)
    make_icon(96, APP / "icons" / "icon-96.png")
    icon_data = "data:image/png;base64," + base64.b64encode((APP / "icons" / "icon-96.png").read_bytes()).decode()

    html = (SRC / "template.html").read_text(encoding="utf-8")
    html = (html.replace("/*__DATA__*/null", data_json).replace("__VERSION__", version)
                .replace("__ICON_DATA__", icon_data))

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
