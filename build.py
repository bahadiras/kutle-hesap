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


# İkon renkleri: zemin ve I-profil için üstten alta gradyan
ICON_PALETTES = {
    "grafit":  {"bg": ((40, 46, 54), (15, 18, 23)), "mark": ((242, 244, 247), (160, 168, 179))},
    "lacivert": {"bg": ((22, 38, 66), (8, 15, 30)), "mark": ((238, 211, 150), (172, 134, 70))},
    "beyaz":   {"bg": ((252, 252, 250), (228, 230, 234)), "mark": ((52, 57, 64), (18, 20, 24))},
}
ICON_PALETTE = "lacivert"


def vgradient(S, top, bottom, y0=0, y1=None):
    """S×S boyutunda dikey gradyan; renk y0..y1 aralığında geçiş yapar."""
    y1 = S if y1 is None else y1
    img = Image.new("RGB", (S, S))
    d = ImageDraw.Draw(img)
    for y in range(S):
        t = min(1, max(0, (y - y0) / max(1, y1 - y0)))
        d.line([(0, y), (S, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(top, bottom)))
    return img


def icon_mark(S, scale=1.0):
    """Köşe yuvarlatmalı (fillet) I-profil kesiti maskesi.
    Koordinatlar 0..1 birim karede; scale, merkeze göre küçültür."""
    def X(v):
        return S * (0.5 + (v - 0.5) * scale)

    def L(v):
        return S * v * scale

    x0, x1, y0, y1 = 0.27, 0.73, 0.25, 0.75   # dış ölçüler
    tf, tw, rf, re = 0.078, 0.092, 0.034, 0.012  # başlık, gövde, iç köşe, uç yuvarlatma
    m = Image.new("L", (S, S), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle([X(x0), X(y0), X(x1), X(y0 + tf)], radius=L(re), fill=255)
    d.rounded_rectangle([X(x0), X(y1 - tf), X(x1), X(y1)], radius=L(re), fill=255)
    wl, wr = 0.5 - tw / 2, 0.5 + tw / 2
    d.rectangle([X(wl), X(y0 + tf / 2), X(wr), X(y1 - tf / 2)], fill=255)
    # Gövde-başlık birleşimlerine iç köşe yuvarlatması
    for ex, sx in ((wl, -1), (wr, 1)):
        for ey, sy in ((y0 + tf, 1), (y1 - tf, -1)):
            ax, ay = ex + sx * rf, ey + sy * rf
            d.rectangle([X(min(ex, ax)), X(min(ey, ay)), X(max(ex, ax)), X(max(ey, ay))], fill=255)
            d.ellipse([X(ax) - L(rf), X(ay) - L(rf), X(ax) + L(rf), X(ay) + L(rf)], fill=0)
    return m, X(y0), X(y1)


def make_icon(size, path, maskable=False, palette=None):
    pal = ICON_PALETTES[palette or ICON_PALETTE]
    S = 2048  # büyük çizip küçültmek kenarları yumuşatır
    shape = Image.new("L", (S, S), 0)
    if maskable:
        shape.paste(255, (0, 0, S, S))
    else:
        ImageDraw.Draw(shape).rounded_rectangle([0, 0, S - 1, S - 1], radius=S * 0.225, fill=255)
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    img.paste(vgradient(S, *pal["bg"]), (0, 0), shape)
    if not maskable:
        # İnce iç kenar parlaklığı
        edge = Image.new("L", (S, S), 0)
        ImageDraw.Draw(edge).rounded_rectangle([4, 4, S - 5, S - 5], radius=S * 0.222, outline=34, width=int(S * 0.006))
        img.paste((255, 255, 255), (0, 0), edge)
    mark, my0, my1 = icon_mark(S, 0.8 if maskable else 1.0)
    img.paste(vgradient(S, *pal["mark"], my0, my1), (0, 0), mark)
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
