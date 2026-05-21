import base64
import io
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image

from layout_spec import (
    A4_HEIGHT_MM,
    A4_WIDTH_MM,
    CARD_GAP_X_MM,
    CARD_GAP_Y_MM,
    CARD_HEIGHT_MM,
    CARD_WIDTH_MM,
    CARDS_PER_PAGE,
    PAGE_MARGIN_BOTTOM_MM,
    PAGE_MARGIN_LEFT_MM,
    PAGE_MARGIN_RIGHT_MM,
    PAGE_MARGIN_TOP_MM,
    PHOTO_HEIGHT_MM,
    PHOTO_WIDTH_MM,
    page_content_height_mm,
    page_content_width_mm,
)

_DIR = Path(__file__).resolve().parent
_ASSETS = _DIR / "assets"


def _read_text(path: Path, default: str = "") -> str:
    try:
        return path.read_text(encoding="utf-8").strip()
    except Exception:
        return default


def _read_b64(path: Path, mime: str) -> str:
    try:
        return f"data:{mime};base64," + base64.b64encode(path.read_bytes()).decode()
    except Exception:
        return ""


def _chrome_bin():
    candidates = [
        os.environ.get("GOOGLE_CHROME_BIN", ""),
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        shutil.which("google-chrome") or "",
        shutil.which("chromium") or "",
        shutil.which("chromium-browser") or "",
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def _esc(text: str) -> str:
    return (
        str(text or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _display_value(text: str) -> str:
    value = str(text or "").strip()
    return f"　{_esc(value)}" if value else ""


def _img_src(value: str, mime_hint: str = "image/jpeg") -> str:
    if not value:
        return ""
    if value.startswith("data:"):
        return value
    return f"data:{mime_hint};base64,{value}"


def _load_logo_b64():
    return _read_b64(_ASSETS / "logo_live.svg", "image/svg+xml")


def _load_star_b64():
    raw = _read_text(_ASSETS / "star_gold_b64.txt")
    if not raw:
        return ""
    try:
        data = base64.b64decode(raw)
        img = Image.open(io.BytesIO(data)).convert("RGBA")
        pixels = img.load()
        width, height = img.size
        for y in range(height):
            for x in range(width):
                r, g, b, a = pixels[x, y]
                if r > 245 and g > 245 and b > 245:
                    pixels[x, y] = (r, g, b, 0)
        out = io.BytesIO()
        img.save(out, format="PNG")
        return "data:image/png;base64," + base64.b64encode(out.getvalue()).decode()
    except Exception:
        return f"data:image/png;base64,{raw}"


def _load_center_bg_b64():
    return _read_b64(_ASSETS / "card_center_bg.png", "image/png")


_CSS = """
html, body {{
  margin: 0;
  padding: 0;
  background: white;
  font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
}}
.sheet {{
  width: __CONTENT_WIDTH__mm;
  height: __CONTENT_HEIGHT__mm;
  margin: 0 auto;
  display: grid;
  grid-template-columns: __CARD_WIDTH__mm __CARD_WIDTH__mm;
  gap: __GAP_Y__mm __GAP_X__mm;
  break-inside: avoid;
  page-break-inside: avoid;
}}
.sheet + .sheet {{
  break-before: page;
  page-break-before: always;
}}
.badge-card {{
  width: __CARD_WIDTH__mm;
  height: __CARD_HEIGHT__mm;
  background: #fff;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  position: relative;
  flex-shrink: 0;
  border: 0.3mm solid #e0e0e0;
  box-sizing: border-box;
}}
.badge-card-bg {{
  position: absolute;
  left: 50%;
  top: 50%;
  width: 45.5mm;
  height: auto;
  transform: translate(calc(-50% - 5px), calc(-50% + 10px));
  opacity: 0.56;
  pointer-events: none;
  z-index: 0;
}}
.bh {{
  background: #E50036;
  height: 11.3mm;
  margin-top: -5px;
  display: flex;
  align-items: center;
  padding: 0 2mm 0 0;
  gap: 2mm;
  flex-shrink: 0;
  position: relative;
  overflow: visible;
  z-index: 2;
}}
.bh-role {{
  display: flex;
  flex-direction: column;
  align-items: center;
  flex-shrink: 0;
  align-self: flex-start;
  width: 10.8mm;
  border-radius: 2mm 2mm 0 0;
  margin-left: 12px;
  margin-top: 10px;
  overflow: visible;
}}
.bh-role-top {{
  width: 10.8mm;
  height: 8.6mm;
  background: #fff;
  border-radius: 2mm 2mm 0 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 0.8mm;
  flex-shrink: 0;
}}
.bh-role-bottom {{
  width: 10.8mm;
  background: #E50036;
  border-radius: 0 0 2mm 2mm;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0.8mm 0.5mm;
}}
.bh-role-crown svg {{
  width: 4mm;
  height: 2.8mm;
  display: block;
}}
.bh-role-cn {{
  font-size: 3.5mm;
  font-weight: 900;
  color: #E50036;
  line-height: 1;
  text-align: center;
  letter-spacing: 0.5mm;
}}
.bh-role-en {{
  font-size: 1.1mm;
  font-weight: 700;
  color: #fff;
  line-height: 1;
  text-align: center;
  white-space: nowrap;
  letter-spacing: 0.1mm;
}}
.bh-logo {{
  position: absolute;
  left: 14.2mm;
  top: 10px;
  height: 6.55mm;
  width: auto;
  object-fit: contain;
  z-index: 200;
}}
.bh-deco {{
  position: absolute;
  right: 0;
  top: 0;
  height: 11.3mm;
  width: 24mm;
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 2mm;
  overflow: visible;
  pointer-events: none;
}}
.bh-deco-circle {{
  width: calc(14.5mm + 58px);
  height: calc(17mm - 5px);
  background: #fff;
  border-radius: 999px;
  flex-shrink: 0;
  transform: translate(calc(4mm + 18px), 3px);
  position: relative;
  z-index: 10;
}}
.bh-deco-circle::before {{
  content: '';
  position: absolute;
  inset: calc(3.6mm - 3px);
  background: #E50036;
  border-radius: inherit;
}}
.bh-deco-circle::after {{
  content: '';
  position: absolute;
  top: 0;
  bottom: 0;
  left: 50%;
  transform: translateX(-50%);
  width: 11px;
  background: #fff;
}}
.bh-deco-pill {{
  width: 5.5mm;
  height: 17mm;
  background: #fff;
  border-radius: 3mm;
  flex-shrink: 0;
  transform: translateX(4mm);
  position: relative;
}}
.bh-deco-pill::before {{
  content: '';
  position: absolute;
  inset: 3.6mm;
  background: #E50036;
  border-radius: calc(3mm - 1mm);
}}
.bb {{
  flex: 1;
  display: grid;
  grid-template-columns: __PHOTO_WIDTH__mm 1fr 20.2mm;
  grid-template-rows: 11.2mm 1fr;
  padding: 2.8mm 2.1mm 2.2mm 2.1mm;
  column-gap: 1.6mm;
  row-gap: 1.2mm;
  background: transparent;
  min-height: 0;
  position: relative;
  z-index: 2;
}}
.bb-photo {{
  grid-column: 1;
  grid-row: 1 / 3;
  padding-top: calc(2.15mm + 3px);
  padding-left: 4px;
  display: flex;
  flex-direction: column;
}}
.bb-photo-frame {{
  width: __PHOTO_WIDTH__mm;
  height: __PHOTO_HEIGHT__mm;
  flex-shrink: 0;
  border: 0.5mm dashed #999;
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
}}
.bb-photo-frame img {{
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  object-fit: cover;
  display: block;
}}
.bb-photo-frame svg.placeholder-x {{
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
}}
.bb-photo-frame .ph-text {{
  position: absolute;
  font-size: 2.8mm;
  color: #aaa;
  font-weight: 500;
  writing-mode: vertical-rl;
  text-orientation: mixed;
  letter-spacing: 1mm;
}}
.bb-stars {{
  grid-column: 2 / 4;
  grid-row: 1;
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 2px;
}}
.bb-star {{
  width: 9.8mm;
  height: 9.8mm;
  flex-shrink: 0;
  background-image: var(--star-src);
  background-size: contain;
  background-repeat: no-repeat;
  background-position: center;
}}
.bb-fields {{
  grid-column: 2;
  grid-row: 2;
  display: flex;
  flex-direction: column;
  gap: 0;
  align-self: end;
  padding-right: 0;
  margin-left: 4px;
  width: calc(100% - 8px);
}}
.bb-field {{
  display: flex;
  align-items: flex-end;
  margin-bottom: 1.8mm;
  gap: 0;
}}
.bb-field-label {{
  font-size: 3.65mm;
  font-weight: 700;
  color: #1a1a1a;
  white-space: nowrap;
  flex-shrink: 0;
  line-height: 1.2;
}}
.bb-field-line {{
  flex: 1;
  border-bottom: 1px solid #1a1a1a;
  min-width: 6mm;
  margin-left: 0.7mm;
  padding-bottom: 0.35mm;
  font-size: 3.45mm;
  font-weight: 600;
  color: #1a1a1a;
  line-height: 1.2;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}}
.bb-field:nth-child(2) .bb-field-line {{
  border-bottom-width: 0.5px;
}}
.bb-qr {{
  grid-column: 3;
  grid-row: 2;
  align-self: end;
}}
.bb-qr-frame {{
  width: 22mm;
  height: 22mm;
  transform: translate(-4px, 3px);
  border: 0.5mm dashed #999;
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
}}
.bb-qr-frame img {{
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  object-fit: contain;
  display: block;
  image-rendering: pixelated;
}}
.bb-qr-frame svg.placeholder-x {{
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
}}
.bb-qr-frame .ph-text {{
  position: absolute;
  font-size: 2.8mm;
  color: #aaa;
  font-weight: 500;
}}
.bb-qr-label {{
  position: absolute;
  bottom: 1.2mm;
  left: 0;
  right: 0;
  text-align: center;
  font-size: 2.2mm;
  color: #888;
  pointer-events: none;
}}
.bt {{
  height: 2.9mm;
  background: #E60036;
  flex-shrink: 0;
}}
@page {{
  size: __A4_WIDTH__mm __A4_HEIGHT__mm;
  margin: __PAGE_TOP__mm __PAGE_RIGHT__mm __PAGE_BOTTOM__mm __PAGE_LEFT__mm;
}}
* {{
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
  color-adjust: exact !important;
  box-sizing: border-box;
}}
""".replace("__CONTENT_WIDTH__", f"{page_content_width_mm():.3f}") \
    .replace("__CONTENT_HEIGHT__", f"{page_content_height_mm():.3f}") \
    .replace("__CARD_WIDTH__", f"{CARD_WIDTH_MM:.3f}") \
    .replace("__CARD_HEIGHT__", f"{CARD_HEIGHT_MM:.3f}") \
    .replace("__GAP_X__", f"{CARD_GAP_X_MM:.3f}") \
    .replace("__GAP_Y__", f"{CARD_GAP_Y_MM:.3f}") \
    .replace("__PHOTO_WIDTH__", f"{PHOTO_WIDTH_MM:.3f}") \
    .replace("__PHOTO_HEIGHT__", f"{PHOTO_HEIGHT_MM:.3f}") \
    .replace("__A4_WIDTH__", f"{A4_WIDTH_MM:.3f}") \
    .replace("__A4_HEIGHT__", f"{A4_HEIGHT_MM:.3f}") \
    .replace("__PAGE_TOP__", f"{PAGE_MARGIN_TOP_MM:.3f}") \
    .replace("__PAGE_RIGHT__", f"{PAGE_MARGIN_RIGHT_MM:.3f}") \
    .replace("__PAGE_BOTTOM__", f"{PAGE_MARGIN_BOTTOM_MM:.3f}") \
    .replace("__PAGE_LEFT__", f"{PAGE_MARGIN_LEFT_MM:.3f}") \
    .replace("{{", "{") \
    .replace("}}", "}")


def _build_card(card: dict, logo_src: str, star_src: str, center_bg_src: str) -> str:
    photo_src = _img_src(card.get("photo_b64", ""), "image/jpeg")
    qr_src = _img_src(card.get("qr_b64", ""), "image/png")

    if photo_src:
        photo_html = f'<img src="{photo_src}" alt="照片">'
        photo_style = ' style="border-color:transparent"'
    else:
        photo_html = """
<svg class="placeholder-x" viewBox="0 0 100 100" preserveAspectRatio="none">
  <line x1="0" y1="0" x2="100" y2="100" stroke="#ccc" stroke-width="1.2"/>
  <line x1="100" y1="0" x2="0" y2="100" stroke="#ccc" stroke-width="1.2"/>
</svg>
<span class="ph-text">一寸照</span>
"""
        photo_style = ""

    if qr_src:
        qr_html = f'<img src="{qr_src}" alt="二维码">'
        qr_style = ' style="border-color:transparent"'
        qr_label = ""
    else:
        qr_html = """
<svg class="placeholder-x" viewBox="0 0 100 100" preserveAspectRatio="none">
  <line x1="0" y1="0" x2="100" y2="100" stroke="#ccc" stroke-width="1.2"/>
  <line x1="100" y1="0" x2="0" y2="100" stroke="#ccc" stroke-width="1.2"/>
</svg>
<span class="ph-text">企微码</span>
"""
        qr_style = ""
        qr_label = '<div class="bb-qr-label">企微码</div>'

    stars = "".join('<span class="bb-star" aria-label="★"></span>' for _ in range(5))

    return f"""
<div class="badge-card">
  <img class="badge-card-bg" src="{center_bg_src}" alt="">
  <div class="bh">
    <div class="bh-role">
      <div class="bh-role-top">
        <div class="bh-role-crown">
          <svg viewBox="0 0 56 40" xmlns="http://www.w3.org/2000/svg" fill="none" stroke="#E60036" stroke-linejoin="round" stroke-linecap="round">
            <rect x="6" y="28" width="44" height="8" rx="2" stroke-width="2.5"/>
            <path d="M6 28 L6 16 L18 24 L28 6 L38 24 L50 16 L50 28 Z" stroke-width="2.5"/>
            <circle cx="28" cy="6" r="3.5" stroke-width="2"/>
            <circle cx="6" cy="16" r="2.5" stroke-width="2"/>
            <circle cx="50" cy="16" r="2.5" stroke-width="2"/>
          </svg>
        </div>
        <div class="bh-role-cn">店长</div>
      </div>
      <div class="bh-role-bottom">
        <div class="bh-role-en">Store Manager</div>
      </div>
    </div>
    <img class="bh-logo" src="{logo_src}" alt="logo">
    <div class="bh-deco">
      <div class="bh-deco-circle"></div>
      <div class="bh-deco-pill"></div>
    </div>
  </div>
  <div class="bb">
    <div class="bb-photo">
      <div class="bb-photo-frame"{photo_style}>{photo_html}</div>
    </div>
    <div class="bb-stars" style="--star-src:url('{star_src}')">{stars}</div>
    <div class="bb-fields">
      <div class="bb-field">
        <span class="bb-field-label">姓名:</span>
        <span class="bb-field-line">{_display_value(card.get("name", ""))}</span>
      </div>
      <div class="bb-field">
        <span class="bb-field-label">职务:</span>
        <span class="bb-field-line">{_display_value(card.get("title", ""))}</span>
      </div>
      <div class="bb-field" style="margin-bottom:0;">
        <span class="bb-field-label">工号:</span>
        <span class="bb-field-line">{_display_value(card.get("staff_id", ""))}</span>
      </div>
    </div>
    <div class="bb-qr">
      <div class="bb-qr-frame"{qr_style}>{qr_html}{qr_label}</div>
    </div>
  </div>
  <div class="bt"></div>
</div>
"""


def build_pdf_html(cards: list[dict]) -> str:
    logo_src = _load_logo_b64()
    star_src = _load_star_b64()
    center_bg_src = _load_center_bg_b64()
    page_chunks = []
    per_page = CARDS_PER_PAGE
    for idx in range(0, len(cards), per_page):
        batch = cards[idx: idx + per_page]
        page_cards = "".join(_build_card(card, logo_src, star_src, center_bg_src) for card in batch)
        blanks = "".join('<div class="badge-card" style="visibility:hidden"></div>' for _ in range(per_page - len(batch)))
        page_chunks.append(f'<section class="sheet">{page_cards}{blanks}</section>')
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>高济药房 · 员工工牌制作系统</title>
<style>{_CSS}</style>
</head>
<body>
{''.join(page_chunks)}
</body>
</html>"""


def build_preview_html(cards: list[dict], scale: float = 0.46) -> str:
    document = build_pdf_html(cards)
    return document.replace(
        "<body>",
        f"""<body style="background:#eef0f4;padding:12px 0 24px;">
<style>
.sheet-wrap {{
  background:#fff;
  border-radius:18px;
  box-shadow:0 12px 28px rgba(15, 23, 42, 0.08);
  padding:18px 14px;
  margin:0 auto 18px;
  width:max-content;
}}
.sheet {{
  transform:scale({scale});
  transform-origin:top left;
  margin:0;
}}
</style>""",
        1,
    ).replace(
        '<section class="sheet">',
        '<div class="sheet-wrap"><section class="sheet">',
    ).replace(
        "</section>",
        "</section></div>",
    )


def export_html_to_pdf(html: str, output_path: str):
    chrome = _chrome_bin()
    if not chrome:
        raise RuntimeError("未找到 Google Chrome，无法使用浏览器版 PDF 导出")

    with tempfile.TemporaryDirectory() as tmpdir:
        html_path = Path(tmpdir) / "badges.html"
        user_data_dir = Path(tmpdir) / "chrome-user-data"
        pdf_path = Path(output_path)
        html_path.write_text(html, encoding="utf-8")
        cmd = [
            chrome,
            "--headless=new",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--disable-extensions",
            "--disable-background-networking",
            f"--user-data-dir={str(user_data_dir)}",
            "--run-all-compositor-stages-before-draw",
            f"--print-to-pdf={str(pdf_path)}",
            "--no-pdf-header-footer",
            html_path.as_uri(),
        ]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("Chrome PDF 导出超时，请减少单次工牌数量后重试") from exc
        if not pdf_path.exists():
            raise RuntimeError("Chrome 已执行，但未生成 PDF 文件")


def export_cards_to_pdf(cards: list[dict], output_path: str):
    export_html_to_pdf(build_pdf_html(cards), output_path)


def export_cards_to_pdf_bytes(cards: list[dict]) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "badges.pdf"
        export_cards_to_pdf(cards, str(pdf_path))
        return pdf_path.read_bytes()


def render_pdf_preview_images(pdf_bytes: bytes, max_pages: int = 3, dpi: int = 144) -> list[bytes]:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        return []

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "preview.pdf"
        out_prefix = Path(tmpdir) / "page"
        pdf_path.write_bytes(pdf_bytes)
        cmd = [
            pdftoppm,
            "-png",
            "-r",
            str(dpi),
            "-f",
            "1",
            "-l",
            str(max_pages),
            str(pdf_path),
            str(out_prefix),
        ]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        except subprocess.TimeoutExpired:
            return []
        images = []
        for idx in range(1, max_pages + 1):
            page_path = Path(f"{out_prefix}-{idx}.png")
            if not page_path.exists():
                break
            images.append(page_path.read_bytes())
        return images
