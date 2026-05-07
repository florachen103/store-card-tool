"""
main.py —— 底层绘图模块
版式对齐「A 星级员工工牌输出版」CDR：
  · 顶栏：白色圆角职务角标 + 店名 + 右侧胶囊装饰
  · 身份区：左照片、居中五星、姓名/职务/工号+下划线
  · 右侧居中：二维码（标注"企微码"）
  · 极淡水印字
"""

from PIL import Image, ImageDraw, ImageFont
import qrcode
import os
import io
import math

DPI = 300


def mm2px(mm):
    return int(mm / 25.4 * DPI)


# ── A4 页面 ──────────────────────────────────────────────
A4_W = mm2px(210)
A4_H = mm2px(297)

# ── 单张工牌（96×55 mm）───────────────────────────────────
CARD_W = mm2px(96)
CARD_H = mm2px(55)

# ── A4 拼版：2列 × 5行 ────────────────────────────────────
COLS = 2
ROWS = 5
MARGIN_X = mm2px(5)
MARGIN_Y = mm2px(2)
GAP_X = (A4_W - 2 * MARGIN_X - COLS * CARD_W) // (COLS - 1)
GAP_Y = (A4_H - 2 * MARGIN_Y - ROWS * CARD_H) // (ROWS - 1)
if GAP_X < 0 or GAP_Y < 0:
    raise ValueError(f"A4 放不下当前布局：GAP_X={GAP_X}, GAP_Y={GAP_Y}")

# ── 工牌版式参数（严格对齐 CDR「A星级员工工牌输出版」）──────
HEADER_H = mm2px(13.5)   # 红色顶栏高（CDR ~25% of 55mm）
FOOTER_H = mm2px(3.0)    # 红色底条高（CDR ~5.5% of 55mm）

# 顶栏左侧白色圆角职务角标（含王冠 + 职务 + Store Manager）
BADGE_W  = mm2px(11.0)   # 宽
BADGE_H  = mm2px(10.5)   # 高
BADGE_X  = mm2px(1.5)
BADGE_Y  = (HEADER_H - BADGE_H) // 2

# 店招文字（可修改为实际门店名）
BRAND_NAME_CN = "高济药房"
BRAND_NAME_EN = "COWELL PHARMACY"

# 照片区（左侧寸照）：22×34 mm
PHOTO_W = mm2px(22)
PHOTO_H = mm2px(34)
PHOTO_X = mm2px(2.0)
PHOTO_Y = HEADER_H + mm2px(2.0)

# 文字区左边界（照片右侧）
TEXT_X    = PHOTO_X + PHOTO_W + mm2px(2.5)
TEXT_LINE = mm2px(6.5)

# 二维码（右侧，与主体区域对齐）
QR_SIZE = mm2px(15)
QR_X    = CARD_W - QR_SIZE - mm2px(2.0)
_BODY_H = CARD_H - FOOTER_H - HEADER_H
QR_Y    = HEADER_H + (_BODY_H - QR_SIZE - mm2px(3.0)) // 2

# 五星（与 CDR 保持一致：较大的 3D 金色星）
STAR_OUTER = mm2px(4.5)   # CDR 中星星约占 body 高度的 25%
STAR_INNER = mm2px(1.8)
STAR_GAP   = mm2px(0.6)

DEFAULT_WATERMARK = "A"

# ── 3D 金色星星资产（与 CDR 完全一致）────────────────────
_STAR_IMG_PATH = os.path.join(os.path.dirname(__file__), "emb_4.png")
_star_img_cache = None


def _get_star_image():
    """加载 3D 金色星星图片，缩放到 STAR_OUTER×2 大小，失败则返回 None。"""
    global _star_img_cache
    if _star_img_cache is not None:
        return _star_img_cache
    try:
        sz = STAR_OUTER * 2
        img = Image.open(_STAR_IMG_PATH).convert("RGBA")
        img = img.resize((sz, sz), Image.LANCZOS)
        _star_img_cache = img
    except Exception:
        _star_img_cache = False
    return _star_img_cache if _star_img_cache else None


# 颜色
COLOR_RED       = "#e50036"   # 品牌红（从 SVG 提取的准确值）
COLOR_RED_LIGHT = "#ff3355"   # 胶囊装饰
COLOR_WHITE     = "#FFFFFF"
COLOR_BLACK     = "#1A1A1A"
COLOR_GRAY      = "#AAAAAA"
COLOR_LIGHT_BG  = "#F5F5F5"
COLOR_GOLD      = "#D4AF37"
COLOR_GOLD_EDGE = "#B8860B"

# ── 字体候选（macOS 中文字体）────────────────────────────
FONT_CANDIDATES = [
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    "/System/Library/Fonts/Supplemental/STHeiti Medium.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
]
_font_cache = {}


def load_font(size):
    size = max(1, size)
    if size in _font_cache:
        return _font_cache[size]
    for path in FONT_CANDIDATES:
        if os.path.exists(path):
            try:
                f = ImageFont.truetype(path, size=size)
                _font_cache[size] = f
                return f
            except Exception:
                continue
    f = ImageFont.load_default()
    _font_cache[size] = f
    return f


# ── 辅助绘图 ─────────────────────────────────────────────

def _star_vertices(cx, cy, r_o, r_i):
    pts = []
    for k in range(10):
        ang = -math.pi / 2 + k * math.pi / 5
        r = r_o if k % 2 == 0 else r_i
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    return pts


def _draw_star(draw, cx, cy, fill, outline=None):
    draw.polygon(_star_vertices(cx, cy, STAR_OUTER, STAR_INNER),
                 fill=fill, outline=outline or fill)


def _make_photo_placeholder():
    """带对角线 X 的照片占位图（与 CDR 模版一致）"""
    img = Image.new("RGB", (PHOTO_W, PHOTO_H), COLOR_LIGHT_BG)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, PHOTO_W - 1, PHOTO_H - 1], outline="#C0C0C0", width=2)
    d.line([3, 3, PHOTO_W - 3, PHOTO_H - 3], fill="#C8C8C8", width=2)
    d.line([PHOTO_W - 3, 3, 3, PHOTO_H - 3], fill="#C8C8C8", width=2)
    font = load_font(mm2px(2.8))
    label = "寸照"
    bbox = d.textbbox((0, 0), label, font=font)
    lw, lh = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text(((PHOTO_W - lw) // 2, (PHOTO_H - lh) // 2), label,
           fill="#BBBBBB", font=font)
    return img


def _make_qr_placeholder():
    """带 X 的二维码占位图"""
    img = Image.new("RGB", (QR_SIZE, QR_SIZE), COLOR_WHITE)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, QR_SIZE - 1, QR_SIZE - 1], outline="#C8C8C8", width=2)
    d.line([3, 3, QR_SIZE - 3, QR_SIZE - 3], fill="#C8C8C8", width=1)
    d.line([QR_SIZE - 3, 3, 3, QR_SIZE - 3], fill="#C8C8C8", width=1)
    return img


def _load_photo(photo_source):
    try:
        if isinstance(photo_source, Image.Image):
            img = photo_source.convert("RGB")
        elif isinstance(photo_source, (bytes, bytearray, io.BytesIO)):
            buf = photo_source if isinstance(photo_source, io.BytesIO) else io.BytesIO(photo_source)
            img = Image.open(buf).convert("RGB")
        elif photo_source and os.path.exists(str(photo_source)):
            img = Image.open(str(photo_source)).convert("RGB")
        else:
            raise FileNotFoundError
        # 等比缩放 + 居中裁切
        src_w, src_h = img.size
        scale = max(PHOTO_W / src_w, PHOTO_H / src_h)
        new_w, new_h = int(src_w * scale), int(src_h * scale)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        left = (new_w - PHOTO_W) // 2
        top  = (new_h - PHOTO_H) // 2
        return img.crop((left, top, left + PHOTO_W, top + PHOTO_H))
    except Exception:
        return _make_photo_placeholder()


def _load_qr(qr_source):
    try:
        if isinstance(qr_source, Image.Image):
            img = qr_source.convert("RGB")
        elif isinstance(qr_source, (bytes, bytearray, io.BytesIO)):
            buf = qr_source if isinstance(qr_source, io.BytesIO) else io.BytesIO(qr_source)
            img = Image.open(buf).convert("RGB")
        elif qr_source and os.path.exists(str(qr_source)):
            img = Image.open(str(qr_source)).convert("RGB")
        elif qr_source:
            img = qrcode.make(str(qr_source)).convert("RGB")
        else:
            raise ValueError
        return img.resize((QR_SIZE, QR_SIZE), Image.LANCZOS)
    except Exception:
        return _make_qr_placeholder()


def _load_brand_logo(logo_source):
    if logo_source is None:
        return None
    try:
        if isinstance(logo_source, Image.Image):
            return logo_source.convert("RGBA")
        if isinstance(logo_source, (bytes, bytearray, io.BytesIO)):
            buf = logo_source if isinstance(logo_source, io.BytesIO) else io.BytesIO(logo_source)
            return Image.open(buf).convert("RGBA")
        if isinstance(logo_source, str) and os.path.exists(logo_source):
            return Image.open(logo_source).convert("RGBA")
    except Exception:
        pass
    return None


# ── 核心绘卡函数 ──────────────────────────────────────────

def draw_card(
    name,
    title,
    staff_id,
    photo_source,
    qr_source,
    *,
    watermark_char=None,
    logo_source=None,
):
    """
    绘制单张工牌，返回 PIL Image。

    版式结构（对齐 CDR「A星级员工工牌输出版」）：
      · 顶栏：[职务角标] [店名/英文] [胶囊装饰]
      · 身份区：[照片] [★★★★★] [姓名/职务/工号+横线]
      · 右侧：[二维码] [企微码]
      · 极浅水印字（默认「A」）
    """
    wm = (str(watermark_char).strip()[:1]
          if watermark_char is not None and str(watermark_char).strip()
          else DEFAULT_WATERMARK)

    card = Image.new("RGB", (CARD_W, CARD_H), COLOR_WHITE)
    draw = ImageDraw.Draw(card)

    # ── 顶栏红底 ─────────────────────────────────────────
    draw.rectangle([0, 0, CARD_W, HEADER_H], fill=COLOR_RED)

    # 职务角标（白色圆角矩形，内含职务文字）
    bx0, by0 = BADGE_X, BADGE_Y
    bx1, by1 = bx0 + BADGE_W, by0 + BADGE_H
    draw.rounded_rectangle([bx0, by0, bx1, by1], radius=mm2px(1.0), fill=COLOR_WHITE)

    role_text = str(title).strip() if title else "员工"
    role_lines = [role_text[i:i+2] for i in range(0, min(len(role_text), 4), 2)]
    font_role = load_font(mm2px(2.4))
    total_role_h = sum(
        draw.textbbox((0, 0), ln, font=font_role)[3] - draw.textbbox((0, 0), ln, font=font_role)[1]
        for ln in role_lines
    ) + mm2px(0.4) * max(0, len(role_lines) - 1)
    ry = by0 + (BADGE_H - total_role_h) // 2
    for ln in role_lines:
        rb = draw.textbbox((0, 0), ln, font=font_role)
        rw = rb[2] - rb[0]
        draw.text((bx0 + (BADGE_W - rw) // 2, ry), ln, fill=COLOR_RED, font=font_role)
        ry += (rb[3] - rb[1]) + mm2px(0.4)

    # 店招（字号随顶栏高度增大）
    brand_x = bx1 + mm2px(2.0)
    font_cn = load_font(mm2px(4.5))
    font_en = load_font(mm2px(2.0))
    bcn = draw.textbbox((0, 0), BRAND_NAME_CN, font=font_cn)
    ben = draw.textbbox((0, 0), BRAND_NAME_EN, font=font_en)
    brand_block_h = (bcn[3] - bcn[1]) + mm2px(0.8) + (ben[3] - ben[1])
    bty = (HEADER_H - brand_block_h) // 2
    draw.text((brand_x, bty), BRAND_NAME_CN, fill=COLOR_WHITE, font=font_cn)
    # 品牌英文前加小图标占位（填充一个小色块）
    en_y = bty + (bcn[3] - bcn[1]) + mm2px(0.6)
    icon_size = mm2px(2.8)
    draw.rounded_rectangle(
        [brand_x, en_y, brand_x + icon_size, en_y + icon_size],
        radius=mm2px(0.4), fill=COLOR_WHITE
    )
    draw.text((brand_x + icon_size + mm2px(0.8), en_y),
              BRAND_NAME_EN, fill=COLOR_WHITE, font=font_en)

    # 顶栏右侧胶囊装饰（两颗椭圆，白/红）
    pcx = CARD_W - mm2px(5.0)
    pcw, pch = mm2px(7.0), mm2px(3.0)
    cy_top = HEADER_H // 2 - mm2px(2.2)
    cy_bot = HEADER_H // 2 + mm2px(2.2)
    draw.ellipse([pcx - pcw // 2, cy_top - pch // 2,
                  pcx + pcw // 2, cy_top + pch // 2], fill=COLOR_WHITE)
    draw.ellipse([pcx - pcw // 2, cy_bot - pch // 2,
                  pcx + pcw // 2, cy_bot + pch // 2],
                 fill=COLOR_RED_LIGHT, outline=COLOR_WHITE, width=1)

    # 底条
    draw.rectangle([0, CARD_H - FOOTER_H, CARD_W, CARD_H], fill=COLOR_RED)

    # 极浅水印字（alpha 18，几乎不可见）
    overlay = Image.new("RGBA", card.size, (0, 0, 0, 0))
    ow = ImageDraw.Draw(overlay)
    font_wm = load_font(mm2px(34))
    wbbox = ow.textbbox((0, 0), wm, font=font_wm)
    ww, wh = wbbox[2] - wbbox[0], wbbox[3] - wbbox[1]
    wx = (CARD_W - ww) // 2
    wy = HEADER_H + (_BODY_H - wh) // 2
    ow.text((wx, wy), wm, fill=(210, 210, 210, 18), font=font_wm)
    card = Image.alpha_composite(card.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(card)

    # 照片（左侧，带占位框）
    photo = _load_photo(photo_source)
    card.paste(photo, (PHOTO_X, PHOTO_Y))
    draw.rectangle([PHOTO_X, PHOTO_Y,
                    PHOTO_X + PHOTO_W - 1, PHOTO_Y + PHOTO_H - 1],
                   outline="#C8C8C8", width=1)

    # 二维码（右侧竖向居中）
    qr_img = _load_qr(qr_source)
    card.paste(qr_img, (QR_X, QR_Y))
    draw.rectangle([QR_X, QR_Y, QR_X + QR_SIZE - 1, QR_Y + QR_SIZE - 1],
                   outline="#C0C0C0", width=1)
    font_qr_label = load_font(mm2px(2.2))
    ql_bbox = draw.textbbox((0, 0), "企微码", font=font_qr_label)
    ql_w = ql_bbox[2] - ql_bbox[0]
    draw.text((QR_X + (QR_SIZE - ql_w) // 2,
               QR_Y + QR_SIZE + mm2px(0.5)),
              "企微码", fill=COLOR_GRAY, font=font_qr_label)

    # 五星（优先使用 3D 金色图片，对齐 CDR）
    content_w = QR_X - mm2px(1.0) - TEXT_X
    star_span = 5 * 2 * STAR_OUTER + 4 * STAR_GAP
    star_left = TEXT_X + max(0, (content_w - star_span) // 2)
    star_cy = PHOTO_Y + STAR_OUTER + mm2px(1.5)
    star_img = _get_star_image()
    for i in range(5):
        cx = star_left + STAR_OUTER + i * (2 * STAR_OUTER + STAR_GAP)
        if star_img is not None:
            sx = cx - STAR_OUTER
            sy = star_cy - STAR_OUTER
            card.paste(star_img, (sx, sy), star_img)
        else:
            _draw_star(draw, cx, star_cy, COLOR_GOLD, COLOR_GOLD_EDGE)

    # 姓名 / 职务 / 工号 + 横线
    font_val = load_font(mm2px(3.2))
    font_id  = load_font(mm2px(3.0))
    line_right = QR_X - mm2px(1.0)
    base_y = star_cy + STAR_OUTER + mm2px(3.5)
    rows = [
        ("姓名：", name,     font_val),
        ("职务：", title,    font_val),
        ("工号：", staff_id, font_id),
    ]
    for i, (lab, val, fnt) in enumerate(rows):
        y = base_y + i * TEXT_LINE
        draw.text((TEXT_X, y), lab, fill=COLOR_BLACK, font=fnt)
        lb = draw.textbbox((TEXT_X, y), lab, font=fnt)
        vx = lb[2] + mm2px(0.5)
        vs = str(val) if val is not None else ""
        draw.text((vx, y), vs, fill=COLOR_BLACK, font=fnt)
        vb = draw.textbbox((vx, y), vs, font=fnt)
        line_y = y + max(lb[3] - lb[1], vb[3] - vb[1]) + mm2px(0.5)
        draw.line([TEXT_X, line_y, line_right, line_y],
                  fill="#333333", width=1)

    # 外框（淡灰）
    draw.rectangle([0, 0, CARD_W - 1, CARD_H - 1], outline="#D0D0D0", width=1)
    return card


# ── 拼版 ──────────────────────────────────────────────────

def build_a4_pages(staff_list, photo_dir="input/photos"):
    """
    将员工列表排版成 A4 页面列表（每页 2×5=10 张）。

    staff_list 每条 dict 字段：
        name, title, staff_id, photo, qr_code
        （可选）watermark, logo
    """
    pages = []
    per_page = COLS * ROWS

    for page_idx in range(0, len(staff_list), per_page):
        batch = staff_list[page_idx: page_idx + per_page]
        a4 = Image.new("RGB", (A4_W, A4_H), COLOR_WHITE)

        for i, staff in enumerate(batch):
            col = i % COLS
            row = i // COLS
            x = MARGIN_X + col * (CARD_W + GAP_X)
            y = MARGIN_Y + row * (CARD_H + GAP_Y)

            photo_src = staff.get("photo", "")
            if photo_src and isinstance(photo_src, str) and not os.path.isabs(photo_src):
                photo_src = os.path.join(photo_dir, photo_src)

            logo_src = staff.get("logo")
            if logo_src and isinstance(logo_src, str) and not os.path.isabs(logo_src):
                logo_src = os.path.join(photo_dir, logo_src)

            wm_val = staff.get("watermark")
            wm_kw  = str(wm_val).strip() if wm_val is not None and str(wm_val).strip() else None

            card_img = draw_card(
                name           = str(staff.get("name", "")),
                title          = str(staff.get("title", "")),
                staff_id       = str(staff.get("staff_id", "")),
                photo_source   = photo_src,
                qr_source      = staff.get("qr_code", ""),
                watermark_char = wm_kw,
                logo_source    = logo_src or None,
            )
            a4.paste(card_img, (x, y))

        pages.append(a4)
    return pages


# ── 导出工具 ──────────────────────────────────────────────

def save_pages_as_pdf(pages, output_path):
    if not pages:
        return
    first = pages[0].convert("RGB")
    rest  = [p.convert("RGB") for p in pages[1:]]
    first.save(output_path, save_all=True, append_images=rest, resolution=DPI)


def save_pages_as_images(pages, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    paths = []
    for i, page in enumerate(pages):
        path = os.path.join(output_dir, f"a4_page_{i + 1}.png")
        page.save(path)
        paths.append(path)
    return paths
