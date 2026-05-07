"""
HTML 工牌模板生成器 - position:absolute 严格还原 SVG 布局

所有坐标直接来自 SVG 解析结果（A星级员工工牌输出版.svg）：
  - SCALE = 210 / 595.27557 mm/SVG-unit
  - 工牌在 A4 上：left=9mm, top=6.07mm，行间距 57.47mm
  - 工牌尺寸：96mm × 55mm
"""

import base64, os, io

_DIR = os.path.dirname(__file__)

# ── 资产路径 ──────────────────────────────────────────────
_STAR_B64_PATH = os.path.join(_DIR, "assets", "star_gold_b64.txt")
_QR_PATH       = os.path.join(_DIR, "assets", "qr_wechat.png")

def _get_star_b64() -> str:
    try:
        return open(_STAR_B64_PATH).read().strip()
    except Exception:
        return ""

def _get_qr_b64() -> str:
    try:
        return base64.b64encode(open(_QR_PATH, "rb").read()).decode()
    except Exception:
        return ""

# ── 来自 SVG 的精确布局常数（所有单位 mm，SCALE=210/595.27557）──────────────────
# 工牌尺寸
CARD_W = 96.000
CARD_H = 55.000

# A4 网格位置
A4_COL1_X = 9.000
A4_COL2_X = 105.000
A4_ROW_Y  = [6.070, 63.540, 121.010, 178.480, 235.950]

# ── 页眉高度：9.535mm（来自 path[395] 两段贝塞尔累积 dy）──
# path[395] 有三个子路径（Z分隔）：
#   1. 主红色条:  x=0~61.586mm, h=9.535mm (右侧 S 曲线内凹)
#   2. 左胶囊pill: x≈63.24~76.51mm, h=9.535mm (左侧弧线)
#   3. 右胶囊pill: x≈79.72~92.99mm, h=9.535mm (右侧弧线)
HDR_H = 9.535

# 页眉三段 SVG 路径（精确贝塞尔，单位 mm，viewBox="0 0 96 9.535"）
# 主红色条
HDR_PATH_MAIN = "M0,0 H61.586 C60.602,1.334 60.020,2.982 60.020,4.767 C60.020,6.551 60.601,8.200 61.586,9.535 H0 Z"
# 左胶囊（左侧弧线凸出至 x≈63.24mm）
HDR_PATH_LEFT_PILL  = "M67.320,0 H76.506 V9.535 H67.320 C65.011,9.177 63.238,7.177 63.238,4.768 C63.238,2.360 65.013,0.358 67.320,0 Z"
# 右胶囊（右侧弧线凸出至 x≈92.99mm）
HDR_PATH_RIGHT_PILL = "M79.722,0 H88.906 C91.216,0.358 92.990,2.359 92.990,4.767 C92.990,7.175 91.216,9.176 88.906,9.535 H79.722 V0 Z"

# 角标（白色圆角方块 + 红色底条）
# path[397]: 圆角矩形，top=1.867mm, left=2.930mm, 总高=10.263mm
# path[399]: 角标底部红色条，y=9.535mm, h=2.594mm（延伸至工牌身体区）
BADGE_L  = 2.930
BADGE_T  = 1.867
BADGE_W  = 10.263
BADGE_H  = 10.263
BADGE_R  = 1.900
BADGE_RED_T = 9.536   # 角标红色底条 top（= 页眉底）
BADGE_RED_H = 2.594   # 角标红色底条高度

# 5 颗星（暗色底板 + 星图）
# BG left 间距=11.598mm，BG width 取星图宽度使各格不重叠
STAR_BG_TOPS  = [30.935, 42.533, 54.131, 65.729, 77.326]  # left
STAR_BG_T     = 11.107   # top
STAR_BG_W     = 11.354   # 与星图等宽（原 14.584mm 会互相重叠）
STAR_BG_H     = 14.000

STAR_TOPS     = [32.550, 44.148, 55.746, 67.344, 78.941]  # left
STAR_T        = 12.722   # top
STAR_W        = 11.354
STAR_H        = 10.770

# 照片框
PHOTO_L = 5.100
PHOTO_T = 16.550
PHOTO_W = 23.540
PHOTO_H = 32.950

# 文字字段（3个，来自 SVG 下划线路径）
FIELD_LABEL_X   = 35.150   # 标签起始 x
FIELD_UL_X      = 43.270   # 下划线起始 x
FIELD_UL_W      = 23.260   # 下划线宽度
FIELD_1_TY      = 30.970   # 字段1 文字 top
FIELD_1_ULY     = 34.550   # 字段1 下划线
FIELD_2_TY      = 38.410
FIELD_2_ULY     = 41.910
FIELD_3_TY      = 45.320
FIELD_3_ULY     = 48.820

# QR 码框（top 与姓名字段顶部对齐）
QR_L = 69.880
QR_T = 30.970   # 与 FIELD_1_TY 对齐（原 28.020mm）
QR_W = 20.040
QR_H = 18.500   # 留出底部标签空间（bottom=49.470mm，label+0.4=49.870mm，footer=52.138mm）

# 底部红色条
FOOT_T = 52.138
FOOT_H = 2.862


# ── CSS ──────────────────────────────────────────────────────
_CSS = """
/* A4 页面 */
.a4-page {
  position: relative;
  width: 210mm;
  height: 297mm;
  background: #f0f0f0;
  box-sizing: border-box;
  overflow: hidden;
}
/* 工牌容器 */
.card-wrap {
  position: absolute;
  width: 96mm;
  height: 55mm;
  overflow: hidden;
  box-shadow: 0 0 0 0.3mm #c0c0c0;
}
/* 工牌内所有子元素均绝对定位 */
.card-wrap * {
  box-sizing: border-box;
}
/* 角标（占位，实际样式均用 inline style）*/
.badge-box {
  position: absolute;
  left: 2.930mm; top: 1.867mm;
  width: 10.263mm; height: 10.263mm;
  background: white;
  border-radius: 1.900mm;
  overflow: hidden;
  display: flex; align-items: center; justify-content: center;
}
/* 品牌文字 */
.brand-cn {
  position: absolute;
  left: 14.5mm; top: 2.3mm;
  font-size: 4.2mm;
  font-weight: 700;
  color: white;
  white-space: nowrap;
  letter-spacing: 0.2mm;
  font-family: "Noto Sans SC", "Source Han Sans CN", "Microsoft YaHei", sans-serif;
}
.brand-en {
  position: absolute;
  left: 14.5mm; top: 8.2mm;
  font-size: 2.0mm;
  font-weight: 600;
  color: rgba(255,255,255,0.90);
  white-space: nowrap;
  letter-spacing: 0.15mm;
  font-family: Arial, Helvetica, sans-serif;
}
/* 星星暗色背景 */
.star-bg {
  position: absolute;
  top: 11.107mm;
  width: 14.584mm; height: 14.000mm;
  background: #1b1918;
}
/* 星星图片 */
.star-img {
  position: absolute;
  top: 12.722mm;
  width: 11.354mm; height: 10.770mm;
  object-fit: contain;
}
.star-fallback {
  position: absolute;
  top: 12.722mm;
  width: 11.354mm; height: 10.770mm;
  display: flex; align-items: center; justify-content: center;
  font-size: 8mm;
  color: #FFD700;
}
/* 照片占位 */
.photo-box {
  position: absolute;
  left: 5.100mm; top: 16.550mm;
  width: 23.540mm; height: 32.950mm;
  background: #f8f8f8;
  border: 0.3mm solid #cccccc;
  overflow: hidden;
}
.photo-box img {
  width: 100%; height: 100%;
  object-fit: cover;
}
.photo-placeholder {
  position: relative;
  width: 100%; height: 100%;
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
}
.photo-placeholder svg {
  position: absolute;
  width: 100%; height: 100%;
  top: 0; left: 0;
}
.photo-label {
  position: relative;
  z-index: 1;
  font-size: 2.8mm;
  color: #888;
  letter-spacing: 0.5mm;
  margin-top: 2mm;
  font-family: "Microsoft YaHei", sans-serif;
}
/* 角色标签（工牌角标）*/
.role-badge {
  position: absolute;
  left: 2.930mm; top: 13.200mm;
  width: 10.262mm;
  background: #b60814;
  border-radius: 0 0 1.5mm 1.5mm;
  text-align: center;
  padding: 0.4mm 0;
}
.role-cn {
  display: block;
  font-size: 2.5mm;
  font-weight: 700;
  color: white;
  font-family: "Microsoft YaHei", sans-serif;
  line-height: 1.3;
}
.role-en {
  display: block;
  font-size: 1.5mm;
  color: rgba(255,255,255,0.85);
  font-family: Arial, sans-serif;
  line-height: 1.2;
  letter-spacing: 0.1mm;
}
/* 文字字段 */
.field-label {
  position: absolute;
  left: 35.150mm;
  font-size: 3.0mm;
  font-weight: 600;
  color: #1b1918;
  font-family: "Microsoft YaHei", sans-serif;
  white-space: nowrap;
}
.field-value {
  position: absolute;
  left: 43.270mm; width: 23.260mm;
  font-size: 3.0mm;
  color: #1b1918;
  font-family: "Microsoft YaHei", sans-serif;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.field-ul {
  position: absolute;
  left: 43.270mm; width: 23.260mm;
  height: 0;
  border-bottom: 0.3mm solid #888;
}
/* QR 码 */
.qr-box {
  position: absolute;
  left: 69.880mm; top: 28.020mm;
  width: 20.040mm; height: 20.040mm;
  overflow: hidden;
  border: 0.2mm solid #bbb;
  background: white;
}
.qr-box img {
  width: 100%; height: 100%;
  object-fit: contain;
}
.qr-placeholder {
  width: 100%; height: 100%;
  position: relative;
  display: flex; align-items: center; justify-content: center;
}
.qr-placeholder svg {
  position: absolute;
  width: 100%; height: 100%;
  top: 0; left: 0;
}
.qr-label {
  position: absolute;
  left: 69.880mm;
  top: 50.000mm;
  width: 20.040mm;
  text-align: center;
  font-size: 2.2mm;
  color: #555;
  font-family: "Microsoft YaHei", sans-serif;
}
/* 底部红条 */
.card-foot {
  position: absolute;
  left: 0; top: 52.138mm;
  width: 96mm; height: 2.862mm;
  background: #e50036;
}
/* 水印 */
.watermark {
  position: absolute;
  left: 0; top: 0;
  width: 96mm; height: 55mm;
  display: flex; align-items: center; justify-content: center;
  pointer-events: none;
  z-index: 50;
}
.watermark-text {
  font-size: 28mm;
  font-weight: 900;
  color: rgba(255,0,0,0.08);
  transform: rotate(-30deg);
  white-space: nowrap;
  font-family: "Microsoft YaHei", sans-serif;
  user-select: none;
}
/* 打印 */
@media print {
  body { margin: 0; }
  .a4-page { box-shadow: none; background: white; }
  .card-wrap { box-shadow: none; }
  .watermark-text { color: rgba(255,0,0,0.06); }
}
"""


# ── 工牌 HTML 生成（单张）─────────────────────────────────────
def _card_html(
    name: str,
    title: str,
    staff_id: str,
    star_count: int,
    photo_b64: str,
    qr_b64: str,
    star_b64: str,
    watermark: str = "",
) -> str:
    n = max(0, min(5, star_count))

    # 星星元素
    star_elems = []
    for i in range(5):
        lx = STAR_BG_TOPS[i]
        if i < n:
            bg_div = f'<div class="star-bg" style="position:absolute;left:{lx:.3f}mm;top:{STAR_BG_T:.3f}mm;width:{STAR_BG_W:.3f}mm;height:{STAR_BG_H:.3f}mm;background:#1b1918;"></div>'
        else:
            bg_div = ""

        if i < n:
            lxi = STAR_TOPS[i]
            if star_b64:
                img = f'<img style="position:absolute;left:{lxi:.3f}mm;top:{STAR_T:.3f}mm;width:{STAR_W:.3f}mm;height:{STAR_H:.3f}mm;object-fit:contain;" src="data:image/png;base64,{star_b64}" alt="★"/>'
            else:
                img = f'<div style="position:absolute;left:{lxi:.3f}mm;top:{STAR_T:.3f}mm;width:{STAR_W:.3f}mm;height:{STAR_H:.3f}mm;display:flex;align-items:center;justify-content:center;font-size:8mm;color:#FFD700;">★</div>'
        else:
            img = ""

        star_elems.append(bg_div + img)

    stars_html = "\n".join(star_elems)

    # 照片
    if photo_b64:
        photo_inner = f'<img src="data:image/jpeg;base64,{photo_b64}" alt="photo"/>'
    else:
        photo_inner = f"""<div class="photo-placeholder">
  <svg viewBox="0 0 100 100" preserveAspectRatio="none">
    <line x1="0" y1="0" x2="100" y2="100" stroke="#ccc" stroke-width="1.5"/>
    <line x1="100" y1="0" x2="0" y2="100" stroke="#ccc" stroke-width="1.5"/>
  </svg>
  <span class="photo-label">寸照</span>
</div>"""

    # QR 码
    if qr_b64:
        qr_inner = f'<img src="data:image/png;base64,{qr_b64}" alt="QR"/>'
    else:
        qr_inner = f"""<div class="qr-placeholder">
  <svg viewBox="0 0 100 100" preserveAspectRatio="none">
    <line x1="0" y1="0" x2="100" y2="100" stroke="#bbb" stroke-width="1.5"/>
    <line x1="100" y1="0" x2="0" y2="100" stroke="#bbb" stroke-width="1.5"/>
  </svg>
</div>"""

    # 水印
    wm_html = ""
    if watermark:
        wm_html = f'<div class="watermark"><span class="watermark-text">{watermark}</span></div>'

    # 角色英文映射
    role_map = {
        "店长": "Store Manager",
        "副店长": "Deputy Manager",
        "营业员": "Sales Associate",
        "收银员": "Cashier",
        "药师": "Pharmacist",
        "主管": "Supervisor",
    }
    role_en = role_map.get(title.strip(), "Staff")

    return f"""<!-- card start -->
<div class="card-wrap">
  <!-- 工牌背景 -->
  <div style="position:absolute;inset:0;background:#fbfaf9;"></div>

  <!-- ══ 页眉区（严格按 SVG path[395] 三子路径）══ -->
  <!-- 页眉 SVG：主红色条 + 左胶囊 + 右胶囊，精确贝塞尔，viewBox 单位=mm -->
  <svg style="position:absolute;left:0;top:0;width:96mm;height:{HDR_H:.3f}mm;overflow:visible;"
       viewBox="0 0 96 9.535" preserveAspectRatio="none"
       xmlns="http://www.w3.org/2000/svg">
    <!-- 主红色条（右侧 S 曲线内凹，x=0→61.586mm） -->
    <path fill="#e50036" d="{HDR_PATH_MAIN}"/>
    <!-- 左胶囊 pill（左侧弧线，x≈63.24→76.51mm） -->
    <path fill="#e50036" d="{HDR_PATH_LEFT_PILL}"/>
    <!-- 右胶囊 pill（右侧弧线，x≈79.72→92.99mm） -->
    <path fill="#e50036" d="{HDR_PATH_RIGHT_PILL}"/>
  </svg>

  <!-- ══ 角标（白色圆角方块，protrudes below header）══ -->
  <!-- path[397]: left=2.930mm top=1.867mm w=10.263mm h=10.263mm r=1.900mm -->
  <div style="position:absolute;left:{BADGE_L:.3f}mm;top:{BADGE_T:.3f}mm;width:{BADGE_W:.3f}mm;height:{BADGE_H:.3f}mm;background:white;border-radius:{BADGE_R:.3f}mm;overflow:hidden;display:flex;flex-direction:column;align-items:center;justify-content:center;z-index:2;">
    <!-- 小药十字（path[398] 红色 cross，居中） -->
    <svg viewBox="0 0 10 10" width="5mm" height="5mm" style="flex-shrink:0;">
      <rect x="3.5" y="1" width="3" height="8" rx="0.5" fill="#e50036"/>
      <rect x="1" y="3.5" width="8" height="3" rx="0.5" fill="#e50036"/>
    </svg>
    <!-- 职务文字（path[400] 小字，在角标下红条内，只显示中文） -->
    <span style="font-size:2.2mm;font-weight:700;color:#e50036;font-family:'Microsoft YaHei',sans-serif;margin-top:0.5mm;line-height:1;">{title}</span>
  </div>

  <!-- path[399]: 角标底部红条（y=9.536mm，h=2.594mm，圆角底） -->
  <div style="position:absolute;left:{BADGE_L:.3f}mm;top:{BADGE_RED_T:.3f}mm;width:{BADGE_W:.3f}mm;height:{BADGE_RED_H:.3f}mm;background:#e50036;border-radius:0 0 {BADGE_R:.3f}mm {BADGE_R:.3f}mm;display:flex;align-items:center;justify-content:center;z-index:2;">
    <span style="font-size:1.4mm;color:white;font-family:Arial,sans-serif;letter-spacing:0.1mm;white-space:nowrap;">{role_en}</span>
  </div>

  <!-- 品牌名称（left=14.5mm，在主红色条内 h={HDR_H:.3f}mm） -->
  <span style="position:absolute;left:14.5mm;top:1.2mm;font-size:3.8mm;font-weight:700;color:white;white-space:nowrap;letter-spacing:0.2mm;font-family:'Noto Sans SC','Microsoft YaHei',sans-serif;">高济药房</span>
  <span style="position:absolute;left:14.5mm;top:6.2mm;font-size:1.8mm;font-weight:600;color:rgba(255,255,255,0.90);white-space:nowrap;letter-spacing:0.15mm;font-family:Arial,Helvetica,sans-serif;">COWELL PHARMACY</span>

  <!-- 星星暗色背景 + 星图 -->
  {stars_html}

  <!-- 照片框：left=5.100mm top=16.550mm w=23.540mm h=32.950mm（来自 SVG path[407]） -->
  <div style="position:absolute;left:{PHOTO_L:.3f}mm;top:{PHOTO_T:.3f}mm;width:{PHOTO_W:.3f}mm;height:{PHOTO_H:.3f}mm;background:#f8f8f8;border:0.3mm solid #ccc;overflow:hidden;">
    {photo_inner}
  </div>

  <!-- 姓名 -->
  <span style="position:absolute;left:{FIELD_LABEL_X:.3f}mm;top:{FIELD_1_TY:.3f}mm;font-size:3.0mm;font-weight:600;color:#1b1918;font-family:'Microsoft YaHei',sans-serif;white-space:nowrap;">姓名：</span>
  <span style="position:absolute;left:{FIELD_UL_X:.3f}mm;top:{FIELD_1_TY:.3f}mm;width:{FIELD_UL_W:.3f}mm;font-size:3.0mm;color:#1b1918;font-family:'Microsoft YaHei',sans-serif;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{name}</span>
  <div  style="position:absolute;left:{FIELD_UL_X:.3f}mm;top:{FIELD_1_ULY:.3f}mm;width:{FIELD_UL_W:.3f}mm;height:0;border-bottom:0.3mm solid #888;"></div>

  <!-- 职务 -->
  <span style="position:absolute;left:{FIELD_LABEL_X:.3f}mm;top:{FIELD_2_TY:.3f}mm;font-size:3.0mm;font-weight:600;color:#1b1918;font-family:'Microsoft YaHei',sans-serif;white-space:nowrap;">职务：</span>
  <span style="position:absolute;left:{FIELD_UL_X:.3f}mm;top:{FIELD_2_TY:.3f}mm;width:{FIELD_UL_W:.3f}mm;font-size:3.0mm;color:#1b1918;font-family:'Microsoft YaHei',sans-serif;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{title}</span>
  <div  style="position:absolute;left:{FIELD_UL_X:.3f}mm;top:{FIELD_2_ULY:.3f}mm;width:{FIELD_UL_W:.3f}mm;height:0;border-bottom:0.3mm solid #888;"></div>

  <!-- 工号 -->
  <span style="position:absolute;left:{FIELD_LABEL_X:.3f}mm;top:{FIELD_3_TY:.3f}mm;font-size:3.0mm;font-weight:600;color:#1b1918;font-family:'Microsoft YaHei',sans-serif;white-space:nowrap;">工号：</span>
  <span style="position:absolute;left:{FIELD_UL_X:.3f}mm;top:{FIELD_3_TY:.3f}mm;width:{FIELD_UL_W:.3f}mm;font-size:3.0mm;color:#1b1918;font-family:'Microsoft YaHei',sans-serif;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{staff_id}</span>
  <div  style="position:absolute;left:{FIELD_UL_X:.3f}mm;top:{FIELD_3_ULY:.3f}mm;width:{FIELD_UL_W:.3f}mm;height:0;border-bottom:0.3mm solid #888;"></div>

  <!-- QR 码（top 与姓名对齐 = {QR_T:.3f}mm） -->
  <div style="position:absolute;left:{QR_L:.3f}mm;top:{QR_T:.3f}mm;width:{QR_W:.3f}mm;height:{QR_H:.3f}mm;overflow:hidden;border:0.2mm solid #bbb;background:white;">
    {qr_inner}
  </div>
  <!-- 企微码标签：QR 正下方 -->
  <span style="position:absolute;left:{QR_L:.3f}mm;top:{QR_T + QR_H + 0.4:.3f}mm;width:{QR_W:.3f}mm;text-align:center;font-size:2.0mm;color:#555;font-family:'Microsoft YaHei',sans-serif;">企微码</span>

  <!-- 底部红条 -->
  <div style="position:absolute;left:0;top:{FOOT_T:.3f}mm;width:96mm;height:{FOOT_H:.3f}mm;background:#e50036;"></div>

  {wm_html}
</div>
<!-- card end -->"""


# ── 构建单卡 HTML（独立预览）────────────────────────────────────
def build_card_html(
    name: str = "张三",
    title: str = "店长",
    staff_id: str = "GJ-001",
    star_count: int = 5,
    photo_b64: str = "",
    qr_b64: str = "",
    watermark: str = "",
) -> str:
    star_b64 = _get_star_b64()
    if not qr_b64:
        qr_b64 = _get_qr_b64()

    card = _card_html(
        name=name, title=title, staff_id=staff_id,
        star_count=star_count, photo_b64=photo_b64,
        qr_b64=qr_b64, star_b64=star_b64, watermark=watermark,
    )

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>员工工牌预览</title>
<style>
{_CSS}
body {{ background: #e0e0e0; display: flex; align-items: center; justify-content: center;
       min-height: 100vh; margin: 0; }}
</style>
</head>
<body>
<div style="transform:scale(3);transform-origin:top left;margin:20px;">
  {card}
</div>
</body>
</html>"""


# ── 构建 A4 完整版（10 张卡）─────────────────────────────────────
def build_a4_html(
    cards_data: list,
    watermark: str = "",
) -> str:
    """
    cards_data: list of dict with keys: name, title, staff_id, star_count,
                photo_b64(opt), qr_b64(opt)
    按 SVG 原始位置排版：2列×5行，精确坐标。
    """
    star_b64 = _get_star_b64()
    default_qr = _get_qr_b64()

    cards_html = []
    for idx, d in enumerate(cards_data[:10]):
        row = idx // 2
        col = idx % 2
        left = A4_COL1_X if col == 0 else A4_COL2_X
        top  = A4_ROW_Y[row]

        qr_b64 = d.get("qr_b64", "") or default_qr
        card = _card_html(
            name=d.get("name", ""),
            title=d.get("title", ""),
            staff_id=d.get("staff_id", ""),
            star_count=d.get("star_count", 5),
            photo_b64=d.get("photo_b64", ""),
            qr_b64=qr_b64,
            star_b64=star_b64,
            watermark=watermark,
        )
        cards_html.append(
            f'<div style="position:absolute;left:{left}mm;top:{top}mm;">\n{card}\n</div>'
        )

    cards_str = "\n".join(cards_html)

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>高济药房员工工牌 A4 打印版</title>
<style>
{_CSS}
html, body {{
  margin: 0; padding: 0;
  background: #888;
}}
.a4-page {{
  margin: 10mm auto;
}}
@media print {{
  body {{ background: white; }}
  .a4-page {{ margin: 0; }}
}}
</style>
</head>
<body>
<div class="a4-page">
  {cards_str}
</div>
</body>
</html>"""


# ── 快速测试 ────────────────────────────────────────────────────
if __name__ == "__main__":
    sample_cards = [
        {"name": f"员工{i+1:02d}", "title": "店长", "staff_id": f"GJ-{i+1:03d}", "star_count": 5}
        for i in range(10)
    ]
    html = build_a4_html(sample_cards)
    out = "/tmp/badge_a4_abs.html"
    open(out, "w").write(html)
    print(f"Saved: {out}")
    # 同时保存单卡预览
    single = build_card_html(name="张三", title="店长", staff_id="GJ-001", star_count=5)
    open("/tmp/badge_single_abs.html", "w").write(single)
    print("Single card: /tmp/badge_single_abs.html")
