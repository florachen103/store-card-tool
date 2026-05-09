"""
excel_image_extractor.py —— 从 Excel 内嵌图片中提取照片和二维码

支持两种插入方式：
  1. 浮动图片（传统插入）：通过 openpyxl ws._images 读取，锚点用 anchor._from.row/col
  2. 单元格绑定图片（新版 Excel "放置在单元格中"）：
     直接解析 xlsx zip 包内的 cellImages XML 节点

行列映射规则：
  - 表头在第 1 行（Excel 1-based），数据从第 2 行开始
  - anchor._from.row 是 0-based，所以数据第一行 = anchor row 1
  - 用表头列名"照片"/"二维码"确定目标列，按锚点列号归类
  - 如果列名匹配不到，按"所有图片列中最小列=照片，次小列=二维码"兜底
"""

import io
import re
import zipfile
from xml.etree import ElementTree as ET
import openpyxl
from PIL import Image
from lxml import etree


# ── 工具函数 ──────────────────────────────────────────────

def _col_index(col_str):
    """
    列标识转 0-based 整数索引。
    支持数字字符串（"3"→2）和字母（"C"→2）。
    """
    col_str = str(col_str).strip()
    if col_str.isdigit():
        return int(col_str) - 1          # 1-based 数字 → 0-based
    # 字母列名，如 A/B/AA
    col_str = col_str.upper()
    result = 0
    for ch in col_str:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def _bytes_to_pil(data: bytes):
    """字节数据转 PIL Image，失败返回 None"""
    try:
        return Image.open(io.BytesIO(data)).convert("RGB")
    except Exception:
        return None


def _normalize_header(text):
    """表头标准化，兼容括号说明和空白差异。"""
    return re.sub(r'\s+|（.*?）|\(.*?\)', '', str(text or ''))


def _parse_dispimg_formula(formula_text):
    """从 =DISPIMG("ID_xxx",1) / _xlfn.DISPIMG(...) 中提取图片 ID。"""
    if not formula_text:
        return None
    match = re.search(r'DISPIMG\("([^"]+)"', str(formula_text), re.IGNORECASE)
    return match.group(1) if match else None


# ── 主提取函数 ────────────────────────────────────────────

def extract_images_from_excel(file_bytes: bytes):
    """
    从 Excel 文件字节中提取内嵌图片。

    返回：
        photo_by_row : dict { Excel行号(1-based, 数据行) -> PIL.Image }
        qr_by_row    : dict { Excel行号(1-based, 数据行) -> PIL.Image }

    行号说明：
        Excel 第1行 = 表头
        Excel 第2行 = 第1条数据  →  photo_by_row key = 2
        以此类推
    """
    # ── Step 1：用 openpyxl 读表头，确定照片/二维码列索引 ──
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    header_row = [str(h).strip() if h is not None else "" for h in header_row]

    photo_col_idx = None   # 0-based
    qr_col_idx    = None   # 0-based
    for i, h in enumerate(header_row):
        normalized = _normalize_header(h)
        if "照片" in normalized:
            photo_col_idx = i
        if "二维码" in normalized or "企微码" in normalized:
            qr_col_idx = i

    print(f"[图片提取] 表头列: {header_row}")
    print(f"[图片提取] 照片列(0-based)={photo_col_idx}, 二维码列(0-based)={qr_col_idx}")

    # ── Step 2：收集所有图片 (行, 列, PIL Image) ──────────
    # 用 set 去重，同一位置只保留第一张
    all_images = {}   # key: (row_1based, col_0based) -> PIL Image

    # 方法A：openpyxl _images（浮动图片）
    _extract_floating(ws, all_images)

    # 方法B：直接解析 xlsx zip（单元格绑定图片 / cellImages）
    _extract_cell_images(file_bytes, all_images)

    print(f"[图片提取] 合计收集到 {len(all_images)} 张图片位置")
    for (r, c), img in sorted(all_images.items()):
        print(f"  → Excel行{r}, 列{c}(0-based), 尺寸{img.size}")

    # ── Step 3：按列归类为 photo / qr ────────────────────
    photo_by_row = {}
    qr_by_row    = {}
    unmatched    = []   # 列号未匹配到表头的图片

    for (row_1based, col_0based), pil_img in all_images.items():
        if photo_col_idx is not None and col_0based == photo_col_idx:
            photo_by_row[row_1based] = pil_img
        elif qr_col_idx is not None and col_0based == qr_col_idx:
            qr_by_row[row_1based] = pil_img
        else:
            unmatched.append((row_1based, col_0based, pil_img))

    # 兜底：列名未匹配时，按列号排序，最小列=照片，次小列=二维码
    if unmatched:
        cols_seen = sorted(set(c for _, c, _ in unmatched))
        print(f"[图片提取] 兜底分配，发现列号: {cols_seen}")
        col_to_type = {}
        if len(cols_seen) >= 1:
            col_to_type[cols_seen[0]] = "photo"
        if len(cols_seen) >= 2:
            col_to_type[cols_seen[1]] = "qr"
        for row_1based, col_0based, pil_img in unmatched:
            t = col_to_type.get(col_0based)
            if t == "photo":
                photo_by_row[row_1based] = pil_img
            elif t == "qr":
                qr_by_row[row_1based] = pil_img

    print(f"[图片提取] 最终：照片 {len(photo_by_row)} 张，二维码 {len(qr_by_row)} 张")
    return photo_by_row, qr_by_row


# ── 方法A：openpyxl 浮动图片 ─────────────────────────────

def _extract_floating(ws, all_images: dict):
    """
    从 openpyxl worksheet._images 提取浮动图片。
    anchor._from.row / col 均为 0-based。
    转换为 1-based 行号存入 all_images。
    """
    images = getattr(ws, '_images', [])
    print(f"[方法A] openpyxl._images 发现 {len(images)} 张浮动图片")

    for img_obj in images:
        # 读取图片二进制
        try:
            img_data = img_obj._data()
        except Exception:
            continue

        # 读取锚点
        anchor = img_obj.anchor
        try:
            row_0based = anchor._from.row   # 0-based
            col_0based = anchor._from.col   # 0-based
        except AttributeError:
            print(f"  [方法A] 跳过 AbsoluteAnchor 图片")
            continue

        row_1based = row_0based + 1   # 转为 1-based Excel 行号
        pil_img = _bytes_to_pil(img_data)
        if pil_img and (row_1based, col_0based) not in all_images:
            all_images[(row_1based, col_0based)] = pil_img
            print(f"  [方法A] 行{row_1based}, 列{col_0based} ✓")


# ── 方法B：直接解析 xlsx zip 包 ───────────────────────────

def _extract_cell_images(file_bytes: bytes, all_images: dict):
    """
    直接解析 xlsx（本质是 zip）内部 XML，提取：
      1. 传统 drawing（xl/drawings/drawingN.xml）中的图片锚点
      2. 新版单元格绑定图片（xl/cellImages/cellImageN.xml）

    这是兜底方案，能处理 openpyxl 读不到的情况。
    """
    try:
        zf = zipfile.ZipFile(io.BytesIO(file_bytes))
    except Exception as e:
        print(f"[方法B] 无法打开 zip: {e}")
        return

    names = zf.namelist()

    # ── B1：传统 drawing XML ──────────────────────────────
    drawing_files = [n for n in names if re.match(r'xl/drawings/drawing\d+\.xml$', n)]
    print(f"[方法B] 发现 drawing 文件: {drawing_files}")

    for drawing_path in drawing_files:
        # 对应的 rels 文件
        rels_path = drawing_path.replace('drawings/drawing', 'drawings/_rels/drawing') \
                                .replace('.xml', '.xml.rels')

        # 读取 rels，建立 rId -> 图片文件路径 映射
        rId_to_imgpath = {}
        if rels_path in names:
            rels_xml = zf.read(rels_path)
            rels_root = etree.fromstring(rels_xml)
            for rel in rels_root:
                rid    = rel.get('Id', '')
                target = rel.get('Target', '')
                rtype  = rel.get('Type', '')
                if 'image' in rtype.lower():
                    # target 是相对路径，如 ../media/image1.png
                    # 转为 zip 内绝对路径
                    base = '/'.join(drawing_path.split('/')[:-1])
                    img_zip_path = _resolve_path(base, target)
                    rId_to_imgpath[rid] = img_zip_path

        # 解析 drawing XML，提取锚点 + rId
        drawing_xml = zf.read(drawing_path)
        root = etree.fromstring(drawing_xml)
        ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }

        # 遍历所有锚点类型
        for anchor_tag in ('xdr:twoCellAnchor', 'xdr:oneCellAnchor'):
            for anchor_el in root.findall(anchor_tag, ns):
                from_el = anchor_el.find('xdr:from', ns)
                if from_el is None:
                    continue
                col_el = from_el.find('xdr:col', ns)
                row_el = from_el.find('xdr:row', ns)
                if col_el is None or row_el is None:
                    continue

                col_0based = int(col_el.text)   # 0-based
                row_0based = int(row_el.text)   # 0-based
                row_1based = row_0based + 1

                # 找图片 rId
                pic_el = anchor_el.find('.//xdr:pic', ns)
                if pic_el is None:
                    continue
                blip_el = pic_el.find('.//a:blip', ns)
                if blip_el is None:
                    continue
                rid = blip_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')

                img_zip_path = rId_to_imgpath.get(rid)
                if not img_zip_path or img_zip_path not in names:
                    print(f"  [方法B] rId={rid} 找不到图片文件")
                    continue

                img_data = zf.read(img_zip_path)
                pil_img  = _bytes_to_pil(img_data)
                if pil_img and (row_1based, col_0based) not in all_images:
                    all_images[(row_1based, col_0based)] = pil_img
                    print(f"  [方法B-drawing] 行{row_1based}, 列{col_0based} ✓ ({img_zip_path})")

    # ── B2：新版单元格绑定图片（cellImages）────────────────
    # 结构1：Excel / WPS 单文件清单 xl/cellimages.xml + DISPIMG("ID_xxx")
    _extract_dispimg_cell_images(zf, names, all_images)

    # 结构2：文件路径形如 xl/cellImages/cellImage1.xml
    cell_image_files = [n for n in names if re.match(r'xl/cellImages/cellImage\d+\.xml$', n)]
    print(f"[方法B] 发现 cellImage 文件: {cell_image_files}")

    for ci_path in cell_image_files:
        rels_path = ci_path.replace('cellImages/cellImage', 'cellImages/_rels/cellImage') \
                           .replace('.xml', '.xml.rels')

        rId_to_imgpath = {}
        if rels_path in names:
            rels_xml  = zf.read(rels_path)
            rels_root = etree.fromstring(rels_xml)
            for rel in rels_root:
                rid    = rel.get('Id', '')
                target = rel.get('Target', '')
                rtype  = rel.get('Type', '')
                if 'image' in rtype.lower():
                    base = '/'.join(ci_path.split('/')[:-1])
                    rId_to_imgpath[rid] = _resolve_path(base, target)

        ci_xml  = zf.read(ci_path)
        ci_root = etree.fromstring(ci_xml)

        # cellImage XML 结构：<etc:cellImage> 内含 <xdr:pic> 和位置信息
        # 位置通常在父级 sheet XML 的 <etc:cellImages> 节点里
        # 这里直接提取图片数据，行列从文件名序号推断（兜底）
        for blip_el in ci_root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
            rid = blip_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
            img_zip_path = rId_to_imgpath.get(rid)
            if not img_zip_path or img_zip_path not in names:
                continue
            img_data = zf.read(img_zip_path)
            pil_img  = _bytes_to_pil(img_data)
            if pil_img:
                print(f"  [方法B-cellImage] 提取到图片 {img_zip_path}，尺寸{pil_img.size}（行列待定）")
                # cellImage 的行列需要从 sheet XML 的 <etc:cellImages> 读取
                # 这里先存到特殊 key (-1, seq)，后续在 _resolve_cell_image_positions 中修正
                seq = len([k for k in all_images if k[0] == -1])
                all_images[(-1, seq)] = pil_img

    # 尝试从 sheet XML 补全 cellImage 的行列位置
    _resolve_cell_image_positions(zf, names, all_images)


def _extract_dispimg_cell_images(zf, names, all_images: dict):
    """
    解析 WPS / 新版 Excel 的 DISPIMG 图片结构：
      - 图片定义在 xl/cellimages.xml
      - 单元格内用 DISPIMG("ID_xxx", 1) 公式引用
    """
    cellimages_path = next((n for n in names if n.lower() == 'xl/cellimages.xml'), None)
    if not cellimages_path:
        return

    rels_path = 'xl/_rels/cellimages.xml.rels'
    if rels_path not in names:
        print("[方法B-dispimg] 缺少 cellimages.xml.rels")
        return

    pic_id_to_image = {}
    rid_to_imgpath = {}

    rels_root = ET.fromstring(zf.read(rels_path))
    rel_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    for rel in rels_root.findall('rel:Relationship', rel_ns):
        rid = rel.attrib.get('Id', '')
        target = rel.attrib.get('Target', '')
        rtype = rel.attrib.get('Type', '')
        if 'image' in rtype.lower():
            rid_to_imgpath[rid] = _resolve_path('xl', target)

    root = ET.fromstring(zf.read(cellimages_path))
    main_ns = {
        'etc': 'http://www.wps.cn/officeDocument/2017/etCustomData',
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    for cell_image in root.findall('etc:cellImage', main_ns):
        pic = cell_image.find('xdr:pic', main_ns)
        if pic is None:
            continue
        c_nv_pr = pic.find('.//xdr:cNvPr', main_ns)
        blip = pic.find('.//a:blip', main_ns)
        if c_nv_pr is None or blip is None:
            continue

        pic_id = c_nv_pr.attrib.get('name') or c_nv_pr.attrib.get('descr')
        rid = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
        img_zip_path = rid_to_imgpath.get(rid)
        if not pic_id or not img_zip_path or img_zip_path not in names:
            continue

        pil_img = _bytes_to_pil(zf.read(img_zip_path))
        if pil_img is not None:
            pic_id_to_image[pic_id] = pil_img
            print(f"  [方法B-dispimg] 图片ID={pic_id} ✓ ({img_zip_path})")

    if not pic_id_to_image:
        print("[方法B-dispimg] 未解析到图片清单")
        return

    sheet_files = [n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
    cell_ns = {'ws': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    for sheet_path in sheet_files:
        sheet_root = ET.fromstring(zf.read(sheet_path))
        for cell in sheet_root.findall('.//ws:c', cell_ns):
            cell_ref = cell.attrib.get('r', '')
            formula_node = cell.find('ws:f', cell_ns)
            pic_id = _parse_dispimg_formula(formula_node.text if formula_node is not None else '')
            if not pic_id or pic_id not in pic_id_to_image:
                continue

            match = re.match(r'([A-Z]+)(\d+)$', cell_ref)
            if not match:
                continue
            col_letters, row_text = match.groups()
            col_0based = _col_index(col_letters)
            row_1based = int(row_text)
            if (row_1based, col_0based) not in all_images:
                all_images[(row_1based, col_0based)] = pic_id_to_image[pic_id]
                print(f"  [方法B-dispimg] 行{row_1based}, 列{col_0based} ✓ (ID={pic_id})")


def _resolve_cell_image_positions(zf, names, all_images):
    """
    从 sheet1.xml 中读取 <etc:cellImages> 节点，
    把 all_images 里 row=-1 的占位图片补全行列信息。
    """
    # 找主 sheet 文件
    sheet_files = [n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
    if not sheet_files:
        return

    sheet_xml  = zf.read(sheet_files[0])
    sheet_root = etree.fromstring(sheet_xml)

    # 命名空间可能不同，用通配符搜索 cellImages
    cell_images_nodes = sheet_root.findall(
        './/{http://schemas.microsoft.com/office/spreadsheetml/2020/11/main}cellImages'
    )
    if not cell_images_nodes:
        # 尝试其他命名空间
        for el in sheet_root.iter():
            if el.tag.endswith('}cellImages') or el.tag == 'cellImages':
                cell_images_nodes.append(el)
                break

    if not cell_images_nodes:
        print("[方法B] 未找到 cellImages 节点，cellImage 行列无法确定")
        return

    # 收集待补全的图片（row=-1）
    pending = {k: v for k, v in list(all_images.items()) if k[0] == -1}
    if not pending:
        return

    pending_list = sorted(pending.items())   # 按 seq 排序

    # 解析每个 cellImage 节点，获取行列
    ci_idx = 0
    for ci_node in cell_images_nodes[0]:
        if ci_idx >= len(pending_list):
            break
        # 找 <xdr:from> 或 <from> 子节点
        from_el = ci_node.find('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from')
        if from_el is None:
            ci_idx += 1
            continue
        col_el = from_el.find('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col')
        row_el = from_el.find('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row')
        if col_el is None or row_el is None:
            ci_idx += 1
            continue

        col_0based = int(col_el.text)
        row_0based = int(row_el.text)
        row_1based = row_0based + 1

        old_key, pil_img = pending_list[ci_idx]
        del all_images[old_key]
        if (row_1based, col_0based) not in all_images:
            all_images[(row_1based, col_0based)] = pil_img
            print(f"  [方法B-cellImage补全] 行{row_1based}, 列{col_0based} ✓")
        ci_idx += 1


def _resolve_path(base: str, target: str) -> str:
    """解析 XML rels 中的相对路径为 zip 内绝对路径"""
    if target.startswith('/'):
        return target.lstrip('/')
    parts = base.split('/')
    for seg in target.split('/'):
        if seg == '..':
            if parts:
                parts.pop()
        elif seg != '.':
            parts.append(seg)
    return '/'.join(parts)
