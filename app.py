"""
app.py —— Streamlit 网页入口
店长岗位挂卡 A4 打印工具
运行方式：streamlit run app.py
"""

import streamlit as st
import pandas as pd
import io
import zipfile
from PIL import Image

import main as card_engine
from excel_image_extractor import extract_images_from_excel

# ── 页面配置 ──────────────────────────────────────────────
st.set_page_config(
    page_title="挂卡打印工具",
    page_icon="🪪",
    layout="wide",
)

st.title("🪪 店长岗位挂卡 A4 打印工具")
st.caption("上传 Excel → 生成预览 → 下载打印文件")

# ── 侧边栏：使用说明 ──────────────────────────────────────
with st.sidebar:
    st.header("📋 使用说明")
    st.markdown("""
1. 在 Excel 中准备员工信息，列名包含：
   - `姓名`
   - `职位`
   - `工号`
   - `照片`（在该列单元格内直接插入图片）
   - `二维码`（在该列单元格内直接插入图片）

2. 保存为 `.xlsx` 格式

3. 上传 Excel 文件

4. 点击「生成预览」

5. 下载 PDF 或 PNG 用于打印
    """)
    st.divider()
    st.markdown("每张挂卡：**96mm × 55mm**")
    st.markdown("每页排列：**2列 × 5行 = 10张**")
    st.divider()
    st.markdown("""
**照片插入方式：**
在 Excel 中选中"照片"列的单元格，
菜单 → 插入 → 图片 → 放置在单元格中
    """)

# ── 主界面 ────────────────────────────────────────────────
st.subheader("① 上传 Excel 文件")
st.info("Excel 中的照片和二维码请直接插入到对应单元格内，无需额外上传图片文件。")

uploaded_excel = st.file_uploader(
    "选择包含员工信息的 Excel 文件（.xlsx）",
    type=["xlsx"],
    help="列名需包含：姓名、职位、工号、照片（内嵌图片）、二维码（内嵌图片）"
)

# ── 解析 Excel ────────────────────────────────────────────
df          = None
photo_by_row = {}
qr_by_row    = {}

if uploaded_excel:
    # 读取文件字节（需要读两次：pandas + openpyxl）
    file_bytes = uploaded_excel.read()

    # 1. pandas 读取文字数据
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), dtype=str).fillna("")
        st.success(f"读取成功，共 {len(df)} 条员工记录")

        with st.expander("查看文字数据预览", expanded=False):
            st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"Excel 文字数据读取失败：{e}")

    # 2. openpyxl 提取内嵌图片
    try:
        photo_by_row, qr_by_row = extract_images_from_excel(file_bytes)

        # 调试信息展示
        with st.expander("📊 图片识别结果", expanded=True):
            col_a, col_b = st.columns(2)
            col_a.metric("识别到照片", f"{len(photo_by_row)} 张")
            col_b.metric("识别到二维码", f"{len(qr_by_row)} 张")

            if df is not None:
                # 找出哪些行缺少图片（数据从第2行起，对应 df 的第0行）
                missing_photo = []
                missing_qr    = []
                for i in range(len(df)):
                    excel_row = i + 2   # Excel 行号（1-based，第1行是表头）
                    name = df.iloc[i].get("姓名", f"第{excel_row}行")
                    if excel_row not in photo_by_row:
                        missing_photo.append(f"行{excel_row}（{name}）")
                    if excel_row not in qr_by_row:
                        missing_qr.append(f"行{excel_row}（{name}）")

                if missing_photo:
                    st.warning(f"以下行未找到照片，将使用占位图：{', '.join(missing_photo)}")
                if missing_qr:
                    st.warning(f"以下行未找到二维码，将使用占位图：{', '.join(missing_qr)}")
                if not missing_photo and not missing_qr:
                    st.success("所有行均已匹配到照片和二维码 ✓")

    except Exception as e:
        st.warning(f"内嵌图片提取失败（将使用占位图）：{e}")

# ── 生成按钮 ──────────────────────────────────────────────
st.divider()
generate_btn = st.button(
    "🖨️ 生成预览",
    type="primary",
    disabled=(df is None)
)

# ── 生成逻辑 ──────────────────────────────────────────────
if generate_btn and df is not None:

    staff_list = []
    for i, row in df.iterrows():
        excel_row = i + 2   # Excel 行号（表头占第1行）

        # 从提取结果中获取图片，没有则传 None（main.py 会用占位图）
        photo_src = photo_by_row.get(excel_row, None)
        qr_src    = qr_by_row.get(excel_row, None)

        staff_list.append({
            "name":     row.get("姓名", ""),
            "title":    row.get("职务", ""),
            "staff_id": row.get("工号", ""),
            "photo":    photo_src,   # PIL Image 或 None
            "qr_code":  qr_src,      # PIL Image 或 None
        })

    with st.spinner("正在生成挂卡，请稍候..."):
        try:
            pages = card_engine.build_a4_pages(staff_list)
            st.session_state["pages"] = pages
            st.success(f"生成完成，共 {len(pages)} 页 A4")
        except Exception as e:
            st.error(f"生成失败：{e}")
            st.stop()

# ── 预览 & 下载 ───────────────────────────────────────────
if "pages" in st.session_state and st.session_state["pages"]:
    pages = st.session_state["pages"]

    st.subheader("② 预览")

    total_pages = len(pages)
    if total_pages == 1:
        st.caption("共 1 页")
        page_idx = 0
    else:
        page_idx = st.slider(
            "选择预览页",
            min_value=1,
            max_value=total_pages,
            value=1,
            step=1,
            format="第 %d 页"
        ) - 1

    # 缩小预览图以适应屏幕
    preview = pages[page_idx].copy()
    preview.thumbnail((900, 1200), Image.LANCZOS)
    st.image(preview, caption=f"第 {page_idx + 1} 页 / 共 {total_pages} 页")

    # ── 下载 ──────────────────────────────────────────────
    st.subheader("③ 下载打印文件")
    dl_col1, dl_col2 = st.columns(2)

    with dl_col1:
        pdf_buf = io.BytesIO()
        first = pages[0].convert("RGB")
        rest  = [p.convert("RGB") for p in pages[1:]]
        first.save(pdf_buf, format="PDF", save_all=True,
                   append_images=rest, resolution=card_engine.DPI)
        pdf_buf.seek(0)
        st.download_button(
            label="� 下载 PDF（推荐打印）",
            data=pdf_buf,
            file_name="挂卡打印.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with dl_col2:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, page in enumerate(pages):
                img_buf = io.BytesIO()
                page.save(img_buf, format="PNG")
                zf.writestr(f"page_{i+1:02d}.png", img_buf.getvalue())
        zip_buf.seek(0)
        st.download_button(
            label="🗜️ 下载 PNG 压缩包",
            data=zip_buf,
            file_name="挂卡打印.zip",
            mime="application/zip",
            use_container_width=True,
        )

    # ── HTML 下载（position:absolute 精确还原版）──────────────
    st.divider()
    st.subheader("④ HTML 打印模板（精确还原 SVG 布局）")

    import html_render as hr
    import base64 as _b64

    def _pil_to_b64(img):
        if img is None:
            return ""
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=90)
        return _b64.b64encode(buf.getvalue()).decode()

    html_cards = []
    for i, s in enumerate(staff_list[:10]):
        html_cards.append({
            "name":      s["name"],
            "title":     s["title"],
            "staff_id":  s["staff_id"],
            "star_count": 5,
            "photo_b64": _pil_to_b64(s.get("photo")),
            "qr_b64":    _pil_to_b64(s.get("qr_code")),
        })

    html_content = hr.build_a4_html(html_cards)
    html_bytes = html_content.encode("utf-8")

    st.download_button(
        label="⬇️ 下载 HTML 模板（浏览器打印 / 精确排版）",
        data=html_bytes,
        file_name="高济药房员工工牌.html",
        mime="text/html",
        use_container_width=True,
    )

    with st.expander("🌐 预览 HTML 模板（缩小版）"):
        import streamlit.components.v1 as components
        # 单卡 HTML 预览
        preview_html = hr.build_card_html(
            name=staff_list[0]["name"] if staff_list else "张三",
            title=staff_list[0]["title"] if staff_list else "店长",
            staff_id=staff_list[0]["staff_id"] if staff_list else "GJ-001",
            star_count=5,
            photo_b64=_pil_to_b64(staff_list[0].get("photo")) if staff_list else "",
            qr_b64=_pil_to_b64(staff_list[0].get("qr_code")) if staff_list else "",
        )
        components.html(preview_html, height=600, scrolling=True)

# ── 示例数据预览 ───────────────────────────────────────────────
st.divider()
with st.expander("🎴 示例数据预览（无需上传文件）", expanded=False):
    import html_render as hr
    import streamlit.components.v1 as components

    sample_html = hr.build_card_html(
        name="李美华",
        title="店长",
        staff_id="GJ-2024001",
        star_count=5,
    )
    components.html(sample_html, height=700, scrolling=False)