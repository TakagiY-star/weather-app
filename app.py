"""
気象レポートPDF 降水量ハイライター
複数PDF同時アップロード対応・PDF/JPG出力選択対応版
"""

import streamlit as st
import tempfile
import os
import io
import zipfile
from itertools import groupby

import pdfplumber
import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c

st.set_page_config(
    page_title="降水量ハイライター",
    page_icon="🌧️",
    layout="centered"
)

st.markdown("""
<style>
    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #00B0F0, #0097d6);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 14px;
        font-size: 16px;
        font-weight: 600;
    }
    .stButton > button:hover { opacity: 0.9; }
    .result-box {
        background: #e8f9f4;
        border: 1px solid #c3f0e4;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin: 16px 0;
    }
    .file-row {
        background: #f4f9fd;
        border: 1px solid #cde4f5;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 6px 0;
        font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)


# ---------- PDF処理ロジック ----------

def extract_rainfall_groups(pdf_path):
    """pdfplumberでセル境界と降水量データを正確に取得してグループ化"""
    page_groups = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words()
            mmh_tops = [w['top'] for w in words if w['text'] == 'mm/h']
            if not mmh_tops:
                continue

            tset = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
            tables = page.find_tables(tset)
            if not tables:
                continue
            all_cells = tables[0].cells

            rainfall_boxes = []
            for w in words:
                in_row = any(abs(w['top'] - mt) <= 8 for mt in mmh_tops)
                if not in_row:
                    continue
                try:
                    val = float(w['text'])
                    if val <= 0:
                        continue
                    cx = (w['x0'] + w['x1']) / 2
                    cy = (w['top'] + w['bottom']) / 2
                    for cell in all_cells:
                        if cell[0] <= cx <= cell[2] and cell[1] <= cy <= cell[3]:
                            rainfall_boxes.append({
                                'value': val,
                                'x0': cell[0], 'top': cell[1],
                                'x1': cell[2], 'bottom': cell[3],
                                'row_top': round(cell[1], 1)
                            })
                            break
                except ValueError:
                    pass

            groups = []
            boxes_sorted = sorted(rainfall_boxes, key=lambda c: (c['row_top'], c['x0']))
            for _, row_boxes in groupby(boxes_sorted, key=lambda c: c['row_top']):
                row_boxes = list(row_boxes)
                current = [row_boxes[0]]
                for box in row_boxes[1:]:
                    if box['x0'] - current[-1]['x1'] <= 25:
                        current.append(box)
                    else:
                        groups.append(current)
                        current = [box]
                groups.append(current)

            page_groups.append({
                'page_num': page_num,
                'page_height': page.height,
                'groups': groups
            })

    return page_groups


def draw_highlights(input_path, output_path, page_groups):
    """水色の枠線＋薄い水色塗りつぶし（Multiplyブレンドで数字を透過）"""
    doc = pdfium.PdfDocument(input_path)
    SKY_R, SKY_G, SKY_B = 0, 176, 240

    for pg in page_groups:
        page = doc[pg['page_num']]
        pdf_h = page.get_height()
        pl_h = pg['page_height']

        for group in pg['groups']:
            gx0 = min(c['x0'] for c in group)
            gx1 = max(c['x1'] for c in group)
            gtop = min(c['top'] for c in group)
            gbottom = max(c['bottom'] for c in group)

            pdf_x0 = gx0
            pdf_y0 = pdf_h - gbottom * (pdf_h / pl_h)
            pdf_y1 = pdf_h - gtop * (pdf_h / pl_h)

            rect = pdfium_c.FPDFPageObj_CreateNewRect(
                pdf_x0, pdf_y0, gx1 - gx0, pdf_y1 - pdf_y0
            )
            pdfium_c.FPDFPageObj_SetFillColor(rect, SKY_R, SKY_G, SKY_B, 40)
            pdfium_c.FPDFPageObj_SetStrokeColor(rect, SKY_R, SKY_G, SKY_B, 255)
            pdfium_c.FPDFPageObj_SetStrokeWidth(rect, 1.5)
            pdfium_c.FPDFPath_SetDrawMode(rect, 1, 1)
            pdfium_c.FPDFPageObj_SetBlendMode(rect, b"Multiply")
            pdfium_c.FPDFPage_InsertObject(page.raw, rect)

        pdfium_c.FPDFPage_GenerateContent(page.raw)

    doc.save(output_path)


def pdf_to_jpg_bytes(pdf_path, dpi=150):
    """PDFの各ページをJPG画像に変換してバイト列のリストで返す"""
    doc = pdfium.PdfDocument(pdf_path)
    scale = dpi / 72
    results = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        bitmap = page.render(scale=scale, rotation=0)
        pil_image = bitmap.to_pil()

        buf = io.BytesIO()
        pil_image.save(buf, format="JPEG", quality=90)
        results.append((page_num, buf.getvalue()))

    return results


def process_single_pdf(uploaded_file, tmpdir, output_format, dpi=150):
    """1つのPDFを処理して出力データのリストを返す"""
    input_path = os.path.join(tmpdir, "input.pdf")
    highlighted_path = os.path.join(tmpdir, "highlighted.pdf")
    base_name = os.path.splitext(uploaded_file.name)[0]

    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    page_groups = extract_rainfall_groups(input_path)
    total_cells = sum(len(g) for pg in page_groups for g in pg['groups'])
    total_groups = sum(len(pg['groups']) for pg in page_groups)

    if total_cells == 0:
        return None, 0, 0

    draw_highlights(input_path, highlighted_path, page_groups)

    outputs = []

    if output_format == "PDF":
        with open(highlighted_path, "rb") as f:
            outputs.append({
                "name": f"{base_name}_ハイライト済.pdf",
                "bytes": f.read(),
                "mime": "application/pdf"
            })
    else:  # JPG
        pages = pdf_to_jpg_bytes(highlighted_path, dpi=dpi)
        if len(pages) == 1:
            _, jpg_bytes = pages[0]
            outputs.append({
                "name": f"{base_name}_ハイライト済.jpg",
                "bytes": jpg_bytes,
                "mime": "image/jpeg"
            })
        else:
            for page_num, jpg_bytes in pages:
                outputs.append({
                    "name": f"{base_name}_ハイライト済_p{page_num + 1}.jpg",
                    "bytes": jpg_bytes,
                    "mime": "image/jpeg"
                })

    return outputs, total_cells, total_groups


# ---------- 画面 ----------

st.title("🌧️ 降水量ハイライター")
st.caption("気象レポートのPDFをアップロードすると、降水量が0より大きいセルを水色の枠と薄い水色でハイライトします。")

st.divider()

uploaded_files = st.file_uploader(
    "気象レポートのPDFを選択してください（複数選択可）",
    type=["pdf"],
    accept_multiple_files=True,
    help="Ctrl（Mac: Command）を押しながらクリックで複数選択できます"
)

if uploaded_files:
    st.markdown(f"**{len(uploaded_files)} 件のファイルが選択されています**")
    for f in uploaded_files:
        st.markdown(f"""
        <div class="file-row">
            📄 {f.name}&nbsp;&nbsp;<span style="color:#6b8aad">{f.size / 1024:.0f} KB</span>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    # 出力形式の選択
    st.markdown("**出力形式を選択してください**")
    output_format = st.radio(
        label="出力形式",
        options=["PDF", "JPG"],
        horizontal=True,
        label_visibility="collapsed"
    )

    if output_format == "JPG":
        dpi = st.select_slider(
            "画質（DPI）",
            options=[72, 96, 150, 200, 300],
            value=150,
            help="数値が大きいほど高画質・ファイルサイズ大。通常は150で十分です"
        )
    else:
        dpi = 150

    st.markdown("")

    if st.button("⚡ ハイライト処理を実行"):
        all_outputs = []
        errors = []

        progress_bar = st.progress(0, text="処理中...")

        with tempfile.TemporaryDirectory() as tmpdir:
            for i, uploaded_file in enumerate(uploaded_files):
                progress_bar.progress(
                    i / len(uploaded_files),
                    text=f"処理中 ({i+1}/{len(uploaded_files)}): {uploaded_file.name}"
                )
                try:
                    outputs, total_cells, total_groups = process_single_pdf(
                        uploaded_file, tmpdir, output_format, dpi
                    )
                    if outputs is None:
                        errors.append(f"{uploaded_file.name}：降水量データが見つかりませんでした")
                    else:
                        for out in outputs:
                            all_outputs.append({
                                **out,
                                "cells": total_cells,
                                "groups": total_groups,
                                "source": uploaded_file.name
                            })
                except Exception as e:
                    errors.append(f"{uploaded_file.name}：エラー ({e})")

            progress_bar.progress(1.0, text="完了！")

            for err in errors:
                st.error(f"⚠️ {err}")

            if all_outputs:
                total_files = len(uploaded_files) - len(errors)
                st.markdown(f"""
                <div class="result-box">
                    <h3 style="color:#00c896; margin:0 0 8px">✓ {total_files} 件の処理が完了しました！</h3>
                </div>
                """, unsafe_allow_html=True)

                # 1ファイル・1ページ → 直接ダウンロード
                if len(all_outputs) == 1:
                    out = all_outputs[0]
                    ext = "PDF" if output_format == "PDF" else "JPG"
                    st.download_button(
                        label=f"📥 {out['name']} をダウンロード（{ext}）",
                        data=out['bytes'],
                        file_name=out['name'],
                        mime=out['mime']
                    )

                # 複数ファイル or 複数ページJPG → ZIPでまとめてダウンロード
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for out in all_outputs:
                            zf.writestr(out['name'], out['bytes'])
                    zip_buf.seek(0)

                    ext_label = "PDF" if output_format == "PDF" else "JPG"
                    st.download_button(
                        label=f"📦 {len(all_outputs)} 件をZIPでまとめてダウンロード（{ext_label}）",
                        data=zip_buf.getvalue(),
                        file_name="ハイライト済み_一括.zip",
                        mime="application/zip"
                    )

                    # 個別ダウンロード
                    with st.expander("📄 個別にダウンロードする"):
                        for out in all_outputs:
                            col1, col2 = st.columns([3, 1])
                            with col1:
                                st.caption(f"{out['name']}")
                            with col2:
                                st.download_button(
                                    label="DL",
                                    data=out['bytes'],
                                    file_name=out['name'],
                                    mime=out['mime'],
                                    key=out['name']
                                )

st.divider()
st.caption("📌 使い方：PDFを選択（複数可）→ 出力形式を選択 → 実行 → ダウンロード")
