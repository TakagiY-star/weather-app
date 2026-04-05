"""
気象レポートPDF 降水量ハイライター
- 降水量セルをハイライト
- 1〜4日目の1時間予報をJPGとしてトリミング出力
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
from PIL import Image, ImageChops
import numpy as np

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


# ============================================================
# 処理ロジック
# ============================================================

def extract_rainfall_groups(pdf_path):
    """降水量 > 0 のセルをグループ化して返す"""
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
    """水色の枠線＋薄い水色塗りつぶしでハイライト"""
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


def find_forecast_crop_top(pdf_path):
    """
    pdfplumberで「現在から24時」または「日付」テキストの位置を検出し、
    1時間予報エリアのクロップ開始位置(pt)を返す
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()
        candidates = []
        for w in words:
            if '現在から' in w['text'] or w['text'] == '日付':
                candidates.append(w['top'])
        if candidates:
            return min(candidates) - 3
    return 210


def find_forecast_crop_bottom(pdf_path):
    """
    4日目の最終行（風のm/s行）下端をpdfplumberで検出し、
    クロップ終了位置(pt)を返す
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()
        ms_tops = [w['top'] for w in words if w['text'] == 'm/s']
        if ms_tops:
            last_ms_top = max(ms_tops)
            last_ms_bottom = max(
                w['bottom'] for w in words
                if w['text'] == 'm/s' and w['top'] == last_ms_top
            )
            return last_ms_bottom + 5
    return 780


def trim_whitespace(img, margin=5):
    """白い余白を自動トリミング"""
    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if bbox is None:
        return img
    return img.crop((
        max(0, bbox[0] - margin),
        max(0, bbox[1] - margin),
        min(img.width, bbox[2] + margin),
        min(img.height, bbox[3] + margin)
    ))


def crop_forecast_jpg(pdf_path, dpi=300):
    """
    ハイライト済みPDFから1〜4日目の1時間予報エリアを
    動的に検出してトリミングし、JPG画像バイト列を返す
    """
    crop_top_pt = find_forecast_crop_top(pdf_path)
    crop_bottom_pt = find_forecast_crop_bottom(pdf_path)

    doc = pdfium.PdfDocument(pdf_path)
    page = doc[0]
    scale = dpi / 72

    bitmap = page.render(scale=scale)
    img = bitmap.to_pil()

    px_top = int(crop_top_pt * scale)
    px_bottom = int(crop_bottom_pt * scale)
    px_bottom = min(px_bottom, img.height)

    cropped = img.crop((0, px_top, img.width, px_bottom))
    trimmed = trim_whitespace(cropped, margin=5)

    buf = io.BytesIO()
    trimmed.save(buf, format="JPEG", quality=95)
    return buf.getvalue()


def process_pdf(uploaded_file, tmpdir, dpi):
    """PDFを処理してハイライト済みPDFと1時間予報JPGを返す"""
    input_path = os.path.join(tmpdir, "input.pdf")
    highlighted_path = os.path.join(tmpdir, "highlighted.pdf")
    base_name = os.path.splitext(uploaded_file.name)[0]

    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    page_groups = extract_rainfall_groups(input_path)
    total_cells = sum(len(g) for pg in page_groups for g in pg['groups'])
    total_groups = sum(len(pg['groups']) for pg in page_groups)

    if total_cells == 0:
        return None, None, 0, 0

    draw_highlights(input_path, highlighted_path, page_groups)

    with open(highlighted_path, "rb") as f:
        highlighted_bytes = f.read()

    jpg_bytes = crop_forecast_jpg(highlighted_path, dpi=dpi)

    return highlighted_bytes, jpg_bytes, total_cells, total_groups


# ============================================================
# 画面
# ============================================================

st.title("🌧️ 降水量ハイライター")
st.caption("気象レポートPDFの降水量をハイライトし、1〜4日目の1時間予報をJPGで切り抜きます。")

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

    dpi = st.select_slider(
        "JPG画質（DPI）",
        options=[150, 200, 300],
        value=300,
        help="300が最高画質です"
    )

    if st.button("⚡ ハイライト＆切り抜きを実行"):
        results = []
        errors = []
        progress_bar = st.progress(0, text="処理中...")

        with tempfile.TemporaryDirectory() as tmpdir:
            for i, uploaded_file in enumerate(uploaded_files):
                progress_bar.progress(
                    i / len(uploaded_files),
                    text=f"処理中 ({i+1}/{len(uploaded_files)}): {uploaded_file.name}"
                )
                try:
                    highlighted_bytes, jpg_bytes, total_cells, total_groups = process_pdf(
                        uploaded_file, tmpdir, dpi
                    )
                    if highlighted_bytes is None:
                        errors.append(f"{uploaded_file.name}：降水量データが見つかりませんでした")
                    else:
                        base_name = os.path.splitext(uploaded_file.name)[0]
                        results.append({
                            "source": uploaded_file.name,
                            "base_name": base_name,
                            "highlighted_bytes": highlighted_bytes,
                            "jpg_bytes": jpg_bytes,
                            "cells": total_cells,
                            "groups": total_groups
                        })
                except Exception as e:
                    errors.append(f"{uploaded_file.name}：エラー ({e})")

            progress_bar.progress(1.0, text="完了！")

        for err in errors:
            st.error(f"⚠️ {err}")

        if results:
            total = len(results)
            st.markdown(f"""
            <div class="result-box">
                <h3 style="color:#00c896; margin:0 0 8px">✓ {total} 件の処理が完了しました！</h3>
                <p style="color:#6b8aad; margin:0">合計 {sum(r['cells'] for r in results)} セルをハイライト</p>
            </div>
            """, unsafe_allow_html=True)

            if len(results) == 1:
                r = results[0]
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 ハイライト済PDF",
                        data=r["highlighted_bytes"],
                        file_name=f"{r['base_name']}_ハイライト済.pdf",
                        mime="application/pdf"
                    )
                with col2:
                    st.download_button(
                        label="📥 1時間予報JPG",
                        data=r["jpg_bytes"],
                        file_name=f"{r['base_name']}_1時間予報.jpg",
                        mime="image/jpeg"
                    )
            else:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for r in results:
                        zf.writestr(f"{r['base_name']}_ハイライト済.pdf", r["highlighted_bytes"])
                        zf.writestr(f"{r['base_name']}_1時間予報.jpg", r["jpg_bytes"])
                zip_buf.seek(0)

                st.download_button(
                    label=f"📦 {total} 件をZIPでまとめてダウンロード",
                    data=zip_buf.getvalue(),
                    file_name="気象レポート_処理済み.zip",
                    mime="application/zip"
                )

                with st.expander("📄 個別にダウンロードする"):
                    for r in results:
                        st.caption(r["source"])
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                label="PDF",
                                data=r["highlighted_bytes"],
                                file_name=f"{r['base_name']}_ハイライト済.pdf",
                                mime="application/pdf",
                                key=f"pdf_{r['source']}"
                            )
                        with col2:
                            st.download_button(
                                label="JPG",
                                data=r["jpg_bytes"],
                                file_name=f"{r['base_name']}_1時間予報.jpg",
                                mime="image/jpeg",
                                key=f"jpg_{r['source']}"
                            )

st.divider()
st.caption("使い方：PDFをアップロード -> 実行 -> ハイライト済PDF と 1時間予報JPG をダウンロード")
