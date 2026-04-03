"""
気象レポートPDF 降水量ハイライター
Streamlit Cloud用 - このファイル1つをGitHubに置くだけで動きます
"""

import streamlit as st
import tempfile
import os
from itertools import groupby

import pdfplumber
import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c

# ページ設定
st.set_page_config(
    page_title="降水量ハイライター",
    page_icon="🌧️",
    layout="centered"
)

# スタイル
st.markdown("""
<style>
    .main { max-width: 700px; margin: 0 auto; }
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
</style>
""", unsafe_allow_html=True)


# ---------- PDF処理ロジック ----------

def extract_rainfall_groups(pdf_path):
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
            pdfium_c.FPDFPageObj_SetStrokeColor(rect, SKY_R, SKY_G, SKY_B, 255)
            pdfium_c.FPDFPageObj_SetStrokeWidth(rect, 1.5)
            pdfium_c.FPDFPath_SetDrawMode(rect, 0, 1)
            pdfium_c.FPDFPage_InsertObject(page.raw, rect)

        pdfium_c.FPDFPage_GenerateContent(page.raw)

    doc.save(output_path)


# ---------- 画面 ----------

st.title("🌧️ 降水量ハイライター")
st.caption("気象レポートのPDFをアップロードすると、降水量が0より大きいセルを水色の枠で自動ハイライトします。")

st.divider()

uploaded_file = st.file_uploader(
    "気象レポートのPDFを選択してください",
    type=["pdf"],
    help="複数ページのPDFにも対応しています"
)

if uploaded_file:
    st.success(f"✅ {uploaded_file.name}（{uploaded_file.size / 1024:.0f} KB）")

    if st.button("⚡ ハイライト処理を実行"):
        with st.spinner("処理中です。しばらくお待ちください…"):
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, "input.pdf")
                output_path = os.path.join(tmpdir, "output.pdf")

                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                try:
                    page_groups = extract_rainfall_groups(input_path)
                    total_cells = sum(len(g) for pg in page_groups for g in pg['groups'])
                    total_groups = sum(len(pg['groups']) for pg in page_groups)

                    if total_cells == 0:
                        st.error("降水量データが見つかりませんでした。気象レポートのPDFか確認してください。")
                    else:
                        draw_highlights(input_path, output_path, page_groups)

                        with open(output_path, "rb") as f:
                            output_bytes = f.read()

                        st.markdown(f"""
                        <div class="result-box">
                            <h3 style="color:#00c896; margin:0 0 8px">✓ 処理完了！</h3>
                            <p style="color:#6b8aad; margin:0">{total_cells} セル検出 / {total_groups} グループでハイライト</p>
                        </div>
                        """, unsafe_allow_html=True)

                        out_name = uploaded_file.name.replace(".pdf", "_ハイライト済.pdf")
                        st.download_button(
                            label="📥 ハイライト済PDFをダウンロード",
                            data=output_bytes,
                            file_name=out_name,
                            mime="application/pdf"
                        )

                except Exception as e:
                    st.error(f"処理中にエラーが発生しました: {e}")

st.divider()
st.caption("📌 使い方：PDFを選択 → 実行ボタン → ダウンロード")
