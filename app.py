"""
気象レポートPDF 降水量ハイライター
複数PDF同時アップロード対応版
"""

import streamlit as st
import tempfile
import os
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
        display: flex;
        align-items: center;
        gap: 10px;
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


def process_single_pdf(uploaded_file, tmpdir):
    """1つのPDFを処理して出力バイト列と統計を返す"""
    input_path = os.path.join(tmpdir, "input.pdf")
    output_path = os.path.join(tmpdir, uploaded_file.name)

    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    page_groups = extract_rainfall_groups(input_path)
    total_cells = sum(len(g) for pg in page_groups for g in pg['groups'])
    total_groups = sum(len(pg['groups']) for pg in page_groups)

    if total_cells == 0:
        return None, 0, 0

    draw_highlights(input_path, output_path, page_groups)

    with open(output_path, "rb") as f:
        output_bytes = f.read()

    return output_bytes, total_cells, total_groups


# ---------- 画面 ----------

st.title("🌧️ 降水量ハイライター")
st.caption("気象レポートのPDFをアップロードすると、降水量が0より大きいセルを水色の枠と薄い水色でハイライトします。複数ファイルの同時処理に対応しています。")

st.divider()

uploaded_files = st.file_uploader(
    "気象レポートのPDFを選択してください（複数選択可）",
    type=["pdf"],
    accept_multiple_files=True,
    help="Ctrl（Mac: Command）を押しながらクリックで複数選択できます"
)

if uploaded_files:
    # アップロード済みファイル一覧
    st.markdown(f"**{len(uploaded_files)} 件のファイルが選択されています**")
    for f in uploaded_files:
        st.markdown(f"""
        <div class="file-row">
            📄 {f.name} <span style="color:#6b8aad; margin-left:auto">{f.size / 1024:.0f} KB</span>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    if st.button("⚡ まとめてハイライト処理を実行"):
        results = []
        errors = []

        progress_bar = st.progress(0, text="処理中...")

        with tempfile.TemporaryDirectory() as tmpdir:
            for i, uploaded_file in enumerate(uploaded_files):
                progress_bar.progress(
                    (i) / len(uploaded_files),
                    text=f"処理中 ({i+1}/{len(uploaded_files)}): {uploaded_file.name}"
                )

                try:
                    output_bytes, total_cells, total_groups = process_single_pdf(
                        uploaded_file, tmpdir
                    )
                    if output_bytes is None:
                        errors.append(f"{uploaded_file.name}：降水量データが見つかりませんでした")
                    else:
                        out_name = uploaded_file.name.replace(".pdf", "_ハイライト済.pdf")
                        results.append({
                            'name': uploaded_file.name,
                            'out_name': out_name,
                            'bytes': output_bytes,
                            'cells': total_cells,
                            'groups': total_groups
                        })
                except Exception as e:
                    errors.append(f"{uploaded_file.name}：エラー ({e})")

            progress_bar.progress(1.0, text="完了！")

            # エラー表示
            for err in errors:
                st.error(f"⚠️ {err}")

            if results:
                total_files = len(results)

                st.markdown(f"""
                <div class="result-box">
                    <h3 style="color:#00c896; margin:0 0 8px">✓ {total_files} 件の処理が完了しました！</h3>
                    <p style="color:#6b8aad; margin:0">合計 {sum(r['cells'] for r in results)} セルをハイライト</p>
                </div>
                """, unsafe_allow_html=True)

                # ファイルが1つの場合：そのままダウンロード
                if total_files == 1:
                    r = results[0]
                    st.download_button(
                        label=f"📥 {r['out_name']} をダウンロード",
                        data=r['bytes'],
                        file_name=r['out_name'],
                        mime="application/pdf"
                    )

                # 複数ファイルの場合：ZIP でまとめてダウンロード
                else:
                    zip_path = os.path.join(tmpdir, "ハイライト済み_一括.zip")
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for r in results:
                            zf.writestr(r['out_name'], r['bytes'])

                    with open(zip_path, "rb") as f:
                        zip_bytes = f.read()

                    st.download_button(
                        label=f"📦 {total_files} 件をZIPでまとめてダウンロード",
                        data=zip_bytes,
                        file_name="ハイライト済み_一括.zip",
                        mime="application/zip"
                    )

                    # 個別ダウンロードも表示
                    with st.expander("📄 個別にダウンロードする"):
                        for r in results:
                            col1, col2 = st.columns([3, 1])
                            with col1:
                                st.caption(f"{r['out_name']}（{r['cells']}セル / {r['groups']}グループ）")
                            with col2:
                                st.download_button(
                                    label="DL",
                                    data=r['bytes'],
                                    file_name=r['out_name'],
                                    mime="application/pdf",
                                    key=r['name']
                                )

st.divider()
st.caption("📌 使い方：PDFを選択（複数可）→ 実行ボタン → ダウンロード")
