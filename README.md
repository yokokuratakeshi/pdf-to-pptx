[app.py](https://github.com/user-attachments/files/25431333/app.py)
"""
PDF → PowerPoint 変換ツール
・画像モード: 各ページを高品質な画像としてスライドに変換
・編集モード: テキストを編集可能なテキストボックス、画像を独立した移動可能オブジェクトに分離
"""

import streamlit as st
import fitz  # pymupdf
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import os

# ========== ページ設定 ==========
st.set_page_config(
    page_title="PDF → PowerPoint 変換",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ========== カスタムCSS ==========
st.markdown("""
<style>
    .main-header { text-align: center; padding: 1.5rem 0; }
    .main-header h1 { font-size: 2.1rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.2rem; }
    .main-header p  { font-size: 0.95rem; color: #666; }
    .info-box {
        background: #f0f4ff; border-left: 4px solid #4a6cf7;
        border-radius: 0 8px 8px 0; padding: 0.9rem 1.1rem;
        margin: 0.8rem 0; font-size: 0.88rem; color: #333;
    }
    .success-box {
        background: #f0fff4; border-left: 4px solid #38a169;
        border-radius: 0 8px 8px 0; padding: 0.9rem 1.1rem; margin: 0.8rem 0;
    }
    .mode-explain {
        background: #fff8e1; border-left: 4px solid #f59e0b;
        border-radius: 0 8px 8px 0; padding: 0.9rem 1.1rem;
        margin: 0.8rem 0; font-size: 0.88rem; color: #444;
    }
    .stButton > button { width: 100%; height: 3rem; font-size: 1rem; font-weight: 600; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ========== ヘッダー ==========
st.markdown("""
<div class="main-header">
    <h1>📄 PDF → PowerPoint 変換</h1>
    <p>PDFをPowerPointスライドに変換します</p>
</div>
""", unsafe_allow_html=True)
st.divider()


# ========== 変換モード選択 ==========
st.markdown("### ⚙️ 変換モードを選択")
mode = st.radio(
    label="変換モード",
    options=["📷  画像モード（高品質・忠実な再現）", "✏️  編集モード（テキスト編集・画像移動が可能）"],
    label_visibility="collapsed",
    horizontal=False,
)
is_edit_mode = "編集モード" in mode

if not is_edit_mode:
    st.markdown("""
    <div class="mode-explain">
        📷 <strong>画像モード</strong>：各PDFページを高品質な画像としてスライドに配置します。<br>
        見た目は原本と同じですが、テキストの直接編集はできません。
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="mode-explain">
        ✏️ <strong>編集モード</strong>：PDFの内容を分析し、テキストと画像を個別のオブジェクトとして配置します。<br>
        &nbsp;&nbsp;• <strong>テキスト</strong> → クリックして直接編集可能なテキストボックス<br>
        &nbsp;&nbsp;• <strong>画像</strong> → 自由に移動・リサイズ可能な画像オブジェクト<br>
        ※ PDFの背景デザインは再現されない場合があります
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ========== 詳細設定 ==========
st.markdown("### 🔧 設定")
col1, col2 = st.columns(2)
with col1:
    dpi = st.select_slider(
        "🎨 画像品質（DPI）",
        options=[72, 96, 144, 200, 300],
        value=144,
        help="画像モードのみに影響します。編集モードでは背景なしのため不要です。"
    )
with col2:
    if not is_edit_mode:
        img_fmt_label = st.selectbox(
            "🖼️ 画像形式",
            ["PNG（高品質）", "JPEG（軽量）"],
            index=0,
        )
        img_fmt = "png" if "PNG" in img_fmt_label else "jpeg"
    else:
        keep_bg = st.checkbox(
            "🖼️ 背景画像も保持する",
            value=False,
            key="keep_bg_val",
            help="チェックを入れると、テキスト・画像レイヤーに加えて、元ページの画像を背景として追加します（参考用）"
        )

st.divider()


# ========== ユーティリティ関数 ==========

def color_int_to_rgb(color_int):
    """PyMuPDF の色整数（0xRRGGBB）を RGBColor に変換"""
    r = (color_int >> 16) & 0xFF
    g = (color_int >> 8) & 0xFF
    b = color_int & 0xFF
    return RGBColor(r, g, b)


def clean_font_name(raw_name):
    """PDF フォント名からサブセットプレフィックスを除去して返す"""
    if not raw_name:
        return "Calibri"
    # "ABCDEF+FontName" → "FontName"
    if "+" in raw_name:
        raw_name = raw_name.split("+", 1)[1]
    # カンマ以降を除去（"Arial,Bold" → "Arial"）
    raw_name = raw_name.split(",")[0]
    # ハイフン以降を除去（"TimesNewRoman-Italic" → "TimesNewRoman"）
    raw_name = raw_name.split("-")[0]
    return raw_name.strip() or "Calibri"


def render_page_image(page, dpi, img_fmt="png"):
    """ページを画像にレンダリングして BytesIO で返す"""
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    buf = io.BytesIO()
    buf.write(pix.tobytes(img_fmt))
    buf.seek(0)
    return buf


def convert_image_mode(doc, dpi, img_fmt, progress_bar):
    """画像モード：各ページを1枚の画像としてスライドに変換"""
    emu_per_point = 914400 / 72
    first_rect = doc[0].rect
    slide_w_emu = int(first_rect.width * emu_per_point)
    slide_h_emu = int(first_rect.height * emu_per_point)

    prs = Presentation()
    prs.slide_width = Emu(slide_w_emu)
    prs.slide_height = Emu(slide_h_emu)
    blank_layout = prs.slide_layouts[6]
    num_pages = len(doc)

    for i, page in enumerate(doc):
        page_rect = page.rect
        page_w_emu = int(page_rect.width * emu_per_point)
        page_h_emu = int(page_rect.height * emu_per_point)
        left = (slide_w_emu - page_w_emu) // 2
        top  = (slide_h_emu - page_h_emu) // 2

        img_buf = render_page_image(page, dpi, img_fmt)
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            img_buf,
            Emu(max(0, left)), Emu(max(0, top)),
            width=Emu(page_w_emu), height=Emu(page_h_emu)
        )
        progress_bar.progress((i + 1) / num_pages, text=f"ページ {i+1}/{num_pages} 変換中...")

    return prs


def is_background_image(bbox, page_w, page_h, threshold=0.80):
    """ページ面積の80%以上を占める画像は背景とみなす"""
    x0, y0, x1, y1 = bbox
    img_area  = max(0, x1 - x0) * max(0, y1 - y0)
    page_area = page_w * page_h
    return page_area > 0 and (img_area / page_area) > threshold


def convert_edit_mode(doc, keep_bg, dpi, progress_bar):
    """
    編集モード：テキストをテキストボックス、画像を独立オブジェクトとして配置。
    keep_bg=True の場合は背景画像も最背面に追加。
    戻り値: (Presentation, total_textboxes, total_images)
    """
    emu_per_point = 914400 / 72
    first_rect = doc[0].rect
    slide_w_emu = int(first_rect.width * emu_per_point)
    slide_h_emu = int(first_rect.height * emu_per_point)

    prs = Presentation()
    prs.slide_width = Emu(slide_w_emu)
    prs.slide_height = Emu(slide_h_emu)
    blank_layout = prs.slide_layouts[6]
    num_pages = len(doc)

    total_textboxes = 0
    total_images    = 0

    for page_idx, page in enumerate(doc):
        page_rect = page.rect
        page_w = page_rect.width
        page_h = page_rect.height

        x_scale = slide_w_emu / page_w
        y_scale = slide_h_emu / page_h

        slide = prs.slides.add_slide(blank_layout)

        # ── 背景画像（オプション） ──────────────────────
        if keep_bg:
            bg_buf = render_page_image(page, dpi, "png")
            slide.shapes.add_picture(
                bg_buf, Emu(0), Emu(0),
                width=Emu(slide_w_emu), height=Emu(slide_h_emu)
            )

        # ── 埋め込み画像を個別オブジェクトとして追加 ────
        added_xrefs = set()
        for img_info in page.get_image_info(xrefs=True):
            xref = img_info.get("xref")
            bbox = img_info.get("bbox")
            if not bbox:
                continue

            x0, y0, x1, y1 = bbox

            # ページをほぼ全面覆う画像は背景とみなしてスキップ
            if is_background_image((x0, y0, x1, y1), page_w, page_h):
                continue

            w_emu    = int((x1 - x0) * x_scale)
            h_emu    = int((y1 - y0) * y_scale)
            left_emu = int(x0 * x_scale)
            top_emu  = int(y0 * y_scale)

            if w_emu <= 0 or h_emu <= 0:
                continue

            try:
                if xref and xref not in added_xrefs:
                    img_data   = doc.extract_image(xref)
                    img_stream = io.BytesIO(img_data["image"])
                    slide.shapes.add_picture(
                        img_stream,
                        Emu(max(0, left_emu)), Emu(max(0, top_emu)),
                        width=Emu(w_emu), height=Emu(h_emu)
                    )
                    added_xrefs.add(xref)
                    total_images += 1
                elif not xref:
                    clip_rect  = fitz.Rect(x0, y0, x1, y1)
                    pix        = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip_rect, alpha=False)
                    img_stream = io.BytesIO(pix.tobytes("png"))
                    slide.shapes.add_picture(
                        img_stream,
                        Emu(max(0, left_emu)), Emu(max(0, top_emu)),
                        width=Emu(w_emu), height=Emu(h_emu)
                    )
                    total_images += 1
            except Exception:
                pass

        # ── テキストブロックをテキストボックスとして追加 ─
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_MEDIABOX_CLIP)

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue

            bx0, by0, bx1, by1 = block["bbox"]
            bw, bh = bx1 - bx0, by1 - by0
            if bw <= 0 or bh <= 0:
                continue

            # ブロック内の実際のテキストを収集
            all_text = "".join(
                span.get("text", "")
                for line in block.get("lines", [])
                for span in line.get("spans", [])
            )
            if not all_text.strip():
                continue  # 空ブロックはスキップ

            left_emu = int(bx0 * x_scale)
            top_emu  = int(by0 * y_scale)
            w_emu    = int(bw  * x_scale)
            h_emu    = int(bh  * y_scale)

            if w_emu < 5000 or h_emu < 5000:
                continue

            txBox = slide.shapes.add_textbox(
                Emu(left_emu), Emu(top_emu),
                Emu(w_emu + 50000), Emu(h_emu + 50000)
            )
            tf = txBox.text_frame
            tf.word_wrap    = True
            tf.auto_size    = None
            tf.margin_left  = Emu(0)
            tf.margin_right = Emu(0)
            tf.margin_top   = Emu(0)
            tf.margin_bottom= Emu(0)

            first_para = True
            for line in block.get("lines", []):
                if first_para:
                    para = tf.paragraphs[0]
                    first_para = False
                else:
                    para = tf.add_paragraph()

                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if not text:
                        continue

                    run      = para.add_run()
                    run.text = text
                    font     = run.font

                    font.size = Pt(max(1, span.get("size", 11)))

                    try:
                        font.color.rgb = color_int_to_rgb(span.get("color", 0))
                    except Exception:
                        pass

                    flags        = span.get("flags", 0)
                    font.bold    = bool(flags & 16)
                    font.italic  = bool(flags & 2)

                    try:
                        font.name = clean_font_name(span.get("font", ""))
                    except Exception:
                        pass

            total_textboxes += 1

        progress_bar.progress((page_idx + 1) / num_pages, text=f"ページ {page_idx+1}/{num_pages} 変換中...")

    return prs, total_textboxes, total_images


# ========== ファイルアップロード ==========
st.markdown("### 📂 PDFファイルをアップロード")
uploaded_file = st.file_uploader(
    "ここにPDFをドラッグ＆ドロップ、またはクリックして選択",
    type=["pdf"],
)

pdf_bytes = None
if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()
    try:
        doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")
        num_pages   = len(doc_preview)
        first_rect  = doc_preview[0].rect
        w_mm = first_rect.width  * 25.4 / 72
        h_mm = first_rect.height * 25.4 / 72
        doc_preview.close()

        st.markdown(f"""
        <div class="info-box">
            📋 <strong>ファイル情報</strong><br>
            &nbsp;&nbsp;• ファイル名: <code>{uploaded_file.name}</code><br>
            &nbsp;&nbsp;• ページ数: <strong>{num_pages} ページ</strong><br>
            &nbsp;&nbsp;• ページサイズ: {w_mm:.0f} × {h_mm:.0f} mm（最初のページ基準）<br>
            &nbsp;&nbsp;• ファイルサイズ: {len(pdf_bytes)/1024:.0f} KB
        </div>
        """, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"PDFの読み込みに失敗しました: {e}")
        pdf_bytes = None

# ========== 変換ボタン ==========
st.markdown("### 🔄 変換")
if pdf_bytes is None:
    st.info("👆 まずPDFファイルをアップロードしてください。")
else:
    btn_label = "⚡ PowerPointに変換する（画像モード）" if not is_edit_mode else "⚡ PowerPointに変換する（編集モード）"
    if st.button(btn_label, type="primary"):
        try:
            progress_bar = st.progress(0, text="変換を開始しています...")
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")

            total_textboxes = 0
            total_images    = 0

            if not is_edit_mode:
                prs = convert_image_mode(doc, dpi, img_fmt, progress_bar)
            else:
                _keep_bg = st.session_state.get("keep_bg_val", False)
                prs, total_textboxes, total_images = convert_edit_mode(doc, _keep_bg, dpi, progress_bar)

            doc.close()

            # PPTX をバイト列として保存
            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)

            progress_bar.progress(1.0, text="✅ 変換完了！")

            # 完了メッセージ
            base_name   = os.path.splitext(uploaded_file.name)[0]
            output_name = f"{base_name}.pptx"
            pptx_kb     = len(pptx_io.getvalue()) / 1024
            mode_label  = "画像モード" if not is_edit_mode else "編集モード"

            st.markdown(f"""
            <div class="success-box">
                ✅ <strong>変換が完了しました！</strong><br>
                &nbsp;&nbsp;• モード: {mode_label}<br>
                &nbsp;&nbsp;• スライド数: <strong>{num_pages} 枚</strong><br>
                &nbsp;&nbsp;• ファイルサイズ: {pptx_kb:.0f} KB
            </div>
            """, unsafe_allow_html=True)

            if is_edit_mode:
                if total_textboxes == 0:
                    st.warning(
                        "⚠️ テキストを抽出できませんでした。\n\n"
                        "このPDFはスキャン画像や特殊なフォントを使用している可能性があります。\n"
                        "「**背景画像も保持する**」にチェックを入れて再変換するか、**画像モード**をお試しください。"
                    )
                else:
                    st.caption(
                        f"💡 テキストボックス {total_textboxes} 個・画像 {total_images} 個を配置しました。"
                        " PowerPointでテキストをクリックすると編集できます。画像はドラッグで移動できます。"
                    )

            st.download_button(
                label=f"📥 {output_name} をダウンロード",
                data=pptx_io.getvalue(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary"
            )

        except Exception as e:
            st.error(f"❌ 変換中にエラーが発生しました: {str(e)}")
            st.exception(e)

# ========== フッター ==========
st.divider()
st.markdown("""
<div style="text-align:center; color:#aaa; font-size:0.78rem;">
    PDF → PowerPoint 変換ツール　|　各PDFページが1枚のスライドに変換されます
</div>
""", unsafe_allow_html=True)
