import io
import os
import tempfile
import zipfile
import re
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pytesseract
import fitz  # PyMuPDF
from docx import Document
import openpyxl

st.set_page_config(page_title="HandyWriter Pro", page_icon="✍️", layout="wide")
st.title("✍️ HandyWriter Pro")
st.caption("Advanced Visual PDF Editor & Document Converter — Free, Offline, and Direct Control")

tab1, tab2, tab3 = st.tabs(
    ["📄 Foxit-Style PDF Direct Editor", "📝 Image → Word / Excel", "📬 Bulk Document Automation"]
)

# ---------------------------------------------------------------------------
# TAB 1: Foxit-Style PDF Direct Editor
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("🎯 Direct In-Place PDF Editor")
    st.write(
        "Upload any PDF to edit text elements directly. Modify sentences, names, or values "
        "in place without altering the surrounding structure, fonts, or alignments."
    )

    uploaded_pdf = st.file_uploader("Upload PDF File to Edit", type=["pdf"], key="pdf_direct_upload")

    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.success(f"Successfully loaded PDF document containing {doc_pdf.page_count} pages.")

        col_nav1, col_nav2 = st.columns([1, 3])
        with col_nav1:
            page_num = st.number_input(
                "Active Page", min_value=1, max_value=doc_pdf.page_count, value=1, key="direct_page_num"
            )
        with col_nav2:
            action_mode = st.radio(
                "Editing Mode",
                ["📝 In-Place Text Correction", "🎨 Visual Cover & Stamp", "🔄 Page Management & Watermark"],
                horizontal=True
            )

        page_idx = page_num - 1
        page = doc_pdf[page_idx]

        if action_mode == "📝 In-Place Text Correction":
            st.markdown("### 🔍 Select & Modify Text Blocks")
            
            text_dict = page.get_text("dict")
            lines_data = []
            for block in text_dict.get("blocks", []):
                for line in block.get("lines", []):
                    line_text = "".join(s["text"] for s in line["spans"])
                    if line_text.strip():
                        first_span = line["spans"][0] if line["spans"] else {"size": 11, "font": "Helvetica", "color": 0}
                        lines_data.append({
                            "bbox": line["bbox"],
                            "text": line_text,
                            "size": first_span.get("size", 11),
                            "font": first_span.get("font", "Helvetica"),
                            "color": first_span.get("color", 0)
                        })

            if not lines_data:
                st.warning("⚠️ No selectable text strings detected on this page.")
            else:
                zoom = 1.5
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                preview_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                draw = ImageDraw.Draw(preview_img)
                
                for i, data in enumerate(lines_data):
                    x0, y0, x1, y1 = [v * zoom for v in data["bbox"]]
                    draw.rectangle([x0, y0, x1, y1], outline="#1E88E5", width=2)
                    draw.text((x0, max(0, y0 - 15)), f"Block {i + 1}", fill="#0D47A1")

                col_layout1, col_layout2 = st.columns([1, 1])
                
                with col_layout1:
                    st.image(preview_img, caption=f"Interactive Workspace Map — Page {page_num}", use_container_width=True)
                
                with col_layout2:
                    st.markdown("#### 🖊️ Edit Element Contents")
                    st.info("Locate the Block Number from the visual map and modify its text values below.")
                    
                    edits_to_apply = {}
                    
                    st.markdown("**Global Font Control (Optional Override)**")
                    c_f1, c_f2, c_f3 = st.columns(3)
                    with c_f1:
                        font_family = st.selectbox(
                            "Font Family", ["Match Original PDF Font", "Helvetica / Arial", "Times New Roman", "Courier New"], key="direct_font"
                        )
                    with c_f2:
                        font_style = st.selectbox("Font Weight", ["Regular", "Bold", "Italic"], key="direct_style")
                    with c_f3:
                        size_override = st.number_input("Font Size Override (0 = Auto-match)", min_value=0, max_value=72, value=0)

                    font_map = {
                        "Helvetica / Arial": {"Regular": "helv", "Bold": "hebo", "Italic": "heit"},
                        "Times New Roman": {"Regular": "tiro", "Bold": "tibo", "Italic": "tiit"},
                        "Courier New": {"Regular": "cour", "Bold": "cobo", "Italic": "coit"},
                    }

                    st.divider()
                    
                    for i, data in enumerate(lines_data):
                        clean_text = data["text"].rstrip("\n")
                        new_val = st.text_input(f"Block #{i + 1} Text", value=clean_text, key=f"direct_inp_{page_idx}_{i}")
                        if new_val != clean_text:
                            edits_to_apply[i] = {
                                "bbox": fitz.Rect(data["bbox"]),
                                "new_text": new_val,
                                "orig_size": data["size"],
                                "orig_font": data["font"]
                            }

                    if edits_to_apply:
                        # Cover only the text area (underline will remain)
for idx, edit in edits_to_apply.items():

    bbox = edit["bbox"]

    tight_bbox = fitz.Rect(
        bbox.x0,
        bbox.y0,
        bbox.x1,
        bbox.y1 - 4
    )

    page.draw_rect(
        tight_bbox,
        fill=(1, 1, 1),
        color=None,
        overlay=True
    )
                            
                            page_rect = page.rect
                            for idx, edit in edits_to_apply.items():
                                bbox = edit["bbox"]
                                base_sz = size_override if size_override > 0 else edit["orig_size"]
                                
                                if font_family == "Match Original PDF Font":
                                    orig_f_lower = edit["orig_font"].lower()
                                    if "bold" in orig_f_lower:
                                        selected_font = "hebo" if "times" not in orig_f_lower else "tibo"
                                    elif "times" in orig_f_lower or "roman" in orig_f_lower:
                                        selected_font = "tiro"
                                    elif "cour" in orig_f_lower:
                                        selected_font = "cour"
                                    else:
                                        selected_font = "helv"
                                else:
                                    selected_font = font_map[font_family][font_style]

                                # Exact baseline placement alignment
                                fixed_rect = fitz.Rect(
                                    bbox.x0, bbox.y0,
                                    min(bbox.x0 + max(bbox.width, 400), page_rect.width - 20),
                                    bbox.y1 + 2
                                )
                                
                                page.insert_textbox(
                                    fixed_rect, 
                                    edit["new_text"], 
                                    fontsize=base_sz, 
                                    fontname=selected_font,
                                    align=0
                                )
                            
                            out_bytes = io.BytesIO(doc_pdf.tobytes())
                            st.success("✨ Underline-safe changes successfully compiled into the document!")
                            st.download_button(
                                "⬇️ Download Edited PDF Document",
                                data=out_bytes,
                                file_name="Foxit_Style_Fixed_Underline.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )

        elif action_mode == "🎨 Visual Cover & Stamp":
            st.markdown("### 🛡️ Precise Text Redaction & Custom Stamping")
            zoom = 1.5
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img_w, img_h = pix.width, pix.height
            
            col_c1, col_c2 = st.columns([1, 1])
            with col_c1:
                st.markdown("**1. Configure Target Boundaries (Pixels)**")
                cx0 = st.number_input("Left Margin (X0)", min_value=0, max_value=img_w, value=50)
                cy0 = st.number_input("Top Margin (Y0)", min_value=0, max_value=img_h, value=50)
                cx1 = st.number_input("Right Margin (X1)", min_value=0, max_value=img_w, value=250)
                cy1 = st.number_input("Bottom Margin (Y1)", min_value=0, max_value=img_h, value=100)
                
                st.markdown("**2. Content Customization**")
                stamp_text = st.text_area("Insertion Content")
                
                cc1, cc2 = st.columns(2)
                with cc1:
                    stamp_font = st.selectbox("Design Font", ["Helvetica", "Times-Roman", "Courier"])
                    stamp_align = st.selectbox("Text Alignment", ["Left", "Center"])
                with cc2:
                    stamp_size = st.number_input("Font Size Scale", min_value=6, max_value=120, value=12)
                    fill_color = st.color_picker("Cover Block Background Color", "#FFFFFF")

            workspace_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            draw_grid = ImageDraw.Draw(workspace_img)
            
            for gx in range(0, img_w, 50):
                draw_grid.line([(gx, 0), (gx, img_h)], fill=(230, 230, 250), width=1)
            for gy in range(0, img_h, 50):
                draw_grid.line([(0, gy), (img_w, gy)], fill=(230, 230, 250), width=1)
            draw_grid.rectangle([cx0, cy0, cx1, cy1], outline="#D32F2F", width=3)
            
            with col_c2:
                st.image(workspace_img, caption="Live Layout Placement & Coverage Grid", use_container_width=True)

            if st.button("🚀 Render Visual Stamp to Document Layer", use_container_width=True):
                pdf_target_rect = fitz.Rect(cx0 / zoom, cy0 / zoom, cx1 / zoom, cy1 / zoom)
                hex_color = fill_color.lstrip('#')
                rgb_tuple = tuple(int(hex_color[i:i+2], 16)/255.0 for i in (0, 2, 4))
                page.draw_rect(pdf_target_rect, color=None, fill=rgb_tuple, fill_opacity=1)
                
                if stamp_text.strip():
                    align_code = 1 if stamp_align == "Center" else 0
                    font_code = "helv" if stamp_font == "Helvetica" else ("tiro" if stamp_font == "Times-Roman" else "cour")
                    page.insert_textbox(pdf_target_rect, stamp_text, fontsize=stamp_size, fontname=font_code, align=align_code)
                
                stamp_out = io.BytesIO(doc_pdf.tobytes())
                st.success("Stamped layer integrated completely.")
                st.download_button(
                    "⬇️ Download Stamped PDF File",
                    data=stamp_out,
                    file_name="Stamped_Output.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

        elif action_mode == "🔄 Page Management & Watermark":
            st.markdown("### 🔧 Document Level Management Tools")
            m_col1, m_col2 = st.columns(2)
            with m_col1:
                st.markdown("#### 🚨 Remove Specified Pages")
                pages_to_drop = st.text_input("Enter Page Numbers to Delete (e.g., 2, 4)")
                if st.button("🗑️ Execute Page Deletion"):
                    if pages_to_drop.strip():
                        try:
                            drop_indices = [int(p.strip()) - 1 for p in pages_to_drop.split(",") if p.strip()]
                            doc_pdf.delete_pages(drop_indices)
                            del_out = io.BytesIO(doc_pdf.tobytes())
                            st.success("Pages successfully dropped.")
                            st.download_button("⬇️ Download Updated PDF", data=del_out, file_name="Updated_Layout.pdf", mime="application/pdf")
                        except Exception as err:
                            st.error(f"Error handling configuration parsing: {err}")
            
            with m_col2:
                st.markdown("#### 🏷️ Overlay Global Security Watermark")
                wm_text = st.text_input("Watermark Label String", value="CONFIDENTIAL")
                if st.button("✨ Apply Watermark Elements"):
                    for p in doc_pdf:
                        r = p.rect
                        center_pt = fitz.Point(r.width / 2, r.height / 2)
                        p.insert_text(
                            (r.width / 4, r.height / 2),
                            wm_text,
                            fontsize=45,
                            color=(0.8, 0.8, 0.8),
                            overlay=True,
                            morph=(center_pt, fitz.Matrix(45))
                        )
                    wm_out = io.BytesIO(doc_pdf.tobytes())
                    st.success("Watermark successfully applied.")
                    st.download_button("⬇️ Download Watermarked PDF", data=wm_out, file_name="Watermarked_Document.pdf", mime="application/pdf")

# ---------------------------------------------------------------------------
# TAB 2: Image -> Word / Excel
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("📝 Intelligent Image-to-Document Extraction (OCR)")
    uploaded_img = st.file_uploader("Upload Image Document (JPG, PNG)", type=["jpg", "jpeg", "png"], key="ocr_img_upload")

    if uploaded_img:
        img_obj = Image.open(uploaded_img).convert("RGB")
        col_img1, col_img2 = st.columns(2)
        with col_img1:
            st.image(img_obj, caption="Original Image Resource", use_container_width=True)

        with st.spinner("Executing OCR..."):
            raw_ocr_text = pytesseract.image_to_string(img_obj, config="--psm 6")

        with col_img2:
            edited_ocr_text = st.text_area("Verified Extraction String", raw_ocr_text, height=350)

        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("📝 Convert to Word (.docx)", use_container_width=True):
                word_doc = Document()
                for line in edited_ocr_text.split("\n"):
                    word_doc.add_paragraph(line)
                buf = io.BytesIO()
                word_doc.save(buf)
                buf.seek(0)
                st.download_button("⬇️ Download Word file", data=buf, file_name="converted.docx", use_container_width=True)
        with c_btn2:
            if st.button("📊 Convert to Excel (.xlsx)", use_container_width=True):
                wb = openpyxl.Workbook()
                ws = wb.active
                for line in edited_ocr_text.split("\n"):
                    if line.strip():
                        ws.append(re.split(r"\t|\s{2,}", line.strip()))
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.download_button("⬇️ Download Excel file", data=buf, file_name="converted.xlsx", use_container_width=True)

# ---------------------------------------------------------------------------
# TAB 3: Bulk Document Automation (Mail Merge)
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("📬 Corporate Bulk Automation and Mail Merge Engine")
    st.write("Generate thousands of personalized PDFs using tabular Excel registers (.xlsx) instantly.")
    
    tpl_pdf = st.file_uploader("1. Upload Template PDF File (with {{TOKENS}} placeholders)", type=["pdf"], key="bulk_tpl_pdf")
    data_sheet = st.file_uploader("2. Upload Excel Dataset Register (.xlsx)", type=["xlsx"], key="bulk_data_sheet")

    if tpl_pdf and data_sheet:
        tpl_bytes = tpl_pdf.read()
        wb = openpyxl.load_workbook(io.BytesIO(data_sheet.read()))
        ws = wb.active
        
        data_headers = [cell.value for cell in ws[1] if cell.value is not None]
        records = list(ws.iter_rows(min_row=2, values_only=True))
        
        st.success(f"📊 Discovered **{len(records)}** active rows matching data sheet criteria. Columns: {', '.join(data_headers)}")
        naming_factor = st.selectbox("Select File Naming Column Mapping (e.g. Employee Name / ID)", options=data_headers)

        if st.button("⚙️ Trigger Bulk Mail-Merge Compilation Loop", use_container_width=True):
            zip_memory_stream = io.BytesIO()
            progress_bar = st.progress(0)
            
            with zipfile.ZipFile(zip_memory_stream, "w") as archive:
                for idx, data_row in enumerate(records):
                    row_context = dict(zip(data_headers, data_row))
                    fresh_doc = fitz.open(stream=tpl_bytes, filetype="pdf")
                    
                    for active_p in fresh_doc:
                        block_text_list = [b for b in active_p.get_text("blocks") if b[4].strip()]
                        
                        sub_actions = []
                        for individual_b in block_text_list:
                            item_text = individual_b[4]
                            was_altered = False
                            for header_key, value_item in row_context.items():
                                target_token = "{{" + str(header_key) + "}}"
                                if target_token in item_text:
                                    item_text = item_text.replace(target_token, str(value_item))
                                    was_altered = True
                            if was_altered:
                                sub_actions.append((fitz.Rect(individual_b[:4]), item_text.rstrip("\n")))

                        # Underline safe batch redaction
                        for coord_box, _ in sub_actions:
                            tight_box = fitz.Rect(coord_box.x0, coord_box.y0, coord_box.x1, coord_box.y1 - 1.5)
                            active_p.add_redact_annot(tight_box, fill=(1, 1, 1))
                        active_p.apply_redactions()
                        
                        for coord_box, formatted_str in sub_actions:
                            fixed_rect = fitz.Rect(coord_box.x0, coord_box.y0, coord_box.x1, coord_box.y1 + 2)
                            active_p.insert_textbox(fixed_rect, formatted_str, fontsize=11, fontname="helv", align=0)
                            
                    file_label = str(row_context.get(naming_factor, f"Record_{idx+1}")).replace(" ", "_")
                    archive.writestr(f"Automated_Outputs/{file_label}.pdf", fresh_doc.tobytes())
                    progress_bar.progress((idx + 1) / len(records))
            
            st.success("✨ Batch transformation sequence completed successfully!")
            st.download_button(
                "⬇️ Download All Compiled Documents (ZIP File)",
                data=zip_memory_stream.getvalue(),
                file_name="Automated_Bulk_Outputs.zip",
                mime="application/zip",
                use_container_width=True
            )

st.divider()
st.caption("HandyWriter Pro v2.6 • Underline Layout Patched Fix")
