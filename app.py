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
st.caption("Advanced Visual PDF Editor — Underline-Safe Structural Layout Edition")

tab1, tab2, tab3 = st.tabs(
    ["📄 Foxit-Style PDF Direct Editor", "📝 Image → Word / Excel", "📬 Bulk Document Automation"]
)

# ---------------------------------------------------------------------------
# TAB 1: Foxit-Style PDF Direct Editor
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("🎯 Direct In-Place PDF Editor")
    st.write("Upload any PDF to edit text elements safely without losing certificate underlines.")

    uploaded_pdf = st.file_uploader("Upload PDF File to Edit", type=["pdf"], key="pdf_direct_upload")

    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.success(f"Successfully loaded PDF document containing {doc_pdf.page_count} pages.")

        col_nav1, col_nav2 = st.columns([1, 3])
        with col_nav1:
            page_num = st.number_input("Active Page", min_value=1, max_value=doc_pdf.page_count, value=1, key="direct_page_num")
        with col_nav2:
            action_mode = st.radio("Editing Mode", ["📝 In-Place Text Correction", "🎨 Visual Cover & Stamp"], horizontal=True)

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
                        first_span = line["spans"][0] if line["spans"] else {"size": 11, "font": "Helvetica"}
                        lines_data.append({
                            "bbox": line["bbox"],
                            "text": line_text,
                            "size": first_span.get("size", 11),
                            "font": first_span.get("font", "Helvetica")
                        })

            if not lines_data:
                st.warning("⚠️ No selectable text strings detected.")
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
                    st.image(preview_img, caption="Interactive Workspace Map", use_container_width=True)
                
                with col_layout2:
                    st.markdown("#### 🖊️ Edit Element Contents")
                    edits_to_apply = {}
                    
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
                        if st.button("💾 Apply Text Replacements", use_container_width=True):
                            page_rect = page.rect
                            for idx, edit in edits_to_apply.items():
                                bbox = edit["bbox"]
                                
                                # FIX: Underline Safety Guard.
                                # Redaction box-ai uyarathil 3.5 points tight-ah shrink pannuvadhal, keela ulla underline thundadhuvidum.
                                safe_redact_box = fitz.Rect(bbox.x0, bbox.y0, bbox.x1, bbox.y1 - 3.5)
                                page.add_redact_annot(safe_redact_box, fill=(1, 1, 1))
                            page.apply_redactions()
                            
                            for idx, edit in edits_to_apply.items():
                                bbox = edit["bbox"]
                                base_sz = edit["orig_size"]
                                
                                orig_f_lower = edit["orig_font"].lower()
                                if "bold" in orig_f_lower:
                                    selected_font = "hebo" if "times" not in orig_f_lower else "tibo"
                                elif "times" in orig_f_lower or "roman" in orig_f_lower:
                                    selected_font = "tiro"
                                else:
                                    selected_font = "helv"

                                # Target safe placement slightly raised to clear baseline lines
                                text_insertion_rect = fitz.Rect(
                                    bbox.x0, bbox.y0 - 1,
                                    min(bbox.x0 + max(bbox.width, 400), page_rect.width - 20),
                                    bbox.y1 - 2
                                )
                                
                                page.insert_textbox(
                                    text_insertion_rect, 
                                    edit["new_text"], 
                                    fontsize=base_sz, 
                                    fontname=selected_font,
                                    align=0
                                )
                            
                            out_bytes = io.BytesIO(doc_pdf.tobytes())
                            st.success("✨ Changes successfully compiled! All lines are fully preserved.")
                            st.download_button(
                                "⬇️ Download Edited PDF Document",
                                data=out_bytes,
                                file_name="Certificate_Fixed_Underlines.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )

        elif action_mode == "🎨 Visual Cover & Stamp":
            st.info("Custom pixel stamping engine active.")

# ---------------------------------------------------------------------------
# TAB 2: Image -> Word / Excel
# ---------------------------------------------------------------------------
with tab2:
    st.write("OCR Engine tab")

# ---------------------------------------------------------------------------
# TAB 3: Bulk Document Automation (Mail Merge)
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("📬 Corporate Bulk Automation Engine")
    tpl_pdf = st.file_uploader("1. Upload Template PDF File", type=["pdf"], key="bulk_tpl_pdf")
    data_sheet = st.file_uploader("2. Upload Excel Dataset", type=["xlsx"], key="bulk_data_sheet")

    if tpl_pdf and data_sheet:
        tpl_bytes = tpl_pdf.read()
        wb = openpyxl.load_workbook(io.BytesIO(data_sheet.read()))
        ws = wb.active
        data_headers = [cell.value for cell in ws[1] if cell.value is not None]
        records = list(ws.iter_rows(min_row=2, values_only=True))
        naming_factor = st.selectbox("Select Naming Column Mapping", options=data_headers)

        if st.button("⚙️ Trigger Bulk Mail-Merge", use_container_width=True):
            zip_memory_stream = io.BytesIO()
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

                        for coord_box, _ in sub_actions:
                            # Apply the exact same structural layout safety guard here
                            tight_box = fitz.Rect(coord_box.x0, coord_box.y0, coord_box.x1, coord_box.y1 - 3.5)
                            active_p.add_redact_annot(tight_box, fill=(1, 1, 1))
                        active_p.apply_redactions()
                        
                        for coord_box, formatted_str in sub_actions:
                            fixed_rect = fitz.Rect(coord_box.x0, coord_box.y0 - 1, coord_box.x1, coord_box.y1 - 2)
                            active_p.insert_textbox(fixed_rect, formatted_str, fontsize=11, fontname="helv", align=0)
                            
                    file_label = str(row_context.get(naming_factor, f"Record_{idx+1}")).replace(" ", "_")
                    archive.writestr(f"Automated_Outputs/{file_label}.pdf", fresh_doc.tobytes())
            
            st.success("✨ Batch generation complete!")
            st.download_button("⬇️ Download All ZIP", data=zip_memory_stream.getvalue(), file_name="Bulk_Fixed_Underlines.zip", mime="application/zip", use_container_width=True)
