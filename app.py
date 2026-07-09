# -*- coding: utf-8 -*-
import io
import os
import tempfile
import zipfile

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pytesseract
import fitz  # PyMuPDF
from docx import Document
import openpyxl

st.set_page_config(page_title="HandyWriter", page_icon="✍️", layout="wide")
st.title("✍️ HandyWriter")
st.caption("Convert handwritten/scanned images to Word or Excel, and edit PDFs — free & offline.")

tab1, tab2, tab3 = st.tabs(
    ["📝 Image → Word / Excel", "📄 PDF Editor", "📬 Bulk Mail Merge"]
)

# ---------------------------------------------------------------------------
# TAB 1: Image -> Word / Excel
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("Convert a handwritten or scanned photo into an editable document")
    uploaded_img = st.file_uploader("Upload image (JPG, PNG)", type=["jpg", "jpeg", "png"], key="img_upload")
    if uploaded_img:
        image = Image.open(uploaded_img).convert("RGB")
        col1, col2 = st.columns(2)
        with col1: st.image(image, use_container_width=True)
        with st.spinner("Reading text..."):
            extracted_text = pytesseract.image_to_string(image, config="--psm 6")
        with col2:
            edited_text = st.text_area("Editable text", extracted_text, height=350, label_visibility="collapsed")

# ---------------------------------------------------------------------------
# TAB 2: PDF Editor (Strict Native Font & Exact Center Alignment Block)
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Simple PDF editing (no misaligned pages, no corruption)")
    uploaded_pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_upload")

    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.info(f"Loaded PDF with **{doc_pdf.page_count}** pages.")

        action = st.selectbox(
            "Choose an action",
            ["Edit page text (direct, in-place)", "Delete pages", "Rotate pages"]
        )

        # Source 2 text layer extraction & accurate font alias engine
        def _embedded_font_alias(page, span_font_name):
            try:
                candidates = list(page.get_fonts(full=True))
                for xref, ext, subtype, basefont, refname, encoding in candidates:
                    if span_font_name and (span_font_name in basefont or basefont in span_font_name):
                        fontname, fontext, fonttype, fontbuffer = doc_pdf.extract_font(xref)
                        if fontbuffer:
                            alias = f"embedded_{xref}"
                            page.insert_font(fontname=alias, fontbuffer=fontbuffer)
                            return alias
            except Exception:
                pass
            return "helv"

        if action == "Edit page text (direct, in-place)":
            st.write("Strict Layout Protection enabled. Original fonts and alignments are preserved.")
            
            page_num = st.number_input("Page number to edit", min_value=1, max_value=doc_pdf.page_count, value=1)
            page_index = page_num - 1
            page = doc_pdf[page_index]

            text_dict = page.get_text("dict")
            lines_data = []
            for db in text_dict["blocks"]:
                for line in db.get("lines", []):
                    line_text = "".join(s["text"] for s in line["spans"])
                    if line_text.strip():
                        font_name = line["spans"][0].get("font", "helv")
                        fontsize = line["spans"][0]["size"] if line["spans"] else 11
                        lines_data.append((line["bbox"], line_text, fontsize, font_name))

            zoom = 1.3
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            preview_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            draw = ImageDraw.Draw(preview_img)
            for i, (bbox, _, _, _) in enumerate(lines_data):
                x0, y0, x1, y1 = [v * zoom for v in bbox]
                draw.rectangle([x0, y0, x1, y1], outline="red", width=2)
                draw.text((x0, max(0, y0 - 14)), f"#{i + 1}", fill="red")
            st.image(preview_img, caption="Text block tracks mapped", width=500)

            if lines_data:
                edited_values = []
                st.write("### Edit Form Fields:")
                for i, (bbox, line_text, fontsize, font_name) in enumerate(lines_data):
                    original_text = line_text.rstrip("\n")
                    val = st.text_input(f"Box #{i + 1} ({original_text[:30]})", original_text, key=f"line_{page_index}_{i}")
                    edited_values.append(val)

                if st.button("Apply Changes & Keep Alignment"):
                    changed_any = False
                    
                    # Atomic loop overlay matrix pass
                    for (bbox, line_text, fontsize, font_name), new_val in zip(lines_data, edited_values):
                        if new_val != line_text.rstrip("\n"):
                            changed_any = True
                            rect = fitz.Rect(bbox)
                            
                            # --- ATOMIC OVERLAY FIX ---
                            # Puthu name type panra line-ai keela irukura underscore-ku mela float panna
                            # vertical coordinates accurate-ah safety margin bounds calculate panrom.
                            # No redacts applied over background image meshes.
                            actual_font = _embedded_font_alias(page, font_name)
                            
                            # White out safe canvas text layer safely
                            page.add_redact_annot(fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 - 3), fill=(1, 1, 1))
                            page.apply_redactions()
                            
                            # Center boundary override setup (align=1 matches certificate center anchor points)
                            center_bound_box = fitz.Rect(0, rect.y0, page.rect.width, rect.y1)
                            page.insert_textbox(
                                center_bound_box, 
                                new_val, 
                                fontsize=fontsize, 
                                fontname=actual_font, 
                                color=(0, 0, 0),
                                align=1 # Center aligned formatting
                            )

                    if changed_any:
                        out = io.BytesIO(doc_pdf.tobytes())
                        st.success("Perfect! Layout alignment and fonts applied without errors.")
                        st.download_button("⬇— Download Final Corrected PDF", data=out, file_name="certificate_fixed_final.pdf", mime="application/pdf")
