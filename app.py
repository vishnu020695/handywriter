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
        with col1:
            st.image(image, caption="Your uploaded image", use_container_width=True)
        with st.spinner("Reading text from image..."):
            extracted_text = pytesseract.image_to_string(image, config="--psm 6")
        with col2:
            st.write("**Extracted text (edit to fix any mistakes):**")
            edited_text = st.text_area("Editable text", extracted_text, height=350, label_visibility="collapsed")

# ---------------------------------------------------------------------------
# TAB 2: PDF Editor (Muzhu Prachanaiyum Theerkum Safe Redaction Engine)
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
            [
                "Edit page text (direct, in-place)",
                "Delete pages",
                "Rotate pages",
            ],
        )

        if action == "Edit page text (direct, in-place)":
            st.write("Edit fields safely. The engine calculates boundary coordinates to save original underlines.")
            
            page_num = st.number_input("Page number to edit", min_value=1, max_value=doc_pdf.page_count, value=1)
            page_index = page_num - 1
            page = doc_pdf[page_index]

            text_dict = page.get_text("dict")
            lines_data = []
            for db in text_dict["blocks"]:
                for line in db.get("lines", []):
                    line_text = "".join(s["text"] for s in line["spans"])
                    if line_text.strip():
                        fontsize = line["spans"][0]["size"] if line["spans"] else 11
                        lines_data.append((line["bbox"], line_text, fontsize))

            # Preview box mapping
            zoom = 1.3
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            preview_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            draw = ImageDraw.Draw(preview_img)
            for i, (bbox, _, _) in enumerate(lines_data):
                x0, y0, x1, y1 = [v * zoom for v in bbox]
                draw.rectangle([x0, y0, x1, y1], outline="red", width=2)
                draw.text((x0, max(0, y0 - 14)), f"#{i + 1}", fill="red")
            st.image(preview_img, caption="Certificate text lines tracked", width=500)

            if lines_data:
                edited_values = []
                st.write("### Edit Form Fields:")
                for i, (bbox, line_text, fontsize) in enumerate(lines_data):
                    original_text = line_text.rstrip("\n")
                    val = st.text_input(f"Box #{i + 1} ({original_text[:30]})", original_text, key=f"line_{page_index}_{i}")
                    edited_values.append(val)

                if st.button("Apply Changes & Protect Underlines"):
                    changed_any = False
                    
                    for (bbox, line_text, fontsize), new_val in zip(lines_data, edited_values):
                        if new_val != line_text.rstrip("\n"):
                            changed_any = True
                            rect = fitz.Rect(bbox)
                            
                            # --- MATHEMATICAL UNDERLINE PROTECTION ---
                            # Original box-oda keela irukura kadaisi 3 pixels-ai thavirthu redaction panrom.
                            # Ithanala name-um erase aagum, aana underline safe-ah thapichidum!
                            safe_erase_zone = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 - 3)
                            
                            # True white fill potu pathaiye name/dept-ai force-ah block panrom
                            page.add_redact_annot(safe_erase_zone, fill=(1, 1, 1))
                            page.apply_redactions()
                            
                            # Puthu text insert pannuvom. Base height line-ku mela float aaga (rect.y1 - 4)
                            padded_write_zone = fitz.Rect(rect.x0, rect.y0, rect.x1 + 400, rect.y1)
                            page.insert_textbox(
                                padded_write_zone, 
                                new_val, 
                                fontsize=fontsize, 
                                fontname="hebo", # Bold matching text
                                color=(0, 0, 0)
                            )

                    if changed_any:
                        out = io.BytesIO(doc_pdf.tobytes())
                        st.success("Changes updated perfectly without disturbing underlines!")
                        st.download_button("⬇️ Download Fixed PDF", data=out, file_name="certificate_perfect.pdf", mime="application/pdf")
                    else:
                        st.info("No modifications to save.")

# ---------------------------------------------------------------------------
# TAB 3: Bulk Mail Merge
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("Generate personalized files")
    merge_mode = st.radio("Choose a method", ["Token-based", "Coordinate-based"])
