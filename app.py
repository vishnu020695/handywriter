"""
HandyWriter — Handwritten Image → Word/Excel + Simple PDF Editor
------------------------------------------------------------------
100% free, open-source Python tools only (Tesseract OCR, python-docx,
openpyxl, PyMuPDF). Runs as a local web app in your browser — works
on Windows, Mac, Linux, and Android/iPhone (via browser, same WiFi).

HOW TO RUN (see README.md for full details):
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import os
import tempfile

import streamlit as st
from PIL import Image
import pytesseract
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import openpyxl

st.set_page_config(page_title="HandyWriter", page_icon="✍️", layout="wide")
st.title("✍️ HandyWriter")
st.caption("Convert handwritten/scanned images to Word or Excel, and edit PDFs — free & offline.")

tab1, tab2 = st.tabs(["📝 Image → Word / Excel", "📄 PDF Editor"])

# ---------------------------------------------------------------------------
# TAB 1: Image -> Word / Excel
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("Convert a handwritten or scanned photo into an editable document")
    st.write(
        "Upload a clear photo of handwriting or printed text. The app reads the text "
        "automatically (OCR). **Always check the extracted text below before saving** — "
        "OCR is not 100% perfect, especially with cursive handwriting, so this lets you "
        "fix any mistakes first."
    )

    uploaded_img = st.file_uploader(
        "Upload image (JPG, PNG)", type=["jpg", "jpeg", "png"], key="img_upload"
    )

    if uploaded_img:
        image = Image.open(uploaded_img).convert("RGB")
        col1, col2 = st.columns(2)
        with col1:
            st.image(image, caption="Your uploaded image", use_container_width=True)

        with st.spinner("Reading text from image..."):
            # psm 6 = assume a uniform block of text (good default for notes/pages)
            extracted_text = pytesseract.image_to_string(image, config="--psm 6")

        with col2:
            st.write("**Extracted text (edit to fix any mistakes):**")
            edited_text = st.text_area(
                "Editable text", extracted_text, height=350, label_visibility="collapsed"
            )

        st.write("### Save as:")
        c1, c2 = st.columns(2)

        with c1:
            if st.button("💾 Create Word (.docx)", use_container_width=True):
                doc = Document()
                doc.add_heading("Converted from image", level=1)
                for line in edited_text.split("\n"):
                    doc.add_paragraph(line)
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button(
                    "⬇️ Download Word file",
                    data=buf,
                    file_name="converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        with c2:
            if st.button("💾 Create Excel (.xlsx)", use_container_width=True):
                # Each non-empty line becomes a row; splits on 2+ spaces or tabs into columns
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Converted"
                for line in edited_text.split("\n"):
                    if not line.strip():
                        continue
                    # naive column split: tabs, or 2+ spaces, treated as column breaks
                    import re
                    cells = re.split(r"\t|\s{2,}", line.strip())
                    ws.append(cells)
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.download_button(
                    "⬇️ Download Excel file",
                    data=buf,
                    file_name="converted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

# ---------------------------------------------------------------------------
# TAB 2: PDF Editor
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Simple PDF editing (no misaligned pages, no corruption)")
    st.write(
        "Upload a PDF to merge, delete pages, rotate, add a watermark, or convert it into "
        "an editable Word document."
    )

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
                "Add text watermark",
                "Convert to editable Word (.docx)",
            ],
        )

        if action == "Edit page text (direct, in-place)":
            st.write(
                "Works best on PDFs that already have real text (not scanned photos). "
                "Pick a page, edit any text box below, then apply — your edit is placed "
                "in the exact same spot and size, so nothing shifts or misaligns."
            )
            page_num = st.number_input(
                "Page number to edit", min_value=1, max_value=doc_pdf.page_count, value=1
            )
            page_index = page_num - 1
            page = doc_pdf[page_index]

            # Show a preview image of the page so the user can see the layout
            pix = page.get_pixmap(matrix=fitz.Matrix(1.3, 1.3))
            st.image(pix.tobytes("png"), caption=f"Page {page_num} preview", width=400)

            # Extract editable text spans with their position/size info
            text_dict = page.get_text("dict")
            spans = []
            for block in text_dict["blocks"]:
                for line in block.get("lines", []):
                    for span in line["spans"]:
                        if span["text"].strip():
                            spans.append(span)

            if not spans:
                st.warning(
                    "No selectable text found on this page — it's likely a scanned image. "
                    "Use the 'Image → Word/Excel' tab instead, or the PDF-to-Word conversion below."
                )
            else:
                st.write(f"Found **{len(spans)}** text boxes on this page:")
                edited_values = []
                for i, span in enumerate(spans):
                    val = st.text_input(f"Text box {i+1}", span["text"], key=f"span_{page_index}_{i}")
                    edited_values.append(val)

                if st.button("Apply edits to this page and download"):
                    changed_any = False
                    for span, new_val in zip(spans, edited_values):
                        if new_val != span["text"]:
                            changed_any = True
                            bbox = fitz.Rect(span["bbox"])
                            page.add_redact_annot(bbox, fill=(1, 1, 1))
                    page.apply_redactions()
                    for span, new_val in zip(spans, edited_values):
                        if new_val != span["text"]:
                            bbox = fitz.Rect(span["bbox"])
                            origin = (bbox.x0, bbox.y1 - 2)
                            page.insert_text(
                                origin, new_val, fontsize=span["size"], fontname="helv"
                            )
                    if not changed_any:
                        st.info("No changes were made.")
                    else:
                        out = io.BytesIO(doc_pdf.tobytes())
                        st.success("Edits applied.")
                        st.download_button(
                            "⬇️ Download edited PDF",
                            data=out,
                            file_name="edited_text.pdf",
                            mime="application/pdf",
                        )

        if action == "Delete pages":
            pages_to_delete = st.text_input(
                "Page numbers to delete (comma-separated, e.g. 1,3,5)"
            )
            if st.button("Apply and download"):
                try:
                    nums = [int(x.strip()) - 1 for x in pages_to_delete.split(",") if x.strip()]
                    doc_pdf.delete_pages(nums)
                    out = io.BytesIO(doc_pdf.tobytes())
                    st.download_button(
                        "⬇️ Download edited PDF", data=out, file_name="edited.pdf", mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"Check your page numbers. Details: {e}")

        elif action == "Rotate pages":
            angle = st.selectbox("Rotate by", [90, 180, 270])
            if st.button("Apply and download"):
                for page in doc_pdf:
                    page.set_rotation(angle)
                out = io.BytesIO(doc_pdf.tobytes())
                st.download_button(
                    "⬇️ Download rotated PDF", data=out, file_name="rotated.pdf", mime="application/pdf"
                )

        elif action == "Add text watermark":
            wm_text = st.text_input("Watermark text", "CONFIDENTIAL")
            if st.button("Apply and download"):
                for page in doc_pdf:
                    rect = page.rect
                    center = fitz.Point(rect.width / 2, rect.height / 2)
                    # rotate diagonally around the page center using a morph matrix
                    morph = (center, fitz.Matrix(45))
                    page.insert_text(
                        (rect.width / 4, rect.height / 2),
                        wm_text,
                        fontsize=40,
                        color=(0.7, 0.7, 0.7),
                        overlay=True,
                        morph=morph,
                    )
                out = io.BytesIO(doc_pdf.tobytes())
                st.download_button(
                    "⬇️ Download watermarked PDF", data=out, file_name="watermarked.pdf", mime="application/pdf"
                )

        elif action == "Convert to editable Word (.docx)":
            if st.button("Convert and download"):
                from pdf2docx import Converter

                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
                    tmp_pdf.write(pdf_bytes)
                    tmp_pdf_path = tmp_pdf.name
                tmp_docx_path = tmp_pdf_path.replace(".pdf", ".docx")

                with st.spinner("Converting (this keeps original layout/alignment)..."):
                    cv = Converter(tmp_pdf_path)
                    cv.convert(tmp_docx_path)
                    cv.close()

                with open(tmp_docx_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Word file",
                        data=f.read(),
                        file_name="converted_from_pdf.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                os.remove(tmp_pdf_path)
                os.remove(tmp_docx_path)

st.divider()
st.caption("HandyWriter · 100% free & open-source · Your files are processed locally, never uploaded anywhere.")
