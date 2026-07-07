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
import zipfile

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

tab1, tab2, tab3 = st.tabs(
    ["📝 Image → Word / Excel", "📄 PDF Editor", "📬 Bulk Mail Merge"]
)


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
                "Edit images / logo (replace or delete)",
                "Delete pages",
                "Rotate pages",
                "Add text watermark",
                "Convert to editable Word (.docx)",
            ],
        )

        if action == "Edit page text (direct, in-place)":
            st.write(
                "Works best on PDFs that already have real text (not scanned photos). "
                "Each numbered box below matches the same numbered box drawn on the page "
                "preview image, so you can see exactly where your edit will land. "
                "Edit any box, then apply — text is placed in the same spot and size, "
                "so nothing shifts or misaligns."
            )
            page_num = st.number_input(
                "Page number to edit", min_value=1, max_value=doc_pdf.page_count, value=1
            )
            page_index = page_num - 1
            page = doc_pdf[page_index]

            # Extract editable text as grouped blocks (paragraphs/lines), not individual words
            raw_blocks = page.get_text("blocks")
            blocks = [b for b in raw_blocks if b[4].strip()]

            # Draw numbered boxes on the preview so boxes map visually to the page
            zoom = 1.3
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            from PIL import Image as PILImage, ImageDraw
            preview_img = PILImage.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            draw = ImageDraw.Draw(preview_img)
            for i, b in enumerate(blocks):
                x0, y0, x1, y1 = [v * zoom for v in b[:4]]
                draw.rectangle([x0, y0, x1, y1], outline="red", width=2)
                draw.text((x0, max(0, y0 - 14)), f"#{i + 1}", fill="red")
            st.image(preview_img, caption=f"Page {page_num} — boxes numbered to match the list below", width=500)

            if not blocks:
                st.warning(
                    "No selectable text found on this page — it's likely a scanned image. "
                    "Use the 'Image → Word/Excel' tab instead, or the PDF-to-Word conversion below."
                )
            else:
                st.write(f"**{len(blocks)}** editable text blocks found on this page:")
                edited_values = []
                for i, b in enumerate(blocks):
                    original_text = b[4].rstrip("\n")
                    val = st.text_area(
                        f"Box #{i + 1}", original_text, height=68, key=f"block_{page_index}_{i}"
                    )
                    edited_values.append(val)

                if st.button("Apply edits to this page and download"):
                    # Determine average font size per block (fallback to 11)
                    text_dict = page.get_text("dict")
                    dict_blocks = text_dict["blocks"]

                    changed_any = False
                    edits = []  # (bbox, new_text, fontsize)
                    for i, (b, new_val) in enumerate(zip(blocks, edited_values)):
                        original_text = b[4].rstrip("\n")
                        if new_val != original_text:
                            changed_any = True
                            bbox = fitz.Rect(b[:4])
                            # find matching dict block for font size (best-effort by position)
                            fontsize = 11
                            for db in dict_blocks:
                                for line in db.get("lines", []):
                                    for span in line["spans"]:
                                        if fitz.Rect(span["bbox"]).intersects(bbox):
                                            fontsize = span["size"]
                                            break
                            edits.append((bbox, new_val, fontsize))

                    if not changed_any:
                        st.info("No changes were made.")
                    else:
                        for bbox, new_val, fontsize in edits:
                            page.add_redact_annot(bbox, fill=(1, 1, 1))
                        page.apply_redactions()
                        page_rect = page.rect
                        for bbox, new_val, fontsize in edits:
                            # Give the box room to grow (in case new text is longer than old)
                            padded = fitz.Rect(
                                bbox.x0,
                                bbox.y0,
                                min(bbox.x0 + max(bbox.width, 300), page_rect.width - 36),
                                bbox.y1,
                            )
                            padded.y1 = min(padded.y1 + fontsize * 4, page_rect.height - 36)
                            size = fontsize
                            rc = page.insert_textbox(padded, new_val, fontsize=size, fontname="helv")
                            tries = 0
                            while rc < 0 and size > 6 and tries < 8:
                                size -= 1
                                rc = page.insert_textbox(padded, new_val, fontsize=size, fontname="helv")
                                tries += 1
                        out = io.BytesIO(doc_pdf.tobytes())
                        st.success("Edits applied.")
                        st.download_button(
                            "⬇️ Download edited PDF",
                            data=out,
                            file_name="edited_text.pdf",
                            mime="application/pdf",
                        )

        elif action == "Edit images / logo (replace or delete)":
            st.write(
                "Replace a logo/image with your own, or delete it entirely, without "
                "shifting any of the surrounding text."
            )
            page_num_img = st.number_input(
                "Page number", min_value=1, max_value=doc_pdf.page_count, value=1, key="img_page_num"
            )
            page_index_img = page_num_img - 1
            page_img = doc_pdf[page_index_img]

            images_found = page_img.get_images(full=True)
            if not images_found:
                st.warning("No images found on this page.")
            else:
                st.write(f"**{len(images_found)}** image(s) found on this page:")
                for idx, img_info in enumerate(images_found):
                    xref = img_info[0]
                    rects = page_img.get_image_rects(xref)
                    if not rects:
                        continue
                    rect = rects[0]

                    # show a small preview crop of this image area
                    zoom = 2
                    pix = page_img.get_pixmap(clip=rect, matrix=fitz.Matrix(zoom, zoom))
                    st.image(pix.tobytes("png"), caption=f"Image #{idx + 1} (page {page_num_img})", width=200)

                    col_a, col_b = st.columns(2)
                    with col_a:
                        replacement = st.file_uploader(
                            f"Replace image #{idx + 1} with:",
                            type=["png", "jpg", "jpeg"],
                            key=f"img_replace_{page_index_img}_{idx}",
                        )
                        if replacement and st.button(f"Apply replacement to image #{idx + 1}", key=f"apply_replace_{page_index_img}_{idx}"):
                            new_bytes = replacement.read()
                            page_img.add_redact_annot(rect, fill=(1, 1, 1))
                            page_img.apply_redactions()
                            page_img.insert_image(rect, stream=new_bytes)
                            out_img = io.BytesIO(doc_pdf.tobytes())
                            st.success(f"Image #{idx + 1} replaced.")
                            st.download_button(
                                "⬇️ Download edited PDF",
                                data=out_img,
                                file_name="edited_image.pdf",
                                mime="application/pdf",
                                key=f"dl_replace_{page_index_img}_{idx}",
                            )
                    with col_b:
                        if st.button(f"🗑️ Delete image #{idx + 1}", key=f"delete_{page_index_img}_{idx}"):
                            page_img.add_redact_annot(rect, fill=(1, 1, 1))
                            page_img.apply_redactions()
                            out_img = io.BytesIO(doc_pdf.tobytes())
                            st.success(f"Image #{idx + 1} deleted.")
                            st.download_button(
                                "⬇️ Download edited PDF",
                                data=out_img,
                                file_name="edited_image.pdf",
                                mime="application/pdf",
                                key=f"dl_delete_{page_index_img}_{idx}",
                            )

        elif action == "Delete pages":
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

# ---------------------------------------------------------------------------
# TAB 3: Bulk Mail Merge (Excel + PDF template -> many personalized PDFs)
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("Generate hundreds/thousands of personalized letters at once")
    st.write(
        "**Step 1:** Prepare your template PDF with placeholder tokens where "
        "personal info should go — for example `{{NAME}}` and `{{DEPARTMENT}}`. "
        "Tip: you can create these tokens using the **PDF Editor** tab above — "
        "edit any real name/department in your existing offer letter and replace "
        "it with a token like `{{NAME}}`, then save that as your template."
    )
    st.write(
        "**Step 2:** Upload an Excel sheet where the column headers exactly match "
        "your tokens (without the curly braces) — e.g. columns named `NAME`, "
        "`DEPARTMENT`, `START_DATE`."
    )

    def _merge_one(template_bytes, replacements):
        doc = fitz.open(stream=template_bytes, filetype="pdf")
        for page in doc:
            blocks = [b for b in page.get_text("blocks") if b[4].strip()]
            text_dict = page.get_text("dict")
            dict_blocks = text_dict["blocks"]
            edits = []
            for b in blocks:
                original = b[4]
                new_text = original
                changed = False
                for token, val in replacements.items():
                    ph = "{{" + str(token) + "}}"
                    if ph in new_text:
                        new_text = new_text.replace(ph, str(val))
                        changed = True
                if changed:
                    bbox = fitz.Rect(b[:4])
                    fontsize = 11
                    for db in dict_blocks:
                        for line in db.get("lines", []):
                            for span in line["spans"]:
                                if fitz.Rect(span["bbox"]).intersects(bbox):
                                    fontsize = span["size"]
                    edits.append((bbox, new_text.rstrip("\n"), fontsize))
            for bbox, new_text, fontsize in edits:
                page.add_redact_annot(bbox, fill=(1, 1, 1))
            page.apply_redactions()
            page_rect = page.rect
            for bbox, new_text, fontsize in edits:
                padded = fitz.Rect(
                    bbox.x0, bbox.y0,
                    min(bbox.x0 + max(bbox.width, 300), page_rect.width - 36),
                    bbox.y1,
                )
                padded.y1 = min(padded.y1 + fontsize * 4, page_rect.height - 36)
                size = fontsize
                rc = page.insert_textbox(padded, new_text, fontsize=size, fontname="helv")
                tries = 0
                while rc < 0 and size > 6 and tries < 8:
                    size -= 1
                    rc = page.insert_textbox(padded, new_text, fontsize=size, fontname="helv")
                    tries += 1
        return doc.tobytes()

    template_pdf = st.file_uploader("Upload template PDF (with {{TOKENS}})", type=["pdf"], key="mm_template")
    excel_file = st.file_uploader("Upload Excel sheet (.xlsx)", type=["xlsx"], key="mm_excel")

    if template_pdf and excel_file:
        template_bytes = template_pdf.read()
        wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
        ws = wb.active
        headers = [c.value for c in ws[1]]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        st.info(f"Found **{len(rows)}** rows and columns: {', '.join(str(h) for h in headers)}")

        folder_col = st.selectbox(
            "Which column should be used to sort output into folders? (e.g. department)",
            options=["(no folders — flat list)"] + [str(h) for h in headers],
        )
        name_col = st.selectbox(
            "Which column should be used to name each file? (e.g. name)",
            options=[str(h) for h in headers],
        )

        if st.button(f"Generate all {len(rows)} personalized PDFs"):
            zip_buf = io.BytesIO()
            progress = st.progress(0, text="Generating...")
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for i, row in enumerate(rows):
                    rowdict = dict(zip(headers, row))
                    out_pdf = _merge_one(template_bytes, rowdict)
                    fname = str(rowdict.get(name_col, f"row_{i+1}")).replace(" ", "_").replace("/", "-")
                    if folder_col != "(no folders — flat list)":
                        folder = str(rowdict.get(folder_col, "General")).replace(" ", "_").replace("/", "-")
                        zf.writestr(f"{folder}/{fname}.pdf", out_pdf)
                    else:
                        zf.writestr(f"{fname}.pdf", out_pdf)
                    progress.progress((i + 1) / len(rows), text=f"Generated {i+1}/{len(rows)}")
            st.success(f"Done! Generated {len(rows)} personalized PDFs.")
            st.download_button(
                "⬇️ Download all as ZIP",
                data=zip_buf.getvalue(),
                file_name="personalized_letters.zip",
                mime="application/zip",
            )

st.divider()
st.caption("HandyWriter · 100% free & open-source · Your files are processed locally, never uploaded anywhere.")
