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
st.caption("Convert images, edit PDFs, and generate perfect bulk certificates without alignment issues.")

tab1, tab2, tab3 = st.tabs(["📝 Image → Word / Excel", "📄 PDF Editor", "📬 Bulk Certificate Merge"])

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

        with st.spinner("Reading text..."):
            extracted_text = pytesseract.image_to_string(image, config="--psm 6")

        with col2:
            edited_text = st.text_area("Editable text", extracted_text, height=350)

        st.write("### Save as:")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Create Word (.docx)", use_container_width=True):
                doc = Document()
                for line in edited_text.split("\n"):
                    doc.add_paragraph(line)
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button("⬇️ Download Word file", data=buf, file_name="converted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

        with c2:
            if st.button("💾 Create Excel (.xlsx)", use_container_width=True):
                wb = openpyxl.Workbook()
                ws = wb.active
                for line in edited_text.split("\n"):
                    if not line.strip(): continue
                    cells = re.split(r"\t|\s{2,}", line.strip())
                    ws.append(cells)
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.download_button("⬇️ Download Excel file", data=buf, file_name="converted.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ---------------------------------------------------------------------------
# TAB 2: PDF Editor (ஒற்றை PDF-ஐ எடிட் செய்யும் வசதி சேர்க்கப்பட்டது)
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Simple PDF Editing")
    uploaded_pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_upload")
    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.info(f"Loaded PDF with {doc_pdf.page_count} pages.")
        
        action = st.selectbox(
            "Choose an action",
            [
                "Edit page text (direct, in-place)",
                "Delete pages",
                "Rotate pages",
                "Add text watermark"
            ],
        )

        if action == "Edit page text (direct, in-place)":
            st.write("PDF-ல் மாற்ற வேண்டிய டெக்ஸ்ட் பாக்ஸ்களை கீழே திருத்திவிட்டு Apply செய்யவும்.")
            page_num = st.number_input("Page number to edit", min_value=1, max_value=doc_pdf.page_count, value=1)
            page_index = page_num - 1
            page = doc_pdf[page_index]

            raw_blocks = page.get_text("blocks")
            blocks = [b for b in raw_blocks if b[4].strip()]

            # Preview Block Numbers
            zoom = 1.3
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            preview_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            draw = ImageDraw.Draw(preview_img)
            for i, b in enumerate(blocks):
                x0, y0, x1, y1 = [v * zoom for v in b[:4]]
                draw.rectangle([x0, y0, x1, y1], outline="red", width=2)
                draw.text((x0, max(0, y0 - 14)), f"#{i + 1}", fill="red")
            st.image(preview_img, caption=f"Page {page_num} — Preview with Box Numbers", width=500)

            if not blocks:
                st.warning("இந்த PDF பக்கத்தில் எடிட் செய்யக்கூடிய டெக்ஸ்ட் எதுவும் கிடைக்கவில்லை. இது ஸ்கேன் செய்யப்பட்ட படமாக இருக்கலாம்.")
            else:
                st.write(f"**{len(blocks)}** எடிட் செய்யக்கூடிய டெக்ஸ்ட் பாக்ஸ்கள் கண்டறியப்பட்டன:")
                edited_values = []
                for i, b in enumerate(blocks):
                    original_text = b[4].rstrip("\n")
                    val = st.text_area(f"Box #{i + 1}", original_text, height=68, key=f"block_{page_index}_{i}")
                    edited_values.append(val)

                if st.button("Apply edits to this page and download"):
                    text_dict = page.get_text("dict")
                    dict_blocks = text_dict["blocks"]

                    changed_any = False
                    edits = []
                    for i, (b, new_val) in enumerate(zip(blocks, edited_values)):
                        original_text = b[4].rstrip("\n")
                        if new_val != original_text:
                            changed_any = True
                            bbox = fitz.Rect(b[:4])
                            fontsize = 11
                            for db in dict_blocks:
                                for line in db.get("lines", []):
                                    for span in line["spans"]:
                                        if fitz.Rect(span["bbox"]).intersects(bbox):
                                            fontsize = span["size"]
                                            break
                            edits.append((bbox, new_val, fontsize))

                    if not changed_any:
                        st.info("மாற்றங்கள் எதுவும் செய்யப்படவில்லை.")
                    else:
                        for bbox, new_val, fontsize in edits:
                            page.add_redact_annot(bbox, fill=(1, 1, 1))
                        page.apply_redactions()
                        page_rect = page.rect
                        for bbox, new_val, fontsize in edits:
                            padded = fitz.Rect(
                                bbox.x0, bbox.y0,
                                min(bbox.x0 + max(bbox.width, 300), page_rect.width - 36),
                                bbox.y1,
                            )
                            padded.y1 = min(padded.y1 + fontsize * 4, page_rect.height - 36)
                            rc = page.insert_textbox(padded, new_val, fontsize=fontsize, fontname="helv")
                        
                        out = io.BytesIO(doc_pdf.tobytes())
                        st.success("மாற்றங்கள் வெற்றிகரமாகச் சேர்க்கப்பட்டன!")
                        st.download_button("⬇️ Download edited PDF", data=out, file_name="edited_text.pdf", mime="application/pdf")

        elif action == "Delete pages":
            pages_to_delete = st.text_input("Page numbers to delete (comma-separated, e.g. 1,3,5)")
            if st.button("Apply and download"):
                try:
                    nums = [int(x.strip()) - 1 for x in pages_to_delete.split(",") if x.strip()]
                    doc_pdf.delete_pages(nums)
                    out = io.BytesIO(doc_pdf.tobytes())
                    st.download_button("⬇️ Download edited PDF", data=out, file_name="edited.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Check your page numbers. Details: {e}")

        elif action == "Rotate pages":
            angle = st.selectbox("Rotate by", [90, 180, 270])
            if st.button("Apply and download"):
                for page in doc_pdf:
                    page.set_rotation(angle)
                out = io.BytesIO(doc_pdf.tobytes())
                st.download_button("⬇️ Download rotated PDF", data=out, file_name="rotated.pdf", mime="application/pdf")

        elif action == "Add text watermark":
            wm_text = st.text_input("Watermark text", "CONFIDENTIAL")
            if st.button("Apply and download"):
                for page in doc_pdf:
                    rect = page.rect
                    center = fitz.Point(rect.width / 2, rect.height / 2)
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
                st.download_button("⬇️ Download watermarked PDF", data=out, file_name="watermarked.pdf", mime="application/pdf")

# ---------------------------------------------------------------------------
# TAB 3: Pixel-Perfect Bulk Certificate Merge (Fixed Alignment)
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("2000+ மாணவர்களுக்கான பர்ஃபெக்ட் பல்க் சான்றிதழ் தயாரிப்பு")
    st.write("இந்த முறையில் சான்றிதழின் அலைன்மென்ட் 1% கூட மாறாது. எடிட் செய்த அடையாளமே தெரியாது.")
    
    st.write("### 1. டெமோ எக்செல் ஷீட்டை டவுன்லோட் செய்யவும்")
    demo_wb = openpyxl.Workbook()
    demo_ws = demo_wb.active
    demo_ws.title = "StudentsData"
    demo_ws.append(["NAME", "DEPARTMENT", "COURSE_NAME", "DURATION", "ROLL_NO"])
    demo_ws.append(["ANITHA VISHNU VINISH", "Information Technology", "Cognitive Skills Enhancement - I", "June 2024 - November 2024", "221IT001"])
    demo_ws.append(["KUMAR RAJ", "Computer Science", "Cognitive Skills Enhancement - I", "June 2024 - November 2024", "221CS045"])
    
    demo_buf = io.BytesIO()
    demo_wb.save(demo_buf)
    demo_buf.seek(0)
    st.download_button(
        label="⬇️ Download Demo Excel Sheet (.xlsx)",
        data=demo_buf,
        file_name="handywriter_demo_students.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("---")
    st.write("### 2. கோப்புகளை அப்லோட் செய்யவும்")
    
    cert_bg = st.file_uploader("Upload Certificate Blank Template (Image - PNG/JPG format ONLY for 100% alignment)", type=["png", "jpg", "jpeg"])
    excel_data = st.file_uploader("Upload Filled Excel Sheet (.xlsx)", type=["xlsx"], key="bulk_excel")

    if cert_bg and excel_data:
        base_image = Image.open(cert_bg).convert("RGB")
        W, H = base_image.size
        
        st.success(f"Template Loaded Successfully! Size: {W}x{H} pixels.")
        
        wb = openpyxl.load_workbook(io.BytesIO(excel_data.read()))
        ws = wb.active
        headers = [c.value for c in ws[1] if c.value is not None]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        
        st.info(f"Excel-லில் மொத்தம் **{len(rows)}** மாணவர்கள் கண்டறியப்பட்டனர்.")
        
        st.write("### 3. அலைன்மென்ட் மற்றும் பொசிஷன் செட்டிங்ஸ் (X, Y Coordinates)")
        st.info("சான்றிதழில் டெக்ஸ்ட் எந்த இடத்தில் (Pixel Position) வர வேண்டும் என்பதை கீழே குறிப்பிடவும்.")
        
        positions = {}
        for h in headers:
            st.markdown(f"**For Variable: `{{{{{h}}}}}`**")
            c_x, c_y, c_s = st.columns(3)
            with c_x:
                x_pos = st.number_input(f"X (Horizontal Position) for {h}", min_value=0, max_value=W, value=int(W/2), key=f"x_{h}")
            with c_y:
                y_pos = st.number_input(f"Y (Vertical Position) for {h}", min_value=0, max_value=H, value=int(H/2), key=f"y_{h}")
            with c_s:
                f_size = st.number_input(f"Font Size for {h}", min_value=10, max_value=200, value=40, key=f"s_{h}")
            
            positions[h] = {"x": x_pos, "y": y_pos, "size": f_size}
        
        name_col = st.selectbox("மாணவர் கோப்பின் பெயராக எதை வைக்க வேண்டும்? (File Name)", options=[str(h) for h in headers])

        if st.button(f"Generate All {len(rows)} Certificates"):
            zip_buf = io.BytesIO()
            progress = st.progress(0, text="Generating certificates...")
            
            font = ImageFont.load_default()

            with zipfile.ZipFile(zip_buf, "w") as zf:
                for idx, row in enumerate(rows):
                    rowdict = dict(zip(headers, row))
                    
                    img_copy = base_image.copy()
                    draw = ImageDraw.Draw(img_copy)
                    
                    for h, pos in positions.items():
                        text_val = str(rowdict.get(h, "") if rowdict.get(h) is not None else "")
                        
                        left, top, right, bottom = draw.textbbox((0, 0), text_val, font=font)
                        text_width = right - left
                        
                        draw.text((pos["x"] - (text_width/2), pos["y"]), text_val, fill=(30, 30, 30), font=font)
                    
                    pdf_out = io.BytesIO()
                    img_copy.save(pdf_out, format="PDF")
                    
                    fname = str(rowdict.get(name_col, f"student_{idx+1}")).replace(" ", "_")
                    zf.writestr(f"{fname}.pdf", pdf_out.getvalue())
                    
                    progress.progress((idx + 1) / len(rows), text=f"Generated {idx+1}/{len(rows)}")
            
            st.success("அனைத்து சான்றிதழ்களும் அலைன்மென்ட் தவறாமல் வெற்றிகரமாக உருவாக்கப்பட்டுவிட்டன!")
            st.download_button(
                "⬇️ Download All Certificates as ZIP",
                data=zip_buf.getvalue(),
                file_name="perfect_certificates.zip",
                mime="application/zip",
                use_container_width=True
            )

st.divider()
st.caption("HandyWriter Pro · Your files are processed safely and securely.")
