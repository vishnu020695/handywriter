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
# TAB 2: PDF Editor (கோட்டின் மேல் துல்லியமாக எழுதும் வகையில் மாற்றப்பட்டது)
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
                "Add Text Over Lines (Perfect Alignment)",
                "Delete pages",
                "Rotate pages"
            ],
        )

        if action == "Add Text Over Lines (Perfect Alignment)":
            st.write("PDF-ல் உள்ள கோடுகளுக்கு மேல் புதிய டெக்ஸ்ட்டை அலைன்மென்ட் மாறாமல் சேர்க்கலாம்.")
            page_num = st.number_input("Page number to edit", min_value=1, max_value=doc_pdf.page_count, value=1)
            page_index = page_num - 1
            page = doc_pdf[page_index]

            # Convert PDF page to image for coordinate plotting
            zoom = 2.0  # Higher zoom for high-res exact selection
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            page_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            W, H = page_img.size

            st.write("### 1. பொசிஷன் செட்டிங்ஸ் (X, Y Coordinates)")
            st.info("கோட்டிற்கு மேல் பெயர் மற்றும் துறை வர வேண்டிய இடத்தை பிக்சல் மதிப்பாக மாற்றவும்.")

            col_x1, col_y1, col_s1 = st.columns(3)
            with col_x1:
                student_name_x = st.number_input("Name X Position (Horizontal)", min_value=0, max_value=W, value=int(W/2))
            with col_y1:
                student_name_y = st.number_input("Name Y Position (Vertical)", min_value=0, max_value=H, value=int(H/2) - 50)
            with col_s1:
                font_size_name = st.number_input("Name Font Size", min_value=10, max_value=150, value=32)

            col_x2, col_y2, col_s2 = st.columns(3)
            with col_x2:
                dept_x = st.number_input("Department X Position (Horizontal)", min_value=0, max_value=W, value=int(W/2))
            with col_y2:
                dept_y = st.number_input("Department Y Position (Vertical)", min_value=0, max_value=H, value=int(H/2) + 50)
            with col_s2:
                font_size_dept = st.number_input("Department Font Size", min_value=10, max_value=150, value=28)

            st.write("### 2. உள்ளிட வேண்டிய விபரங்கள்")
            input_name = st.text_input("Enter Student Name (NAME)", "AJAY K")
            input_dept = st.text_input("Enter Department (DEPARTMENT)", "B.Sc. AIML")

            if st.button("Apply Text on PDF and Download"):
                # Work directly on a high-res image canvas to ensure exact overlaying without text box splitting
                draw = ImageDraw.Draw(page_img)
                font = ImageFont.load_default()

                # Write Name
                left, top, right, bottom = draw.textbbox((0, 0), input_name, font=font)
                w_name = right - left
                draw.text((student_name_x - (w_name/2), student_name_y), input_name, fill=(0, 0, 0), font=font)

                # Write Dept
                left, top, right, bottom = draw.textbbox((0, 0), input_dept, font=font)
                w_dept = right - left
                draw.text((dept_x - (w_dept/2), dept_y), input_dept, fill=(0, 0, 0), font=font)

                # Convert the modified high-res image directly back to a clean PDF page
                pdf_buffer = io.BytesIO()
                page_img.save(pdf_buffer, format="PDF")
                
                st.success("கோட்டிற்கு மேல் டெக்ஸ்ட் பர்ஃபெக்ட்டாக பொருத்தப்பட்டது!")
                st.download_button("⬇️ Download Perfect PDF", data=pdf_buffer.getvalue(), file_name="perfect_edited.pdf", mime="application/pdf")

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
        st.info("சான்றிதழில் டெக்ಸ್ಟ್ எந்த இடத்தில் (Pixel Position) வர வேண்டும் என்பதை கீழே குறிப்பிடவும்.")
        
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
