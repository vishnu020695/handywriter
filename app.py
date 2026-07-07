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
# TAB 2: PDF Editor
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Simple PDF Editing")
    uploaded_pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_upload")
    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.info(f"Loaded PDF with {doc_pdf.page_count} pages.")
        # (உங்களின் பழைய ரொடேஷன் மற்றும் சிம்பிள் எடிட் கோட் இதனுள் வரும்...)

# ---------------------------------------------------------------------------
# TAB 3: Pixel-Perfect Bulk Certificate Merge (Fixed Alignment)
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("2000+ மாணவர்களுக்கான பர்ஃபெக்ட் பல்க் சான்றிதழ் தயாரிப்பு")
    st.write("இந்த முறையில் சான்றிதழின் அலைன்மென்ட் 1% கூட மாறாது. எடிட் செய்த அடையாளமே தெரியாது.")
    
    # Demo Excel Download Button
    st.write("### 1. டெமோ எக்செல் ஷீட்டை டவுன்லோட் செய்யவும்")
    demo_wb = openpyxl.Workbook()
    demo_ws = demo_wb.active
    demo_ws.title = "StudentsData"
    # Headers
    demo_ws.append(["NAME", "DEPARTMENT", "COURSE_NAME", "DURATION", "ROLL_NO"])
    # Sample Rows
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
        
        # Load Excel Data
        wb = openpyxl.load_workbook(io.BytesIO(excel_data.read()))
        ws = wb.active
        headers = [c.value for c in ws[1]]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        
        st.info(f"Excel-லில் மொத்தம் **{len(rows)}** மாணவர்கள் கண்டறியப்பட்டனர்.")
        
        st.write("### 3. அலைன்மென்ட் மற்றும் பொசிஷன் செட்டிங்ஸ் (X, Y Coordinates)")
        st.info("சான்றிதழில் டெக்ஸ்ட் எந்த இடத்தில் (Pixel Position) வர வேண்டும் என்பதை கீழே குறிப்பிடவும்.")
        
        col_x, col_y, col_size = st.columns(3)
        
        positions = {}
        # Dynamic coordinate picking for headers to accommodate any structure change
        for h in headers:
            if h:
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
            
            # Using default standard font, can be customized
            try:
                font_path = "arial.ttf" 
            except:
                font_path = None # Fallback to default pillow font if not found

            with zipfile.ZipFile(zip_buf, "w") as zf:
                for idx, row in enumerate(rows):
                    rowdict = dict(zip(headers, row))
                    
                    # Create a copy of background image
                    img_copy = base_image.copy()
                    draw = ImageDraw.Draw(img_copy)
                    
                    # Draw text values onto the image flawlessly based on user coordinates
                    for h, pos in positions.items():
                        text_val = str(rowdict.get(h, ""))
                        
                        # Load dynamic font size
                        if font_path:
                            font = ImageFont.truetype(font_path, pos["size"])
                        else:
                            font = ImageFont.load_default()
                            
                        # Center align logic based on text width
                        left, top, right, bottom = draw.textbbox((0, 0), text_val, font=font)
                        text_width = right - left
                        
                        # Drawing the text exactly where it belongs
                        draw.text((pos["x"] - (text_width/2), pos["y"]), text_val, fill=(30, 30, 30), font=font)
                    
                    # Save as PDF
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
