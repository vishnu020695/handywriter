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

st.set_page_config(page_title="HandyWriter Pro Ultimate", page_icon="✍️", layout="wide")
st.title("✍️ HandyWriter Pro Ultimate")
st.caption("Certificates, Offer Letters என அனைத்தையும் கோடுகளுக்கு மேல் (Overlay) அலைன்மென்ட் மாறாமல் பர்ஃபெக்ட்டாக தயாரிக்கும் இன்ஜின்.")

tab1, tab2, tab3 = st.tabs(["📝 Image → Word / Excel", "📄 Single PDF Overlayer (No Line Break)", "📬 Universal Bulk Merge (2000+ Students)"])

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
# TAB 2: Single PDF Overlayer (கோட்டை உடைக்காமல் மேலே எழுதும் புதிய வசதி)
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("கோட்டின் மேல் (Overlay) துல்லியமாக எழுதும் சிங்கிள் எடிட்டர்")
    st.write("இந்த முறையில் ஒரிஜினல் PDF-ல் இருக்கும் கோடுகள் அப்படியே இருக்கும், அதற்கு மேல் டெக்ஸ்ட் பர்ஃபெக்ட்டாக உட்காரும்.")
    
    uploaded_pdf = st.file_uploader("Upload Your Template PDF", type=["pdf"], key="pdf_single_perfect")
    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        page_num = st.number_input("Page number", min_value=1, max_value=doc_pdf.page_count, value=1)
        page = doc_pdf[page_num - 1]
        
        # Convert PDF to high-res image canvas so lines are never broken
        zoom = 2.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        page_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        W, H = page_img.size
        
        st.write("### 1. என்னென்ன விபரங்கள் கோட்டின் மேல் எழுதப்பட வேண்டும்?")
        
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            val_name = st.text_input("Enter Student Name:", "AJAY K")
            val_dept = st.text_input("Enter Department:", "B.Sc. AIML")
        with col_in2:
            val_course = st.text_input("Enter Course Name (Optional):", "Cognitive Skills Enhancement - I")
            val_duration = st.text_input("Enter Duration/Date (Optional):", "June 2024 - November 2024")

        st.markdown("---")
        st.write("### 2. அலைன்மென்ட் மற்றும் பொசிஷன் செட்டிங்ஸ் (X, Y Coordinates)")
        st.info("மாற்றங்களைச் சரிபார்க்க கீழே உள்ள எண்களை மாற்றி பொசிஷன் செய்யவும்.")

        fields = {
            "Student Name": {"val": val_name, "default_y": int(H*0.52), "size": 32},
            "Department": {"val": val_dept, "default_y": int(H*0.56), "size": 28},
            "Course Name": {"val": val_course, "default_y": int(H*0.60), "size": 26},
            "Duration/Date": {"val": val_duration, "default_y": int(H*0.64), "size": 26}
        }
        
        final_positions = {}
        for f_name, f_info in fields.items():
            if f_info["val"]:
                st.markdown(f"**📍 Position Setup for: {f_name}**")
                cx, cy, cs = st.columns(3)
                with cx:
                    x_p = st.number_input(f"X (Horizontal Center) - {f_name}", min_value=0, max_value=W, value=int(W/2), key=f"s_x_{f_name}")
                with cy:
                    y_p = st.number_input(f"Y (Vertical Height) - {f_name}", min_value=0, max_value=H, value=f_info["default_y"], key=f"s_y_{f_name}")
                with cs:
                    s_p = st.number_input(f"Font Size - {f_name}", min_value=10, max_value=150, value=f_info["size"], key=f"s_s_{f_name}")
                
                final_positions[f_name] = {"x": x_p, "y": y_p, "size": s_p, "val": f_info["val"]}

        if st.button("Apply Text over PDF Lines and Download"):
            img_copy = page_img.copy()
            draw = ImageDraw.Draw(img_copy)
            
            # Use standard cloud font fallback
            font = ImageFont.load_default()

            for f_name, data in final_positions.items():
                txt = data["val"]
                left, top, right, bottom = draw.textbbox((0, 0), txt, font=font)
                tw = right - left
                # Draws text directly over the line centered
                draw.text((data["x"] - (tw/2), data["y"]), txt, fill=(0, 0, 0), font=font)
            
            pdf_buffer = io.BytesIO()
            img_copy.save(pdf_buffer, format="PDF")
            
            st.success("கோட்டிற்கு மேல் டெக்ஸ்ட் பர்ஃபெக்ட்டாக அலைன் செய்யப்பட்டது!")
            st.download_button("⬇_ Download Perfect PDF", data=pdf_buffer.getvalue(), file_name="perfect_edited.pdf", mime="application/pdf")

# ---------------------------------------------------------------------------
# TAB 3: Universal Bulk Merge (2000+ மாணவர்களுக்கும் அலைன்மென்ட் மாறாத பல்க் இன்ஜின்)
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("2000+ மாணவர்களின் சான்றிதழ்கள் / ஆஃபர் லெட்டர்கள் தயாரிப்பு")
    st.write("எக்செல் ஷீட்டில் உள்ள அத்தனை விபரங்களையும் கோடுகளை உடைக்காமல் பல்க்காக அச்சிடலாம்.")
    
    st.write("### 1. மாதிரி எக்செல் கோப்பு (Demo Template)")
    demo_wb = openpyxl.Workbook()
    demo_ws = demo_wb.active
    demo_ws.title = "BulkData"
    demo_ws.append(["NAME", "DEPARTMENT", "COURSE_NAME", "DURATION", "ROLL_NO"])
    demo_ws.append(["AJAY K", "B.Sc. AIML", "Cognitive Skills Enhancement - I", "June 2024 - November 2024", "221AI005"])
    demo_ws.append(["ANITHA VISHNU VINISH", "Information Technology", "Cognitive Skills Enhancement - I", "June 2024 - November 2024", "221IT001"])
    
    demo_buf = io.BytesIO()
    demo_wb.save(demo_buf)
    demo_buf.seek(0)
    st.download_button("⬇_ Download Universal Demo Excel (.xlsx)", data=demo_buf, file_name="handywriter_bulk_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")
    st.write("### 2. ஆவணத்தின் வெற்றுப் படம் மற்றும் எக்செல் அப்லோடு")
    st.info("முக்கிய குறிப்பு: PDF-ஐ 'PDF to JPG' ஆன்லைன் கன்வெர்ட்டர் மூலம் படமாக மாற்றி இங்கே அப்லோட் செய்யவும்.")
    
    template_file = st.file_uploader("Upload Blank Template Image (PNG or JPG format)", type=["png", "jpg", "jpeg"], key="bulk_template_img")
    excel_file = st.file_uploader("Upload Student Filled Excel Sheet (.xlsx)", type=["xlsx"], key="main_bulk_excel")

    if template_file and excel_file:
        base_image = Image.open(template_file).convert("RGB")
        W, H = base_image.size
        st.success(f"Template Loaded! Resolution: {W}x{H} Pixels.")
        
        wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
        ws = wb.active
        headers = [c.value for c in ws[1] if c.value is not None]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        
        st.info(f"கண்டறியப்பட்ட பத்திகள் (Columns): {', '.join(headers)} (மொத்தம் {len(rows)} மாணவர்கள்)")
        
        st.write("### 3. அலைன்மென்ட் செட்டிங்ஸ் (X, Y Coordinates)")
        
        positions = {}
        for index, h in enumerate(headers):
            st.markdown(f"**⚙_ Position for column: `{{{{{h}}}}}`**")
            cx, cy, cs, ca = st.columns([2, 2, 2, 2])
            with cx:
                x_pos = st.number_input(f"X (Horizontal) - {h}", min_value=0, max_value=W, value=int(W/2), key=f"bx_{h}")
            with cy:
                y_pos = st.number_input(f"Y (Vertical Height) - {h}", min_value=0, max_value=H, value=int(H*0.5) + (index * 50), key=f"by_{h}")
            with cs:
                f_size = st.number_input(f"Font Size - {h}", min_value=10, max_value=200, value=32, key=f"bs_{h}")
            with ca:
                align_type = st.selectbox(f"Alignment - {h}", options=["Center", "Left"], key=f"ba_{h}")
            
            positions[h] = {"x": x_pos, "y": y_pos, "size": f_size, "align": align_type}
            st.markdown("<br>", unsafe_with_html=True)
            
        name_col = st.selectbox("கோப்புகளுக்கு பெயரிட எந்த பத்தியைப் பயன்படுத்த வேண்டும்? (File Name Field)", options=headers)

        if st.button(f"🚀 Generate All {len(rows)} Perfect Documents"):
            zip_buf = io.BytesIO()
            progress = st.progress(0, text="Generating documents...")
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
                        
                        if pos["align"] == "Center":
                            final_x = pos["x"] - (text_width / 2)
                        else:
                            final_x = pos["x"]
                            
                        draw.text((final_x, pos["y"]), text_val, fill=(30, 30, 30), font=font)
                    
                    pdf_out = io.BytesIO()
                    img_copy.save(pdf_out, format="PDF")
                    
                    fname = str(rowdict.get(name_col, f"doc_{idx+1}")).replace(" ", "_").replace("/", "-")
                    zf.writestr(f"{fname}.pdf", pdf_out.getvalue())
                    
                    progress.progress((idx + 1) / len(rows), text=f"Generated {idx+1}/{len(rows)}")
                    
            st.success("அனைத்து ஆவணங்களும் அலைன்மென்ட் மாறாமல் பர்ஃபெக்ட்டாகத் தயார் செய்யப்பட்டுவிட்டன!")
            st.download_button("⬇_ Download All Bulk PDFs as ZIP", data=zip_buf.getvalue(), file_name="bulk_documents.zip", mime="application/zip", use_container_width=True)

st.divider()
st.caption("HandyWriter Pro Ultimate · Secure, Cloud-ready local rendering engine.")
