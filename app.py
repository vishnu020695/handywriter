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

st.set_page_config(page_title="HandyWriter Ultimate", page_icon="✍️", layout="wide")
st.title("✍️ HandyWriter Ultimate")
st.caption("Certificates, Offer Letters, Letters என 2000+ மாணவர்களின் அனைத்து ஃபார்மட்களையும் அலைன்மென்ட் மாறாமல் பல்க்காக தயாரிக்கும் ஒரே கருவி.")

tab1, tab2, tab3 = st.tabs(["📝 Image → Word / Excel", "📄 Single PDF Editor", "📬 Universal Bulk Merge (All Formats)"])

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
# TAB 2: Single PDF Editor (கோட்டை உடைக்காமல் மேலே எழுதும் பிக்சல் முறை - FIXED)
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("ஒற்றை PDF / ஆஃபர் லெட்டரில் எதை வேண்டுமானாலும் திருத்திக் கொள்ளும் வசதி")
    st.write("இந்த முறையில் ஒரிஜினல் PDF-ல் உள்ள கோடுகள் உடையாது. அதற்கு மேல் (Overlay) டெக்ஸ்ட் பர்ஃபெக்ட்டாக அமரும்.")
    
    uploaded_pdf = st.file_uploader("Upload Single PDF / Offer Letter", type=["pdf"], key="pdf_upload_single")
    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.info(f"Loaded PDF with {doc_pdf.page_count} pages.")
        
        page_num = st.number_input("Page number to edit", min_value=1, max_value=doc_pdf.page_count, value=1, key="s_page")
        page = doc_pdf[page_num - 1]
        
        # PDF பக்கத்தை இமேஜாக மாற்றி துல்லியமான பிக்சல் அலைன்மென்ட் செய்கிறோம்
        zoom = 2.0  
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        page_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        W, H = page_img.size

        # Session State பயன்படுத்தி எத்தனை கஸ்டம் வரிகள் வேண்டுமானாலும் உருவாக்கிக் கொள்ளலாம்
        if "single_fields" not in st.session_state:
            st.session_state.single_fields = ["FIELD_1", "FIELD_2"]

        st.write("### 🛠️ எடிட் செய்ய வேண்டிய புதிய ஃபீல்டுகளை உருவாக்கவும்")
        col_add, col_rem = st.columns(2)
        with col_add:
            f_name = st.text_input("பாக்ஸ் பெயர் (எ.கா: NAME, DEPT, COURSE):").upper().strip()
            if st.button("➕ Add Box") and f_name:
                if f_name not in st.session_state.single_fields:
                    st.session_state.single_fields.append(f_name)
                    st.rerun()
        with col_rem:
            f_rem = st.selectbox("நீக்க வேண்டிய பாக்ஸ்:", options=["-- Select --"] + st.session_state.single_fields)
            if st.button("🗑️ Remove Box") and f_rem != "-- Select --":
                st.session_state.single_fields.remove(f_rem)
                st.rerun()

        st.markdown("---")
        st.write("### 1. பொசிஷன் மற்றும் டைப் செய்ய வேண்டிய விபரங்கள்")
        
        single_positions = {}
        for index, field in enumerate(st.session_state.single_fields):
            st.markdown(f"**📍 Configuration for: `{field}`**")
            cx, cy, cs, ca, ct = st.columns([2, 2, 2, 2, 4])
            with cx:
                x_pos = st.number_input(f"X (Idathu/Valathu) - {field}", min_value=0, max_value=W, value=int(W/2), key=f"sf_x_{field}")
            with cy:
                y_pos = st.number_input(f"Y (Mele/Keezhe) - {field}", min_value=0, max_value=H, value=int(H/2) + (index * 70) - 100, key=f"sf_y_{field}")
            with cs:
                f_size = st.number_input(f"Font Size - {field}", min_value=10, max_value=150, value=32, key=f"sf_s_{field}")
            with ca:
                align_type = st.selectbox(f"Align - {field}", options=["Center", "Left"], key=f"sf_a_{field}")
            with ct:
                text_val = st.text_input(f"Enter Text to print over line:", value="", key=f"sf_v_{field}", placeholder="இங்கு டைப் செய்யவும்...")
            
            single_positions[field] = {"x": x_pos, "y": y_pos, "size": f_size, "align": align_type, "text": text_val}
            st.markdown("<br>", unsafe_with_html=True)

        if st.button("Apply Text on PDF and Download"):
            draw = ImageDraw.Draw(page_img)
            
            for field, data in single_positions.items():
                val_to_print = data["text"]
                if val_to_print:
                    try:
                        font = ImageFont.truetype("LiberationSans-Regular.ttf", data["size"])
                    except:
                        try:
                            font = ImageFont.truetype("DejaVuSans.ttf", data["size"])
                        except:
                            font = ImageFont.load_default()
                    
                    left, top, right, bottom = draw.textbbox((0, 0), val_to_print, font=font)
                    text_width = right - left
                    
                    if data["align"] == "Center":
                        final_x = data["x"] - (text_width / 2)
                    else:
                        final_x = data["x"]
                        
                    # கோட்டின் மேல் (Overlay) பர்ஃபெக்ட்டாக பிரிண்ட் செய்கிறது
                    draw.text((final_x, data["y"]), val_to_print, fill=(0, 0, 0), font=font)

            pdf_buffer = io.BytesIO()
            page_img.save(pdf_buffer, format="PDF")
            
            st.success("டெக்ஸ்ட் கோடுகளுக்கு மேல் துல்லியமாகப் பொருத்தப்பட்டது!")
            st.download_button("⬇️ Download Perfect PDF", data=pdf_buffer.getvalue(), file_name="perfect_edited.pdf", mime="application/pdf", use_container_width=True)

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
# TAB 3: Universal Bulk Merge
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("2000+ மாணவர்களின் சான்றிதழ் / ஆஃபர் லெட்டர் பல்க் தயாரிப்பு")
    st.write("உங்கள் எக்செல் ஷீட்டில் நீங்கள் என்னென்ன பத்திகள் (Columns) வைத்துள்ளீர்களோ, அவை அனைத்தையும் இந்த சிஸ்டம் தானாகவே கண்டறிந்து மாற்றித் தரும்.")
    
    st.write("### 1. மாதிரி எக்செல் கோப்பு (Demo Template)")
    is_offer = st.checkbox("ஆஃபர் லெட்டருக்கான மாதிரி எக்செல் தேவையா? (Check for Offer Letter Template)")
    if is_offer:
        headers_demo = ["NAME", "ROLL_NO", "DEPARTMENT", "OFFER_DATE", "SALARY", "JOIN_DATE"]
        row1 = ["AJAY K", "221AI005", "B.Sc. AIML", "07-July-2026", "Rs. 25,000", "01-August-2026"]
    else:
        headers_demo = ["NAME", "DEPARTMENT", "COURSE_NAME", "DURATION", "ROLL_NO"]
        row1 = ["AJAY K", "B.Sc. AIML", "Cognitive Skills Enhancement - I", "June 2024 - November 2024", "221AI005"]
        
    demo_wb = openpyxl.Workbook()
    demo_ws = demo_wb.active
    demo_ws.title = "BulkData"
    demo_wb.append(headers_demo)
    demo_wb.append(row1)
    
    demo_buf = io.BytesIO()
    demo_wb.save(demo_buf)
    demo_buf.seek(0)
    st.download_button("⬇️ Download Sample Excel Template (.xlsx)", data=demo_buf, file_name="handywriter_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")
    st.write("### 2. ஆவணத்தின் வெற்றுப் படம் மற்றும் எக்செல் அப்லோடு")
    
    template_file = st.file_uploader("Upload Blank Template Image (PNG or JPG only)", type=["png", "jpg", "jpeg"])
    excel_file = st.file_uploader("Upload Student Excel Sheet (.xlsx)", type=["xlsx"], key="main_bulk_excel")

    if template_file and excel_file:
        base_image = Image.open(template_file).convert("RGB")
        W, H = base_image.size
        st.success(f"Template Loaded Successfully! Size: {W}x{H} Pixels.")
        
        wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
        ws = wb.active
        headers = [c.value for c in ws[1] if c.value is not None]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        
        st.info(f"எக்செல் ஷீட்டில் கண்டறியப்பட்ட பத்திகள் (Columns): {', '.join(headers)}")
        st.write(f"மொத்தம் **{len(rows)}** மாணவர்களின் விபரங்கள் உள்ளன.")
        
        st.write("### 3. அлайнமென்ட் செட்டிங்ஸ் (X, Y Coordinates for EVERY Column)")
        
        positions = {}
        for index, h in enumerate(headers):
            st.markdown(f"**⚙️ Settings for column: `{{{{{h}}}}}`**")
            cx, cy, cs, ca = st.columns([2, 2, 2, 2])
            with cx:
                x_pos = st.number_input(f"X Position - {h}", min_value=0, max_value=W, value=int(W/2), key=f"bx_{h}")
            with cy:
                y_pos = st.number_input(f"Y Position - {h}", min_value=0, max_value=H, value=int(H/2) + (index * 70) - 100, key=f"by_{h}")
            with cs:
                f_size = st.number_input(f"Font Size - {h}", min_value=10, max_value=200, value=32, key=f"bs_{h}")
            with ca:
                align_type = st.selectbox(f"Text Align - {h}", options=["Center", "Left"], key=f"ba_{h}")
            
            positions[h] = {"x": x_pos, "y": y_pos, "size": f_size, "align": align_type}
            st.markdown("<br>", unsafe_with_html=True)
            
        name_col = st.selectbox("கோப்புகளுக்கு பெயரிட எந்த பத்தியைப் பயன்படுத்த வேண்டும்? (File Name)", options=headers)

        if st.button(f"🚀 Generate All {len(rows)} Custom PDF Documents"):
            zip_buf = io.BytesIO()
            progress = st.progress(0, text="Generating documents...")

            with zipfile.ZipFile(zip_buf, "w") as zf:
                for idx, row in enumerate(rows):
                    rowdict = dict(zip(headers, row))
                    img_copy = base_image.copy()
                    draw = ImageDraw.Draw(img_copy)
                    
                    for h, pos in positions.items():
                        text_val = str(rowdict.get(h, "") if rowdict.get(h) is not None else "")
                        
                        try:
                            font = ImageFont.truetype("LiberationSans-Regular.ttf", pos["size"])
                        except:
                            try:
                                font = ImageFont.truetype("DejaVuSans.ttf", pos["size"])
                            except:
                                font = ImageFont.load_default()
                        
                        left, top, right, bottom = draw.textbbox((0, 0), text_val, font=font)
                        text_width = right - left
                        
                        if pos["align"] == "Center":
                            final_x = pos["x"] - (text_width / 2)
                        else:
                            final_x = pos["x"]
                            
                        draw.text((final_x, pos["y"]), text_val, fill=(0, 0, 0), font=font)
                    
                    pdf_out = io.BytesIO()
                    img_copy.save(pdf_out, format="PDF")
                    
                    fname = str(rowdict.get(name_col, f"student_{idx+1}")).replace(" ", "_").replace("/", "-")
                    zf.writestr(f"{fname}.pdf", pdf_out.getvalue())
                    
                    progress.progress((idx + 1) / len(rows), text=f"Generated {idx+1}/{len(rows)}")
                    
            st.success("அனைத்து ஆவணங்களும் எரர் இல்லாமல் வெற்றிகரமாகத் தயார் செய்யப்பட்டுவிட்டன!")
            st.download_button("⬇️ Download All PDFs as ZIP", data=zip_buf.getvalue(), file_name="bulk_custom_documents.zip", mime="application/zip", use_container_width=True)

st.divider()
st.caption("HandyWriter Ultimate · Safe, Local Processing inside Streamlit Cloud.")
