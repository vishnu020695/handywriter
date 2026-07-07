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
st.caption("Certificates, 4-Page Offer Letters, Documents என எதை வேண்டுமானாலும் எரர் மற்றும் அலைன்மென்ட் பிழையின்றி மாற்றும் ஒரே டூல்.")

tab1, tab2, tab3 = st.tabs(["📝 Image → Word / Excel", "📄 Intelligent PDF Editor (No Boxes)", "📬 Universal Bulk Merge (All Formats)"])

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
# TAB 2: Intelligent PDF Editor (நூற்றுக்கணக்கான பெட்டிகள் இல்லாத ஸ்மார்ட் எடிட்டர்)
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("🔍 தேடித் திருத்தும் ஸ்மார்ட் PDF எடிட்டர் (ஆஃபர் லெட்டர்களுக்கு மிகச்சிறந்தது)")
    st.write("PDF-ல் உள்ள நூற்றுக்கணக்கான பெட்டிகளைத் தேடத் தேவையில்லை. எந்த வார்த்தையை மாற்ற வேண்டுமோ அதை மட்டும் கீழே கொடுத்து மாற்றிக் கொள்ளலாம்.")
    
    uploaded_pdf = st.file_uploader("Upload Single PDF / Scanned Offer Letter", type=["pdf"], key="pdf_upload_single")
    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.success(f"Loaded PDF with {doc_pdf.page_count} pages.")
        
        st.write("### ✍️ எடிட் செய்ய வேண்டிய விபரங்கள் (நீங்கள் எத்தனை வரிகள் வேண்டுமானாலும் மாற்றலாம்):")
        
        # Dynamic Search and Replace fields using Streamlit session state
        if "replacements_count" not in st.session_state:
            st.session_state.replacements_count = 2

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("➕ இன்னும் ஒரு வார்த்தையை மாற்ற வேண்டுமா? (Add Word)"):
                st.session_state.replacements_count += 1
                st.rerun()
        with col_btn2:
            if st.button("🗑️ கடைசியாக சேர்த்ததை நீக்கு (Remove)") and st.session_state.replacements_count > 1:
                st.session_state.replacements_count -= 1
                st.rerun()

        change_dict = {}
        for idx in range(st.session_state.replacements_count):
            st.markdown(f"**வார்த்தை #{idx+1}**")
            c_find, c_replace = st.columns(2)
            with c_find:
                find_txt = st.text_input(f"மாற்ற வேண்டிய பழைய வார்த்தை (எ.கா: AJAY K அல்லது B.Sc. AIML):", key=f"f_{idx}")
            with c_replace:
                replace_txt = st.text_input(f"வர வேண்டிய புதிய வார்த்தை (எ.கா: VISHNU அல்லது B.Com):", key=f"r_{idx}")
            
            if find_txt.strip():
                change_dict[find_txt.strip()] = replace_txt.strip()

        if st.button("🚀 Apply Changes and Download PDF", use_container_width=True):
            if not change_dict:
                st.warning("மாற்றுவதற்கு எந்த வார்த்தையையும் நீங்கள் உள்ளிடவில்லை.")
            else:
                with st.spinner("ஆவணத்தில் உள்ள வரிகளைத் தேடி மாற்றியமைக்கிறது..."):
                    for page in doc_pdf:
                        for find_str, replace_str in change_dict.items():
                            # Find the exact text location in PDF coordinates
                            text_instances = page.search_for(find_str)
                            for rect in text_instances:
                                # Get exact font details before wiping out
                                text_dict = page.get_text("dict", clip=rect)
                                fontsize = 11
                                try:
                                    fontsize = text_dict["blocks"][0]["lines"][0]["spans"][0]["size"]
                                except:
                                    pass

                                # 1. Wipe out the old text cleanly
                                page.add_redact_annot(rect, fill=(1, 1, 1))
                                page.apply_redactions()
                                
                                # 2. Overlay new text on top of the original coordinate box perfectly
                                page.insert_textbox(rect, replace_str, fontsize=fontsize, fontname="helv", align=1)
                    
                    out_pdf = io.BytesIO(doc_pdf.tobytes())
                    st.success("அனைத்து பக்கங்களிலும் மாற்றங்கள் வெற்றிகரமாகச் செய்யப்பட்டன!")
                    st.download_button("⬇️ Download Final Edited PDF", data=out_pdf, file_name="perfect_edited.pdf", mime="application/pdf", use_container_width=True)

# ---------------------------------------------------------------------------
# TAB 3: Universal Bulk Merge (Fixed Engine)
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("2000+ மாணவர்களின் சான்றிதழ் / ஆஃபர் லெட்டர் பல்க் தயாரிப்பு")
    st.write("உங்கள் எக்செல் ஷீட்டில் நீங்கள் என்னென்ன பத்திகள் (Columns) வைத்துள்ளீர்களோ, அவை அனைத்தையும் இந்த சிஸ்டம் தானாகவே கண்டறிந்து மாற்றித் தரும்.")
    
    st.write("### 1. மாதிரி எக்செல் கோப்பு (Demo Template)")
    is_offer = st.checkbox("ஆஃபர் லெட்டருக்கான மாதிரி எக்செல் தேவையா? (Check for Offer Letter Template)")
    
    def generate_demo_excel(offer_mode):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BulkData"
        if offer_mode:
            ws.append(["NAME", "ROLL_NO", "DEPARTMENT", "OFFER_DATE", "SALARY", "JOIN_DATE"])
            ws.append(["AJAY K", "221AI005", "B.Sc. AIML", "07-July-2026", "Rs. 25,000", "01-August-2026"])
        else:
            ws.append(["NAME", "DEPARTMENT", "COURSE_NAME", "DURATION", "ROLL_NO"])
            ws.append(["AJAY K", "B.Sc. AIML", "Cognitive Skills Enhancement - I", "June 2024 - November 2024", "221AI005"])
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.getvalue()

    demo_data = generate_demo_excel(is_offer)
    st.download_button("⬇️ Download Sample Excel Template (.xlsx)", data=demo_data, file_name="handywriter_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
        
        st.info(f"Columns found: {', '.join(headers)}")
        
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
                    
            st.success("அனைத்து ஆவணங்களும் வெற்றிகரமாகத் தயார் செய்யப்பட்டுவிட்டன!")
            st.download_button("⬇️ Download All PDFs as ZIP", data=zip_buf.getvalue(), file_name="bulk_custom_documents.zip", mime="application/zip", use_container_width=True)

st.divider()
st.caption("HandyWriter Ultimate · Safe, Local Processing inside Streamlit Cloud.")
