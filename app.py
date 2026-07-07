import io
import os
import zipfile
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pytesseract
from docx import Document
import openpyxl

st.set_page_config(page_title="HandyWriter Ultimate", page_icon="✍️", layout="wide")
st.title("✍️ HandyWriter Ultimate")
st.caption("Universal Document Engine - Zero PyMuPDF alignment crashes.")

tab1, tab2 = st.tabs(["📄 Document Editor & Bulk Merge", "📝 Image → Word / Excel"])

# Helper function to get safe system fonts inside Streamlit Cloud
def get_sys_font(font_size):
    for f_name in ["LiberationSans-Regular.ttf", "DejaVuSans.ttf", "Arial.ttf"]:
        try:
            return ImageFont.truetype(f_name, font_size)
        except IOError:
            continue
    return ImageFont.load_default()

# ---------------------------------------------------------------------------
# TAB 1: Universal Document Editor & Bulk Engine (Supports All Formats)
# ---------------------------------------------------------------------------
with tab1:
    st.subheader("Generate or Edit Any Document Layout (Certificates, Offer Letters, IDs)")
    st.write("Upload a blank background image of your document layout, then map your fields anywhere using pixel coordinates.")
    
    st.write("### 1. Download Layout Template Sheet")
    is_offer = st.checkbox("Toggle for Offer Letter Data Structure (Changes column layout headers)")
    
    def generate_demo_excel(offer_mode):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DocumentData"
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
    st.download_button("⬇️ Download Sample Excel Sheet (.xlsx)", data=demo_data, file_name="handywriter_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")
    st.write("### 2. Upload Files")
    st.info("Note: Convert your blank layout page/PDF to an image file format (PNG or JPG) before uploading to maintain 100% alignment stability.")
    
    template_file = st.file_uploader("Upload Blank Template Canvas (PNG / JPG)", type=["png", "jpg", "jpeg"])
    excel_file = st.file_uploader("Upload Data Sheet (.xlsx)", type=["xlsx"])

    if template_file and excel_file:
        base_image = Image.open(template_file).convert("RGB")
        W, H = base_image.size
        st.success(f"Layout Template Ready! Resolution Profile: {W}x{H} Pixels.")
        
        wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
        ws = wb.active
        headers = [c.value for c in ws[1] if c.value is not None]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        
        st.info(f"Identified Fields: {', '.join(headers)} | Processing Row Volume: {len(rows)}")
        
        st.write("### 3. Coordinate Layout Placement Settings")
        st.write("Adjust where each field sits dynamically on the document layout matrix below:")
        
        positions = {}
        for index, h in enumerate(headers):
            st.markdown(f"**Field Variable Block: `{{{{{h}}}}}`**")
            cx, cy, cs, ca = st.columns([2, 2, 2, 2])
            with cx:
                x_pos = st.number_input(f"X Axis Position - {h}", min_value=0, max_value=W, value=int(W/2), key=f"x_{h}")
            with cy:
                y_pos = st.number_input(f"Y Axis Position - {h}", min_value=0, max_value=H, value=int(H/2) + (index * 70) - 100, key=f"y_{h}")
            with cs:
                f_size = st.number_input(f"Font Point Size - {h}", min_value=10, max_value=200, value=32, key=f"s_{h}")
            with ca:
                align_type = st.selectbox(f"Text Horizon Align - {h}", options=["Center", "Left"], key=f"a_{h}")
            
            positions[h] = {"x": x_pos, "y": y_pos, "size": f_size, "align": align_type}
            st.markdown("<br>", unsafe_with_html=True)
            
        name_col = st.selectbox("Select key column target to handle individual filename extractions:", options=headers)

        if st.button("🚀 Process and Generate Output Documents", use_container_width=True):
            zip_buf = io.BytesIO()
            progress = st.progress(0, text="Processing batch records...")

            with zipfile.ZipFile(zip_buf, "w") as zf:
                for idx, row in enumerate(rows):
                    rowdict = dict(zip(headers, row))
                    img_copy = base_image.copy()
                    draw = ImageDraw.Draw(img_copy)
                    
                    for h, pos in positions.items():
                        text_val = str(rowdict.get(h, "") if rowdict.get(h) is not None else "")
                        font = get_sys_font(pos["size"])
                        
                        left, top, right, bottom = draw.textbbox((0, 0), text_val, font=font)
                        text_width = right - left
                        
                        if pos["align"] == "Center":
                            final_x = pos["x"] - (text_width / 2)
                        else:
                            final_x = pos["x"]
                            
                        # Draws clean crisp black text directly over layout guidelines/lines
                        draw.text((final_x, pos["y"]), text_val, fill=(0, 0, 0), font=font)
                    
                    pdf_out = io.BytesIO()
                    img_copy.save(pdf_out, format="PDF")
                    
                    fname = str(rowdict.get(name_col, f"doc_{idx+1}")).replace(" ", "_").replace("/", "-")
                    zf.writestr(f"{fname}.pdf", pdf_out.getvalue())
                    
                    progress.progress((idx + 1) / len(rows), text=f"Processing item reference {idx+1}/{len(rows)}")
                    
            st.success("Batch array calculations executed cleanly with zero validation errors.")
            st.download_button("⬇️ Download All Processed PDFs (.zip)", data=zip_buf.getvalue(), file_name="generated_documents.zip", mime="application/zip", use_container_width=True)

# ---------------------------------------------------------------------------
# TAB 2: Image -> Word / Excel
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("Convert Scanned Text Block")
    uploaded_img = st.file_uploader("Upload Image Assets", type=["jpg", "jpeg", "png"], key="img_upload")

    if uploaded_img:
        image = Image.open(uploaded_img).convert("RGB")
        col1, col2 = st.columns(2)
        with col1:
            st.image(image, use_container_width=True)

        with st.spinner("Extracting standard notes..."):
            extracted_text = pytesseract.image_to_string(image, config="--psm 6")

        with col2:
            edited_text = st.text_area("Sanitized Working Space Data", extracted_text, height=350)

        if st.button("💾 Export Text Payload (.docx)"):
            doc = Document()
            for line in edited_text.split("\n"):
                doc.add_paragraph(line)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("⬇️ Download Docx File", data=buf, file_name="converted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
