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
st.caption("Universal Layout Mapping Engine — Zero Invisible Fonts, Zero Layout Stretches.")

tab1, tab2, tab3 = st.tabs([
    "📝 Image → Word / Excel", 
    "📄 Targeted Single Document Editor", 
    "📬 Universal Bulk Merge"
])

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
# TAB 2: Image Canvas Single PDF/Image Editor (No Invisible Fonts)
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("🔍 Canvas Overlay Single Editor")
    st.write("Upload a PDF or an Image. This mode converts pages into a canvas background so your text overlays perfectly above lines at any scale.")
    
    single_file = st.file_uploader("Upload PDF or Image Layout Template", type=["pdf", "png", "jpg", "jpeg"], key="single_canvas_upload")
    
    if single_file:
        # Load background layout smoothly whether it is a native image or a PDF page
        if single_file.name.lower().endswith('.pdf'):
            pdf_doc = fitz.open(stream=single_file.read(), filetype="pdf")
            page_idx = st.number_input("Page number to edit:", min_value=1, max_value=pdf_doc.page_count, value=1, key="single_pdf_pg")
            pix = pdf_doc[page_idx - 1].get_pixmap(matrix=fitz.Matrix(2.0, 2.0)) # high-res baseline conversion
            base_image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        else:
            base_image = Image.open(single_file).convert("RGB")
            
        W, H = base_image.size
        st.success(f"Layout Template Rendered! Canvas Area Dimensions: {W}x{H} Pixels.")
        
        # Setup session variables dynamically
        if "single_fields_count" not in st.session_state:
            st.session_state.single_fields_count = 2
            
        col_s_add, col_s_rem = st.columns(2)
        with col_s_add:
            if st.button("➕ Add Custom Text Layer Field", key="add_s_field"):
                st.session_state.single_fields_count += 1
                st.rerun()
        with col_s_rem:
            if st.button("🗑️ Remove Last Layer Field", key="rem_s_field") and st.session_state.single_fields_count > 1:
                st.session_state.single_fields_count -= 1
                st.rerun()
                
        st.write("### ⚙️ Position Configuration Settings & Value Inputs")
        layers_map = {}
        for idx in range(st.session_state.single_fields_count):
            st.markdown(f"**Custom Field Element #{idx+1}**")
            cx, cy, cs, ca, ct = st.columns([2, 2, 1.5, 2, 4.5])
            with cx:
                x_pos = st.number_input(f"X (Horizontal)", min_value=0, max_value=W, value=int(W/2), key=f"s_x_{idx}")
            with cy:
                y_pos = st.number_input(f"Y (Vertical)", min_value=0, max_value=H, value=int(H/2) + (idx * 80) - 100, key=f"s_y_{idx}")
            with cs:
                f_size = st.number_input(f"Font Size", min_value=10, max_value=200, value=36, key=f"s_s_{idx}")
            with ca:
                align_type = st.selectbox(f"Align Type", options=["Center", "Left"], key=f"s_a_{idx}")
            with ct:
                text_val = st.text_input(f"Text String Value:", value=f"Sample Field Text #{idx+1}", key=f"s_t_{idx}")
                
            layers_map[idx] = {"x": x_pos, "y": y_pos, "size": f_size, "align": align_type, "text": text_val.strip()}
            st.markdown("<br>", unsafe_with_html=True)
            
        if st.button("🚀 Render Canvas & Download Document", use_container_width=True, key="process_single_canvas"):
            img_copy = base_image.copy()
            draw = ImageDraw.Draw(img_copy)
            
            for idx, data in layers_map.items():
                if not data["text"]:
                    continue
                try:
                    font = ImageFont.truetype("LiberationSans-Regular.ttf", data["size"])
                except:
                    try:
                        font = ImageFont.truetype("DejaVuSans.ttf", data["size"])
                    except:
                        font = ImageFont.load_default()
                        
                left, top, right, bottom = draw.textbbox((0, 0), data["text"], font=font)
                text_width = right - left
                
                final_x = data["x"] - (text_width / 2) if data["align"] == "Center" else data["x"]
                draw.text((final_x, data["y"]), data["text"], fill=(0, 0, 0), font=font)
                
            pdf_out = io.BytesIO()
            img_copy.save(pdf_out, format="PDF")
            st.success("Document text compiled with 100% visibility scaling!")
            st.download_button("⬇️ Download Final Document PDF", data=pdf_out.getvalue(), file_name="canvas_output.pdf", mime="application/pdf", use_container_width=True)

# ---------------------------------------------------------------------------
# TAB 3: Universal Bulk Merge
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("Bulk Document / Certificate / Offer Letter Generation")
    st.write("Upload any blank layout background along with your custom Excel data to process mass files smoothly.")
    
    st.write("### 1. Download Demo Excel Data Structure")
    is_offer = st.checkbox("Check here if generating Offer Letters (Changes data structure layout)", key="bulk_offer_toggle")
    
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
    st.download_button("⬇️ Download Sample Excel Template (.xlsx)", data=demo_data, file_name="handywriter_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_demo_excel")

    st.markdown("---")
    st.write("### 2. Upload Document Layout Template and Data Sheets")
    
    template_file = st.file_uploader("Upload Blank Template Image (PNG or JPG only)", type=["png", "jpg", "jpeg"], key="bulk_template_upload")
    excel_file = st.file_uploader("Upload Student Excel Sheet (.xlsx)", type=["xlsx"], key="main_bulk_excel")

    if template_file and excel_file:
        base_image = Image.open(template_file).convert("RGB")
        W, H = base_image.size
        st.success(f"Template Loaded Successfully! Size: {W}x{H} Pixels.")
        
        wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
        ws = wb.active
        headers = [c.value for c in ws[1] if c.value is not None]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        
        st.info(f"Detected columns from Excel file: {', '.join(headers)}")
        st.write(f"Total rows discovered: **{len(rows)}** records.")
        
        st.write("### 3. Coordinate Position Configuration Settings (X, Y Axes)")
        
        positions = {}
        for index, h in enumerate(headers):
            st.markdown(f"**⚙️ Position Strategy Configurations for column: `{{{{{h}}}}}`**")
            cx, cy, cs, ca = st.columns([2, 2, 2, 2])
            with cx:
                x_pos = st.number_input(f"X (Horizontal Axis) - {h}", min_value=0, max_value=W, value=int(W/2), key=f"bx_{h}")
            with cy:
                y_pos = st.number_input(f"Y (Vertical Axis) - {h}", min_value=0, max_value=H, value=int(H/2) + (index * 70) - 100, key=f"by_{h}")
            with cs:
                f_size = st.number_input(f"Font Size Limit - {h}", min_value=10, max_value=200, value=32, key=f"bs_{h}")
            with ca:
                align_type = st.selectbox(f"Data Alignment Alignment Type - {h}", options=["Center", "Left"], key=f"ba_{h}")
            
            positions[h] = {"x": x_pos, "y": y_pos, "size": f_size, "align": align_type}
            st.markdown("<br>", unsafe_with_html=True)
            
        name_col = st.selectbox("Select column header to use for individual document filenames:", options=headers, key="filename_select")

        if st.button(f"🚀 Mass Generate All {len(rows)} Custom PDF Files", key="run_bulk_engine"):
            zip_buf = io.BytesIO()
            progress = st.progress(0, text="Processing files dynamically...")

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
                    
                    fname = str(rowdict.get(name_col, f"file_{idx+1}")).replace(" ", "_").replace("/", "-")
                    zf.writestr(f"{fname}.pdf", pdf_out.getvalue())
                    
                    progress.progress((idx + 1) / len(rows), text=f"Processed metadata item {idx+1}/{len(rows)}")
                    
            st.success("All dynamic file operations executed cleanly!")
            st.download_button("⬇️ Download All Extracted Files as ZIP", data=zip_buf.getvalue(), file_name="bulk_generated_documents.zip", mime="application/zip", use_container_width=True, key="dl_bulk_zip")

st.divider()
st.caption("HandyWriter Ultimate Pro · Secure Sandbox Local Instance Execution.")
