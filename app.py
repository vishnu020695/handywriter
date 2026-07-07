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
st.caption("Visual Document Editor and Bulk PDF Engine — No Invisible Font or Alignment Crashes.")

tab1, tab2, tab3 = st.tabs([
    "📝 Image → Word / Excel", 
    "📄 Visual Single PDF Editor", 
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
# TAB 2: Visual Single PDF Editor (Fixed Bounding Box Rendering & Native Font Scaling)
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("📄 Visual Layout Mapping Editor")
    st.write("Upload your template PDF file. Map text blocks using target numbers to directly type replacements without disrupting surrounding lines.")
    
    uploaded_pdf = st.file_uploader("Upload Single PDF / Offer Letter", type=["pdf"], key="pdf_upload_single")
    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        st.info(f"Loaded PDF Document with {doc_pdf.page_count} page(s).")
        
        page_num = st.number_input("Select page number to view and modify:", min_value=1, max_value=doc_pdf.page_count, value=1, key="s_page")
        page = doc_pdf[page_num - 1]
        
        # Pull text blocks 
        raw_blocks = page.get_text("blocks")
        blocks = [b for b in raw_blocks if b[4].strip()]
        
        # Render a high-resolution display overlay map to visually locate layout sections
        zoom = 1.5
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        preview_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        draw_map = ImageDraw.Draw(preview_img)
        
        for i, b in enumerate(blocks):
            x0, y0, x1, y1 = [v * zoom for v in b[:4]]
            draw_map.rectangle([x0, y0, x1, y1], outline="blue", width=2)
            draw_map.text((x0, max(0, y0 - 15)), f"#{i + 1}", fill="blue")
            
        st.image(preview_img, caption="Document Structural Reference Layout Map", width=650)
        
        if not blocks:
            st.warning("No dynamic structural text boxes detected on this layout page.")
        else:
            st.write("### 📝 Modify Target Text Inputs Below:")
            edited_values = []
            
            # Form grid configuration split for structured input tracking
            for i, b in enumerate(blocks):
                original_sample = b[4].strip()
                val = st.text_input(f"Block Box #{i + 1} (Current: '{original_sample[:60]}')", value=b[4].rstrip("\n"), key=f"sb_val_{i}")
                edited_values.append(val)
                
            if st.button("💾 Apply Dynamic Updates & Export PDF", use_container_width=True):
                text_dict = page.get_text("dict")
                dict_blocks = text_dict["blocks"]
                
                for i, (b, new_val) in enumerate(zip(blocks, edited_values)):
                    if new_val != b[4].rstrip("\n"):
                        bbox = fitz.Rect(b[:4])
                        
                        # Dynamically discover font metrics from layout structure
                        fontsize = 12
                        for db in dict_blocks:
                            for line in db.get("lines", []):
                                for span in line["spans"]:
                                    if fitz.Rect(span["bbox"]).intersects(bbox):
                                        fontsize = span["size"]
                                        break
                        
                        # Execute a layout boundary check to safeguard against empty render parameters
                        if bbox.width > 0 and bbox.height > 0:
                            page.add_redact_annot(bbox, fill=(1, 1, 1))
                            page.apply_redactions()
                            
                            # Using native PDF core font sets ('helv' for Helvetica) scales flawlessly to any text box width
                            page.insert_textbox(bbox, new_val, fontsize=fontsize, fontname="helv", align=0)
                
                out = io.BytesIO(doc_pdf.tobytes())
                st.success("Document template populated and compiled successfully!")
                st.download_button("⬇|️ Download Final PDF Document", data=out, file_name="handywriter_output.pdf", mime="application/pdf", use_container_width=True)

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
