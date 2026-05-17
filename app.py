import streamlit as st
import openpyxl
from google import genai
from google.genai import types
from pdf2image import convert_from_bytes
import io
import json
import os
import re

st.set_page_config(page_title="Procurement Automation Tool", layout="wide")
st.title("📊 Automated Vendor PFI Comparative Analysis (Line-Audit Engine)")
st.subheader("Upload your template and scanned/digital vendor PFIs to auto-populate names and prices.")

# Sidebar for API Key configuration
with st.sidebar:
    st.header("Setup")
    api_key = st.text_input("Enter Gemini API Key:", type="password")
    st.markdown("[Get a free API Key here](https://aistudio.google.com/)")

# File Uploaders
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("1. Upload Excel Template (.xlsx)", type=["xlsx"])
with col2:
    pfi_files = st.file_uploader("2. Upload ALL Vendor PFIs together (PDFs)", type=["pdf"], accept_multiple_files=True)

def convert_pdf_to_images(pdf_file):
    """Converts a scanned PDF into a list of PIL Images using pdf2image."""
    try:
        pdf_bytes = pdf_file.read()
        pdf_file.seek(0)
        images = convert_from_bytes(pdf_bytes)
        return images
    except Exception as e:
        st.error(f"Failed to process PDF pages into images: {e}")
        return []

def parse_pfi_images_with_ai(pfi_images, row_item_dict, api_key_str):
    """Sends document images to Gemini to perform OCR and precise line auditing."""
    try:
        client = genai.Client(api_key=api_key_str)
        
        prompt_text = f"""You are a meticulous procurement forensic auditor. Your job is to extract data from the attached invoice image(s) with 100% horizontal accuracy. 
        Do not mismatch columns or lines.
        
        Our Master Spreadsheet Items (Format is "ROW_NUMBER": "ITEM DESCRIPTION"):
        {json.dumps(row_item_dict, indent=2)}
        
        Your Mission:
        1. Extract the official Vendor Name from the top header/logo.
        2. Go line-by-line through the vendor's invoice image. For EACH printed item, find the matching item in our spreadsheet list using your procurement knowledge (abbreviations like BLT, SS, MS, DIA match formal names).
        3. Double check the table columns on the image. Make sure you extract the UNIT PRICE, not the quantity, serial number, or line total.
        
        CRITICAL TWO-STEP OUTPUT FORMAT:
        You must structure your response exactly like this JSON layout. 
        Use the "audit_trail" section to write down your reasoning line-by-line first to lock in your accuracy. Then put the final mappings in "row_prices". Do not use markdown wrappers.
        
        {{
            "vendor_name": "Official Vendor Name",
            "audit_trail": [
                "Invoice Line 1 says 'HEX BLT 10MM' with unit price 15.00. This matches Spreadsheet Row 9 ('10mm Hexagonal Stainless Bolt')."
            ],
            "row_prices": {{
                "9": 15.00
            }}
        }}
        """
        
        contents = [prompt_text]
        for img in pfi_images:
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            img_bytes = img_byte_arr.getvalue()
            
            image_part = types.Part.from_bytes(
                data=img_bytes,
                mime_type="image/jpeg"
            )
            contents.append(image_part)
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
            ),
        )
        
        response_text = response.text.strip()
        
        # Display the audit trail logs live in Streamlit
        with st.expander("🔍 View AI Line-by-Line Match Explanations"):
            st.code(response_text, language="json")
            
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(0))
        else:
            return {"vendor_name": "Unknown Vendor", "row_prices": {}}
            
    except Exception as e:
        st.error(f"Error communicating with Gemini Vision: {e}")
        return {"vendor_name": "Unknown Vendor", "row_prices": {}}

if st.button("🚀 Process Invoices & Match Prices") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Running Audited Line-by-Line Vision Extraction..."):
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            row_item_dict = {}
            for row in range(9, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value
                if item_desc:
                    clean_desc = re.sub(r'\s+', ' ', str(item_desc)).strip()
                    if "ITEM DESCRIPTION" in clean_desc.upper() or not clean_desc:
                        continue
                    row_item_dict[str(row)] = clean_desc
            
            if not row_item_dict:
                st.error("No item descriptions found in Column C from row 9 down.")
            else:
                st.info(f"Loaded {len(row_item_dict)} target items from your comparative matrix.")
                
                vendor_start_cols = [5, 8, 11, 14] # E, H, K, N
                
                for index, pfi in enumerate(pfi_files[:4]):
                    st.write(f"📷 Analyzing Layout & Alignment for: **{pfi.name}**...")
                    
                    pfi_images = convert_pdf_to_images(pfi)
                    if not pfi_images:
                        st.warning(f"Skipping {pfi.name} due to image generation error.")
                        continue
                        
                    extracted_data = parse_pfi_images_with_ai(pfi_images, row_item_dict, api_key)
                    
                    vendor_name = extracted_data.get("vendor_name", f"Vendor {index + 1}")
                    row_prices = extracted_data.get("row_prices", {})
                    
                    target_col = vendor_start_cols[index]
                    ws.cell(row=8, column=target_col).value = vendor_name
                    
                    match_count = 0
                    for row_str, price in row_prices.items():
                        try:
                            target_row = int(row_str)
                            ws.cell(row=target_row, column=target_col).value = float(price)
                            match_count += 1
                        except (ValueError, TypeError):
                            pass
                    
                    st.success(f"✅ Finished **{vendor_name}** ({pfi.name}). Mapped {match_count} verified prices.")
                
                output_filename = "Populated_Comparative_Analysis.xlsx"
                wb.save(output_filename)
                
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download Corrected Excel Sheet",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                os.remove(output_filename)
