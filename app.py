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
st.title("📊 Automated Vendor PFI Comparative Analysis (Precision Engine)")
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

def parse_pfi_images_with_ai(pfi_images, items_list, api_key_str):
    """Sends document images to Gemini 2.5 Flash to perform OCR and price extraction."""
    try:
        client = genai.Client(api_key=api_key_str)
        
        prompt_text = f"""You are a senior procurement analyst auditing a vendor Proforma Invoice (PFI).
        Look closely at the attached invoice image(s). Your job is to extract pricing data and match it to our internal Target Items list.
        
        Our Internal Target Items List:
        {json.dumps(items_list)}
        
        INSTRUCTIONS:
        1. Look for the vendor name at the top letterhead, logo, or stamp, and extract it.
        2. Go line-by-line through the items on the vendor PFI image. 
        3. Match the vendor's line items to our Target Items List by identifying core nouns, sizes, specifications, and materials (even if abbreviated, shortened, or in a different word order). 
           - For example, if our target is "10mm Hexagonal Stainless Bolt" and the vendor wrote "HEX BLT 10MM SS", that is a match.
        4. Extract the numeric UNIT PRICE for each matched item. Ignore currency symbols, commas, or total values.
        
        CRITICAL: Return your output strictly as a JSON object matching this structure. Do not include markdown code block syntax around the JSON:
        {{
            "vendor_name": "Name of Vendor",
            "prices": {{
                "EXACT TARGET ITEM DESCRIPTION FROM OUR LIST 1": 1500.50,
                "EXACT TARGET ITEM DESCRIPTION FROM OUR LIST 2": 420.00
            }}
        }}"""
        
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
        
        # Setting a strict json configuration to force compliant structured mapping
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
            ),
        )
        
        response_text = response.text.strip()
        return json.loads(response_text)
            
    except Exception as e:
        st.error(f"Error communicating with Gemini Vision: {e}")
        return {"vendor_name": "Unknown Vendor", "prices": {}}

if st.button("🚀 Process Invoices & Match Prices") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Executing OCR Vision Analysis..."):
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            items = []
            row_mapping = {}
            
            # Read items from Column C starting from row 9 down
            for row in range(9, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value
                if item_desc and str(item_desc).strip():
                    clean_desc = str(item_desc).strip()
                    if "ITEM DESCRIPTION" in clean_desc.upper():
                        continue
                    items.append(clean_desc)
                    row_mapping[clean_desc] = row
            
            if not items:
                st.error("No item descriptions discovered in Column C (checked from row 9 down).")
            else:
                st.info(f"Targeting {len(items)} matrix items for pricing updates.")
                
                # Excel column numbers for matching coordinates (E, H, K, N)
                vendor_start_cols = [5, 8, 11, 14]
                
                for index, pfi in enumerate(pfi_files[:4]):
                    st.write(f"📷 Performing OCR Character Identification on: **{pfi.name}**...")
                    
                    pfi_images = convert_pdf_to_images(pfi)
                    if not pfi_images:
                        st.warning(f"Skipping {pfi.name} due to PDF processing error.")
                        continue
                        
                    extracted_data = parse_pfi_images_with_ai(pfi_images, items, api_key)
                    
                    vendor_name = extracted_data.get("vendor_name", f"Vendor {index + 1}")
                    extracted_prices = extracted_data.get("prices", {})
                    
                    target_col = vendor_start_cols[index]
                    
                    # 1. Place Vendor Name in Row 8
                    ws.cell(row=8, column=target_col).value = vendor_name
                    
                    # 2. Match pricing rows
                    match_count = 0
                    for item, price in extracted_prices.items():
                        # Standardize checking to clean up minor whitespace mismatches
                        clean_item_key = str(item).strip()
                        if clean_item_key in row_mapping:
                            target_row = row_mapping[clean_item_key]
                            try:
                                ws.cell(row=target_row, column=target_col).value = float(price)
                                match_count += 1
                            except (ValueError, TypeError):
                                pass
                    
                    st.success(f"✅ Extracted **{vendor_name}** ({pfi.name}). Successfully identified and mapped {match_count}/{len(items)} prices.")
                
                output_filename = "Populated_Comparative_Analysis.xlsx"
                wb.save(output_filename)
                
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download Completed Excel Sheet",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                os.remove(output_filename)
