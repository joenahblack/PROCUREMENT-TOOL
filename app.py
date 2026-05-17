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
st.title("📊 Automated Vendor PFI Comparative Analysis")
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
        
        prompt_text = f"""You are a master procurement auditor. Analyze the attached invoice/proforma invoice images.
        
        Step 1: Extract the official vendor name from the header/logo.
        Step 2: Read every line item printed on this document. Match them against our Target Items list.
        
        Target Items List:
        {json.dumps(items_list)}
        
        Matching Rules:
        - Be smart and flexible. Vendors use abbreviations (e.g. 'BLT', 'SS', 'MS', 'DIA', 'W/', 'THK').
        - Match based on specifications, sizes, dimensions, and core item nouns. If they clearly refer to the same physical material, it is a match.
        - Extract the numeric unit price.
        
        CRITICAL OUTPUT FORMATTING:
        You must output your final answer wrapped inside a clean JSON structure block. Ensure the JSON is completely filled with your findings:
        {{
            "vendor_name": "Name of Vendor",
            "prices": {{
                "EXACT DESCRIPTIONS FROM TARGET LIST": 1500.00
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
        
        # We drop the strict system json mode configuration to allow full reasoning/OCR capability
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents
        )
        
        response_text = response.text.strip()
        
        # Display response diagnostic console
        with st.expander("🔍 Visual Inspection Log"):
            st.code(response_text)
            
        # Isolate the JSON string from raw markdown text safely
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(0))
        else:
            return {"vendor_name": "Unknown Vendor", "prices": {}}
            
    except Exception as e:
        st.error(f"Error communicating with Gemini Vision: {e}")
        return {"vendor_name": "Unknown Vendor", "prices": {}}

if st.button("🚀 Process Invoices & Match Prices") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Executing OCR Vision Analysis... Please keep this window open."):
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            items = []
            row_mapping = {}
            
            # Read items from Column C starting from row 9 down
            for row in range(9, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value
                if item_desc:
                    clean_desc = re.sub(r'\s+', ' ', str(item_desc)).strip()
                    if "ITEM DESCRIPTION" in clean_desc.upper() or not clean_desc:
                        continue
                    items.append(clean_desc)
                    row_mapping[clean_desc] = row
            
            if not items:
                st.error("No item descriptions discovered in Column C (checked from row 9 down).")
            else:
                st.info(f"Targeting {len(items)} matrix items for pricing updates.")
                
                vendor_start_cols = [5, 8, 11, 14]
                
                for index, pfi in enumerate(pfi_files[:4]):
                    st.write(f"📷 Processing Document: **{pfi.name}**...")
                    
                    pfi_images = convert_pdf_to_images(pfi)
                    if not pfi_images:
                        st.warning(f"Skipping {pfi.name} due to PDF processing error.")
                        continue
                        
                    extracted_data = parse_pfi_images_with_ai(pfi_images, items, api_key)
                    
                    vendor_name = extracted_data.get("vendor_name", f"Vendor {index + 1}")
                    extracted_prices = extracted_data.get("prices", {})
                    
                    target_col = vendor_start_cols[index]
                    ws.cell(row=8, column=target_col).value = vendor_name
                    
                    match_count = 0
                    for item, price in extracted_prices.items():
                        clean_item_key = re.sub(r'\s+', ' ', str(item)).strip()
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
