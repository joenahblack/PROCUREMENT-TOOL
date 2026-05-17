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
st.title("📊 Automated Vendor PFI Comparative Analysis (with OCR Vision)")
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
        # Reset file pointer just in case
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
        
        # 1. Base Prompt instructions
        prompt_text = f"""You are an expert procurement data extraction tool with advanced OCR vision capabilities.
        Analyze the attached image(s) of a Vendor's Proforma Invoice (PFI).
        
        Tasks:
        1. Extract the official VENDOR NAME (usually found at the top letterhead or logo area).
        2. Match the items on this invoice to our target item descriptions list.
        
        Target Items to look for:
        {json.dumps(items_list)}
        
        CRITICAL MATCHING INSTRUCTIONS:
        - Vendors will use heavily abbreviated names (e.g., 'BLT 10MM' instead of '10mm Hexagonal Stainless Bolt'). Use your procurement knowledge to determine if they refer to the same item.
        - Extract the UNIT PRICE for each matched item.
        - If an item from our target list is clearly missing from the vendor invoice, omit it from your output.
        
        Return your answer STRICTLY as a raw JSON dictionary with exactly two keys: "vendor_name" and "prices".
        The "prices" key must be a dictionary where keys are the EXACT item descriptions from our target list, and values are numbers (floats) representing the unit price.
        
        Example Output Format:
        {{
            "vendor_name": "ABC Industrial Supplies Ltd",
            "prices": {{
                "Target Item Description Example 1": 1500.50,
                "Target Item Description Example 2": 420.00
            }}
        }}"""
        
        # Start the contents list with our prompt text
        contents = [prompt_text]
        
        # 2. Safely add the images using Google's correct Types formatting
        for img in pfi_images:
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            img_bytes = img_byte_arr.getvalue()
            
            # Using the official Part.from_bytes wrapper to fix the validation crash
            image_part = types.Part.from_bytes(
                data=img_bytes,
                mime_type="image/jpeg"
            )
            contents.append(image_part)
        
        # Request content from Gemini 2.5 Flash using visual data objects
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents
        )
        
        response_text = response.text
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        
        if json_match:
            return json.loads(json_match.group(0))
        else:
            st.warning("Could not isolate JSON records from the vision model response.")
            return {"vendor_name": "Unknown Vendor", "prices": {}}
            
    except Exception as e:
        st.error(f"Error communicating with Gemini Vision: {e}")
        return {"vendor_name": "Unknown Vendor", "prices": {}}

if st.button("🚀 Process Invoices & Match Prices") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Executing OCR Vision Analysis... This may take a moment per page."):
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
                        if item in row_mapping:
                            target_row = row_mapping[item]
                            try:
                                ws.cell(row=target_row, column=target_col).value = float(price)
                                match_count += 1
                            except ValueError:
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
