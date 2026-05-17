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
st.title("📊 Automated Vendor PFI Comparative Analysis (Row-Mapping Engine)")
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
    """Sends document images to Gemini to perform OCR and direct row mapping."""
    try:
        client = genai.Client(api_key=api_key_str)
        
        prompt_text = f"""You are a master procurement auditor mapping vendor invoice prices directly to our spreadsheet row coordinates.
        Look closely at the attached invoice image(s).
        
        Your Core Instructions:
        1. Extract the official Vendor Name from the letterhead/logo at the top.
        2. Match the items printed on this invoice image to our Master Spreadsheet Items list below.
        3. For every item you match, extract its numeric UNIT PRICE and pair it directly with its corresponding SPREADSHEET ROW NUMBER.
        
        Master Spreadsheet Items (Format is "ROW_NUMBER": "ITEM DESCRIPTION"):
        {json.dumps(row_item_dict, indent=2)}
        
        Matching Guardrails:
        - Vendors will use heavy abbreviations, different word orders, or shortened specs (e.g., 'BLT' vs 'Bolt', 'SS' vs 'Stainless', '10mm' vs 'M10'). Match them flexibly using your engineering and procurement knowledge.
        - Only return matches where you are confident the physical item is the same.
        
        OUTPUT FORMAT:
        You must output your findings strictly as a clean JSON object structure. Do not wrap it in markdown code blocks.
        {{
            "vendor_name": "Official Vendor Name",
            "row_prices": {{
                "ROW_NUMBER_AS_STRING": 1250.00
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
        
        # Display response diagnostic logs
        with st.expander("🔍 Live AI Extraction Inspection Log"):
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
        with st.spinner("Executing Coordinate OCR Vision Analysis..."):
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            # Map Row Numbers directly to Descriptions to eliminate Python matching errors
            row_item_dict = {}
            
            for row in range(9, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value # Column C
                if item_desc:
                    clean_desc = re.sub(r'\s+', ' ', str(item_desc)).strip()
                    if "ITEM DESCRIPTION" in clean_desc.upper() or not clean_desc:
                        continue
                    # Store row as string key for seamless JSON processing
                    row_item_dict[str(row)] = clean_desc
            
            if not row_item_dict:
                st.error("No item descriptions found in Column C from row 9 down.")
            else:
                st.info(f"Loaded {len(row_item_dict)} items from your spreadsheet matrix.")
                
                # Target columns for Vendor 1, 2, 3, 4 (E, H, K, N)
                vendor_start_cols = [5, 8, 11, 14]
                
                for index, pfi in enumerate(pfi_files[:4]):
                    st.write(f"📷 Processing Document: **{pfi.name}**...")
                    
                    pfi_images = convert_pdf_to_images(pfi)
                    if not pfi_images:
                        st.warning(f"Skipping {pfi.name} due to PDF transformation error.")
                        continue
                        
                    # Request data from Gemini using direct row mapping payload
                    extracted_data = parse_pfi_images_with_ai(pfi_images, row_item_dict, api_key)
                    
                    vendor_name = extracted_data.get("vendor_name", f"Vendor {index + 1}")
                    row_prices = extracted_data.get("row_prices", {})
                    
                    target_col = vendor_start_cols[index]
                    
                    # Write vendor name to Row 8
                    ws.cell(row=8, column=target_col).value = vendor_name
                    
                    match_count = 0
                    # Write values directly into the row numbers specified by the AI
                    for row_str, price in row_prices.items():
                        try:
                            target_row = int(row_str)
                            ws.cell(row=target_row, column=target_col).value = float(price)
                            match_count += 1
                        except (ValueError, TypeError):
                            pass
                    
                    st.success(f"✅ Extracted **{vendor_name}** ({pfi.name}). Successfully mapped {match_count} items directly to rows.")
                
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
