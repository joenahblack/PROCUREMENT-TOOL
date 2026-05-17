import streamlit as st
import openpyxl
from google import genai
from google.genai import types
from pdf2image import convert_from_bytes
import io
import json
import os
import re

st.set_page_config(page_title="Multi-Quotation Comparative Analysis Engine", layout="wide")
st.title("📊 Multi-Quotation Vendor Matching Engine")
st.subheader("Upload your template and multiple quotation PDFs to automatically map vendor names and material prices.")

# Sidebar setup
with st.sidebar:
    st.header("Authentication")
    api_key = st.text_input("Enter Gemini API Key:", type="password")
    st.markdown("[Get an API Key here](https://aistudio.google.com/)")

# Inputs
template_file = st.file_uploader("1. Upload Master Excel Template (.xlsx)", type=["xlsx"])
pfi_files = st.file_uploader("2. Upload ALL Vendor Quotations / PFIs (Select multiple PDFs)", type=["pdf"], accept_multiple_files=True)

def convert_pdf_to_images(pdf_file):
    """Converts multi-page PDFs cleanly into standard images for visual layout auditing."""
    try:
        pdf_bytes = pdf_file.read()
        pdf_file.seek(0)
        return convert_from_bytes(pdf_bytes)
    except Exception as e:
        st.error(f"Failed processing image conversions for {pdf_file.name}: {e}")
        return []

def extract_all_quotations(all_document_packages, row_item_dict, api_key_str):
    """
    Sends all quotations simultaneously to Gemini to map vendor names and 
    line items directly to spreadsheet rows without shifting errors.
    """
    try:
        client = genai.Client(api_key=api_key_str)
        
        prompt_text = f"""You are an elite procurement data pipeline. You are reviewing multiple vendor quotations simultaneously.
        
        Our Master Spreadsheet target lines (Format is "ROW_NUMBER": "EXACT MATERIAL DESCRIPTION"):
        {json.dumps(row_item_dict, indent=2)}
        
        INSTRUCTIONS:
        1. Examine each attached vendor quotation package individually.
        2. Extract the official Vendor Name from the top header/logo of that package.
        3. Match the line items printed on that specific document to our Master Spreadsheet target lines. 
           - Look past abbreviations (e.g., 'BLT' -> 'Bolt', 'SS' -> 'Stainless Steel', 'DIA' -> 'Diameter'). Match by true materials and sizing specifications.
        4. Extract ONLY the numeric UNIT PRICE for that row. Discard quantities, serial numbers, or line totals.
        
        CRITICAL OUTPUT FORMATTING:
        Return your results exactly matching the structured format below. Group the data by each individual quotation file index provided to you. Do not use markdown wrappers.
        
        {{
            "quotations": [
                {{
                    "file_index": 0,
                    "vendor_name": "Official Name of Vendor 1",
                    "row_prices": {{
                        "9": 1450.00,
                        "12": 320.50
                    }}
                }}
            ]
        }}"""
        
        contents = [prompt_text]
        
        # Build image context parameters tagged with file indexes
        for idx, package in enumerate(all_document_packages):
            contents.append(f"\n--- START OF QUOTATION FILE INDEX {idx} (Named: {package['name']}) ---")
            for img in package['images']:
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='JPEG')
                contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type="image/jpeg"))
            contents.append(f"--- END OF QUOTATION FILE INDEX {idx} ---")
            
        # Call Gemini with Native JSON enforcement
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        
        return json.loads(response.text.strip())
    except Exception as e:
        st.error(f"Error processing visual layout schema matching: {e}")
        return {"quotations": []}

if st.button("🚀 Process & Map All Quotations Simultaneously") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar configuration.")
    else:
        with st.spinner("Analyzing multiple quotation layouts and mapping matrices..."):
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            # 1. Isolate spreadsheet material definitions from Column C
            row_item_dict = {}
            for row in range(9, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=3).value
                if cell_val:
                    clean_desc = re.sub(r'\s+', ' ', str(cell_val)).strip()
                    if "ITEM DESCRIPTION" in clean_desc.upper() or not clean_desc:
                        continue
                    row_item_dict[str(row)] = clean_desc
            
            if not row_item_dict:
                st.error("No valid material descriptions found in Column C (checking rows 9 down).")
            else:
                st.info(f"Loaded {len(row_item_dict)} clear material items to match from spreadsheet.")
                
                # Convert up to 4 quotation PDFs concurrently
                all_packages = []
                for pfi in pfi_files[:4]:
                    st.write(f"📷 Scanning text coordinates for: **{pfi.name}**...")
                    imgs = convert_pdf_to_images(pfi)
                    if imgs:
                        all_packages.append({"name": pfi.name, "images": imgs})
                
                # Execute unified master execution run
                extracted_results = extract_all_quotations(all_packages, row_item_dict, api_key)
                
                # View full live payload trace tracking
                with st.expander("🔍 Live AI Extraction Mapping Log"):
                    st.json(extracted_results)
                
                # 2. Map directly to column tracks (E=Vendor 1, H=Vendor 2, K=Vendor 3, N=Vendor 4)
                vendor_columns = [5, 8, 11, 14]
                quotation_data_list = extracted_results.get("quotations", [])
                
                for q_data in quotation_data_list:
                    f_idx = q_data.get("file_index")
                    if f_idx is None or f_idx >= len(vendor_columns):
                        continue
                        
                    target_col = vendor_columns[f_idx]
                    v_name = q_data.get("vendor_name", f"Vendor Slot {f_idx + 1}")
                    row_prices = q_data.get("row_prices", {})
                    
                    # Insert Vendor Name at row 8 of respective column track
                    ws.cell(row=8, column=target_col).value = v_name
                    
                    # Direct coordinate matrix pricing injection
                    successful_inserts = 0
                    for r_str, price in row_prices.items():
                        try:
                            target_row = int(r_str)
                            clean_price = float(str(price).replace(",", "").strip())
                            ws.cell(row=target_row, column=target_col).value = clean_price
                            successful_inserts += 1
                        except (ValueError, TypeError):
                            pass
                            
                    st.success(f"✅ Columns Successfully Formatted for **{v_name}**: Injected {successful_inserts} unit prices.")
                
                # File Generation Assembly
                output_filename = "Multi_Vendor_Populated_Analysis.xlsx"
                wb.save(output_filename)
                
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download Ready Comparative Matrix Sheet",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                os.remove(output_filename)
