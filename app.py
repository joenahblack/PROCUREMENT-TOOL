import streamlit as st
import openpyxl
from google import genai
from google.genai import types
from pdf2image import convert_from_bytes
import io
import json
import os
import re

st.set_page_config(page_title="Procurement Tool - Clean Reset", layout="wide")
st.title("📊 Procurement Automation (Reset & Diagnose)")

# Sidebar for Setup
with st.sidebar:
    st.header("Authentication")
    api_key = st.text_input("Enter Gemini API Key:", type="password")
    st.markdown("[Get a fresh API Key here](https://aistudio.google.com/)")

# Uploads
template_file = st.file_uploader("1. Upload Excel Template (.xlsx)", type=["xlsx"])
pfi_file = st.file_uploader("2. Upload JUST ONE Vendor PFI (PDF)", type=["pdf"])

def convert_pdf_to_images(pdf_file):
    try:
        pdf_bytes = pdf_file.read()
        pdf_file.seek(0)
        return convert_from_bytes(pdf_bytes)
    except Exception as e:
        st.error(f"Failed to turn PDF into images: {e}")
        return []

if st.button("🚀 Run Diagnostic Extraction") and template_file and pfi_file:
    if not api_key:
        st.warning("Please enter your API Key.")
    else:
        # Load your spreadsheet row coordinates
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active
        
        row_item_dict = {}
        for row in range(9, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=3).value # Column C
            if cell_val:
                clean_text = str(cell_val).strip()
                if "ITEM DESCRIPTION" in clean_text.upper():
                    continue
                row_item_dict[str(row)] = clean_text

        st.write(f"### Loaded {len(row_item_dict)} target descriptions from Excel Column C.")
        
        # Slicing PDF pages
        images = convert_pdf_to_images(pfi_file)
        st.write(f"📷 Successfully split invoice into **{len(images)} page image(s)**.")
        
        # Connect to Gemini
        try:
            client = genai.Client(api_key=api_key)
            
            prompt = f"""You are a data entry assistant. Look closely at the image of the invoice provided.
            
            Your tasks:
            1. Identify the official VENDOR NAME.
            2. Match the items on this invoice page to our spreadsheet target list below.
            3. Find the matching item's numeric UNIT PRICE (do not grab quantities, line numbers, or line totals).
            
            Our Spreadsheet Target List (Format is "ROW_NUMBER": "DESCRIPTION"):
            {json.dumps(row_item_dict, indent=2)}
            
            Return your answer STRICTLY as a raw JSON dictionary matching this format exactly:
            {{
                "vendor_name": "Name of Vendor",
                "row_prices": {{
                    "9": 1500.00
                }}
            }}"""
            
            # Package inputs
            contents = [prompt]
            for img in images:
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='JPEG')
                contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type="image/jpeg"))
            
            with st.spinner("AI is examining the invoice layout..."):
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=contents,
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
            
            # Print EXACTLY what came out of the AI engine
            st.success("🤖 **Raw Data Received from Gemini:**")
            raw_result = json.loads(response.text.strip())
            st.json(raw_result)
            
            # Test writing to Excel live
            st.write("### ⚙️ Writing to Excel Data stream:")
            vendor_name = raw_result.get("vendor_name", "Unknown Vendor")
            row_prices = raw_result.get("row_prices", {})
            
            # Write Name to E8 (Vendor 1 Column Anchor)
            ws.cell(row=8, column=5).value = vendor_name
            st.write(f"✅ Assigned Vendor Name `'{vendor_name}'` to cell `E8`")
            
            mapped_count = 0
            for r_str, price in row_prices.items():
                try:
                    r_num = int(r_str)
                    p_num = float(str(price).replace(",", "").strip())
                    ws.cell(row=r_num, column=5).value = p_num # Put prices in Column E
                    st.write(f"➡️ Placed price `{p_num}` onto Row `{r_num}`, Column `E` (Matches: *{row_item_dict.get(r_str)}*)")
                    mapped_count += 1
                except Exception as cell_err:
                    st.error(f"❌ Failed processing row {r_str} with value {price}: {cell_err}")
            
            # Generate Download Link if mapping works
            output_name = "Diagnostic_Output.xlsx"
            wb.save(output_name)
            st.download_button("📥 Download Populated Excel Sheet", data=open(output_name, "rb"), file_name=output_name)
            os.remove(output_name)
            
        except Exception as api_err:
            st.error(f"API pipeline broke down: {api_err}")
