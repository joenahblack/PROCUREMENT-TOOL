def extract_all_vendors_simultaneously(all_document_packages, row_item_dict, api_key_str):
    """Sends all vendor invoices together to Gemini with hyper-permissive matching rules."""
    try:
        client = genai.Client(api_key=api_key_str)
        
        prompt_text = f"""You are a senior procurement forensics manager. You are auditing multiple vendor quotations simultaneously, including documents from vendors like Moseng.
        
        Our Master Spreadsheet target lines (Format is "ROW_NUMBER": "EXACT MATERIAL DESCRIPTION"):
        {json.dumps(row_item_dict, indent=2)}
        
        YOUR INSTRUCTIONS:
        1. Look through each attached vendor quotation package individually.
        2. Extract the official Vendor Name from the letterhead/logo area. Strip out any stray quotes.
        3. Match the items on that invoice to our Master Spreadsheet target list.
        
        CRITICAL OVERRIDE FOR VENDORS WITH UNUSUAL LAYOUTS (e.g., Moseng):
        - Some vendors use internal item codes, list dimensions differently, or quote functional substitutes. 
        - Use your deep engineering and procurement domain knowledge to force a match if the physical material is clearly intended for that slot.
        - Ignore manufacturer part numbers, line order numbers, or minor brand differences. If it is a functional match, link it to our spreadsheet row!
        - Make absolutely sure you are pulling the UNIT PRICE, not the quantity, item index, or total row cost.
        4. If a vendor completely omitted an item from their quote, omit its row number from that vendor's dictionary block.
        
        OUTPUT FORMATTING RULE:
        Return your findings strictly as a clean JSON object matching this schema layout. Group your answers by the exact 'file_index' provided to you. Do not include markdown code block syntax formatting or backslashes.
        
        {{
            "quotations": [
                {{
                    "file_index": 0,
                    "vendor_name": "Official Vendor Name 1",
                    "row_prices": {{
                        "9": 1500.00,
                        "12": 435.50
                    }}
                }}
            ]
        }}"""
        
        contents = [prompt_text]
        
        for idx, package in enumerate(all_document_packages):
            contents.append(f"\n--- START OF VENDOR QUOTATION FILE INDEX {idx} (Filename: {package['name']}) ---")
            for img in package['images']:
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='JPEG')
                contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type="image/jpeg"))
            contents.append(f"--- END OF VENDOR QUOTATION FILE INDEX {idx} ---")
            
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        
        return json.loads(response.text.strip())
    except Exception as e:
        st.error(f"Failed to communicate with master alignment engine: {e}")
        return {"quotations": []}
