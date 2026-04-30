import pandas as pd
import json
import os

def convert_excel_to_json(excel_path, json_path):
    df = pd.read_excel(excel_path)
    
    products = []
    for i, row in df.iterrows():
        # Multi-image support
        img_raw = str(row.get('Prod_Image', ''))
        images = [url.strip() for url in img_raw.split(',') if url.strip()]
        primary_image = images[0] if images else ""
        
        # Pricing Grid
        pricing = []
        for j in range(1, 11):
            q_val = row.get(f'Q{j}')
            p_val = row.get(f'P{j}')
            d_val = row.get(f'D{j}')
            if pd.notna(q_val) and pd.notna(p_val):
                pricing.append({
                    "qty": int(q_val),
                    "price": float(p_val),
                    "code": str(d_val) if pd.notna(d_val) else ""
                })

        product = {
            "name": str(row.get('Product_Name', "N/A")).strip(),
            "product_number": str(row.get('Product_Number', "")).strip(),
            "price": float(row.get('P1', 0)) if pd.notna(row.get('P1')) else 0.0,
            "category": str(row.get('Category', "Uncategorized")).split(',')[0].strip(),
            "description": str(row.get('Description', "")).strip(),
            "summary": str(row.get('Summary', "")).strip(),
            "image": primary_image,
            "images": images, # Array of all images
            "featured": i < 8,
            "details": {
                "colors": str(row.get('Product_Color', "")).strip(),
                "materials": str(row.get('Material', "")).strip(),
                "size": str(row.get('Size_Values', "")).strip(),
                "imprint_method": str(row.get('Imprint_Method', "")).strip(),
                "imprint_color": str(row.get('Imprint_Color', "")).strip(),
                "production_time": str(row.get('Production_Time', "")).strip(),
                "price_includes": str(row.get('Price_Includes', "")).strip()
            },
            "pricing_grid": pricing
        }
        products.append(product)
    
    output = {"products": products}
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=4, ensure_ascii=False)
    
    print(f"Successfully converted {len(products)} products with full ASI details.")

if __name__ == "__main__":
    convert_excel_to_json('F:/Accio Work/5855589_USD.xlsx', 'F:/Accio Work/products.json')
