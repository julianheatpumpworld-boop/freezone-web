import pandas as pd
import json
import os
import re

def slugify(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9]+', '-', text)
    return text.strip('-')

def build_bundled_json():
    input_dir = 'F:/Accio Work/content/products'
    output_path = 'F:/Accio Work/products.json'
    
    products = []
    if not os.path.exists(input_dir):
        return

    files = [f for f in os.listdir(input_dir) if f.endswith('.json')]
    for filename in files:
        file_path = os.path.join(input_dir, filename)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                product_data = json.load(f)
                products.append(product_data)
        except:
            continue
            
    products.sort(key=lambda x: x.get('name', '').lower())
    output = {"products": products}
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=4, ensure_ascii=False)
    print(f"Successfully bundled {len(products)} products into products.json")

def convert_csv_to_json(csv_path):
    output_dir = 'F:/Accio Work/content/products'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    df = pd.read_csv(csv_path, encoding='utf-8')
    df.columns = [c.strip() for c in df.columns]
    
    count = 0
    for i, row in df.iterrows():
        def clean(val):
            if pd.isna(val): return ""
            return str(val).strip().replace('\t', '')

        img_raw = clean(row.get('Prod_Image', ''))
        images = [url.strip() for url in img_raw.split(',') if url.strip()]
        primary_image = images[0] if images else ""
        
        pricing = []
        for j in range(1, 11):
            q_val = row.get(f'Q{j}')
            p_val = row.get(f'P{j}')
            d_val = row.get(f'D{j}')
            if pd.notna(q_val) and pd.notna(p_val):
                try:
                    qty = int(float(str(q_val).replace('\t', '').strip()))
                    price = float(str(p_val).replace('\t', '').strip())
                    pricing.append({"qty": qty, "price": price, "code": clean(d_val)})
                except: continue

        product = {
            "name": clean(row.get('Product_Name', "N/A")),
            "product_number": clean(row.get('Product_Number', "")),
            "price": float(pricing[0]['price']) if pricing else 0.0,
            "category": clean(row.get('Category', "Uncategorized")).split(',')[0].strip(),
            "description": clean(row.get('Description', "")),
            "summary": clean(row.get('Summary', "")),
            "image": primary_image,
            "images": images,
            "featured": i < 12,
            "details": {
                "colors": clean(row.get('Product_Color', "")),
                "materials": clean(row.get('Material', "")),
                "size": clean(row.get('Size_Values', "")),
                "imprint_method": clean(row.get('Imprint_Method', "")),
                "imprint_color": clean(row.get('Imprint_Color', "")),
                "production_time": clean(row.get('Production_Time', "")),
                "price_includes": clean(row.get('Price_Includes', ""))
            },
            "pricing_grid": pricing
        }
        
        # Save individual file
        item_num = product['product_number']
        filename = slugify(item_num) if item_num else slugify(product['name'])[:50]
        file_path = os.path.join(output_dir, f"{filename}.json")
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(product, f, indent=4, ensure_ascii=False)
        count += 1
    
    print(f"Successfully converted {count} products to individual files.")
    build_bundled_json()

if __name__ == "__main__":
    convert_csv_to_json('F:/Accio Work/5855589_USD.csv')
