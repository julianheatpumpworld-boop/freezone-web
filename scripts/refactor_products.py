import json
import os
import re

def slugify(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9]+', '-', text)
    return text.strip('-')

def split_products():
    json_path = 'F:/Accio Work/products.json'
    output_dir = 'F:/Accio Work/content/products'
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    products = data.get('products', [])
    for p in products:
        # Use product number or slugified name as filename
        item_num = p.get('product_number', '')
        if not item_num:
            filename = slugify(p.get('name', 'product'))[:50]
        else:
            filename = slugify(item_num)
            
        file_path = os.path.join(output_dir, f"{filename}.json")
        
        # Avoid overwriting with same name if item_num is missing
        count = 1
        while os.path.exists(file_path):
            file_path = os.path.join(output_dir, f"{filename}-{count}.json")
            count += 1
            
        with open(file_path, 'w', encoding='utf-8') as pf:
            json.dump(p, pf, indent=4, ensure_ascii=False)
            
    print(f"Successfully split {len(products)} products into individual files.")

if __name__ == "__main__":
    split_products()
