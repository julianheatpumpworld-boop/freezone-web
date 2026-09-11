import json
import os

def build_bundled_json():
    input_dir = 'F:/Accio Work/content/products'
    output_path = 'F:/Accio Work/products.json'
    
    products = []
    if not os.path.exists(input_dir):
        print("Error: Input directory does not exist.")
        return

    files = [f for f in os.listdir(input_dir) if f.endswith('.json')]
    for filename in files:
        file_path = os.path.join(input_dir, filename)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                product_data = json.load(f)
                products.append(product_data)
        except Exception as e:
            print(f"Skipping {filename} due to error: {e}")
            
    # Keep consistent sorting if possible, e.g., by name
    products.sort(key=lambda x: x.get('name', '').lower())
    
    output = {"products": products}
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=4, ensure_ascii=False)
        
    print(f"Successfully bundled {len(products)} products into products.json")

if __name__ == "__main__":
    build_bundled_json()
