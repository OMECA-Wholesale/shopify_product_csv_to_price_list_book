import pandas as pd
from typing import Dict, List, Optional, Any
import os

class ProductExtractor:
    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.products = {}
        self.raw_df = None

    def load_data(self) -> bool:
        try:
            self.raw_df = pd.read_csv(self.csv_path)
            return True
        except Exception as e:
            print(f"Error loading CSV: {e}")
            return False

    def extract_products(self) -> Dict[str, Dict]:
        if self.raw_df is None:
            self.load_data()

        products = {}

        for _, row in self.raw_df.iterrows():
            handle = row.get('Handle')
            if pd.isna(handle):
                continue

            if handle not in products:
                products[handle] = {
                    'handle': handle,
                    'title': row.get('Title', ''),
                    'body_html': row.get('Body (HTML)', ''),
                    'vendor': row.get('Vendor', ''),
                    'product_category': row.get('Product Category', ''),
                    'type': row.get('Type', ''),
                    'tags': row.get('Tags', ''),
                    'published': row.get('Published', True),
                    'variants': [],
                    'images': []
                }

                if pd.notna(row.get('Image Src')):
                    products[handle]['images'].append({
                        'src': row.get('Image Src'),
                        'position': row.get('Image Position', 1),
                        'alt_text': row.get('Image Alt Text', '')
                    })

            if pd.notna(row.get('Option1 Value')) or pd.notna(row.get('Variant SKU')):
                variant = {
                    'sku': row.get('Variant SKU', ''),
                    'price': row.get('Variant Price', 0),
                    'compare_at_price': row.get('Variant Compare At Price', ''),
                    'inventory_qty': row.get('Variant Inventory Qty', 0),
                    'weight': row.get('Variant Grams', 0),
                    'weight_unit': row.get('Variant Weight Unit', 'g'),
                    'barcode': row.get('Variant Barcode', ''),
                    'option1': row.get('Option1 Value', ''),
                    'option2': row.get('Option2 Value', ''),
                    'option3': row.get('Option3 Value', ''),
                    'option1_name': row.get('Option1 Name', ''),
                    'option2_name': row.get('Option2 Name', ''),
                    'option3_name': row.get('Option3 Name', ''),
                    'taxable': row.get('Variant Taxable', True),
                    'requires_shipping': row.get('Variant Requires Shipping', True)
                }
                products[handle]['variants'].append(variant)

            if pd.notna(row.get('Image Src')) and row.get('Image Position', 1) > 1:
                products[handle]['images'].append({
                    'src': row.get('Image Src'),
                    'position': row.get('Image Position', 1),
                    'alt_text': row.get('Image Alt Text', '')
                })

        self.products = products
        return products

    def get_products_by_tag(self, tag: str) -> List[Dict]:
        filtered = []
        for handle, product in self.products.items():
            tags = str(product.get('tags', '')).lower()
            if tag.lower() in tags:
                filtered.append(product)
        return filtered

    def get_all_tags(self) -> set:
        tags = set()
        for product in self.products.values():
            product_tags = str(product.get('tags', ''))
            if product_tags:
                for tag in product_tags.split(','):
                    tags.add(tag.strip())
        return tags

    def get_product_by_handle(self, handle: str) -> Optional[Dict]:
        return self.products.get(handle)

    def group_products_by_tag(self) -> Dict[str, List[Dict]]:
        grouped = {}
        for product in self.products.values():
            product_tags = str(product.get('tags', ''))
            if not product_tags:
                if 'untagged' not in grouped:
                    grouped['untagged'] = []
                grouped['untagged'].append(product)
            else:
                for tag in product_tags.split(','):
                    tag = tag.strip()
                    if tag:
                        if tag not in grouped:
                            grouped[tag] = []
                        grouped[tag].append(product)
        return grouped