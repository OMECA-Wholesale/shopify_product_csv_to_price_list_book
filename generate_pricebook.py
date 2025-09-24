import json
import os
import sys
import pandas as pd
from datetime import datetime
import requests
from io import BytesIO
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from typing import Dict, List, Optional
import glob

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from lib.product_extractor import ProductExtractor
from lib.translation_extractor import TranslationExtractor

class PriceBookGenerator:
    def __init__(self, config_path: str = "config.json"):
        self.config = self.load_config(config_path)
        self.product_extractor = None
        self.translation_extractor = None
        self.wb = None
        self.ws = None
        self.temp_images = []  # Store temporary image paths

    def load_config(self, config_path: str) -> Dict:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def initialize_extractors(self):
        product_csv_files = glob.glob("inputs/shopify_product_csv/*.csv")
        if not product_csv_files:
            raise FileNotFoundError("No product CSV files found in inputs/shopify_product_csv/")

        self.product_extractor = ProductExtractor(product_csv_files[0])
        self.product_extractor.load_data()
        self.product_extractor.extract_products()

        translation_csv_files = glob.glob("inputs/shopify_translate_csv/*.csv")
        if translation_csv_files:
            self.translation_extractor = TranslationExtractor(translation_csv_files[0])
            self.translation_extractor.load_data()
            self.translation_extractor.extract_translations()

    def create_workbook(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "Price List"

    def add_header(self, start_row: int = 1) -> int:
        header_font = Font(size=16, bold=True)
        company_font = Font(size=20, bold=True)

        self.ws.merge_cells(f'A{start_row}:E{start_row}')
        self.ws[f'A{start_row}'] = self.config.get('company_name', 'Company Name')
        self.ws[f'A{start_row}'].font = company_font
        self.ws[f'A{start_row}'].alignment = Alignment(horizontal='center', vertical='center')

        info_row = start_row + 1
        self.ws.merge_cells(f'A{info_row}:E{info_row}')
        contact_info = f"PHONE: {self.config.get('phone', '')}  "
        if self.config.get('email'):
            contact_info += f"Email: {self.config.get('email', '')}"
        self.ws[f'A{info_row}'] = contact_info
        self.ws[f'A{info_row}'].alignment = Alignment(horizontal='center')

        if self.config.get('address'):
            addr_row = info_row + 1
            self.ws.merge_cells(f'A{addr_row}:E{addr_row}')
            self.ws[f'A{addr_row}'] = self.config.get('address', '')
            self.ws[f'A{addr_row}'].alignment = Alignment(horizontal='center')
            return addr_row + 2

        return info_row + 2

    def add_product_section(self, products: List[Dict], section_title: str, start_row: int) -> int:
        section_font = Font(size=14, bold=True)
        header_font = Font(size=11, bold=True)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

        self.ws.merge_cells(f'A{start_row}:E{start_row}')
        self.ws[f'A{start_row}'] = section_title.upper()
        self.ws[f'A{start_row}'].font = section_font
        self.ws[f'A{start_row}'].alignment = Alignment(horizontal='center', vertical='center')
        self.ws[f'A{start_row}'].fill = header_fill

        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            self.ws[f'{col}{start_row}'].border = border

        current_row = start_row + 1

        if not products:
            return current_row + 1

        headers = ['Image', 'SKU', 'Product Name', 'Variant', 'Wholesale Price']
        for col_idx, header in enumerate(headers, 1):
            cell = self.ws.cell(row=current_row, column=col_idx, value=header)
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
            cell.border = border

        current_row += 1

        for product in products:
            max_height = 1

            image_added = False
            if product.get('images'):
                image_url = product['images'][0].get('src', '')
                if image_url:
                    try:
                        response = requests.get(image_url, timeout=10)
                        if response.status_code == 200:
                            img = Image.open(BytesIO(response.content))

                            max_size = (100, 100)
                            img.thumbnail(max_size, Image.Resampling.LANCZOS)

                            # Create temp directory if it doesn't exist
                            os.makedirs("temp", exist_ok=True)

                            # Use unique filename with product handle to avoid conflicts
                            temp_path = f"temp/temp_img_{product.get('handle', 'unknown')}_{current_row}.png"
                            img.save(temp_path)
                            self.temp_images.append(temp_path)  # Track temp files

                            xl_img = XLImage(temp_path)
                            xl_img.width = 100
                            xl_img.height = 100

                            self.ws.add_image(xl_img, f'A{current_row}')
                            self.ws.row_dimensions[current_row].height = 80
                            image_added = True

                            # Don't delete here - will cleanup after save
                    except Exception as e:
                        print(f"Error loading image: {e}")

            if not image_added:
                self.ws[f'A{current_row}'] = ""
                self.ws.row_dimensions[current_row].height = 30

            for variant in product.get('variants', [{}]):
                # Get language list
                languages = self.config.get('target_language', ['default'])

                # For each language, create a separate row
                for lang_idx, lang in enumerate(languages):
                    row_data = []

                    # Image column - only for first row
                    row_data.append("")

                    # SKU - only for first language row
                    if lang_idx == 0:
                        row_data.append(variant.get('sku', ''))
                    else:
                        row_data.append("")  # Empty for translation rows

                    # Product name in specific language
                    product_name = product.get('title', '')
                    if self.translation_extractor and lang != 'default':
                        translated = self.translation_extractor.get_translated_title(
                            product.get('handle', ''),
                            lang
                        )
                        if translated:
                            product_name = translated

                    row_data.append(product_name)

                    # Variant string - only for first language row
                    if lang_idx == 0:
                        variant_parts = []
                        option1 = variant.get('option1', '')
                        if option1 and option1 != 'Default Title' and str(option1).lower() != 'nan':
                            variant_parts.append(str(option1))
                        option2 = variant.get('option2', '')
                        if option2 and str(option2).lower() != 'nan':
                            variant_parts.append(str(option2))
                        option3 = variant.get('option3', '')
                        if option3 and str(option3).lower() != 'nan':
                            variant_parts.append(str(option3))

                        variant_str = ' / '.join(variant_parts) if variant_parts else ''
                        row_data.append(variant_str)
                    else:
                        row_data.append("")  # Empty for translation rows

                    # Price - only for first language row
                    if lang_idx == 0:
                        price = variant.get('price', 0)
                        try:
                            price_val = float(price) if price else 0
                            row_data.append(f"${price_val:.2f}")
                        except:
                            row_data.append(str(price))
                    else:
                        row_data.append("")  # Empty for translation rows

                    # Write row data to cells
                    for col_idx, value in enumerate(row_data[1:], 2):
                        cell = self.ws.cell(row=current_row, column=col_idx, value=value)
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        cell.border = border

                        # Style translation rows slightly different (lighter text)
                        if lang_idx > 0:
                            cell.font = Font(color="666666")

                    current_row += 1

                # Add spacing between variants if multiple exist
                if len(product.get('variants', [])) > 1:
                    current_row += 1

        return current_row + 1

    def set_column_widths(self):
        column_widths = {
            'A': 15,
            'B': 15,
            'C': 50,
            'D': 12,
            'E': 12
        }

        for col, width in column_widths.items():
            self.ws.column_dimensions[col].width = width

    def generate(self):
        print("Initializing extractors...")
        self.initialize_extractors()

        print("Creating workbook...")
        self.create_workbook()

        print("Adding header...")
        current_row = self.add_header()

        print("Grouping products by tag...")
        grouped_products = self.product_extractor.group_products_by_tag()

        # Filter by target_tag if specified
        target_tags = self.config.get('target_tag', [])
        if target_tags and isinstance(target_tags, list) and len(target_tags) > 0:
            filtered_groups = {}
            for tag in target_tags:
                if tag in grouped_products:
                    filtered_groups[tag] = grouped_products[tag]
            if filtered_groups:  # Only use filtered groups if we found matches
                grouped_products = filtered_groups
            else:
                print(f"Warning: No products found with tags: {target_tags}")

        print(f"Found {len(grouped_products)} product groups")

        for tag, products in grouped_products.items():
            print(f"Adding section: {tag} with {len(products)} products")
            current_row = self.add_product_section(products, tag, current_row)

        print("Setting column widths...")
        self.set_column_widths()

        os.makedirs("outputs", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"outputs/pricebook_{timestamp}.xlsx"

        print(f"Saving to {output_file}...")
        self.wb.save(output_file)
        print(f"Price book generated successfully: {output_file}")

        # Cleanup temporary images after save
        self.cleanup_temp_images()

        return output_file

    def cleanup_temp_images(self):
        """Remove temporary image files after workbook is saved"""
        for temp_path in self.temp_images:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception as e:
                print(f"Warning: Could not remove temp file {temp_path}: {e}")
        self.temp_images = []

def main():
    generator = PriceBookGenerator()
    generator.generate()

if __name__ == "__main__":
    main()