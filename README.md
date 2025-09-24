# Shopify Product CSV to Price List Book

A tool to convert Shopify product export data into a professional Excel price book with multi-language support.

## Features

- Convert Shopify product CSV exports to formatted Excel price books
- Multi-language support via Shopify translation exports
- Group products by tags/collections
- Include product images in Excel
- Customizable company header information
- Professional layout similar to wholesale catalogs

## Project Structure

```
├── inputs/
│   ├── shopify_product_csv/       # Shopify product export CSV files
│   └── shopify_translate_csv/     # Shopify translation export CSV files
├── lib/
│   ├── product_extractor.py       # Product data extraction module
│   └── translation_extractor.py   # Translation data extraction module
├── outputs/                       # Generated Excel price books
├── config.json                     # Configuration file
├── generate_pricebook.py          # Main generation script
└── requirements.txt               # Python dependencies
```

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd shopify_product_csv_to_-price_list_book
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

Edit `config.json` to customize your price book:

```json
{
  "company_name": "Your Company Name",
  "phone": "(xxx) xxx-xxxx",
  "website": "yourwebsite.com",
  "address": "Your Company Address",
  "email": "your@email.com",
  "target_tag": [],              // Tags to filter products (empty = all products)
  "target_language": ["default"] // Language codes for multi-language support
}
```

### Configuration Options

- `target_tag`: Filter products by tags
  - `[]` - Include all products
  - `["bonsai", "cups"]` - Only include products with "bonsai" or "cups" tags
  - Products will be grouped by their tags in the Excel output

- `target_language`: Language codes for multi-language product names
  - `["default"]` - Use only default language from product CSV
  - `["default", "zh-CN"]` - Include both English and Chinese names
  - Example output: "Red T-shirt 红短袖"

## Usage

1. Export your products from Shopify:
   - Go to Shopify Admin → Products → Export
   - Save CSV file to `inputs/shopify_product_csv/`

2. (Optional) Export translations:
   - Install Shopify Translate & Adapt app
   - Export translations to CSV
   - Save to `inputs/shopify_translate_csv/`

3. Configure `config.json` with your company information

4. Run the generator:
```bash
python generate_pricebook.py
```

5. Find your price book in `outputs/` folder

## Input File Format

### Shopify Product CSV
Standard Shopify product export with columns:
- Handle, Title, Vendor, Tags
- Product Category, Type
- Variant SKU, Price
- Image Src, Option Values
- And other standard Shopify fields

### Translation CSV (Optional)
Shopify Translate & Adapt export with columns:
- Type, Identification, Field
- Locale, Default content, Translated content

## Output

The generated Excel price book includes:
- Company header with contact information
- Products grouped by tags/collections
- Product images (when available)
- SKU, name, description, price columns
- Multi-language product names (if configured)
- Professional wholesale catalog formatting

## Example

With configuration:
```json
{
  "target_tag": ["bonsai"],           // Only include products tagged "bonsai"
  "target_language": ["default", "zh-CN"]  // Show English and Chinese names
}
```

Product output:
```
SKU: EP-5
Name: PP HINGED CONTAINER 5*5 250PCS 餐盒5*5 250个
Price: $20.97
```

## Requirements

- Python 3.7+
- pandas
- openpyxl
- Pillow (for image handling)
- requests (for downloading images)