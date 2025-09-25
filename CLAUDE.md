# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Purpose

This is a Shopify product data to Excel price book converter that creates professional wholesale catalogs with multi-language support. It processes Shopify product CSVs and optional translation files to generate formatted Excel files suitable for B2B wholesale distribution.

## Key Commands

### Run the generator
```bash
python generate_pricebook.py
```

### Install dependencies
```bash
pip install -r requirements.txt
```

## Architecture Overview

### Core Processing Flow
1. **ProductExtractor** (`lib/product_extractor.py`) - Parses Shopify product CSV exports and structures product data with variants, images, and tags
2. **TranslationExtractor** (`lib/translation_extractor.py`) - Processes Shopify Translate & Adapt CSV exports to provide multi-language support
3. **PriceBookGenerator** (`generate_pricebook.py`) - Main orchestrator that:
   - Loads configuration from `config.json`
   - Initializes extractors with input files
   - Groups products by tags based on `target_tag` filter
   - Creates styled Excel workbook with company branding
   - Handles image embedding and temporary file management

### Data Flow
- Input CSVs must be placed in `inputs/shopify_product_csv/` (required) and `inputs/shopify_translate_csv/` (optional)
- Configuration in `config.json` controls filtering (`target_tag`), languages (`target_language`), and company branding
- Generated Excel files are saved to `outputs/` with timestamps
- Temporary images are stored in `temp/` during processing and cleaned up after

### Key Configuration Fields
- `target_tag`: Array of tags to filter products (empty = all products)
- `target_language`: Array of language codes for multi-language display (e.g., ["default", "zh-CN"])
- `logo`: Path to company logo file (optional, typically in `assets/`)

### Excel Output Structure
- Header section with company info and optional logo
- Products grouped by tag into sections
- Each product shows: Image, SKU, Product Name (multi-language in same cell), Variant, Wholesale Price
- Professional styling with alternating row colors and modern borders

## Important Implementation Details

### Image Handling
- Product images are downloaded and resized to 100x100px thumbnails
- Temporary images are tracked in `self.temp_images` list and cleaned up after workbook save
- Logo is resized to max 60px height while maintaining aspect ratio

### Translation Display
- When multiple languages are configured, product names appear on separate lines within the same cell
- Translation rows use `\n` line breaks and cell wrap_text is enabled
- Row height automatically adjusts based on number of languages

### Variant Handling
- Products with multiple variants (sizes, colors) are properly expanded
- Variant options with "nan" values are filtered out
- "Default Title" variant option is hidden from display

## Common Modifications

To add new fields to the Excel output, modify:
1. Headers array in `add_product_section()` method
2. Row data building logic in the variant loop
3. Column widths in `set_column_widths()` method

To change styling, update PatternFill colors and Font properties in `add_product_section()` and `add_header()` methods.