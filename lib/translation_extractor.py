import pandas as pd
from typing import Dict, List, Optional
import os

class TranslationExtractor:
    def __init__(self, csv_path: Optional[str] = None):
        self.csv_path = csv_path
        self.translations = {}
        self.raw_df = None

    def load_data(self) -> bool:
        if not self.csv_path or not os.path.exists(self.csv_path):
            print(f"Translation file not found: {self.csv_path}")
            return False

        try:
            self.raw_df = pd.read_csv(self.csv_path)
            return True
        except Exception as e:
            print(f"Error loading translation CSV: {e}")
            return False

    def extract_translations(self) -> Dict:
        if self.raw_df is None:
            if not self.load_data():
                return {}

        translations = {}

        for _, row in self.raw_df.iterrows():
            item_type = row.get('Type', '')
            identification = str(row.get('Identification', ''))
            field = row.get('Field', '')
            locale = row.get('Locale', '')
            default_content = row.get('Default content', '')
            translated_content = row.get('Translated content', '')

            if pd.isna(item_type) or pd.isna(identification):
                continue

            item_id = identification.split(',')[0].strip("'")

            if item_type not in translations:
                translations[item_type] = {}

            if item_id not in translations[item_type]:
                translations[item_type][item_id] = {}

            if locale not in translations[item_type][item_id]:
                translations[item_type][item_id][locale] = {}

            if field not in translations[item_type][item_id][locale]:
                translations[item_type][item_id][locale][field] = {}

            translations[item_type][item_id][locale][field] = {
                'default': default_content,
                'translated': translated_content
            }

        self.translations = translations
        return translations

    def get_product_translations(self, product_handle: str, locale: str) -> Dict:
        if 'PRODUCT' not in self.translations:
            return {}

        for product_id, locales in self.translations['PRODUCT'].items():
            if locale in locales:
                if 'handle' in locales[locale]:
                    handle_data = locales[locale]['handle']
                    if handle_data.get('default') == product_handle:
                        return locales[locale]
        return {}

    def get_translated_title(self, product_handle: str, locale: str) -> Optional[str]:
        translations = self.get_product_translations(product_handle, locale)
        if 'title' in translations:
            return translations['title'].get('translated', '')
        return None

    def get_variant_translations(self, variant_id: str, locale: str) -> Dict:
        if 'PRODUCT_VARIANT' not in self.translations:
            return {}

        if variant_id in self.translations['PRODUCT_VARIANT']:
            if locale in self.translations['PRODUCT_VARIANT'][variant_id]:
                return self.translations['PRODUCT_VARIANT'][variant_id][locale]
        return {}

    def get_collection_translations(self, collection_id: str, locale: str) -> Dict:
        if 'COLLECTION' not in self.translations:
            return {}

        if collection_id in self.translations['COLLECTION']:
            if locale in self.translations['COLLECTION'][collection_id]:
                return self.translations['COLLECTION'][collection_id][locale]
        return {}

    def get_available_locales(self) -> set:
        locales = set()
        for item_type in self.translations.values():
            for item_id in item_type.values():
                locales.update(item_id.keys())
        return locales

    def build_multilingual_name(self, product_handle: str, default_name: str, target_languages: List[str]) -> str:
        if not target_languages or len(target_languages) == 0:
            return default_name

        name_parts = []

        for lang in target_languages:
            if lang == "default" or lang == "":
                name_parts.append(default_name)
            else:
                translated = self.get_translated_title(product_handle, lang)
                if translated:
                    name_parts.append(translated)
                elif lang != "default":
                    for product_id, locales in self.translations.get('PRODUCT', {}).items():
                        if lang in locales and 'title' in locales[lang]:
                            title_info = locales[lang]['title']
                            if title_info.get('default', '').lower() == default_name.lower():
                                trans = title_info.get('translated', '')
                                if trans:
                                    name_parts.append(trans)
                                    break

        return " ".join(name_parts) if name_parts else default_name