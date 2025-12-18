# -*- coding: utf-8 -*-
"""
Google Sheets Client for CI Generator
Fetches master data (customers, products, staff) from Google Sheets
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pathlib import Path
from typing import List, Dict, Optional
import pandas as pd


class SheetsClient:
    """Google Sheets API client for master data"""

    SCOPES = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    def __init__(self, credentials_path: str, spreadsheet_name: str = "CI_Master_Data"):
        """
        Initialize Google Sheets client

        Args:
            credentials_path: Path to Google service account credentials JSON
            spreadsheet_name: Name of the Google Sheets document
        """
        self.credentials_path = Path(credentials_path)
        self.spreadsheet_name = spreadsheet_name
        self._client = None
        self._spreadsheet = None

    def connect(self):
        """Connect to Google Sheets API"""
        if not self.credentials_path.exists():
            raise FileNotFoundError(f"Credentials file not found: {self.credentials_path}")

        creds = ServiceAccountCredentials.from_json_keyfile_name(
            str(self.credentials_path),
            self.SCOPES
        )
        self._client = gspread.authorize(creds)
        self._spreadsheet = self._client.open(self.spreadsheet_name)

    def get_customers(self) -> List[Dict]:
        """
        Get all customers from Customers sheet

        Returns:
            List of customer dictionaries
        """
        if not self._spreadsheet:
            self.connect()

        worksheet = self._spreadsheet.worksheet("Customers")
        records = worksheet.get_all_records()
        return records

    def get_products(self) -> List[Dict]:
        """
        Get all products from Products sheet

        Returns:
            List of product dictionaries
        """
        if not self._spreadsheet:
            self.connect()

        worksheet = self._spreadsheet.worksheet("Products")
        records = worksheet.get_all_records()
        return records

    def get_staff(self) -> List[Dict]:
        """
        Get all staff from Staff sheet

        Returns:
            List of staff dictionaries
        """
        if not self._spreadsheet:
            self.connect()

        worksheet = self._spreadsheet.worksheet("Staff")
        records = worksheet.get_all_records()
        return records

    def get_customer_by_code(self, customer_code: str) -> Optional[Dict]:
        """Get a single customer by code"""
        customers = self.get_customers()
        for c in customers:
            if c.get('customer_code') == customer_code:
                return c
        return None

    def get_product_by_sku(self, sku_id: str) -> Optional[Dict]:
        """Get a single product by SKU ID"""
        products = self.get_products()
        for p in products:
            if p.get('sku_id') == sku_id:
                return p
        return None

    def get_product_by_barcode(self, barcode: str) -> Optional[Dict]:
        """Get a single product by barcode"""
        products = self.get_products()
        for p in products:
            if str(p.get('barcode')) == str(barcode):
                return p
        return None

    def search_products(self, query: str) -> List[Dict]:
        """Search products by SKU or description"""
        products = self.get_products()
        query_lower = query.lower()
        results = []
        for p in products:
            if (query_lower in str(p.get('sku_id', '')).lower() or
                query_lower in str(p.get('description', '')).lower() or
                query_lower in str(p.get('barcode', ''))):
                results.append(p)
        return results


class MockSheetsClient:
    """Mock client for testing without Google Sheets"""

    def __init__(self):
        self._customers = [
            {
                'customer_code': 'SK',
                'company_name': 'Soko Glam Inc',
                'country': 'USA',
                'address': '123 5th Ave, New York, NY 10001',
                'email': 'orders@sokoglam.com',
                'currency': 'USD',
                'trade_terms': 'FOB'
            },
            {
                'customer_code': 'OT',
                'company_name': 'Orien Trade OÜ',
                'country': 'Estonia',
                'address': 'Tallinn, Estonia',
                'email': 'nga.vu@orientrade.com',
                'currency': 'KRW',
                'trade_terms': 'DAP'
            },
            {
                'customer_code': 'ST',
                'company_name': '주식회사 실리콘투',
                'country': 'Korea',
                'address': '서울특별시 강남구',
                'email': 'rosie@siliconii.net',
                'currency': 'KRW',
                'trade_terms': 'N/A'
            },
            {
                'customer_code': 'SP',
                'company_name': 'Sephora Mexico',
                'country': 'Mexico',
                'address': 'Mexico City, Mexico',
                'email': 'vanessa.rojas@sephora.com',
                'currency': 'USD',
                'trade_terms': 'FOB'
            },
        ]

        self._products = [
            {
                'sku_id': 'BIO-S095885625',
                'barcode': '8809891184613',
                'description': '[Biodance] Rejuvenating Caviar PDRN Real Deep Mask (34g*4ea)',
                'hs_code': '3307.90-4000',
                'price_krw': 7800,
                'price_usd': 6.2
            },
            {
                'sku_id': 'BIO-S000004059',
                'barcode': '8809937361060',
                'description': '[Biodance] Bio Collagen-Real Deep Mask 1Box (34g*4ea)_EN',
                'hs_code': '3307.90-4000',
                'price_krw': 7100,
                'price_usd': 8.75
            },
            {
                'sku_id': 'BIO-S000004224',
                'barcode': '8809937361190',
                'description': '[Biodance] Collagen Gel Toner Pads 60Pads',
                'hs_code': '3307.90-4000',
                'price_krw': 8700,
                'price_usd': 11.97
            },
            {
                'sku_id': 'BIO-S023027889',
                'barcode': '8809891185139',
                'description': '[Biodance] Collagen Gel Toner Pads 10Pads',
                'hs_code': '3307.90-4000',
                'price_krw': 10,
                'price_usd': 0
            },
        ]

        self._staff = [
            {'staff_id': 'JBP', 'name': '박정빈', 'email': 'jungbin.park@beautyselection.co.kr'},
            {'staff_id': 'ETL', 'name': '이의태', 'email': 'euitae.lee@beautyselection.co.kr'},
            {'staff_id': 'ESK', 'name': '김은석', 'email': 'eunseok.kim@beautyselection.co.kr'},
            {'staff_id': 'KMK', 'name': '강경미', 'email': 'kyeongmi.kang@beautyselection.co.kr'},
        ]

    def connect(self):
        pass

    def get_customers(self) -> List[Dict]:
        return self._customers

    def get_products(self) -> List[Dict]:
        return self._products

    def get_staff(self) -> List[Dict]:
        return self._staff

    def get_customer_by_code(self, customer_code: str) -> Optional[Dict]:
        for c in self._customers:
            if c['customer_code'] == customer_code:
                return c
        return None

    def get_product_by_sku(self, sku_id: str) -> Optional[Dict]:
        for p in self._products:
            if p['sku_id'] == sku_id:
                return p
        return None

    def get_product_by_barcode(self, barcode: str) -> Optional[Dict]:
        for p in self._products:
            if str(p['barcode']) == str(barcode):
                return p
        return None

    def search_products(self, query: str) -> List[Dict]:
        query_lower = query.lower()
        results = []
        for p in self._products:
            if (query_lower in p['sku_id'].lower() or
                query_lower in p['description'].lower() or
                query_lower in str(p['barcode'])):
                results.append(p)
        return results
