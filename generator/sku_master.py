# -*- coding: utf-8 -*-
"""
SKU Master - CSV/Google Sheets 기반 SKU 검색 모듈
BigQuery → Google Sheets → CI Generator 연동
"""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional, Union
import os
import io


# Google Sheets 설정
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1tQsyTHMhJXfV5xDzNsV2BgTZB7Pjm4OIhYxzG2267PI/edit"
GSHEET_WORKSHEET = "sku_master"  # 시트 이름 (gid=1951994186)


class SKUMaster:
    """SKU Master 데이터 검색 클래스"""

    # 기본 CSV 경로 (로컬용)
    DEFAULT_CSV_PATH = r"C:\Users\user\Downloads\bs_sku_master_251217.csv"

    def __init__(self, csv_path: Optional[str] = None):
        """
        SKU Master 초기화

        Args:
            csv_path: CSV 파일 경로 (없으면 기본 경로 사용)
        """
        self.csv_path = csv_path or self.DEFAULT_CSV_PATH
        self._df: Optional[pd.DataFrame] = None
        self._source: str = ""  # 데이터 소스 (local, gsheet, upload)

    def load_from_gsheet(self) -> bool:
        """
        Google Sheets에서 SKU Master 로드 (Streamlit Cloud용)
        st.secrets에 GCP 서비스 계정 정보가 필요

        Returns:
            성공 여부

        Raises:
            Exception: 연결 실패 시 상세 에러 메시지와 함께 예외 발생
        """
        import streamlit as st
        from streamlit_gsheets import GSheetsConnection

        # Streamlit secrets에서 연결
        conn = st.connection("gsheets", type=GSheetsConnection)

        # 시트 읽기 (캐시 TTL 1시간)
        self._df = conn.read(
            spreadsheet=GSHEET_URL,
            worksheet=GSHEET_WORKSHEET,
            ttl=3600  # 1시간 캐시
        )

        self._process_barcode()
        self._source = "gsheet"
        return True

    def get_source(self) -> str:
        """데이터 소스 반환"""
        return self._source

    def load_from_dataframe(self, df: pd.DataFrame):
        """
        DataFrame에서 직접 로드 (Streamlit Cloud용)

        Args:
            df: 로드된 DataFrame
        """
        self._df = df.copy()
        self._process_barcode()
        self._source = "dataframe"

    def load_from_bytes(self, file_bytes: Union[bytes, io.BytesIO], encoding: str = 'utf-8'):
        """
        바이트 데이터에서 로드 (Streamlit 파일 업로드용)

        Args:
            file_bytes: CSV 파일 바이트
            encoding: 인코딩 (기본 utf-8, 한글 cp949)
        """
        if isinstance(file_bytes, bytes):
            file_bytes = io.BytesIO(file_bytes)

        # 여러 인코딩 시도
        encodings = [encoding, 'utf-8', 'cp949', 'euc-kr']
        for enc in encodings:
            try:
                file_bytes.seek(0)
                self._df = pd.read_csv(file_bytes, encoding=enc, dtype={'barcode': str})
                self._process_barcode()
                self._source = "upload"
                return
            except (UnicodeDecodeError, LookupError):
                continue

        raise ValueError("CSV 파일 인코딩을 인식할 수 없습니다")

    def _process_barcode(self):
        """바코드 컬럼 정제"""
        if self._df is None or 'barcode' not in self._df.columns:
            return

        def convert_barcode(val):
            if pd.isna(val) or val == '':
                return ''
            try:
                num = float(val)
                if num > 0:
                    return str(int(num))
            except:
                return str(val)
            return ''

        self._df['barcode_clean'] = self._df['barcode'].apply(convert_barcode)

    def is_loaded(self) -> bool:
        """데이터 로드 여부 확인"""
        return self._df is not None and not self._df.empty

    def _load_data(self) -> pd.DataFrame:
        """CSV 데이터 로드 (lazy loading)"""
        if self._df is None:
            # 로컬 파일이 있으면 로드
            if Path(self.csv_path).exists():
                self._df = pd.read_csv(
                    self.csv_path,
                    encoding='cp949',
                    dtype={'barcode': str}
                )
                self._process_barcode()
                self._source = "local"
            else:
                # 빈 DataFrame 반환 (Cloud 환경)
                return pd.DataFrame()

        return self._df

    def search(self, query: str, limit: int = 20) -> List[Dict]:
        """
        SKU 검색 (variant_code, barcode, variant_name, product_name)

        Args:
            query: 검색어
            limit: 최대 결과 수

        Returns:
            검색 결과 리스트 [{sku_id, barcode, description, hs_code}, ...]
        """
        df = self._load_data()
        query_lower = query.lower().strip()

        if not query_lower:
            return []

        # 검색 조건 (대소문자 무시)
        mask = (
            df['variant_code'].fillna('').str.lower().str.contains(query_lower, regex=False) |
            df['barcode_clean'].fillna('').str.contains(query_lower, regex=False) |
            df['variant_name'].fillna('').str.lower().str.contains(query_lower, regex=False) |
            df['product_name'].fillna('').str.lower().str.contains(query_lower, regex=False)
        )

        results = df[mask].head(limit)

        return [
            {
                'sku_id': row.get('variant_code', ''),
                'barcode': row.get('barcode_clean', ''),
                'description': row.get('variant_full_name', ''),
                'hs_code': '',  # 추후 추가 예정
                'product_name': row.get('product_name', ''),
                'variant_name': row.get('variant_name', ''),
                'manufacturer': row.get('manufacturer', ''),
                'variant_status': row.get('variant_status', '')
            }
            for _, row in results.iterrows()
            if row.get('variant_code')  # variant_code가 있는 것만
        ]

    def get_by_sku_id(self, sku_id: str) -> Optional[Dict]:
        """
        SKU ID로 정확히 조회

        Args:
            sku_id: SKU ID (variant_code)

        Returns:
            SKU 정보 dict 또는 None
        """
        df = self._load_data()
        match = df[df['variant_code'] == sku_id]

        if match.empty:
            return None

        row = match.iloc[0]
        return {
            'sku_id': row.get('variant_code', ''),
            'barcode': row.get('barcode_clean', ''),
            'description': row.get('variant_full_name', ''),
            'hs_code': '',
            'product_name': row.get('product_name', ''),
            'variant_name': row.get('variant_name', ''),
            'manufacturer': row.get('manufacturer', ''),
            'variant_status': row.get('variant_status', '')
        }

    def get_by_barcode(self, barcode: str) -> Optional[Dict]:
        """
        바코드로 정확히 조회

        Args:
            barcode: 바코드 (EAN-13)

        Returns:
            SKU 정보 dict 또는 None
        """
        df = self._load_data()
        barcode_clean = barcode.strip()
        match = df[df['barcode_clean'] == barcode_clean]

        if match.empty:
            return None

        row = match.iloc[0]
        return {
            'sku_id': row.get('variant_code', ''),
            'barcode': row.get('barcode_clean', ''),
            'description': row.get('variant_full_name', ''),
            'hs_code': '',
            'product_name': row.get('product_name', ''),
            'variant_name': row.get('variant_name', ''),
            'manufacturer': row.get('manufacturer', ''),
            'variant_status': row.get('variant_status', '')
        }

    def get_active_products(self, limit: int = 100) -> List[Dict]:
        """
        활성 상품 목록 조회 (variant_status가 '대기 예정'이 아닌 것)

        Args:
            limit: 최대 결과 수

        Returns:
            상품 리스트
        """
        df = self._load_data()

        # variant_code가 있고 판매불가가 아닌 것
        active = df[
            (df['variant_code'].notna()) &
            (df['variant_code'] != '') &
            (df['variant_status'] != '판매불가')
        ].head(limit)

        return [
            {
                'sku_id': row.get('variant_code', ''),
                'barcode': row.get('barcode_clean', ''),
                'description': row.get('variant_full_name', ''),
                'hs_code': '',
                'product_name': row.get('product_name', ''),
                'variant_name': row.get('variant_name', ''),
                'manufacturer': row.get('manufacturer', ''),
                'variant_status': row.get('variant_status', '')
            }
            for _, row in active.iterrows()
        ]


# 싱글톤 인스턴스
_sku_master: Optional[SKUMaster] = None


def get_sku_master(csv_path: Optional[str] = None) -> SKUMaster:
    """SKU Master 싱글톤 인스턴스 반환"""
    global _sku_master
    if _sku_master is None:
        _sku_master = SKUMaster(csv_path)
    return _sku_master


if __name__ == "__main__":
    # 테스트
    master = SKUMaster()

    print("=== SKU 검색 테스트 ===")
    results = master.search("바이오댄스", limit=5)
    for r in results:
        print(f"  {r['sku_id']} | {r['barcode']} | {r['description'][:50]}...")

    print("\n=== 바코드 조회 테스트 ===")
    sku = master.get_by_barcode("8809890063063")
    if sku:
        print(f"  Found: {sku['sku_id']} - {sku['description'][:50]}...")
    else:
        print("  Not found")
