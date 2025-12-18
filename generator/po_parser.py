# -*- coding: utf-8 -*-
"""
PO Parser - 발주서(Purchase Order) Excel 파싱 모듈
다양한 PO 형식을 자동으로 감지하여 SKU ID와 수량을 추출
"""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional, Tuple
import io

try:
    from .sku_master import get_sku_master
except ImportError:
    from sku_master import get_sku_master


class POParser:
    """PO(발주서) 파싱 클래스"""

    # SKU ID 컬럼 후보
    SKU_COLUMN_CANDIDATES = ['SKU ID', 'SKU_ID', 'SKUID', 'SKU', 'ITEM CODE', 'ITEM_CODE']

    # 수량 컬럼 후보
    QTY_COLUMN_CANDIDATES = [
        'TOTAL PRODUCT QTY',
        'ORDER QUANTITIES',
        'ORDER QTY',
        'QTY',
        'QUANTITY',
        '수량'
    ]

    # 영문 상품명 컬럼 후보 (우선)
    NAME_EN_COLUMN_CANDIDATES = [
        'Product Name_EN',
        'PRODUCT NAME_EN',
        'ENGLISH NAME',
        'DESCRIPTION_EN',
        'PRODUCT NAME',
        'DESCRIPTION'
    ]

    # 국문 상품명 컬럼 후보
    NAME_KR_COLUMN_CANDIDATES = [
        'Product Name_KR',
        'PRODUCT NAME_KR',
        '상품명',
        'KOREAN NAME',
        'ITEM NAME'
    ]

    # 바코드 컬럼 후보
    BARCODE_COLUMN_CANDIDATES = [
        'GTIN-8 CODE',
        'BARCODE',
        'EAN',
        'UPC'
    ]

    # 단가 컬럼 후보 (FOB, VAT 제외 가격)
    PRICE_COLUMN_CANDIDATES = [
        'Supply Price',
        'SUPPLY PRICE',
        'FOB',
        'UNIT PRICE',
        'PRICE',
        '단가',
        '공급가',
        'FOB PRICE'
    ]

    # 금액 컬럼 후보
    AMOUNT_COLUMN_CANDIDATES = [
        'AMOUNT',
        'TOTAL AMOUNT',
        'AMOUNT(KRW)',
        'AMOUNT (KRW)',
        'AMOUNT(USD)',
        'AMOUNT (USD)',
        '금액',
        'TOTAL'
    ]

    # 통화 컬럼 후보
    CURRENCY_COLUMN_CANDIDATES = [
        'CURRENCY',
        '통화',
        'CURR'
    ]

    # HS CODE 컬럼 후보
    HS_CODE_COLUMN_CANDIDATES = [
        'HS CODE',
        'HS_CODE',
        'HSCODE',
        'HS CODE(TARIFF)',
        'TARIFF CODE'
    ]

    def __init__(self):
        self.sku_master = get_sku_master()

    def detect_currency_from_column(self, col_name: str) -> Optional[str]:
        """
        컬럼명에서 통화 감지
        예: 'AMOUNT(KRW)' -> 'KRW', 'Supply Price (USD)' -> 'USD'
        """
        if not col_name:
            return None
        col_upper = str(col_name).upper()
        if 'KRW' in col_upper or '원' in col_name:
            return 'KRW'
        if 'USD' in col_upper or '$' in col_name:
            return 'USD'
        if 'EUR' in col_upper or '€' in col_name:
            return 'EUR'
        if 'JPY' in col_upper or '¥' in col_name:
            return 'JPY'
        return None

    def find_header_row(self, df: pd.DataFrame) -> int:
        """
        SKU ID 컬럼이 있는 header row 찾기

        Args:
            df: header=None으로 읽은 DataFrame

        Returns:
            header row index
        """
        for idx, row in df.iterrows():
            row_values = [str(v).upper().strip() for v in row.values if pd.notna(v)]
            for candidate in self.SKU_COLUMN_CANDIDATES:
                if candidate.upper() in row_values:
                    return idx
        return 0

    def find_column(self, columns: List[str], candidates: List[str]) -> Optional[str]:
        """
        컬럼명 후보 중 일치하는 컬럼 찾기

        Args:
            columns: DataFrame 컬럼명 리스트
            candidates: 찾을 컬럼명 후보 리스트

        Returns:
            일치하는 컬럼명 또는 None
        """
        columns_upper = {c.upper().strip(): c for c in columns if pd.notna(c)}

        for candidate in candidates:
            candidate_upper = candidate.upper().strip()
            if candidate_upper in columns_upper:
                return columns_upper[candidate_upper]

            # 부분 일치
            for col_upper, col_original in columns_upper.items():
                if candidate_upper in col_upper:
                    return col_original

        return None

    def parse(self, file, filename: str = "", options: dict = None) -> Tuple[List[Dict], List[str]]:
        """
        PO Excel 파일 파싱

        Args:
            file: 업로드된 파일 객체 (BytesIO 또는 파일 경로)
            filename: 파일명 (확장자 확인용)
            options: 파싱 옵션
                - include_price: 단가/통화 포함 (기본 True)
                - include_hs_code: HS CODE 포함 (기본 False)
                - include_name_kr: 국문명 별도 포함 (기본 False)

        Returns:
            (파싱된 아이템 리스트, 에러/경고 메시지 리스트)
        """
        # 기본 옵션
        if options is None:
            options = {}
        include_price = options.get('include_price', True)
        include_hs_code = options.get('include_hs_code', False)
        include_name_kr = options.get('include_name_kr', False)
        items = []
        messages = []

        try:
            # 1. 전체 읽기 (header 없이)
            # openpyxl 필터 에러 회피를 위해 nrows 제한
            import warnings
            warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

            if isinstance(file, (str, Path)):
                try:
                    raw_df = pd.read_excel(file, header=None, engine='openpyxl')
                except ValueError:
                    # 필터 에러 시 xlrd 시도
                    raw_df = pd.read_excel(file, header=None, engine='xlrd')
            else:
                try:
                    raw_df = pd.read_excel(file, header=None, engine='openpyxl')
                except ValueError:
                    file.seek(0)
                    raw_df = pd.read_excel(file, header=None, engine='xlrd')

            # 2. Header row 찾기
            header_row = self.find_header_row(raw_df)
            messages.append(f"Header row detected: row {header_row + 1}")

            # 3. 해당 row를 header로 다시 읽기
            if isinstance(file, (str, Path)):
                try:
                    df = pd.read_excel(file, header=header_row, engine='openpyxl')
                except ValueError:
                    df = pd.read_excel(file, header=header_row, engine='xlrd')
            else:
                file.seek(0)  # BytesIO의 경우 처음으로 돌아가기
                try:
                    df = pd.read_excel(file, header=header_row, engine='openpyxl')
                except ValueError:
                    file.seek(0)
                    df = pd.read_excel(file, header=header_row, engine='xlrd')

            # 4. 컬럼 찾기
            sku_col = self.find_column(df.columns.tolist(), self.SKU_COLUMN_CANDIDATES)
            qty_col = self.find_column(df.columns.tolist(), self.QTY_COLUMN_CANDIDATES)
            name_en_col = self.find_column(df.columns.tolist(), self.NAME_EN_COLUMN_CANDIDATES)
            name_kr_col = self.find_column(df.columns.tolist(), self.NAME_KR_COLUMN_CANDIDATES)
            barcode_col = self.find_column(df.columns.tolist(), self.BARCODE_COLUMN_CANDIDATES)
            price_col = self.find_column(df.columns.tolist(), self.PRICE_COLUMN_CANDIDATES)
            amount_col = self.find_column(df.columns.tolist(), self.AMOUNT_COLUMN_CANDIDATES)
            hs_code_col = self.find_column(df.columns.tolist(), self.HS_CODE_COLUMN_CANDIDATES) if include_hs_code else None

            if not sku_col:
                messages.append("⚠️ SKU ID 컬럼을 찾을 수 없습니다")
                return items, messages

            if not qty_col:
                messages.append("⚠️ 수량(QTY) 컬럼을 찾을 수 없습니다")
                return items, messages

            # 통화 컬럼 또는 컬럼명에서 통화 감지
            currency_col = self.find_column(df.columns.tolist(), self.CURRENCY_COLUMN_CANDIDATES)
            detected_currency = None

            # 1. 통화 컬럼이 있으면 사용
            # 2. 없으면 금액/단가 컬럼명에서 감지
            if not currency_col:
                if amount_col:
                    detected_currency = self.detect_currency_from_column(amount_col)
                if not detected_currency and price_col:
                    detected_currency = self.detect_currency_from_column(price_col)

            messages.append(f"SKU column: {sku_col}")
            messages.append(f"QTY column: {qty_col}")
            if name_en_col:
                messages.append(f"Name (EN) column: {name_en_col}")
            if name_kr_col:
                messages.append(f"Name (KR) column: {name_kr_col}")
            if price_col and include_price:
                messages.append(f"Price column: {price_col}")
            if amount_col and include_price:
                messages.append(f"Amount column: {amount_col}")
            if currency_col:
                messages.append(f"Currency column: {currency_col}")
            elif detected_currency:
                messages.append(f"Detected currency: {detected_currency}")
            if hs_code_col:
                messages.append(f"HS CODE column: {hs_code_col}")

            # 5. 데이터 추출
            not_found_skus = []

            for idx, row in df.iterrows():
                sku_id = str(row.get(sku_col, '')).strip()
                if not sku_id or sku_id == 'nan':
                    continue

                # 수량 파싱
                qty_value = row.get(qty_col, 0)
                try:
                    qty = int(float(qty_value)) if pd.notna(qty_value) else 0
                except (ValueError, TypeError):
                    qty = 0

                if qty <= 0:
                    continue

                # PO에서 직접 가져온 정보 (영문 우선, 없으면 국문)
                po_name_en = str(row.get(name_en_col, '')) if name_en_col else ''
                po_name_kr = str(row.get(name_kr_col, '')) if name_kr_col else ''
                if po_name_en == 'nan':
                    po_name_en = ''
                if po_name_kr == 'nan':
                    po_name_kr = ''
                # Description: 영문 우선, 없으면 국문
                po_description = po_name_en if po_name_en else po_name_kr

                po_barcode = str(row.get(barcode_col, '')) if barcode_col else ''
                if po_barcode == 'nan':
                    po_barcode = ''

                # HS CODE 파싱
                po_hs_code = ''
                if hs_code_col:
                    hs_value = str(row.get(hs_code_col, ''))
                    if hs_value and hs_value != 'nan':
                        po_hs_code = hs_value

                # 단가 파싱
                unit_price = 0.0
                if price_col:
                    price_value = row.get(price_col, 0)
                    try:
                        unit_price = float(price_value) if pd.notna(price_value) else 0.0
                    except (ValueError, TypeError):
                        unit_price = 0.0

                # 금액 파싱
                amount = 0.0
                if amount_col:
                    amount_value = row.get(amount_col, 0)
                    try:
                        amount = float(amount_value) if pd.notna(amount_value) else 0.0
                    except (ValueError, TypeError):
                        amount = 0.0

                # 통화 파싱 (컬럼에서 또는 감지된 통화 사용)
                currency = None
                if currency_col:
                    currency_value = str(row.get(currency_col, '')).strip().upper()
                    if currency_value and currency_value != 'NAN':
                        currency = currency_value
                if not currency:
                    currency = detected_currency  # 컬럼명에서 감지된 통화

                # SKU Master에서 조회
                sku_info = self.sku_master.get_by_sku_id(sku_id)

                if sku_info:
                    # SKU Master description 우선, 없으면 PO description
                    description = sku_info.get('description', '') or po_description
                    items.append({
                        'sku_id': sku_id,
                        'barcode': sku_info.get('barcode', po_barcode),
                        'description': description,
                        'name_kr': po_name_kr if include_name_kr else '',
                        'qty': qty,
                        'unit_price': unit_price if include_price else 0.0,
                        'amount': amount if include_price else 0.0,
                        'currency': currency if include_price else None,
                        'hs_code': po_hs_code or sku_info.get('hs_code', ''),
                        'source': 'master',
                        'manufacturer': sku_info.get('manufacturer', ''),
                        'variant_status': sku_info.get('variant_status', '')
                    })
                else:
                    # SKU Master에 없는 경우
                    not_found_skus.append(sku_id)
                    description = po_description if po_description else f'[미등록] {sku_id}'
                    items.append({
                        'sku_id': sku_id,
                        'barcode': po_barcode,
                        'description': description,
                        'name_kr': po_name_kr if include_name_kr else '',
                        'qty': qty,
                        'unit_price': unit_price if include_price else 0.0,
                        'amount': amount if include_price else 0.0,
                        'currency': currency if include_price else None,
                        'hs_code': po_hs_code,
                        'source': 'po_only',
                        'manufacturer': '',
                        'variant_status': ''
                    })

            # 결과 메시지
            messages.append(f"✅ 총 {len(items)}개 상품 파싱 완료")
            if not_found_skus:
                messages.append(f"⚠️ SKU Master에 없는 SKU: {', '.join(not_found_skus[:5])}" +
                              (f" 외 {len(not_found_skus) - 5}개" if len(not_found_skus) > 5 else ""))

        except Exception as e:
            messages.append(f"❌ 파싱 오류: {str(e)}")

        return items, messages

    def parse_multiple_sheets(self, file, filename: str = "") -> Tuple[List[Dict], List[str]]:
        """
        여러 시트가 있는 Excel 파일 파싱

        Args:
            file: 업로드된 파일 객체
            filename: 파일명

        Returns:
            (파싱된 아이템 리스트, 메시지 리스트)
        """
        all_items = []
        all_messages = []

        try:
            # 시트 목록 확인
            if isinstance(file, (str, Path)):
                xl = pd.ExcelFile(file, engine='openpyxl')
            else:
                file.seek(0)
                xl = pd.ExcelFile(file, engine='openpyxl')

            sheet_names = xl.sheet_names
            all_messages.append(f"발견된 시트: {', '.join(sheet_names)}")

            # 각 시트 파싱
            for sheet_name in sheet_names:
                if isinstance(file, (str, Path)):
                    sheet_df = pd.read_excel(file, sheet_name=sheet_name, header=None, engine='openpyxl')
                else:
                    file.seek(0)
                    sheet_df = pd.read_excel(file, sheet_name=sheet_name, header=None, engine='openpyxl')

                # SKU ID 컬럼이 있는 시트만 파싱
                header_row = self.find_header_row(sheet_df)
                if header_row > 0 or any(
                    any(c.upper() in str(v).upper() for c in self.SKU_COLUMN_CANDIDATES)
                    for v in sheet_df.iloc[0].values if pd.notna(v)
                ):
                    if isinstance(file, (str, Path)):
                        items, messages = self.parse(file, filename)
                    else:
                        file.seek(0)
                        items, messages = self.parse(file, filename)

                    all_items.extend(items)
                    all_messages.extend([f"[{sheet_name}] {m}" for m in messages])
                    break  # 첫 번째 유효한 시트만 처리

        except Exception as e:
            all_messages.append(f"❌ Excel 파일 처리 오류: {str(e)}")

        return all_items, all_messages


def parse_po_file(file, filename: str = "", options: dict = None) -> Tuple[List[Dict], List[str]]:
    """
    PO 파일 파싱 헬퍼 함수

    Args:
        file: 업로드된 파일 객체
        filename: 파일명
        options: 파싱 옵션 (include_price, include_hs_code, include_name_kr)

    Returns:
        (파싱된 아이템 리스트, 메시지 리스트)
    """
    parser = POParser()
    return parser.parse(file, filename, options)


if __name__ == "__main__":
    # 테스트
    import sys

    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        parser = POParser()
        items, messages = parser.parse(test_file, Path(test_file).name)

        print("=== Messages ===")
        for msg in messages:
            print(f"  {msg}")

        print("\n=== Items ===")
        for item in items[:10]:
            print(f"  {item['sku_id']} | {item['qty']}EA | {item['description'][:40]}...")
    else:
        print("Usage: python po_parser.py <po_file.xlsx>")
