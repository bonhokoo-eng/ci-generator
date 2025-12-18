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

    # 제품명 컬럼 후보
    NAME_COLUMN_CANDIDATES = [
        'Product Name_EN',
        'PRODUCT NAME',
        'DESCRIPTION',
        '상품명',
        'ITEM NAME'
    ]

    # 바코드 컬럼 후보
    BARCODE_COLUMN_CANDIDATES = [
        'GTIN-8 CODE',
        'BARCODE',
        'EAN',
        'UPC'
    ]

    def __init__(self):
        self.sku_master = get_sku_master()

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

    def parse(self, file, filename: str = "") -> Tuple[List[Dict], List[str]]:
        """
        PO Excel 파일 파싱

        Args:
            file: 업로드된 파일 객체 (BytesIO 또는 파일 경로)
            filename: 파일명 (확장자 확인용)

        Returns:
            (파싱된 아이템 리스트, 에러/경고 메시지 리스트)
        """
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
            name_col = self.find_column(df.columns.tolist(), self.NAME_COLUMN_CANDIDATES)
            barcode_col = self.find_column(df.columns.tolist(), self.BARCODE_COLUMN_CANDIDATES)

            if not sku_col:
                messages.append("⚠️ SKU ID 컬럼을 찾을 수 없습니다")
                return items, messages

            if not qty_col:
                messages.append("⚠️ 수량(QTY) 컬럼을 찾을 수 없습니다")
                return items, messages

            messages.append(f"SKU column: {sku_col}")
            messages.append(f"QTY column: {qty_col}")

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

                # PO에서 직접 가져온 정보
                po_name = str(row.get(name_col, '')) if name_col else ''
                po_barcode = str(row.get(barcode_col, '')) if barcode_col else ''
                if po_barcode == 'nan':
                    po_barcode = ''

                # SKU Master에서 조회
                sku_info = self.sku_master.get_by_sku_id(sku_id)

                if sku_info:
                    items.append({
                        'sku_id': sku_id,
                        'barcode': sku_info.get('barcode', po_barcode),
                        'description': sku_info.get('description', po_name),
                        'qty': qty,
                        'unit_price': 0.0,
                        'hs_code': sku_info.get('hs_code', ''),
                        'source': 'master',  # SKU Master에서 매칭됨
                        'manufacturer': sku_info.get('manufacturer', ''),
                        'variant_status': sku_info.get('variant_status', '')
                    })
                else:
                    # SKU Master에 없는 경우
                    not_found_skus.append(sku_id)
                    items.append({
                        'sku_id': sku_id,
                        'barcode': po_barcode,
                        'description': po_name if po_name and po_name != 'nan' else f'[미등록] {sku_id}',
                        'qty': qty,
                        'unit_price': 0.0,
                        'hs_code': '',
                        'source': 'po_only',  # PO에서만 가져옴
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


def parse_po_file(file, filename: str = "") -> Tuple[List[Dict], List[str]]:
    """
    PO 파일 파싱 헬퍼 함수

    Args:
        file: 업로드된 파일 객체
        filename: 파일명

    Returns:
        (파싱된 아이템 리스트, 메시지 리스트)
    """
    parser = POParser()
    return parser.parse(file, filename)


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
