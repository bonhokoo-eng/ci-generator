# -*- coding: utf-8 -*-
"""
Data Store - RECEIVER/Staff 정보 저장소
JSON 파일 기반 영구 저장
"""

import json
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime


class DataStore:
    """RECEIVER/Staff 데이터 저장소"""

    def __init__(self, data_dir: Optional[str] = None):
        """
        초기화

        Args:
            data_dir: 데이터 저장 디렉토리 (기본: ./data)
        """
        if data_dir is None:
            data_dir = Path(__file__).parent.parent / "data"
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)

        self.receivers_file = self.data_dir / "receivers.json"
        self.staff_file = self.data_dir / "staff.json"
        self.history_file = self.data_dir / "invoice_history.json"

        # 파일이 없으면 초기화
        self._init_files()

    def _init_files(self):
        """JSON 파일 초기화"""
        if not self.receivers_file.exists():
            self._save_json(self.receivers_file, [])

        if not self.staff_file.exists():
            # 기본 Staff 목록 (이름 → 이메일 규칙)
            default_staff = [
                {"name": "구본호", "email": "bonho.koo@beautyselection.com", "phone": ""},
                {"name": "김민수", "email": "minsu.kim@beautyselection.com", "phone": ""},
            ]
            self._save_json(self.staff_file, default_staff)

        if not self.history_file.exists():
            self._save_json(self.history_file, [])

    def _load_json(self, path: Path) -> List[Dict]:
        """JSON 파일 로드"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return []

    def _save_json(self, path: Path, data: List[Dict]):
        """JSON 파일 저장"""
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    # ========== RECEIVER ==========

    def get_receivers(self) -> List[Dict]:
        """
        저장된 RECEIVER 목록 조회

        Returns:
            [{customer_code, company_name, address, country, email, phone, currency, trade_terms}, ...]
        """
        return self._load_json(self.receivers_file)

    def get_receiver_by_code(self, customer_code: str) -> Optional[Dict]:
        """
        고객 코드로 RECEIVER 조회

        Args:
            customer_code: 고객 코드 (e.g., 'SK', 'OT')

        Returns:
            RECEIVER dict 또는 None
        """
        receivers = self.get_receivers()
        for r in receivers:
            if r.get('customer_code', '').upper() == customer_code.upper():
                return r
        return None

    def save_receiver(self, receiver: Dict) -> bool:
        """
        RECEIVER 저장 (기존 있으면 업데이트)

        Args:
            receiver: RECEIVER dict (customer_code 필수)

        Returns:
            성공 여부
        """
        if not receiver.get('customer_code'):
            return False

        receivers = self.get_receivers()
        code = receiver['customer_code'].upper()

        # 기존 있으면 업데이트
        updated = False
        for i, r in enumerate(receivers):
            if r.get('customer_code', '').upper() == code:
                receivers[i] = receiver
                updated = True
                break

        # 없으면 추가
        if not updated:
            receiver['customer_code'] = code
            receivers.append(receiver)

        self._save_json(self.receivers_file, receivers)
        return True

    def delete_receiver(self, customer_code: str) -> bool:
        """RECEIVER 삭제"""
        receivers = self.get_receivers()
        code = customer_code.upper()
        new_list = [r for r in receivers if r.get('customer_code', '').upper() != code]

        if len(new_list) < len(receivers):
            self._save_json(self.receivers_file, new_list)
            return True
        return False

    # ========== STAFF ==========

    def get_staff_list(self) -> List[Dict]:
        """
        Staff 목록 조회

        Returns:
            [{name, email, phone}, ...]
        """
        return self._load_json(self.staff_file)

    def get_staff_by_name(self, name: str) -> Optional[Dict]:
        """이름으로 Staff 조회"""
        staff_list = self.get_staff_list()
        for s in staff_list:
            if s.get('name') == name:
                return s
        return None

    def get_staff_by_email(self, email: str) -> Optional[Dict]:
        """이메일로 Staff 조회"""
        staff_list = self.get_staff_list()
        for s in staff_list:
            if s.get('email', '').lower() == email.lower():
                return s
        return None

    def save_staff(self, staff: Dict) -> bool:
        """
        Staff 저장 (이메일 기준 업데이트)

        Args:
            staff: {name, email, phone}

        Returns:
            성공 여부
        """
        if not staff.get('email'):
            return False

        staff_list = self.get_staff_list()
        email = staff['email'].lower()

        # 기존 있으면 업데이트
        updated = False
        for i, s in enumerate(staff_list):
            if s.get('email', '').lower() == email:
                staff_list[i] = staff
                updated = True
                break

        if not updated:
            staff_list.append(staff)

        self._save_json(self.staff_file, staff_list)
        return True

    def delete_staff(self, email: str) -> bool:
        """
        Staff 삭제

        Args:
            email: 삭제할 Staff 이메일

        Returns:
            성공 여부
        """
        staff_list = self.get_staff_list()
        new_list = [s for s in staff_list if s.get('email', '').lower() != email.lower()]

        if len(new_list) < len(staff_list):
            self._save_json(self.staff_file, new_list)
            return True
        return False

    # ========== INVOICE HISTORY ==========

    def get_invoice_history(self, limit: int = 50) -> List[Dict]:
        """
        Invoice 발행 히스토리 조회

        Args:
            limit: 최대 개수

        Returns:
            [{invoice_no, customer_code, date, total, currency}, ...]
        """
        history = self._load_json(self.history_file)
        return history[-limit:]

    def add_invoice_history(self, invoice_info: Dict):
        """
        Invoice 발행 기록 추가

        Args:
            invoice_info: {invoice_no, customer_code, date, total, currency, item_count}
        """
        history = self._load_json(self.history_file)
        invoice_info['created_at'] = datetime.now().isoformat()
        history.append(invoice_info)

        # 최대 500개까지만 유지
        if len(history) > 500:
            history = history[-500:]

        self._save_json(self.history_file, history)

    def get_next_sequence(self, customer_code: str, date: datetime) -> int:
        """
        해당 날짜의 다음 Invoice 시퀀스 번호

        Args:
            customer_code: 고객 코드
            date: 날짜

        Returns:
            다음 시퀀스 번호
        """
        history = self._load_json(self.history_file)
        date_str = date.strftime('%y%m%d')
        prefix = f"INV{customer_code.upper()}{date_str}-"

        max_seq = 0
        for h in history:
            inv_no = h.get('invoice_no', '')
            if inv_no.startswith(prefix):
                try:
                    seq = int(inv_no.split('-')[-1])
                    max_seq = max(max_seq, seq)
                except:
                    pass

        return max_seq + 1


# 싱글톤 인스턴스
_data_store: Optional[DataStore] = None


def get_data_store(data_dir: Optional[str] = None) -> DataStore:
    """DataStore 싱글톤 인스턴스 반환"""
    global _data_store
    if _data_store is None:
        _data_store = DataStore(data_dir)
    return _data_store


if __name__ == "__main__":
    # 테스트
    store = DataStore()

    print("=== Staff 목록 ===")
    for s in store.get_staff_list():
        print(f"  {s['name']}: {s['email']}")

    print("\n=== RECEIVER 저장 테스트 ===")
    store.save_receiver({
        'customer_code': 'SK',
        'company_name': 'SokoGlam',
        'address': '123 Test St, New York, NY',
        'country': 'USA',
        'email': 'contact@sokoglam.com',
        'currency': 'USD',
        'trade_terms': 'FOB'
    })
    print(f"  Saved: {store.get_receiver_by_code('SK')}")

    print("\n=== 시퀀스 테스트 ===")
    next_seq = store.get_next_sequence('SK', datetime.now())
    print(f"  Next sequence for SK: {next_seq}")
