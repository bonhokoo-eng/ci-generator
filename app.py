# -*- coding: utf-8 -*-
"""
CI Generator - Streamlit Web Application
Commercial Invoice Generator for Beautyselection
"""

import streamlit as st
from datetime import datetime
from pathlib import Path
import io
import sys

# Add generator module to path
sys.path.insert(0, str(Path(__file__).parent))

from generator.ci_generator import (
    CIGenerator, InvoiceData, ReceiverInfo, ShippingTerms, LineItem, BANK_DETAILS
)
from generator.sku_master import get_sku_master
from generator.data_store import get_data_store
from generator.po_parser import parse_po_file

# Page config
st.set_page_config(
    page_title="CI Generator",
    page_icon="📄",
    layout="wide"
)

# shadcn/ui Mira Style - Custom CSS
st.markdown("""
<style>
/* ===== shadcn/ui Mira Theme ===== */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* Root Variables */
:root {
    --background: #ffffff;
    --foreground: #09090b;
    --card: #ffffff;
    --card-foreground: #09090b;
    --primary: #18181b;
    --primary-foreground: #fafafa;
    --secondary: #f4f4f5;
    --secondary-foreground: #18181b;
    --muted: #f4f4f5;
    --muted-foreground: #71717a;
    --accent: #f4f4f5;
    --accent-foreground: #18181b;
    --destructive: #ef4444;
    --border: #e4e4e7;
    --input: #e4e4e7;
    --ring: #18181b;
    --radius: 0.5rem;
}

/* Base Typography */
html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* Main App Container */
.stApp {
    background: var(--background);
}

/* Sidebar Styling */
[data-testid="stSidebar"] {
    background: #fafafa;
    border-right: 1px solid var(--border);
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 1.5rem;
}

/* Headers */
h1, h2, h3 {
    font-weight: 600 !important;
    color: var(--foreground) !important;
    letter-spacing: -0.025em !important;
}

h1 { font-size: 1.875rem !important; }
h2 { font-size: 1.25rem !important; }
h3 { font-size: 1.125rem !important; }

/* Subheader in sidebar */
[data-testid="stSidebar"] .stMarkdown h3 {
    font-size: 0.875rem !important;
    font-weight: 500 !important;
    color: var(--muted-foreground) !important;
    text-transform: uppercase;
    letter-spacing: 0.05em !important;
    margin-bottom: 0.75rem !important;
}

/* Text Input */
.stTextInput > div > div > input {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 0.5rem 0.75rem !important;
    font-size: 0.875rem !important;
    background: var(--background) !important;
    transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
}

.stTextInput > div > div > input:focus {
    border-color: var(--ring) !important;
    box-shadow: 0 0 0 2px rgba(24, 24, 27, 0.1) !important;
    outline: none !important;
}

/* Number Input */
.stNumberInput > div > div > input {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 0.5rem 0.75rem !important;
    font-size: 0.875rem !important;
    background: var(--background) !important;
}

/* Select Box */
.stSelectbox > div > div {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    background: var(--background) !important;
}

.stSelectbox > div > div > div {
    font-size: 0.875rem !important;
}

/* Buttons - Primary */
.stButton > button[kind="primary"],
.stButton > button[data-baseweb="button"][kind="primary"] {
    background: var(--primary) !important;
    color: var(--primary-foreground) !important;
    border: none !important;
    border-radius: var(--radius) !important;
    padding: 0.5rem 1rem !important;
    font-weight: 500 !important;
    font-size: 0.875rem !important;
    transition: opacity 0.15s ease !important;
}

.stButton > button[kind="primary"]:hover {
    opacity: 0.9 !important;
}

/* Buttons - Secondary/Default */
.stButton > button {
    background: var(--background) !important;
    color: var(--foreground) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 0.5rem 1rem !important;
    font-weight: 500 !important;
    font-size: 0.875rem !important;
    transition: background 0.15s ease !important;
}

.stButton > button:hover {
    background: var(--accent) !important;
}

/* Sidebar columns alignment for select + delete button */
[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] {
    align-items: flex-end !important;
    gap: 0.5rem !important;
}

/* Delete button in sidebar - shadcn ghost variant style */
[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] .stButton > button {
    width: 36px !important;
    height: 36px !important;
    padding: 0 !important;
    min-height: 36px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    background: transparent !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    color: var(--muted-foreground) !important;
    font-size: 1rem !important;
    transition: all 0.15s ease !important;
}

[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] .stButton > button:hover:not(:disabled) {
    background: var(--destructive) !important;
    border-color: var(--destructive) !important;
    color: white !important;
}

[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] .stButton > button:disabled {
    opacity: 0.4 !important;
    cursor: not-allowed !important;
}

/* Save button in sidebar - primary style */
[data-testid="stSidebar"] .stButton > button:not([data-testid="stHorizontalBlock"] .stButton > button) {
    width: 100% !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 0 !important;
    background: var(--muted) !important;
    border-radius: var(--radius) !important;
    padding: 0.25rem !important;
}

.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border-radius: calc(var(--radius) - 2px) !important;
    padding: 0.5rem 1rem !important;
    font-weight: 500 !important;
    font-size: 0.875rem !important;
    color: var(--muted-foreground) !important;
}

.stTabs [aria-selected="true"] {
    background: var(--background) !important;
    color: var(--foreground) !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1) !important;
}

/* Checkbox */
.stCheckbox > label > div {
    border-radius: 4px !important;
}

/* Divider */
hr {
    border: none !important;
    border-top: 1px solid var(--border) !important;
    margin: 1rem 0 !important;
}

/* Cards (using containers) */
[data-testid="stVerticalBlock"] > div:has(> .stMarkdown) {
    background: var(--card);
    border-radius: var(--radius);
}

/* Info/Warning/Error boxes */
.stAlert {
    border-radius: var(--radius) !important;
    border: 1px solid var(--border) !important;
}

[data-baseweb="notification"] {
    border-radius: var(--radius) !important;
}

/* Dataframe */
.stDataFrame {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
}

/* Expander */
.streamlit-expanderHeader {
    font-weight: 500 !important;
    font-size: 0.875rem !important;
    background: var(--muted) !important;
    border-radius: var(--radius) !important;
}

/* File Uploader */
[data-testid="stFileUploader"] {
    border: 2px dashed var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 1.5rem !important;
}

[data-testid="stFileUploader"]:hover {
    border-color: var(--muted-foreground) !important;
}

/* Caption text */
.stCaption, [data-testid="stCaptionContainer"] {
    color: var(--muted-foreground) !important;
    font-size: 0.75rem !important;
}

/* Success message */
.stSuccess {
    background: #f0fdf4 !important;
    border: 1px solid #bbf7d0 !important;
    color: #166534 !important;
    border-radius: var(--radius) !important;
}

/* Download button */
.stDownloadButton > button {
    background: var(--primary) !important;
    color: var(--primary-foreground) !important;
    border: none !important;
    border-radius: var(--radius) !important;
}

/* Toggle */
[data-testid="stToggle"] label > div {
    border-radius: 9999px !important;
}

/* Metric styling */
[data-testid="stMetric"] {
    background: var(--card) !important;
    padding: 1rem !important;
    border-radius: var(--radius) !important;
    border: 1px solid var(--border) !important;
}

/* Footer */
footer {
    color: var(--muted-foreground) !important;
}

/* Hide Streamlit branding */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'invoice_items' not in st.session_state:
    st.session_state.invoice_items = []


def check_password():
    """비밀번호 인증"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.title("🔐 CI Generator")
    st.caption("로그인이 필요합니다")

    password = st.text_input("Password", type="password", key="password_input")

    if st.button("Login", type="primary"):
        # secrets에서 비밀번호 가져오기 (없으면 기본값 사용)
        correct_password = st.secrets.get("app_password", "b2b7788")
        if password == correct_password:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("잘못된 비밀번호입니다")

    return False


def render_invoice_items(currency, staff_email, receiver_info, data_store, invoice_date,
                         customer_code, order_no, staff_phone, is_domestic, shipping,
                         tax_rate, total_transaction, custom_remarks, key_prefix="main"):
    """상품 목록 표시 및 CI 생성 UI (재사용 가능)"""
    if st.session_state.invoice_items:
        # 단가 수정 가능한 아이템 목록
        st.caption("단가를 수정하면 Total이 자동 계산됩니다")

        total_amount = 0

        for idx, item in enumerate(st.session_state.invoice_items):
            col_info, col_qty, col_price, col_total, col_del = st.columns([4, 1, 2, 2, 1])

            with col_info:
                desc_short = item.description[:30] + '...' if len(item.description) > 30 else item.description
                st.text(f"{idx+1}. {item.sku_id}")
                st.caption(desc_short)

            with col_qty:
                st.text(f"{item.qty} EA")

            with col_price:
                # 단가 수정 가능
                new_price = st.number_input(
                    "단가",
                    min_value=0.0,
                    value=float(item.unit_price),
                    step=0.01 if currency == 'USD' else 100.0,
                    key=f"price_{key_prefix}_{idx}",
                    label_visibility="collapsed"
                )
                # 변경사항 반영
                if new_price != item.unit_price:
                    st.session_state.invoice_items[idx].unit_price = new_price

            with col_total:
                line_total = item.qty * new_price
                total_amount += line_total
                if currency == 'USD':
                    st.text(f"${line_total:,.2f}")
                else:
                    st.text(f"₩{line_total:,.0f}")

            with col_del:
                if st.button("X", key=f"del_item_{key_prefix}_{idx}"):
                    st.session_state.invoice_items.pop(idx)
                    st.rerun()

        # 합계
        st.divider()
        col_total, col_clear = st.columns([3, 1])
        with col_total:
            total_display = f"${total_amount:,.2f}" if currency == 'USD' else f"₩{total_amount:,.0f}"
            st.markdown(f"### Total: {currency} {total_display}")
        with col_clear:
            if st.button("전체 삭제", key=f"clear_btn_{key_prefix}"):
                st.session_state.invoice_items = []
                st.rerun()

        st.divider()

        # ===== CI 생성 =====
        # 필수 필드 검증
        validation_errors = []
        if not staff_email:
            validation_errors.append("담당자 이메일 필수")
        if not receiver_info or not receiver_info.company_name:
            validation_errors.append("고객사 회사명 필수")
        if not receiver_info or not receiver_info.address:
            validation_errors.append("고객사 주소 필수")
        if not st.session_state.invoice_items:
            validation_errors.append("상품 1개 이상 필수")

        if validation_errors:
            st.warning("필수 항목: " + ", ".join(validation_errors))

        generate_disabled = len(validation_errors) > 0

        if st.button("CI 생성", type="primary", disabled=generate_disabled, key=f"gen_btn_{key_prefix}"):
            try:
                generator = CIGenerator()

                # Invoice 번호 생성
                seq = data_store.get_next_sequence(
                    customer_code if customer_code else 'XX',
                    datetime.combine(invoice_date, datetime.min.time())
                )
                invoice_no = generator.generate_invoice_number(
                    customer_code if customer_code else 'XX',
                    datetime.combine(invoice_date, datetime.min.time()),
                    seq
                )

                # InvoiceData 생성
                invoice_data = InvoiceData(
                    invoice_date=datetime.combine(invoice_date, datetime.min.time()),
                    invoice_no=invoice_no,
                    order_no=order_no,
                    staff_email=staff_email,
                    staff_phone=staff_phone,
                    is_domestic=is_domestic,
                    receiver=receiver_info,
                    shipping=shipping,
                    currency=currency,
                    items=st.session_state.invoice_items,
                    tax_rate=tax_rate,
                    total_transaction=total_transaction,
                    custom_remarks=custom_remarks
                )

                # 생성
                wb = generator.generate(invoice_data)

                # 히스토리 저장
                data_store.add_invoice_history({
                    'invoice_no': invoice_no,
                    'customer_code': customer_code if customer_code else 'XX',
                    'date': invoice_date.isoformat(),
                    'total': total_amount,
                    'currency': currency,
                    'item_count': len(st.session_state.invoice_items)
                })

                # 다운로드
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                st.download_button(
                    label="Download Excel",
                    data=output,
                    file_name=f"{invoice_no}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{key_prefix}"
                )
                st.success(f"생성 완료: {invoice_no}")

            except ValueError as e:
                st.error(f"Validation Error: {e}")
            except Exception as e:
                st.error(f"Error: {e}")

    else:
        st.info("상품을 추가하세요")


def main():
    # 인증 체크
    if not check_password():
        return

    st.title("📄 Commercial Invoice Generator")

    # Data stores
    sku_master = get_sku_master()
    data_store = get_data_store()

    # ========== SIDEBAR ==========
    with st.sidebar:
        st.header("Invoice Settings")

        # ===== SKU Master 설정 =====
        with st.expander("SKU Master 설정", expanded=not sku_master.is_loaded()):
            if sku_master.is_loaded():
                source_map = {
                    "local": "로컬 CSV",
                    "gsheet": "Google Sheets",
                    "upload": "업로드됨",
                    "dataframe": "DataFrame"
                }
                source_label = source_map.get(sku_master.get_source(), "로드됨")
                st.success(f"SKU Master: {source_label}")
            else:
                # Google Sheets 자동 연결 시도
                if st.button("Google Sheets 연결", key="gsheet_connect"):
                    with st.spinner("Google Sheets 연결 중..."):
                        if sku_master.load_from_gsheet():
                            st.success("Google Sheets 연결 완료!")
                            st.rerun()
                        else:
                            st.error("연결 실패 - secrets 설정을 확인하세요")

                st.caption("또는 CSV 직접 업로드:")
                sku_file = st.file_uploader(
                    "SKU Master CSV",
                    type=['csv'],
                    key="sku_master_upload",
                    help="bs_sku_master CSV 파일 업로드"
                )

                if sku_file:
                    try:
                        sku_master.load_from_bytes(sku_file.getvalue())
                        st.success("로드 완료!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"로드 실패: {e}")

        st.divider()

        # 국내/해외 모드
        is_domestic = st.toggle("국내 거래 (Domestic)", value=False)
        mode_label = "국내" if is_domestic else "해외 (International)"
        st.caption(f"Mode: {mode_label}")

        st.divider()

        # ===== STAFF (담당자) =====
        st.subheader("담당자 (Staff)")
        staff_list = data_store.get_staff_list()
        staff_names = ["(직접 입력)"] + [s['name'] for s in staff_list]

        # 담당자 선택 + 삭제 버튼
        col_staff, col_staff_del = st.columns([5, 1])
        with col_staff:
            selected_staff_name = st.selectbox("담당자 선택", staff_names, label_visibility="collapsed")
        with col_staff_del:
            can_delete_staff = selected_staff_name != "(직접 입력)"
            if st.button("✕", key="del_staff_btn", disabled=not can_delete_staff, help="담당자 삭제"):
                staff_info = data_store.get_staff_by_name(selected_staff_name)
                if staff_info and data_store.delete_staff(staff_info['email']):
                    st.rerun()

        if selected_staff_name == "(직접 입력)":
            staff_email = st.text_input("이메일 *", placeholder="name@beautyselection.com")
            staff_phone = st.text_input("연락처", placeholder="+82-10-1234-5678")

            # 새 담당자 저장
            if staff_email and st.button("담당자 저장"):
                name_from_email = staff_email.split('@')[0].replace('.', ' ').title()
                data_store.save_staff({
                    'name': name_from_email,
                    'email': staff_email,
                    'phone': staff_phone
                })
                st.success(f"저장됨: {name_from_email}")
                st.rerun()
        else:
            staff_info = data_store.get_staff_by_name(selected_staff_name)
            staff_email = staff_info['email'] if staff_info else ""
            staff_phone = staff_info.get('phone', '') if staff_info else ""
            st.text_input("이메일", value=staff_email, disabled=True)

        st.divider()

        # ===== RECEIVER (고객사) =====
        st.subheader("고객사 (Receiver)")
        receivers = data_store.get_receivers()
        receiver_options = ["(신규 입력)"] + [
            f"{r['customer_code']} - {r['company_name']}" for r in receivers
        ]

        # 고객사 선택 + 삭제 버튼
        col_recv, col_recv_del = st.columns([5, 1])
        with col_recv:
            selected_receiver = st.selectbox("고객사 선택", receiver_options, label_visibility="collapsed")
        with col_recv_del:
            can_delete_recv = selected_receiver != "(신규 입력)"
            if st.button("✕", key="del_recv_btn", disabled=not can_delete_recv, help="고객사 삭제"):
                code_to_del = selected_receiver.split(' - ')[0]
                if data_store.delete_receiver(code_to_del):
                    st.rerun()

        if selected_receiver == "(신규 입력)":
            col1, col2 = st.columns([1, 2])
            with col1:
                customer_code = st.text_input("코드 *", max_chars=3, placeholder="SK")
            with col2:
                company_name = st.text_input("회사명 *", placeholder="SokoGlam Inc.")

            receiver_address = st.text_area("주소 *", height=80, placeholder="123 Test St, New York, NY 10001")
            receiver_country = st.text_input("국가", placeholder="USA") if not is_domestic else ""
            receiver_email = st.text_input("Email")
            receiver_phone = st.text_input("Tel")
            receiver_business_no = st.text_input("사업자번호") if is_domestic else ""

            receiver_info = ReceiverInfo(
                company_name=company_name,
                address=receiver_address,
                country=receiver_country,
                email=receiver_email,
                phone=receiver_phone,
                business_no=receiver_business_no
            )

            # 새 고객사 저장
            if customer_code and company_name and receiver_address:
                if st.button("고객사 저장"):
                    data_store.save_receiver({
                        'customer_code': customer_code.upper(),
                        'company_name': company_name,
                        'address': receiver_address,
                        'country': receiver_country,
                        'email': receiver_email,
                        'phone': receiver_phone,
                        'business_no': receiver_business_no,
                        'currency': 'KRW' if is_domestic else 'USD',
                        'trade_terms': 'FOB'
                    })
                    st.success(f"저장됨: {customer_code}")
                    st.rerun()
        else:
            # 저장된 고객사 로드
            code = selected_receiver.split(' - ')[0]
            saved_receiver = data_store.get_receiver_by_code(code)
            if saved_receiver:
                customer_code = saved_receiver['customer_code']
                company_name = saved_receiver['company_name']
                receiver_address = saved_receiver['address']
                receiver_country = saved_receiver.get('country', '')
                receiver_email = saved_receiver.get('email', '')
                receiver_phone = saved_receiver.get('phone', '')
                receiver_business_no = saved_receiver.get('business_no', '')

                st.text_input("회사명", value=company_name, disabled=True)
                st.text_input("주소", value=receiver_address, disabled=True)

                receiver_info = ReceiverInfo(
                    company_name=company_name,
                    address=receiver_address,
                    country=receiver_country,
                    email=receiver_email,
                    phone=receiver_phone,
                    business_no=receiver_business_no
                )
            else:
                customer_code = ""
                receiver_info = None

        st.divider()

        # ===== INVOICE INFO =====
        st.subheader("Invoice Info")
        invoice_date = st.date_input("발행일 *", value=datetime.now())
        order_no = st.text_input("주문번호 (PO)", placeholder="PO25001234")
        currency = st.selectbox("통화 *", ["KRW", "USD", "EUR"], index=0 if is_domestic else 1)

        # SHIPPING TERMS (해외만)
        shipping = None
        if not is_domestic:
            with st.expander("Shipping Terms (선택)"):
                trade_terms = st.selectbox("Trade Terms", ["", "FOB", "CIF", "EXW", "DDP"])
                loading_port = st.text_input("Loading Port", placeholder="KOREA")
                destination_port = st.text_input("Destination Port")
                shipping_method = st.selectbox("Shipping Method", ["", "BY AIR", "BY SEA"])
                reason_of_export = st.selectbox("Reason of Export", ["", "SALE", "SAMPLE", "FOC"])

                if trade_terms:
                    shipping = ShippingTerms(
                        terms=trade_terms,
                        loading_port=loading_port,
                        destination_port=destination_port,
                        shipping_method=shipping_method,
                        reason_of_export=reason_of_export
                    )

        # TAX & REMARKS
        with st.expander("Tax & Remarks (선택)"):
            tax_rate = st.number_input("세율 (%)", min_value=0.0, max_value=30.0, value=0.0, step=1.0) / 100
            total_transaction = st.number_input("실거래금액 (FOC용)", min_value=0.0, value=0.0)
            custom_remarks = st.text_area("추가 Remarks", height=80)

    # ========== MAIN CONTENT ==========
    # 탭: 수동 입력 / PO 파일 업로드
    tab_manual, tab_po = st.tabs(["수동 입력", "PO 파일 업로드"])

    # ===== TAB 1: 수동 입력 =====
    with tab_manual:
        col_left, col_right = st.columns([2, 3])

        # ===== 상품 추가 =====
        with col_left:
            st.subheader("상품 추가")

            search_query = st.text_input(
                "SKU / 바코드 / 상품명 검색",
                placeholder="바이오댄스, BIO-S178...",
                key="sku_search"
            )

            if search_query:
                try:
                    results = sku_master.search(search_query, limit=10)
                    if results:
                        # 검색 결과 표시
                        product_options = {
                            f"{p['sku_id']} | {p['description'][:40]}...": p
                            for p in results if p['sku_id']
                        }

                        if product_options:
                            selected_product_key = st.selectbox(
                                "상품 선택",
                                options=list(product_options.keys())
                            )
                            selected_product = product_options[selected_product_key]

                            # 상품 정보 표시
                            st.info(f"""
**SKU ID:** {selected_product['sku_id']}
**Barcode:** {selected_product['barcode']}
**Description:** {selected_product['description']}
""")

                            # 수량 및 가격 입력
                            col_qty, col_price = st.columns(2)
                            with col_qty:
                                qty = st.number_input("수량 (EA) *", min_value=1, value=100, step=10)
                                qty_outbox = st.number_input("QTY(OUTBOX)", min_value=0.0, value=0.0, step=1.0)

                            with col_price:
                                is_foc = st.checkbox("F.O.C. (무상)")
                                # F.O.C.도 가격 입력 가능 (기본값 0)
                                unit_price = st.number_input(
                                    f"단가 ({currency})" + (" (기본 0)" if is_foc else " *"),
                                    min_value=0.0,
                                    value=0.0,
                                    step=0.01 if currency == 'USD' else 100.0,
                                    help="F.O.C.도 0.1불 등 일부 금액 입력 가능" if is_foc else None
                                )

                            hs_code = st.text_input("HS Code", placeholder="3304.99")

                            # 추가 버튼
                            if st.button("➕ 상품 추가", type="primary"):
                                if qty > 0 and (is_foc or unit_price > 0):
                                    item = LineItem(
                                        sku_id=selected_product['sku_id'],
                                        barcode=selected_product['barcode'],
                                        description=selected_product['description'],
                                        hs_code=hs_code,
                                        qty=qty,
                                        qty_outbox=qty_outbox,
                                        unit_price=unit_price,
                                        is_foc=is_foc
                                    )
                                    st.session_state.invoice_items.append(item)
                                    st.success(f"추가됨: {qty} x {selected_product['sku_id']}")
                                    st.rerun()
                                else:
                                    st.error("수량과 단가를 확인하세요")
                        else:
                            st.warning("SKU ID가 있는 상품이 없습니다")
                    else:
                        st.warning("검색 결과가 없습니다")
                except FileNotFoundError:
                    st.error("SKU Master CSV 파일을 찾을 수 없습니다")

            # 수동 입력 옵션
            with st.expander("직접 입력 (SKU Master에 없는 경우)"):
                manual_sku = st.text_input("SKU ID *", key="manual_sku")
                manual_barcode = st.text_input("Barcode", key="manual_barcode")
                manual_desc = st.text_input("상품명 *", key="manual_desc")
                manual_hs = st.text_input("HS Code", key="manual_hs")

                col_m1, col_m2 = st.columns(2)
                with col_m1:
                    manual_qty = st.number_input("수량 *", min_value=1, value=1, key="manual_qty")
                    manual_outbox = st.number_input("QTY(OUTBOX)", min_value=0.0, value=0.0, key="manual_outbox")
                with col_m2:
                    manual_foc = st.checkbox("F.O.C.", key="manual_foc")
                    manual_price = st.number_input(
                        f"단가 ({currency})" + (" (기본 0)" if manual_foc else ""),
                        min_value=0.0,
                        value=0.0,
                        key="manual_price",
                        help="F.O.C.도 일부 금액 입력 가능" if manual_foc else None
                    )

                if st.button("➕ 수동 추가"):
                    if manual_sku and manual_desc and manual_qty > 0:
                        item = LineItem(
                            sku_id=manual_sku,
                            barcode=manual_barcode,
                            description=manual_desc,
                            hs_code=manual_hs,
                            qty=manual_qty,
                            qty_outbox=manual_outbox,
                            unit_price=manual_price,  # F.O.C.도 입력값 그대로 사용
                            is_foc=manual_foc
                        )
                        st.session_state.invoice_items.append(item)
                        st.success(f"추가됨: {manual_qty} x {manual_sku}")
                        st.rerun()
                    else:
                        st.error("SKU ID, 상품명, 수량은 필수입니다")

        # ===== 상품 목록 (수동 입력 탭 내) =====
        with col_right:
            render_invoice_items(currency, staff_email, receiver_info, data_store, invoice_date,
                                 customer_code, order_no, staff_phone, is_domestic, shipping,
                                 tax_rate, total_transaction, custom_remarks, key_prefix="manual")

    # ===== TAB 2: PO 파일 업로드 =====
    with tab_po:
        st.subheader("PO (발주서) 파일 업로드")
        st.caption("Excel 파일(.xlsx, .xls)을 업로드하면 자동으로 SKU와 수량을 파싱합니다")

        uploaded_file = st.file_uploader(
            "PO 파일 선택",
            type=['xlsx', 'xls'],
            key="po_uploader"
        )

        if uploaded_file:
            st.info(f"업로드된 파일: {uploaded_file.name}")

            # 파싱 실행
            with st.spinner("PO 파일 파싱 중..."):
                items, messages = parse_po_file(uploaded_file, uploaded_file.name)

            # 메시지 표시
            for msg in messages:
                if msg.startswith("❌"):
                    st.error(msg)
                elif msg.startswith("⚠️"):
                    st.warning(msg)
                else:
                    st.caption(msg)

            if items:
                # 파싱 결과 미리보기
                st.subheader(f"파싱 결과: {len(items)}개 상품")

                preview_data = []
                for item in items:
                    source_icon = "✅" if item['source'] == 'master' else "⚠️"
                    preview_data.append({
                        '': source_icon,
                        'SKU ID': item['sku_id'],
                        'Description': item['description'][:40] + '...' if len(item['description']) > 40 else item['description'],
                        'QTY': item['qty'],
                        'Barcode': item['barcode']
                    })

                st.dataframe(preview_data, use_container_width=True, hide_index=True)
                st.caption("✅ = SKU Master 매칭, ⚠️ = PO에서만 (Master에 없음)")

                # 단가 일괄 입력
                st.divider()
                col_price_input, col_foc = st.columns([3, 1])
                with col_price_input:
                    default_price = st.number_input(
                        f"기본 단가 ({currency})",
                        min_value=0.0,
                        value=0.0,
                        step=0.01 if currency == 'USD' else 100.0,
                        key="po_default_price"
                    )
                with col_foc:
                    all_foc = st.checkbox("전체 F.O.C.", key="po_all_foc")

                # 전체 추가 버튼
                if st.button("➕ 전체 상품 추가", type="primary", key="po_add_all"):
                    added_count = 0
                    for item in items:
                        line_item = LineItem(
                            sku_id=item['sku_id'],
                            barcode=item['barcode'],
                            description=item['description'],
                            hs_code=item.get('hs_code', ''),
                            qty=item['qty'],
                            qty_outbox=0.0,
                            unit_price=0.0 if all_foc else default_price,
                            is_foc=all_foc
                        )
                        st.session_state.invoice_items.append(line_item)
                        added_count += 1

                    st.success(f"✅ {added_count}개 상품이 추가되었습니다")
                    st.rerun()

        # PO 탭에서도 현재 아이템 목록 표시
        st.divider()
        st.subheader("현재 Invoice Items")
        render_invoice_items(currency, staff_email, receiver_info, data_store, invoice_date,
                             customer_code, order_no, staff_phone, is_domestic, shipping,
                             tax_rate, total_transaction, custom_remarks, key_prefix="po")

    # ===== FOOTER =====
    st.divider()
    st.caption("CI Generator v2.0 | Beautyselection")


if __name__ == "__main__":
    main()
