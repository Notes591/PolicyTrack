# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import requests
import xml.etree.ElementTree as ET
import re
from streamlit_autorefresh import st_autorefresh
import pandas as pd
import time
from deep_translator import GoogleTranslator

# ====== إعداد الصفحة ======
st.set_page_config(page_title="تتبع الشحنات", page_icon="🚚", layout="wide")

# ====== CSS مخصص ======
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .stDataFrame { border-radius: 10px; }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
    }
    h1 { color: #1a1a2e; }
    .stButton>button {
        background-color: #4361ee;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 10px 24px;
        font-weight: bold;
    }
    .stButton>button:hover { background-color: #3a0ca3; }
</style>
""", unsafe_allow_html=True)

st.title("🚚 نظام تتبع الشحنات")
st_autorefresh(interval=600000, key="auto_refresh")

# ====== الاتصال بـ Google Sheets ======
@st.cache_resource
def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        st.secrets["gcp_service_account"], scope)
    return gspread.authorize(creds)

client = get_gspread_client()

SHEET_NAME     = "Complaints"
POLICY_SHEET   = "Policy number"
DELIVERED_SHEET = "تم التسليم"
RETURNED_SHEET  = "تم الإرجاع"
ORDERS_SHEET    = "Order Number"
DELIVERED_ARCHIVE = "Delivered Archive"
RETURNED_ARCHIVE  = "Returned Archive"

def get_or_create_sheet(name):
    try:
        return client.open(SHEET_NAME).worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = client.open(SHEET_NAME).add_worksheet(title=name, rows="100", cols="10")
        ws.append_row(["Order Number","Policy Number","Date","Status","Days Since Shipment"])
        return ws

policy_sheet          = client.open(SHEET_NAME).worksheet(POLICY_SHEET)
delivered_sheet       = get_or_create_sheet(DELIVERED_SHEET)
returned_sheet        = get_or_create_sheet(RETURNED_SHEET)
delivered_archive_sheet = get_or_create_sheet(DELIVERED_ARCHIVE)
returned_archive_sheet  = get_or_create_sheet(RETURNED_ARCHIVE)

order_sheet = client.open(SHEET_NAME).worksheet(ORDERS_SHEET)
order_data  = order_sheet.get_all_values()
order_dict  = {row[1]: row[3] for row in order_data[1:]
               if len(row) > 3 and row[3].strip()}

# ====== Aramex ======
CLIENT_INFO = {
    "UserName": "fitnessworld525@gmail.com",
    "Password": "Aa12345678@",
    "Version": "v1",
    "AccountNumber": "71958996",
    "AccountPin": "657448",
    "AccountEntity": "RUH",
    "AccountCountryCode": "SA"
}

def remove_xml_namespaces(xml_str):
    xml_str = re.sub(r'xmlns(:\w+)?="[^"]+"', '', xml_str)
    xml_str = re.sub(r'(<\/?)(\w+:)', r'\1', xml_str)
    return xml_str

@st.cache_data(ttl=300)
def get_aramex_status(awb_number):
    try:
        payload = {
            "ClientInfo": CLIENT_INFO,
            "Shipments": [awb_number],
            "Transaction": {"Reference1":"","Reference2":"","Reference3":"","Reference4":"","Reference5":""},
            "LabelInfo": None
        }
        resp = requests.post(
            "https://ws.aramex.net/ShippingAPI.V2/Tracking/Service_1_0.svc/json/TrackShipments",
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=10
        )
        if resp.status_code != 200:
            return f"❌ فشل الاتصال ({resp.status_code})"

        xml_content = remove_xml_namespaces(resp.content.decode('utf-8'))
        root = ET.fromstring(xml_content)

        tracking_results = root.find('TrackingResults')
        if tracking_results is None or len(tracking_results) == 0:
            return "❌ لا توجد حالة"

        keyvalue = tracking_results.find('KeyValueOfstringArrayOfTrackingResultmFAkxlpY')
        if keyvalue is None:
            return "❌ لا توجد حالة"

        tracking_array = keyvalue.find('Value')
        if tracking_array is None:
            return "❌ لا توجد حالة"

        tracks = tracking_array.findall('TrackingResult')
        if not tracks:
            return "❌ لا توجد حالة"

        last = sorted(
            tracks,
            key=lambda t: t.find('UpdateDateTime').text if t.find('UpdateDateTime') is not None else '',
            reverse=True
        )[0]

        desc_en = last.find('UpdateDescription').text if last.find('UpdateDescription') is not None else "—"
        try:
            desc_ar = GoogleTranslator(source='en', target='ar').translate(desc_en)
        except:
            desc_ar = "—"
        return f"{desc_en} | {desc_ar}"

    except Exception as e:
        return f"⚠️ خطأ: {e}"

# ====== تصنيف الحالة ======
DELIVERED_KEYWORDS = ["delivered","تم التسليم","shipment charges paid",
                      "customer id received","collected by consignee"]
RETURNED_KEYWORDS  = ["returned","تم الإرجاع","returned to shipper"]

def check_status(status_text: str) -> str:
    text = status_text.lower()
    if any(k in text for k in DELIVERED_KEYWORDS): return "delivered"
    if any(k in text for k in RETURNED_KEYWORDS):  return "returned"
    return "other"

# ====== حساب أيام الشحن ======
def calc_days(date_str: str) -> int:
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return (datetime.now() - datetime.strptime(date_str.strip(), fmt)).days
        except:
            continue
    return 0

# ====== تحميل بيانات الشيت ======
def load_policy_data():
    return policy_sheet.get_all_values()

policy_data = load_policy_data()

# ====== تحديث عمود الأيام ======
if len(policy_data) > 1:
    cells = policy_sheet.range(f'E2:E{len(policy_data)}')
    for idx, row in enumerate(policy_data[1:]):
        date_str = row[2] if len(row) > 2 else ""
        days = calc_days(date_str)
        # تحديث في الذاكرة
        while len(row) < 6:
            row.append("—")
        row[4] = days
        row[5] = "مشحون" if str(row[0]) in order_dict else "غير مشحون"
        cells[idx].value = days
    policy_sheet.update_cells(cells)

# ====== واجهة البحث ======
st.header("🔍 البحث عن شحنة")
search_order = st.text_input("أدخل رقم الطلب", placeholder="مثال: 12345")

if search_order.strip():
    result = next((r for r in policy_data[1:]
                   if len(r) >= 2 and str(r[0]) == search_order.strip()), None)
    if result:
        while len(result) < 6:
            result.append("—")
        col1, col2, col3 = st.columns(3)
        col1.success(f"✅ رقم الطلب: {result[0]}")
        col2.info(f"📦 رقم الشحنة: {result[1]}")
        col3.write(f"📅 التاريخ: {result[2]}")
        st.write(f"🔄 **الحالة:** {result[3]}")
        st.write(f"⏳ **أيام منذ الشحن:** {result[4]}")
        st.write(f"🚚 **حالة الشحن:** {result[5]}")
    else:
        st.error("⚠️ لم يتم العثور على الطلب")

# ====== تحديث الحالات ======
st.markdown("---")
if st.button("🔄 تحديث جميع الحالات الآن"):
    progress = st.progress(0)
    status_msg = st.empty()
    rows_to_update = [
        (idx, row) for idx, row in enumerate(policy_data[1:], start=2)
        if len(row) >= 2 and row[1].strip() and check_status(row[3] if len(row) > 3 else "") == "other"
    ]
    for i, (sheet_idx, row) in enumerate(rows_to_update):
        status_msg.info(f"جاري تحديث {i+1}/{len(rows_to_update)}: {row[1]}")
        new_status = get_aramex_status(row[1])
        row[3] = new_status
        progress.progress((i+1) / max(len(rows_to_update), 1))

    if rows_to_update:
        cells = policy_sheet.range(f'D2:D{len(policy_data)}')
        for idx, row in enumerate(policy_data[1:]):
            cells[idx].value = row[3] if len(row) > 3 else "—"
        policy_sheet.update_cells(cells)

    status_msg.empty()
    st.success("✅ تم تحديث جميع الحالات بنجاح!")

# ====== نقل المُسلَّم والمُرجَع ======
def append_in_batches(sheet, rows, batch_size=20):
    for i in range(0, len(rows), batch_size):
        sheet.append_rows(rows[i:i+batch_size], value_input_option='USER_ENTERED')
        time.sleep(1)

delivered_existing = {r[1] for r in delivered_sheet.get_all_values()[1:] if len(r) > 1}
returned_existing  = {r[1] for r in returned_sheet.get_all_values()[1:]  if len(r) > 1}

new_delivered, new_returned, rows_to_delete = [], [], []

for idx, row in enumerate(policy_data[1:], start=2):
    if len(row) < 2 or not row[1].strip():
        continue
    status = check_status(row[3] if len(row) > 3 else "")
    if status == "delivered" and row[1] not in delivered_existing:
        new_delivered.append(row[:5])
        rows_to_delete.append(idx)
    elif status == "returned" and row[1] not in returned_existing:
        new_returned.append(row[:5])
        rows_to_delete.append(idx)

if new_delivered:
    append_in_batches(delivered_sheet, new_delivered)
    append_in_batches(delivered_archive_sheet, new_delivered)

if new_returned:
    append_in_batches(returned_sheet, new_returned)
    append_in_batches(returned_archive_sheet, new_returned)

# حذف الصفوف من الأسفل للأعلى (مهم!)
for row_idx in sorted(rows_to_delete, reverse=True):
    policy_sheet.delete_rows(row_idx)
    time.sleep(0.3)

# ====== عرض الإحصائيات ======
st.markdown("---")
all_active = [r for r in policy_data[1:] if check_status(r[3] if len(r) > 3 else "") == "other"]

def safe_days(r):
    try: return int(r[4])
    except: return 0

delayed  = [r for r in all_active if safe_days(r) > 3]
current  = [r for r in all_active if safe_days(r) <= 3]

col1, col2, col3, col4 = st.columns(4)
col1.metric("📦 إجمالي النشطة",  len(all_active))
col2.metric("⚠️ متأخرة",         len(delayed))
col3.metric("✅ في الوقت",        len(current))
col4.metric("🚚 تم التسليم",      len(new_delivered))

# ====== تطبيع الصفوف ======
def normalize(data, n=6):
    result = []
    for row in data:
        row = list(row[:n])
        row += ["—"] * (n - len(row))
        result.append(row)
    return result

cols = ["رقم الطلب","رقم الشحنة","التاريخ","الحالة","أيام الشحن","حالة الشحن"]

st.markdown("---")
st.subheader("⚠️ الشحنات المتأخرة (أكثر من 3 أيام)")
if delayed:
    st.dataframe(
        pd.DataFrame(normalize(delayed), columns=cols),
        use_container_width=True,
        height=400
    )
else:
    st.success("✅ لا توجد شحنات متأخرة!")

st.markdown("---")
st.subheader("📦 الشحنات الحالية")
if current:
    st.dataframe(
        pd.DataFrame(normalize(current), columns=cols),
        use_container_width=True,
        height=400
    )
else:
    st.info("لا توجد شحنات حالياً.")

st.markdown("---")
st.caption(f"آخر تحديث: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
