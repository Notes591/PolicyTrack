# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import time
import requests
import xml.etree.ElementTree as ET
import re
from streamlit_autorefresh import st_autorefresh
import gspread.exceptions

# ====== تحديث تلقائي كل 10 دقائق ======
st_autorefresh(interval=600000, key="auto_refresh")

# ====== إعداد الاتصال بجوجل شيت ======
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# ====== اسم ملف Google Sheet ======
SHEET_NAME = "Complaints"  # نفس الملف
POLICY_SHEET = "Policy number"

# ====== الوصول إلى ورقة Policy number ======
try:
    policy_sheet = client.open(SHEET_NAME).worksheet(POLICY_SHEET)
except Exception as e:
    st.error(f"❌ خطأ في الوصول إلى الورقة: {e}")
    st.stop()

# ====== إعداد صفحة Streamlit ======
st.set_page_config(page_title="📦 تتبع الشحنات", page_icon="🚚", layout="wide")
st.title("🚚 نظام تتبع الشحنات (Policy number)")

# ====== بيانات Aramex ======
client_info = {
    "UserName": "fitnessworld525@gmail.com",
    "Password": "Aa12345678@",
    "Version": "v1",
    "AccountNumber": "71958996",
    "AccountPin": "657448",
    "AccountEntity": "RUH",
    "AccountCountryCode": "SA"
}

# ====== دوال مساعدة ======
def remove_xml_namespaces(xml_str):
    xml_str = re.sub(r'xmlns(:\w+)?="[^"]+"', '', xml_str)
    xml_str = re.sub(r'(<\/?)(\w+:)', r'\1', xml_str)
    return xml_str

def extract_reference(tracking_result):
    for ref_tag in ['Reference1', 'Reference2', 'Reference3', 'Reference4', 'Reference5']:
        ref_elem = tracking_result.find(ref_tag)
        if ref_elem is not None and ref_elem.text and ref_elem.text.strip() != "":
            return ref_elem.text.strip()
    return ""

def get_aramex_status(awb_number):
    try:
        headers = {"Content-Type": "application/json"}
        payload = {
            "ClientInfo": client_info,
            "Shipments": [awb_number],
            "Transaction": {"Reference1": "", "Reference2": "", "Reference3": "", "Reference4": "", "Reference5": ""},
            "LabelInfo": None
        }
        url = "https://ws.aramex.net/ShippingAPI.V2/Tracking/Service_1_0.svc/json/TrackShipments"
        response = requests.post(url, json=payload, headers=headers, timeout=10)
        if response.status_code != 200:
            return f"❌ فشل الاتصال ({response.status_code})"
        xml_content = response.content.decode('utf-8')
        xml_content = remove_xml_namespaces(xml_content)
        root = ET.fromstring(xml_content)
        tracking_results = root.find('TrackingResults')
        if tracking_results is None or len(tracking_results) == 0:
            return "❌ لا توجد حالة متاحة"
        keyvalue = tracking_results.find('KeyValueOfstringArrayOfTrackingResultmFAkxlpY')
        if keyvalue is not None:
            tracking_array = keyvalue.find('Value')
            if tracking_array is not None:
                tracks = tracking_array.findall('TrackingResult')
                if tracks:
                    last_track = sorted(tracks, key=lambda tr: tr.find('UpdateDateTime').text if tr.find('UpdateDateTime') is not None else '', reverse=True)[0]
                    desc = last_track.find('UpdateDescription').text if last_track.find('UpdateDescription') is not None else "—"
                    return desc
        return "❌ لا توجد حالة متاحة"
    except Exception as e:
        return f"⚠️ خطأ في جلب الحالة: {e}"

# ====== تحميل بيانات الشيت ======
try:
    policy_data = policy_sheet.get_all_values()
except Exception:
    policy_data = []

# ====== البحث ======
st.header("🔍 البحث عن شحنة")
search_order = st.text_input("أدخل رقم الطلب للبحث")

if search_order.strip():
    found = False
    for i, row in enumerate(policy_data[1:], start=2):  # تخطي العناوين
        if len(row) >= 2 and str(row[0]) == search_order:
            found = True
            policy_number = row[1]
            date_added = row[2] if len(row) > 2 else "—"
            status = row[3] if len(row) > 3 else "—"

            st.success(f"✅ تم العثور على الطلب رقم: {search_order}")
            st.info(f"📦 رقم الشحنة: {policy_number}")
            st.write(f"📅 التاريخ: {date_added}")
            st.write(f"🔄 الحالة الحالية: {status}")

            # تحديث مباشر للحالة
            if policy_number.strip():
                new_status = get_aramex_status(policy_number)
                if new_status and new_status != status:
                    try:
                        policy_sheet.update_cell(i, 4, new_status)
                        st.success(f"✅ تم تحديث الحالة إلى: {new_status}")
                    except Exception as e:
                        st.error(f"⚠️ لم يتم تحديث الحالة: {e}")
            break

    if not found:
        st.error("⚠️ لم يتم العثور على الطلب في الشيت")

# ====== تحديث تلقائي لكل الحالات ======
st.markdown("---")
st.header("🔄 تحديث تلقائي لكل الشحنات")

def update_status_sheets():
    """توزيع الشحنات على تبويبات حسب الحالة"""
    for idx, row in enumerate(policy_data[1:], start=2):
        if len(row) >= 4:
            policy_number = row[1]
            status = row[3]  # العمود الرابع = الحالة
            if not status.strip():
                continue
            try:
                # تحقق إذا كان هناك ورقة باسم الحالة، إذا لا توجد أنشئها
                try:
                    status_sheet = client.open(SHEET_NAME).worksheet(status)
                except gspread.exceptions.WorksheetNotFound:
                    status_sheet = client.open(SHEET_NAME).add_worksheet(title=status, rows="100", cols="10")
                    status_sheet.append_row(["Order Number", "Policy Number", "Date Added", "Status"])
                
                # أضف الشحنة في ورقة الحالة إذا لم تكن موجودة مسبقًا
                existing_orders = status_sheet.col_values(1)  # عمود رقم الطلب
                if row[0] not in existing_orders:
                    status_sheet.append_row(row[:4])
            except Exception as e:
                st.error(f"⚠️ خطأ أثناء تحديث تبويب الحالة {status}: {e}")

if st.button("تحديث جميع الحالات الآن"):
    if len(policy_data) <= 1:
        st.warning("❌ لا توجد بيانات لتحديثها")
    else:
        progress = st.progress(0)
        updated_count = 0
        for idx, row in enumerate(policy_data[1:], start=2):
            if len(row) >= 2:
                policy_number = row[1]
                if policy_number.strip():
                    status = get_aramex_status(policy_number)
                    try:
                        policy_sheet.update_cell(idx, 4, status)
                        updated_count += 1
                        policy_data[idx-1][3] = status  # تحديث البيانات محليًا
                    except gspread.exceptions.APIError:
                        time.sleep(1)
            progress.progress(idx / len(policy_data))
        st.success(f"✅ تم تحديث {updated_count} شحنة بنجاح")
        # تحديث تبويبات الحالات
        update_status_sheets()

# ====== عرض كل البيانات ======
st.markdown("---")
st.header("📋 جميع الشحنات المسجلة")
if len(policy_data) > 1:
    st.dataframe(policy_data[1:], use_container_width=True)
else:
    st.info("لا توجد بيانات في الشيت حالياً.")
