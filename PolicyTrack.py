# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import requests
import xml.etree.ElementTree as ET
import re
from streamlit_autorefresh import st_autorefresh

# ====== تحديث تلقائي كل 10 دقائق ======
st_autorefresh(interval=600000, key="auto_refresh")

# ====== إعداد الاتصال بجوجل شيت ======
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# ====== اسم ملف Google Sheet ======
SHEET_NAME = "Complaints"
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

# ====== تحديث حالات الأيام ======
for idx, row in enumerate(policy_data[1:], start=2):
    if len(row) < 5:
        row += ["0"] * (5 - len(row))
    date_added_str = row[2] if len(row) > 2 else None
    days_diff = 0
    if date_added_str and date_added_str.strip():
        for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
            try:
                date_added = datetime.strptime(date_added_str, fmt)
                days_diff = (datetime.now() - date_added).days
                break
            except:
                continue
    row[4] = days_diff
    try:
        policy_sheet.update_cell(idx, 5, days_diff)
    except:
        pass

# ====== تحديث جميع الحالات من Aramex ======
if st.button("تحديث جميع الحالات الآن"):
    import time
    progress = st.progress(0)
    for idx, row in enumerate(policy_data[1:], start=2):
        if len(row) >= 2 and row[1].strip():
            new_status = get_aramex_status(row[1])
            row[3] = new_status
            try:
                policy_sheet.update_cell(idx, 4, new_status)
            except:
                pass
        progress.progress(idx / len(policy_data))
    st.success("✅ تم تحديث جميع الحالات")

# ====== تصنيف البيانات ======
delayed_shipments = [row for row in policy_data[1:] if int(row[4]) > 3]
delivered_shipments = [row for row in policy_data[1:] if row[3].strip().lower() == "delivered"]
current_shipments = [row for row in policy_data[1:] if int(row[4]) <= 3 and row[3].strip().lower() != "delivered"]

# ====== عرض الجداول ======
st.markdown("---")
st.subheader("⏳ الشحنات المتأخرة (>3 أيام)")
if delayed_shipments:
    st.dataframe(delayed_shipments, use_container_width=True)
else:
    st.info("لا توجد شحنات متأخرة حالياً.")

st.markdown("---")
st.subheader("✅ الشحنات التي تم توصيلها")
if delivered_shipments:
    st.dataframe(delivered_shipments, use_container_width=True)
else:
    st.info("لا توجد شحنات تم توصيلها حالياً.")

st.markdown("---")
st.subheader("📦 الشحنات الحالية")
if current_shipments:
    st.dataframe(current_shipments, use_container_width=True)
else:
    st.info("لا توجد شحنات حالياً.")
