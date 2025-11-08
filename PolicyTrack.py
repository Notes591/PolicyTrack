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

# ====== البحث ======
st.header("🔍 البحث عن شحنة")
search_order = st.text_input("أدخل رقم الطلب للبحث")

if search_order.strip():
    found = False
    for i, row in enumerate(policy_data[1:], start=2):
        if len(row) >= 2 and str(row[0]) == search_order:
            found = True
            policy_number = row[1]
            date_added = row[2] if len(row) > 2 else "—"
            status = row[3] if len(row) > 3 else "—"
            days_since = row[4] if len(row) > 4 else "—"

            st.success(f"✅ تم العثور على الطلب رقم: {search_order}")
            st.info(f"📦 رقم الشحنة: {policy_number}")
            st.write(f"📅 التاريخ: {date_added}")
            st.write(f"🔄 الحالة الحالية: {status}")
            st.write(f"⏳ أيام منذ الشحن: {days_since}")

            if policy_number.strip():
                new_status = get_aramex_status(policy_number)
                if new_status and new_status != status:
                    try:
                        policy_sheet.update_cell(i, 4, new_status)
                        row[3] = new_status
                        st.success(f"✅ تم تحديث الحالة إلى: {new_status}")
                    except Exception as e:
                        st.error(f"⚠️ لم يتم تحديث الحالة: {e}")
            break

    if not found:
        st.error("⚠️ لم يتم العثور على الطلب في الشيت")

# ====== تحديث الشحنات + التبويبات ======
st.markdown("---")
st.header("🔄 تحديث جميع الشحنات")

def update_special_sheets():
    delayed_name = "متأخرة"
    delivered_name = "تم التسليم"

    try:
        # إنشاء/مسح تبويب المتأخرة
        try:
            delayed_sheet = client.open(SHEET_NAME).worksheet(delayed_name)
            delayed_sheet.clear()
            delayed_sheet.append_row(["Order Number", "Policy Number", "Date Added", "Status", "Days Since Shipment"])
        except gspread.exceptions.WorksheetNotFound:
            delayed_sheet = client.open(SHEET_NAME).add_worksheet(title=delayed_name, rows="100", cols="10")
            delayed_sheet.append_row(["Order Number", "Policy Number", "Date Added", "Status", "Days Since Shipment"])
        
        # إنشاء/مسح تبويب التسليم
        try:
            delivered_sheet = client.open(SHEET_NAME).worksheet(delivered_name)
            delivered_sheet.clear()
            delivered_sheet.append_row(["Order Number", "Policy Number", "Date Added", "Status", "Days Since Shipment"])
        except gspread.exceptions.WorksheetNotFound:
            delivered_sheet = client.open(SHEET_NAME).add_worksheet(title=delivered_name, rows="100", cols="10")
            delivered_sheet.append_row(["Order Number", "Policy Number", "Date Added", "Status", "Days Since Shipment"])

        # تحديث البيانات
        for idx, row in enumerate(policy_data[1:], start=2):
            if len(row) < 5:
                row.append(0)  # عمود الأيام
            status = row[3].strip()
            date_added_str = row[2] if len(row) > 2 else None

            # حساب عدد الأيام منذ الشحنة
            if date_added_str and date_added_str.strip():
                try:
                    date_added = datetime.strptime(date_added_str, "%Y-%m-%d")
                    days_diff = (datetime.now() - date_added).days
                    row[4] = days_diff
                    policy_sheet.update_cell(idx, 5, days_diff)
                except:
                    pass

            # الشحنات التي وصلت
            if status.lower() == "delivered":
                delivered_sheet.append_row(row[:5])
                continue

            # الشحنات المتأخرة
            if date_added_str:
                try:
                    if row[4] > 3:
                        delayed_sheet.append_row(row[:5])
                except:
                    continue
    except Exception as e:
        st.error(f"⚠️ خطأ أثناء تحديث التبويبات: {e}")

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
                        row[3] = status
                        updated_count += 1
                    except gspread.exceptions.APIError:
                        time.sleep(1)
            progress.progress(idx / len(policy_data))
        st.success(f"✅ تم تحديث {updated_count} شحنة بنجاح")
        update_special_sheets()

# ====== عرض كل البيانات ======
st.markdown("---")
st.header("📋 جميع الشحنات المسجلة")
if len(policy_data) > 1:
    st.dataframe(policy_data[1:], use_container_width=True)
else:
    st.info("لا توجد بيانات في الشيت حالياً.")
