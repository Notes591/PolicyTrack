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

# ====== تحديث تلقائي كل 10 دقائق ======
st_autorefresh(interval=600000, key="auto_refresh")

# ====== إعداد الاتصال بجوجل شيت ======
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# ====== شيتات Google ======
SHEET_NAME = "Complaints"
POLICY_SHEET = "Policy number"
DELIVERED_SHEET = "تم التسليم"
ORDERS_SHEET = "Order Number"

# ====== تحميل ورقة Policy number ======
policy_sheet = client.open(SHEET_NAME).worksheet(POLICY_SHEET)

# ====== إنشاء/تحميل تبويب "تم التسليم" ======
try:
    delivered_sheet = client.open(SHEET_NAME).worksheet(DELIVERED_SHEET)
except gspread.exceptions.WorksheetNotFound:
    delivered_sheet = client.open(SHEET_NAME).add_worksheet(title=DELIVERED_SHEET, rows="100", cols="10")
    delivered_sheet.append_row(["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment"])

# ====== تحميل شيت Order Number ======
order_sheet = client.open(SHEET_NAME).worksheet(ORDERS_SHEET)
order_data = order_sheet.get_all_values()
order_dict = {row[0]: row[1] for row in order_data[1:]}  # رقم الطلب : المندوب

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
policy_data = policy_sheet.get_all_values()

# ====== تحديث أيام الشحن وحالة الشحن ======
for idx, row in enumerate(policy_data[1:], start=2):
    if len(row) < 6:
        row += ["0", "غير معروف"] * (6 - len(row))
    # حساب الأيام منذ الشحنة
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
    # حالة الشحن (مشحون / غير مشحون)
    order_num = str(row[0])
    if order_num in order_dict and order_dict[order_num].strip():
        row[5] = "مشحون"
    else:
        row[5] = "غير مشحون"

# ====== البحث عن شحنة ======
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
            shipping_state = row[5] if len(row) > 5 else "غير معروف"

            st.success(f"✅ تم العثور على الطلب رقم: {search_order}")
            st.info(f"📦 رقم الشحنة: {policy_number}")
            st.write(f"📅 التاريخ: {date_added}")
            st.write(f"🔄 الحالة الحالية: {status}")
            st.write(f"⏳ أيام منذ الشحن: {days_since}")
            st.write(f"🚚 حالة الشحن: {shipping_state}")
            break
    if not found:
        st.error("⚠️ لم يتم العثور على الطلب في الشيت")

# ====== تحديث جميع الحالات ======
if st.button("تحديث جميع الحالات الآن"):
    progress = st.progress(0)
    for idx, row in enumerate(policy_data[1:], start=2):
        if len(row) >= 2 and row[1].strip():
            if row[3].strip().lower() != "delivered":
                new_status = get_aramex_status(row[1])
                row[3] = new_status
                try:
                    policy_sheet.update_cell(idx, 4, new_status)
                except:
                    pass
        progress.progress(idx / len(policy_data))
    st.success("✅ تم تحديث جميع الحالات")

# ====== تصنيف البيانات لعرضها ======
delayed_shipments = [row for row in policy_data[1:] if int(row[4]) > 3 and row[3].strip().lower() != "delivered"]
current_shipments = [row for row in policy_data[1:] if int(row[4]) <= 3 and row[3].strip().lower() != "delivered"]

# ====== تحديث تبويب "تم التسليم" تلقائياً ======
delivered_shipments = [row for row in delivered_sheet.get_all_values()[1:]]  # من تبويب التسليم
for row in policy_data[1:]:
    if row[3].strip().lower() == "delivered":
        existing = [r[1] for r in delivered_shipments]
        if row[1] not in existing:
            delivered_sheet.append_row(row[:5])
            delivered_shipments.append(row)

# ====== عرض الجداول ======
st.markdown("---")
st.subheader("الشحنات المتأخرة")
if delayed_shipments:
    st.dataframe(pd.DataFrame(delayed_shipments, columns=["Order Number","Policy Number","Date","Status","Days Since Shipment","حالة الشحن"]), use_container_width=True)
else:
    st.info("لا توجد شحنات متأخرة حالياً.")

st.markdown("---")
st.subheader("✅ الشحنات التي تم توصيلها")
if delivered_shipments:
    df_delivered = pd.DataFrame(delivered_shipments, columns=["Order Number","Policy Number","Date","Status","Days Since Shipment"])
    for i, row in df_delivered.iterrows():
        st.write(row.to_dict())
        if st.button(f"حذف {row['Order Number']}"):
            delivered_sheet.delete_rows(i+2)
            st.success(f"✅ تم حذف {row['Order Number']}")
else:
    st.info("لا توجد شحنات تم توصيلها حالياً.")

st.markdown("---")
st.subheader("📦 الشحنات الحالية")
if current_shipments:
    st.dataframe(pd.DataFrame(current_shipments, columns=["Order Number","Policy Number","Date","Status","Days Since Shipment","حالة الشحن"]), use_container_width=True)
else:
    st.info("لا توجد شحنات حالياً.")
