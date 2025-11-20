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
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ================= إعداد صفحة Streamlit =================
st.set_page_config(page_title="📦 تتبع الشحنات", page_icon="🚚", layout="wide")
st.title("🚚 نظام تتبع الشحنات (Policy number)")

# تحديث تلقائي كل 10 دقائق
st_autorefresh(interval=600000, key="auto_refresh")

# ================= إعداد الاتصال بجوجل شيت =================
scope = ["https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive"]

creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# ================= الشيتات =================
SHEET_NAME = "Complaints"
POLICY_SHEET = "Policy number"
DELIVERED_SHEET = "تم التسليم"
RETURNED_SHEET = "تم الإرجاع"
ORDERS_SHEET = "Order Number"
DELIVERED_ARCHIVE = "Delivered Archive"
RETURNED_ARCHIVE = "Returned Archive"

def get_or_create_sheet(sheet_name):
    try:
        return client.open(SHEET_NAME).worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = client.open(SHEET_NAME).add_worksheet(title=sheet_name, rows="100", cols="10")
        sheet.append_row(["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment"])
        return sheet

policy_sheet = client.open(SHEET_NAME).worksheet(POLICY_SHEET)
delivered_sheet = get_or_create_sheet(DELIVERED_SHEET)
returned_sheet = get_or_create_sheet(RETURNED_SHEET)
delivered_archive_sheet = get_or_create_sheet(DELIVERED_ARCHIVE)
returned_archive_sheet = get_or_create_sheet(RETURNED_ARCHIVE)

# ================= شيت الاوردر =================
order_sheet = client.open(SHEET_NAME).worksheet(ORDERS_SHEET)
order_data = order_sheet.get_all_values()
order_dict = {row[1]: row[3] for row in order_data[1:] if len(row) > 3 and row[3].strip()}

# ================= بيانات Aramex =================
client_info = {
    "UserName": "fitnessworld525@gmail.com",
    "Password": "Aa12345678@",
    "Version": "v1",
    "AccountNumber": "71958996",
    "AccountPin": "657448",
    "AccountEntity": "RUH",
    "AccountCountryCode": "SA"
}

# ================= دوال =================
def remove_xml_namespaces(xml_str):
    xml_str = re.sub(r'xmlns(:\w+)?="[^"]+"', '', xml_str)
    xml_str = re.sub(r'(<\/?)(\w+:)', r'\1', xml_str)
    return xml_str

# ================= إرسال الإيميل (من secrets.toml) =================
def send_delay_email(delayed_rows):
    if not delayed_rows:
        return

    try:
        email_user = st.secrets["email"]["username"]
        email_pass = st.secrets["email"]["password"]
        send_to = st.secrets["email"]["send_to"]

    except Exception as e:
        st.error(f"❌ لم يتم تحميل بيانات الإيميل من secrets.toml: {e}")
        return

    subject = "🚨 تنبيه: شحنات متأخرة (Noon – Aramex)"

    message = "يوجد شحنات متأخرة تجاوزت 3 أيام:\n\n"
    for row in delayed_rows:
        order = row[0]
        policy = row[1]
        days = row[4]
        message += f"- Order: {order} | Policy: {policy} | Days: {days}\n"

    msg = MIMEMultipart()
    msg["From"] = email_user
    msg["To"] = ", ".join(send_to)
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "plain"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email_user, email_pass)
        server.sendmail(email_user, send_to, msg.as_string())
        server.quit()

    except Exception as e:
        st.error(f"❌ فشل إرسال الإيميل: {e}")


# ================= جلب حالة أرامكس =================
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

        xml_content = remove_xml_namespaces(response.content.decode('utf-8'))
        root = ET.fromstring(xml_content)
        tracking_results = root.find('TrackingResults')

        if tracking_results is None:
            return "❌ لا توجد حالة متاحة"

        keyvalue = tracking_results.find('KeyValueOfstringArrayOfTrackingResultmFAkxlpY')
        if keyvalue is None:
            return "❌ لا توجد حالة متاحة"

        tracking_array = keyvalue.find('Value')
        if tracking_array is None:
            return "❌ لا توجد حالة متاحة"

        tracks = tracking_array.findall('TrackingResult')
        if not tracks:
            return "❌ لا توجد حالة متاحة"

        last_track = sorted(
            tracks,
            key=lambda tr: tr.find('UpdateDateTime').text if tr.find('UpdateDateTime') is not None else '',
            reverse=True
        )[0]

        desc_en = last_track.find('UpdateDescription').text or "—"

        try:
            desc_ar = GoogleTranslator(source='en', target='ar').translate(desc_en)
        except:
            desc_ar = "—"

        return f"{desc_en} - {desc_ar}"

    except Exception as e:
        return f"⚠️ خطأ: {e}"

# ================= تحميل بيانات الشيت =================
policy_data = policy_sheet.get_all_values()

# ================= تحديث الأيام والحالة =================
cells = policy_sheet.range(f'E2:E{len(policy_data)}')

for idx, row in enumerate(policy_data[1:]):
    if len(row) < 6:
        row += ["0", "غير معروف"] * (6 - len(row))

    date_added = row[2].strip() if len(row) > 2 else None
    days_diff = 0

    if date_added:
        for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
            try:
                dt = datetime.strptime(date_added, fmt)
                days_diff = (datetime.now() - dt).days
                break
            except:
                continue

    row[4] = days_diff
    cells[idx].value = days_diff

    order_num = str(row[0])
    row[5] = "مشحون" if order_num in order_dict else "غير مشحون"

policy_sheet.update_cells(cells)

# ================= البحث =================
st.header("🔍 البحث عن شحنة")
search_order = st.text_input("أدخل رقم الطلب للبحث")

if search_order.strip():
    found = False
    for row in policy_data[1:]:
        if len(row) >= 2 and str(row[0]) == search_order:
            found = True
            st.success(f"تم العثور على الطلب {search_order}")
            st.info(f"📦 رقم الشحنة: {row[1]}")
            st.write(f"📅 التاريخ: {row[2]}")
            st.write(f"🔄 الحالة: {row[3]}")
            st.write(f"⏳ الأيام: {row[4]}")
            break

    if not found:
        st.error("⚠️ لم يتم العثور على الطلب")

# ================= دالة التصنيف =================
def check_status(status_text):
    text = status_text.lower()

    if any(w in text for w in ["delivered", "تم التسليم", "collected"]):
        return "delivered"

    if any(w in text for w in ["returned", "تم الإرجاع"]):
        return "returned"

    return "other"

# ================= تأخير =================
delayed_shipments = [
    row for row in policy_data[1:]
    if int(row[4]) > 3 and check_status(row[3]) == "other"
]

current_shipments = [
    row for row in policy_data[1:]
    if int(row[4]) <= 3 and check_status(row[3]) == "other"
]

def normalize_rows(rows, n):
    fixed = []
    for r in rows:
        r = r[:n]
        r += ["—"] * (n - len(r))
        fixed.append(r)
    return fixed

delayed_shipments = normalize_rows(delayed_shipments, 6)
current_shipments = normalize_rows(current_shipments, 6)

# ========== إرسال الإيميل للتأخيرات ==========
send_delay_email(delayed_shipments)

# ================= عرض الجداول =================
st.markdown("---")
st.subheader("🚨 الشحنات المتأخرة")

if delayed_shipments:
    st.dataframe(pd.DataFrame(
        delayed_shipments,
        columns=["Order", "Policy", "Date", "Status", "Days", "حالة الشحن"]
    ), use_container_width=True)
else:
    st.info("لا توجد شحنات متأخرة.")

st.markdown("---")
st.subheader("📦 الشحنات الجارية")

if current_shipments:
    st.dataframe(pd.DataFrame(
        current_shipments,
        columns=["Order", "Policy", "Date", "Status", "Days", "حالة الشحن"]
    ), use_container_width=True)
else:
    st.info("لا توجد شحنات جارية.")

st.success("🚀 التطبيق يعمل الآن بكل الوظائف بما فيها الإرسال التلقائي!")
