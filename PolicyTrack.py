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
from email.mime.application import MIMEApplication
import io

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

# ================= إزالة XML Namespace =================
def remove_xml_namespaces(xml_str):
    xml_str = re.sub(r'xmlns(:\w+)?="[^"]+"', '', xml_str)
    xml_str = re.sub(r'(<\/?)(\w+:)', r'\1', xml_str)
    return xml_str

# ================= إعداد واجهة الإيميل =================
st.markdown("### 📧 إعدادات الإيميل (اختياري)")

custom_subject = st.text_input(
    "عنوان الإيميل (اختياري)",
    value="🚨 تنبيه: شحنات متأخرة (Noon – Aramex)"
)

custom_emails_input = st.text_input(
    "إيميلات المستلمين (افصلهم بفاصلة , إذا تريد تغييرهم — اختيارية)",
    ""
)

# ================= إرسال الإيميل =================
def send_delay_email(delayed_rows, custom_emails=None, custom_subject=None):
    if not delayed_rows:
        return

    try:
        email_user = st.secrets["email"]["username"]
        email_pass = st.secrets["email"]["password"]
        default_emails = st.secrets["email"]["send_to"]
    except:
        st.error("❌ لم يتم تحميل بيانات الإيميل من secrets.toml")
        return

    # اختيار الإيميلات
    if custom_emails:
        send_to = [e.strip() for e in custom_emails.split(",") if e.strip()]
    else:
        send_to = default_emails

    # عنوان الإيميل
    subject = custom_subject if custom_subject else "🚨 تنبيه: شحنات متأخرة (Noon – Aramex)"

    # نص الرسالة
    message = "يوجد شحنات متأخرة تجاوزت 3 أيام:\n\n"
    for row in delayed_rows:
        message += f"- Order: {row[0]} | Policy: {row[1]} | Days: {row[4]}\n"

    msg = MIMEMultipart()
    msg["From"] = email_user
    msg["To"] = ", ".join(send_to)
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "plain"))

    # إرفاق Excel
    df = pd.DataFrame(delayed_rows, columns=["Order", "Policy", "Date", "Status", "Days", "Shipping State"])
    output = io.BytesIO()
    df.to_excel(output, index=False, sheet_name="Delayed Shipments")
    output.seek(0)

    part = MIMEApplication(output.read(), Name="Delayed_Shipments.xlsx")
    part['Content-Disposition'] = 'attachment; filename="Delayed_Shipments.xlsx"'
    msg.attach(part)

    # إرسال
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
            "Transaction": {},
            "LabelInfo": None
        }
        url = "https://ws.aramex.net/ShippingAPI.V2/Tracking/Service_1_0.svc/json/TrackShipments"
        response = requests.post(url, json=payload, headers=headers, timeout=10)

        if response.status_code != 200:
            return "❌ فشل الاتصال"

        xml_content = remove_xml_namespaces(response.content.decode('utf-8'))
        root = ET.fromstring(xml_content)

        tracks = root.findall(".//TrackingResult")
        if not tracks:
            return "❌ لا توجد حالة"

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

    except:
        return "⚠️ خطأ أثناء جلب الحالة"

# ================= تحميل بيانات الشيت =================
policy_data = policy_sheet.get_all_values()

# ================= تحديث الأيام =================
cells = policy_sheet.range(f'E2:E{len(policy_data)}')

for idx, row in enumerate(policy_data[1:]):
    if len(row) < 6:
        row += ["0", "غير معروف"] * (6 - len(row))

    date_added = row[2].strip()
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

policy_sheet.update_cells(cells)

# ================= زر تحديث الحالات =================
st.markdown("---")
st.subheader("🔄 تحديث الحالات الآن")

if st.button("تحديث جميع الحالات الآن"):

    updated_rows = []
    for row in policy_data[1:]:
        policy = row[1]
        status = get_aramex_status(policy)
        row[3] = status
        updated_rows.append(row)

    policy_sheet.update("A2", updated_rows)

    # استخراج الشحنات المتأخرة
    delayed_shipments = [
        r for r in updated_rows
        if int(r[4]) > 3 and "delivered" not in r[3].lower() and "returned" not in r[3].lower()
    ]

    # تجهيز الأعمدة
    for r in delayed_shipments:
        if len(r) < 6:
            r += ["—"] * (6 - len(r))

    # إرسال الإيميل
    send_delay_email(
        delayed_shipments,
        custom_emails=custom_emails_input,
        custom_subject=custom_subject
    )

    st.success("✔️ تم تحديث الحالات وإرسال الإيميل.")

# ================= عرض الجداول =================
def check_status(text):
    text = text.lower()
    if "delivered" in text or "تم التسليم" in text:
        return "delivered"
    if "returned" in text or "تم الإرجاع" in text:
        return "returned"
    return "other"

delayed_shipments = [
    row for row in policy_data[1:]
    if int(row[4]) > 3 and check_status(row[3]) == "other"
]

st.markdown("---")
st.subheader("🚨 الشحنات المتأخرة")

if delayed_shipments:
    df = pd.DataFrame(delayed_shipments, columns=["Order", "Policy", "Date", "Status", "Days", "Shipping State"])
    st.dataframe(df, use_container_width=True)
else:
    st.info("لا توجد شحنات متأخرة.")

st.success("🚀 التطبيق جاهز ويعمل بكل الوظائف!")
