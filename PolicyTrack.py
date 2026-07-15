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

st_autorefresh(interval=600000, key="auto_refresh")

# ================= إعداد الاتصال بجوجل شيت =================
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

try:
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
except Exception as e:
    st.error("❌ خطأ في تحميل بيانات GCP من secrets.toml: " + str(e))
    st.stop()

# ================= أسماء الشيتات =================
SHEET_NAME        = "Complaints"
POLICY_SHEET      = "Policy number"
DELIVERED_SHEET   = "تم التسليم"
RETURNED_SHEET    = "تم الإرجاع"
ORDERS_SHEET      = "Order Number"
DELIVERED_ARCHIVE = "Delivered Archive"
RETURNED_ARCHIVE  = "Returned Archive"

def get_or_create_sheet(sheet_name):
    try:
        return client.open(SHEET_NAME).worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sh = client.open(SHEET_NAME)
        sheet = sh.add_worksheet(title=sheet_name, rows="100", cols="10")
        sheet.append_row(["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment", "حالة الشحن"])
        return sheet

try:
    policy_sheet = client.open(SHEET_NAME).worksheet(POLICY_SHEET)
except Exception as e:
    st.error(f"❌ لا يمكن فتح الشيت الرئيسي '{SHEET_NAME}': {e}")
    st.stop()

delivered_sheet         = get_or_create_sheet(DELIVERED_SHEET)
returned_sheet          = get_or_create_sheet(RETURNED_SHEET)
delivered_archive_sheet = get_or_create_sheet(DELIVERED_ARCHIVE)
returned_archive_sheet  = get_or_create_sheet(RETURNED_ARCHIVE)

try:
    order_sheet = client.open(SHEET_NAME).worksheet(ORDERS_SHEET)
    order_data  = order_sheet.get_all_values()
    order_dict  = {row[1]: row[3] for row in order_data[1:] if len(row) > 3 and row[3].strip()}
except Exception:
    order_dict = {}

# ================= بيانات حسابَي Aramex =================
ARAMEX_ACCOUNTS = [
    {
        "label": "الحساب الأول",
        "UserName": "fitnessworld525@gmail.com",
        "Password": "Aa12345678@",
        "Version": "v1",
        "AccountNumber": "71958996",
        "AccountPin": "657448",
        "AccountEntity": "RUH",
        "AccountCountryCode": "SA"
    },
    {
        "label": "الحساب الثاني",
        "UserName": "homeentryh5556@gmail.com",
        "Password": "Aa12345678@",
        "Version": "v1",
        "AccountNumber": "4004297",
        "AccountPin": "216216",
        "AccountEntity": "RUH",
        "AccountCountryCode": "SA"
    }
]

# ================= دوال مساعدة =================
def remove_xml_namespaces(xml_str):
    xml_str = re.sub(r'xmlns(:\w+)?="[^"]+"', '', xml_str)
    xml_str = re.sub(r'(<\/?)(\w+:)', r'\1', xml_str)
    return xml_str

def _fetch_aramex_status(awb_number, account):
    """محاولة جلب الحالة من حساب أرامكس واحد. يرجع النص أو None لو فشل."""
    try:
        client_info = {k: v for k, v in account.items() if k != "label"}
        payload = {
            "ClientInfo": client_info,
            "Shipments": [awb_number],
            "Transaction": {"Reference1": "", "Reference2": "", "Reference3": "", "Reference4": "", "Reference5": ""},
            "LabelInfo": None
        }
        url = "https://ws.aramex.net/ShippingAPI.V2/Tracking/Service_1_0.svc/json/TrackShipments"
        response = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)

        if response.status_code != 200:
            return None

        xml_content = remove_xml_namespaces(response.content.decode('utf-8'))
        root = ET.fromstring(xml_content)

        tracking_results = root.find('TrackingResults')
        if tracking_results is not None:
            keyvalue = tracking_results.find('KeyValueOfstringArrayOfTrackingResultmFAkxlpY')
            tracks = []
            if keyvalue is not None:
                tracking_array = keyvalue.find('Value')
                if tracking_array is not None:
                    tracks = tracking_array.findall('TrackingResult')
        else:
            tracks = []

        if not tracks:
            tracks = root.findall(".//TrackingResult")

        if not tracks:
            return None  # لم يجد بيانات، جرّب الحساب التاني

        last_track = sorted(
            tracks,
            key=lambda tr: tr.find('UpdateDateTime').text if tr.find('UpdateDateTime') is not None else '',
            reverse=True
        )[0]

        desc_en = last_track.find('UpdateDescription').text if last_track.find('UpdateDescription') is not None else "—"
        if desc_en == "—" or not desc_en:
            return None

        try:
            desc_ar = GoogleTranslator(source='en', target='ar').translate(desc_en)
        except:
            desc_ar = "—"

        return f"{desc_en} - {desc_ar}"

    except Exception:
        return None


def get_aramex_status(awb_number):
    """يجرب الحسابين بالترتيب. لو الأول رجع نتيجة يستخدمها، لو لا يجرب الثاني."""
    for account in ARAMEX_ACCOUNTS:
        result = _fetch_aramex_status(awb_number, account)
        if result:
            return result
    return "❌ لا توجد حالة متاحة من أي حساب"


def check_status(status_text):
    text = (status_text or "").lower()
    delivered_keywords = ["delivered", "تم التسليم", "shipment charges paid",
                          "customer id received", "collected by consignee", "delivery",
                          "delivered to consignee"]
    returned_keywords  = ["returned", "تم الإرجاع", "returned to shipper",
                          "return to sender", "return"]
    for k in delivered_keywords:
        if k in text:
            return "delivered"
    for k in returned_keywords:
        if k in text:
            return "returned"
    return "other"


def normalize_rows(rows, n=6):
    result = []
    for r in rows:
        r = list(r)[:n]
        r += ["—"] * (n - len(r))
        result.append(r)
    return result


def append_in_batches(sheet, rows, batch_size=50):
    if not rows:
        return
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i+batch_size]
        try:
            sheet.append_rows(batch, value_input_option='USER_ENTERED')
        except Exception:
            for row in batch:
                try:
                    sheet.append_row(row, value_input_option='USER_ENTERED')
                    time.sleep(0.2)
                except Exception:
                    pass
        time.sleep(0.5)


def delete_policy_rows(policy_numbers):
    """حذف صفوف من policy_sheet بناءً على قائمة أرقام البوليصة."""
    for pol in policy_numbers:
        try:
            all_policy = policy_sheet.get_all_values()
            for i, row in enumerate(all_policy[1:], start=2):
                if len(row) > 1 and row[1] == pol:
                    policy_sheet.delete_rows(i)
                    time.sleep(0.3)
                    break
        except Exception:
            continue

# ================= تحميل بيانات policy =================
try:
    policy_data = policy_sheet.get_all_values()
except Exception as e:
    st.error("❌ خطأ في قراءة بيانات الشيت الرئيسي: " + str(e))
    st.stop()

# ================= تحديث عمود الأيام وحالة الشحن =================
if len(policy_data) >= 2:
    cells = policy_sheet.range(f'E2:E{len(policy_data)}')
    for idx, row in enumerate(policy_data[1:]):
        if len(row) < 6:
            row += [""] * (6 - len(row))
        date_str   = row[2] if len(row) > 2 else ""
        days_diff  = 0
        if date_str and date_str.strip():
            for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"):
                try:
                    days_diff = (datetime.now() - datetime.strptime(date_str.strip(), fmt)).days
                    break
                except:
                    continue
        row[4] = days_diff
        cells[idx].value = days_diff
        row[5] = "مشحون" if str(row[0]) in order_dict else "غير مشحون"
    try:
        policy_sheet.update_cells(cells)
    except Exception as e:
        st.warning("تحذير: لم يتم تحديث عمود الأيام: " + str(e))

# ================= واجهة البحث =================
st.header("🔍 البحث عن شحنة")
search_order = st.text_input("أدخل رقم الطلب للبحث")
if search_order.strip():
    found = False
    for row in policy_data[1:]:
        if len(row) >= 2 and str(row[0]) == search_order.strip():
            found = True
            col1, col2, col3 = st.columns(3)
            col1.success(f"✅ رقم الطلب: {row[0]}")
            col2.info(f"📦 رقم الشحنة: {row[1]}")
            col3.write(f"📅 التاريخ: {row[2] if len(row) > 2 else '—'}")
            st.write(f"🔄 الحالة: {row[3] if len(row) > 3 else '—'}")
            st.write(f"⏳ أيام منذ الشحن: {row[4] if len(row) > 4 else '—'}")
            st.write(f"🚚 حالة الشحن: {row[5] if len(row) > 5 else '—'}")
            break
    if not found:
        st.error("⚠️ لم يتم العثور على الطلب في الشيت")

st.markdown("---")

# ================= زر ١: تحديث الحالات فقط =================
st.subheader("🔄 تحديث الحالات")
if st.button("🔄 تحديث جميع الحالات الآن", use_container_width=True):
    progress    = st.progress(0)
    status_msg  = st.empty()
    total       = max(len(policy_data) - 1, 1)
    updated_rows = []

    for idx, row in enumerate(policy_data[1:], start=2):
        if len(row) < 6:
            row += [""] * (6 - len(row))
        if check_status(row[3]) == "other" and row[1].strip():
            status_msg.info(f"جاري تحديث {idx-1}/{total}: {row[1]}")
            row[3] = get_aramex_status(row[1])
        updated_rows.append(row)
        progress.progress((idx - 1) / total)

    # حفظ العمود D
    try:
        cells = policy_sheet.range(f'D2:D{len(policy_data)}')
        for i, r in enumerate(updated_rows):
            cells[i].value = r[3] if len(r) > 3 else "—"
        policy_sheet.update_cells(cells)
    except Exception as e:
        st.warning("تحذير: لم يتم حفظ الحالات: " + str(e))

    # نقل المُسلَّم والمُرجَع
    try:
        delivered_existing = {r[1] for r in delivered_sheet.get_all_values()[1:] if len(r) > 1}
        returned_existing  = {r[1] for r in returned_sheet.get_all_values()[1:]  if len(r) > 1}
    except Exception:
        delivered_existing, returned_existing = set(), set()

    new_delivered, new_returned = [], []
    for r in updated_rows:
        flag = check_status(r[3])
        if flag == "delivered" and r[1] not in delivered_existing:
            new_delivered.append(r[:6])
        elif flag == "returned" and r[1] not in returned_existing:
            new_returned.append(r[:6])

    if new_delivered:
        append_in_batches(delivered_sheet, new_delivered)
        append_in_batches(delivered_archive_sheet, new_delivered)
        delete_policy_rows([r[1] for r in new_delivered])

    if new_returned:
        append_in_batches(returned_sheet, new_returned)
        append_in_batches(returned_archive_sheet, new_returned)
        delete_policy_rows([r[1] for r in new_returned])

    # إعادة تحميل البيانات
    try:
        policy_data = policy_sheet.get_all_values()
    except Exception:
        policy_data = [policy_data[0]] + updated_rows

    status_msg.empty()
    st.success(f"✅ تم تحديث الحالات | نُقل للتسليم: {len(new_delivered)} | نُقل للإرجاع: {len(new_returned)}")

st.markdown("---")

# ================= زر ٢: إرسال الإيميل فقط =================
st.subheader("📧 إرسال تنبيه الشحنات المتأخرة")

col_a, col_b = st.columns(2)
with col_a:
    custom_subject = st.text_input(
        "عنوان الإيميل",
        value="🚨 تنبيه: شحنات متأخرة (Aramex)"
    )
with col_b:
    custom_emails_input = st.text_input(
        "إيميلات المستلمين (افصلهم بفاصلة، أو اتركه فارغاً للافتراضي)",
        ""
    )

if st.button("📧 إرسال تنبيه الشحنات المتأخرة الآن", use_container_width=True):
    # استخراج الشحنات المتأخرة من البيانات الحالية (بس اللي مشحون فعلاً، مش غير مشحون)
    delayed_rows = []
    for row in policy_data[1:]:
        try:
            days = int(str(row[4]).strip()) if len(row) > 4 and str(row[4]).strip() else 0
        except:
            days = 0
        shipped = len(row) > 5 and row[5].strip() == "مشحون"
        if days > 3 and shipped and check_status(row[3] if len(row) > 3 else "") == "other":
            delayed_rows.append(row)

    delayed_rows = normalize_rows(delayed_rows, 6)

    if not delayed_rows:
        st.info("✅ لا توجد شحنات متأخرة لإرسال تنبيه عنها.")
    else:
        try:
            email_user     = st.secrets["email"]["username"]
            email_pass     = st.secrets["email"]["password"]
            default_emails = st.secrets["email"]["send_to"]
        except Exception as e:
            st.error("❌ لم يتم تحميل بيانات الإيميل من secrets.toml: " + str(e))
            st.stop()

        send_to = [e.strip() for e in custom_emails_input.split(",") if e.strip()] or default_emails
        subject = custom_subject or "🚨 تنبيه: شحنات متأخرة (Aramex)"

        message_body = f"يوجد {len(delayed_rows)} شحنة متأخرة تجاوزت 3 أيام:\n\n"
        for row in delayed_rows:
            message_body += f"- Order: {row[0]} | Policy: {row[1]} | Days: {row[4]}\n"

        msg = MIMEMultipart()
        msg["From"]    = email_user
        msg["To"]      = ", ".join(send_to)
        msg["Subject"] = subject
        msg.attach(MIMEText(message_body, "plain"))

        # مرفق Excel
        df_delayed = pd.DataFrame(
            delayed_rows,
            columns=["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment", "حالة الشحن"]
        )
        output = io.BytesIO()
        df_delayed.to_excel(output, index=False, sheet_name="Delayed Shipments")
        output.seek(0)
        part = MIMEApplication(output.read(), Name="Delayed_Shipments.xlsx")
        part['Content-Disposition'] = 'attachment; filename="Delayed_Shipments.xlsx"'
        msg.attach(part)

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_user, email_pass)
            server.sendmail(email_user, send_to, msg.as_string())
            server.quit()
            st.success(f"✅ تم إرسال التنبيه لـ {len(delayed_rows)} شحنة متأخرة إلى: {', '.join(send_to)}")
        except Exception as e:
            st.error(f"❌ فشل إرسال الإيميل: {e}")

# ================= عرض الإحصائيات والجداول =================
st.markdown("---")

def get_days_val(r):
    return int(str(r[4]).strip()) if len(r) > 4 and str(r[4]).strip().lstrip('-').isdigit() else 0

all_active       = [r for r in policy_data[1:] if check_status(r[3] if len(r) > 3 else "") == "other"]
delayed_display  = [r for r in all_active if get_days_val(r) > 3]
current_display  = [r for r in all_active if get_days_val(r) <= 3]

# غير مشحون من عندنا وعليه 3 أيام فأكثر
not_shipped_display = [
    r for r in policy_data[1:]
    if len(r) >= 6 and r[5].strip() == "غير مشحون" and get_days_val(r) >= 3
]

# مشحون عندنا لكن لسه واقف/متحركش فى أرامكس وعليه 3 أيام فأكثر
stuck_display = [
    r for r in all_active
    if len(r) >= 6 and r[5].strip() == "مشحون" and get_days_val(r) >= 3
]

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("📦 إجمالي النشطة", len(all_active))
col2.metric("⚠️ متأخرة (+3 أيام)", len(delayed_display))
col3.metric("✅ في الوقت",        len(current_display))
col4.metric("🚫 غير مشحون (+3 أيام)", len(not_shipped_display))
col5.metric("🐌 عالقة بأرامكس (+3 أيام)", len(stuck_display))

COLS = ["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment", "حالة الشحن"]

st.markdown("---")
tab_main, tab_not_shipped, tab_stuck = st.tabs([
    "🏠 الرئيسية",
    "🚫 غير مشحون (+3 أيام)",
    "🐌 عالقة بأرامكس (+3 أيام)"
])

with tab_main:
    st.subheader("🚨 الشحنات المتأخرة (أكثر من 3 أيام)")
    if delayed_display:
        st.dataframe(
            pd.DataFrame(normalize_rows(delayed_display, 6), columns=COLS),
            use_container_width=True,
            height=400
        )
    else:
        st.success("✅ لا توجد شحنات متأخرة!")

    st.markdown("---")
    st.subheader("📦 الشحنات الحالية")
    if current_display:
        st.dataframe(
            pd.DataFrame(normalize_rows(current_display, 6), columns=COLS),
            use_container_width=True,
            height=400
        )
    else:
        st.info("لا توجد شحنات حالياً.")

with tab_not_shipped:
    st.subheader("🚫 شحنات لم تُشحن من عندنا بعد (3 أيام فأكثر)")
    if not_shipped_display:
        st.dataframe(
            pd.DataFrame(normalize_rows(not_shipped_display, 6), columns=COLS),
            use_container_width=True,
            height=400
        )
    else:
        st.success("✅ لا توجد شحنات متأخرة فى الشحن!")

with tab_stuck:
    st.subheader("🐌 شحنات مشحونة ولم تتحرك فى أرامكس (3 أيام فأكثر)")
    if stuck_display:
        st.dataframe(
            pd.DataFrame(normalize_rows(stuck_display, 6), columns=COLS),
            use_container_width=True,
            height=400
        )
    else:
        st.success("✅ لا توجد شحنات عالقة!")

st.caption(f"آخر تحديث: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
