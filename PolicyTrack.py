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

# تحديث تلقائي كل 10 دقائق (يبقى لكن الإيميل لن يرسل إلا عند الضغط على زر التحديث)
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

# ================= المتغيرات لاسماء الشيتات =================
SHEET_NAME = "Complaints"
POLICY_SHEET = "Policy number"
DELIVERED_SHEET = "تم التسليم"
RETURNED_SHEET = "تم الإرجاع"
ORDERS_SHEET = "Order Number"
DELIVERED_ARCHIVE = "Delivered Archive"
RETURNED_ARCHIVE = "Returned Archive"

# ================= دوال مساعدة لإنشاء الأوراق لو مش موجودة =================
def get_or_create_sheet(sheet_name):
    try:
        return client.open(SHEET_NAME).worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sh = client.open(SHEET_NAME)
        sheet = sh.add_worksheet(title=sheet_name, rows="100", cols="10")
        sheet.append_row(["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment", "حالة الشحن"])
        return sheet

# تحميل أوراقنا
try:
    policy_sheet = client.open(SHEET_NAME).worksheet(POLICY_SHEET)
except Exception as e:
    st.error(f"❌ لا يمكن فتح الشيت الرئيسي '{SHEET_NAME}': {e}")
    st.stop()

delivered_sheet = get_or_create_sheet(DELIVERED_SHEET)
returned_sheet = get_or_create_sheet(RETURNED_SHEET)
delivered_archive_sheet = get_or_create_sheet(DELIVERED_ARCHIVE)
returned_archive_sheet = get_or_create_sheet(RETURNED_ARCHIVE)

# ================= تحميل شيت الاوردر =================
try:
    order_sheet = client.open(SHEET_NAME).worksheet(ORDERS_SHEET)
    order_data = order_sheet.get_all_values()
    order_dict = {row[1]: row[3] for row in order_data[1:] if len(row) > 3 and row[3].strip()}
except Exception:
    order_dict = {}

# ================= اعداد Aramex =================
client_info = {
    "UserName": "fitnessworld525@gmail.com",
    "Password": "Aa12345678@",
    "Version": "v1",
    "AccountNumber": "71958996",
    "AccountPin": "657448",
    "AccountEntity": "RUH",
    "AccountCountryCode": "SA"
}

# ================= دوال مساعدة =================
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
        response = requests.post(url, json=payload, headers=headers, timeout=15)
        if response.status_code != 200:
            return f"❌ فشل الاتصال ({response.status_code})"

        xml_content = response.content.decode('utf-8')
        xml_content = remove_xml_namespaces(xml_content)
        root = ET.fromstring(xml_content)

        tracking_results = root.find('TrackingResults')
        if tracking_results is None:
            # حاول إيجاد TrackingResult مباشرة
            tracks = root.findall(".//TrackingResult")
        else:
            keyvalue = tracking_results.find('KeyValueOfstringArrayOfTrackingResultmFAkxlpY')
            tracks = []
            if keyvalue is not None:
                tracking_array = keyvalue.find('Value')
                if tracking_array is not None:
                    tracks = tracking_array.findall('TrackingResult')
            if not tracks:
                tracks = root.findall(".//TrackingResult")

        if not tracks:
            return "❌ لا توجد حالة متاحة"

        last_track = sorted(
            tracks,
            key=lambda tr: tr.find('UpdateDateTime').text if tr.find('UpdateDateTime') is not None else '',
            reverse=True
        )[0]

        desc_en = last_track.find('UpdateDescription').text if last_track.find('UpdateDescription') is not None else "—"
        try:
            desc_ar = GoogleTranslator(source='en', target='ar').translate(desc_en)
        except:
            desc_ar = "—"

        return f"{desc_en} - {desc_ar}"
    except Exception as e:
        return f"⚠️ خطأ في جلب الحالة: {e}"

# ================= تحميل بيانات الشيت policy =================
try:
    policy_data = policy_sheet.get_all_values()
except Exception as e:
    st.error("❌ خطأ في قراءة بيانات الشيت الرئيسي: " + str(e))
    st.stop()

# ================= واجهة تعديل الايميل و العنوان =================
st.markdown("### 📧 إعدادات الإيميل (اختياري)")
custom_subject = st.text_input(
    "عنوان الإيميل (اختياري)",
    value="🚨 تنبيه: شحنات متأخرة (Noon – Aramex)"
)
custom_emails_input = st.text_input(
    "إيميلات المستلمين (افصلهم بفاصلة , إذا تريد تغييرهم — اختيارية)",
    ""
)

# ================= دالة توحيد الصفوف =================
def normalize_rows(rows, n):
    fixed = []
    for r in rows:
        # ensure list
        r = list(r)
        r = r[:n]
        r += ["—"] * (n - len(r))
        fixed.append(r)
    return fixed

# ================= دالة إرسال الإيميل مع اكسل مرفق =================
def send_delay_email(delayed_rows, custom_emails=None, custom_subject=None):
    if not delayed_rows:
        return

    try:
        email_user = st.secrets["email"]["username"]
        email_pass = st.secrets["email"]["password"]
        default_emails = st.secrets["email"]["send_to"]
    except Exception as e:
        st.error("❌ لم يتم تحميل بيانات الإيميل من secrets.toml: " + str(e))
        return

    # اختيار من المستلمين
    if custom_emails:
        send_to = [e.strip() for e in custom_emails.split(",") if e.strip()]
    else:
        send_to = default_emails

    # عنوان
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
    df = pd.DataFrame(delayed_rows, columns=["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment", "حالة الشحن"])
    output = io.BytesIO()
    df.to_excel(output, index=False, sheet_name="Delayed Shipments")
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
    except Exception as e:
        st.error(f"❌ فشل إرسال الإيميل: {e}")

# ================= تحديث Days و حالة الشحن (كما في الكود القديم) =================
# التأكد من وجود على الأقل صف الرأس
if len(policy_data) < 2:
    st.info("ملف الشحنات فارغ أو لا يحتوي على بيانات بعد الصف الأول.")
else:
    cells = policy_sheet.range(f'E2:E{len(policy_data)}')
    for idx, row in enumerate(policy_data[1:]):
        if len(row) < 6:
            row += ["", ""] * (6 - len(row))
        date_added_str = row[2] if len(row) > 2 else None
        days_diff = 0
        if date_added_str and str(date_added_str).strip():
            for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
                try:
                    dt = datetime.strptime(date_added_str, fmt)
                    days_diff = (datetime.now() - dt).days
                    break
                except:
                    continue
        # ضع قيمة صفر افتراضية لو لم يتعرف
        try:
            row[4] = int(days_diff)
        except:
            row[4] = 0
        cells[idx].value = row[4]
        # حالة الشحن بناء على شيت الاوردر
        order_num = str(row[0]) if len(row) > 0 else ""
        row[5] = "مشحون" if order_num in order_dict else "غير مشحون"
    try:
        policy_sheet.update_cells(cells)
    except Exception as e:
        st.warning("تحذير: لم يتم تحديث عمود الأيام في الشيت بسبب: " + str(e))

# ================= واجهة البحث =================
st.header("🔍 البحث عن شحنة")
search_order = st.text_input("أدخل رقم الطلب للبحث")
if search_order.strip():
    found = False
    for row in policy_data[1:]:
        if len(row) >= 2 and str(row[0]) == search_order:
            found = True
            st.success(f"✅ تم العثور على الطلب رقم: {search_order}")
            st.info(f"📦 رقم الشحنة: {row[1]}")
            st.write(f"📅 التاريخ: {row[2] if len(row) > 2 else '—'}")
            st.write(f"🔄 الحالة الحالية: {row[3] if len(row) > 3 else '—'}")
            st.write(f"⏳ أيام منذ الشحن: {row[4] if len(row) > 4 else '—'}")
            st.write(f"🚚 حالة الشحن: {row[5] if len(row) > 5 else 'غير معروف'}")
            break
    if not found:
        st.error("⚠️ لم يتم العثور على الطلب في الشيت")

# ================= دالة التصنيف (نفس منطق الكود القديم) =================
def check_status(status_text):
    text = (status_text or "").lower()
    delivered_conditions = [
        "delivered", "تم التسليم", "shipment charges paid", "customer id received",
        "collected by consignee", "delivery", "delivered to consignee"
    ]
    returned_conditions = [
        "returned", "تم الإرجاع", "returned to shipper", "return to sender", "return"
    ]
    for cond in delivered_conditions:
        if cond in text:
            return "delivered"
    for cond in returned_conditions:
        if cond in text:
            return "returned"
    return "other"

# ================= زر تحديث جميع الحالات الآن (التحديث + نقل الوصلات) =================
if st.button("تحديث جميع الحالات الآن"):
    progress = st.progress(0)
    total = len(policy_data) - 1 if len(policy_data) > 1 else 1
    updated_rows = []
    # نعمل نسخة محدثة من policy_data لنستخدمها لاحقاً
    for idx, row in enumerate(policy_data[1:], start=2):
        # تأكد من طول الصف
        if len(row) < 6:
            row += [""] * (6 - len(row))
        # إذا لم تكن الحالة مُصنفة كـ delivered/returned، نجلب الحالة
        if check_status(row[3]) == "other":
            try:
                new_status = get_aramex_status(row[1])
                row[3] = new_status
            except Exception as e:
                row[3] = row[3]  # احتفظ بالحالة القديمة لو فشل
        updated_rows.append(row)
        # تحديث البار
        try:
            progress.progress((idx-1) / total)
        except:
            pass

    # تحديث العمود D (الحالة) دفعة واحدة
    try:
        cells = policy_sheet.range(f'D2:D{len(policy_data)}')
        for i, r in enumerate(updated_rows):
            cells[i].value = r[3]
        policy_sheet.update_cells(cells)
    except Exception as e:
        st.warning("⚓ تحذير: لم يتم تحديث عمود الحالة في الشيت: " + str(e))

    st.success("✅ تم تحديث الحالات من مزود الشحن (Aramex)")

    # ================= الآن نطبق منطق النقل للشحنات التي وصلت أو عادت (نفس منطق الكود القديم) =================
    # جلب القوائم الحالية من delivered/returned sheets (بدون رؤوس)
    try:
        delivered_shipments_existing = delivered_sheet.get_all_values()[1:]
    except Exception:
        delivered_shipments_existing = []
    try:
        returned_shipments_existing = returned_sheet.get_all_values()[1:]
    except Exception:
        returned_shipments_existing = []

    # جهز صفوف جديدة للوصول والارتجاع (نأخذ أول 5 أعمدة منهم كما في الكود القديم)
    new_delivered = []
    new_returned = []
    for r in updated_rows:
        try:
            status_flag = check_status(r[3])
            if status_flag == "delivered":
                # إذا لم تكن موجودة مسبقاً في delivered_sheet
                if r[1] not in [x[1] for x in delivered_shipments_existing]:
                    new_delivered.append(r[:6])  # خذ 6 أعمدة للحفاظ على البنية
            elif status_flag == "returned":
                if r[1] not in [x[1] for x in returned_shipments_existing]:
                    new_returned.append(r[:6])
        except Exception:
            continue

    # دالة لإضافة في دفعات
    def append_in_batches(sheet, rows, batch_size=50):
        if not rows:
            return
        for i in range(0, len(rows), batch_size):
            batch = rows[i:i+batch_size]
            try:
                sheet.append_rows(batch, value_input_option='USER_ENTERED')
            except Exception as e:
                # حاول واحد واحد لو فشل الدفعة
                for row in batch:
                    try:
                        sheet.append_row(row, value_input_option='USER_ENTERED')
                        time.sleep(0.2)
                    except Exception:
                        pass
            time.sleep(0.5)

    # أضف الجديد إلى delivered و الأرشيف، ثم امسح من policy_sheet
    if new_delivered:
        append_in_batches(delivered_sheet, [nd[:6] for nd in new_delivered])
        append_in_batches(delivered_archive_sheet, [nd[:6] for nd in new_delivered])
        # حذف الصفوف من policy_sheet بناء على رقم البوليصه
        for nd in new_delivered:
            pol = nd[1]
            # البحث عن الصف الذي يحتوي على البوليصه
            try:
                all_policy = policy_sheet.get_all_values()
                for i, row in enumerate(all_policy[1:], start=2):
                    try:
                        if len(row) > 1 and row[1] == pol:
                            policy_sheet.delete_rows(i)
                            break
                    except Exception:
                        continue
            except Exception:
                continue

    if new_returned:
        append_in_batches(returned_sheet, [nr[:6] for nr in new_returned])
        append_in_batches(returned_archive_sheet, [nr[:6] for nr in new_returned])
        for nr in new_returned:
            pol = nr[1]
            try:
                all_policy = policy_sheet.get_all_values()
                for i, row in enumerate(all_policy[1:], start=2):
                    try:
                        if len(row) > 1 and row[1] == pol:
                            policy_sheet.delete_rows(i)
                            break
                    except Exception:
                        continue
            except Exception:
                continue

    # ================= بعد النقل - نعيد تحميل policy_data لعرض صحيح =================
    try:
        policy_data = policy_sheet.get_all_values()
    except Exception:
        policy_data = updated_rows  # fallback

    # ================= استخراج الشحنات المتأخرة (Days > 3) كما في الكود القديم =================
    delayed_shipments = [row for row in policy_data[1:] if int(row[4] if str(row[4]).strip() else 0) > 3 and check_status(row[3]) == "other"]
    # تأكد من البنية
    delayed_shipments = normalize_rows(delayed_shipments, 6)

    # ================= إرسال الإيميل بعد انتهاء عملية التحديث بالكامل =================
    send_delay_email(
        delayed_shipments,
        custom_emails=custom_emails_input,
        custom_subject=custom_subject
    )

    st.success("✔️ تم إرسال تنبيه الشحنات المتأخرة (إن وُجدت) بالبريد الإلكتروني.")

# ================= تصنيف البيانات للعرض كما في الكود القديم =================
# تأكد بترتيب واحد: شحنات متأخرة (Days>3) ثم الشحنات الحالية (Days<=3) بنفس سياسة check_status
delayed_shipments = [row for row in policy_data[1:] if int(row[4] if str(row[4]).strip() else 0) > 3 and check_status(row[3]) == "other"]
current_shipments = [row for row in policy_data[1:] if int(row[4] if str(row[4]).strip() else 0) <= 3 and check_status(row[3]) == "other"]

delayed_shipments = normalize_rows(delayed_shipments, 6)
current_shipments = normalize_rows(current_shipments, 6)

# ================= عرض الشحنات المتأخرة =================
st.markdown("---")
st.subheader("🚨 الشحنات المتأخرة")
if delayed_shipments:
    try:
        df_delayed = pd.DataFrame(delayed_shipments, columns=["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment", "حالة الشحن"])
        st.dataframe(df_delayed, use_container_width=True)
    except Exception as e:
        st.error("خطأ في عرض جدول الشحنات المتأخرة: " + str(e))
else:
    st.info("لا توجد شحنات متأخرة حالياً.")

# ================= عرض الشحنات الحالية =================
st.markdown("---")
st.subheader("📦 الشحنات الحالية")
if current_shipments:
    try:
        df_current = pd.DataFrame(current_shipments, columns=["Order Number", "Policy Number", "Date", "Status", "Days Since Shipment", "حالة الشحن"])
        st.dataframe(df_current, use_container_width=True)
    except Exception as e:
        st.error("خطأ في عرض جدول الشحنات الحالية: " + str(e))
else:
    st.info("لا توجد شحنات حالياً.")

st.success("🚀 ")
