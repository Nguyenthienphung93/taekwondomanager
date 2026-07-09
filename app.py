from datetime import datetime, date, timedelta, timezone
import calendar
import json
import os
from io import BytesIO
from PIL import Image, ImageOps, ImageFilter
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, send_file
from supabase_client import supabase
from utils import auto_generate_license, is_active, calc_tuition, format_money, build_month_codes
import re
import unicodedata
from pathlib import Path
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-this-secret")
BASE_DIR = Path(__file__).resolve().parent
STUDENT_PHOTO_DIR = BASE_DIR / "static" / "student_photos"
STUDENT_PHOTO_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_PHOTO_EXTENSIONS = {"jpg", "jpeg", "png", "webp"}
SUPABASE_URL = os.environ.get(
    "SUPABASE_URL",
    "https://hbmqqvrhjnhxgrxvlthm.supabase.co"
)

STUDENT_PHOTO_BUCKET = "student-photos"
CLUB_ASSET_BUCKET = "system-assets"


DEFAULT_APP_SETTINGS = {
    "header": {
        "club_small_title": "CÂU LẠC BỘ TAEKWONDO",
        "club_name": "HOA HƯỚNG DƯƠNG",
        "logo_url": "https://hbmqqvrhjnhxgrxvlthm.supabase.co/storage/v1/object/public/system-assets/logo.png"
    },
    "fees": {
        "tuition_fee": 500000,
        "exam_fee": 300000,
        "dan_fees": {
            "1 Đẳng": 950000,
            "2 Đẳng": 650000,
            "3 Đẳng": 800000,
            "4 Đẳng": 1250000,
            "5 Đẳng": 1500000,
            "6 Đẳng": 2000000
        }
    },
    "exam": {
        "exam_number_prefix": "Cấp_"
    },
    "class_options": {
        "classrooms": ["2 - 4 - 6", "3 - 5 - 7", "T7 - CN", "Hẹn hò"],
        "timeclasses": ["Ca 1", "Ca 2"],
        "clubs": ["Hoa Hướng Dương Q1(CLB_00019)", "Hoa Hướng Dương Q10(CLB_00104)"]
    },

    "club_info": {
        "intro_title": "CLB Taekwondo Sunflower - Hoa Hướng Dương Diên Hồng",
        "intro_subtitle": "Sunflower Taekwondo Club",
        "intro_content": (
            "CLB Taekwondo Sunflower, còn được biết đến với tên CLB Taekwondo Hoa Hướng Dương, "
            "là môi trường tập luyện Taekwondo dành cho võ sinh nhiều lứa tuổi. Hiện nay cơ sở Q10 "
            "hoạt động với tên CLB Hoa Hướng Dương - Diên Hồng, định hướng đào tạo võ sinh theo tinh thần "
            "kỷ luật, tự tin, lễ phép và phát triển thể chất toàn diện.\n\n"
            "CLB chú trọng nền tảng kỹ thuật căn bản, rèn luyện thể lực, tác phong võ đạo, đồng thời tạo điều kiện "
            "cho võ sinh tham gia thi nâng cấp đai, biểu diễn, giao lưu và các hoạt động phong trào phù hợp."
        ),

        "head_coach": {
            "name": "Nguyễn Thiên Phụng",
            "role": "HLV Trưởng",
            "phone": "",
            "photo_url": "",
            "description": (
                "Nguyễn Thiên Phụng là VĐV Taekwondo Poomsae Việt Nam, nhiều năm thi đấu đỉnh cao "
                "và đạt nhiều thành tích nổi bật trong nước cũng như quốc tế. Hiện tại phụ trách định hướng "
                "chuyên môn chung, xây dựng chương trình huấn luyện, theo dõi chất lượng đào tạo và phát triển "
                "phong trào của CLB Hoa Hướng Dương - Diên Hồng."
            )
        },

        "registrar": {
            "name": "Đang cập nhật",
            "role": "Người ghi danh",
            "phone": "",
            "photo_url": "",
            "description": (
                "Phụ trách tư vấn, tiếp nhận thông tin đăng ký, hỗ trợ phụ huynh và võ sinh "
                "trong quá trình ghi danh, đóng học phí và theo dõi lịch học tại CLB."
            )
        },

        "coaches": [
            {
                "group": "246",
                "group_title": "HLV lớp 2-4-6",
                "group_subtitle": "Lớp buổi chiều",
                "name": "Nguyễn Hoàng Long",
                "role": "Phụ Trách HLV",
                "phone": "",
                "photo_url": ""
            },
            {
                "group": "246",
                "group_title": "HLV lớp 2-4-6",
                "group_subtitle": "Lớp buổi chiều",
                "name": "Nguyễn Duy Thông",
                "role": "Trợ giảng",
                "phone": "",
                "photo_url": ""
            },
            {
                "group": "357",
                "group_title": "HLV lớp 3-5-7",
                "group_subtitle": "Lớp chính trong tuần",
                "name": "Nông Thạch Khiêm",
                "role": "Phụ Trách HLV",
                "phone": "",
                "photo_url": ""
            },
            {
                "group": "357",
                "group_title": "HLV lớp 3-5-7",
                "group_subtitle": "Lớp chính trong tuần",
                "name": "Nguyễn Trung Hiếu",
                "role": "Phụ Trách HLV",
                "phone": "",
                "photo_url": ""
            },
            {
                "group": "357",
                "group_title": "HLV lớp 3-5-7",
                "group_subtitle": "Lớp chính trong tuần",
                "name": "Trần Ngọc Hà",
                "role": "Trợ giảng",
                "phone": "",
                "photo_url": ""
            },
            {
                "group": "weekend",
                "group_title": "HLV lớp sáng & cuối tuần",
                "group_subtitle": "Thứ 7, Chủ nhật và lớp sáng",
                "name": "Nguyễn Lê Cường",
                "role": "Phụ Trách HLV",
                "phone": "",
                "photo_url": ""
            }
        ],

        "regular_schedules": [
            {"title": "Lớp 2-4-6", "time": "18:00 - 19:30 và 19:30 - 21:00"},
            {"title": "Lớp 3-5-7", "time": "18:00 - 19:30 và 19:30 - 21:00"},
            {"title": "Lớp Thứ 7 và Chủ nhật", "time": "8:30 - 10:30"}
        ],

        "summer_schedules": [
            {"title": "Lớp hè", "time": "Từ đầu tháng 6 đến cuối tháng 8"},
            {"title": "Sáng 2-4-6 hoặc 3-5-7", "time": "8:00 - 9:30 và 9:30 - 11:00"}
        ],

        "fees_info": {
            "monthly_fee": "500.000đ",
            "three_month_discount": "Giảm 10%",
            "six_month_bonus": "Tặng 1 tháng miễn phí",
            "family_discount": "Giảm thêm 10% nếu đóng chung",
            "uniform_fee": "420.000đ / bộ",
            "exam_fee": "300.000đ / lần",
            "exam_notes": [
                "3 tháng CLB sẽ tổ chức thi nâng cấp đai 1 lần.",
                "Võ sinh cần đảm bảo thời gian tập luyện và chuyên môn, được HLV phụ trách xác nhận trước khi thi.",
                "Võ sinh mới thi lần đầu đóng thêm 20.000đ một lần duy nhất để Liên đoàn Taekwondo Việt Nam cấp mã, phục vụ hoạt động quản lý chung và đảm bảo quyền lợi trong quá trình tập luyện."
            ]
        }
    },

    "profile": {
        "owner": "Nguyễn Thiên Phụng",
        "phone_zalo": "0989 03 04 93",
        "email": "nhoctotokute93@gmail.com",
        "role": "VĐV Taekwondo Poomsae Việt Nam",
        "system_name": "Hệ thống quản lý hội viên CLB Taekwondo"
    },
}


def merge_settings(defaults, saved):
    result = {}

    for key, value in defaults.items():
        if isinstance(value, dict):
            result[key] = merge_settings(value, (saved or {}).get(key, {}))
        else:
            result[key] = (saved or {}).get(key, value)

    return result


def load_app_settings():
    """
    Đọc setup hệ thống từ Supabase.
    Nếu Supabase chưa có dữ liệu thì tự tạo từ DEFAULT_APP_SETTINGS.
    """
    try:
        rows = supabase.table(APP_SETTINGS_TABLE) \
            .select("key,value") \
            .execute().data or []

        if not rows:
            save_app_settings(DEFAULT_APP_SETTINGS)
            return DEFAULT_APP_SETTINGS.copy()

        saved = {}

        for row in rows:
            key = str(row.get("key") or "").strip()
            value = row.get("value") or {}

            if key:
                saved[key] = value

        return merge_settings(DEFAULT_APP_SETTINGS, saved)

    except Exception as e:
        print("[LOAD APP SETTINGS SUPABASE ERROR]", e)

        # Dự phòng nếu Supabase lỗi thì vẫn cho app chạy bằng mặc định
        return DEFAULT_APP_SETTINGS.copy()


def save_app_settings(settings):
    """
    Lưu setup hệ thống lên Supabase.
    Mỗi nhóm setup là 1 dòng:
    header, fees, exam, class_options, club_info, profile...
    """
    try:
        now_iso = datetime.now(timezone.utc).isoformat()

        for key, value in settings.items():
            payload = {
                "key": key,
                "value": value,
                "updated_at": now_iso,
            }

            existing = supabase.table(APP_SETTINGS_TABLE) \
                .select("key") \
                .eq("key", key) \
                .limit(1) \
                .execute().data or []

            if existing:
                supabase.table(APP_SETTINGS_TABLE) \
                    .update(payload) \
                    .eq("key", key) \
                    .execute()
            else:
                supabase.table(APP_SETTINGS_TABLE) \
                    .insert(payload) \
                    .execute()

    except Exception as e:
        print("[SAVE APP SETTINGS SUPABASE ERROR]", e)
        raise e


def money_to_int_web(v):
    return int(str(v or "0")
        .replace("đ", "")
        .replace("Đ", "")
        .replace(".", "")
        .replace(",", "")
        .replace(" ", "")
        .strip() or 0)


def get_app_setting(path, default=None):
    data = load_app_settings()
    cur = data

    for part in str(path).split("."):
        if not isinstance(cur, dict):
            return default
        cur = cur.get(part)

    return default if cur is None else cur


def get_app_setting_int(path, default=0):
    try:
        return int(get_app_setting(path, default) or default)
    except:
        return default

def clean_transfer_note_no_accent(text):
    text = str(text or "").strip()

    # Bỏ dấu tiếng Việt
    text = remove_accents(text)

    # Chỉ giữ ký tự an toàn cho nội dung chuyển khoản
    text = re.sub(r"[^A-Za-z0-9\s_\-\.]", "", text)

    # Gom khoảng trắng
    text = re.sub(r"\s+", " ", text).strip()

    return text


def build_transfer_note(template, student):
    template = str(template or "").strip()

    if not template:
        template = "{student_name}"

    student = student or {}

    note = (
        template
        .replace("{student_name}", str(student.get("name") or "").strip())
        .replace("{student_license}", str(student.get("license") or "").strip())
        .replace("{ma_hv}", str(student.get("license") or "").strip())
    ).strip()

    return clean_transfer_note_no_accent(note)


def build_vietqr_url(payment_info, student=None, amount=None):
    payment_info = payment_info or {}

    bank_code = str(payment_info.get("bank_code") or "").strip()
    account_number = str(payment_info.get("account_number") or "").strip()
    account_name = str(payment_info.get("account_name") or "").strip()
    note = build_transfer_note(payment_info.get("transfer_note"), student or {})

    if not bank_code or not account_number:
        return ""

    params = []

    if amount:
        try:
            amount_int = int(amount or 0)
            if amount_int > 0:
                params.append(f"amount={amount_int}")
        except:
            pass

    if note:
        params.append(f"addInfo={note}")

    if account_name:
        params.append(f"accountName={account_name}")

    query = "&".join(params)

    if query:
        return f"https://img.vietqr.io/image/{bank_code}-{account_number}-compact2.png?{query}"

    return f"https://img.vietqr.io/image/{bank_code}-{account_number}-compact2.png"

# =========================
# ADMIN LOGIN - BỘ QUẢN LÝ
# =========================
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "nhoctotokute93")
ADMIN_PASSWORD_HASH = generate_password_hash(
    os.environ.get("ADMIN_PASSWORD", "Nguyenthienphung#93")
)

ADMIN_PUBLIC_ENDPOINTS = {
    "admin_login",
    "static",
    "index",

    # Các route học viên để không bị khóa bởi admin login
    "student_login",
    "student_logout",
    "student_portal_home",
    "student_portal_notifications",
    "student_portal_notification_detail",
    "student_portal_info",
    "student_portal_club_info",
    "student_portal_fees",
    "student_portal_exams",
    "student_portal_activities",
    "student_portal_settings",

    "coach_login",
    "coach_logout",
    "coach_exam",
    "coach_dan",
    "coach_exam_update_status",
}

ADMIN_PUBLIC_PATH_PREFIXES = (
    "/static/",
    "/student-login",
    "/student-logout",
    "/student-portal",
    "/coach-login",
    "/coach-logout",
    "/coach/",
)

def get_default_payment_settings():
    return {
        "id": "club_payment",
        "account_name": "",
        "account_number": "",
        "bank_code": "ACB",
        "bank_name": "ACB",
        "transfer_note": "{student_name}",
    }


def get_payment_settings():
    try:
        rows = supabase.table(PAYMENT_SETTINGS_TABLE) \
            .select("*") \
            .eq("id", "club_payment") \
            .limit(1) \
            .execute().data or []

        if rows:
            data = rows[0]
        else:
            data = get_default_payment_settings()

            supabase.table(PAYMENT_SETTINGS_TABLE) \
                .insert(data) \
                .execute()

        default_data = get_default_payment_settings()

        for key, value in default_data.items():
            if not data.get(key):
                data[key] = value

        return data

    except Exception as e:
        print("[GET PAYMENT SETTINGS ERROR]", e)
        return get_default_payment_settings()


def save_payment_settings(form):
    bank_code = str(form.get("bank_code") or "").strip()

    bank_map = {
        "VCB": "Vietcombank",
        "TCB": "Techcombank",
        "MB": "MB Bank",
        "ACB": "ACB",
        "BIDV": "BIDV",
        "VTB": "VietinBank",
        "VPB": "VPBank",
        "TPB": "TPBank",
        "VIB": "VIB",
        "OCB": "OCB",
        "STB": "Sacombank",
        "HDB": "HDBank",
        "SHB": "SHB",
        "EIB": "Eximbank",
        "MSB": "MSB",
        "BAB": "Bac A Bank",
        "SEAB": "SeABank",
        "LPB": "LPBank",
        "VAB": "VietABank",
        "ABB": "ABBank",
        "NAB": "Nam A Bank",
        "PGB": "PGBank",
        "PVCB": "PVcomBank",
    }

    payload = {
        "id": "club_payment",
        "account_name": str(form.get("account_name") or "").strip(),
        "account_number": str(form.get("account_number") or "").strip(),
        "bank_code": bank_code,
        "bank_name": bank_map.get(bank_code, bank_code),
        "transfer_note": str(form.get("transfer_note") or "{student_name}").strip() or "{student_name}",
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }

    existing = supabase.table(PAYMENT_SETTINGS_TABLE) \
        .select("id") \
        .eq("id", "club_payment") \
        .limit(1) \
        .execute().data or []

    if existing:
        supabase.table(PAYMENT_SETTINGS_TABLE) \
            .update(payload) \
            .eq("id", "club_payment") \
            .execute()
    else:
        supabase.table(PAYMENT_SETTINGS_TABLE) \
            .insert(payload) \
            .execute()

    return payload


def is_admin_logged_in():
    return session.get("admin_logged_in") is True


def is_safe_next_url(next_url):
    next_url = str(next_url or "").strip()
    return next_url.startswith("/") and not next_url.startswith("//")

def back_to_current_page(default_endpoint):
    """
    Quay lại đúng trang vừa thao tác.
    Dùng cho xóa/sửa để không bị nhảy về trang 1.
    """
    next_url = request.form.get("next") or request.referrer or url_for(default_endpoint)
    next_url = str(next_url or "").strip()

    # Chỉ cho redirect nội bộ, tránh redirect ra link ngoài
    if next_url.startswith(request.host_url):
        next_url = next_url.replace(request.host_url.rstrip("/"), "", 1)

    if not is_safe_next_url(next_url):
        next_url = url_for(default_endpoint)

    return redirect(next_url)

@app.before_request
def require_admin_login_for_management_pages():
    path = request.path or "/"

    # Cho phép các endpoint public
    if request.endpoint in ADMIN_PUBLIC_ENDPOINTS:
        return None

    # Cho phép static và bộ học viên
    if any(path.startswith(prefix) for prefix in ADMIN_PUBLIC_PATH_PREFIXES):
        return None

    # Cho phép trang đăng nhập admin
    if path == "/admin-login":
        return None

    # Nếu admin đã đăng nhập thì cho qua
    if is_admin_logged_in():
        return None

    # Chưa đăng nhập thì chuyển về login admin
    return redirect(url_for("admin_login", next=path))


@app.route("/admin-login", methods=["GET", "POST"])
def admin_login():
    if is_admin_logged_in():
        next_url = request.args.get("next") or url_for("students")

        if not is_safe_next_url(next_url):
            next_url = url_for("students")

        return redirect(next_url)

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        remember = request.form.get("remember") == "on"
        next_url = request.form.get("next") or url_for("students")

        if not is_safe_next_url(next_url):
            next_url = url_for("students")

        if username == ADMIN_USERNAME and check_password_hash(ADMIN_PASSWORD_HASH, password):
            # Không dùng session.clear(), vì sẽ làm out Student và Coach
            session["admin_logged_in"] = True
            session["admin_username"] = username
            session.permanent = remember

            flash("Đăng nhập admin thành công", "success")
            return redirect(next_url)

        flash("ID hoặc mật khẩu admin chưa đúng")
        return redirect(url_for("admin_login", next=next_url))

    next_url = request.args.get("next") or url_for("students")

    if not is_safe_next_url(next_url):
        next_url = url_for("students")

    return render_template("admin_login.html", next_url=next_url)


@app.get("/admin-logout")
def admin_logout():
    session.pop("admin_logged_in", None)
    session.pop("admin_username", None)

    flash("Đã đăng xuất admin")
    return redirect(url_for("admin_login"))

@app.context_processor
def inject_app_settings():
    settings = load_app_settings()

    fees = settings.get("fees", {})
    exam = settings.get("exam", {})
    class_options = settings.get("class_options", DEFAULT_APP_SETTINGS["class_options"])

    return {
        "app_settings": settings,
        "admin_logged_in": is_admin_logged_in(),
        "admin_username": session.get("admin_username", ADMIN_USERNAME),
        "tuition_fee": int(fees.get("tuition_fee", 500000) or 500000),
        "exam_fee": int(fees.get("exam_fee", 300000) or 300000),
        "dan_fees": fees.get("dan_fees", {}),
        "exam_number_prefix": exam.get("exam_number_prefix", "Cấp_"),
        "class_options": class_options
    }

@app.template_filter('money')
def money_filter(v):
    return format_money(v or 0)

@app.template_filter("datetime_vn")
def datetime_vn_filter(value):
    raw = str(value or "").strip()

    if not raw:
        return "—"

    try:
        dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))

        return dt.strftime("%d/%m/%Y %H:%M")

    except Exception:
        return raw

@app.get('/')
def index():
    return redirect(url_for('student_login'))

def remove_accents(text):
    text = str(text or "").replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def normalize_birthdate_web(raw):
    s = re.sub(r"\D", "", str(raw or ""))

    if len(s) == 8:
        day, month, year = s[:2], s[2:4], s[4:]
    elif len(s) == 6:
        day, month, yy = s[:2], s[2:4], s[4:]
        year = "19" + yy if int(yy) >= 30 else "20" + yy
    else:
        return ""

    try:
        d = datetime(int(year), int(month), int(day))
        return d.strftime("%d/%m/%Y")
    except:
        return ""


def auto_hv_code_web(name, birthdate):
    name = remove_accents(name).lower().strip()
    parts = [p for p in name.split() if p]

    if not parts:
        return ""

    ten = parts[-1]
    initials = "".join(p[0] for p in parts[:-1])

    birth = normalize_birthdate_web(birthdate)
    if not birth:
        return ""

    dd, mm, yyyy = birth.split("/")
    return f"HV_{ten}{initials}_{dd}{mm}{yyyy[-2:]}"

@app.get('/students')
def students():
    q = request.args.get('q', '').strip()
    new_license = request.args.get('new', '').strip()

    try:
        query = supabase.table(STUDENT_TABLE).select('*')

        if q:
            query = query.or_(f'license.ilike.%{q}%,name.ilike.%{q}%,phonenumber.ilike.%{q}%')

        rows = query.execute().data or []

    except Exception as e:
        print("[SUPABASE ERROR /students]", e)
        flash("Không kết nối được Supabase. Ken kiểm tra lại mạng, DNS hoặc SUPABASE_URL trong file .env.")
        rows = []

    def sort_student(r):
        license_code = str(r.get('license') or '').strip()

        # Hội viên vừa thêm luôn nằm đầu danh sách
        new_rank = 0 if new_license and license_code == new_license else 1

        # Sau đó vẫn ưu tiên HV đang hoạt động
        active_rank = 0 if is_active(r.get('active')) else 1

        # Sau đó sắp xếp theo id/idx mới nhất
        new_id = r.get('id') or r.get('idx') or 0
        try:
            new_id = int(new_id)
        except:
            new_id = 0

        return (new_rank, active_rank, -new_id)

    rows.sort(key=sort_student)

    return render_template(
        'students.html',
        rows=rows,
        q=q,
        new_license=new_license
    )

@app.post('/students/add')
def students_add():
    form = request.form

    birthdate = normalize_birthdate_web(form.get("birthdate", ""))

    if not birthdate:
        flash("Ngày sinh không hợp lệ")
        return redirect(url_for("students"))

    license_code = form.get("license", "").strip()

    if not license_code:
        license_code = auto_hv_code_web(form.get("name", ""), birthdate)

    if not license_code:
        flash("Không tạo được Mã HV. Ken kiểm tra lại Họ tên và Ngày sinh.")
        return redirect(url_for("students"))

    # =========================
    # CHẶN TRÙNG MÃ HỘI VIÊN
    # =========================
    duplicated = supabase.table(STUDENT_TABLE) \
        .select("license,name,birthdate") \
        .eq("license", license_code) \
        .limit(1) \
        .execute().data or []

    if duplicated:
        old = duplicated[0]
        flash(
            f"Mã HV {license_code} đã tồn tại cho học viên "
            f"{old.get('name', '')} - {old.get('birthdate', '')}. "
            f"Ken không thể thêm trùng mã."
        )
        return redirect(url_for("students"))
    payload = {
        'license': license_code,
        'name': form.get('name','').strip(),
        'birthdate': birthdate,
        'gender': form.get('gender','').strip(),
        'classroom': form.get('classroom','').strip(),
        'timeclass': form.get('timeclass','').strip(),
        'clup': form.get('clup','').strip(),
        'phonenumber': form.get('phonenumber','').strip(),
        'address': form.get('address', '').strip(),
        'belt': form.get('belt','Cấp 10').strip() or 'Cấp 10',
        'family': form.get('family','Không'),
        'active': form.get('active','Có'),
        'telegram_id': form.get('telegram_id','').strip(),
    }
    supabase.table(STUDENT_TABLE).insert(payload).execute()
    flash(f'Đã thêm học viên: {license_code}')

    return redirect(url_for('students', new=license_code))

@app.post('/students/<license_code>/delete')
def students_delete(license_code):
    supabase.table(STUDENT_TABLE).delete().eq('license', license_code).execute()
    flash(f'Đã xoá học viên: {license_code}', "success")
    return back_to_current_page("students")
    
@app.post('/students/<license_code>/update')
def students_update(license_code):
    form = request.form

    birthdate = normalize_birthdate_web(form.get("birthdate", ""))

    payload = {
        'name': form.get('name', '').strip(),
        'birthdate': birthdate,
        'gender': form.get('gender', '').strip(),
        'classroom': form.get('classroom', '').strip(),
        'timeclass': form.get('timeclass', '').strip(),
        'clup': form.get('clup', '').strip(),
        'phonenumber': form.get('phonenumber', '').strip(),
        'address': form.get('address', '').strip(),
        'belt': form.get('belt', 'Cấp 10').strip(),
        'family': form.get('family', 'Không'),
        'active': form.get('active', 'Có'),
        'telegram_id': form.get('telegram_id', '').strip(),
    }

    supabase.table(STUDENT_TABLE).update(payload).eq('license', license_code).execute()
    flash(f"Đã cập nhật học viên: {license_code}", "success")
    return back_to_current_page("students")

@app.get('/fees')
def fees():
    q = request.args.get('q','').strip()
    selected_license = request.args.get('ma_hv', '').strip()
    query = supabase.table(HOCPHI_TABLE).select('*').order('thoi_gian', desc=True).limit(300)

    if q:
        query = query.or_(f'ma_hv.ilike.%{q}%,ho_ten.ilike.%{q}%')

    rows = query.execute().data or []

    students = supabase.table(STUDENT_TABLE) \
        .select('license,name,birthdate,gender,classroom,timeclass,phonenumber,belt,family') \
        .order('name') \
        .execute().data or []

    now = datetime.now()

    return render_template(
        'fees.html',
        rows=rows,
        students=students,
        q=q,
        current_month=now.month,
        current_year=now.year,
        selected_license=selected_license
    )

@app.post('/fees/add')
def fees_add():
    f = request.form
    ma_hv = f.get('ma_hv', '').strip()

    sv = supabase.table(STUDENT_TABLE) \
        .select('*') \
        .eq('license', ma_hv) \
        .limit(1) \
        .execute().data

    if not sv:
        flash('Không tìm thấy Mã HV')
        return redirect(url_for('fees'))

    sv = sv[0]

    def money_to_int(v):
        return int(
            str(v or '0')
            .replace('đ', '')
            .replace('Đ', '')
            .replace('.', '')
            .replace(',', '')
            .replace(' ', '')
            .strip() or 0
        )

    # =========================
    # THÁNG HỌC PHÍ
    # =========================
    months, years = [], []

    for i in range(1, 13):
        m = f.get(f'month{i}', '').strip()
        y = f.get(f'year{i}', '').strip()

        if m and y:
            months.append(int(m))
            years.append(int(y))

    month_label, month_codes = build_month_codes(months, years)

    # =========================
    # THI CẤP / THI ĐẲNG
    # =========================
    current_year = str(datetime.now().year)

    exam_quarter = f.get("exam_quarter", "").strip().upper()
    exam_amount = money_to_int(f.get("exam_amount"))

    dan_quarter = f.get("dan_quarter", "").strip().upper()
    dan_level = f.get("dan_level", "").strip()
    dan_amount = money_to_int(f.get("dan_amount"))

    grand_total = money_to_int(f.get('grand_total'))

    discount_type = f.get('discount_type', '').strip()
    discount_value = money_to_int(f.get('discount_value'))

    payment_method = 'CK' if f.get('chuyen_khoan') else 'TM'

    # =========================
    # CỘT ĐÓNG PHÍ
    # =========================
    fee_parts = []

    if month_label:
        fee_parts.append(f"Học phí {month_label}")

    ma_quy = ""

    # Thi cấp: Q1/Q2/Q3/Q4 + năm
    if exam_quarter:
        ma_quy = f"{exam_quarter}{current_year}"

        quarter_number = exam_quarter.replace("Q", "")
        fee_parts.append(f"Thi cấp quý {quarter_number}-{current_year}")

    # Thi đẳng: L1/L2/MN/MT/MB/QG - năm
    # Nếu có thi đẳng thì ưu tiên ma_quy thi đẳng
    if dan_quarter and dan_level:
        ma_quy = f"{dan_quarter}-{current_year}"

        dan_label_map = {
            "L1": "Lần 1",
            "L2": "Lần 2",
            "MN": "KV miền nam",
            "MT": "KV miền trung",
            "MB": "KV miền bắc",
            "QG": "Quốc Gia",
        }

        dan_quarter_label = dan_label_map.get(dan_quarter, dan_quarter)
        fee_parts.append(f"Thi đẳng {dan_level} - {dan_quarter_label}-{current_year}")

    dong_phi_text = " - ".join(fee_parts)

    # =========================
    # CỘT GHI CHÚ
    # =========================
    main_note = f.get('note', '').strip()
    auto_fee_note = f.get('auto_fee_note', '').strip()

    sub_notes = [payment_method]

    if auto_fee_note:
        sub_notes.append(auto_fee_note)

    # Dự phòng nếu frontend không gửi auto_fee_note
    if not auto_fee_note:
        if discount_type == "half_month":
            sub_notes.append("Đóng nửa tháng - giảm 50% học phí")

        elif discount_type == "percent" and discount_value > 0:
            sub_notes.append(f"Giảm giá {discount_value}%")

        elif discount_type == "money" and discount_value > 0:
            sub_notes.append(f"Giảm giá {format_money(discount_value)}")

        family_value = str(sv.get("family") or "").strip().lower()
        family_value = remove_accents(family_value)

        if family_value in ["co", "yes", "true", "1"]:
            sub_notes.append("Gia đình - giảm 10% học phí")

    sub_note_text = " - ".join(sub_notes)

    if main_note:
        note = f"{main_note} ({sub_note_text})"
    else:
        note = f"({sub_note_text})"

    payload = {
        'thoi_gian': datetime.now().isoformat(),
        'ma_hv': ma_hv,
        'ho_ten': sv.get('name', ''),
        'ngay_sinh': sv.get('birthdate', ''),
        'gioi_tinh': sv.get('gender', ''),
        'lop': sv.get('classroom', ''),
        'ca': sv.get('timeclass', ''),
        'tong_tien': grand_total,
        'thang_dong_phi': dong_phi_text,
        'ghi_chu': note,
        'ma_thang': month_codes,
        'ma_quy': ma_quy,
        'chuyen_khoan': payment_method,
    }

    supabase.table(HOCPHI_TABLE).insert(payload).execute()

    # Nếu HV đang tạm nghỉ đúng tháng vừa đóng, tự chuyển về bình thường
    clear_student_temp_leave_if_paid(ma_hv, month_codes)

    flash(f'Đã lưu học phí: {sv.get("name")} - {format_money(grand_total)}', "success")
    return back_to_current_page("fees")

@app.post("/fees/delete/<fee_id>")
def delete_fee(fee_id):
    try:
        supabase.table(HOCPHI_TABLE) \
            .delete() \
            .eq("id", fee_id) \
            .execute()

        flash("Đã xóa phiếu thu", "success")

    except Exception as e:
        flash(f"Lỗi xóa: {e}", "danger")

    return back_to_current_page("fees")


def get_student_paid_months_count(ma_hv):
    """
    Đếm số tháng học viên đã đóng học phí.
    Dựa theo bảng hocphi và cột ma_thang.
    VD ma_thang: 062026 hoặc 062026 - 072026 - 082026
    """
    try:
        ma_hv = str(ma_hv or "").strip()

        if not ma_hv:
            return 0

        rows = supabase.table(HOCPHI_TABLE) \
            .select("ma_hv,ma_thang") \
            .eq("ma_hv", ma_hv) \
            .execute().data or []

        paid_set = set()

        for r in rows:
            ma_thang = str(r.get("ma_thang") or "").strip()

            if not ma_thang:
                continue

            month_codes = [
                x.strip()
                for x in re.split(r"\s*-\s*", ma_thang)
                if x.strip()
            ]

            for code in month_codes:
                # Chỉ nhận mã tháng dạng 062026, 072026...
                if re.fullmatch(r"\d{6}", code):
                    paid_set.add(code)

        return len(paid_set)

    except Exception as e:
        print("[WELCOME POPUP PAID MONTHS ERROR]", e)
        return 0


@app.context_processor
def inject_student_welcome_popup():
    try:
        student_license = str(session.get("student_license") or "").strip()
        student_name = str(session.get("student_name") or "").strip()

        show_popup = session.pop("show_student_welcome_popup", False)

        months_count = 0
        if student_license:
            months_count = get_student_paid_months_count(student_license)

        return {
            "show_student_welcome_popup": show_popup,
            "student_name": student_name,
            "welcome_months_count": months_count,
        }

    except Exception as e:
        print("[INJECT STUDENT WELCOME POPUP ERROR]", e)

        return {
            "show_student_welcome_popup": False,
            "student_name": "",
            "welcome_months_count": 0,
        }

@app.get('/tracking')
def tracking():
    now = datetime.now()

    year = request.args.get('year', str(now.year))
    month = request.args.get('month', str(now.month))

    code = f'{int(month):02d}{year}'

    rows_all = supabase.table(HOCPHI_TABLE).select('*').execute().data or []

    rows = []
    total = 0
    cash = 0
    bank = 0

    for r in rows_all:
        ma_thang = str(r.get('ma_thang') or '').strip()

        # Chỉ xét những phiếu có đóng học phí của tháng đang lọc
        if code not in ma_thang:
            continue

        amount = int(r.get('tong_tien') or 0)

        month_codes = [
            x.strip()
            for x in re.split(r'\s*-\s*', ma_thang)
            if x.strip()
        ]

        first_month_code = month_codes[0] if month_codes else ""

        # =========================
        # Tách riêng tiền thi cấp
        # Theo form hiện tại phí thi cấp là 300.000đ
        # Tab theo dõi học phí chỉ tính tiền học phí, không tính phí thi cấp
        # =========================
        exam_fee = 0

        has_exam_fee = bool(str(r.get('ma_quy') or '').strip())

        if has_exam_fee:
            exam_fee = get_app_setting_int("fees.exam_fee", 300000)

        tuition_amount = amount - exam_fee

        if tuition_amount < 0:
            tuition_amount = 0

        # =========================
        # Quy tắc hiển thị tiền:
        # - Đóng 1 tháng: hiện tiền học phí ở tháng đó
        # - Đóng nhiều tháng 1 lần: chỉ tính tiền ở tháng đầu tiên của chuỗi
        #   các tháng sau vẫn hiện dòng nhưng số tiền = 0đ
        # =========================
        if len(month_codes) <= 1:
            display_amount = tuition_amount
        elif code == first_month_code:
            display_amount = tuition_amount
        else:
            display_amount = 0

        r['display_tong_tien'] = display_amount

        if display_amount > 0:
            total += display_amount

            if r.get('chuyen_khoan') == 'TM':
                cash += display_amount
            elif r.get('chuyen_khoan') == 'CK':
                bank += display_amount

        rows.append(r)

    rows.sort(key=lambda x: str(x.get('thoi_gian') or ''), reverse=True)

    total_count = len([r for r in rows if int(r.get('display_tong_tien') or 0) > 0])
    cash_count = len([r for r in rows if int(r.get('display_tong_tien') or 0) > 0 and r.get('chuyen_khoan') == 'TM'])
    bank_count = len([r for r in rows if int(r.get('display_tong_tien') or 0) > 0 and r.get('chuyen_khoan') == 'CK'])

    return render_template(
        'tracking.html',
        rows=rows,
        year=year,
        month=month,
        current_year=now.year,
        total=total,
        cash=cash,
        bank=bank,
        total_count=total_count,
        cash_count=cash_count,
        bank_count=bank_count
    )

def get_exam_info_web(ky_thi):
    ky_thi = str(ky_thi or "").strip()

    if not ky_thi:
        return {}

    try:
        rows = supabase.table(EXAM_INFO_TABLE) \
            .select("*") \
            .eq("ky_thi", ky_thi) \
            .limit(1) \
            .execute().data or []

        info = rows[0] if rows else {}

        judges_raw = info.get("judges_json") if info else ""

        if isinstance(judges_raw, str):
            try:
                info["judges"] = json.loads(judges_raw or "[]")
            except:
                info["judges"] = []
        elif isinstance(judges_raw, list):
            info["judges"] = judges_raw
        else:
            info["judges"] = []

        return info

    except Exception as e:
        print("[GET EXAM INFO ERROR]", e)
        return {}


def validate_exam_info_form(form):
    exam_date = str(form.get("exam_date") or "").strip()
    exam_time = str(form.get("exam_time") or "").strip()
    venue = str(form.get("venue") or "").strip()

    supervisor_name = str(form.get("supervisor_name") or "").strip()
    supervisor_active = str(form.get("supervisor_active") or "on").strip().lower()
    supervisor_active = "on" if supervisor_active == "on" else "off"

    judges_json_raw = str(form.get("judges_json") or "[]").strip()

    errors = []

    if not exam_date:
        errors.append("Ken chưa nhập <b>Ngày thi</b>.")

    if not exam_time:
        errors.append("Ken chưa nhập <b>Giờ thi</b>.")

    if not venue:
        errors.append("Ken chưa nhập <b>Địa điểm thi</b>.")

    if not supervisor_name:
        errors.append("Ken chưa nhập <b>Giám sát</b>.")

    try:
        judges = json.loads(judges_json_raw or "[]")

        if not isinstance(judges, list):
            judges = []
    except:
        judges = []

    cleaned_judges = []

    for item in judges:
        if not isinstance(item, dict):
            continue

        name = str(item.get("name") or "").strip()
        active = str(item.get("active") or "on").strip().lower()
        active = "on" if active == "on" else "off"

        if name:
            cleaned_judges.append({
                "name": name,
                "active": active,
            })

    if not cleaned_judges:
        errors.append("Ken chưa nhập <b>Giám khảo 1</b>.")

    return {
        "ok": not errors,
        "errors": errors,
        "data": {
            "exam_date": exam_date,
            "exam_time": exam_time,
            "venue": venue,
            "supervisor_name": supervisor_name,
            "supervisor_active": supervisor_active,
            "judges": cleaned_judges,
        }
    }


def save_exam_info_web(ky_thi, year, quarter, data):
    ky_thi = str(ky_thi or "").strip()
    year = str(year or "").strip()
    quarter = str(quarter or "").strip().upper()

    if not ky_thi:
        return

    payload = {
        "ky_thi": ky_thi,
        "year": year,
        "quarter": quarter,
        "exam_date": data.get("exam_date", ""),
        "exam_time": data.get("exam_time", ""),
        "venue": data.get("venue", ""),
        "supervisor_name": data.get("supervisor_name", ""),
        "supervisor_active": data.get("supervisor_active", "on"),
        "judges_json": json.dumps(data.get("judges", []), ensure_ascii=False),
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }

    existing = supabase.table(EXAM_INFO_TABLE) \
        .select("id") \
        .eq("ky_thi", ky_thi) \
        .limit(1) \
        .execute().data or []

    if existing:
        supabase.table(EXAM_INFO_TABLE) \
            .update(payload) \
            .eq("ky_thi", ky_thi) \
            .execute()
    else:
        payload["created_at"] = datetime.now(timezone.utc).isoformat()

        supabase.table(EXAM_INFO_TABLE) \
            .insert(payload) \
            .execute()


@app.post("/exam-list/save-info")
def exam_list_save_info():
    data = request.get_json(silent=True) or {}

    year = str(data.get("year") or "").strip()
    quarter = str(data.get("quarter") or "").strip().upper()

    if not year or not quarter:
        return jsonify({
            "ok": False,
            "message": "Thiếu năm hoặc quý thi."
        }), 400

    dan_quarters = ["L1", "L2", "MN", "MT", "MB", "QG"]

    if quarter in dan_quarters:
        ky_thi = f"{quarter}-{year}"
    else:
        ky_thi = f"{year}-{quarter}"

    form_like_data = {
        "exam_date": str(data.get("exam_date") or "").strip(),
        "exam_time": str(data.get("exam_time") or "").strip(),
        "venue": str(data.get("venue") or "").strip(),
        "supervisor_name": str(data.get("supervisor_name") or "").strip(),
        "supervisor_active": str(data.get("supervisor_active") or "on").strip(),
        "judges_json": json.dumps(data.get("judges") or [], ensure_ascii=False),
    }

    exam_info_check = validate_exam_info_form(form_like_data)

    if not exam_info_check["ok"]:
        return jsonify({
            "ok": False,
            "message": "Chưa đủ thông tin kỳ thi.",
            "errors": exam_info_check["errors"]
        }), 400

    try:
        save_exam_info_web(
            ky_thi,
            year,
            quarter,
            exam_info_check["data"]
        )

        return jsonify({
            "ok": True,
            "message": "Đã lưu thông tin kỳ thi lên Supabase.",
            "ky_thi": ky_thi
        })

    except Exception as e:
        print("[SAVE EXAM INFO API ERROR]", e)

        return jsonify({
            "ok": False,
            "message": f"Lỗi lưu thông tin kỳ thi: {e}"
        }), 500

def is_dan_exam_belt_web(belt):
    """
    True nếu cấp dự thi là Đẳng.
    VD: 1 Đẳng, 2 Đẳng...
    """
    belt = normalize_belt_name_web(belt)
    return "Đẳng" in belt


def build_exam_ma_quy_web(year, quarter, list_type="cap"):
    """
    Thi cấp:
        Q1, Q2, Q3, Q4  -> Q12026, Q22026...

    Thi đẳng:
        L1, L2, MN, MT, MB, QG -> L1-2026, MN-2026...
    """
    year = str(year or "").strip()
    quarter = str(quarter or "").strip().upper()
    list_type = str(list_type or "cap").strip().lower()

    dan_quarters = ["L1", "L2", "MN", "MT", "MB", "QG"]

    if list_type == "dan" or quarter in dan_quarters:
        return f"{quarter}-{year}"

    return f"{quarter}{year}"

def get_exam_list_rows_by_type_web(year, quarter, list_type="cap"):
    """
    list_type = cap  → chỉ lấy thi cấp: Cấp 9 đến Cấp 1
    list_type = dan  → chỉ lấy thi đẳng: 1 Đẳng trở lên
    """
    ma_quy = build_exam_ma_quy_web(year, quarter, list_type)

    year_str = str(year or "").strip()
    quarter_str = str(quarter or "").strip().upper()
    list_type_str = str(list_type or "cap").strip().lower()

    dan_quarters = ["L1", "L2", "MN", "MT", "MB", "QG"]

    if list_type_str == "dan" or quarter_str in dan_quarters:
        ky_thi = f"{quarter_str}-{year_str}"
    else:
        ky_thi = f"{year_str}-{quarter_str}"

    saved_result_map = {}

    try:
        saved_results = supabase.table(KETQUA_TABLE) \
            .select("ma_hv,ket_qua,so_thi,cap_dai_thi") \
            .eq("ky_thi", ky_thi) \
            .execute().data or []

        for item in saved_results:
            ma_hv_saved = str(item.get("ma_hv") or "").strip()

            if ma_hv_saved:
                saved_result_map[ma_hv_saved] = item

    except Exception as e:
        print("[GET SAVED EXAM RESULTS ERROR]", e)
        saved_result_map = {}

    fees = supabase.table(HOCPHI_TABLE) \
        .select("id,ma_hv,ma_quy,thoi_gian,coach_check_status,coach_check_note,coach_checked_by,coach_checked_by_name,coach_checked_at") \
        .eq("ma_quy", ma_quy) \
        .execute().data or []

    fee_map = {}

    for f in fees:
        ma_hv_fee = str(f.get("ma_hv") or "").strip()

        if not ma_hv_fee:
            continue

        old = fee_map.get(ma_hv_fee)

        if not old:
            fee_map[ma_hv_fee] = f
        else:
            old_time = str(old.get("thoi_gian") or "")
            new_time = str(f.get("thoi_gian") or "")

            if new_time > old_time:
                fee_map[ma_hv_fee] = f

    ids = []

    for x in fees:
        ma_hv = str(x.get("ma_hv") or "").strip()

        if ma_hv and ma_hv not in ids:
            ids.append(ma_hv)

    students = []

    if ids:
        students = supabase.table(STUDENT_TABLE) \
            .select("*") \
            .in_("license", ids) \
            .execute().data or []

    rows = []

    for s in students:
        current_belt = normalize_belt_name_web(s.get("belt"))
        license_code = str(s.get("license") or "").strip()
        saved_result = saved_result_map.get(license_code)

        # =========================
        # Nếu đã chốt kết quả rồi:
        # Lấy cấp thi đã lưu trong bảng ketqua để hiển thị lại lịch sử.
        # Không tính lại theo student.belt nữa, vì student.belt đã được nâng cấp sau khi Đạt.
        # =========================
        if saved_result:
            saved_exam_belt = normalize_belt_name_web(saved_result.get("cap_dai_thi"))

            # VD đã lưu Cấp 9 nghĩa là lúc đó học viên đang Cấp 10 và thi lên Cấp 9.
            display_current_belt = get_previous_belt_web(saved_exam_belt)

            cap_du_thi = saved_exam_belt
            current_belt_for_display = display_current_belt

        else:
            cap_du_thi = get_next_belt_web(current_belt)

            # Khóa chắc lỗi Cấp 1 -> 1 Đẳng
            if current_belt == "Cấp 1":
                cap_du_thi = "1 Đẳng"

            current_belt_for_display = current_belt

        s["belt"] = current_belt_for_display
        s["cap_du_thi"] = cap_du_thi
        s["ket_qua_mac_dinh"] = "Đạt"

        # Dữ liệu kết quả đã chốt trong bảng ketqua.
        # Bắt buộc có để exam_list/admin biết dòng nào đã lưu kết quả,
        # từ đó hiện badge "Đã lưu kết quả" thay vì còn combobox Đạt/Không đạt/Vắng.
        s["result_saved"] = bool(saved_result)
        s["saved_ket_qua"] = saved_result.get("ket_qua", "") if saved_result else ""
        s["saved_so_thi"] = saved_result.get("so_thi", "") if saved_result else ""
        s["saved_cap_thi"] = saved_result.get("cap_dai_thi", "") if saved_result else ""

        fee_check = fee_map.get(license_code, {}) or {}

        s["hocphi_id"] = fee_check.get("id", "")
        s["coach_check_status"] = fee_check.get("coach_check_status") or "Chưa KTra"
        s["coach_check_note"] = fee_check.get("coach_check_note") or ""
        s["coach_checked_by"] = fee_check.get("coach_checked_by") or ""
        s["coach_checked_by_name"] = fee_check.get("coach_checked_by_name") or ""
        s["coach_checked_at"] = fee_check.get("coach_checked_at") or ""
        is_dan = is_dan_exam_belt_web(cap_du_thi)

        if list_type == "dan" and is_dan:
            rows.append(s)

        elif list_type == "cap" and not is_dan:
            rows.append(s)

    return rows

@app.get('/exam-list')
def exam_list():
    year = request.args.get('year', str(datetime.now().year))
    quarter = request.args.get('quarter', f'Q{(datetime.now().month - 1)//3 + 1}').strip().upper()

    ky_thi = f"{year}-{quarter}"

    students = get_exam_list_rows_by_type_web(
        year=year,
        quarter=quarter,
        list_type="cap"
    )

    finalize_errors = session.pop("exam_finalize_errors", [])
    exam_info = get_exam_info_web(ky_thi)
    saved_count = len([x for x in students if x.get("result_saved")])
    all_saved = bool(students) and saved_count == len(students)

    return render_template(
        'exam_list.html',
        rows=students,
        year=year,
        quarter=quarter,
        ky_thi=ky_thi,
        finalize_errors=finalize_errors,
        exam_info=exam_info,
        list_type="cap",
        saved_count=saved_count,
        all_saved=all_saved
    )

@app.get('/dan-list')
def dan_list():
    year = request.args.get('year', str(datetime.now().year))
    quarter = request.args.get('quarter', 'L1').strip().upper()

    ky_thi = f"{quarter}-{year}"

    students = get_exam_list_rows_by_type_web(
        year=year,
        quarter=quarter,
        list_type="dan"
    )

    finalize_errors = session.pop("exam_finalize_errors", [])
    exam_info = get_exam_info_web(ky_thi)
    saved_count = len([x for x in students if x.get("result_saved")])
    all_saved = bool(students) and saved_count == len(students)

    return render_template(
        'dan_list.html',
        rows=students,
        year=year,
        quarter=quarter,
        ky_thi=ky_thi,
        finalize_errors=finalize_errors,
        exam_info=exam_info,
        saved_count=saved_count,
        all_saved=all_saved,
        list_type="dan"
    )

def normalize_excel_header_web(value):
    """
    Chuẩn hóa tên cột Excel:
    VD: "Mã HV" -> "ma hv"
    """
    value = remove_accents(str(value or "").strip()).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def find_excel_col_index_web(headers, candidates):
    """
    Tìm vị trí cột theo nhiều tên có thể có.
    headers: list text đã chuẩn hóa
    candidates: list tên cột có thể gặp
    """
    normalized_candidates = [
        normalize_excel_header_web(x)
        for x in candidates
    ]

    for idx, header in enumerate(headers):
        if header in normalized_candidates:
            return idx

    return None


def extract_exam_number_web(raw):
    """
    Lấy số thi từ Excel.
    Nhận:
    - 12
    - 012
    - Cấp 5_12
    - Cap 5 - 12
    """
    raw = str(raw or "").strip()

    if not raw:
        return ""

    # Nếu Excel đọc số dạng 12.0 thì đưa về 12
    if re.fullmatch(r"\d+\.0", raw):
        raw = raw.split(".")[0]

    # Lấy cụm số cuối cùng trong chuỗi
    nums = re.findall(r"\d+", raw)

    if not nums:
        return ""

    return str(int(nums[-1]))


def normalize_exam_period_web(raw):
    """
    Chuẩn hóa kỳ thi từ Excel để so với trang hiện tại.

    Nhận các dạng:
    - Q12026
    - Q1-2026
    - Q1/2026
    - 2026-Q1
    - L1-2026
    - 2026-L1
    """
    raw = str(raw or "").strip().upper()

    if not raw:
        return ""

    raw = raw.replace("_", "-").replace("/", "-").replace(" ", "")

    # Thi cấp: Q12026 / Q1-2026 / Q1.2026
    m = re.search(r"Q([1-4])\D*(20\d{2})", raw)
    if m:
        return f"{m.group(2)}-Q{m.group(1)}"

    # Thi cấp: 2026-Q1 / 2026.Q1
    m = re.search(r"(20\d{2})\D*Q([1-4])", raw)
    if m:
        return f"{m.group(1)}-Q{m.group(2)}"

    # Thi đẳng: L1-2026 / MN-2026...
    dan_codes = "L1|L2|MN|MT|MB|QG"

    m = re.search(rf"({dan_codes})\D*(20\d{{2}})", raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}"

    m = re.search(rf"(20\d{{2}})\D*({dan_codes})", raw)
    if m:
        return f"{m.group(2)}-{m.group(1)}"

    return raw


def current_exam_period_web(year, quarter, list_type):
    year = str(year or "").strip()
    quarter = str(quarter or "").strip().upper()
    list_type = str(list_type or "cap").strip().lower()

    dan_quarters = ["L1", "L2", "MN", "MT", "MB", "QG"]

    if list_type == "dan" or quarter in dan_quarters:
        return f"{quarter}-{year}"

    return f"{year}-{quarter}"




@app.post("/exam-list/import-numbers")
def exam_list_import_numbers():
    """
    Upload Excel để tự load Số thi vào danh sách đang lọc.

    Kiểm tra:
    1. Đúng năm/quý đang mở.
    2. Đúng Mã HV có trong danh sách hiện tại.
    3. Đúng Cấp dự thi nếu Excel có cột Cấp dự thi.
    4. Không trùng mã HV trong Excel.
    5. Không trùng số thi giữa các dòng được áp dụng.
    """
    year = request.form.get("year", str(datetime.now().year)).strip()
    quarter = request.form.get("quarter", "").strip().upper()
    list_type = request.form.get("list_type", "cap").strip().lower()

    back_period = current_exam_period_web(year, quarter, list_type)

    file = request.files.get("excel_file")

    if not file or not file.filename:
        return jsonify({
            "ok": False,
            "message": "Ken chưa chọn file Excel."
        }), 400

    filename = secure_filename(file.filename or "")
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if ext not in ["xlsx", "xlsm"]:
        return jsonify({
            "ok": False,
            "message": "File phải là Excel .xlsx hoặc .xlsm."
        }), 400

    try:
        current_rows = get_exam_list_rows_by_type_web(
            year=year,
            quarter=quarter,
            list_type=list_type
        )

        current_map = {}

        for r in current_rows:
            license_code = str(r.get("license") or "").strip()

            if not license_code:
                continue

            current_map[license_code] = {
                "license": license_code,
                "name": str(r.get("name") or "").strip(),
                "cap_du_thi": normalize_belt_name_web(r.get("cap_du_thi")),
            }

        if not current_map:
            return jsonify({
                "ok": False,
                "message": "Danh sách hiện tại chưa có học viên để cập nhật số thi."
            }), 400

        wb = load_workbook(file, data_only=True)
        ws = wb.active

        if ws.max_row < 2:
            return jsonify({
                "ok": False,
                "message": "File Excel chưa có dữ liệu."
            }), 400

        raw_headers = [
            ws.cell(row=1, column=col).value
            for col in range(1, ws.max_column + 1)
        ]

        headers = [
            normalize_excel_header_web(x)
            for x in raw_headers
        ]

        col_ma_hv = find_excel_col_index_web(headers, [
            "Mã HV", "Ma HV",
            "Mã hội viên", "Ma hoi vien",
            "Mã học viên", "Ma hoc vien",
            "Mã võ sinh", "Ma vo sinh",
            "Mã thi sinh", "Ma thi sinh",
            "Mã thí sinh", "Ma thi sinh",
            "Mã dự thi", "Ma du thi",
            "License",
            "Student code",
            "Student ID",
            "ID",
            "ma_hv",
            "mahv",
            "license"
        ])

        col_so_thi = find_excel_col_index_web(headers, [
            "Số thi", "So thi",
            "Số báo danh", "So bao danh",
            "Số dự thi", "So du thi",
            "Số thứ tự thi", "So thu tu thi",
            "Mã số thi", "Ma so thi",
            "Mã số dự thi", "Ma so du thi",
            "SBD",
            "Exam number",
            "Exam No",
            "Candidate number",
            "Contest number",
            "Bib number",
            "so_thi",
            "sothi",
            "so_bao_danh",
            "sbd"
        ])

        col_cap_du_thi = find_excel_col_index_web(headers, [
            "Cấp dự thi", "Cap du thi",
            "Cấp thi", "Cap thi",
            "Cấp", "Cap",
            "Đẳng dự thi", "Dang du thi",
            "Đẳng thi", "Dang thi",
            "Belt",
            "Exam belt",
            "cap_du_thi",
            "cap_thi"
        ])

        col_ky_thi = find_excel_col_index_web(headers, [
            "Quý thi", "Quy thi",
            "Kỳ thi", "Ky thi",
            "Quý", "Quy",
            "Lần thi", "Lan thi",
            "Mã quý", "Ma quy",
            "Mã kỳ thi", "Ma ky thi",
            "ma_quy",
            "ky_thi"
        ])

        if col_ky_thi is None or col_ma_hv is None or col_cap_du_thi is None or col_so_thi is None:
            return jsonify({
                "ok": False,
                "message": "Excel bắt buộc phải có đủ các cột: Quý thi, Mã HV/Mã thí sinh, Cấp dự thi, Số thi/Số báo danh.",
                "errors": [
                    "Thiếu cột Quý thi." if col_ky_thi is None else "",
                    "Thiếu cột Mã HV/Mã thí sinh." if col_ma_hv is None else "",
                    "Thiếu cột Cấp dự thi." if col_cap_du_thi is None else "",
                    "Thiếu cột Số thi/Số báo danh." if col_so_thi is None else "",
                ],
                "numbers": {}
            }), 400

        numbers = {}
        errors = []
        outside_students = []
        seen_license = set()
        seen_number = {}

        for row_idx in range(2, ws.max_row + 1):
            values = [
                ws.cell(row=row_idx, column=col).value
                for col in range(1, ws.max_column + 1)
            ]

            ma_hv = str(values[col_ma_hv] or "").strip()
            so_thi = extract_exam_number_web(values[col_so_thi])

            if not ma_hv and not so_thi:
                continue

            excel_period = normalize_exam_period_web(values[col_ky_thi])
            excel_cap = normalize_belt_name_web(values[col_cap_du_thi])

            if not excel_period:
                errors.append(f"Dòng {row_idx}: thiếu Quý thi.")
                continue

            if not ma_hv:
                errors.append(f"Dòng {row_idx}: thiếu Mã HV/Mã thí sinh.")
                continue

            if not excel_cap:
                errors.append(f"{ma_hv}: thiếu Cấp dự thi.")
                continue

            if not so_thi:
                errors.append(f"{ma_hv}: thiếu Số thi/Số báo danh.")
                continue

            if ma_hv in seen_license:
                errors.append(f"{ma_hv}: bị trùng Mã HV trong file Excel.")
                continue

            seen_license.add(ma_hv)

            # =========================
            # KIỂM TRA TRÙNG SỐ THI TRÊN TOÀN BỘ FILE EXCEL
            # Làm trước khi kiểm tra mã có trong hệ thống hay không.
            # Vì VĐV ngoài CLB vẫn có thể bị trùng số thi với học viên của mình.
            # =========================
            excel_number_key = f"{excel_cap}_{so_thi}"

            if excel_number_key in seen_number:
                errors.append(
                    f"Số thi {excel_cap}_{so_thi} bị trùng giữa {seen_number[excel_number_key]} và {ma_hv}."
                )
                continue

            seen_number[excel_number_key] = ma_hv

            current_student = current_map.get(ma_hv)

            if not current_student:
                outside_students.append(
                    f"{ma_hv}: không có trong danh sách đang lọc, có thể là VĐV/CLB khác thi chung."
                )
                continue

            if excel_period != back_period:
                errors.append(
                    f"{ma_hv}: sai Quý thi. Excel là {excel_period}, hệ thống là {back_period}."
                )
                continue

            system_cap = current_student["cap_du_thi"]

            if excel_cap != system_cap:
                errors.append(
                    f"{ma_hv}: sai Cấp dự thi. Excel là {excel_cap}, hệ thống là {system_cap}."
                )
                continue

            numbers[ma_hv] = so_thi

        if errors:
            return jsonify({
                "ok": False,
                "message": "File Excel có lỗi cần xử lý trước.",
                "errors": errors,
                "outside_students": outside_students,
                "numbers": {}
            }), 400

        if not numbers:
            return jsonify({
                "ok": False,
                "message": "Không có dòng nào thuộc danh sách hiện tại để cập nhật số thi.",
                "errors": errors,
                "outside_students": outside_students,
                "numbers": {}
            }), 400

        return jsonify({
            "ok": True,
            "message": f"Đã dò xong Excel. Có thể cập nhật {len(numbers)} số thi.",
            "applied_count": len(numbers),
            "outside_count": len(outside_students),
            "numbers": numbers,
            "outside_students": outside_students,
            "require_confirm": len(outside_students) > 0,
            "errors": []
        })

    except Exception as e:
        print("[IMPORT EXAM NUMBERS ERROR]", e)

        return jsonify({
            "ok": False,
            "message": f"Lỗi đọc Excel: {e}"
        }), 500

@app.post('/exam-list/finalize')
def exam_list_finalize():
    year = request.form.get("year", str(datetime.now().year)).strip()
    quarter = request.form.get("quarter", f'Q{(datetime.now().month - 1)//3 + 1}').strip().upper()
    list_type = request.form.get("list_type", "cap").strip()

    dan_quarters = ["L1", "L2", "MN", "MT", "MB", "QG"]

    if list_type == "dan" or quarter in dan_quarters:
        ky_thi = f"{quarter}-{year}"
    else:
        ky_thi = f"{year}-{quarter}"

    back_endpoint = "dan_list" if list_type == "dan" else "exam_list"

    exam_info_check = validate_exam_info_form(request.form)

    if not exam_info_check["ok"]:
        session["exam_finalize_errors"] = [
            "<b>Chưa đủ thông tin kỳ thi</b><br>" + "<br>".join(exam_info_check["errors"])
        ]

        return redirect(url_for(
            back_endpoint,
            year=year,
            quarter=quarter
        ))

    current_ky_value = parse_ky_thi_web(ky_thi)

    ma_hv_list = request.form.getlist("ma_hv")

    payloads = []
    pass_updates = []
    errors = []
    seen_ma_hv = set()

    for ma_hv in ma_hv_list:
        ma_hv = str(ma_hv or "").strip()

        if not ma_hv:
            continue

        if ma_hv in seen_ma_hv:
            errors.append(format_result_error_web(ma_hv, "", "Mã hội viên bị trùng trong danh sách đang chốt."))
            continue

        seen_ma_hv.add(ma_hv)

        ho_ten = request.form.get(f"name_{ma_hv}", "").strip()
        cap_du_thi = request.form.get(f"cap_du_thi_{ma_hv}", "").strip()
        so_thi_raw = request.form.get(f"so_thi_{ma_hv}", "").strip()
        ket_qua = request.form.get(f"ket_qua_{ma_hv}", "Đạt").strip()

        so_thi = ""
        if so_thi_raw:
            exam_prefix = str(get_app_setting("exam.exam_number_prefix", "Cấp_") or "")

            # Cách dùng trong Setup:
            # "Cấp_"   -> Cấp_01
            # "{cap}_" -> Cấp 5_01
            # ""       -> 01
            exam_prefix = exam_prefix.replace("{cap}", cap_du_thi)

            so_thi = f"{exam_prefix}{so_thi_raw}"

        # 1) Lấy thông tin học viên hiện tại
        student_rows = safe_rows(STUDENT_TABLE, "*", license=ma_hv)
        student = student_rows[0] if student_rows else {}

        if not student:
            errors.append(format_result_error_web(ma_hv, ho_ten, "Không tìm thấy học viên trong bảng student."))
            continue

        current_student_belt = str(student.get("belt") or "").strip()

        # 2) Lấy toàn bộ lịch sử thi cấp của học viên
        all_results = safe_rows(KETQUA_TABLE, "*", ma_hv=ma_hv)

        # 3) Mỗi quý chỉ được chốt 1 lần / 1 mã HV
        same_quarter_results = [
            r for r in all_results
            if str(r.get("ky_thi") or "").strip().upper() == ky_thi.upper()
        ]

        if same_quarter_results:
            errors.append(
                f"""
                <b>{ma_hv} - {ho_ten}</b><br>
                Lý do: Học viên này đã có kết quả trong kỳ <b>{ky_thi}</b>.<br>
                Cách xử lý: Nếu Ken muốn chốt lại, hãy vào tab <b>Kết quả thi cấp</b>,
                click phải dòng kết quả cũ và xóa trước. Sau đó quay lại chốt lại.
                """
            )
            continue

        # 4) Chỉ xét các kết quả trước quý hiện tại
        previous_results = [
            r for r in all_results
            if parse_ky_thi_web(r.get("ky_thi")) < current_ky_value
        ]

        future_or_invalid_results = [
            r for r in all_results
            if parse_ky_thi_web(r.get("ky_thi")) > current_ky_value
        ]

        if future_or_invalid_results:
            errors.append(
                f"""
                <b>{ma_hv} - {ho_ten}</b><br>
                Lý do: Học viên này đã có kết quả ở quý sau <b>{ky_thi}</b>.<br>
                Cách xử lý: Không thể chốt ngược lại kỳ cũ khi đã có dữ liệu kỳ sau.
                Ken cần kiểm tra lại lịch sử trong tab <b>Kết quả thi cấp</b>.
                """
            )
            continue

        # 5) Kiểm tra đúng trình tự cấp thi
        expected_belt = get_expected_exam_belt_web(current_student_belt, previous_results)

        if cap_du_thi != expected_belt:
            errors.append(
                f"""
                <b>{ma_hv} - {ho_ten}</b><br>
                Lý do: Sai trình tự cấp thi.<br>
                Cấp được phép thi hiện tại: <b>{expected_belt}</b><br>
                Cấp đang chốt: <b>{cap_du_thi}</b><br>
                Quy tắc: Nếu quý trước <b>Đạt</b> thì quý sau được thi cấp tiếp theo.
                Nếu <b>Không đạt</b> hoặc <b>Vắng</b> thì quý sau phải thi lại cấp đó.
                """
            )
            continue

        ghi_chu = ""

        if ket_qua == "Vắng":
            ghi_chu = "Vắng thi"
        elif ket_qua == "Không đạt":
            ghi_chu = "Không đạt"

        payloads.append({
            "ky_thi": ky_thi,
            "ma_hv": ma_hv,
            "ho_ten": ho_ten,
            "cap_dai_thi": cap_du_thi,
            "so_thi": so_thi,
            "ket_qua": ket_qua,
            "ghi_chu": ghi_chu,
        })

        # 6) Nếu đạt thì cập nhật cấp mới cho student sau khi insert thành công
        if ket_qua == "Đạt":
            pass_updates.append({
                "ma_hv": ma_hv,
                "new_belt": cap_du_thi,
            })

    # Nếu có lỗi thì không chốt bất kỳ ai, tránh dữ liệu nửa đúng nửa sai
    if errors:
        session["exam_finalize_errors"] = errors

        return redirect(url_for(
            back_endpoint,
            year=year,
            quarter=quarter
        ))

    if not payloads:
        flash("Không có dữ liệu hợp lệ để chốt kết quả.")
        return redirect(url_for(
            back_endpoint,
            year=year,
            quarter=quarter
        ))

    try:
        # Lưu thông tin kỳ thi trước khi chốt kết quả
        save_exam_info_web(ky_thi, year, quarter, exam_info_check["data"])

        # Không xóa kết quả cũ nữa.
        # Vì mỗi mã HV / mỗi quý chỉ được chốt 1 lần.
        supabase.table(KETQUA_TABLE).insert(payloads).execute()

        # Cập nhật cấp mới cho học viên đạt
        for item in pass_updates:
            supabase.table(STUDENT_TABLE) \
                .update({"belt": item["new_belt"]}) \
                .eq("license", item["ma_hv"]) \
                .execute()

        flash(
            f"Đã chốt kết quả kỳ thi {ky_thi}: "
            f"{len(payloads)} võ sinh. "
            f"Đã cập nhật cấp mới cho {len(pass_updates)} võ sinh đạt."
        )

    except Exception as e:
        print("[FINALIZE EXAM ERROR]", e)
        flash(f"Lỗi chốt kết quả: {e}")

        return redirect(url_for(
            back_endpoint,
            year=year,
            quarter=quarter
        ))

    return redirect(url_for("results", year=year, quarter=quarter))

@app.get('/api/student/<license_code>')
def api_student(license_code):
    data = supabase.table(STUDENT_TABLE).select('*').eq('license', license_code).limit(1).execute().data
    return jsonify(data[0] if data else {})

STUDENT_TABLE = "student"
COACH_TABLE = "coaches"
HOCPHI_TABLE = "hocphi"
KETQUA_TABLE = "ketqua"
HOATDONG_TABLE = "hoatdong"
NOTIFICATION_TABLE = "notifications"
PAYMENT_SETTINGS_TABLE = "payment_settings"
APP_SETTINGS_TABLE = "app_settings"
NOTIFICATION_READ_TABLE = "notification_reads"
EXAM_INFO_TABLE = "exam_infos"
BELT_FLOW_WEB = [
    "Cấp 10", "Cấp 9", "Cấp 8", "Cấp 7", "Cấp 6",
    "Cấp 5", "Cấp 4", "Cấp 3", "Cấp 2", "Cấp 1",
    "1 Đẳng", "2 Đẳng", "3 Đẳng", "4 Đẳng", "5 Đẳng",
    "6 Đẳng", "7 Đẳng", "8 Đẳng", "9 Đẳng", "10 Đẳng"
]

def normalize_belt_name_web(belt):
    raw = str(belt or "").strip()

    if not raw:
        return ""

    # Chuẩn hóa Unicode: xử lý lỗi Cấp khác Cấp
    raw = unicodedata.normalize("NFC", raw)

    # Chuẩn hóa khoảng trắng
    raw = re.sub(r"\s+", " ", raw).strip()

    # Bản bỏ dấu để nhận mọi kiểu: Cấp, Cấp, cap, CẤP
    text_no_accent = remove_accents(raw).lower()
    text_no_accent = re.sub(r"\s+", " ", text_no_accent).strip()

    # Chuẩn hóa Cấp: Cấp 1, Cấp 1, cap 1, CAP 1 -> Cấp 1
    cap_match = re.search(r"cap\s*(\d+)", text_no_accent)
    if cap_match:
        return f"Cấp {int(cap_match.group(1))}"

    # Chuẩn hóa Đẳng: 1 đẳng, 1 Đẳng, 1 dang, 1 DANG -> 1 Đẳng
    dang_match = re.search(r"(\d+)\s*dang", text_no_accent)
    if dang_match:
        return f"{int(dang_match.group(1))} Đẳng"

    return raw

def supabase_public_url(bucket, filename):
    filename = str(filename or "").strip().lstrip("/")
    return f"{SUPABASE_URL}/storage/v1/object/public/{bucket}/{filename}"


def supabase_student_photo_public_url(filename):
    filename = str(filename or "").strip().lstrip("/")
    return supabase_public_url(STUDENT_PHOTO_BUCKET, filename)


def upload_club_asset_to_supabase(file_storage, folder="club-info"):
    """
    Upload ảnh logo/HLV/thông tin CLB lên Supabase Storage.
    """
    if not file_storage or not file_storage.filename:
        return ""

    filename = secure_filename(file_storage.filename)
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if ext not in ALLOWED_PHOTO_EXTENSIONS:
        raise ValueError("Ảnh chỉ nhận file JPG, JPEG, PNG hoặc WEBP.")

    storage_path = f"{folder}/{datetime.now().strftime('%Y%m%d%H%M%S%f')}.{ext}"
    content = file_storage.read()

    content_type_map = {
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "png": "image/png",
        "webp": "image/webp",
    }

    try:
        supabase.storage.from_(CLUB_ASSET_BUCKET).upload(
            storage_path,
            content,
            {
                "content-type": content_type_map.get(ext, "application/octet-stream"),
                "upsert": "true"
            }
        )
    except Exception:
        supabase.storage.from_(CLUB_ASSET_BUCKET).update(
            storage_path,
            content,
            {
                "content-type": content_type_map.get(ext, "application/octet-stream"),
                "upsert": "true"
            }
        )

    return supabase_public_url(CLUB_ASSET_BUCKET, storage_path)

def prepare_student_photo_for_upload(file_storage, target_width=780, target_height=1040, quality=98):
    """
    Xử lý ảnh học viên trước khi upload Supabase:
    - Sửa xoay EXIF
    - Convert RGB
    - Crop đúng tỷ lệ 3x4
    - Resize vừa đủ để hiển thị web, tránh bị sharpen quá mạnh
    - Làm mềm rất nhẹ để giảm lốm đốm vùng da/áo
    - Xuất JPEG chất lượng cao
    """

    img = Image.open(file_storage.stream)
    img = ImageOps.exif_transpose(img)
    img = img.convert("RGB")

    src_w, src_h = img.size
    target_ratio = target_width / target_height
    src_ratio = src_w / src_h

    # Crop đúng tỷ lệ 3x4
    if src_ratio > target_ratio:
        new_w = int(src_h * target_ratio)
        left = (src_w - new_w) // 2
        img = img.crop((left, 0, left + new_w, src_h))

    elif src_ratio < target_ratio:
        new_h = int(src_w / target_ratio)

        # Ưu tiên giữ mặt và phần đầu, không cắt quá thấp
        top = max(0, int((src_h - new_h) * 0.10))

        if top + new_h > src_h:
            top = src_h - new_h

        img = img.crop((0, top, src_w, top + new_h))

    # Dùng BICUBIC mềm hơn LANCZOS, tránh bị rít nét
    img = img.resize((target_width, target_height), Image.Resampling.BICUBIC)

    # Làm mềm cực nhẹ, giảm lốm đốm nhưng không mất chi tiết mặt
    img = img.filter(ImageFilter.GaussianBlur(radius=0.18))

    output = BytesIO()
    img.save(
        output,
        format="JPEG",
        quality=quality,
        optimize=True,
        progressive=False,
        subsampling=0
    )
    output.seek(0)

    return output.getvalue()

def get_student_photo_url(license_code):
    license_code = str(license_code or "").strip()

    if not license_code:
        return ""

    try:
        rows = supabase.table(STUDENT_TABLE) \
            .select("photo_url") \
            .eq("license", license_code) \
            .limit(1) \
            .execute().data or []

        if rows and rows[0].get("photo_url"):
            return rows[0].get("photo_url")

    except Exception as e:
        print("[GET STUDENT PHOTO URL ERROR]", e)

    return ""

def rename_student_photo_file(old_license, new_license):
    old_license = str(old_license or "").strip()
    new_license = str(new_license or "").strip()

    if not old_license or not new_license or old_license == new_license:
        return

    for ext in ["jpg", "jpeg", "png", "webp"]:
        old_path = STUDENT_PHOTO_DIR / f"{old_license}.{ext}"
        new_path = STUDENT_PHOTO_DIR / f"{new_license}.{ext}"

        if old_path.exists():
            if new_path.exists():
                new_path.unlink()
            old_path.rename(new_path)
            break

@app.get("/api/student-detail/<license_code>")
def api_student_detail(license_code):
    license_code = str(license_code or "").strip()

    student_rows = safe_rows(STUDENT_TABLE, "*", license=license_code)
    student = student_rows[0] if student_rows else {}

    results = safe_rows(KETQUA_TABLE, "*", ma_hv=license_code)
    activities = safe_rows(HOATDONG_TABLE, "*", ma_hv=license_code)

    results.sort(key=lambda r: str(r.get("ky_thi") or ""), reverse=True)
    activities.sort(key=lambda r: str(r.get("thoi_gian") or ""), reverse=True)

    return jsonify({
        "student": student,
        "results": results,
        "activities": activities,
        "photo_url": get_student_photo_url(license_code)
    })

@app.post("/api/student-photo-upload/<license_code>")
def api_student_photo_upload(license_code):
    license_code = str(license_code or "").strip()

    if not license_code:
        return jsonify({"ok": False, "message": "Thiếu mã hội viên"}), 400

    if "photo" not in request.files:
        return jsonify({"ok": False, "message": "Chưa chọn hình"}), 400

    file = request.files["photo"]

    if not file or not file.filename:
        return jsonify({"ok": False, "message": "File hình không hợp lệ"}), 400

    filename = secure_filename(file.filename)
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if ext not in ALLOWED_PHOTO_EXTENSIONS:
        return jsonify({
            "ok": False,
            "message": "Chỉ cho phép ảnh JPG, JPEG, PNG hoặc WEBP"
        }), 400

    try:
        # Luôn xử lý và lưu ảnh học viên thành JPG chuẩn 3x4
        storage_path = f"{license_code}.jpg"
        content = prepare_student_photo_for_upload(file)

        # Xóa toàn bộ ảnh cũ khác đuôi để mỗi học viên chỉ còn 1 ảnh
        old_paths = [
            f"{license_code}.jpg",
            f"{license_code}.jpeg",
            f"{license_code}.png",
            f"{license_code}.webp",
        ]

        try:
            supabase.storage.from_(STUDENT_PHOTO_BUCKET).remove(old_paths)
        except Exception:
            pass

        # Upload mới / ghi đè ảnh cũ
        try:
            supabase.storage.from_(STUDENT_PHOTO_BUCKET).upload(
                storage_path,
                content,
                {
                    "content-type": "image/jpeg",
                    "upsert": "true"
                }
            )
        except Exception:
            supabase.storage.from_(STUDENT_PHOTO_BUCKET).update(
                storage_path,
                content,
                {
                    "content-type": "image/jpeg",
                    "upsert": "true"
                }
            )

        photo_url = supabase_student_photo_public_url(storage_path)

        # Thêm version chống cache để web hiện ảnh mới ngay
        photo_url_with_version = f"{photo_url}?v={int(datetime.now().timestamp())}"

        # Lưu link ảnh vào bảng student
        supabase.table(STUDENT_TABLE) \
            .update({"photo_url": photo_url_with_version}) \
            .eq("license", license_code) \
            .execute()

        return jsonify({
            "ok": True,
            "photo_url": photo_url_with_version
        })

    except Exception as e:
        print("[UPLOAD STUDENT PHOTO SUPABASE ERROR]", repr(e))
        return jsonify({"ok": False, "message": str(e)}), 500

def sync_student_profile_to_related_tables(license_code):
    """
    Đồng bộ thông tin học viên từ bảng student sang các bảng liên quan.
    student là bảng gốc.
    """

    license_code = str(license_code or "").strip()

    if not license_code:
        return

    try:
        rows = supabase.table(STUDENT_TABLE) \
            .select("*") \
            .eq("license", license_code) \
            .limit(1) \
            .execute().data or []

        if not rows:
            return

        s = rows[0]

        name = str(s.get("name") or "").strip()
        birthdate = str(s.get("birthdate") or "").strip()
        gender = str(s.get("gender") or "").strip()
        classroom = str(s.get("classroom") or "").strip()
        timeclass = str(s.get("timeclass") or "").strip()
        clup = str(s.get("clup") or "").strip()
        phone = str(s.get("phonenumber") or "").strip()

        # =========================
        # 1. HỌC PHÍ
        # hocphi đang lưu snapshot: họ tên, ngày sinh, giới tính, lớp, ca
        # =========================
        try:
            supabase.table(HOCPHI_TABLE) \
                .update({
                    "ho_ten": name,
                    "ngay_sinh": birthdate,
                    "gioi_tinh": gender,
                    "lop": classroom,
                    "ca": timeclass,
                }) \
                .eq("ma_hv", license_code) \
                .execute()
        except Exception as e:
            print("[SYNC HOCPHI PROFILE ERROR]", e)

        # =========================
        # 2. KẾT QUẢ THI CẤP
        # ketqua đang lưu ho_ten
        # =========================
        try:
            supabase.table(KETQUA_TABLE) \
                .update({
                    "ho_ten": name,
                }) \
                .eq("ma_hv", license_code) \
                .execute()
        except Exception as e:
            print("[SYNC KETQUA PROFILE ERROR]", e)

        # =========================
        # 3. HOẠT ĐỘNG
        # hoatdong đang lưu ho_ten
        # =========================
        try:
            supabase.table(HOATDONG_TABLE) \
                .update({
                    "ho_ten": name,
                }) \
                .eq("ma_hv", license_code) \
                .execute()
        except Exception as e:
            print("[SYNC HOATDONG PROFILE ERROR]", e)

        # =========================
        # 4. THÔNG BÁO GỬI RIÊNG
        # notifications đang lưu target_name
        # =========================
        try:
            supabase.table(NOTIFICATION_TABLE) \
                .update({
                    "target_name": name,
                }) \
                .eq("target_license", license_code) \
                .execute()
        except Exception as e:
            print("[SYNC NOTIFICATION TARGET NAME ERROR]", e)

    except Exception as e:
        print("[SYNC STUDENT PROFILE ERROR]", e)

@app.post("/api/student-field-update")
def api_student_field_update():
    data = request.get_json(silent=True) or {}

    license_code = str(data.get("license") or "").strip()
    field = str(data.get("field") or "").strip()
    value = str(data.get("value") or "").strip()

    allowed_fields = {
        "name",
        "birthdate",
        "gender",
        "phonenumber",
        "address",
        "telegram_id",
        "license",
        "classroom",
        "timeclass",
        "clup",
        "belt",
        "family",
        "active",
    }

    if not license_code or not field:
        return jsonify({
            "ok": False,
            "message": "Thiếu mã hội viên hoặc trường cần sửa"
        }), 400

    if field not in allowed_fields:
        return jsonify({
            "ok": False,
            "message": "Trường này không được phép sửa"
        }), 400

    try:
        # =========================
        # ĐỔI MÃ HỘI VIÊN TOÀN HỆ THỐNG
        # =========================
        if field == "license":
            old_license = license_code.strip()
            new_license = value.strip()

            if not new_license:
                return jsonify({
                    "ok": False,
                    "message": "Mã hội viên không được trống"
                }), 400

            if new_license == old_license:
                sync_student_profile_to_related_tables(old_license)

                return jsonify({
                    "ok": True,
                    "new_license": new_license,
                    "message": "Mã hội viên không thay đổi. Đã kiểm tra đồng bộ dữ liệu liên quan."
                })

            # Kiểm tra mã mới có trùng không
            duplicated = supabase.table(STUDENT_TABLE) \
                .select("license,name,birthdate") \
                .eq("license", new_license) \
                .limit(1) \
                .execute().data or []

            if duplicated:
                old = duplicated[0]

                return jsonify({
                    "ok": False,
                    "message": (
                        f"Mã HV {new_license} đã tồn tại cho học viên "
                        f"{old.get('name', '')} - {old.get('birthdate', '')}. "
                        f"Không thể đổi sang mã bị trùng."
                    )
                }), 400

            # Lấy học viên hiện tại
            student_rows = supabase.table(STUDENT_TABLE) \
                .select("*") \
                .eq("license", old_license) \
                .limit(1) \
                .execute().data or []

            if not student_rows:
                return jsonify({
                    "ok": False,
                    "message": f"Không tìm thấy mã HV cũ: {old_license}"
                }), 404

            old_student = student_rows[0]

            student_update = {
                "license": new_license
            }

            old_username = str(old_student.get("portal_username") or "").strip()

            # Nếu username portal đang là mã cũ thì đổi sang mã mới
            if not old_username or old_username == old_license:
                student_update["portal_username"] = new_license

            # 1. Đổi mã trong student trước
            supabase.table(STUDENT_TABLE) \
                .update(student_update) \
                .eq("license", old_license) \
                .execute()

            # 2. Đổi mã trong các bảng liên quan
            supabase.table(HOCPHI_TABLE) \
                .update({"ma_hv": new_license}) \
                .eq("ma_hv", old_license) \
                .execute()

            supabase.table(KETQUA_TABLE) \
                .update({"ma_hv": new_license}) \
                .eq("ma_hv", old_license) \
                .execute()

            supabase.table(HOATDONG_TABLE) \
                .update({"ma_hv": new_license}) \
                .eq("ma_hv", old_license) \
                .execute()

            try:
                supabase.table(NOTIFICATION_TABLE) \
                    .update({"target_license": new_license}) \
                    .eq("target_license", old_license) \
                    .execute()
            except Exception as e:
                print("[UPDATE NOTIFICATION TARGET LICENSE ERROR]", e)

            try:
                supabase.table(NOTIFICATION_READ_TABLE) \
                    .update({"student_license": new_license}) \
                    .eq("student_license", old_license) \
                    .execute()
            except Exception as e:
                print("[UPDATE NOTIFICATION READ LICENSE ERROR]", e)

            # 3. Sau khi đổi mã xong, lấy student làm gốc và đồng bộ lại tên/ngày sinh/lớp/ca...
            sync_student_profile_to_related_tables(new_license)

            return jsonify({
                "ok": True,
                "new_license": new_license,
                "message": (
                    f"Đã đổi mã HV từ {old_license} sang {new_license}. "
                    f"Đã đồng bộ lại học phí, thi cấp, hoạt động và thông báo."
                )
            })

        # =========================
        # CẬP NHẬT CÁC TRƯỜNG KHÁC
        # Sau khi sửa student, quét lại các bảng liên quan.
        # =========================
        supabase.table(STUDENT_TABLE) \
            .update({field: value}) \
            .eq("license", license_code) \
            .execute()

        # Nếu sửa các thông tin có thể ảnh hưởng bảng khác thì đồng bộ lại
        fields_need_sync = {
            "name",
            "birthdate",
            "gender",
            "classroom",
            "timeclass",
            "clup",
            "phonenumber",
            "address",
        }

        if field in fields_need_sync:
            sync_student_profile_to_related_tables(license_code)

        return jsonify({
            "ok": True,
            "message": "Đã cập nhật và kiểm tra đồng bộ dữ liệu liên quan."
        })

    except Exception as e:
        print("[STUDENT FIELD UPDATE ERROR]", e)

        return jsonify({
            "ok": False,
            "message": str(e)
        }), 500

def get_next_belt_web(current_belt):
    current_belt = normalize_belt_name_web(current_belt)

    # Khóa riêng các trường hợp dễ lỗi
    if current_belt == "Cấp 1":
        return "1 Đẳng"

    if current_belt == "1 Đẳng":
        return "2 Đẳng"

    if current_belt == "2 Đẳng":
        return "3 Đẳng"

    if current_belt == "3 Đẳng":
        return "4 Đẳng"

    next_map = {
        "Cấp 10": "Cấp 9",
        "Cấp 9": "Cấp 8",
        "Cấp 8": "Cấp 7",
        "Cấp 7": "Cấp 6",
        "Cấp 6": "Cấp 5",
        "Cấp 5": "Cấp 4",
        "Cấp 4": "Cấp 3",
        "Cấp 3": "Cấp 2",
        "Cấp 2": "Cấp 1",

        "Cấp 1": "1 Đẳng",
        "1 Đẳng": "2 Đẳng",
        "2 Đẳng": "3 Đẳng",
        "3 Đẳng": "4 Đẳng",
        "4 Đẳng": "5 Đẳng",
        "5 Đẳng": "6 Đẳng",
        "6 Đẳng": "7 Đẳng",
        "7 Đẳng": "8 Đẳng",
        "8 Đẳng": "9 Đẳng",
        "9 Đẳng": "10 Đẳng",
    }

    return next_map.get(current_belt, current_belt)

def belt_css_class_web(belt_name):
    s = str(belt_name or "").strip().lower()
    s = remove_accents(s)
    s = s.replace(" ", "-")
    return s

def get_previous_belt_web(current_belt):
    current_belt = normalize_belt_name_web(current_belt)

    if current_belt in BELT_FLOW_WEB:
        idx = BELT_FLOW_WEB.index(current_belt)
        if idx - 1 >= 0:
            return BELT_FLOW_WEB[idx - 1]

    return current_belt or "Cấp 10"

def get_belt_index_web(belt):
    belt = str(belt or "").strip()
    if belt in BELT_FLOW_WEB:
        return BELT_FLOW_WEB.index(belt)
    return -1


def parse_ky_thi_web(ky_thi):
    """
    Nhận dạng:
    - Thi cấp: 2026-Q1, 2026-Q2, 2026-Q3, 2026-Q4
    - Thi đẳng: L1-2026, L2-2026, MN-2026, MT-2026, MB-2026, QG-2026

    Trả về số để so sánh thứ tự kỳ thi.
    """
    s = str(ky_thi or "").strip().upper()

    if not s:
        return 0

    # Thi cấp: 2026-Q3
    m = re.match(r"^(20\d{2})-Q([1-4])$", s)
    if m:
        year = int(m.group(1))
        q = int(m.group(2))
        return year * 100 + q

    # Thi đẳng: L1-2026, MN-2026...
    dan_order = {
        "L1": 11,
        "L2": 12,
        "MN": 13,
        "MT": 14,
        "MB": 15,
        "QG": 16,
    }

    m = re.match(r"^(L1|L2|MN|MT|MB|QG)-(20\d{2})$", s)
    if m:
        code = m.group(1)
        year = int(m.group(2))
        return year * 100 + dan_order.get(code, 0)

    # Dự phòng nếu dữ liệu cũ bị đảo: 2026-L1
    m = re.match(r"^(20\d{2})-(L1|L2|MN|MT|MB|QG)$", s)
    if m:
        year = int(m.group(1))
        code = m.group(2)
        return year * 100 + dan_order.get(code, 0)

    return 0


def get_expected_exam_belt_web(student_belt, previous_results):
    """
    Tính cấp dự thi đúng theo hiện trạng mới nhất.

    Quy tắc:
    - Nếu chưa từng thi: cấp dự thi = cấp kế tiếp của cấp hiện tại trong student.
    - Nếu lần gần nhất Vắng hoặc Không đạt: thi lại đúng cấp đó.
    - Nếu lần gần nhất Đạt: cấp dự thi = cấp kế tiếp của cấp hiện tại trong student.
      Vì student.belt đã được cập nhật sau khi chốt Đạt.
    """

    student_belt = normalize_belt_name_web(student_belt)

    if not previous_results:
        return get_next_belt_web(student_belt)

    previous_results = sorted(
        previous_results,
        key=lambda r: parse_ky_thi_web(r.get("ky_thi")),
        reverse=True
    )

    last = previous_results[0]
    last_exam_belt = normalize_belt_name_web(last.get("cap_dai_thi"))
    last_result = str(last.get("ket_qua") or "").strip()

    # Nếu quý trước vắng hoặc không đạt thì quý sau thi lại cấp đó
    if last_result in ["Không đạt", "Vắng"]:
        return last_exam_belt

    # Nếu quý trước đạt thì lấy cấp hiện tại trong student để tính cấp tiếp theo
    if last_result == "Đạt":
        if student_belt in BELT_FLOW_WEB:
            return get_next_belt_web(student_belt)

        return get_next_belt_web(last_exam_belt)

    # Trường hợp dữ liệu lạ thì ưu tiên cấp hiện tại
    if student_belt in BELT_FLOW_WEB:
        return get_next_belt_web(student_belt)

    return get_next_belt_web(last_exam_belt)

def recalc_student_belt_after_delete_web(student_license, deleted_exam_belt=""):
    """
    Tính lại cấp hiện tại của học viên sau khi xóa kết quả.

    Nguyên tắc:
    - Nếu còn kết quả Đạt trước đó => cấp hiện tại = cấp Đạt gần nhất.
    - Nếu không còn kết quả Đạt nào => cấp hiện tại = cấp trước của cấp vừa xóa.
      VD: xóa kết quả Đạt Cấp 7 => học viên quay về Cấp 8.
    """

    student_rows = safe_rows(STUDENT_TABLE, "*", license=student_license)
    student = student_rows[0] if student_rows else {}

    if not student:
        return ""

    current_belt = str(student.get("belt") or "").strip() or "Cấp 10"

    results = safe_rows(KETQUA_TABLE, "*", ma_hv=student_license)

    results = sorted(
        results,
        key=lambda r: parse_ky_thi_web(r.get("ky_thi")),
        reverse=True
    )

    # Nếu còn lịch sử Đạt thì lấy cấp Đạt mới nhất
    for r in results:
        if str(r.get("ket_qua") or "").strip() == "Đạt":
            passed_belt = str(r.get("cap_dai_thi") or "").strip()
            if passed_belt:
                return passed_belt

    # Nếu không còn kết quả Đạt nào, quay về cấp trước khi thi cấp vừa xóa
    deleted_exam_belt = str(deleted_exam_belt or "").strip()

    if deleted_exam_belt:
        return get_previous_belt_web(deleted_exam_belt)

    return current_belt or "Cấp 10"

def format_result_error_web(ma_hv, ho_ten, message):
    return f"{ma_hv} - {ho_ten}: {message}"

def clear_student_temp_leave_if_paid(license_code, month_codes):
    """
    Khi học viên đóng học phí lại, nếu tháng đóng trùng với tháng đang tạm nghỉ
    thì tự hủy trạng thái tạm nghỉ.
    """
    license_code = str(license_code or "").strip()

    if not license_code:
        return

    raw = str(month_codes or "")
    codes = [
        x.strip()
        for x in re.split(r"\s*-\s*", raw)
        if x.strip()
    ]

    if not codes:
        return

    try:
        rows = supabase.table(STUDENT_TABLE) \
            .select("license,temp_leave_month,temp_leave_year") \
            .eq("license", license_code) \
            .limit(1) \
            .execute().data or []

        if not rows:
            return

        student = rows[0]

        temp_month = student.get("temp_leave_month")
        temp_year = student.get("temp_leave_year")

        if not temp_month or not temp_year:
            return

        temp_code = f"{int(temp_month):02d}{int(temp_year)}"

        if temp_code in codes:
            supabase.table(STUDENT_TABLE) \
                .update({
                    "temp_leave_month": None,
                    "temp_leave_year": None,
                    "temp_leave_note": None,
                }) \
                .eq("license", license_code) \
                .execute()

    except Exception as e:
        print("[CLEAR TEMP LEAVE IF PAID ERROR]", e)

def safe_rows(table_name, select_text="*", **filters):
    try:
        q = supabase.table(table_name).select(select_text)

        for k, v in filters.items():
            q = q.eq(k, v)

        return q.execute().data or []

    except Exception as e:
        # Không dùng tiếng Việt ở print để tránh lỗi UnicodeEncodeError trên Windows terminal
        try:
            print(f"[WARN] Cannot read table {table_name}: {repr(e)}")
        except:
            pass

        return []


def get_prev_month_year(month, year):
    month = int(month)
    year = int(year)

    if month == 1:
        return 12, year - 1

    return month - 1, year


def get_month_code_web(month, year):
    return f"{int(month):02d}{int(year)}"


def parse_fee_datetime_web(raw):
    raw = str(raw or "").strip()

    if not raw:
        return None

    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except:
        return None


def safe_due_date_web(year, month, day):
    year = int(year)
    month = int(month)
    day = int(day)

    last_day = calendar.monthrange(year, month)[1]
    day = min(day, last_day)

    return date(year, month, day)


def get_compare_date_web(year, month):
    """
    Nếu đang lọc tháng hiện tại: so với ngày hôm nay.
    Nếu lọc tháng cũ: so với ngày cuối tháng đó để xem cuối tháng còn ai quá hạn.
    Nếu lọc tháng tương lai: so với ngày đầu tháng đó.
    """
    today = date.today()
    year = int(year)
    month = int(month)

    if today.year == year and today.month == month:
        return today

    selected_first = date(year, month, 1)

    if selected_first > today.replace(day=1):
        return selected_first

    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, last_day)


@app.get("/unpaid")
def unpaid():
    now = datetime.now()

    year = request.args.get("year", str(now.year))
    month = request.args.get("month", str(now.month))

    try:
        year_int = int(year)
        month_int = int(month)
    except:
        year_int = now.year
        month_int = now.month

    current_code = get_month_code_web(month_int, year_int)

    prev_month, prev_year = get_prev_month_year(month_int, year_int)
    prev_code = get_month_code_web(prev_month, prev_year)

    compare_date = get_compare_date_web(year_int, month_int)

    students = safe_rows(STUDENT_TABLE)
    fees = safe_rows(HOCPHI_TABLE)

    # =========================
    # 1) Ai đã đóng tháng đang lọc thì loại khỏi danh sách chưa đóng
    # =========================
    paid_current_ids = set()

    for f in fees:
        ma_thang = str(f.get("ma_thang") or "")
        ma_hv = str(f.get("ma_hv") or "").strip()

        if ma_hv and current_code in ma_thang:
            paid_current_ids.add(ma_hv)

    # =========================
    # 2) Lấy lần đóng gần nhất của tháng trước cho từng HV
    # =========================
    last_prev_fee_by_student = {}

    for f in fees:
        ma_thang = str(f.get("ma_thang") or "")
        ma_hv = str(f.get("ma_hv") or "").strip()

        if not ma_hv:
            continue

        if prev_code not in ma_thang:
            continue

        paid_dt = parse_fee_datetime_web(f.get("thoi_gian"))

        if not paid_dt:
            continue

        old = last_prev_fee_by_student.get(ma_hv)

        if not old or paid_dt > old["paid_dt"]:
            last_prev_fee_by_student[ma_hv] = {
                "paid_dt": paid_dt,
                "fee": f
            }

    rows = []

    summary = {
        "total": 0,
        "vang_lai": 0,
        "not_due": 0,
        "near_7_4": 0,
        "near_3_0": 0,
        "overdue": 0,
        "temp_leave": 0,
    }

    for s in students:
        active = str(s.get("active") or "").strip().lower()

        # Chỉ xét HV đang hoạt động
        if active not in ["có", "1", "true", "yes"]:
            continue

        license_code = str(s.get("license") or "").strip()

        if not license_code:
            continue

        # Đã đóng tháng đang lọc rồi thì không nằm trong danh sách chưa đóng
        if license_code in paid_current_ids:
            continue

        row = dict(s)

        # =========================
        # TẠM NGHỈ 1 THÁNG
        # Nếu học viên được set tạm nghỉ đúng tháng/năm đang lọc
        # thì ưu tiên hiện trạng thái này.
        # =========================
        temp_leave_month = s.get("temp_leave_month")
        temp_leave_year = s.get("temp_leave_year")

        try:
            is_temp_leave = (
                int(temp_leave_month or 0) == int(month_int)
                and int(temp_leave_year or 0) == int(year_int)
            )
        except:
            is_temp_leave = False

        if is_temp_leave:
            row["fee_status"] = "temp_leave"
            row["fee_status_label"] = "Tạm nghỉ 1 tháng"
            row["fee_due_date"] = "—"
            row["fee_days_left"] = "—"
            row["fee_row_class"] = "fee-row-temp-leave"

            summary["temp_leave"] += 1
            rows.append(row)
            continue

        prev_fee_info = last_prev_fee_by_student.get(license_code)

        # =========================
        # 3) Không có đóng tháng trước => HV vãng lai
        # =========================
        if not prev_fee_info:
            row["fee_status"] = "vang_lai"
            row["fee_status_label"] = "HV vãng lai"
            row["fee_due_date"] = "—"
            row["fee_days_left"] = "—"
            row["fee_row_class"] = "fee-row-vang-lai"

            summary["vang_lai"] += 1
            rows.append(row)
            continue

        # =========================
        # 4) Có đóng tháng trước => lấy ngày đóng tháng trước làm ngày đến hạn tháng này
        # VD đóng 15/6 -> hạn tháng 7 là 15/7
        # =========================
        last_paid_dt = prev_fee_info["paid_dt"]
        due_date = safe_due_date_web(year_int, month_int, last_paid_dt.day)

        days_left = (due_date - compare_date).days

        row["fee_due_date"] = due_date.strftime("%d/%m/%Y")
        row["fee_days_left"] = days_left

        if days_left < 0:
            row["fee_status"] = "overdue"
            row["fee_status_label"] = "Quá hạn"
            row["fee_row_class"] = "fee-row-overdue"
            summary["overdue"] += 1

        elif 0 <= days_left <= 3:
            row["fee_status"] = "near_3_0"
            row["fee_status_label"] = "Gần đến hạn 3-0 ngày"
            row["fee_row_class"] = "fee-row-near-orange"
            summary["near_3_0"] += 1

        elif 4 <= days_left <= 7:
            row["fee_status"] = "near_7_4"
            row["fee_status_label"] = "Gần đến hạn 7-4 ngày"
            row["fee_row_class"] = "fee-row-near-yellow"
            summary["near_7_4"] += 1

        else:
            row["fee_status"] = "not_due"
            row["fee_status_label"] = "Chưa đến hạn"
            row["fee_row_class"] = "fee-row-not-due"
            summary["not_due"] += 1

        rows.append(row)

    summary["total"] = len(rows)

    # Sắp xếp: Quá hạn -> cam -> vàng -> chưa đến hạn -> vãng lai
    status_order = {
        "overdue": 0,
        "near_3_0": 1,
        "near_7_4": 2,
        "not_due": 3,
        "temp_leave": 4,
        "vang_lai": 5,
    }

    rows.sort(key=lambda r: (
        status_order.get(r.get("fee_status"), 9),
        str(r.get("classroom") or ""),
        str(r.get("timeclass") or ""),
        str(r.get("name") or "")
    ))

    return render_template(
        "unpaid.html",
        rows=rows,
        year=str(year_int),
        month=str(month_int),
        current_year=now.year,
        summary=summary,
        compare_date=compare_date.strftime("%d/%m/%Y"),
        prev_month=prev_month,
        prev_year=prev_year
    )

@app.post("/unpaid/temp-leave")
def unpaid_temp_leave():
    license_code = request.form.get("license", "").strip()
    month = request.form.get("month", "").strip()
    year = request.form.get("year", "").strip()

    if not license_code or not month or not year:
        flash("Thiếu thông tin để set tạm nghỉ.", "danger")
        return redirect(url_for("unpaid"))

    try:
        supabase.table(STUDENT_TABLE) \
            .update({
                "temp_leave_month": int(month),
                "temp_leave_year": int(year),
                "temp_leave_note": "Tạm nghỉ 1 tháng",
            }) \
            .eq("license", license_code) \
            .execute()

        flash(f"Đã set tạm nghỉ 1 tháng cho {license_code}.", "success")

    except Exception as e:
        print("[SET TEMP LEAVE ERROR]", e)
        flash(f"Lỗi set tạm nghỉ: {e}", "danger")

    return back_to_current_page("unpaid")


@app.post("/unpaid/temp-leave-cancel")
def unpaid_temp_leave_cancel():
    license_code = request.form.get("license", "").strip()
    month = request.form.get("month", "").strip()
    year = request.form.get("year", "").strip()

    if not license_code:
        flash("Thiếu mã học viên để hủy tạm nghỉ.", "danger")
        return back_to_current_page("unpaid")

    try:
        supabase.table(STUDENT_TABLE) \
            .update({
                "temp_leave_month": None,
                "temp_leave_year": None,
                "temp_leave_note": None,
            }) \
            .eq("license", license_code) \
            .execute()

        flash(f"Đã hủy tạm nghỉ cho {license_code}.", "success")

    except Exception as e:
        print("[CANCEL TEMP LEAVE ERROR]", e)
        flash(f"Lỗi hủy tạm nghỉ: {e}", "danger")

    return back_to_current_page("unpaid")

@app.get("/exam-tracking")
def exam_tracking():
    now = datetime.now()

    year = request.args.get("year", str(now.year)).strip()
    quarter = request.args.get("quarter", f"Q{(now.month - 1)//3 + 1}").strip().upper()

    dan_quarters = ["L1", "L2", "MN", "MT", "MB", "QG"]

    if quarter in dan_quarters:
        ma_quy = f"{quarter}-{year}"
        ky_thi = f"{quarter}-{year}"
        tracking_title = "Theo dõi thi cấp đẳng"
    else:
        ma_quy = f"{quarter}{year}"
        ky_thi = f"{year}-{quarter}"
        tracking_title = "Theo dõi thi cấp đẳng"

    current_ky_value = parse_ky_thi_web(ky_thi)

    rows = safe_rows(HOCPHI_TABLE)
    rows = [
        r for r in rows
        if str(r.get("ma_quy") or "").strip().upper() == ma_quy.upper()
    ]

    rows.sort(key=lambda x: str(x.get("thoi_gian") or ""), reverse=True)

    # =========================
    # LẤY DANH SÁCH MÃ HV
    # =========================
    ids = []

    for r in rows:
        ma_hv = str(r.get("ma_hv") or "").strip()
        if ma_hv and ma_hv not in ids:
            ids.append(ma_hv)

    students_map = {}
    results_map = {}

    if ids:
        students = supabase.table(STUDENT_TABLE) \
            .select("license,name,birthdate,gender,belt") \
            .in_("license", ids) \
            .execute().data or []

        students_map = {
            str(s.get("license") or "").strip(): s
            for s in students
        }

        all_results = supabase.table(KETQUA_TABLE) \
            .select("*") \
            .in_("ma_hv", ids) \
            .execute().data or []

        for result in all_results:
            ma_hv = str(result.get("ma_hv") or "").strip()
            if ma_hv:
                results_map.setdefault(ma_hv, []).append(result)

    # =========================
    # TÍNH TIỀN THEO PHIẾU THU THỰC TẾ
    # =========================
    total_count = len(rows)

    total_exam_fee = sum([
        int(r.get("tong_tien") or 0)
        for r in rows
    ])

    cash_rows = [
        r for r in rows
        if str(r.get("chuyen_khoan") or "").strip().upper() == "TM"
    ]

    bank_rows = [
        r for r in rows
        if str(r.get("chuyen_khoan") or "").strip().upper() == "CK"
    ]

    cash_count = len(cash_rows)
    bank_count = len(bank_rows)

    cash_total = sum([
        int(r.get("tong_tien") or 0)
        for r in cash_rows
    ])

    bank_total = sum([
        int(r.get("tong_tien") or 0)
        for r in bank_rows
    ])

    # =========================
    # THỐNG KÊ CẤP DỰ THI
    # =========================
    belt_order = [
        "Cấp 9", "Cấp 8", "Cấp 7", "Cấp 6", "Cấp 5",
        "Cấp 4", "Cấp 3", "Cấp 2", "Cấp 1",
        "1 Đẳng", "2 Đẳng", "3 Đẳng", "4 Đẳng", "5 Đẳng",
        "6 Đẳng", "7 Đẳng", "8 Đẳng", "9 Đẳng", "10 Đẳng"
    ]

    belt_stats_map = {}

    for r in rows:
        ma_hv = str(r.get("ma_hv") or "").strip()
        student = students_map.get(ma_hv, {})

        current_belt = normalize_belt_name_web(student.get("belt"))
        exam_belt = get_next_belt_web(current_belt)

        if current_belt == "Cấp 1":
            exam_belt = "1 Đẳng"

        r["ho_ten"] = student.get("name") or r.get("ho_ten") or ""
        r["ngay_sinh"] = student.get("birthdate") or r.get("ngay_sinh") or ""
        r["gioi_tinh"] = student.get("gender") or r.get("gioi_tinh") or ""
        r["cap_hien_tai"] = current_belt
        r["cap_du_thi"] = exam_belt

        if exam_belt:
            belt_stats_map[exam_belt] = belt_stats_map.get(exam_belt, 0) + 1

    belt_stats = []

    for belt_name in belt_order:
        count = belt_stats_map.get(belt_name, 0)

        if count > 0:
            belt_stats.append({
                "name": belt_name,
                "count": count,
                "class": belt_css_class_web(belt_name)
            })

    return render_template(
        "exam_tracking.html",
        rows=rows,
        year=year,
        quarter=quarter,
        current_year=now.year,
        total_count=total_count,
        total_exam_fee=total_exam_fee,
        cash_count=cash_count,
        bank_count=bank_count,
        cash_total=cash_total,
        bank_total=bank_total,
        belt_stats=belt_stats,
        tracking_title=tracking_title
    )

@app.get("/exam-tracking/export")
def exam_tracking_export():
    now = datetime.now()

    year = request.args.get("year", str(now.year)).strip()
    quarter = request.args.get("quarter", f"Q{(now.month - 1)//3 + 1}").strip().upper()

    ma_quy = build_exam_ma_quy_web(year, quarter)

    rows = safe_rows(HOCPHI_TABLE)
    rows = [
        r for r in rows
        if str(r.get("ma_quy") or "").strip().upper() == ma_quy.upper()
    ]

    rows.sort(key=lambda x: str(x.get("thoi_gian") or ""), reverse=True)

    ids = []

    for r in rows:
        ma_hv = str(r.get("ma_hv") or "").strip()
        if ma_hv and ma_hv not in ids:
            ids.append(ma_hv)

    students_map = {}

    if ids:
        students = supabase.table(STUDENT_TABLE) \
            .select("license,name,birthdate,gender,belt") \
            .in_("license", ids) \
            .execute().data or []

        students_map = {
            str(s.get("license") or "").strip(): s
            for s in students
        }

    export_rows = []

    for r in rows:
        ma_hv = str(r.get("ma_hv") or "").strip()
        student = students_map.get(ma_hv, {})

        current_belt = normalize_belt_name_web(student.get("belt"))
        exam_belt = get_next_belt_web(current_belt)

        if current_belt == "Cấp 1":
            exam_belt = "1 Đẳng"

        export_rows.append({
            "ma_quy": r.get("ma_quy", ""),
            "ma_hv": ma_hv,
            "ho_ten": student.get("name") or r.get("ho_ten") or "",
            "ngay_sinh": student.get("birthdate") or r.get("ngay_sinh") or "",
            "gioi_tinh": student.get("gender") or r.get("gioi_tinh") or "",
            "cap_hien_tai": current_belt,
            "cap_du_thi": exam_belt,
            "ghi_chu": r.get("ghi_chu", ""),
        })

    wb = Workbook()
    ws = wb.active
    ws.title = "Theo doi thi cap dang"

    headers = [
        "Mã quý",
        "Mã HV",
        "Họ tên",
        "Ngày sinh",
        "Giới tính",
        "Cấp hiện tại",
        "Cấp dự thi",
        "Ghi chú",
    ]

    # =========================
    # STYLE
    # =========================
    title_font = Font(bold=True, size=16, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="0F2A4A")

    sub_font = Font(bold=True, size=11, color="1E293B")

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1D4ED8")

    thin_side = Side(style="thin", color="D9E2EF")
    table_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # =========================
    # TITLE
    # =========================
    ws.merge_cells("A1:H1")
    ws["A1"] = "DANH SÁCH THI CẤP ĐẲNG"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = center

    ws.merge_cells("A2:H2")
    ws["A2"] = f"Bộ lọc: Năm {year} - Quý {quarter} | Tổng số: {len(export_rows)} võ sinh"
    ws["A2"].font = sub_font
    ws["A2"].alignment = center

    ws.merge_cells("A3:H3")
    ws["A3"] = f"Xuất lúc: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    ws["A3"].font = Font(italic=True, color="475569")
    ws["A3"].alignment = center

    # =========================
    # HEADER
    # =========================
    header_row = 5

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = table_border

    # =========================
    # DATA
    # =========================
    data_start_row = header_row + 1

    for row_idx, item in enumerate(export_rows, start=data_start_row):
        values = [
            item["ma_quy"],
            item["ma_hv"],
            item["ho_ten"],
            item["ngay_sinh"],
            item["gioi_tinh"],
            item["cap_hien_tai"],
            item["cap_du_thi"],
            item["ghi_chu"],
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = table_border
            cell.alignment = center if col_idx not in [3, 8] else left

            if col_idx in [6, 7]:
                cell.font = Font(bold=True)

    # =========================
    # AUTO WIDTH THEO ĐỘ DÀI CHỮ
    # =========================
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)

        max_len = len(str(header))

        for row in range(1, ws.max_row + 1):
            value = ws.cell(row=row, column=col_idx).value
            if value is not None:
                max_len = max(max_len, len(str(value)))

        width = max_len + 4

        # Giới hạn để không quá nhỏ hoặc quá dài
        if width < 12:
            width = 12
        if width > 36:
            width = 36

        ws.column_dimensions[col_letter].width = width

    # =========================
    # ROW HEIGHT + FREEZE
    # =========================
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[5].height = 24

    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:H{ws.max_row}"

    # Canh giữa tiêu đề sau khi merge
    for row in [1, 2, 3]:
        for col in range(1, 9):
            ws.cell(row=row, column=col).alignment = center

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"theo_doi_thi_cap_dang_{year}_{quarter}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/results")
def results():
    now = datetime.now()

    year = request.args.get("year", "").strip()
    quarter = request.args.get("quarter", "").strip()
    name = request.args.get("name", "").strip()
    cap_thi = request.args.get("cap_thi", "").strip()
    ket_qua = request.args.get("ket_qua", "").strip()

    rows = safe_rows(KETQUA_TABLE)

    # Sắp xếp mới nhất lên trên
    rows.sort(key=lambda r: str(r.get("ky_thi") or ""), reverse=True)

    # Lấy danh sách năm, cấp thi để đổ vào combobox
    year_options = sorted(
        {
            str(r.get("ky_thi") or "").split("-")[0]
            for r in rows
            if str(r.get("ky_thi") or "").split("-")[0]
        },
        reverse=True
    )

    cap_thi_options = sorted(
        {
            str(r.get("cap_dai_thi") or "").strip()
            for r in rows
            if str(r.get("cap_dai_thi") or "").strip()
        }
    )

    # Nếu chưa chọn năm thì mặc định năm hiện tại
    if not year:
        year = str(now.year)

    filtered = []

    for r in rows:
        ky_thi = str(r.get("ky_thi") or "")
        r_year = ky_thi.split("-")[0] if "-" in ky_thi else ""
        r_quarter = ky_thi.split("-")[1] if "-" in ky_thi else ""

        if year and r_year != year:
            continue

        if quarter and r_quarter != quarter:
            continue

        if name:
            text = f"{r.get('ma_hv', '')} {r.get('ho_ten', '')}".lower()
            if name.lower() not in text:
                continue

        if cap_thi and str(r.get("cap_dai_thi") or "") != cap_thi:
            continue

        if ket_qua and str(r.get("ket_qua") or "") != ket_qua:
            continue

        filtered.append(r)

    delete_errors = session.pop("result_delete_errors", [])

    return render_template(
        "results.html",
        rows=filtered,
        year=year,
        quarter=quarter,
        name=name,
        cap_thi=cap_thi,
        ket_qua=ket_qua,
        current_year=now.year,
        year_options=year_options,
        cap_thi_options=cap_thi_options,
        delete_errors=delete_errors
    )

@app.post("/results/delete")
def results_delete():
    ky_thi = request.form.get("ky_thi", "").strip()
    ma_hv = request.form.get("ma_hv", "").strip()
    so_thi = request.form.get("so_thi", "").strip()

    back_year = request.args.get("year", "").strip()
    back_quarter = request.args.get("quarter", "").strip()

    if not ky_thi or not ma_hv:
        session["result_delete_errors"] = [
            """
            <b>Thiếu thông tin xóa kết quả</b><br>
            Lý do: Không nhận được kỳ thi hoặc mã hội viên.<br>
            Cách xử lý: Ken tải lại trang rồi click phải xóa lại.
            """
        ]
        return redirect(url_for("results", year=back_year, quarter=back_quarter))

    try:
        all_results = safe_rows(KETQUA_TABLE, "*", ma_hv=ma_hv)

        if not all_results:
            session["result_delete_errors"] = [
                f"""
                <b>{ma_hv}</b><br>
                Lý do: Không tìm thấy lịch sử thi cấp của học viên này.<br>
                Cách xử lý: Ken kiểm tra lại bảng kết quả thi cấp.
                """
            ]
            return redirect(url_for("results", year=back_year, quarter=back_quarter))

        target_result = None

        for r in all_results:
            same_ky = str(r.get("ky_thi") or "").strip() == ky_thi
            same_so_thi = True

            if so_thi:
                same_so_thi = str(r.get("so_thi") or "").strip() == so_thi

            if same_ky and same_so_thi:
                target_result = r
                break

        if not target_result:
            session["result_delete_errors"] = [
                f"""
                <b>{ma_hv}</b><br>
                Lý do: Không tìm thấy dòng kết quả cần xóa trong kỳ <b>{ky_thi}</b>.<br>
                Cách xử lý: Ken tải lại trang hoặc kiểm tra lại bảng Supabase.
                """
            ]
            return redirect(url_for("results", year=back_year, quarter=back_quarter))

        target_ky_value = parse_ky_thi_web(target_result.get("ky_thi"))

        newer_results = [
            r for r in all_results
            if parse_ky_thi_web(r.get("ky_thi")) > target_ky_value
        ]

        if newer_results:
            newer_results = sorted(
                newer_results,
                key=lambda r: parse_ky_thi_web(r.get("ky_thi")),
                reverse=True
            )

            newest = newer_results[0]

            session["result_delete_errors"] = [
                f"""
                <b>{ma_hv} - {target_result.get('ho_ten', '')}</b><br>
                Lý do: Không thể xóa kết quả kỳ <b>{target_result.get('ky_thi')}</b>
                vì học viên này đã có kết quả mới hơn.<br><br>

                Kết quả mới nhất hiện tại:<br>
                Kỳ thi: <b>{newest.get('ky_thi')}</b><br>
                Cấp thi: <b>{newest.get('cap_dai_thi')}</b><br>
                Số thi: <b>{newest.get('so_thi')}</b><br>
                Kết quả: <b>{newest.get('ket_qua')}</b><br><br>

                Cách xử lý: Ken phải xóa từ kết quả mới nhất về cũ nhất.
                Ví dụ đã thi Cấp 7 → Cấp 6 → Cấp 5 thì phải xóa <b>Cấp 5</b> trước.
                """
            ]

            return redirect(url_for("results", year=back_year, quarter=back_quarter))

        q = supabase.table(KETQUA_TABLE) \
            .delete() \
            .eq("ky_thi", ky_thi) \
            .eq("ma_hv", ma_hv)

        if so_thi:
            q = q.eq("so_thi", so_thi)

        q.execute()

        deleted_exam_belt = str(target_result.get("cap_dai_thi") or "").strip()
        new_belt = recalc_student_belt_after_delete_web(ma_hv, deleted_exam_belt)

        if new_belt:
            supabase.table(STUDENT_TABLE) \
                .update({"belt": new_belt}) \
                .eq("license", ma_hv) \
                .execute()

        flash(
            f"Đã xóa kết quả {ma_hv} - {ky_thi}. "
            f"Đã cập nhật cấp hiện tại về {new_belt}."
        )

    except Exception as e:
        print("[DELETE RESULT ERROR]", e)

        session["result_delete_errors"] = [
            f"""
            <b>Lỗi xóa kết quả</b><br>
            Lý do hệ thống: {e}<br>
            Cách xử lý: Ken kiểm tra lại kết nối Supabase hoặc dữ liệu trong bảng ketqua.
            """
        ]

    return redirect(url_for("results", year=back_year, quarter=back_quarter))

def parse_web_date(raw):
    """
    Nhận yyyy-mm-dd từ input type=date.
    Trả về datetime hoặc None.
    """
    raw = str(raw or "").strip()

    try:
        return datetime.strptime(raw, "%Y-%m-%d")
    except:
        return None


def format_activity_date_range(start_raw, end_raw):
    start = parse_web_date(start_raw)
    end = parse_web_date(end_raw)

    if not start and not end:
        return ""

    if start and not end:
        return start.strftime("%d/%m/%Y")

    if end and not start:
        return end.strftime("%d/%m/%Y")

    # Nếu cùng ngày
    if start.date() == end.date():
        return start.strftime("%d/%m/%Y")

    # Cùng tháng, cùng năm: 14 - 16/07/2026
    if start.month == end.month and start.year == end.year:
        return f"{start.strftime('%d')} - {end.strftime('%d/%m/%Y')}"

    # Khác tháng hoặc khác năm: 28/06 - 04/07/2026
    if start.year == end.year:
        return f"{start.strftime('%d/%m')} - {end.strftime('%d/%m/%Y')}"

    # Khác năm: 28/12/2025 - 04/01/2026
    return f"{start.strftime('%d/%m/%Y')} - {end.strftime('%d/%m/%Y')}"

@app.get("/activities")
def activities():
    rows = safe_rows(HOATDONG_TABLE)
    rows.sort(key=lambda r: str(r.get("created_at") or r.get("thoi_gian") or ""), reverse=True)

    students = safe_rows(STUDENT_TABLE)
    students.sort(key=lambda s: str(s.get("name") or ""))

    return render_template(
        "activities.html",
        rows=rows,
        students=students
    )

@app.post("/activities/add")
def activities_add():
    f = request.form

    ma_hv = f.get("ma_hv", "").strip()

    student_rows = safe_rows(STUDENT_TABLE, "*", license=ma_hv)
    student = student_rows[0] if student_rows else {}

    if not student:
        flash("Không tìm thấy hội viên.")
        return redirect(url_for("activities"))

    start_date = f.get("start_date", "").strip()
    end_date = f.get("end_date", "").strip()

    thoi_gian_text = format_activity_date_range(start_date, end_date)

    payload = {
        "ma_hv": ma_hv,
        "ho_ten": student.get("name", ""),
        "hoat_dong": f.get("hoat_dong", "").strip(),
        "thoi_gian": thoi_gian_text,
        "dia_diem": f.get("dia_diem", "").strip(),
        "noi_dung": f.get("noi_dung", "").strip(),
        "ket_qua": f.get("ket_qua", "").strip(),
    }

    try:
        supabase.table(HOATDONG_TABLE).insert(payload).execute()
        flash(f"Đã thêm hoạt động cho {student.get('name', '')}")
    except Exception as e:
        print("[ADD ACTIVITY ERROR]", e)
        flash(f"Lỗi thêm hoạt động: {e}")

    return back_to_current_page("activities")


@app.post("/activities/update/<activity_id>")
def activities_update(activity_id):
    f = request.form

    ma_hv = f.get("ma_hv", "").strip()

    student_rows = safe_rows(STUDENT_TABLE, "*", license=ma_hv)
    student = student_rows[0] if student_rows else {}

    if not student:
        flash("Không tìm thấy hội viên để cập nhật hoạt động.", "danger")
        return back_to_current_page("activities")

    payload = {
        "ma_hv": ma_hv,
        "ho_ten": student.get("name", ""),
        "hoat_dong": f.get("hoat_dong", "").strip(),
        "thoi_gian": f.get("thoi_gian", "").strip(),
        "dia_diem": f.get("dia_diem", "").strip(),
        "noi_dung": f.get("noi_dung", "").strip(),
        "ket_qua": f.get("ket_qua", "").strip(),
    }

    try:
        supabase.table(HOATDONG_TABLE) \
            .update(payload) \
            .eq("id", activity_id) \
            .execute()

        flash(f"Đã cập nhật hoạt động cho {student.get('name', '')}", "success")

    except Exception as e:
        print("[UPDATE ACTIVITY ERROR]", e)
        flash(f"Lỗi cập nhật hoạt động: {e}", "danger")

    return back_to_current_page("activities")


@app.post("/activities/delete/<activity_id>")
def activities_delete(activity_id):
    try:
        supabase.table(HOATDONG_TABLE) \
            .delete() \
            .eq("id", activity_id) \
            .execute()

        flash("Đã xóa hoạt động.", "success")

    except Exception as e:
        print("[DELETE ACTIVITY ERROR]", e)
        flash(f"Lỗi xóa hoạt động: {e}", "danger")

    return back_to_current_page("activities")


def clean_option_list(values):
    cleaned = []
    seen = set()

    for value in values or []:
        value = str(value or "").strip()

        if not value:
            continue

        key = value.lower()

        if key in seen:
            continue

        seen.add(key)
        cleaned.append(value)

    return cleaned


def get_class_options_settings(settings=None):
    settings = settings or load_app_settings()
    defaults = DEFAULT_APP_SETTINGS["class_options"]

    class_options = settings.setdefault("class_options", {})

    for key, default_values in defaults.items():
        current_values = class_options.get(key)

        if not isinstance(current_values, list):
            current_values = default_values.copy()

        class_options[key] = clean_option_list(current_values)

        if not class_options[key]:
            class_options[key] = default_values.copy()

    return class_options


OPTION_META = {
    "classrooms": "Lớp",
    "timeclasses": "Ca",
    "clubs": "CLB",
}

def get_list_count(form, name):
    try:
        return int(form.get(name, "0") or 0)
    except:
        return 0


def read_club_info_from_form(form, files, old_club_info):
    old_club_info = old_club_info or {}

    head_old = old_club_info.get("head_coach", {}) or {}

    head_photo_url = str(form.get("head_photo_url") or head_old.get("photo_url") or "").strip()
    head_photo_file = files.get("head_photo_file")

    if head_photo_file and head_photo_file.filename:
        head_photo_url = upload_club_asset_to_supabase(head_photo_file, "club-info/coaches")

    head_coach = {
        "name": str(form.get("head_name") or "").strip(),
        "role": str(form.get("head_role") or "HLV Trưởng").strip(),
        "phone": str(form.get("head_phone") or "").strip(),
        "photo_url": head_photo_url,
        "description": str(form.get("head_description") or "").strip(),
    }

    registrar_old = old_club_info.get("registrar", {}) or {}

    registrar_photo_url = str(
        form.get("registrar_photo_url") or registrar_old.get("photo_url") or ""
    ).strip()

    registrar_photo_file = files.get("registrar_photo_file")

    if registrar_photo_file and registrar_photo_file.filename:
        registrar_photo_url = upload_club_asset_to_supabase(
            registrar_photo_file,
            "club-info/registrar"
        )

    registrar = {
        "name": str(form.get("registrar_name") or "").strip(),
        "role": str(form.get("registrar_role") or "Người ghi danh").strip(),
        "phone": str(form.get("registrar_phone") or "").strip(),
        "photo_url": registrar_photo_url,
        "description": str(form.get("registrar_description") or "").strip(),
    }

    coaches = []
    coach_count = get_list_count(form, "coach_count")

    group_map = {
        "246": {
            "group_title": "HLV lớp 2-4-6",
            "group_subtitle": "Lớp buổi chiều",
        },
        "357": {
            "group_title": "HLV lớp 3-5-7",
            "group_subtitle": "Lớp chính trong tuần",
        },
        "weekend": {
            "group_title": "HLV lớp sáng & cuối tuần",
            "group_subtitle": "Thứ 7, Chủ nhật và lớp sáng",
        },
    }

    for i in range(coach_count):
        if str(form.get(f"coach_delete_{i}") or "") == "1":
            continue

        name = str(form.get(f"coach_name_{i}") or "").strip()
        role = str(form.get(f"coach_role_{i}") or "").strip()
        group = str(form.get(f"coach_group_{i}") or "246").strip()
        phone = str(form.get(f"coach_phone_{i}") or "").strip()

        if not name and not role and not phone:
            continue

        photo_url = str(form.get(f"coach_photo_url_{i}") or "").strip()
        photo_file = files.get(f"coach_photo_file_{i}")

        if photo_file and photo_file.filename:
            photo_url = upload_club_asset_to_supabase(photo_file, "club-info/coaches")

        group_info = group_map.get(group, group_map["246"])

        coaches.append({
            "group": group,
            "group_title": group_info["group_title"],
            "group_subtitle": group_info["group_subtitle"],
            "name": name,
            "role": role,
            "phone": phone,
            "photo_url": photo_url,
        })

    regular_schedules = []
    regular_count = get_list_count(form, "regular_schedule_count")

    for i in range(regular_count):
        if str(form.get(f"regular_schedule_delete_{i}") or "") == "1":
            continue

        title = str(form.get(f"regular_schedule_title_{i}") or "").strip()
        time_text = str(form.get(f"regular_schedule_time_{i}") or "").strip()

        if title or time_text:
            regular_schedules.append({
                "title": title,
                "time": time_text,
            })

    summer_schedules = []
    summer_count = get_list_count(form, "summer_schedule_count")

    for i in range(summer_count):
        if str(form.get(f"summer_schedule_delete_{i}") or "") == "1":
            continue

        title = str(form.get(f"summer_schedule_title_{i}") or "").strip()
        time_text = str(form.get(f"summer_schedule_time_{i}") or "").strip()

        if title or time_text:
            summer_schedules.append({
                "title": title,
                "time": time_text,
            })

    exam_notes = []
    note_count = get_list_count(form, "exam_note_count")

    for i in range(note_count):
        if str(form.get(f"exam_note_delete_{i}") or "") == "1":
            continue

        note = str(form.get(f"exam_note_{i}") or "").strip()

        if note:
            exam_notes.append(note)

    return {
        "intro_title": str(form.get("intro_title") or "").strip(),
        "intro_subtitle": str(form.get("intro_subtitle") or "").strip(),
        "intro_content": str(form.get("intro_content") or "").strip(),

        "head_coach": head_coach,
        "registrar": registrar,
        "coaches": coaches,

        "regular_schedules": regular_schedules,
        "summer_schedules": summer_schedules,

        "fees_info": {
            "monthly_fee": str(form.get("monthly_fee") or "").strip(),
            "three_month_discount": str(form.get("three_month_discount") or "").strip(),
            "six_month_bonus": str(form.get("six_month_bonus") or "").strip(),
            "family_discount": str(form.get("family_discount") or "").strip(),
            "uniform_fee": str(form.get("uniform_fee") or "").strip(),
            "exam_fee": str(form.get("club_exam_fee") or "").strip(),
            "exam_notes": exam_notes,
        }
    }


@app.post("/setup/class-options/add")
def setup_class_options_add():
    option_type = request.form.get("option_type", "").strip()
    value = request.form.get("value", "").strip()

    if option_type not in OPTION_META:
        flash("Nhóm thông tin không hợp lệ.", "danger")
        return redirect(url_for("setup", tab="classes"))

    if not value:
        flash(f"{OPTION_META[option_type]} không được để trống.", "danger")
        return redirect(url_for("setup", tab="classes"))

    settings = load_app_settings()
    class_options = get_class_options_settings(settings)

    current = clean_option_list(class_options.get(option_type, []))

    if value.lower() in [x.lower() for x in current]:
        flash(f"{OPTION_META[option_type]} này đã có trong danh sách.", "danger")
        return redirect(url_for("setup", tab="classes"))

    current.append(value)

    class_options[option_type] = current
    settings["class_options"] = class_options

    save_app_settings(settings)

    flash(f"Đã thêm {OPTION_META[option_type]}: {value}", "success")
    return redirect(url_for("setup", tab="classes"))


@app.post("/setup/class-options/delete")
def setup_class_options_delete():
    option_type = request.form.get("option_type", "").strip()
    value = request.form.get("value", "").strip()

    if option_type not in OPTION_META:
        flash("Nhóm thông tin không hợp lệ.", "danger")
        return redirect(url_for("setup", tab="classes"))

    settings = load_app_settings()
    class_options = get_class_options_settings(settings)

    current = clean_option_list(class_options.get(option_type, []))
    class_options[option_type] = [x for x in current if x != value]

    settings["class_options"] = class_options

    save_app_settings(settings)

    flash(f"Đã xóa {OPTION_META[option_type]}: {value}", "success")
    return redirect(url_for("setup", tab="classes"))

@app.get("/information")
def information():
    return render_template("information.html", settings=load_app_settings())


@app.route("/setup", methods=["GET", "POST"])
def setup():
    settings = load_app_settings()
    active_tab = request.args.get("tab", "header")

    if request.method == "POST":
        active_tab = request.form.get("active_tab", "header").strip() or "header"

        try:
            if active_tab == "header":
                settings["header"]["club_small_title"] = request.form.get(
                    "club_small_title",
                    settings["header"]["club_small_title"]
                ).strip() or DEFAULT_APP_SETTINGS["header"]["club_small_title"]

                settings["header"]["club_name"] = request.form.get(
                    "club_name",
                    settings["header"]["club_name"]
                ).strip() or DEFAULT_APP_SETTINGS["header"]["club_name"]

                logo_file = request.files.get("logo_file")

                if logo_file and logo_file.filename:
                    filename = secure_filename(logo_file.filename)
                    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

                    if ext not in ALLOWED_PHOTO_EXTENSIONS:
                        flash("Logo chỉ nhận file jpg, jpeg, png, webp.", "danger")
                        return redirect(url_for("setup", tab="header"))

                    settings["header"]["logo_url"] = upload_club_asset_to_supabase(
                        logo_file,
                        "club-info/header"
                    )

            elif active_tab == "fees":
                settings["fees"]["tuition_fee"] = money_to_int_web(
                    request.form.get("tuition_fee", settings["fees"].get("tuition_fee", 500000))
                )

                settings["fees"]["exam_fee"] = money_to_int_web(
                    request.form.get("exam_fee", settings["fees"].get("exam_fee", 300000))
                )

                dan_fees = {}

                for dan_name in ["1 Đẳng", "2 Đẳng", "3 Đẳng", "4 Đẳng", "5 Đẳng", "6 Đẳng"]:
                    field_name = f"dan_fee_{dan_name.replace(' ', '_')}"
                    dan_fees[dan_name] = money_to_int_web(
                        request.form.get(
                            field_name,
                            settings["fees"].get("dan_fees", {}).get(dan_name, 0)
                        )
                    )

                settings["fees"]["dan_fees"] = dan_fees

            elif active_tab == "exam":
                settings["exam"]["exam_number_prefix"] = request.form.get(
                    "exam_number_prefix",
                    settings["exam"]["exam_number_prefix"]
                )

            elif active_tab == "club_info":
                settings["club_info"] = read_club_info_from_form(
                    request.form,
                    request.files,
                    settings.get("club_info", DEFAULT_APP_SETTINGS.get("club_info", {}))
                )

            elif active_tab == "payment":
                save_payment_settings(request.form)
                flash("Đã lưu thông tin thanh toán lên Supabase")
            save_app_settings(settings)
            flash("Đã lưu cài đặt.", "success")



        except Exception as e:
            print("[SETUP SAVE ERROR]", e)
            flash(f"Lỗi lưu cài đặt: {e}", "danger")

        return redirect(url_for("setup", tab=active_tab))

    get_class_options_settings(settings)

    return render_template(
        "setup.html",
        settings=settings,
        active_tab=active_tab,
        payment_settings=get_payment_settings()
    )

def ensure_student_portal_account(student):
    """
    Tạo tài khoản mặc định cho học viên nếu chưa có:
    username = Mã HV
    password = SĐT
    """
    if not student:
        return {}

    license_code = str(student.get("license") or "").strip()
    phone = str(student.get("phonenumber") or "").strip()

    if not license_code:
        return student

    update_data = {}

    if not str(student.get("portal_username") or "").strip():
        update_data["portal_username"] = license_code

    if not str(student.get("portal_password_hash") or "").strip() and phone:
        update_data["portal_password_hash"] = generate_password_hash(phone)

    if update_data:
        try:
            supabase.table(STUDENT_TABLE) \
                .update(update_data) \
                .eq("license", license_code) \
                .execute()

            student.update(update_data)
        except Exception as e:
            print("[PORTAL ACCOUNT INIT ERROR]", e)

    return student


def get_logged_student():
    license_code = str(session.get("student_license") or "").strip()

    if not license_code:
        return None

    rows = safe_rows(STUDENT_TABLE, "*", license=license_code)

    if not rows:
        session.pop("student_license", None)
        return None

    return rows[0]


def require_student_login():
    student = get_logged_student()

    if not student:
        return None

    return student


def get_student_notifications(student_license):
    student_license = str(student_license or "").strip()

    all_notifications = safe_rows(NOTIFICATION_TABLE)

    rows = []

    for n in all_notifications:
        target_type = str(n.get("target_type") or "all").strip()
        target_license = str(n.get("target_license") or "").strip()

        if target_type == "all" or target_license == student_license:
            rows.append(n)

    rows.sort(key=lambda r: str(r.get("created_at") or ""), reverse=True)

    reads = safe_rows(NOTIFICATION_READ_TABLE, "*", student_license=student_license)
    read_ids = {
        str(r.get("notification_id"))
        for r in reads
    }

    for n in rows:
        n["is_read"] = str(n.get("id")) in read_ids

    unread_count = len([n for n in rows if not n.get("is_read")])

    return rows, unread_count


def mark_notification_read(notification_id, student_license):
    notification_id = str(notification_id or "").strip()
    student_license = str(student_license or "").strip()

    if not notification_id or not student_license:
        return

    try:
        existing = supabase.table(NOTIFICATION_READ_TABLE) \
            .select("id") \
            .eq("notification_id", notification_id) \
            .eq("student_license", student_license) \
            .limit(1) \
            .execute().data or []

        if existing:
            return

        supabase.table(NOTIFICATION_READ_TABLE).insert({
            "notification_id": notification_id,
            "student_license": student_license
        }).execute()

    except Exception as e:
        print("[MARK NOTIFICATION READ ERROR]", e)

NOTIFICATION_CONTACT_FOOTER = (
    "Mọi thắc mắc xin liên hệ SĐT/Zalo: "
    "0963830315 - Tâm (Người ghi danh)."
)


def append_notification_footer(content):
    content = str(content or "").strip()

    if not content:
        return ""

    # Tránh bị chèn lặp nếu Ken đã nhập sẵn dòng này
    if "0963830315" in content and "Tâm" in content:
        return content

    return f"{content}\n\n{NOTIFICATION_CONTACT_FOOTER}"

def cleanup_old_notifications():
    """
    Tự dọn thông báo cũ để Supabase không phình dữ liệu.

    Quy tắc:
    - Thông báo đã đọc: xóa sau 15 ngày.
    - Thông báo chưa đọc: xóa sau 30 ngày.
    """

    now = datetime.now(timezone.utc)

    read_cutoff = (now - timedelta(days=15)).isoformat()
    unread_cutoff = (now - timedelta(days=30)).isoformat()

    try:
        # =========================
        # 1. Lấy các thông báo đã đọc
        # =========================
        read_rows = supabase.table(NOTIFICATION_READ_TABLE) \
            .select("notification_id") \
            .execute().data or []

        read_ids = []

        for r in read_rows:
            nid = r.get("notification_id")

            if nid is not None and nid not in read_ids:
                read_ids.append(nid)

        # =========================
        # 2. Xóa thông báo đã đọc quá 15 ngày
        # =========================
        if read_ids:
            supabase.table(NOTIFICATION_TABLE) \
                .delete() \
                .in_("id", read_ids) \
                .lt("created_at", read_cutoff) \
                .execute()

        # =========================
        # 3. Xóa thông báo chưa đọc quá 30 ngày
        # =========================
        old_notifications = supabase.table(NOTIFICATION_TABLE) \
            .select("id,created_at") \
            .lt("created_at", unread_cutoff) \
            .execute().data or []

        old_unread_ids = []

        for n in old_notifications:
            nid = n.get("id")

            if nid is None:
                continue

            # Nếu không nằm trong read_ids thì xem là chưa đọc
            if nid not in read_ids:
                old_unread_ids.append(nid)

        if old_unread_ids:
            supabase.table(NOTIFICATION_TABLE) \
                .delete() \
                .in_("id", old_unread_ids) \
                .execute()

        # =========================
        # 4. Xóa read logs mồ côi nếu notification đã mất
        # =========================
        existing_notifications = supabase.table(NOTIFICATION_TABLE) \
            .select("id") \
            .execute().data or []

        existing_ids = {
            str(n.get("id"))
            for n in existing_notifications
            if n.get("id") is not None
        }

        all_read_rows = supabase.table(NOTIFICATION_READ_TABLE) \
            .select("id,notification_id") \
            .execute().data or []

        orphan_read_ids = []

        for r in all_read_rows:
            notification_id = str(r.get("notification_id") or "")

            if notification_id and notification_id not in existing_ids:
                orphan_read_ids.append(r.get("id"))

        if orphan_read_ids:
            supabase.table(NOTIFICATION_READ_TABLE) \
                .delete() \
                .in_("id", orphan_read_ids) \
                .execute()

    except Exception as e:
        print("[CLEANUP NOTIFICATIONS ERROR]", e)

def personalize_notification_content(content, student_name="", is_all=False):
    content = str(content or "").strip()
    student_name = str(student_name or "").strip()

    if not content:
        return ""

    lines = content.splitlines()

    greeting_all = "Kính gửi quý Phụ huynh và võ sinh!"
    greeting_student = f"Kính gửi quý Phụ huynh và võ sinh: {student_name}" if student_name else "Kính gửi quý Phụ huynh và võ sinh:"

    greeting_regex = re.compile(r"^Kính gửi quý Phụ huynh và võ sinh[:!].*", re.IGNORECASE)

    if is_all:
        new_greeting = greeting_all
    else:
        new_greeting = greeting_student

    if lines:
        first_line = lines[0].strip()

        if greeting_regex.match(first_line):
            lines[0] = new_greeting
        else:
            lines.insert(0, new_greeting)
            lines.insert(1, "")
    else:
        lines = [new_greeting, ""]

    return "\n".join(lines).strip()

def build_notification_filter_ma_quy(target_mode, year, exam_quarter="", dan_round=""):
    """
    Tạo mã kỳ thi để lọc thông báo theo học phí đã đăng ký thi.
    - Thi cấp: Q12026, Q22026...
    - Thi đẳng: L1-2026, L2-2026, MN-2026...
    """
    target_mode = str(target_mode or "").strip()
    year = str(year or datetime.now().year).strip()
    exam_quarter = str(exam_quarter or "").strip().upper()
    dan_round = str(dan_round or "").strip().upper()

    if target_mode in ["exam_registered", "exam_not_registered"]:
        if not exam_quarter:
            return ""
        return f"{exam_quarter}{year}"

    if target_mode in ["dan_registered", "dan_not_registered"]:
        if not dan_round:
            return ""
        return f"{dan_round}-{year}"

    return ""


def get_students_by_notification_target(target_mode, students, year, exam_quarter="", dan_round=""):
    """
    Trả về danh sách học viên theo tùy chọn gửi thông báo:
    - exam_registered: đã đăng ký thi cấp kỳ chọn.
    - exam_not_registered: chưa đăng ký thi cấp kỳ chọn.
    - dan_registered: đã đăng ký thi đẳng kỳ/lần chọn.
    - dan_not_registered: chưa đăng ký thi đẳng kỳ/lần chọn.
    """
    target_mode = str(target_mode or "").strip()

    ma_quy = build_notification_filter_ma_quy(
        target_mode=target_mode,
        year=year,
        exam_quarter=exam_quarter,
        dan_round=dan_round
    )

    if not ma_quy:
        return [], ""

    fees = safe_rows(HOCPHI_TABLE, "ma_hv,ma_quy")

    registered_ids = {
        str(f.get("ma_hv") or "").strip()
        for f in fees
        if str(f.get("ma_quy") or "").strip().upper() == ma_quy.upper()
    }

    active_students = [
        s for s in students
        if is_active(s.get("active")) and str(s.get("license") or "").strip()
    ]

    if target_mode in ["exam_registered", "dan_registered"]:
        filtered_students = [
            s for s in active_students
            if str(s.get("license") or "").strip() in registered_ids
        ]

    elif target_mode in ["exam_not_registered", "dan_not_registered"]:
        filtered_students = [
            s for s in active_students
            if str(s.get("license") or "").strip() not in registered_ids
        ]

    else:
        filtered_students = []

    return filtered_students, ma_quy


def get_notification_target_label(target_mode, year, exam_quarter="", dan_round=""):
    target_mode = str(target_mode or "").strip()
    year = str(year or datetime.now().year).strip()
    exam_quarter = str(exam_quarter or "").strip().upper()
    dan_round = str(dan_round or "").strip().upper()

    exam_label_map = {
        "Q1": "Quý 1",
        "Q2": "Quý 2",
        "Q3": "Quý 3",
        "Q4": "Quý 4",
    }

    dan_label_map = {
        "L1": "Lần 1",
        "L2": "Lần 2",
        "MN": "KV miền Nam",
        "MT": "KV miền Trung",
        "MB": "KV miền Bắc",
        "QG": "Quốc Gia",
    }

    if target_mode == "exam_registered":
        return f"Thi cấp - Đã đăng ký - {exam_label_map.get(exam_quarter, exam_quarter)} - {year}"

    if target_mode == "exam_not_registered":
        return f"Thi cấp - Chưa đăng ký - {exam_label_map.get(exam_quarter, exam_quarter)} - {year}"

    if target_mode == "dan_registered":
        return f"Thi đẳng - Đã đăng ký - {dan_label_map.get(dan_round, dan_round)} - {year}"

    if target_mode == "dan_not_registered":
        return f"Thi đẳng - Chưa đăng ký - {dan_label_map.get(dan_round, dan_round)} - {year}"

    return "Tùy chọn"

@app.route("/notifications", methods=["GET", "POST"])
def notifications():
    cleanup_old_notifications()

    students = safe_rows(STUDENT_TABLE)
    students.sort(key=lambda s: str(s.get("name") or ""))

    now_year = datetime.now().year
    year_options = [now_year - 1, now_year, now_year + 1]

    if request.method == "POST":
        target_mode = request.form.get("target_mode", "all").strip()
        target_licenses = request.form.getlist("target_licenses")
        title = request.form.get("title", "").strip()
        content = request.form.get("content", "").strip()

        notify_year = request.form.get("notify_year", str(now_year)).strip() or str(now_year)
        exam_quarter = request.form.get("exam_quarter", "").strip().upper()
        dan_round = request.form.get("dan_round", "").strip().upper()

        target_licenses = [
            str(x or "").strip()
            for x in target_licenses
            if str(x or "").strip()
        ]

        if not title:
            flash("Ken chưa nhập tiêu đề thông báo.", "danger")
            return redirect(url_for("notifications"))

        if not content:
            flash("Ken chưa nhập nội dung thông báo.", "danger")
            return redirect(url_for("notifications"))

        try:
            # =========================
            # GỬI CHO TẤT CẢ
            # =========================
            if target_mode == "all":
                final_content = personalize_notification_content(
                    content,
                    student_name="",
                    is_all=True
                )

                payload = {
                    "target_type": "all",
                    "target_license": None,
                    "target_name": "Tất cả học viên",
                    "title": title,
                    "content": append_notification_footer(final_content),
                }

                supabase.table(NOTIFICATION_TABLE).insert(payload).execute()
                flash("Đã gửi thông báo cho tất cả học viên.", "success")
                return redirect(url_for("notifications"))

            student_map = {
                str(s.get("license") or "").strip(): s
                for s in students
            }

            payloads = []

            # =========================
            # GỬI THEO NHÓM THI CẤP / THI ĐẲNG
            # =========================
            group_target_modes = [
                "exam_registered",
                "exam_not_registered",
                "dan_registered",
                "dan_not_registered",
            ]

            if target_mode in group_target_modes:
                if target_mode in ["exam_registered", "exam_not_registered"] and not exam_quarter:
                    flash("Ken chưa chọn quý thi cấp.", "danger")
                    return redirect(url_for("notifications"))

                if target_mode in ["dan_registered", "dan_not_registered"] and not dan_round:
                    flash("Ken chưa chọn lần/khu vực thi đẳng.", "danger")
                    return redirect(url_for("notifications"))

                target_students, ma_quy = get_students_by_notification_target(
                    target_mode=target_mode,
                    students=students,
                    year=notify_year,
                    exam_quarter=exam_quarter,
                    dan_round=dan_round
                )

                target_label = get_notification_target_label(
                    target_mode=target_mode,
                    year=notify_year,
                    exam_quarter=exam_quarter,
                    dan_round=dan_round
                )

                if not target_students:
                    flash(f"Không có học viên phù hợp với nhóm: {target_label}.", "danger")
                    return redirect(url_for("notifications"))

                for student in target_students:
                    license_code = str(student.get("license") or "").strip()
                    student_name = student.get("name", "")

                    final_content = personalize_notification_content(
                        content,
                        student_name=student_name,
                        is_all=False
                    )

                    payloads.append({
                        "target_type": target_mode,
                        "target_license": license_code,
                        "target_name": student_name,
                        "title": title,
                        "content": append_notification_footer(final_content),
                    })

                supabase.table(NOTIFICATION_TABLE).insert(payloads).execute()

                flash(
                    f"Đã gửi thông báo cho {len(payloads)} học viên thuộc nhóm: {target_label}.",
                    "success"
                )
                return redirect(url_for("notifications"))

            # =========================
            # GỬI TÙY CHỌN NHIỀU HỌC VIÊN
            # =========================
            if not target_licenses:
                flash("Ken chưa chọn học viên nhận thông báo.", "danger")
                return redirect(url_for("notifications"))

            for license_code in target_licenses:
                student = student_map.get(license_code)

                if not student:
                    continue

                student_name = student.get("name", "")

                final_content = personalize_notification_content(
                    content,
                    student_name=student_name,
                    is_all=False
                )

                payloads.append({
                    "target_type": "student",
                    "target_license": license_code,
                    "target_name": student_name,
                    "title": title,
                    "content": append_notification_footer(final_content),
                })

            if not payloads:
                flash("Không tìm thấy học viên hợp lệ để gửi thông báo.", "danger")
                return redirect(url_for("notifications"))

            supabase.table(NOTIFICATION_TABLE).insert(payloads).execute()

            flash(f"Đã gửi thông báo cho {len(payloads)} học viên.", "success")
            return redirect(url_for("notifications"))

        except Exception as e:
            print("[ADD NOTIFICATION ERROR]", e)
            flash(f"Lỗi gửi thông báo: {e}", "danger")
            return redirect(url_for("notifications"))

    notifications_rows = safe_rows(NOTIFICATION_TABLE)
    notifications_rows.sort(key=lambda r: str(r.get("created_at") or ""), reverse=True)

    return render_template(
        "notifications.html",
        rows=students,
        notifications=notifications_rows,
        current_year=now_year,
        year_options=year_options
    )
    
@app.route("/student-login", methods=["GET", "POST"])
def student_login():
    if request.method == "POST":
        action = request.form.get("action", "login").strip()

        # =========================
        # QUÊN MẬT KHẨU HỌC VIÊN
        # Reset về mặc định:
        # username = Mã HV hiện tại
        # password = SĐT đăng ký
        # =========================
        if action == "forgot_password":
            username = request.form.get("reset_username", "").strip()
            birthdate_raw = request.form.get("reset_birthdate", "").strip()
            phone_raw = request.form.get("reset_phone", "").strip()

            birthdate = normalize_birthdate_web(birthdate_raw)
            phone = re.sub(r"\D", "", phone_raw)

            if not username or not birthdate or not phone:
                flash("Vui lòng nhập đủ Mã HV/Username, ngày sinh và SĐT đăng ký.", "danger")
                return redirect(url_for("student_login"))

            students = safe_rows(STUDENT_TABLE)

            matched = None

            for s in students:
                license_code = str(s.get("license") or "").strip()
                portal_username = str(s.get("portal_username") or "").strip()
                student_birthdate = str(s.get("birthdate") or "").strip()
                student_phone = re.sub(r"\D", "", str(s.get("phonenumber") or ""))

                same_account = username in [license_code, portal_username]
                same_birthdate = birthdate == student_birthdate
                same_phone = phone == student_phone

                if same_account and same_birthdate and same_phone:
                    matched = s
                    break

            if not matched:
                flash("Thông tin xác minh chưa đúng. Ken kiểm tra lại Mã HV, ngày sinh hoặc SĐT đăng ký.", "danger")
                return redirect(url_for("student_login"))

            license_code = str(matched.get("license") or "").strip()
            student_phone = re.sub(r"\D", "", str(matched.get("phonenumber") or ""))

            if not license_code or not student_phone:
                flash("Tài khoản thiếu Mã HV hoặc SĐT nên chưa thể reset mật khẩu.", "danger")
                return redirect(url_for("student_login"))

            supabase.table(STUDENT_TABLE) \
                .update({
                    "portal_username": license_code,
                    "portal_password_hash": generate_password_hash(student_phone)
                }) \
                .eq("license", license_code) \
                .execute()

            flash("Đã cài lại tài khoản về mặc định. Tên đăng nhập = Mã HV, mật khẩu = SĐT.", "success")
            return redirect(url_for("student_login"))

        # =========================
        # ĐĂNG NHẬP BÌNH THƯỜNG
        # =========================
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        if not username or not password:
            flash("Vui lòng nhập tên đăng nhập và mật khẩu.", "danger")
            return redirect(url_for("student_login"))

        students = safe_rows(STUDENT_TABLE)

        matched = None

        for s in students:
            ensure_student_portal_account(s)

            portal_username = str(s.get("portal_username") or "").strip()
            license_code = str(s.get("license") or "").strip()
            phone = str(s.get("phonenumber") or "").strip()

            # Cho phép đăng nhập bằng username, mã HV hoặc SĐT
            if username in [portal_username, license_code, phone]:
                matched = s
                break

        if not matched:
            flash("Không tìm thấy tài khoản học viên.", "danger")
            return redirect(url_for("student_login"))

        password_hash = str(matched.get("portal_password_hash") or "").strip()

        # Nếu học viên cũ chưa có hash, dùng SĐT làm mật khẩu lần đầu
        if not password_hash:
            matched = ensure_student_portal_account(matched)
            password_hash = str(matched.get("portal_password_hash") or "").strip()

        if not check_password_hash(password_hash, password):
            flash("Mật khẩu không đúng.", "danger")
            return redirect(url_for("student_login"))

        session["student_logged_in"] = True
        session["student_license"] = matched.get("license")
        session["student_name"] = matched.get("name")
        session["show_student_welcome_popup"] = True

        flash("Đăng nhập thành công.", "success")
        return redirect(url_for("student_portal_notifications"))

    return render_template("student_login.html")


@app.get("/student-logout")
def student_logout():
    session.pop("student_logged_in", None)
    session.pop("student_license", None)
    session.pop("student_name", None)
    session.pop("show_student_welcome_popup", None)

    flash("Đã đăng xuất.", "success")
    return redirect(url_for("student_login"))

# =========================
# COACH PORTAL - HLV
# =========================

def normalize_phone_web(value):
    return re.sub(r"\D", "", str(value or ""))


def is_active_text(value):
    text = remove_accents(str(value or "")).strip().lower()
    return text in ["co", "yes", "true", "1", "active", "dang hoat dong"]


def get_logged_coach():
    coach_code = str(session.get("coach_code") or "").strip()

    if not coach_code:
        return None

    try:
        rows = supabase.table(COACH_TABLE) \
            .select("*") \
            .eq("coach_code", coach_code) \
            .limit(1) \
            .execute().data or []

        if not rows:
            session.pop("coach_logged_in", None)
            session.pop("coach_code", None)
            return None

        coach = rows[0]

        if not is_active_text(coach.get("active")):
            session.pop("coach_logged_in", None)
            session.pop("coach_code", None)
            return None

        return coach

    except Exception as e:
        print("[GET LOGGED COACH ERROR]", e)
        return None


def require_coach_login():
    if not session.get("coach_logged_in"):
        return None

    return get_logged_coach()


@app.route("/coach-login", methods=["GET", "POST"])
def coach_login():
    if request.method == "POST":
        coach_code = request.form.get("coach_code", "").strip()
        phone = normalize_phone_web(request.form.get("phone", ""))

        if not coach_code or not phone:
            flash("Vui lòng nhập Mã HLV và SĐT.", "danger")
            return redirect(url_for("coach_login"))

        try:
            rows = supabase.table(COACH_TABLE) \
                .select("*") \
                .eq("coach_code", coach_code) \
                .limit(1) \
                .execute().data or []

            if not rows:
                flash("Không tìm thấy Mã HLV.", "danger")
                return redirect(url_for("coach_login"))

            coach = rows[0]

            if not is_active_text(coach.get("active")):
                flash("Tài khoản HLV này đang bị khóa.", "danger")
                return redirect(url_for("coach_login"))

            db_phone = normalize_phone_web(coach.get("phone"))

            if not db_phone or db_phone != phone:
                flash("SĐT đăng nhập chưa đúng.", "danger")
                return redirect(url_for("coach_login"))

            # Không dùng session.clear(), vì sẽ làm out Admin và Student
            session["coach_logged_in"] = True
            session["coach_code"] = coach.get("coach_code")
            session["coach_name"] = coach.get("name")

            flash("Đăng nhập HLV thành công.", "success")
            return redirect(url_for("coach_exam"))

        except Exception as e:
            print("[COACH LOGIN ERROR]", e)
            flash(f"Lỗi đăng nhập HLV: {e}", "danger")
            return redirect(url_for("coach_login"))

    return render_template("coach_login.html")


@app.get("/coach-logout")
def coach_logout():
    session.pop("coach_logged_in", None)
    session.pop("coach_code", None)
    session.pop("coach_name", None)

    flash("Đã đăng xuất HLV.", "success")
    return redirect(url_for("coach_login"))


@app.get("/coach/exam")
def coach_exam():
    coach = require_coach_login()

    if not coach:
        return redirect(url_for("coach_login"))

    now = datetime.now()

    year = request.args.get("year", str(now.year)).strip()
    quarter = request.args.get(
        "quarter",
        f"Q{(now.month - 1) // 3 + 1}"
    ).strip().upper()

    if quarter not in ["Q1", "Q2", "Q3", "Q4"]:
        quarter = f"Q{(now.month - 1) // 3 + 1}"

    rows = get_exam_list_rows_by_type_web(
        year=year,
        quarter=quarter,
        list_type="cap"
    )

    rows.sort(key=lambda r: (
        get_belt_index_web(r.get("cap_du_thi")),
        str(r.get("gender") or ""),
        str(r.get("name") or "")
    ))

    status_summary = {
        "total": len(rows),
        "Đúng": 0,
        "Sai": 0,
        "Ktra lại": 0,
        "Chưa KTra": 0,
    }

    for r in rows:
        status = str(r.get("coach_check_status") or "").strip()

        if not status:
            status = "Chưa KTra"

        if status not in status_summary:
            status_summary[status] = 0

        status_summary[status] += 1

    years = [str(y) for y in range(now.year - 2, now.year + 3)]
    quarters = ["Q1", "Q2", "Q3", "Q4"]

    return render_template(
        "coach_exam.html",
        coach=coach,
        rows=rows,
        year=year,
        quarter=quarter,
        years=years,
        quarters=quarters,
        status_summary=status_summary
    )

@app.get("/coach/dan")
def coach_dan():
    coach = require_coach_login()

    if not coach:
        return redirect(url_for("coach_login"))

    now = datetime.now()

    year = request.args.get("year", str(now.year)).strip()

    # Nhận cả 2 tên để tránh lỗi HTML cũ / mới:
    # - quarter: đang dùng trong coach_dan.html
    # - dan_round: dự phòng nếu sau này đổi tên input
    dan_round = (
        request.args.get("quarter")
        or request.args.get("dan_round")
        or "L1"
    )
    dan_round = str(dan_round or "L1").strip().upper()

    dan_round_options = [
        {"value": "L1", "label": "Lần 1"},
        {"value": "L2", "label": "Lần 2"},
        {"value": "MN", "label": "KV miền Nam"},
        {"value": "MT", "label": "KV miền Trung"},
        {"value": "MB", "label": "KV miền Bắc"},
        {"value": "QG", "label": "Quốc Gia"},
    ]

    dan_round_values = [
        item["value"]
        for item in dan_round_options
    ]

    if dan_round not in dan_round_values:
        dan_round = "L1"

    quarter_label = next(
        (
            item["label"]
            for item in dan_round_options
            if item["value"] == dan_round
        ),
        dan_round
    )

    rows = get_exam_list_rows_by_type_web(
        year=year,
        quarter=dan_round,
        list_type="dan"
    )

    rows.sort(key=lambda r: (
        get_belt_index_web(r.get("cap_du_thi")),
        str(r.get("gender") or ""),
        str(r.get("name") or "")
    ))

    status_summary = {
        "total": len(rows),
        "Đúng": 0,
        "Sai": 0,
        "Ktra lại": 0,
        "Chưa KTra": 0,
    }

    for r in rows:
        status = str(r.get("coach_check_status") or "").strip()

        if not status:
            status = "Chưa KTra"

        if status not in status_summary:
            status_summary[status] = 0

        status_summary[status] += 1

    years = [str(y) for y in range(now.year - 2, now.year + 3)]

    return render_template(
        "coach_dan.html",
        coach=coach,
        rows=rows,
        year=year,

        # Biến dùng cho coach_dan.html hiện tại
        quarter=dan_round,
        quarters=dan_round_options,
        quarter_label=quarter_label,

        # Biến dự phòng nếu sau này HTML dùng dan_round
        dan_round=dan_round,
        dan_round_options=dan_round_options,

        years=years,
        status_summary=status_summary
    )

@app.post("/coach/exam/update-status")
def coach_exam_update_status():
    coach = require_coach_login()

    if not coach:
        return jsonify({
            "ok": False,
            "message": "HLV chưa đăng nhập."
        }), 401

    data = request.get_json(silent=True) or {}

    hocphi_id = str(data.get("hocphi_id") or "").strip()
    status = str(data.get("status") or "").strip()
    note = str(data.get("note") or "").strip()

    allowed_status = ["Chưa KTra", "Đúng", "Sai", "Ktra lại"]

    if not hocphi_id:
        return jsonify({
            "ok": False,
            "message": "Thiếu ID học phí."
        }), 400

    if status not in allowed_status:
        return jsonify({
            "ok": False,
            "message": "Trạng thái không hợp lệ."
        }), 400

    try:
        payload = {
            "coach_check_status": status,
            "coach_check_note": note,
            "coach_checked_by": str(coach.get("coach_code") or ""),
            "coach_checked_by_name": str(coach.get("name") or ""),
            "coach_checked_at": datetime.now(timezone.utc).isoformat(),
        }

        supabase.table(HOCPHI_TABLE) \
            .update(payload) \
            .eq("id", hocphi_id) \
            .execute()

        return jsonify({
            "ok": True,
            "message": "Đã lưu kiểm tra HLV.",
            "status": status,
            "checked_by_name": payload["coach_checked_by_name"],
            "checked_at": payload["coach_checked_at"],
        })

    except Exception as e:
        print("[COACH EXAM UPDATE STATUS ERROR]", e)

        return jsonify({
            "ok": False,
            "message": str(e)
        }), 500

@app.get("/student-portal")
def student_portal_home():
    return redirect(url_for("student_portal_notifications"))


@app.get("/student-portal/notifications")
def student_portal_notifications():
    cleanup_old_notifications()

    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    notifications_rows, unread_count = get_student_notifications(student.get("license"))

    return render_template(
        "student_portal_notifications.html",
        student=student,
        notifications=notifications_rows,
        unread_count=unread_count
    )


@app.get("/student-portal/notifications/<notification_id>")
def student_portal_notification_detail(notification_id):
    cleanup_old_notifications()

    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    rows = safe_rows(NOTIFICATION_TABLE, "*", id=notification_id)
    notification = rows[0] if rows else None

    if not notification:
        flash("Không tìm thấy thông báo.", "danger")
        return redirect(url_for("student_portal_notifications"))

    notification["content"] = append_notification_footer(notification.get("content", ""))

    target_type = str(notification.get("target_type") or "all")
    target_license = str(notification.get("target_license") or "")

    if target_type != "all" and target_license != str(student.get("license")):
        flash("Thông báo này không thuộc tài khoản của bạn.", "danger")
        return redirect(url_for("student_portal_notifications"))

    mark_notification_read(notification_id, student.get("license"))

    notifications_rows, unread_count = get_student_notifications(student.get("license"))

    return render_template(
        "student_portal_notification_detail.html",
        student=student,
        notification=notification,
        unread_count=unread_count
    )


@app.get("/student-portal/info")
def student_portal_info():
    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    notifications_rows, unread_count = get_student_notifications(student.get("license"))

    return render_template(
        "student_portal_info.html",
        student=student,
        unread_count=unread_count,
        photo_url=get_student_photo_url(student.get("license"))
    )

@app.get("/student-portal/club-info")
def student_portal_club_info():
    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    license_code = student.get("license")
    notifications_rows, unread_count = get_student_notifications(license_code)

    settings = load_app_settings()
    club_info = settings.get("club_info", DEFAULT_APP_SETTINGS.get("club_info", {}))

    return render_template(
        "student_portal_club_info.html",
        student=student,
        unread_count=unread_count,
        club_info=club_info
    )

@app.get("/student-portal/fees")
def student_portal_fees():
    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    license_code = student.get("license")

    fees = safe_rows(HOCPHI_TABLE, "*", ma_hv=license_code)
    fees.sort(key=lambda r: str(r.get("thoi_gian") or ""), reverse=True)

    def is_exam_fee_row(f):
        title = str(f.get("thang_dong_phi") or "").lower()
        ma_quy = str(f.get("ma_quy") or "").strip()

        return bool(
            ma_quy
            or "thi cấp" in title
            or "thi đẳng" in title
        )


    def format_month_code_for_portal(raw):
        raw = str(raw or "").strip()

        if not raw:
            return "Chưa có"

        # Nếu có nhiều tháng: 062026 - 072026
        first_code = raw.split("-")[0].strip()

        if re.fullmatch(r"\d{6}", first_code):
            return f"{first_code[:2]}/{first_code[2:]}"

        return raw


    def format_exam_code_for_portal(f):
        ma_quy = str(f.get("ma_quy") or "").strip()
        title = str(f.get("thang_dong_phi") or "").strip()

        if title:
            # VD: Thi cấp quý 2-2026
            return title.replace("Thi cấp quý ", "Q").replace("-2026", "/2026")

        if ma_quy:
            return format_ma_quy_label_web(ma_quy)

        return "Chưa có"


    tuition_fees = [
        f for f in fees
        if not is_exam_fee_row(f)
    ]

    exam_fees = [
        f for f in fees
        if is_exam_fee_row(f)
    ]

    latest_tuition_text = "Chưa có"
    latest_exam_text = "Chưa có"

    if tuition_fees:
        latest_tuition_text = format_month_code_for_portal(
            tuition_fees[0].get("ma_thang") or tuition_fees[0].get("thang_dong_phi")
        )

    if exam_fees:
        latest_exam_text = format_exam_code_for_portal(exam_fees[0])

    tuition_count = len(tuition_fees)
    exam_count = len(exam_fees)
    # =========================
    # NHẮC HẠN ĐÓNG HỌC PHÍ
    # Quy tắc:
    # - Nếu đã đóng tháng hiện tại: báo đã đóng.
    # - Nếu có đóng tháng trước: lấy ngày đóng tháng trước làm hạn tháng này.
    # - Nếu không có tháng trước: xem là võ sinh vãng lai / chưa có dữ liệu liên tiếp.
    # =========================
    today = date.today()
    current_code = f"{today.month:02d}{today.year}"

    if today.month == 1:
        prev_month = 12
        prev_year = today.year - 1
    else:
        prev_month = today.month - 1
        prev_year = today.year

    prev_code = f"{prev_month:02d}{prev_year}"

    fee_due_info = {
        "status": "casual",
        "status_class": "fee-due-casual",
        "title": "Chưa có dữ liệu tháng trước",
        "main": "Võ sinh tạm ngưng tập",
        "sub": "Nhanh chóng đóng phí để tiếp tục tập luyện nhé.",
        "paid_date": "—",
        "due_date": "—",
        "days_text": "—",
    }

    current_month_fees = []
    prev_month_fees = []

    for f in fees:
        ma_thang = str(f.get("ma_thang") or "")

        if current_code in ma_thang:
            current_month_fees.append(f)

        if prev_code in ma_thang:
            prev_month_fees.append(f)

    def parse_fee_dt(raw):
        try:
            return datetime.fromisoformat(str(raw or "").replace("Z", "+00:00"))
        except:
            return None

    if current_month_fees:
        latest_current = current_month_fees[0]
        paid_dt = parse_fee_dt(latest_current.get("thoi_gian"))

        fee_due_info = {
            "status": "paid",
            "status_class": "fee-due-green",
            "title": "Đã đóng tháng này",
            "main": "Học phí tháng hiện tại đã hoàn tất",
            "sub": "Cảm ơn phụ huynh và võ sinh đã đồng hành cùng CLB.",
            "paid_date": paid_dt.strftime("%d/%m/%Y") if paid_dt else "—",
            "due_date": "Đã đóng",
            "days_text": "Đã hoàn tất",
        }

    elif prev_month_fees:
        latest_prev = prev_month_fees[0]
        paid_dt = parse_fee_dt(latest_prev.get("thoi_gian"))

        if paid_dt:
            last_day = calendar.monthrange(today.year, today.month)[1]
            due_day = min(paid_dt.day, last_day)
            due_date = date(today.year, today.month, due_day)

            days_left = (due_date - today).days

            if days_left >= 8:
                status_class = "fee-due-green"
                title = "Chưa đến hạn"
                days_text = f"Còn {days_left} ngày"
                sub = "Học phí vẫn còn hạn."

            elif 4 <= days_left <= 7:
                status_class = "fee-due-yellow"
                title = "Gần đến hạn"
                days_text = f"Còn {days_left} ngày"
                sub = "Sắp đến hạn đóng học phí, phụ huynh/võ sinh lưu ý giúp CLB."

            elif 0 <= days_left <= 3:
                status_class = "fee-due-orange"
                title = "Sắp đến hạn"
                days_text = f"Còn {days_left} ngày"
                sub = "Phụ huynh/võ sinh vui lòng chuẩn bị đóng phí để không gián đoạn tập luyện."

            else:
                status_class = "fee-due-red"
                title = "Quá hạn"
                days_text = f"Quá hạn {abs(days_left)} ngày"
                sub = "Phụ huynh/võ sinh nhanh chóng đóng phí để tiếp tục tập luyện nhé."

            fee_due_info = {
                "status": "normal",
                "status_class": status_class,
                "title": title,
                "main": days_text,
                "sub": sub,
                "paid_date": paid_dt.strftime("%d/%m/%Y"),
                "due_date": due_date.strftime("%d/%m/%Y"),
                "days_text": days_text,
            }

    notifications_rows, unread_count = get_student_notifications(license_code)

    payment_info = get_payment_settings()

    payment_transfer_note = build_transfer_note(
        payment_info.get("transfer_note", "{student_name}"),
        student
    )

    payment_qr_url = build_vietqr_url(payment_info, student)

    return render_template(
        "student_portal_fees.html",
        student=student,
        fees=fees,
        payment_info=payment_info,
        payment_qr_url=payment_qr_url,
        payment_transfer_note=payment_transfer_note,
        fee_due_info=fee_due_info,
        latest_tuition_text=latest_tuition_text,
        latest_exam_text=latest_exam_text,
        tuition_count=tuition_count,
        exam_count=exam_count
    )

def format_ma_quy_label_web(ma_quy):
    """
    Q22026 -> Quý 2 - 2026
    Q2-2026 -> Quý 2 - 2026
    2026-Q2 -> Quý 2 - 2026
    """
    s = str(ma_quy or "").strip().upper()

    if not s:
        return "Sẽ bổ sung"

    m = re.search(r"Q([1-4])\D*(20\d{2})", s)
    if m:
        return f"Quý {m.group(1)} - {m.group(2)}"

    m = re.search(r"(20\d{2})\D*Q([1-4])", s)
    if m:
        return f"Quý {m.group(2)} - {m.group(1)}"

    return s


def ma_quy_to_ky_thi_web(ma_quy):
    """
    Q22026 -> 2026-Q2
    Q2-2026 -> 2026-Q2
    2026-Q2 -> 2026-Q2
    """
    s = str(ma_quy or "").strip().upper()

    if not s:
        return ""

    m = re.search(r"Q([1-4])\D*(20\d{2})", s)
    if m:
        return f"{m.group(2)}-Q{m.group(1)}"

    m = re.search(r"(20\d{2})\D*Q([1-4])", s)
    if m:
        return f"{m.group(1)}-Q{m.group(2)}"

    return s


@app.get("/student-portal/exams")
def student_portal_exams():
    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    license_code = student.get("license")

    # =========================
    # KẾT QUẢ ĐÃ THI
    # =========================
    results = safe_rows(KETQUA_TABLE, "*", ma_hv=license_code)
    results.sort(key=lambda r: parse_ky_thi_web(r.get("ky_thi")), reverse=True)

    for r in results:
        ky = str(r.get("ky_thi") or "").strip()
        info = get_exam_info_web(ky)

        judges = info.get("judges", []) if info else []
        judge_names = []

        for j in judges:
            if not isinstance(j, dict):
                continue

            name = str(j.get("name") or "").strip()
            active = str(j.get("active") or "on").strip().lower()

            if name:
                status = "ON" if active == "on" else "OFF"
                judge_names.append(f"{name} ({status})")

        r["exam_date_text"] = info.get("exam_date") or "Sẽ bổ sung"
        r["exam_time_text"] = info.get("exam_time") or "Sẽ bổ sung"
        r["venue_text"] = info.get("venue") or "Sẽ bổ sung"

        supervisor_name = info.get("supervisor_name") or ""
        supervisor_active = str(info.get("supervisor_active") or "on").lower()
        supervisor_status = "ON" if supervisor_active == "on" else "OFF"

        r["supervisor_text"] = (
            f"{supervisor_name} ({supervisor_status})"
            if supervisor_name else "Sẽ bổ sung"
        )

        r["judges_text"] = " | ".join(judge_names) if judge_names else "Sẽ bổ sung"

    result_ky_set = {
        str(r.get("ky_thi") or "").strip().upper()
        for r in results
        if str(r.get("ky_thi") or "").strip()
    }

    # =========================
    # TÍNH CẤP DỰ THI TIẾP THEO CHO HỌC VIÊN
    # Logic đúng:
    # - Nếu kết quả mới nhất Đạt: cấp tiếp theo = cấp kế tiếp của cấp đã đạt
    # - Nếu Không đạt hoặc Vắng: thi lại đúng cấp đó
    # - Nếu chưa có kết quả: lấy cấp hiện tại trong student để tính cấp tiếp theo
    # =========================
    current_belt = normalize_belt_name_web(student.get("belt"))
    next_belt = get_next_belt_web(current_belt)

    latest_result = results[0] if results else None

    if latest_result:
        latest_exam_belt = normalize_belt_name_web(latest_result.get("cap_dai_thi"))
        latest_exam_result = str(latest_result.get("ket_qua") or "").strip()

        if latest_exam_result == "Đạt":
            next_belt = get_next_belt_web(latest_exam_belt)

        elif latest_exam_result in ["Không đạt", "Vắng"]:
            next_belt = latest_exam_belt

    fee_rows = safe_rows(HOCPHI_TABLE, "*", ma_hv=license_code)

    exam_fee_rows = []

    for f in fee_rows:
        ma_quy = str(f.get("ma_quy") or "").strip()

        if not ma_quy:
            continue

        ky_thi = ma_quy_to_ky_thi_web(ma_quy)

        if not ky_thi:
            continue

        exam_fee_rows.append({
            "raw": f,
            "ma_quy": ma_quy,
            "ky_thi": ky_thi,
            "label": format_ma_quy_label_web(ma_quy),
            "order": parse_ky_thi_web(ky_thi),
            "thoi_gian": str(f.get("thoi_gian") or "")
        })

    # Lấy kỳ thi đã đăng ký mới nhất
    exam_fee_rows.sort(
        key=lambda x: (
            int(x.get("order") or 0),
            str(x.get("thoi_gian") or "")
        ),
        reverse=True
    )

    selected_exam_row = exam_fee_rows[0] if exam_fee_rows else None

    not_registered_text = "Chưa đăng ký kỳ mới"

    if selected_exam_row:
        next_ky_thi = selected_exam_row["ky_thi"]
        next_quarter = selected_exam_row["label"]

        # Kiểm tra kỳ đăng ký mới nhất đã có kết quả chưa
        matched_result = None

        for r in results:
            if str(r.get("ky_thi") or "").strip().upper() == next_ky_thi.upper():
                matched_result = r
                break

        if matched_result:
            # Nếu kỳ đã có kết quả rồi thì không xem là kỳ sắp thi nữa.
            # Chuyển sang hiển thị cấp tiếp theo sau kết quả.
            next_status = "Chưa đăng ký kỳ mới"
            next_quarter = not_registered_text
            next_ky_thi = ""
            next_belt_display = next_belt
        else:
            # Nếu đã đóng phí kỳ mới nhưng chưa có kết quả thì đây là kỳ sắp thi.
            next_status = "Đã đăng ký thi"
            next_belt_display = next_belt

    else:
        next_ky_thi = ""
        next_quarter = not_registered_text
        next_status = not_registered_text
        next_belt_display = next_belt

    next_exam_info = get_exam_info_web(next_ky_thi) if next_ky_thi else {}

    next_exam_info = get_exam_info_web(next_ky_thi) if next_ky_thi else {}

    next_judges = next_exam_info.get("judges", []) if next_exam_info else []
    next_judge_names = []

    for j in next_judges:
        if not isinstance(j, dict):
            continue

        name = str(j.get("name") or "").strip()
        active = str(j.get("active") or "on").lower().strip()

        if name:
            status = "ON" if active == "on" else "OFF"
            next_judge_names.append(f"{name} ({status})")

    next_supervisor_name = next_exam_info.get("supervisor_name") or ""
    next_supervisor_active = str(next_exam_info.get("supervisor_active") or "on").lower().strip()
    next_supervisor_status = "ON" if next_supervisor_active == "on" else "OFF"

    next_exam = {
        "cap_du_thi": next_belt_display,
        "status": next_status,
        "quarter": next_quarter,
        "exam_date": next_exam_info.get("exam_date") or not_registered_text,
        "exam_time": next_exam_info.get("exam_time") or not_registered_text,
        "venue": next_exam_info.get("venue") or not_registered_text,
        "supervisor": (
            f"{next_supervisor_name} ({next_supervisor_status})"
            if next_supervisor_name else not_registered_text
        ),
        "judges": " | ".join(next_judge_names) if next_judge_names else not_registered_text,
    }

    notifications_rows, unread_count = get_student_notifications(license_code)

    return render_template(
        "student_portal_exams.html",
        student=student,
        results=results,
        next_exam=next_exam,
        unread_count=unread_count
    )

@app.get("/student-portal/activities")
def student_portal_activities():
    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    license_code = student.get("license")

    activities = safe_rows(HOATDONG_TABLE, "*", ma_hv=license_code)
    activities.sort(key=lambda r: str(r.get("thoi_gian") or ""), reverse=True)

    notifications_rows, unread_count = get_student_notifications(license_code)

    return render_template(
        "student_portal_activities.html",
        student=student,
        activities=activities,
        unread_count=unread_count
    )


@app.route("/student-portal/settings", methods=["GET", "POST"])
def student_portal_settings():
    student = require_student_login()
    if not student:
        return redirect(url_for("student_login"))

    license_code = student.get("license")

    if request.method == "POST":
        new_username = request.form.get("portal_username", "").strip()
        old_password = request.form.get("old_password", "").strip()
        new_password = request.form.get("new_password", "").strip()
        confirm_password = request.form.get("confirm_password", "").strip()

        update_data = {}

        if new_username:
            duplicated = supabase.table(STUDENT_TABLE) \
                .select("license,portal_username") \
                .eq("portal_username", new_username) \
                .neq("license", license_code) \
                .limit(1) \
                .execute().data or []

            if duplicated:
                flash("Tên đăng nhập này đã có người dùng.", "danger")
                return redirect(url_for("student_portal_settings"))

            update_data["portal_username"] = new_username

        if old_password or new_password or confirm_password:
            password_hash = str(student.get("portal_password_hash") or "").strip()

            if not check_password_hash(password_hash, old_password):
                flash("Mật khẩu cũ không đúng.", "danger")
                return redirect(url_for("student_portal_settings"))

            if not new_password:
                flash("Ken chưa nhập mật khẩu mới.", "danger")
                return redirect(url_for("student_portal_settings"))

            if new_password != confirm_password:
                flash("Xác nhận mật khẩu mới không khớp.", "danger")
                return redirect(url_for("student_portal_settings"))

            update_data["portal_password_hash"] = generate_password_hash(new_password)

        if update_data:
            supabase.table(STUDENT_TABLE) \
                .update(update_data) \
                .eq("license", license_code) \
                .execute()

            flash("Đã cập nhật tài khoản.", "success")

        return redirect(url_for("student_portal_settings"))

    notifications_rows, unread_count = get_student_notifications(license_code)

    return render_template(
        "student_portal_settings.html",
        student=student,
        unread_count=unread_count
    )


if __name__ == '__main__':
    app.run(
        host='0.0.0.0',
        port=5050,
        debug=False,
        use_reloader=False
    )
