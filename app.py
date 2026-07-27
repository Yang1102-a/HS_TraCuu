import streamlit as st
import pandas as pd
import re, json, html, unicodedata
import bcrypt
from io import BytesIO
from supabase import create_client

st.set_page_config(page_title="Tra cứu mã HS", layout="wide")

ADMIN_USER = "admin"
ADMIN_PASS_DEFAULT = "000"

BUCKET_NAME = "userdata"
SETTINGS_ROW_ID = 1
HISTORY_LIMIT = 30

st.markdown("""
<style>
.stApp { background:#0f1117; color:#eee; }
.card { border:1px solid #333846; border-radius:16px; padding:16px; margin-bottom:12px; background:#171a22; }
.hs { font-size:25px; font-weight:800; color:#7CFFB2; }
.score { color:#aaa; font-size:13px; }
mark { background:#ffe066; color:black; padding:2px 4px; border-radius:4px; }
</style>
""", unsafe_allow_html=True)


# ================= SUPABASE CLIENT =================
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_SERVICE_KEY"]
except Exception:
    st.error(
        "Thiếu cấu hình SUPABASE_URL / SUPABASE_SERVICE_KEY trong Secrets.\n\n"
        "Vào Streamlit Cloud > Settings > Secrets và thêm 2 giá trị này."
    )
    st.stop()


@st.cache_resource
def get_supabase_client():
    return create_client(SUPABASE_URL, SUPABASE_KEY)


supabase = get_supabase_client()


# ================= USER / SETTINGS =================
def hash_pw(pw):
    return bcrypt.hashpw(str(pw).encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_pw(pw, hashed):
    try:
        return bcrypt.checkpw(str(pw).encode("utf-8"), str(hashed).encode("utf-8"))
    except Exception:
        return False


def clean_username(username):
    username = str(username).strip().lower()
    return re.sub(r"[^a-zA-Z0-9_-]", "", username)


def default_admin_record():
    return {
        "username": ADMIN_USER,
        "password": hash_pw(ADMIN_PASS_DEFAULT),
        "role": "admin",
        "locked": False,
    }


def ensure_admin_exists():
    # chỉ tạo admin mặc định nếu CHƯA tồn tại, không ghi đè admin đã có
    resp = supabase.table("users").select("username").eq("username", ADMIN_USER).execute()
    if not resp.data:
        supabase.table("users").insert(default_admin_record()).execute()


def get_user_by_username(username):
    # truy vấn đúng 1 user, dùng cho đăng nhập thay vì tải cả bảng
    resp = supabase.table("users").select("*").eq("username", username).execute()
    if not resp.data:
        return None

    row = resp.data[0]
    return {
        "password": row.get("password", ""),
        "role": row.get("role", "user"),
        "locked": row.get("locked", False),
    }


def load_users():
    ensure_admin_exists()

    resp = supabase.table("users").select("*").execute()
    users = {}

    for row in resp.data:
        uname = row["username"]
        users[uname] = {
            "password": row.get("password", ""),
            "role": row.get("role", "user"),
            "locked": row.get("locked", False),
        }

    return users


def save_users(users):
    records = []
    for uname, info in users.items():
        records.append({
            "username": uname,
            "password": info.get("password", ""),
            "role": info.get("role", "user"),
            "locked": info.get("locked", False),
        })

    supabase.table("users").upsert(records).execute()


def default_settings():
    return {
        "allow_register": True,
        "allow_self_delete": True,
        "allow_user_download": True,
        "allow_change_password": True,
        "show_history": True,
        "allow_view_sheet": True
    }


def save_settings(settings):
    supabase.table("settings").upsert({"id": SETTINGS_ROW_ID, "data": settings}).execute()


def load_settings():
    resp = supabase.table("settings").select("*").eq("id", SETTINGS_ROW_ID).execute()

    if resp.data:
        data = resp.data[0].get("data") or {}
    else:
        data = {}

    base = default_settings()
    base.update(data)

    if not resp.data:
        save_settings(base)

    return base


def storage_path(username):
    return f"{username}/data.xlsx"


def _excel_state_key(username):
    return f"__excel_exists_{username}"


@st.cache_data(show_spinner=False)
def download_user_excel_bytes(username):
    # thử tải trực tiếp theo đúng path; nếu file không tồn tại thì Storage sẽ
    # ném exception, ta bắt lại và coi như user chưa có dữ liệu
    #
    # Lưu ý: hàm này cache theo `username` (không theo nội dung file), nên đây
    # là điểm DUY NHẤT có thể trả dữ liệu cũ sau khi upload/xóa — mọi thao tác
    # ghi phải gọi kèm reset_excel_cache() để làm mới đúng cache này.
    try:
        return supabase.storage.from_(BUCKET_NAME).download(storage_path(username))
    except Exception:
        return None


def user_has_excel(username):
    # Ưu tiên trạng thái vừa biết chắc chắn trong phiên hiện tại (vừa upload
    # hoặc vừa xóa xong) thay vì phụ thuộc thời điểm cache_data được làm mới —
    # nhờ vậy kết quả đúng ngay lập tức, không cần F5.
    state_key = _excel_state_key(username)
    if state_key in st.session_state:
        return st.session_state[state_key]

    return download_user_excel_bytes(username) is not None


def reset_excel_cache(username, exists):
    # Xóa đúng cache bị ảnh hưởng (không đụng tới cache của user/dữ liệu khác)
    # và ghi lại trạng thái đã biết vào session_state cho lần kiểm tra tiếp theo
    # trong phiên này.
    download_user_excel_bytes.clear()
    st.session_state[_excel_state_key(username)] = exists


def upload_user_excel_bytes(username, file_bytes):
    # upsert=true => ghi đè file cũ thay vì tạo bản mới
    supabase.storage.from_(BUCKET_NAME).upload(
        storage_path(username),
        file_bytes,
        file_options={
            "upsert": "true",
            "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
    )
    reset_excel_cache(username, True)


def delete_user_excel(username):
    try:
        supabase.storage.from_(BUCKET_NAME).remove([storage_path(username)])
    except Exception:
        pass
    reset_excel_cache(username, False)


def delete_user_completely(username):
    if username == ADMIN_USER:
        return False

    supabase.table("users").delete().eq("username", username).execute()
    supabase.table("search_history").delete().eq("username", username).execute()
    delete_user_excel(username)

    return True


# ================= TEXT =================
def remove_accents(text):
    text = unicodedata.normalize("NFD", str(text))
    return text.encode("ascii", "ignore").decode("utf-8")


def normalize_text(text):
    text = remove_accents(str(text).lower())

    replacements = {
        "xylanh": "xy lanh",
        "xilanh": "xy lanh",
        "xi lanh": "xy lanh",
        "cylinder": "xy lanh",
        "hydraulic": "thuy luc",
        "shaft": "truc",
        "pump": "bom",
        "seal": "phot",
        "steel": "thep",
        "iron": "sat",
        "bearing": "vong bi",
        "gear": "banh rang",
        "chain": "xich",
        "filter": "loc",
        "bolt": "bulong",
        "screw": "vit",
        "harvester": "may gat",
        "tractor": "may cay",
        "rice": "lua",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def format_hs(value):
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"\D", "", s)


def find_name_col(columns):
    keys = ["tên hàng", "ten hang", "mô tả", "mo ta", "description", "name", "tên", "ten"]
    for key in keys:
        k = normalize_text(key)
        for col in columns:
            if k in normalize_text(col):
                return col
    return None


def find_hs_col(columns):
    for col in columns:
        t = normalize_text(col).replace(" ", "")
        if t in ["hs", "mahs", "hscode", "mahscode"]:
            return col
        if "mahs" in t or "hscode" in t:
            return col
    for col in columns:
        if "hs" in str(col).lower():
            return col
    return None


def token_match_positions(query_tokens, text_tokens):
    """
    Tìm vị trí các từ query trong text.
    Ưu tiên khớp chính xác, sau đó mới khớp chứa nhau.
    """
    positions = []
    used_positions = set()

    for q in query_tokens:
        found = None

        # Ưu tiên khớp đúng token
        for i, t in enumerate(text_tokens):
            if i in used_positions:
                continue
            if q == t:
                found = i
                break

        # Nếu không có, mới cho phép q nằm trong token hoặc token nằm trong q
        if found is None:
            for i, t in enumerate(text_tokens):
                if i in used_positions:
                    continue
                if q in t or t in q:
                    found = i
                    break

        if found is not None:
            positions.append(found)
            used_positions.add(found)

    return positions


def ordered_subsequence_count(query_tokens, text_tokens):
    """
    Đếm số từ query xuất hiện theo đúng thứ tự trong text.
    """
    qi = 0
    count = 0

    for t in text_tokens:
        if qi >= len(query_tokens):
            break

        q = query_tokens[qi]

        if q == t or q in t or t in q:
            count += 1
            qi += 1

    return count


def similarity_score(query, text):
    """
    Điểm tìm kiếm theo đúng logic nghiệp vụ:

    1. Đủ toàn bộ từ khóa + đúng cụm liên tiếp => 100
       Ví dụ: query "con chó mực"
       text "abc con chó mực màu vàng chanh xyz" => 100

    2. Đủ toàn bộ từ khóa + đúng thứ tự nhưng có từ chen giữa => 90-99

    3. Đủ toàn bộ từ khóa nhưng sai thứ tự => 75-85

    4. Thiếu từ => không được điểm cao, tối đa 70
    """
    q = normalize_text(query)
    t = normalize_text(text)

    if not q or not t:
        return 0

    q_tokens = q.split()
    t_tokens = t.split()

    if not q_tokens or not t_tokens:
        return 0

    phrase = " ".join(q_tokens)

    # Đúng nguyên cụm, không trừ vì câu dài
    if phrase in t:
        return 100

    positions = token_match_positions(q_tokens, t_tokens)
    matched = len(positions)
    coverage = matched / len(q_tokens)

    # Thiếu từ thì điểm thấp, không cho leo lên đầu
    if coverage < 1:
        return round(coverage * 70, 2)

    # Đủ từ rồi mới xét thứ tự
    ordered_count = ordered_subsequence_count(q_tokens, t_tokens)
    ordered_ratio = ordered_count / len(q_tokens)

    # Đủ từ + đúng thứ tự
    if ordered_ratio == 1:
        gaps = sum(
            max(0, b - a - 1)
            for a, b in zip(positions, positions[1:])
        )

        # Có từ chen giữa thì trừ nhẹ, nhưng vẫn cao hơn sai thứ tự
        return round(max(90, 99 - gaps), 2)

    # Đủ từ nhưng sai thứ tự
    return round(75 + ordered_ratio * 10, 2)


def highlight_text(original, query):
    text = str(original)
    q_tokens = [x for x in normalize_text(query).split() if len(x) >= 2]

    if not q_tokens:
        return html.escape(text)

    def repl(m):
        word = m.group(0)
        norm = normalize_text(word)
        if any(q in norm or norm in q for q in q_tokens):
            return f"<mark>{html.escape(word)}</mark>"
        return html.escape(word)

    return re.sub(r"\S+", repl, text)


# ================= EXCEL =================
@st.cache_data(show_spinner=False)
def _parse_excel_all(file_bytes):
    xf = pd.ExcelFile(BytesIO(file_bytes))
    frames = []

    for sheet in xf.sheet_names:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, dtype=str)
        df = df.dropna(how="all")
        if df.empty:
            continue
        df["Sheet"] = sheet
        frames.append(df)

    if not frames:
        return pd.DataFrame(), xf.sheet_names

    return pd.concat(frames, ignore_index=True), xf.sheet_names


def load_excel_all(username):
    file_bytes = download_user_excel_bytes(username)
    if file_bytes is None:
        raise FileNotFoundError("Không tìm thấy file Excel.")
    return _parse_excel_all(file_bytes)


@st.cache_data(show_spinner=False)
def _parse_one_sheet(file_bytes, sheet_name):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, dtype=str)
    return df.dropna(how="all")


def load_one_sheet(username, sheet_name):
    file_bytes = download_user_excel_bytes(username)
    if file_bytes is None:
        raise FileNotFoundError("Không tìm thấy file Excel.")
    return _parse_one_sheet(file_bytes, sheet_name)


def run_search(query, df):
    name_col = find_name_col(df.columns)
    hs_col = find_hs_col(df.columns)

    if not name_col:
        raise Exception("Không tìm thấy cột tên hàng.")
    if not hs_col:
        raise Exception("Không tìm thấy cột mã HS.")

    result = df.copy()
    result[hs_col] = result[hs_col].apply(format_hs)
    result["Tên tìm kiếm"] = result[name_col].fillna("").astype(str)
    result["Điểm giống"] = result["Tên tìm kiếm"].apply(lambda x: similarity_score(query, x))

    result = result[result["Điểm giống"] >= 25].copy()
    result = result.sort_values("Điểm giống", ascending=False)
    result = result.drop_duplicates(subset=[name_col, hs_col], keep="first")

    return result.head(100), name_col, hs_col


# ================= HISTORY =================
def save_history(username, keyword):
    keyword = str(keyword).strip()
    if not keyword:
        return

    # Xóa bản ghi trùng để đưa từ khóa lên đầu (mới nhất)
    supabase.table("search_history").delete().eq("username", username).eq("keyword", keyword).execute()
    supabase.table("search_history").insert({"username": username, "keyword": keyword}).execute()

    resp = (
        supabase.table("search_history")
        .select("id")
        .eq("username", username)
        .order("created_at", desc=True)
        .execute()
    )
    ids = [r["id"] for r in resp.data]

    if len(ids) > HISTORY_LIMIT:
        stale_ids = ids[HISTORY_LIMIT:]
        supabase.table("search_history").delete().in_("id", stale_ids).execute()


def clear_history(username):
    supabase.table("search_history").delete().eq("username", username).execute()


def delete_history(username, keyword):
    supabase.table("search_history").delete().eq("username", username).eq("keyword", keyword).execute()


def get_history(username):
    resp = (
        supabase.table("search_history")
        .select("keyword")
        .eq("username", username)
        .order("created_at", desc=True)
        .limit(HISTORY_LIMIT)
        .execute()
    )
    return [row["keyword"] for row in resp.data]


# ================= LOGIN =================
def show_login():
    st.title("🔐 Đăng nhập")
    ensure_admin_exists()
    settings = load_settings()
    allow_register = settings.get("allow_register", True)

    tab_login, tab_register = st.tabs(["Đăng nhập", "Đăng ký"])

    with tab_login:
        username = st.text_input("Tên đăng nhập", key="login_user")
        password = st.text_input("Mật khẩu", type="password", key="login_pass")

        if st.button("Đăng nhập", type="primary", use_container_width=True):
            username = clean_username(username)
            user = get_user_by_username(username)

            if user and verify_pw(password, user.get("password", "")):
                if user.get("locked", False):
                    st.error("Tài khoản đã bị khóa. Liên hệ admin.")
                    return

                st.session_state.clear()
                st.session_state["logged_in"] = True
                st.session_state["username"] = username
                st.session_state["role"] = user["role"]
                st.rerun()
            else:
                st.error("Sai tài khoản hoặc mật khẩu.")

    with tab_register:
        if not allow_register:
            st.warning("Đăng ký tài khoản đang tạm khóa. Liên hệ admin để được tạo tài khoản.")
        else:
            users = load_users()
            new_user = st.text_input("Tên đăng nhập mới", key="reg_user")
            new_pass = st.text_input("Mật khẩu", type="password", key="reg_pass")
            new_pass2 = st.text_input("Nhập lại mật khẩu", type="password", key="reg_pass2")

            if st.button("Tạo tài khoản", use_container_width=True):
                new_user = clean_username(new_user)

                if not new_user or not new_pass:
                    st.error("Nhập thiếu.")
                elif new_user == ADMIN_USER:
                    st.error("Tên đăng nhập không hợp lệ.")
                elif new_user in users:
                    st.error("Tài khoản đã tồn tại.")
                elif new_pass != new_pass2:
                    st.error("Mật khẩu không khớp.")
                else:
                    users[new_user] = {
                        "password": hash_pw(new_pass),
                        "role": "user",
                        "locked": False
                    }
                    save_users(users)
                    st.success("Tạo tài khoản thành công. Hãy đăng nhập.")


# ================= ADMIN =================
def show_admin_panel():
    st.title("👑 Admin Panel")
    users = load_users()

    tab_users, tab_files, tab_settings = st.tabs(["👤 Tài khoản", "📁 Dữ liệu user", "⚙️ Cài đặt"])

    with tab_users:
        st.subheader("Quản lý tài khoản")

        for uname, info in users.items():
            c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 2, 3, 2, 1])

            c1.write(f"**{uname}**")
            c2.write(info["role"])

            if uname == ADMIN_USER:
                c3.success("🔒 Hệ thống")
                c4.info("Không thể đổi MK admin")
                c5.info("Luôn hoạt động")
            else:
                status = "🔴 Đã khóa" if info.get("locked", False) else "🟢 Hoạt động"
                c3.write(status)
                c4.write("🔒 Đã mã hóa")

                if info.get("locked", False):
                    if c5.button("🔓 Mở khóa", key=f"unlock_{uname}"):
                        users[uname]["locked"] = False
                        save_users(users)
                        st.success(f"Đã mở khóa {uname}")
                        st.rerun()
                else:
                    if c5.button("🔒 Khóa", key=f"lock_{uname}"):
                        users[uname]["locked"] = True
                        save_users(users)
                        st.success(f"Đã khóa {uname}")
                        st.rerun()

                new_pw = st.text_input(
                    f"MK mới cho {uname}",
                    type="password",
                    key=f"reset_pw_{uname}",
                    placeholder="Nhập MK mới",
                    label_visibility="collapsed"
                )

                if st.button("Reset MK", key=f"reset_btn_{uname}"):
                    if not new_pw:
                        st.warning("Nhập mật khẩu mới trước.")
                    else:
                        users[uname]["password"] = hash_pw(new_pw)
                        save_users(users)
                        st.success(f"Đã reset mật khẩu cho {uname}")
                        st.rerun()

                confirm_key = f"confirm_delete_{uname}"

                if c6.button("🗑", key=f"delete_user_{uname}"):
                    st.session_state[confirm_key] = True

                if st.session_state.get(confirm_key):
                    st.warning(
                        f"Bạn chắc chắn muốn xóa tài khoản **{uname}** chứ?\n\n"
                        "Toàn bộ file dữ liệu, lịch sử tìm kiếm và thông tin tài khoản sẽ mất vĩnh viễn."
                    )

                    col_yes, col_no = st.columns(2)

                    if col_yes.button("✅ Xác nhận xóa", key=f"yes_delete_{uname}"):
                        delete_user_completely(uname)
                        st.success(f"Đã xóa {uname}")
                        st.rerun()

                    if col_no.button("❌ Hủy", key=f"no_delete_{uname}"):
                        st.session_state[confirm_key] = False
                        st.rerun()

            st.markdown("---")

    with tab_files:
        st.subheader("Xem / tải / upload dữ liệu từng user")

        user_list = [u for u in users.keys() if u != ADMIN_USER]

        if not user_list:
            st.warning("Chưa có tài khoản user nào.")
        else:
            target_user = st.selectbox("Chọn user", user_list)
            has_data = user_has_excel(target_user)

            if has_data:
                st.success("User này đã có file dữ liệu.")

                file_bytes = download_user_excel_bytes(target_user)

                if file_bytes:
                    st.download_button(
                        "⬇️ Tải file Excel của user này",
                        data=file_bytes,
                        file_name=f"{target_user}_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                try:
                    df, sheets = load_excel_all(target_user)
                    st.write(f"Đọc được **{len(df)} dòng** / **{len(sheets)} sheet**")
                    st.dataframe(df.head(100), use_container_width=True)
                except Exception as e:
                    st.error(f"Lỗi đọc file: {e}")

                if st.button("🗑 Xóa file dữ liệu user này"):
                    delete_user_excel(target_user)
                    st.rerun()
            else:
                st.warning("User này chưa có file dữ liệu.")

            uploaded = st.file_uploader(
                f"Upload Excel cho {target_user}",
                type=["xlsx"],
                key=f"admin_upload_{target_user}"
            )

            if uploaded:
                upload_user_excel_bytes(target_user, uploaded.getvalue())
                st.success("Đã upload file.")
                st.rerun()

    with tab_settings:
        st.subheader("Cài đặt hệ thống")

        settings = load_settings()

        options = [
            ("allow_register", "Cho phép người dùng tự đăng ký tài khoản"),
            ("allow_self_delete", "Cho phép user tự xóa tài khoản"),
            ("allow_user_download", "Cho phép user tải file dữ liệu của mình"),
            ("allow_change_password", "Cho phép user đổi mật khẩu"),
            ("show_history", "Hiển thị lịch sử tìm kiếm"),
            ("allow_view_sheet", "Cho phép user xem sheet dữ liệu"),
        ]

        changed = False

        for key, label in options:
            old_val = settings.get(key, True)
            new_val = st.toggle(label, value=old_val, key=f"setting_{key}")
            if new_val != old_val:
                settings[key] = new_val
                changed = True

        if changed:
            save_settings(settings)
            st.success("Đã cập nhật cài đặt hệ thống.")
            st.rerun()


# ================= SEARCH =================
def show_search_page(username):
    st.title("🔎 Tra cứu mã HS")
    settings = load_settings()

    search_key = f"search_text_{username}"

    if not user_has_excel(username):
        st.info("Chưa có file dữ liệu. Upload Excel ở sidebar trước.")
        return

    try:
        df, sheets = load_excel_all(username)
    except Exception as e:
        st.error(f"Lỗi đọc Excel: {e}")
        return

    if df.empty:
        st.error("File Excel không có dữ liệu.")
        return

    name_col = find_name_col(df.columns)
    hs_col = find_hs_col(df.columns)

    st.caption(f"Đọc được **{len(df)} dòng** / **{len(sheets)} sheet**")

    if not name_col:
        st.error("Không tìm thấy cột tên hàng.")
        st.write("Các cột đang có:", list(df.columns))
        return

    if not hs_col:
        st.error("Không tìm thấy cột mã HS.")
        st.write("Các cột đang có:", list(df.columns))
        return

    st.caption(f"Cột tên hàng: **{name_col}** | Cột mã HS: **{hs_col}**")

    if search_key not in st.session_state:
        st.session_state[search_key] = ""

    pending_key = f"pending_search_text_{username}"

    # Streamlit không cho sửa st.session_state[search_key]
    # sau khi text_input đã render. Vì vậy click lịch sử sẽ ghi vào pending_key,
    # rồi trước khi tạo text_input mới chuyển pending_key sang search_key.
    if pending_key in st.session_state:
        st.session_state[search_key] = st.session_state.pop(pending_key)

    query = st.text_input(
        "Nhập tên hàng cần tra",
        key=search_key,
        placeholder="Ví dụ: phớt sắt, xy lanh, hydraulic cylinder..."
    )

    if query:
        save_history(username, query)

        try:
            result, name_col, hs_col = run_search(query, df)
        except Exception as e:
            st.error(f"Lỗi tìm kiếm: {e}")
            return

        if result.empty:
            st.warning("Không tìm thấy kết quả phù hợp.")
            return

        hs_stats = result[hs_col].value_counts().reset_index()
        hs_stats.columns = ["Mã HS", "Số lần xuất hiện"]

        best_hs = hs_stats.iloc[0]["Mã HS"]
        best_count = hs_stats.iloc[0]["Số lần xuất hiện"]

        st.success(f"💡 Gợi ý mã đáng tin nhất: **{best_hs}** — xuất hiện {best_count} lần")

        with st.expander("📊 Thống kê mã HS"):
            st.dataframe(hs_stats, use_container_width=True, hide_index=True)

        st.subheader(f"Tìm thấy {len(result)} kết quả")

        for _, row in result.iterrows():
            hs = html.escape(str(row[hs_col]))
            name = highlight_text(row[name_col], query)
            score = round(float(row["Điểm giống"]), 1)
            sheet = html.escape(str(row.get("Sheet", "")))

            st.markdown(f"""
            <div class="card">
                <div class="hs">Mã HS: {hs}</div>
                <p><b>Tên hàng:</b> {name}</p>
                <p class="score">Độ giống: {score}% | Sheet: {sheet}</p>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("Xem toàn bộ dòng"):
                show_row = row.drop(labels=["Tên tìm kiếm", "Điểm giống"], errors="ignore")
                st.write(show_row.to_frame("Giá trị"))

    else:
        if settings.get("show_history", True):
            history_list = get_history(username)

            if history_list:
                with st.expander("🕘 Lịch sử tìm kiếm", expanded=True):
                    if st.button("🧹 Xóa toàn bộ lịch sử", key=f"clear_history_{username}"):
                        clear_history(username)
                        st.rerun()

                    for i, kw in enumerate(history_list):
                        c1, c2, c3 = st.columns([5, 1, 1])
                        c1.write(kw)

                        if c2.button("🔍", key=f"hist_search_{username}_{i}_{kw}"):
                            st.session_state[f"pending_search_text_{username}"] = kw
                            st.rerun()

                        if c3.button("🗑", key=f"hist_delete_{username}_{i}_{kw}"):
                            delete_history(username, kw)
                            st.rerun()


# ================= SHEET =================
def show_sheet_page(username):
    st.title("📋 Xem dữ liệu sheet")

    if not user_has_excel(username):
        st.info("Chưa có file dữ liệu.")
        return

    try:
        _, sheets = load_excel_all(username)
    except Exception as e:
        st.error(f"Lỗi đọc Excel: {e}")
        return

    if not sheets:
        st.error("File không có sheet.")
        return

    selected_sheet = st.selectbox("Chọn sheet", sheets, key=f"sheet_select_{username}")

    try:
        df = load_one_sheet(username, selected_sheet)
    except Exception as e:
        st.error(f"Lỗi đọc sheet: {e}")
        return

    if df.empty:
        st.warning("Sheet này rỗng.")
        return

    hs_col = find_hs_col(df.columns)
    if hs_col:
        df[hs_col] = df[hs_col].apply(format_hs)

    st.caption(f"Sheet **{selected_sheet}** — {len(df)} dòng")

    detail_key = f"opened_row_{username}_{selected_sheet}"

    if detail_key not in st.session_state:
        st.session_state[detail_key] = None

    row_num = st.number_input(
        "Mở chi tiết dòng",
        min_value=1,
        max_value=len(df),
        value=1,
        step=1,
        key=f"row_num_{username}_{selected_sheet}"
    )

    if st.button("🃏 Mở thẻ dòng này", key=f"open_card_{username}_{selected_sheet}"):
        st.session_state[detail_key] = int(row_num) - 1
        st.rerun()

    if st.session_state[detail_key] is not None:
        idx = st.session_state[detail_key]
        row = df.iloc[idx]

        name_col = find_name_col(df.columns)
        hs_col = find_hs_col(df.columns)

        name = str(row[name_col]) if name_col else "—"
        hs = str(row[hs_col]) if hs_col else "—"

        c1, c2, c3 = st.columns([1, 1, 4])

        if c1.button("❌ Đóng thẻ", key=f"close_card_{username}_{selected_sheet}"):
            st.session_state[detail_key] = None
            st.rerun()

        if idx > 0:
            if c2.button("⬅ Dòng trước", key=f"prev_row_{username}_{selected_sheet}"):
                st.session_state[detail_key] = idx - 1
                st.rerun()

        if idx < len(df) - 1:
            if c3.button("Dòng tiếp ➡", key=f"next_row_{username}_{selected_sheet}"):
                st.session_state[detail_key] = idx + 1
                st.rerun()

        st.markdown(f"""
        <div class="card">
            <div class="hs">Mã HS: {html.escape(hs)}</div>
            <p><b>Tên hàng:</b><br>{html.escape(name)}</p>
            <p class="score">Sheet: {html.escape(selected_sheet)} | Dòng: {idx + 1}</p>
        </div>
        """, unsafe_allow_html=True)

        st.write(row.to_frame("Giá trị"))

    table = df.copy()
    table.index = table.index + 1
    table.index.name = "Dòng"
    st.dataframe(table, use_container_width=True)


# ================= MAIN =================
def show_main_app():
    username = st.session_state["username"]
    role = st.session_state["role"]
    settings = load_settings()

    with st.sidebar:
        st.markdown(f"### 👤 {username}")

        if role == "admin":
            st.markdown("👑 Admin")
            st.info("Admin chỉ dùng để quản trị.")
        else:
            st.markdown("---")
            st.markdown("### 📁 File dữ liệu")

            has_data = user_has_excel(username)

            if has_data:
                st.success("✅ Đã có file dữ liệu")
            else:
                st.warning("Chưa có file dữ liệu")

            uploaded = st.file_uploader("Upload file Excel", type=["xlsx"], key=f"upload_{username}")

            if uploaded:
                upload_user_excel_bytes(username, uploaded.getvalue())
                st.success("Đã lưu file Excel.")
                st.rerun()

            if has_data and settings.get("allow_user_download", True):
                file_bytes = download_user_excel_bytes(username)
                if file_bytes:
                    st.download_button(
                        "⬇️ Tải file hiện tại",
                        data=file_bytes,
                        file_name=f"{username}_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{username}"
                    )

            st.markdown("---")

            if settings.get("allow_change_password", True):
                with st.expander("🔑 Đổi mật khẩu"):
                    old_pw = st.text_input("Mật khẩu cũ", type="password", key=f"old_pw_{username}")
                    new_pw = st.text_input("Mật khẩu mới", type="password", key=f"new_pw_{username}")
                    new_pw2 = st.text_input("Nhập lại mật khẩu mới", type="password", key=f"new_pw2_{username}")

                    if st.button("Lưu mật khẩu", key=f"save_pw_{username}"):
                        users = load_users()

                        ok_old = verify_pw(old_pw, users[username].get("password", ""))

                        if not ok_old:
                            st.error("Mật khẩu cũ sai.")
                        elif not new_pw:
                            st.error("Mật khẩu mới không được trống.")
                        elif new_pw != new_pw2:
                            st.error("Mật khẩu mới không khớp.")
                        else:
                            users[username]["password"] = hash_pw(new_pw)
                            save_users(users)
                            st.success("Đã đổi mật khẩu.")

            if settings.get("allow_self_delete", True):
                with st.expander("⚠️ Xóa tài khoản"):
                    st.warning("Xóa tài khoản sẽ xóa toàn bộ file Excel và lịch sử tìm kiếm của bạn.")

                    if st.button("🗑 Xóa tài khoản của tôi", key=f"self_delete_{username}"):
                        st.session_state[f"confirm_self_delete_{username}"] = True

                    if st.session_state.get(f"confirm_self_delete_{username}"):
                        st.error("Bạn chắc chắn muốn xóa tài khoản này chứ? Dữ liệu sẽ mất vĩnh viễn.")

                        c1, c2 = st.columns(2)

                        if c1.button("✅ Xác nhận xóa", key=f"yes_self_delete_{username}"):
                            delete_user_completely(username)
                            st.session_state.clear()
                            st.rerun()

                        if c2.button("❌ Hủy", key=f"no_self_delete_{username}"):
                            st.session_state[f"confirm_self_delete_{username}"] = False
                            st.rerun()

        st.markdown("---")

        if st.button("🚪 Đăng xuất", use_container_width=True, key=f"logout_{username}"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    if role == "admin":
        show_admin_panel()
    else:
        if settings.get("allow_view_sheet", True):
            tab_search, tab_sheet = st.tabs(["🔎 Tra cứu", "📋 Xem sheet"])

            with tab_search:
                show_search_page(username)

            with tab_sheet:
                show_sheet_page(username)
        else:
            show_search_page(username)


# ================= ROUTER =================
if not st.session_state.get("logged_in"):
    show_login()
else:
    show_main_app()