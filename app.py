import streamlit as st
import pandas as pd
import os
import re
import html
import json
import hashlib
import unicodedata
from pathlib import Path
from difflib import SequenceMatcher

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Tra cứu mã HS", layout="wide")

BASE_DIR = Path(__file__).parent
USERDATA_DIR = BASE_DIR / "userdata"
USERDATA_DIR.mkdir(exist_ok=True)

USERS_FILE = BASE_DIR / "users.json"
ADMIN_USER = "admin"
ADMIN_PASS_DEFAULT = "000"

st.markdown("""
<style>
.stApp { background: #0f1117; color: #eee; }
div[data-testid="stTextInput"] input { font-size: 16px; padding: 10px; }
.card {
    border: 1px solid #333846;
    border-radius: 16px;
    padding: 16px;
    margin-bottom: 12px;
    background: #171a22;
}
.hs { font-size: 25px; font-weight: 800; color: #7CFFB2; }
.score { color: #aaa; font-size: 13px; }
mark { background: #ffe066; color: black; padding: 2px 4px; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)


# =========================
# AUTH
# =========================
def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()


def load_users() -> dict:
    if USERS_FILE.exists():
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    users = {
        ADMIN_USER: {
            "password": hash_pw(ADMIN_PASS_DEFAULT),
            "role": "admin"
        }
    }
    save_users(users)
    return users


def save_users(users: dict):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def clean_username(username: str) -> str:
    username = username.strip().lower()
    username = re.sub(r"[^a-zA-Z0-9_-]", "", username)
    return username


def user_dir(username: str) -> Path:
    d = USERDATA_DIR / username
    d.mkdir(parents=True, exist_ok=True)
    return d


def data_file(username: str) -> Path:
    return user_dir(username) / "data.xlsx"


def history_file(username: str) -> Path:
    return user_dir(username) / "history.csv"


# =========================
# TEXT SEARCH
# =========================
def remove_accents(text):
    text = unicodedata.normalize("NFD", str(text))
    text = text.encode("ascii", "ignore").decode("utf-8")
    return text


def normalize_text(text):
    text = remove_accents(str(text).lower())

    replacements = {
        "xylanh": "xy lanh",
        "xi lanh": "xy lanh",
        "xilanh": "xy lanh",
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
    text = re.sub(r"\s+", " ", text).strip()
    return text


def format_hs(value):
    if pd.isna(value):
        return ""

    s = str(value).strip()

    if re.match(r"^\d+\.0$", s):
        s = s[:-2]

    s = re.sub(r"\D", "", s)

    return s


def find_name_col(columns):
    priority = [
        "tên hàng", "ten hang", "mô tả", "mo ta",
        "description", "name", "tên", "ten"
    ]

    norm_cols = {col: normalize_text(col) for col in columns}

    for key in priority:
        key_norm = normalize_text(key)
        for col, norm in norm_cols.items():
            if key_norm in norm:
                return col

    return None


def find_hs_col(columns):
    for col in columns:
        t = normalize_text(col).replace(" ", "")
        if t in ["hs", "mahs", "mahscode", "hscode"]:
            return col
        if "mahs" in t or "hscode" in t or "hs" == t:
            return col

    for col in columns:
        if "hs" in str(col).lower():
            return col

    return None


def similarity_score(query, text):
    q = normalize_text(query)
    t = normalize_text(text)

    if not q or not t:
        return 0

    if q == t:
        return 100

    if q in t:
        return 95

    q_tokens = q.split()
    t_tokens = t.split()

    matched = 0
    for qt in q_tokens:
        if any(qt == tt or qt in tt or tt in qt for tt in t_tokens):
            matched += 1

    token_score = matched / len(q_tokens) * 80 if q_tokens else 0
    seq_score = SequenceMatcher(None, q, t).ratio() * 100

    return round(max(token_score, seq_score), 2)


def highlight_text(original_text, query):
    safe = html.escape(str(original_text))
    q_tokens = [w for w in normalize_text(query).split() if len(w) >= 2]

    if not q_tokens:
        return safe

    def repl(m):
        word = m.group(0)
        norm = normalize_text(word)
        if any(q in norm or norm in q for q in q_tokens):
            return f"<mark>{html.escape(word)}</mark>"
        return html.escape(word)

    return re.sub(r"\S+", repl, str(original_text))


# =========================
# EXCEL
# =========================
@st.cache_data(show_spinner=False)
def load_excel_all(path_str: str):
    path = Path(path_str)

    if not path.exists():
        raise FileNotFoundError("Không thấy file Excel.")

    xf = pd.ExcelFile(path)
    frames = []

    for sheet in xf.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
        df = df.dropna(how="all")

        if len(df) == 0:
            continue

        df["Sheet"] = sheet
        frames.append(df)

    if not frames:
        return pd.DataFrame(), xf.sheet_names

    full_df = pd.concat(frames, ignore_index=True)
    return full_df, xf.sheet_names


@st.cache_data(show_spinner=False)
def load_one_sheet(path_str: str, sheet_name: str):
    df = pd.read_excel(path_str, sheet_name=sheet_name, dtype=str)
    return df.dropna(how="all")


def run_search(query, df):
    if df.empty:
        return pd.DataFrame()

    name_col = find_name_col(df.columns)
    hs_col = find_hs_col(df.columns)

    if not name_col:
        raise ValueError("Không tìm thấy cột tên hàng / mô tả hàng.")

    if not hs_col:
        raise ValueError("Không tìm thấy cột mã HS.")

    result = df.copy()
    result[hs_col] = result[hs_col].apply(format_hs)

    result["Tên tìm kiếm"] = result[name_col].fillna("").astype(str)
    result["Điểm giống"] = result["Tên tìm kiếm"].apply(
        lambda x: similarity_score(query, x)
    )

    result = result[result["Điểm giống"] > 25].copy()
    result = result.sort_values("Điểm giống", ascending=False)
    result = result.drop_duplicates(subset=[name_col, hs_col], keep="first")

    return result.head(100)


# =========================
# HISTORY
# =========================
def save_history(username, keyword):
    keyword = keyword.strip()
    if not keyword:
        return

    hfile = history_file(username)
    new = pd.DataFrame([{"Từ khóa": keyword}])

    if hfile.exists():
        old = pd.read_csv(hfile)
        data = pd.concat([new, old], ignore_index=True)
    else:
        data = new

    data = data.drop_duplicates(subset=["Từ khóa"], keep="first").head(30)
    data.to_csv(hfile, index=False)


def clear_history(username):
    hfile = history_file(username)
    if hfile.exists():
        hfile.unlink()


def delete_history(username, keyword):
    hfile = history_file(username)
    if not hfile.exists():
        return

    df = pd.read_csv(hfile)
    df = df[df["Từ khóa"] != keyword]
    df.to_csv(hfile, index=False)


# =========================
# UI LOGIN
# =========================
def show_login():
    st.title("🔐 Đăng nhập")

    users = load_users()

    tab1, tab2 = st.tabs(["Đăng nhập", "Đăng ký"])

    with tab1:
        username = st.text_input("Tên đăng nhập")
        password = st.text_input("Mật khẩu", type="password")

        if st.button("Đăng nhập", type="primary", use_container_width=True):
            username = clean_username(username)

            if username in users and users[username]["password"] == hash_pw(password):
                st.session_state["logged_in"] = True
                st.session_state["username"] = username
                st.session_state["role"] = users[username]["role"]
                st.rerun()
            else:
                st.error("Sai tên đăng nhập hoặc mật khẩu.")

    with tab2:
        new_user = st.text_input("Tên đăng nhập mới")
        new_pass = st.text_input("Mật khẩu mới", type="password")
        new_pass2 = st.text_input("Nhập lại mật khẩu", type="password")

        if st.button("Tạo tài khoản", use_container_width=True):
            new_user = clean_username(new_user)

            if not new_user or not new_pass:
                st.error("Nhập thiếu.")
            elif new_user == ADMIN_USER:
                st.error("Tên này không dùng được.")
            elif new_user in users:
                st.error("Tài khoản đã tồn tại.")
            elif new_pass != new_pass2:
                st.error("Mật khẩu không khớp.")
            else:
                users[new_user] = {
                    "password": hash_pw(new_pass),
                    "role": "user"
                }
                save_users(users)
                user_dir(new_user)
                st.success("Tạo tài khoản thành công. Giờ đăng nhập được rồi.")


# =========================
# ADMIN
# =========================
def show_admin_panel():
    st.title("👑 Admin Panel")

    users = load_users()

    tab_users, tab_files = st.tabs(["👤 Tài khoản", "📁 File dữ liệu"])

    with tab_users:
        st.subheader("Quản lý tài khoản")

        for uname, info in users.items():
            c1, c2, c3, c4 = st.columns([3, 2, 2, 2])

            c1.write(f"**{uname}**")
            c2.write(info["role"])

            with c3:
                new_pw = st.text_input(
                    "MK mới",
                    key=f"newpw_{uname}",
                    type="password",
                    label_visibility="collapsed",
                    placeholder="MK mới"
                )

                if st.button("Reset MK", key=f"reset_{uname}"):
                    if new_pw:
                        users[uname]["password"] = hash_pw(new_pw)
                        save_users(users)
                        st.success(f"Đã reset mật khẩu cho {uname}")
                        st.rerun()
                    else:
                        st.warning("Nhập mật khẩu mới trước.")

            with c4:
                if uname != ADMIN_USER:
                    if st.button("Xóa", key=f"del_{uname}"):
                        del users[uname]
                        save_users(users)
                        st.success(f"Đã xóa {uname}")
                        st.rerun()

        st.info("Không hiển thị mật khẩu cũ vì mật khẩu đang lưu dạng mã hóa hash. Admin chỉ nên reset, không nên xem.")

    with tab_files:
        st.subheader("Dữ liệu tất cả người dùng")

        target_user = st.selectbox("Chọn user", list(users.keys()))
        path = data_file(target_user)

        st.write(f"Đường dẫn file: `{path}`")

        if path.exists():
            st.success("User này đã có file dữ liệu.")

            with open(path, "rb") as f:
                st.download_button(
                    "⬇️ Tải file Excel của user này",
                    data=f,
                    file_name=f"{target_user}_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if st.button("🗑 Xóa file dữ liệu user này"):
                path.unlink()
                st.cache_data.clear()
                st.rerun()

            try:
                df, sheets = load_excel_all(str(path))
                st.write(f"Đọc được **{len(df)} dòng**, **{len(sheets)} sheet**.")
                st.dataframe(df.head(50), use_container_width=True)
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}")
        else:
            st.warning("User này chưa có file dữ liệu.")

        uploaded = st.file_uploader(
            f"Upload file Excel cho {target_user}",
            type=["xlsx"],
            key=f"admin_upload_{target_user}"
        )

        if uploaded:
            with open(path, "wb") as f:
                f.write(uploaded.getbuffer())

            st.cache_data.clear()
            st.success("Đã upload file cho user.")
            st.rerun()


# =========================
# SEARCH UI
# =========================
def show_search_page(username):
    st.title("🔎 Tra cứu mã HS")

    path = data_file(username)

    if not path.exists():
        st.info("Chưa có file dữ liệu. Upload file Excel ở sidebar trước.")
        return

    try:
        df, sheets = load_excel_all(str(path))
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
        st.error("Không tìm thấy cột tên hàng. Cột nên có chữ: Tên hàng / Mô tả / Description.")
        st.write("Các cột đang có:", list(df.columns))
        return

    if not hs_col:
        st.error("Không tìm thấy cột mã HS. Cột nên có chữ: HS / Mã HS / HS Code.")
        st.write("Các cột đang có:", list(df.columns))
        return

    st.caption(f"Cột tên hàng: **{name_col}** | Cột mã HS: **{hs_col}**")

    if "search_text" not in st.session_state:
        st.session_state["search_text"] = ""

    query = st.text_input(
        "Nhập tên hàng cần tra",
        key="search_text",
        placeholder="Ví dụ: phớt sắt, xy lanh thủy lực, hydraulic cylinder..."
    )

    col_a, col_b = st.columns([1, 5])
    with col_a:
        do_search = st.button("🔍 Tìm", type="primary", use_container_width=True)

    if query and (do_search or query):
        save_history(username, query)

        try:
            result = run_search(query, df)
        except Exception as e:
            st.error(f"Lỗi tìm kiếm: {e}")
            return

        if result.empty:
            st.warning("Không tìm thấy kết quả phù hợp.")
            return

        hs_stats = result[hs_col].value_counts().reset_index()
        hs_stats.columns = ["Mã HS", "Số lần"]

        best_hs = hs_stats.iloc[0]["Mã HS"]
        best_count = hs_stats.iloc[0]["Số lần"]

        st.success(f"💡 Gợi ý mã đáng tin nhất: **{best_hs}** — xuất hiện {best_count} lần")

        with st.expander("📊 Thống kê mã HS"):
            st.dataframe(hs_stats, use_container_width=True, hide_index=True)

        st.subheader(f"Tìm thấy {len(result)} kết quả")

        for i, row in result.iterrows():
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
                clean_row = row.drop(labels=["Tên tìm kiếm", "Điểm giống"], errors="ignore")
                st.write(clean_row.to_frame("Giá trị"))

    else:
        hfile = history_file(username)
        if hfile.exists():
            hist = pd.read_csv(hfile)
            with st.expander("🕘 Lịch sử tìm kiếm", expanded=True):
                if st.button("🧹 Xóa toàn bộ lịch sử"):
                    clear_history(username)
                    st.rerun()

                for _, r in hist.iterrows():
                    kw = r["Từ khóa"]
                    c1, c2, c3 = st.columns([5, 1, 1])
                    c1.write(kw)

                    if c2.button("🔍", key=f"search_hist_{kw}"):
                        st.session_state["search_text"] = kw
                        st.rerun()

                    if c3.button("🗑", key=f"del_hist_{kw}"):
                        delete_history(username, kw)
                        st.rerun()


# =========================
# SHEET VIEW
# =========================
def show_sheet_page(username):
    st.title("📋 Xem dữ liệu sheet")

    path = data_file(username)

    if not path.exists():
        st.info("Chưa có file dữ liệu.")
        return

    try:
        _, sheets = load_excel_all(str(path))
    except Exception as e:
        st.error(f"Lỗi đọc Excel: {e}")
        return

    if not sheets:
        st.error("File không có sheet.")
        return

    selected = st.selectbox("Chọn sheet", sheets)

    try:
        df = load_one_sheet(str(path), selected)
    except Exception as e:
        st.error(f"Lỗi đọc sheet: {e}")
        return

    if df.empty:
        st.warning("Sheet này rỗng.")
        return

    hs_col = find_hs_col(df.columns)
    if hs_col:
        df[hs_col] = df[hs_col].apply(format_hs)

    st.caption(f"Sheet **{selected}** — {len(df)} dòng")

    row_num = st.number_input(
        "Mở chi tiết dòng",
        min_value=1,
        max_value=max(1, len(df)),
        value=1,
        step=1
    )

    if st.button("🃏 Mở thẻ dòng này"):
        row = df.iloc[int(row_num) - 1]

        name_col = find_name_col(df.columns)
        hs_col = find_hs_col(df.columns)

        name = str(row[name_col]) if name_col else "—"
        hs = str(row[hs_col]) if hs_col else "—"

        st.markdown(f"""
        <div class="card">
            <div class="hs">Mã HS: {html.escape(hs)}</div>
            <p><b>Tên hàng:</b><br>{html.escape(name)}</p>
        </div>
        """, unsafe_allow_html=True)

        st.write(row.to_frame("Giá trị"))

    table = df.copy()
    table.index = table.index + 1
    table.index.name = "Dòng"
    st.dataframe(table, use_container_width=True)


# =========================
# MAIN APP
# =========================
def show_main_app():
    username = st.session_state["username"]
    role = st.session_state["role"]

    path = data_file(username)

    with st.sidebar:
        st.markdown(f"### 👤 {username}")

        if role == "admin":
            st.markdown("👑 Admin")

        st.markdown("---")
        st.markdown("### 📁 File dữ liệu")

        if path.exists():
            st.success("✅ Đã có file dữ liệu")
        else:
            st.warning("Chưa có file dữ liệu")

        uploaded = st.file_uploader("Upload file Excel", type=["xlsx"])

        if uploaded:
            with open(path, "wb") as f:
                f.write(uploaded.getbuffer())

            st.cache_data.clear()
            st.success("Đã lưu file Excel.")
            st.rerun()

        if path.exists():
            with open(path, "rb") as f:
                st.download_button(
                    "⬇️ Tải file hiện tại",
                    data=f,
                    file_name=f"{username}_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        st.markdown("---")

        with st.expander("🔑 Đổi mật khẩu"):
            old_pw = st.text_input("Mật khẩu cũ", type="password")
            new_pw = st.text_input("Mật khẩu mới", type="password")
            new_pw2 = st.text_input("Nhập lại mật khẩu mới", type="password")

            if st.button("Lưu mật khẩu"):
                users = load_users()

                if users[username]["password"] != hash_pw(old_pw):
                    st.error("Mật khẩu cũ sai.")
                elif not new_pw:
                    st.error("Mật khẩu mới không được trống.")
                elif new_pw != new_pw2:
                    st.error("Mật khẩu mới không khớp.")
                else:
                    users[username]["password"] = hash_pw(new_pw)
                    save_users(users)
                    st.success("Đã đổi mật khẩu.")

        if st.button("🚪 Đăng xuất", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    if role == "admin":
        tab1, tab2, tab3 = st.tabs(["🔎 Tra cứu", "📋 Xem sheet", "👑 Admin Panel"])
    else:
        tab1, tab2 = st.tabs(["🔎 Tra cứu", "📋 Xem sheet"])
        tab3 = None

    with tab1:
        show_search_page(username)

    with tab2:
        show_sheet_page(username)

    if role == "admin" and tab3:
        with tab3:
            show_admin_panel()


# =========================
# ROUTER
# =========================
if not st.session_state.get("logged_in"):
    show_login()
else:
    show_main_app()