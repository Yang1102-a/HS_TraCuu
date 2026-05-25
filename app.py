import streamlit as st
import pandas as pd
import os
import re
import html
import unicodedata

st.set_page_config(page_title="Tra cứu mã HS", layout="wide")

st.markdown("""
<style>
.stApp { background: #0f1117; }
div[data-testid="stTextInput"] input { font-size: 18px; padding: 14px; }
.card {
    border: 1px solid #333846;
    border-radius: 16px;
    padding: 18px;
    margin-bottom: 14px;
    background: #171a22;
}
.hs {
    font-size: 26px;
    font-weight: 800;
    color: #7CFFB2;
}
.score { color: #aaa; }
mark {
    background: #ffe066;
    color: black;
    padding: 2px 4px;
    border-radius: 4px;
}
</style>
""", unsafe_allow_html=True)

st.title("🔎 Tra cứu mã HS")

DATA_FILE = "data.xlsx"
HISTORY_FILE = "search_history.csv"


uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

if uploaded_file:
    with open(DATA_FILE, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("Đã lưu file Excel thành công!")
    st.cache_data.clear()


# ── Helpers ──────────────────────────────────────────────
def remove_accents(text):
    text = unicodedata.normalize("NFD", str(text))
    return text.encode("ascii", "ignore").decode("utf-8")


def normalize_text(text):
    text = remove_accents(str(text).lower())

    replacements = {
        "xylanh": "xy lanh",
        "xi lanh": "xy lanh",
        "cylinder": "xy lanh",
        "hydraulic": "thuy luc",
        "shaft": "truc",
        "pump": "bom",
        "seal": "phot",
        "steel": "thep",
        "iron": "sat",
        "bolt": "bulong",
        "screw": "vit",
        "bearing": "vong bi",
        "gear": "banh rang",
        "chain": "xich",
        "filter": "loc",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    return text


def find_name_col(columns):
    for col in columns:
        if "tên hàng" in str(col).lower():
            return col

    for col in columns:
        t = str(col).lower()
        if "tên" in t or "hàng" in t or "hang" in t:
            return col

    return None


def find_hs_col(columns):
    for col in columns:
        t = str(col).lower().replace(" ", "")
        if "mãhs" in t or "mahs" in t or t == "hs" or "hs" in t:
            return col

    return None


def highlight_text(original_text, query):
    safe_text = html.escape(str(original_text))

    query_tokens = set(
        w for w in normalize_text(query).split()
        if len(w) >= 2 and not w.isdigit()
    )

    if not query_tokens:
        return safe_text

    def replace_token(m):
        token = m.group(0)
        norm = normalize_text(token)

        if not norm or len(norm) < 2:
            return token

        if norm in query_tokens:
            return f"<mark>{token}</mark>"

        if any(qt in norm for qt in query_tokens if len(qt) >= 4):
            return f"<mark>{token}</mark>"

        return token

    return re.sub(r"\S+", replace_token, safe_text)


def save_history(keyword):
    if st.session_state.get("last_saved_keyword") == keyword:
        return

    st.session_state["last_saved_keyword"] = keyword

    new_row = pd.DataFrame([{"Từ khóa": keyword}])

    if os.path.exists(HISTORY_FILE):
        old = pd.read_csv(HISTORY_FILE)
        history = pd.concat([new_row, old], ignore_index=True)
    else:
        history = new_row

    history.drop_duplicates().head(20).to_csv(HISTORY_FILE, index=False)


@st.cache_data
def load_excel(path):
    xf = pd.ExcelFile(path)
    frames = []

    for sheet in xf.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        df = df.dropna(how="all")
        df["Sheet"] = sheet
        frames.append(df)

    return pd.concat(frames, ignore_index=True), xf.sheet_names


@st.cache_data
def load_sheet(path, sheet_name):
    return pd.read_excel(path, sheet_name=sheet_name).dropna(how="all")


# ── Search scoring mới ──────────────────────────────────
def token_match_positions(query_tokens, text_tokens):
    positions = []
    used_positions = set()

    for q in query_tokens:
        found = None

        for i, t in enumerate(text_tokens):
            if i in used_positions:
                continue

            if q == t or q in t:
                found = i
                break

        if found is not None:
            positions.append(found)
            used_positions.add(found)

    return positions


def longest_ordered_subsequence_score(query_tokens, text_tokens):
    """
    Đo mức đúng thứ tự của các từ query trong text.
    Đúng thứ tự hoàn toàn => 1.0
    Sai thứ tự nhiều => thấp hơn
    """
    q_index = 0
    matched = 0

    for t in text_tokens:
        if q_index >= len(query_tokens):
            break

        q = query_tokens[q_index]

        if q == t or q in t:
            matched += 1
            q_index += 1

    return matched / len(query_tokens) if query_tokens else 0


def ordered_match_score(query_norm, text_norm):
    q_tokens = query_norm.split()
    t_tokens = text_norm.split()

    if not q_tokens or not t_tokens:
        return 0

    positions = token_match_positions(q_tokens, t_tokens)
    matched_count = len(positions)
    coverage = matched_count / len(q_tokens)

    # Thiếu từ thì không thể điểm cao
    if coverage < 1:
        return round(coverage * 70, 2)

    phrase = " ".join(q_tokens)

    # 1. Đủ cụm nguyên văn, đúng thứ tự liên tiếp
    # Câu dài không bị trừ điểm.
    if phrase in text_norm:
        return 100

    # 2. Đủ từ và đúng thứ tự, nhưng có chữ chen giữa
    ordered_score = longest_ordered_subsequence_score(q_tokens, t_tokens)

    if ordered_score == 1:
        gaps = sum(
            max(0, b - a - 1)
            for a, b in zip(positions, positions[1:])
        )

        # càng ít chữ chen giữa càng cao
        return round(max(90, 99 - gaps), 2)

    # 3. Đủ từ nhưng sai thứ tự
    # Vẫn có điểm, nhưng đứng dưới nhóm đúng thứ tự.
    return round(75 + ordered_score * 10, 2)


def run_search(query, search_df, name_col, hs_col):
    q_norm = normalize_text(query)

    result = search_df.copy()

    result["Tên chuẩn hóa"] = result[name_col].astype(str).apply(normalize_text)

    result["Điểm giống"] = result["Tên chuẩn hóa"].apply(
        lambda x: ordered_match_score(q_norm, x)
    )

    result = result[result["Điểm giống"] > 0].copy()

    result[hs_col] = (
        result[hs_col]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
    )

    result = result.sort_values(
        by=["Điểm giống"],
        ascending=False
    )

    result = result.drop_duplicates(
        subset=[name_col, hs_col],
        keep="first"
    )

    return result.head(50)


def show_results(top_result, search_text, search_df, name_col, hs_col):
    if len(top_result) == 0:
        st.warning("Không tìm thấy kết quả phù hợp.")
        return

    hs_stats = top_result[hs_col].value_counts().reset_index()
    hs_stats.columns = ["Mã HS", "Số lần xuất hiện"]

    best_hs = hs_stats.iloc[0]["Mã HS"]
    best_count = hs_stats.iloc[0]["Số lần xuất hiện"]

    st.success(
        f"💡 Gợi ý mã đáng tin nhất: **{best_hs}** — xuất hiện {best_count} lần"
    )

    with st.expander("📊 Thống kê mã HS — click để xem sản phẩm"):
        st.caption("Bấm một mã HS để xem sản phẩm có mã đó trong kết quả.")

        for _, r in hs_stats.iterrows():
            col1, col2, col3 = st.columns([3, 2, 2])

            col1.write(f"**{r['Mã HS']}**")
            col2.write(f"{r['Số lần xuất hiện']} lần")

            if col3.button("🔍 Xem sản phẩm", key=f"hs_btn_{r['Mã HS']}"):
                st.session_state["hs_filter"] = r["Mã HS"]
                st.rerun()

    with st.expander("🧩 Nhóm các mã HS khác nhau"):
        grouped = top_result.groupby(hs_col).size().reset_index(name="Số dòng")
        st.dataframe(grouped, use_container_width=True, hide_index=True)

    st.subheader(f"Tìm thấy {len(top_result)} kết quả phù hợp nhất")

    for _, row in top_result.iterrows():
        highlighted = highlight_text(row[name_col], search_text)

        st.markdown(
            f"""
            <div class="card">
                <div class="hs">Mã HS: {html.escape(str(row[hs_col]))}</div>
                <p><b>Tên hàng:</b> {highlighted}</p>
                <p class="score">
                    Độ giống: {round(row['Điểm giống'], 1)}%
                    | Sheet: {html.escape(str(row['Sheet']))}
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        with st.expander("Xem toàn bộ thông tin dòng này"):
            st.write(row.drop(labels=["Tên chuẩn hóa"]).to_frame("Giá trị"))


# ── Main ─────────────────────────────────────────────────
if os.path.exists(DATA_FILE):
    search_df, sheet_names = load_excel(DATA_FILE)

    name_col = find_name_col(search_df.columns)
    hs_col = find_hs_col(search_df.columns)

    tab_search, tab_sheet = st.tabs(["🔎 Tra cứu", "📋 Xem dữ liệu sheet"])

    with tab_search:
        if "search_from_history" in st.session_state:
            st.session_state["search_text"] = st.session_state.pop(
                "search_from_history"
            )

        search_text = st.text_input(
            "🔎 Nhập tên hàng cần tra",
            key="search_text",
            placeholder="Ví dụ: phớt sắt, xylanh, hydraulic cylinder..."
        )

        if search_text:
            save_history(search_text)

            if not name_col:
                st.error("Không tìm thấy cột tên hàng.")

            elif not hs_col:
                st.error("Không tìm thấy cột mã HS.")

            else:
                top_result = run_search(
                    search_text,
                    search_df,
                    name_col,
                    hs_col
                )

                st.session_state["last_top_result"] = top_result

                hs_filter_val = st.session_state.get("hs_filter", "")

                if hs_filter_val:
                    st.subheader(
                        f"📦 Sản phẩm có mã HS: {hs_filter_val} — trong kết quả tìm kiếm"
                    )

                    if st.button("← Quay lại toàn bộ kết quả"):
                        st.session_state["hs_filter"] = ""
                        st.rerun()

                    else:
                        filtered = top_result[top_result[hs_col] == hs_filter_val]

                        st.info(f"Tìm thấy **{len(filtered)}** sản phẩm")

                        for _, row in filtered.iterrows():
                            highlighted = highlight_text(row[name_col], search_text)

                            st.markdown(
                                f"""
                                <div class="card">
                                    <div class="hs">Mã HS: {html.escape(str(row[hs_col]))}</div>
                                    <p><b>Tên hàng:</b> {highlighted}</p>
                                    <p class="score">
                                        Độ giống: {round(row['Điểm giống'], 1)}%
                                        | Sheet: {html.escape(str(row['Sheet']))}
                                    </p>
                                </div>
                                """,
                                unsafe_allow_html=True
                            )

                else:
                    show_results(top_result, search_text, search_df, name_col, hs_col)

        else:
            st.session_state["hs_filter"] = ""

            st.info("Nhập tên hàng vào ô bên trên để bắt đầu tra cứu.")

            if os.path.exists(HISTORY_FILE):
                history_df = pd.read_csv(HISTORY_FILE)

                with st.expander("🕘 Lịch sử tìm kiếm — click để tra lại", expanded=True):
                    st.caption("Bấm nút để tìm kiếm lại từ khóa cũ.")

                    for _, hr in history_df.iterrows():
                        kw = hr["Từ khóa"]

                        c1, c2 = st.columns([5, 1])

                        c1.write(kw)

                        if c2.button("🔍", key=f"hist_{kw}"):
                            st.session_state["search_from_history"] = kw
                            st.rerun()

    with tab_sheet:
        selected_sheet = st.selectbox("Chọn sheet để xem", sheet_names)

        view_df = load_sheet(DATA_FILE, selected_sheet)

        st.caption(
            f"Sheet: {selected_sheet} — {len(view_df)} dòng | Tổng {len(sheet_names)} sheet"
        )

        st.dataframe(
            view_df,
            use_container_width=True,
            hide_index=True
        )

else:
    st.info("Hãy upload file Excel để bắt đầu.")