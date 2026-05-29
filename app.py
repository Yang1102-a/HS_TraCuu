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
div[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }
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


def remove_accents(text):
    text = unicodedata.normalize("NFD", str(text))
    return text.encode("ascii", "ignore").decode("utf-8")


def normalize_text(text):
    text = remove_accents(str(text).lower())
    replacements = {
        "xylanh": "xy lanh", "xi lanh": "xy lanh", "cylinder": "xy lanh",
        "hydraulic": "thuy luc", "shaft": "truc", "pump": "bom",
        "seal": "phot", "steel": "thep", "iron": "sat", "bolt": "bulong",
        "screw": "vit", "bearing": "vong bi", "gear": "banh rang",
        "chain": "xich", "filter": "loc",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


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


def delete_history_keyword(keyword):
    if not os.path.exists(HISTORY_FILE):
        return
    history = pd.read_csv(HISTORY_FILE)
    history = history[history["Từ khóa"] != keyword]
    history.to_csv(HISTORY_FILE, index=False)


def clear_history():
    if os.path.exists(HISTORY_FILE):
        os.remove(HISTORY_FILE)


@st.cache_data
def load_excel(path):
    xf = pd.ExcelFile(path)
    frames = []
    for sheet in xf.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet).dropna(how="all")
        df["Sheet"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True), xf.sheet_names


@st.cache_data
def load_sheet(path, sheet_name):
    return pd.read_excel(path, sheet_name=sheet_name).dropna(how="all")


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
    if coverage < 1:
        return round(coverage * 70, 2)
    phrase = " ".join(q_tokens)
    if phrase in text_norm:
        return 100
    ordered_score = longest_ordered_subsequence_score(q_tokens, t_tokens)
    if ordered_score == 1:
        gaps = sum(max(0, b - a - 1) for a, b in zip(positions, positions[1:]))
        return round(max(90, 99 - gaps), 2)
    return round(75 + ordered_score * 10, 2)


def run_search(query, search_df, name_col, hs_col):
    q_norm = normalize_text(query)
    result = search_df.copy()
    result["Tên chuẩn hóa"] = result[name_col].astype(str).apply(normalize_text)
    result["Điểm giống"] = result["Tên chuẩn hóa"].apply(
        lambda x: ordered_match_score(q_norm, x)
    )
    result = result[result["Điểm giống"] > 0].copy()
    result[hs_col] = result[hs_col].astype(str).str.replace(r"\.0$", "", regex=True)
    result = result.sort_values("Điểm giống", ascending=False)
    result = result.drop_duplicates(subset=[name_col, hs_col], keep="first")
    return result.head(50)


def show_results(top_result, search_text, search_df, name_col, hs_col):
    if len(top_result) == 0:
        st.warning("Không tìm thấy kết quả phù hợp.")
        return

    hs_stats = top_result[hs_col].value_counts().reset_index()
    hs_stats.columns = ["Mã HS", "Số lần xuất hiện"]
    best_hs = hs_stats.iloc[0]["Mã HS"]
    best_count = hs_stats.iloc[0]["Số lần xuất hiện"]
    st.success(f"💡 Gợi ý mã đáng tin nhất: **{best_hs}** — xuất hiện {best_count} lần")

    with st.expander("📊 Thống kê mã HS — click để xem sản phẩm"):
        for _, r in hs_stats.iterrows():
            col1, col2, col3 = st.columns([3, 2, 2])
            col1.write(f"**{r['Mã HS']}**")
            col2.write(f"{r['Số lần xuất hiện']} lần")
            if col3.button("🔍 Xem sản phẩm", key=f"hs_btn_{r['Mã HS']}"):
                st.session_state["hs_filter"] = r["Mã HS"]
                st.rerun()

    with st.expander("🎯 Xếp hạng mã HS theo mức phù hợp với tìm kiếm"):
        hs_accuracy = (
            top_result.groupby(hs_col)["Điểm giống"]
            .agg(["mean", "count", "max"])
            .reset_index()
            .rename(columns={
                hs_col: "Mã HS",
                "mean": "Điểm TB (%)",
                "count": "Số kết quả",
                "max": "Điểm cao nhất (%)",
            })
            .sort_values("Điểm TB (%)", ascending=False)
        )
        hs_accuracy["Điểm TB (%)"] = hs_accuracy["Điểm TB (%)"].apply(lambda x: f"{x:.1f}")
        hs_accuracy["Điểm cao nhất (%)"] = hs_accuracy["Điểm cao nhất (%)"].apply(lambda x: f"{x:.1f}")
        st.dataframe(hs_accuracy, use_container_width=True, hide_index=True)

    st.subheader(f"Tìm thấy {len(top_result)} kết quả phù hợp nhất")

    for _, row in top_result.iterrows():
        highlighted = highlight_text(row[name_col], search_text)
        st.markdown(f"""
        <div class="card">
            <div class="hs">Mã HS: {html.escape(str(row[hs_col]))}</div>
            <p><b>Tên hàng:</b> {highlighted}</p>
            <p class="score">
                Độ giống: {round(row['Điểm giống'], 1)}%
                | Sheet: {html.escape(str(row['Sheet']))}
            </p>
        </div>
        """, unsafe_allow_html=True)
        with st.expander("Xem toàn bộ thông tin dòng này"):
            st.write(row.drop(labels=["Tên chuẩn hóa"]).to_frame("Giá trị"))


if os.path.exists(DATA_FILE):
    search_df, sheet_names = load_excel(DATA_FILE)
    name_col = find_name_col(search_df.columns)
    hs_col = find_hs_col(search_df.columns)

    tab_search, tab_sheet = st.tabs(["🔎 Tra cứu", "📋 Xem dữ liệu sheet"])

    with tab_search:
        if "search_from_history" in st.session_state:
            st.session_state["search_text"] = st.session_state.pop("search_from_history")

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
                top_result = run_search(search_text, search_df, name_col, hs_col)
                st.session_state["last_top_result"] = top_result
                hs_filter_val = st.session_state.get("hs_filter", "")

                if hs_filter_val:
                    st.subheader(f"📦 Sản phẩm có mã HS: {hs_filter_val}")
                    if st.button("← Quay lại toàn bộ kết quả"):
                        st.session_state["hs_filter"] = ""
                        st.rerun()
                    filtered = top_result[top_result[hs_col] == hs_filter_val]
                    st.info(f"Tìm thấy **{len(filtered)}** sản phẩm")
                    for _, row in filtered.iterrows():
                        highlighted = highlight_text(row[name_col], search_text)
                        st.markdown(f"""
                        <div class="card">
                            <div class="hs">Mã HS: {html.escape(str(row[hs_col]))}</div>
                            <p><b>Tên hàng:</b> {highlighted}</p>
                            <p class="score">
                                Độ giống: {round(row['Điểm giống'], 1)}%
                                | Sheet: {html.escape(str(row['Sheet']))}
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    show_results(top_result, search_text, search_df, name_col, hs_col)
        else:
            st.session_state["hs_filter"] = ""
            st.info("Nhập tên hàng vào ô bên trên để bắt đầu tra cứu.")

            if os.path.exists(HISTORY_FILE):
                history_df = pd.read_csv(HISTORY_FILE)
                with st.expander("🕘 Lịch sử tìm kiếm — click để tra lại", expanded=True):
                    st.caption("Bấm 🔍 để tìm lại, bấm 🗑 để xóa từng dòng.")
                    if st.button("🧹 Xóa toàn bộ lịch sử"):
                        clear_history()
                        st.rerun()
                    for _, hr in history_df.iterrows():
                        kw = hr["Từ khóa"]
                        c1, c2, c3 = st.columns([5, 1, 1])
                        c1.write(kw)
                        if c2.button("🔍", key=f"hist_search_{kw}"):
                            st.session_state["search_from_history"] = kw
                            st.rerun()
                        if c3.button("🗑", key=f"hist_delete_{kw}"):
                            delete_history_keyword(kw)
                            st.rerun()

    with tab_sheet:
        # ── Router: "sheet_detail_row" có giá trị → đang ở trang chi tiết ──
        if "sheet_detail_row" not in st.session_state:
            st.session_state["sheet_detail_row"] = None
        if "sheet_detail_sheet" not in st.session_state:
            st.session_state["sheet_detail_sheet"] = None

        detail_idx   = st.session_state["sheet_detail_row"]
        detail_sheet = st.session_state["sheet_detail_sheet"]

        # ════════════════════════════════════════════
        # TRANG CHI TIẾT
        # ════════════════════════════════════════════
        if detail_idx is not None and detail_sheet is not None:
            detail_df = load_sheet(DATA_FILE, detail_sheet).copy()
            detail_df["Sheet"] = detail_sheet

            d_name_col = find_name_col(detail_df.columns) or name_col
            d_hs_col   = find_hs_col(detail_df.columns) or hs_col

            if d_hs_col and d_hs_col in detail_df.columns:
                detail_df[d_hs_col] = (
                    detail_df[d_hs_col]
                    .astype(str)
                    .str.replace(r"\.0$", "", regex=True)
                )

            row = detail_df.iloc[detail_idx]
            _name_val = str(row[d_name_col]) if d_name_col and d_name_col in row.index else "—"
            _hs_val   = str(row[d_hs_col])   if d_hs_col   and d_hs_col   in row.index else "—"

            # Nút quay lại
            if st.button("← Quay lại danh sách", type="secondary"):
                st.session_state["sheet_detail_row"]   = None
                st.session_state["sheet_detail_sheet"] = None
                st.rerun()

            st.markdown(f"### 🃏 Chi tiết dòng {detail_idx + 1} — Sheet: {detail_sheet}")

            # Card lớn mã HS + tên hàng
            st.markdown(f"""
            <div class="card">
                <div class="hs">Mã HS: {html.escape(_hs_val)}</div>
                <p style="font-size:16px"><b>Tên hàng:</b><br>{html.escape(_name_val)}</p>
                <p class="score">Sheet: {html.escape(detail_sheet)} | Dòng: {detail_idx + 1}</p>
            </div>
            """, unsafe_allow_html=True)

            # Toàn bộ các trường còn lại dạng thẻ nhỏ theo lưới 2 cột
            skip_cols = {"Tên chuẩn hóa", "Điểm giống", "Sheet",
                         d_name_col or "", d_hs_col or ""}
            other_fields = [(c, row[c]) for c in row.index if c not in skip_cols]

            if other_fields:
                st.markdown("#### 📋 Thông tin khác")
                cols = st.columns(2)
                for i, (field, val) in enumerate(other_fields):
                    with cols[i % 2]:
                        st.markdown(f"""
                        <div class="card" style="padding:12px 16px; margin-bottom:10px">
                            <p class="score" style="margin:0;font-size:12px">{html.escape(str(field))}</p>
                            <p style="margin:4px 0 0 0;font-size:15px;font-weight:600">
                                {html.escape(str(val))}
                            </p>
                        </div>
                        """, unsafe_allow_html=True)

            # Điều hướng dòng trước / sau
            st.markdown("---")
            total_rows = len(detail_df)
            nav_prev, nav_next = st.columns(2)
            with nav_prev:
                if detail_idx > 0:
                    if st.button("⬅ Dòng trước"):
                        st.session_state["sheet_detail_row"] = detail_idx - 1
                        st.rerun()
            with nav_next:
                if detail_idx < total_rows - 1:
                    if st.button("Dòng tiếp ➡"):
                        st.session_state["sheet_detail_row"] = detail_idx + 1
                        st.rerun()

        # ════════════════════════════════════════════
        # TRANG BẢNG (mặc định)
        # ════════════════════════════════════════════
        else:
            selected_sheet = st.selectbox("Chọn sheet để xem", sheet_names)
            view_df = load_sheet(DATA_FILE, selected_sheet).copy()
            view_df["Sheet"] = selected_sheet

            sheet_name_col = find_name_col(view_df.columns) or name_col
            sheet_hs_col   = find_hs_col(view_df.columns) or hs_col

            if sheet_hs_col and sheet_hs_col in view_df.columns:
                view_df[sheet_hs_col] = (
                    view_df[sheet_hs_col]
                    .astype(str)
                    .str.replace(r"\.0$", "", regex=True)
                )

            st.caption(
                f"Sheet: **{selected_sheet}** — {len(view_df)} dòng | "
                f"Tổng {len(sheet_names)} sheet"
            )

            # Nhập dòng + nút mở thẻ — đặt ngay trên bảng cho gọn
            ca, cb, cc = st.columns([1, 1, 4])
            with ca:
                row_num = st.number_input(
                    "Dòng",
                    min_value=1, max_value=len(view_df), step=1,
                    key=f"row_input_{selected_sheet}",
                    label_visibility="visible",
                )
            with cb:
                st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
                if st.button("🃏 Mở thẻ", key=f"open_card_{selected_sheet}", type="primary"):
                    st.session_state["sheet_detail_row"]   = int(row_num) - 1
                    st.session_state["sheet_detail_sheet"] = selected_sheet
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

            # Bảng gọn bên dưới
            display_cols = [c for c in view_df.columns if c != "Sheet"]
            table_df = view_df[display_cols].reset_index(drop=True)
            table_df.index = table_df.index + 1
            table_df.index.name = "DÒNG"
            st.dataframe(table_df, use_container_width=True, hide_index=False)

else:
    st.info("Hãy upload file Excel để bắt đầu.")