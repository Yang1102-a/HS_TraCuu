import streamlit as st
import pandas as pd
import os
import re
import html
import unicodedata
from rapidfuzz import fuzz

st.set_page_config(page_title="Tra cứu mã HS", layout="wide")

st.markdown("""
<style>
.stApp { background: #0f1117; }
div[data-testid="stTextInput"] input {
    font-size: 18px;
    padding: 14px;
}
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
.score {
    color: #aaa;
}
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


def remove_accents(text):
    text = str(text)
    text = unicodedata.normalize("NFD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    return text


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
        text = str(col).lower()
        if "tên hàng" in text:
            return col

    for col in columns:
        text = str(col).lower()
        if "tên" in text or "hàng" in text or "hang" in text:
            return col

    return None


def find_hs_col(columns):
    for col in columns:
        text = str(col).lower().replace(" ", "")
        if "mãhs" in text or "mahs" in text or text == "hs" or "hs" in text:
            return col

    return None


def highlight_text(original_text, query):
    safe_text = html.escape(str(original_text))
    words = normalize_text(query).split()

    for word in words:
        if len(word) >= 2:
            safe_text = re.sub(
                f"({re.escape(word)})",
                r"<mark>\1</mark>",
                safe_text,
                flags=re.IGNORECASE
            )

    return safe_text


def save_history(keyword):
    new_row = pd.DataFrame([{"Từ khóa": keyword}])

    if os.path.exists(HISTORY_FILE):
        old = pd.read_csv(HISTORY_FILE)
        history = pd.concat([new_row, old], ignore_index=True)
    else:
        history = new_row

    history = history.drop_duplicates().head(20)
    history.to_csv(HISTORY_FILE, index=False)


if os.path.exists(DATA_FILE):

    excel_file = pd.ExcelFile(DATA_FILE)

    all_data = []

    for sheet in excel_file.sheet_names:
        temp_df = pd.read_excel(DATA_FILE, sheet_name=sheet)
        temp_df = temp_df.dropna(how="all")
        temp_df["Sheet"] = sheet
        all_data.append(temp_df)

    search_df = pd.concat(all_data, ignore_index=True)

    selected_sheet = st.selectbox("Chọn sheet để xem", excel_file.sheet_names)

    view_df = pd.read_excel(DATA_FILE, sheet_name=selected_sheet)
    view_df = view_df.dropna(how="all")

    name_col = find_name_col(search_df.columns)
    hs_col = find_hs_col(search_df.columns)

    st.caption(
        f"Đang xem sheet: {selected_sheet} | Tra cứu trên toàn bộ {len(excel_file.sheet_names)} sheet"
    )

    search_text = st.text_input(
        "🔎 Nhập tên hàng cần tra",
        placeholder="Ví dụ: phớt sắt, xylanh, hydraulic cylinder, shaft pump..."
    )

    if search_text:

        save_history(search_text)

        if not name_col:
            st.error("Không tìm thấy cột tên hàng.")

        elif not hs_col:
            st.error("Không tìm thấy cột mã HS.")

        else:
            query_norm = normalize_text(search_text)

            result = search_df.copy()

            result["Tên chuẩn hóa"] = result[name_col].astype(str).apply(normalize_text)

            result["Điểm giống"] = result["Tên chuẩn hóa"].apply(
                lambda x: max(
                    fuzz.token_set_ratio(query_norm, x),
                    fuzz.partial_ratio(query_norm, x)
                )
            )

            result = result[result["Điểm giống"] > 0]

            result[hs_col] = result[hs_col].astype(str).str.replace(".0", "", regex=False)

            result = result.sort_values(
                by="Điểm giống",
                ascending=False
            )

            result = result.drop_duplicates(
                subset=[name_col, hs_col],
                keep="first"
            )

            top_result = result.head(50)

            st.subheader(f"Tìm thấy {len(top_result)} kết quả phù hợp nhất")

            if len(top_result) > 0:

                hs_stats = top_result[hs_col].value_counts().reset_index()
                hs_stats.columns = ["Mã HS", "Số lần xuất hiện"]

                best_hs = hs_stats.iloc[0]["Mã HS"]
                best_count = hs_stats.iloc[0]["Số lần xuất hiện"]

                st.success(
                    f"💡 Gợi ý mã đáng tin nhất: {best_hs} — xuất hiện {best_count} lần trong nhóm kết quả phù hợp"
                )

                with st.expander("📊 Thống kê mã HS trong kết quả"):
                    st.dataframe(
                        hs_stats,
                        use_container_width=True,
                        hide_index=True
                    )

                grouped = top_result.groupby(hs_col).size().reset_index(name="Số dòng")

                with st.expander("🧩 Nhóm các mã HS khác nhau của mặt hàng gần giống"):
                    st.dataframe(
                        grouped,
                        use_container_width=True,
                        hide_index=True
                    )

                for index, row in top_result.iterrows():

                    hs_value = row[hs_col]
                    name_value = row[name_col]
                    sheet_value = row["Sheet"]
                    score_value = round(row["Điểm giống"], 1)

                    highlighted_name = highlight_text(name_value, search_text)

                    st.markdown(
                        f"""
                        <div class="card">
                            <div class="hs">Mã HS: {html.escape(str(hs_value))}</div>
                            <p><b>Tên hàng:</b> {highlighted_name}</p>
                            <p class="score">Độ giống: {score_value}% | Sheet: {html.escape(str(sheet_value))}</p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    with st.expander("Xem toàn bộ thông tin dòng này"):
                        st.write(row.drop(labels=["Tên chuẩn hóa"]).to_frame("Giá trị"))

            else:
                st.warning("Không tìm thấy kết quả phù hợp.")

    else:
        st.subheader("Dữ liệu sheet đang chọn")

        st.dataframe(
            view_df,
            use_container_width=True,
            hide_index=True
        )

        if os.path.exists(HISTORY_FILE):
            history_df = pd.read_csv(HISTORY_FILE)
            with st.expander("🕘 Lịch sử tìm kiếm gần đây"):
                st.dataframe(
                    history_df,
                    use_container_width=True,
                    hide_index=True
                )

else:
    st.info("Hãy upload file Excel để bắt đầu.")