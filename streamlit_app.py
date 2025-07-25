import streamlit as st
import pandas as pd
import re
import io
import phonenumbers
from phonenumbers import geocoder

# Danh sách mã quốc gia phổ biến để tự động thêm dấu +
COUNTRY_CODES = {
    '886': 'Taiwan',
    '1': 'USA/Canada',
    '81': 'Japan',
    '82': 'South Korea',
    '85': 'Hong Kong',
    '86': 'China',
    '855': 'Cambodia',
    '856': 'Laos',
    '95': 'Myanmar',
    '44': 'UK',
    '61': 'Australia',
    '65': 'Singapore',
    '66': 'Thailand',
}
# Bản đồ chuyển đổi đầu số cũ ➜ đầu số mới tại Việt Nam
VIETNAM_OLD_PREFIX_MAP = {
    '0162': '032', '0163': '033', '0164': '034',
    '0165': '035', '0166': '036', '0167': '037',
    '0168': '038', '0169': '039',
    '0120': '070', '0121': '079', '0122': '077',
    '0126': '076', '0128': '078',
    '0123': '083', '0124': '084', '0125': '085',
    '0127': '081', '0129': '082',
    '0186': '056', '0188': '058',
    '0199': '059'
}
def normalize_phone(phone):
    if pd.isna(phone):
        return None
    phone = str(phone).strip()
    phone = phone.replace('O', '0').replace('o', '0') 
    phone = re.sub(r'[^\d+]', '', phone)
    phone = phone.replace("’", "").replace("‘", "")  # loại bỏ dấu lạ
    phone = phone.lstrip("=+'\"")  # loại bỏ các ký tự dính từ Excel

    if phone.startswith('00'):
        phone = '+' + phone[2:]

    # 🔄 Nếu số bắt đầu bằng 84 và đủ dài → thêm lại tiền tố 0 để trigger map đầu số cũ
    if phone.startswith('84') and len(phone) >= 11:
        phone = '0' + phone[2:]

    # 🔁 Chuyển đầu số cũ sang đầu số mới nếu có
    for old_prefix, new_prefix in VIETNAM_OLD_PREFIX_MAP.items():
        if phone.startswith(old_prefix) and len(phone) == 11:
            phone = new_prefix + phone[4:]
            break

    # 🇻🇳 Chuẩn hóa +84 ➜ 0
    if phone.startswith('+84'):
        phone = '0' + phone[3:]
    elif phone.startswith('84') and len(phone) in [10, 11]:
        phone = '0' + phone[2:]

    # ✅ Check số Việt Nam (di động & bàn)
    if (phone.startswith('02') and len(phone) == 11) or \
       (phone.startswith(('03', '04', '05', '06', '07', '08', '09')) and len(phone) == 10):
        return phone

    # 📦 Nếu 9 số, thêm 0 rồi thử lại
    if len(phone) == 9 and phone[0] in '3456789':
        phone = '0' + phone
        if (phone.startswith('02') and len(phone) == 11) or \
           (phone.startswith(('03', '04', '05', '06', '07', '08', '09')) and len(phone) == 10):
            return phone

    # 🌍 Xử lý số quốc tế dạng +...
    if phone.startswith('+'):
        try:
            parsed = phonenumbers.parse(phone, None)
            if phonenumbers.is_valid_number(parsed):
                country = geocoder.description_for_number(parsed, 'en')
                return f"{phone} / {country}"
        except:
            return None

    # ➕ Nếu không có dấu + nhưng là mã quốc gia
    for code in sorted(COUNTRY_CODES.keys(), key=lambda x: -len(x)):
        if phone.startswith(code) and len(phone) >= len(code) + 7:
            fake_plus = '+' + phone
            try:
                parsed = phonenumbers.parse(fake_plus, None)
                if phonenumbers.is_valid_number(parsed):
                    country = geocoder.description_for_number(parsed, 'en')
                    return f"{fake_plus} / {country}"
            except:
                continue
    # ❌ Không hợp lệ
    return None



# Giao diện Streamlit
st.set_page_config(page_title="Chuẩn hóa SĐT từ file Excel", layout="wide")
st.title("📱 Chuẩn Hóa Số Điện Thoại Theo Cột Bạn Chọn")

uploaded_file = st.file_uploader("📥 Kéo thả file Excel có nhiều sheet", type=["xlsx"])
st.markdown("---")
st.subheader("📲 Hoặc chuẩn hóa SĐT từ danh sách bạn nhập bên dưới (không cần file)")
manual_input = st.text_area("📥 Nhập danh sách SĐT, mỗi dòng 1 số", height=200, placeholder="vd:\n0912345678\n+886912345678")

if st.button("🚀 Chuẩn hóa danh sách nhập tay"):
    if manual_input.strip() == "":
        st.warning("⚠️ Bạn chưa nhập số nào cả!")
    else:
        raw_numbers = [line.strip() for line in manual_input.splitlines() if line.strip()]
        normalized_numbers = [normalize_phone(num) for num in raw_numbers]

        result_manual_df = pd.DataFrame({
            "Giá trị gốc ban đầu": raw_numbers,
            "SĐT đã chuẩn hóa": normalized_numbers
        })

        st.success("✅ Đã chuẩn hóa xong danh sách bạn nhập.")
        st.dataframe(result_manual_df, use_container_width=True)

        buffer_manual = io.BytesIO()
        result_manual_df.to_excel(buffer_manual, index=False)
        buffer_manual.seek(0)

        st.download_button(
            "📥 Tải danh sách đã chuẩn hóa (nhập tay)",
            data=buffer_manual.getvalue(),
            file_name="sdt_nhap_tay_chuan_hoa.xlsx",
            key="download_manual"
        )

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sample_df = xls.parse(xls.sheet_names[0])  # Đọc sheet đầu tiên
    st.subheader(f"📄 Sheet đầu tiên: `{xls.sheet_names[0]}`")
    # st.dataframe(sample_df.head())
    st.dataframe(sample_df, use_container_width=True, height=500)

    selected_col = st.selectbox("🔍 Chọn cột chứa số điện thoại cần chuẩn hóa:", sample_df.columns)

    # ✅ Sau khi chuẩn hóa
    if st.button("🚀 Bắt đầu chuẩn hóa"):
        all_data = []
        for sheet in xls.sheet_names:
            try:
                df = xls.parse(sheet)
                if selected_col not in df.columns:
                    st.warning(f"⚠️ Sheet '{sheet}' không có cột '{selected_col}'")
                    continue
                df["SĐT đã chuẩn hóa"] = df[selected_col].apply(normalize_phone)
                df["Giá trị gốc ban đầu"] = df[selected_col]
                df["Tên sheet"] = sheet
                all_data.append(df)
            except Exception as e:
                st.warning(f"❌ Lỗi ở sheet '{sheet}': {e}")

        if all_data:
            result_df = pd.concat(all_data, ignore_index=True)
            st.session_state["result_df"] = result_df  # 🔒 Lưu vào session
            st.success("✅ Đã chuẩn hóa xong toàn bộ dữ liệu!")

    # ✅ Hiển thị nếu đã có result_df
    if "result_df" in st.session_state:
        result_df = st.session_state["result_df"]
        st.dataframe(result_df, use_container_width=True)

        buffer = io.BytesIO()
        result_df.to_excel(buffer, index=False)
        st.download_button(
            "📤 Tải file kết quả về Excel",
            data=buffer.getvalue(),
            file_name="ket_qua_chuan_hoa_sdt.xlsx",
            key="download_normalized"
        )

        # Nút lọc dòng hợp lệ
        if st.button("🧹 Lọc và tải danh sách sạch"):
            clean_df = result_df.dropna(subset=["SĐT đã chuẩn hóa"]).reset_index(drop=True)
            st.session_state["clean_df"] = clean_df

        # Nếu đã lọc thì hiển thị nút tải
        if "clean_df" in st.session_state:
            clean_df = st.session_state["clean_df"]
            st.success(f"✅ Đã lọc xong, còn lại {len(clean_df)} dòng hợp lệ.")
            st.dataframe(clean_df, use_container_width=True)

            buffer_clean = io.BytesIO()
            clean_df.to_excel(buffer_clean, index=False)
            buffer_clean.seek(0)

            st.download_button(
                "📥 Tải danh sách sạch (không có dòng None)",
                data=buffer_clean.getvalue(),
                file_name="sdt_sach_khong_none.xlsx",
                key="download_cleaned_clean"
            )
        # ===================== 📌 BẢNG TỔNG HỢP TOÀN BỘ GỐC + CHUẨN HÓA =====================
        if "result_df" in st.session_state:
            full_df = st.session_state["result_df"].copy()
        
            # Hiển thị thêm cột giá trị gốc (nếu chưa có)
            if "Giá trị gốc ban đầu" not in full_df.columns and selected_col in full_df.columns:
                full_df["Giá trị gốc ban đầu"] = full_df[selected_col]
        
            # Cho đẹp: đưa cột "Giá trị gốc" và "SĐT đã chuẩn hóa" ra đầu
            cols = full_df.columns.tolist()
            ordered_cols = ["Giá trị gốc ban đầu", "SĐT đã chuẩn hóa"] + [c for c in cols if c not in ["Giá trị gốc ban đầu", "SĐT đã chuẩn hóa"]]
            full_df = full_df[ordered_cols]
        
            st.subheader("📊 Bảng đầy đủ (toàn bộ dòng + giá trị gốc + chuẩn hóa)")
            st.dataframe(full_df, use_container_width=True, height=600)
        
            buffer_full = io.BytesIO()
            full_df.to_excel(buffer_full, index=False)
            buffer_full.seek(0)
        
            st.download_button(
                "📥 Tải bảng đầy đủ (có giá trị gốc + chuẩn hóa)",
                data=buffer_full.getvalue(),
                file_name="bang_day_du_co_goc_va_chuan_hoa.xlsx",
                key="download_full_goc_va_chuan"
            )
                    


    else:
        st.error("❌ Không có sheet nào được xử lý thành công.")
else:
    st.info("📂 Vui lòng upload file Excel để bắt đầu.")
