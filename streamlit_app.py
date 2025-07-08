import streamlit as st
import pandas as pd
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="📊 Thống kê tham gia lớp học", layout="wide")
st.title("📊 Tổng hợp số lần tham gia & các buổi học theo nhân viên")

# ======================= Hàm đọc toàn bộ tên từ tất cả sheet =======================
def extract_all_names(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    result = []
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2, usecols="A", engine="openpyxl", dtype=str)

            for name in df.iloc[:, 0]:
                result.append((name, sheet_name))
        except Exception as e:
            st.warning(f"❌ Lỗi sheet `{sheet_name}` trong file `{uploaded_file.name}`: {e}")
    return result

# ======================= Hàm xuất ra Excel =======================
def to_excel_file(df, sheet_name="ThamGia"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ======================= Giao diện =======================
uploaded_files = st.file_uploader("📤 Tải lên các file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_names = []

    # Gộp toàn bộ tên và sheet
    for file in uploaded_files:
        extracted = extract_all_names(file)
        all_names.extend(extracted)

    if all_names:
        name_sheet_map = defaultdict(list)
        count_map = defaultdict(int)

        for name, sheet in all_names:
            if pd.notna(name):
                name_str = str(name).strip()
                name_sheet_map[name_str].append(sheet)
                count_map[name_str] += 1

        # Tạo DataFrame tổng hợp
        df_summary = pd.DataFrame([
            {
                "Tên Nhân Viên": name,
                "Tổng số lần xuất hiện": count_map[name],
                "Xuất hiện ở các buổi": ", ".join(sorted(set(name_sheet_map[name])))
            }
            for name in sorted(name_sheet_map.keys())
        ])


        # ======================= BẢNG LIỆT KÊ TÊN NHÂN VIÊN & BUỔI HỌC =======================
        st.subheader("📋 Danh sách tên nhân viên và buổi học (không gộp)")

        df_raw = pd.DataFrame([
            {"Tên Nhân Viên": str(name).strip(), "Buổi Học (Sheet)": sheet}
            for name, sheet in all_names if pd.notna(name)
        ])

        st.dataframe(df_raw, use_container_width=True)

        excel_raw = to_excel_file(df_raw, sheet_name="TenVaBuoiHoc")
        st.download_button(
            label="📥 Tải Excel tên & buổi học",
            data=excel_raw,
            file_name="bang_ten_va_buoi_hoc.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
                # ======================= 📌 TỔNG HỢP THAM GIA (KHÔNG THAY ĐỔI TÊN) =======================
        st.subheader("📌 Tổng hợp tham gia lớp học theo từng nhân viên (từ bảng gốc)")

        df_raw["Tên Nhân Viên"] = df_raw["Tên Nhân Viên"].astype(str)

        df_summary_v2 = (
            df_raw.groupby("Tên Nhân Viên")
            .agg(
                **{
                    "Tổng số lần xuất hiện": ("Buổi Học (Sheet)", "count"),
                    "Số buổi học khác nhau": ("Buổi Học (Sheet)", lambda x: len(set(x))),
                    "Danh sách buổi học": ("Buổi Học (Sheet)", lambda x: ", ".join(sorted(set(map(str, x)))))
                }
            )
            .reset_index()
            .sort_values(by="Tổng số lần xuất hiện", ascending=False)
        )

        st.dataframe(df_summary_v2, use_container_width=True)

        excel_summary_v2 = to_excel_file(df_summary_v2, sheet_name="TongHopV2")
        st.download_button(
            label="📥 Tải Excel (Tổng hợp V2)",
            data=excel_summary_v2,
            file_name="tong_hop_v2_nhan_vien.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # ======================= ✅ BẢNG CHẤM CÔNG (THEO TÊN & SHEET) =======================
        st.subheader("✅ Bảng chấm công (Tên nhân viên x Buổi học)")

        df_raw["Tên Nhân Viên"] = df_raw["Tên Nhân Viên"].astype(str)
        df_raw["Buổi Học (Sheet)"] = df_raw["Buổi Học (Sheet)"].astype(str)

        # Tạo bảng pivot (chấm công)
        df_attendance_matrix = pd.pivot_table(
            df_raw,
            index="Tên Nhân Viên",
            columns="Buổi Học (Sheet)",
            aggfunc="size",
            fill_value=0
        )

        # Đổi số > 0 thành "✅"
        df_attendance_matrix = df_attendance_matrix.applymap(lambda x: "✅" if x > 0 else "")

        # Thêm cột tổng
        df_attendance_matrix["Tổng buổi tham gia"] = (df_attendance_matrix == "✅").sum(axis=1)

        # Reset index để hiển thị đẹp
        df_attendance_matrix = df_attendance_matrix.reset_index()

        st.dataframe(df_attendance_matrix, use_container_width=True)

        # Tải về
        excel_attendance_matrix = to_excel_file(df_attendance_matrix, sheet_name="ChamCong")
        st.download_button(
            label="📥 Tải Excel (Bảng chấm công)",
            data=excel_attendance_matrix,
            file_name="bang_cham_cong_nhan_vien.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("⚠️ Không tìm thấy tên nào trong cột A.")
