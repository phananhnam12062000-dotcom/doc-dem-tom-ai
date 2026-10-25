# -*- coding: utf-8 -*-
import os
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DEFAULT_EXCEL_FILENAME = "Bao_Cao_Tom_Tu_Dong.xlsx"

# =============================================================================
# BẢNG MÀU & FONT CHUẨN THỦY SẢN (NAVY BLUE #1F4E79 & ZEBRA #F2F5F8)
# =============================================================================
NAVY_BLUE_HEX = "1F4E79"       # Màu xanh dương đậm chủ đạo (#1F4E79)
ZEBRA_LIGHT_HEX = "F2F5F8"     # Màu xanh nhạt xen kẽ (#F2F5F8)
WHITE_HEX = "FFFFFF"           # Màu nền trắng (#FFFFFF)
ORANGE_HIGHLIGHT_HEX = "FCE4D6"  # Màu cam nhạt nổi bật cho Khối tổng số Sheet 2
BORDER_GRAY_HEX = "D0D7DE"     # Viền mảnh thẩm mỹ

FONT_FAMILY = "Segoe UI"


def _auto_fit_columns(ws, min_width: int = 15, max_width: int = 44):
    """
    Tự động căn rộng kích thước cột (Auto-fit column width) dựa theo độ dài của chữ,
    tuyệt đối không để chữ bị che khuất hoặc tràn lề.
    """
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = min_width
        for cell in col:
            # Bỏ qua dòng tiêu đề lớn gộp ô (dòng 1) để tránh kéo giãn cột A quá mức
            if cell.row in (1, 6, 7) and cell.column == 1:
                continue
            val_str = str(cell.value or "")
            # Nếu ô chứa công thức Excel, ước lượng độ dài hiển thị thực tế
            if val_str.startswith("="):
                val_str = "123,456.78 con/kg"
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 5, min_width), max_width)


def build_shrimp_report_workbook(
    pond_name: str,
    staff_name: str,
    weight_records: list,
    batch_records: list,
    total_batch_count: int = 0,
) -> openpyxl.Workbook:
    """
    Tạo đối tượng openpyxl.Workbook chuẩn 2 Sheet với:
    - Màu chủ đạo Navy Blue (#1F4E79), chữ trắng, in đậm, căn giữa.
    - Dòng dữ liệu xen kẽ trắng (#FFFFFF) và xanh nhạt (#F2F5F8).
    - Font chữ "Segoe UI", tiêu đề cỡ 16, nội dung bảng cỡ 11.
    - Công thức tính toán tự động 100% (COUNT, SUM, AVERAGE, COUNTIF, SUMIF, IFERROR).
    - Tự động căn độ rộng cột (Auto-fit column width).
    """
    wb = openpyxl.Workbook()
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    pond_display = pond_name.strip() if pond_name else "Ao Số 1 - Khu A"
    staff_display = staff_name.strip() if staff_name else "Kỹ thuật viên 1"

    # Định nghĩa các Style chuẩn
    font_title = Font(name=FONT_FAMILY, size=16, bold=True, color=WHITE_HEX)
    font_header = Font(name=FONT_FAMILY, size=11, bold=True, color=WHITE_HEX)
    font_bold = Font(name=FONT_FAMILY, size=11, bold=True, color="1C2833")
    font_Formula_navy = Font(name=FONT_FAMILY, size=11, bold=True, color=NAVY_BLUE_HEX)
    font_size_highlight = Font(name=FONT_FAMILY, size=11, bold=True, color="D35400")
    font_regular = Font(name=FONT_FAMILY, size=11, bold=False, color="1C2833")
    font_strike = Font(name=FONT_FAMILY, size=11, strike=True, color="922B21")

    fill_navy = PatternFill(start_color=NAVY_BLUE_HEX, end_color=NAVY_BLUE_HEX, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_LIGHT_HEX, end_color=ZEBRA_LIGHT_HEX, fill_type="solid")
    fill_white = PatternFill(start_color=WHITE_HEX, end_color=WHITE_HEX, fill_type="solid")
    fill_orange = PatternFill(start_color=ORANGE_HIGHLIGHT_HEX, end_color=ORANGE_HIGHLIGHT_HEX, fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)

    thin_side = Side(border_style="thin", color=BORDER_GRAY_HEX)
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # =========================================================================
    # SHEET 1: CÂN TRỌNG LƯỢNG & GIỚI TÍNH (Mẫu tôm thương phẩm)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "1_Can_Trong_Luong_Gioi_Tinh"
    ws1.views.sheetView[0].showGridLines = True

    # 1. PHẦN THÔNG TIN CHUNG (Dòng 1 đến Dòng 5)
    ws1.merge_cells("A1:D1")
    ws1["A1"] = "BÁO CÁO CÂN TRỌNG LƯỢNG, GIỚI TÍNH & ĐẾM TÔM THÔNG MINH"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_navy
    ws1["A1"].alignment = align_center
    ws1.row_dimensions[1].height = 36

    info_rows = [
        ("Mã Lô / Ao:", pond_display),
        ("Kỹ thuật viên:", staff_display),
        ("Thời gian xuất báo cáo:", now_str),
    ]
    for idx, (lbl, val) in enumerate(info_rows, start=2):
        ws1.cell(row=idx, column=1, value=lbl).font = font_bold
        ws1.cell(row=idx, column=1).alignment = align_left
        ws1.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)
        val_cell = ws1.cell(row=idx, column=2, value=val)
        val_cell.font = font_regular
        val_cell.alignment = align_left
        ws1.row_dimensions[idx].height = 20

    ws1.row_dimensions[5].height = 10

    # 2. BẢNG THỐNG KÊ TỔNG HỢP (Nằm ở trên: Dòng 6 đến Dòng 12 - Dùng công thức Excel)
    ws1.merge_cells("A6:D6")
    ws1["A6"] = "BẢNG THỐNG KÊ TỔNG HỢP MẪU TÔM (CÔNG THỨC TỰ ĐỘNG)"
    for col_i in range(1, 5):
        c = ws1.cell(row=6, column=col_i)
        c.font = font_header
        c.fill = fill_navy
        c.alignment = align_center
        c.border = thin_border
    ws1.row_dimensions[6].height = 26

    detail_start_1 = 15
    detail_end_1 = max(detail_start_1, detail_start_1 + len(weight_records) - 1)

    summary_specs = [
        (
            "Tổng số mẫu đã cân",
            f"=COUNT(B{detail_start_1}:B{detail_end_1})",
            "#,##0",
            "con",
        ),
        (
            "Tổng trọng lượng (g)",
            f"=SUM(B{detail_start_1}:B{detail_end_1})",
            "#,##0.00",
            "g",
        ),
        (
            "Trọng lượng trung bình",
            f"=IFERROR(AVERAGE(B{detail_start_1}:B{detail_end_1}), 0)",
            "0.00",
            "g/con",
        ),
        (
            "Size tôm trung bình (con/kg)",
            "=IFERROR((C7/C8)*1000, 0)",
            "0.0",
            "con/kg",
        ),
        (
            "Tỷ lệ Tôm Đực - Male (%)",
            f'=IFERROR((COUNTIF(C{detail_start_1}:C{detail_end_1}, "Male")/C7)*100, 0)',
            "0.0",
            "%",
        ),
        (
            "Tỷ lệ Tôm Cái - Female (%)",
            f'=IFERROR((COUNTIF(C{detail_start_1}:C{detail_end_1}, "Female")/C7)*100, 0)',
            "0.0",
            "%",
        ),
    ]

    for i, (label_text, formula_str, num_fmt, unit_text) in enumerate(summary_specs):
        r_idx = 7 + i
        row_fill = fill_zebra if i % 2 == 1 else fill_white
        ws1.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=2)

        c_label = ws1.cell(row=r_idx, column=1, value=label_text)
        c_label.font = font_bold if i == 3 else font_regular
        c_label.alignment = align_left

        c_val = ws1.cell(row=r_idx, column=3, value=formula_str)
        c_val.font = font_size_highlight if i == 3 else font_Formula_navy
        c_val.number_format = num_fmt
        c_val.alignment = align_right

        c_unit = ws1.cell(row=r_idx, column=4, value=unit_text)
        c_unit.font = font_regular
        c_unit.alignment = align_center

        for col_i in range(1, 5):
            cell_item = ws1.cell(row=r_idx, column=col_i)
            cell_item.fill = row_fill
            cell_item.border = thin_border
        ws1.row_dimensions[r_idx].height = 22

    ws1.row_dimensions[13].height = 10

    # 3. BẢNG CHI TIẾT CÂN MẪU (Từ Dòng 14: Cột A STT, B Trọng Lượng - g, C Giới Tính, D Thời Gian Nhận Diện)
    headers_sheet1 = ["STT", "Trọng Lượng (g)", "Giới Tính", "Thời Gian Nhận Diện"]
    for col_i, h_text in enumerate(headers_sheet1, start=1):
        c = ws1.cell(row=14, column=col_i, value=h_text)
        c.font = font_header
        c.fill = fill_navy
        c.alignment = align_center
        c.border = thin_border
    ws1.row_dimensions[14].height = 26

    if not weight_records:
        # Nếu chưa có dữ liệu, tạo 1 dòng trống định dạng chuẩn để công thức không bị lỗi #REF!
        c_stt = ws1.cell(row=15, column=1, value=1)
        c_w = ws1.cell(row=15, column=2, value=0.0)
        c_g = ws1.cell(row=15, column=3, value="Chưa rõ")
        c_t = ws1.cell(row=15, column=4, value="--:--:--")
        c_stt.alignment = align_center
        c_w.alignment = align_right
        c_w.number_format = "0.00"
        c_g.alignment = align_center
        c_t.alignment = align_center
        for col_i in range(1, 5):
            cell_item = ws1.cell(row=15, column=col_i)
            cell_item.font = font_regular
            cell_item.fill = fill_white
            cell_item.border = thin_border
        ws1.row_dimensions[15].height = 21
    else:
        for idx, rec in enumerate(weight_records):
            r_idx = detail_start_1 + idx
            row_fill = fill_zebra if idx % 2 == 1 else fill_white

            stt_val = rec.get("stt", idx + 1)
            weight_val = float(rec.get("weight", 0.0))
            gender_val = str(rec.get("gender", "Chưa rõ"))
            time_val = str(rec.get("time", ""))

            c_stt = ws1.cell(row=r_idx, column=1, value=stt_val)
            c_w = ws1.cell(row=r_idx, column=2, value=round(weight_val, 2))
            c_g = ws1.cell(row=r_idx, column=3, value=gender_val)
            c_t = ws1.cell(row=r_idx, column=4, value=time_val)

            c_stt.alignment = align_center
            c_w.alignment = align_right
            c_w.number_format = "0.00"
            c_g.alignment = align_center
            c_t.alignment = align_center

            for col_i in range(1, 5):
                cell_item = ws1.cell(row=r_idx, column=col_i)
                cell_item.font = font_regular
                cell_item.fill = row_fill
                cell_item.border = thin_border
            ws1.row_dimensions[r_idx].height = 21

    # Dòng Tổng cộng / Trung bình cuối bảng Sheet 1 (Sử dụng công thức SUM & AVERAGE)
    footer_row_1 = detail_end_1 + 1
    ws1.cell(row=footer_row_1, column=1, value="TỔNG / TB").alignment = align_center
    c_sum_w = ws1.cell(row=footer_row_1, column=2, value=f"=SUM(B{detail_start_1}:B{detail_end_1})")
    c_sum_w.number_format = "#,##0.00"
    c_sum_w.alignment = align_right
    c_avg_w = ws1.cell(row=footer_row_1, column=3, value=f"=IFERROR(AVERAGE(B{detail_start_1}:B{detail_end_1}), 0)")
    c_avg_w.number_format = '"TB: "0.00" g"'
    c_avg_w.alignment = align_center
    c_cnt_w = ws1.cell(row=footer_row_1, column=4, value=f'=COUNT(B{detail_start_1}:B{detail_end_1})&" mẫu"')
    c_cnt_w.alignment = align_center

    for col_i in range(1, 5):
        cell_item = ws1.cell(row=footer_row_1, column=col_i)
        cell_item.font = font_Formula_navy
        cell_item.fill = fill_zebra
        cell_item.border = thin_border
    ws1.row_dimensions[footer_row_1].height = 24

    _auto_fit_columns(ws1, min_width=18)

    # =========================================================================
    # SHEET 2: ĐẾM SỐ LƯỢNG CỘNG DỒN (Mẻ đếm tốc độ cao)
    # =========================================================================
    ws2 = wb.create_sheet(title="2_Dem_So_Luong_Cong_Don")
    ws2.views.sheetView[0].showGridLines = True

    # 1. PHẦN THÔNG TIN CHUNG (Dòng 1 đến Dòng 5)
    ws2.merge_cells("A1:F1")
    ws2["A1"] = "BÁO CÁO CÂN TRỌNG LƯỢNG, GIỚI TÍNH & ĐẾM TÔM THÔNG MINH"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_navy
    ws2["A1"].alignment = align_center
    ws2.row_dimensions[1].height = 36

    for idx, (lbl, val) in enumerate(info_rows, start=2):
        ws2.cell(row=idx, column=1, value=lbl).font = font_bold
        ws2.cell(row=idx, column=1).alignment = align_left
        ws2.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=6)
        val_cell = ws2.cell(row=idx, column=2, value=val)
        val_cell.font = font_regular
        val_cell.alignment = align_left
        ws2.row_dimensions[idx].height = 20

    ws2.row_dimensions[5].height = 10

    # 2. KHỐI TỔNG SỐ (Dòng 6 đến Dòng 7 - Highlight màu cam nhạt, gộp ô, chữ to in đậm, công thức SUM)
    detail_start_2 = 10
    detail_end_2 = max(detail_start_2, detail_start_2 + len(batch_records) - 1)

    ws2.merge_cells("A6:D7")
    ws2["A6"] = "TỔNG SỐ LƯỢNG TÔM ĐÃ ĐẾM (CON)"
    ws2["A6"].font = Font(name=FONT_FAMILY, size=14, bold=True, color="9C3400")
    ws2["A6"].alignment = align_center

    ws2.merge_cells("E6:F7")
    ws2["E6"] = f"=SUM(C{detail_start_2}:C{detail_end_2})"
    ws2["E6"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="C0392B")
    ws2["E6"].number_format = '#,##0 "con"'
    ws2["E6"].alignment = align_right

    for r_i in (6, 7):
        ws2.row_dimensions[r_i].height = 22
        for c_i in range(1, 7):
            cell_item = ws2.cell(row=r_i, column=c_i)
            cell_item.fill = fill_orange
            cell_item.border = thin_border

    ws2.row_dimensions[8].height = 10

    # 3. BẢNG CHI TIẾT CÁC MẺ ĐẾM (Từ Dòng 9: Cột A STT, B Mã Mẻ, C Số Lượng Mẻ, D Tổng Tích Lũy, E Thời Gian, F Trạng Thái)
    headers_sheet2 = [
        "STT",
        "Mã Mẻ",
        "Số Lượng Mẻ",
        "Tổng Tích Lũy",
        "Thời Gian",
        "Trạng Thái",
    ]
    for col_i, h_text in enumerate(headers_sheet2, start=1):
        c = ws2.cell(row=9, column=col_i, value=h_text)
        c.font = font_header
        c.fill = fill_navy
        c.alignment = align_center
        c.border = thin_border
    ws2.row_dimensions[9].height = 26

    if not batch_records:
        ws2.cell(row=10, column=1, value=1).alignment = align_center
        ws2.cell(row=10, column=2, value="Mẻ #1").alignment = align_center
        c_cnt = ws2.cell(row=10, column=3, value=0)
        c_cnt.number_format = "#,##0"
        c_cnt.alignment = align_right
        c_cum = ws2.cell(row=10, column=4, value="=SUM($C$10:C10)")
        c_cum.number_format = "#,##0"
        c_cum.alignment = align_right
        ws2.cell(row=10, column=5, value="--:--:--").alignment = align_center
        ws2.cell(row=10, column=6, value="Đã cộng").alignment = align_center
        for col_i in range(1, 7):
            cell_item = ws2.cell(row=10, column=col_i)
            cell_item.font = font_regular
            cell_item.fill = fill_white
            cell_item.border = thin_border
        ws2.row_dimensions[10].height = 21
    else:
        for idx, rec in enumerate(batch_records):
            r_idx = detail_start_2 + idx
            row_fill = fill_zebra if idx % 2 == 1 else fill_white

            stt_val = rec.get("stt", rec.get("id", idx + 1))
            batch_code = str(rec.get("batch_name", rec.get("batchCode", f"Mẻ #{idx + 1}")))
            count_val = int(rec.get("count", 0))
            time_val = str(rec.get("time", ""))
            status_val = str(rec.get("status", "Đã cộng"))

            c_stt = ws2.cell(row=r_idx, column=1, value=stt_val)
            c_code = ws2.cell(row=r_idx, column=2, value=batch_code)
            c_cnt = ws2.cell(row=r_idx, column=3, value=count_val)
            # Cột D (Tổng Tích Lũy): Dùng công thức cộng dồn tự động =SUM($C$10:C{r_idx})
            c_cum = ws2.cell(row=r_idx, column=4, value=f"=SUM($C${detail_start_2}:C{r_idx})")
            c_time = ws2.cell(row=r_idx, column=5, value=time_val)
            c_status = ws2.cell(row=r_idx, column=6, value=status_val)

            c_stt.alignment = align_center
            c_code.alignment = align_center
            c_cnt.alignment = align_right
            c_cnt.number_format = "#,##0"
            c_cum.alignment = align_right
            c_cum.number_format = "#,##0"
            c_time.alignment = align_center
            c_status.alignment = align_center

            is_undone = "hoàn tác" in status_val.lower() or "undo" in status_val.lower()
            for col_i in range(1, 7):
                cell_item = ws2.cell(row=r_idx, column=col_i)
                if is_undone:
                    cell_item.font = font_strike
                elif col_i == 4:
                    cell_item.font = font_bold
                else:
                    cell_item.font = font_regular
                cell_item.fill = row_fill
                cell_item.border = thin_border
            ws2.row_dimensions[r_idx].height = 21

    # Dòng Tổng cộng & Trung bình mẻ cuối bảng Sheet 2 (Dùng công thức SUM & AVERAGE)
    footer_row_2 = detail_end_2 + 1
    ws2.merge_cells(start_row=footer_row_2, start_column=1, end_row=footer_row_2, end_column=2)
    ws2.cell(row=footer_row_2, column=1, value="TỔNG CỘNG / TRUNG BÌNH MẺ").alignment = align_center

    c_total_2 = ws2.cell(row=footer_row_2, column=3, value=f"=SUM(C{detail_start_2}:C{detail_end_2})")
    c_total_2.number_format = "#,##0"
    c_total_2.alignment = align_right

    c_avg_2 = ws2.cell(row=footer_row_2, column=4, value=f"=IFERROR(AVERAGE(C{detail_start_2}:C{detail_end_2}), 0)")
    c_avg_2.number_format = '"TB: "#,##0.0" con/mẻ"'
    c_avg_2.alignment = align_right

    ws2.cell(row=footer_row_2, column=5, value=f'=COUNT(C{detail_start_2}:C{detail_end_2})&" mẻ"').alignment = align_center
    ws2.cell(row=footer_row_2, column=6, value="Hoàn tất").alignment = align_center

    for col_i in range(1, 7):
        cell_item = ws2.cell(row=footer_row_2, column=col_i)
        cell_item.font = font_Formula_navy
        cell_item.fill = fill_zebra
        cell_item.border = thin_border
    ws2.row_dimensions[footer_row_2].height = 24

    _auto_fit_columns(ws2, min_width=16)

    return wb


def save_shrimp_report(
    file_path: str,
    pond_name: str,
    staff_name: str,
    weight_records: list,
    batch_records: list,
    total_batch_count: int = 0,
):
    """
    Lưu trực tiếp báo cáo Excel (.xlsx) ra đường dẫn file_path trên máy tính.
    Trả về (success: bool, final_path: str, message: str).
    """
    try:
        wb = build_shrimp_report_workbook(
            pond_name=pond_name,
            staff_name=staff_name,
            weight_records=weight_records,
            batch_records=batch_records,
            total_batch_count=total_batch_count,
        )
        wb.save(file_path)
        return True, file_path, "Đã lưu báo cáo Excel 2 Sheet thành công."
    except PermissionError:
        alt_name = os.path.splitext(file_path)[0] + f"_{int(datetime.datetime.now().timestamp())}.xlsx"
        try:
            wb = build_shrimp_report_workbook(
                pond_name=pond_name,
                staff_name=staff_name,
                weight_records=weight_records,
                batch_records=batch_records,
                total_batch_count=total_batch_count,
            )
            wb.save(alt_name)
            return True, alt_name, "File gốc đang mở, đã lưu sang bản sao mới."
        except Exception as e2:
            return False, file_path, f"Không thể ghi file Excel (Đang mở bởi phần mềm khác): {e2}"
    except Exception as e:
        return False, file_path, f"Lỗi khi xuất file Excel: {e}"


def export_shrimp_excel_bytes(
    pond_name: str,
    staff_name: str,
    weight_records: list,
    batch_records: list,
) -> bytes:
    """
    Xuất báo cáo Excel (.xlsx) dưới dạng bộ nhớ đệm BytesIO để tích hợp trực tiếp
    với nút `st.download_button` của Streamlit.
    """
    wb = build_shrimp_report_workbook(
        pond_name=pond_name,
        staff_name=staff_name,
        weight_records=weight_records,
        batch_records=batch_records,
    )
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
