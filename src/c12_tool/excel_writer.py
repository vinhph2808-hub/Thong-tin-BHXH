from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .parser import C12Record


class ExcelWriteError(ValueError):
    pass


def _display_number(value: float) -> int | float:
    return int(value) if value.is_integer() else value


def _apply_sheet_styling(ws) -> None:
    """Ensure YN and IC sheets have clean, executive-level corporate styling."""
    font_name = "Segoe UI"
    navy_fill = PatternFill("solid", fgColor="1F4E78")
    gray_fill = PatternFill("solid", fgColor="F2F4F8")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    alt_fill = PatternFill("solid", fgColor="F8FAFC")
    dk_fill = PatternFill("solid", fgColor="F9FAFB")
    ck_fill = PatternFill("solid", fgColor="F0F4F8")
    tot_fill = PatternFill("solid", fgColor="E9EEF4")

    thin_gray = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    tot_border = Border(
        left=thin_gray,
        right=thin_gray,
        top=Side(style="thin", color="475569"),
        bottom=Side(style="double", color="475569"),
    )

    num_fmt_standard = '#,##0;(#,##0);"-"'
    num_fmt_balance = '#,##0;[Red](#,##0);"-"'

    try:
        ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["E"].width = 25

    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 18
    for r in (1, 2):
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.fill = navy_fill
            cell.font = Font(name=font_name, size=10.5, bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border

    ws.row_dimensions[3].height = 22
    for c in range(1, 6):
        cell = ws.cell(3, c)
        cell.fill = gray_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = cell_border
        if c == 1:
            cell.font = Font(name=font_name, size=10, bold=True, color="475569")
        elif c in (2, 3, 4):
            cell.font = Font(name=font_name, size=10.5, bold=True, color="C00000")
        else:
            cell.font = Font(name=font_name, size=10, bold=True, color="475569")

    ws.row_dimensions[4].height = 22
    for c in range(1, 6):
        cell = ws.cell(4, c)
        cell.fill = dk_fill
        cell.border = cell_border
        if c == 4:
            cell.font = Font(name=font_name, size=10, italic=True, bold=True, color="475569")
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif c == 5:
            cell.font = Font(name=font_name, size=10.5, bold=True, color="0F172A")
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = num_fmt_standard
        else:
            cell.font = Font(name=font_name, size=10, color="475569")
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for m in range(1, 13):
        r = m + 4
        ws.row_dimensions[r].height = 21
        row_fill = alt_fill if m % 2 == 0 else white_fill
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.fill = row_fill
            cell.border = cell_border
            if c == 1:
                cell.font = Font(name=font_name, size=10, bold=True, color="1E293B")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in (2, 3, 4):
                cell.font = Font(name=font_name, size=10, color="0F172A")
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = num_fmt_standard
            elif c == 5:
                cell.font = Font(name=font_name, size=10, bold=True, color="0F172A")
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = num_fmt_balance

    ws.row_dimensions[17].height = 23
    for c in range(1, 6):
        cell = ws.cell(17, c)
        cell.fill = ck_fill
        cell.border = cell_border
        if c == 4:
            cell.font = Font(name=font_name, size=10, italic=True, bold=True, color="1F4E78")
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif c == 5:
            cell.font = Font(name=font_name, size=10.5, bold=True, color="1F4E78")
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = num_fmt_balance
        else:
            cell.font = Font(name=font_name, size=10, color="475569")
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[18].height = 26
    if not ws.cell(18, 3).value:
        ws.cell(18, 3).value = "=SUM(C5:C16)"
    for c in range(1, 6):
        cell = ws.cell(18, c)
        cell.fill = tot_fill
        cell.border = tot_border
        if c == 1:
            cell.font = Font(name=font_name, size=10.5, bold=True, color="0F172A")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c in (2, 3, 4):
            cell.font = Font(name=font_name, size=10.5, bold=True, color="0F172A")
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = num_fmt_standard
        else:
            cell.font = Font(name=font_name, size=10, color="475569")


def export_records(template_path: str | Path, records: Iterable[C12Record], output_dir: str | Path) -> Path:
    template = Path(template_path)
    output_folder = Path(output_dir)
    output_folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_folder / f"THONG_BAO_BHXH_KET_QUA_{stamp}.xlsx"
    shutil.copy2(template, output_path)

    try:
        workbook = load_workbook(output_path)
    except Exception as exc:
        output_path.unlink(missing_ok=True)
        raise ExcelWriteError(f"Không mở được file Excel mẫu: {exc}") from exc

    sheet_map: dict[str, str] = {}
    for name in workbook.sheetnames:
        norm = name.strip().upper()
        if norm == "YN":
            sheet_map["YN"] = name
        elif norm in ("IC", "IN"):
            sheet_map["IC"] = name
            sheet_map["IN"] = name

    for required in ("YN", "IC"):
        if required not in sheet_map:
            raise ExcelWriteError(f"File Excel thiếu sheet '{required}'.")

    deduped: dict[tuple[str, int], C12Record] = {}
    for record in records:
        key = (record.unit, record.month)
        if key not in deduped:
            deduped[key] = record
        else:
            old = deduped[key]
            deduped[key] = C12Record(
                pdf_path=record.pdf_path if record.b_total is not None else old.pdf_path,
                unit=old.unit,
                unit_code=record.unit_code or old.unit_code,
                month=old.month,
                year=record.year or old.year,
                b_total=record.b_total if record.b_total is not None else old.b_total,
                row2_bhtnld=record.row2_bhtnld if record.row2_bhtnld is not None else old.row2_bhtnld,
                c_total=record.c_total if record.c_total is not None else old.c_total,
                a_total=record.a_total if record.a_total is not None else old.a_total,
            )

    for (unit, month), record in sorted(deduped.items()):
        sheet_name = sheet_map.get(unit, unit)
        if sheet_name not in workbook.sheetnames:
            continue
        ws = workbook[sheet_name]
        row = month + 4
        if record.b_total is not None:
            ws.cell(row=row, column=2, value=_display_number(record.b_total))
        if record.row2_bhtnld is not None:
            ws.cell(row=row, column=3, value=_display_number(record.row2_bhtnld))
        if record.c_total is not None:
            ws.cell(row=row, column=4, value=_display_number(record.c_total))
        # Dòng E4: SDĐK mang sang tương ứng dòng STT A cột CỘNG và CHỈ điền của Tháng 1
        if month == 1:
            val_a = record.a_total if record.a_total is not None else 0.0
            ws["E4"] = _display_number(val_a)

    # Style YN and IC/IN sheets for pristine consistency
    styled_sheets = set()
    for unit in ("YN", "IC"):
        if unit in sheet_map:
            s_name = sheet_map[unit]
            if s_name not in styled_sheets:
                _apply_sheet_styling(workbook[s_name])
                styled_sheets.add(s_name)

    # Recreate KIỂM TRA PDF sheet with pristine executive styling
    if "KIỂM TRA PDF" in workbook.sheetnames:
        del workbook["KIỂM TRA PDF"]
    check = workbook.create_sheet("KIỂM TRA PDF")
    try:
        check.sheet_view.showGridLines = True
    except Exception:
        pass

    headers = ["Đơn vị", "Tháng", "Trạng thái", "Năm", "Mã đơn vị", "Tên file PDF"]
    check.append(headers)

    navy = PatternFill("solid", fgColor="1F4E78")
    white_font = Font(name="Segoe UI", size=10.5, color="FFFFFF", bold=True)
    header_border = Border(
        left=Side(style="thin", color="17365D"),
        right=Side(style="thin", color="17365D"),
        top=Side(style="thin", color="17365D"),
        bottom=Side(style="medium", color="17365D"),
    )

    check.row_dimensions[1].height = 26
    for cell in check[1]:
        cell.fill = navy
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border

    # Fills & Fonts for status
    green_fill = PatternFill("solid", fgColor="E2F0D9")
    green_font = Font(name="Segoe UI", size=10, bold=True, color="1E7E34")

    rose_fill = PatternFill("solid", fgColor="FCE4D6")
    rose_font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

    amber_fill = PatternFill("solid", fgColor="FFF2CC")
    amber_font = Font(name="Segoe UI", size=10, bold=True, color="B25E00")

    thin_gray = Side(style="thin", color="D9E2F3")
    data_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    row_no = 2
    for unit in ("YN", "IC"):
        for month in range(1, 13):
            record = deduped.get((unit, month))
            missing_fields: list[str] = []
            if record:
                if record.b_total is None:
                    missing_fields.append("B")
                if record.row2_bhtnld is None:
                    missing_fields.append("2")
                if record.c_total is None:
                    # Nếu C không có số liệu thanh toán trong kỳ, mặc định là 0 VNĐ đã nộp
                    record = C12Record(
                        pdf_path=record.pdf_path,
                        unit=record.unit,
                        unit_code=record.unit_code,
                        month=record.month,
                        year=record.year,
                        b_total=record.b_total,
                        row2_bhtnld=record.row2_bhtnld,
                        c_total=0.0,
                        a_total=record.a_total,
                    )
                    deduped[(unit, month)] = record
                if month == 1 and record.a_total is None:
                    record = C12Record(
                        pdf_path=record.pdf_path,
                        unit=record.unit,
                        unit_code=record.unit_code,
                        month=record.month,
                        year=record.year,
                        b_total=record.b_total,
                        row2_bhtnld=record.row2_bhtnld,
                        c_total=record.c_total,
                        a_total=0.0,
                    )
                    deduped[(unit, month)] = record

            status = "THIẾU PDF"
            if record and missing_fields:
                status = "Thiếu chỉ tiêu: " + ", ".join(missing_fields)
            elif record:
                status = "Đã nhập"

            values = [
                unit,
                month,
                status,
                record.year if record else None,
                record.unit_code if record else None,
                record.pdf_path.name if record else None,
            ]
            check.append(values)
            check.row_dimensions[row_no].height = 22

            # Status styling
            if record and not missing_fields:
                current_fill = green_fill
                status_font = green_font
            elif record and missing_fields:
                current_fill = amber_fill
                status_font = amber_font
            else:
                current_fill = rose_fill
                status_font = rose_font

            for c_idx, cell in enumerate(check[row_no], 1):
                cell.border = data_border
                cell.fill = current_fill
                if c_idx in (1, 2):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="1E293B")
                elif c_idx == 3:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = status_font
                elif c_idx == 4:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Segoe UI", size=10, color="475569")
                elif c_idx == 5:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
                elif c_idx == 6:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.font = Font(name="Segoe UI", size=10, color="334155")

            row_no += 1

    check.freeze_panes = "A2"
    check.auto_filter.ref = f"A1:F{check.max_row}"
    widths = {"A": 12, "B": 10, "C": 26, "D": 12, "E": 20, "F": 52}
    for col, width in widths.items():
        check.column_dimensions[col].width = width

    check.conditional_formatting.add(
        f"A2:F{check.max_row}",
        FormulaRule(formula=['$C2="THIẾU PDF"'], fill=rose_fill),
    )
    check.sheet_properties.pageSetUpPr.fitToPage = True
    check.page_setup.fitToWidth = 1
    check.page_setup.fitToHeight = 0

    try:
        workbook.save(output_path)
    except PermissionError as exc:
        raise ExcelWriteError("Không thể lưu Excel. Hãy đóng file nếu đang mở rồi thử lại.") from exc
    return output_path
