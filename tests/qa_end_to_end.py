from pathlib import Path
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.c12_tool.parser import parse_c12, parse_c12_records, _extract_from_tables, _key
from src.c12_tool.excel_writer import export_records


def make_pdf(path: Path, code: str, month: int, year: int, offset: int, blank_c: bool = False):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    data = [
        ["STT", "NOI DUNG", "BHXH", "BHYT", "BHTN", "BHTNLD", "CONG"],
        ["A", "B", "1", "2", "3", "4", "5=1+2+3+4"],
        ["A", "Ky truoc mang sang", "10", "20", "30", "40", str(1000 + offset)],
        ["B", "Phat sinh trong ky", "11", "21", "31", "41", str(2000 + offset)],
        ["2", "Quy luong", "12", "22", "42", str(3000 + offset), ""],
        ["C", "So tien da nop trong ky", "13", "23", "33", "43", "" if blank_c else str(4000 + offset)],
    ]
    table = Table(data, colWidths=[45, 72, 68, 68, 72, 68, 85])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    doc.build([
        Paragraph("THONG BAO KET QUA DONG BHXH C12", styles["Heading2"]),
        Paragraph(f"Ma don vi: {code}", styles["Normal"]),
        Paragraph(f"Thang {month} nam {year}", styles["Normal"]),
        Spacer(1, 18),
        table,
    ])


def make_multi_pdf(path: Path):
    styles = getSampleStyleSheet()
    story = []
    for index, (code, month, offset) in enumerate((("YN123456", 1, 10), ("YN123456", 2, 20))):
        data = [
            ["STT", "NOI DUNG", "BHXH", "BHYT", "BHTN", "BHTNLD", "CONG"],
            ["A", "B", "1", "2", "3", "4", "5=1+2+3+4"],
            ["A", "Ky truoc mang sang", "10", "20", "30", "40", str(1000 + offset)],
            ["B", "Phat sinh trong ky", "11", "21", "31", "41", str(2000 + offset)],
            ["2", "Quy luong", "12", "22", "42", str(3000 + offset), ""],
            ["C", "So tien da nop trong ky", "13", "23", "33", "43", str(4000 + offset)],
        ]
        table = Table(data, colWidths=[45, 72, 68, 68, 72, 68, 85])
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ]))
        story.extend([
            Paragraph("THONG BAO KET QUA DONG BHXH C12", styles["Heading2"]),
            Paragraph(f"Ma don vi: {code}", styles["Normal"]),
            Paragraph(f"Thang {month} nam 2026", styles["Normal"]), Spacer(1, 18), table,
        ])
        if index == 0:
            story.append(PageBreak())
    SimpleDocTemplate(str(path), pagesize=A4).build(story)


def main():
    assert _key("Mã đơn vị - BHTNLĐ") == "MA DON VI - BHTNLD"
    real_layout_table = [[
        "STT", "NỘI DUNG", "BHXH", "BHYT", "BHTN", "BHTNLĐ", "CỘNG"
    ], [
        "A", "B", "1", "2", "3", "4", "5 = 1+2+3+4"
    ], [
        "A", "Kỳ trước mang sang", "1.344.786", "", "", "", "1.344.786"
    ], [
        "B", "Phát sinh trong kỳ", "56.725.000", "10.210.500", "4.538.000", "1.134.500", "72.608.000"
    ], [
        "2", "Quỹ lương", "226.900.000", "226.900.000", "226.900.000", "226.900.000", ""
    ], [
        "C", "Tổng", "1", "2", "3", "4", "99.999.999"
    ]]
    layout_values = _extract_from_tables([real_layout_table])
    assert layout_values["2"] == 226900000
    assert layout_values["B"] == 72608000
    main_without_c = real_layout_table[:-1]
    continuation_table = [[
        "C", "Số tiền đã nộp trong kỳ", "", "", "", "", "72.608.000"
    ], [
        "D", "Phân bổ tiền đóng", "56.725.000", "10.210.500", "4.538.000", "1.134.500", "72.608.000"
    ]]
    continuation_values = _extract_from_tables([main_without_c, continuation_table])
    assert continuation_values["C"] == 72608000

    qa = ROOT / "tmp" / "qa"
    qa.mkdir(parents=True, exist_ok=True)
    yn = qa / "C12_YN_01.pdf"
    ic = qa / "C12_IC_03.pdf"
    bw = qa / "C12_BW_01.pdf"
    make_pdf(yn, "YN123456", 1, 2026, 10)
    make_pdf(ic, "IC987654", 3, 2026, 20)
    make_pdf(bw, "BW0001D", 1, 2026, 40)
    records = [parse_c12(yn), parse_c12(ic), parse_c12(bw)]
    assert records[0].unit == "YN" and records[0].month == 1
    assert records[0].b_total == 2010 and records[0].row2_bhtnld == 3010
    assert records[0].c_total == 4010 and records[0].a_total == 1010
    assert records[1].unit == "IC" and records[1].month == 3
    assert records[2].unit == "IC" and records[2].unit_code == "BW0001D"
    partial_pdf = qa / "C12_YN_11_partial.pdf"
    make_pdf(partial_pdf, "YN123456", 11, 2026, 30, blank_c=True)
    partial_record = parse_c12(partial_pdf)
    assert partial_record.month == 11 and partial_record.c_total is None
    multi = qa / "C12_YN_multi.pdf"
    make_multi_pdf(multi)
    multi_records = parse_c12_records(multi)
    assert [record.month for record in multi_records] == [1, 2]
    assert [record.row2_bhtnld for record in multi_records] == [3010, 3020]
    output = export_records(ROOT / "THÔNG BÁO BXH.xlsx", records, qa / "output")
    wb = load_workbook(output, data_only=False)
    assert wb["YN"]["B5"].value == 2010
    assert wb["YN"]["C5"].value == 3010
    assert wb["YN"]["D5"].value == 4010
    assert wb["YN"]["E4"].value == 1010
    assert wb["IC"]["B7"].value == 2020
    assert wb["IC"]["C7"].value == 3020
    assert wb["IC"]["D7"].value == 4020
    assert wb["IC"]["B5"].value == 2040
    assert wb["IC"]["C5"].value == 3040
    assert wb["IC"]["D5"].value == 4040
    assert wb["IC"]["E4"].value == 1040
    assert "KIỂM TRA PDF" in wb.sheetnames
    check = wb["KIỂM TRA PDF"]
    assert check.max_row == 25
    assert check["C2"].value == "Đã nhập"
    assert check["C3"].value == "THIẾU PDF"
    assert check["C16"].value == "Đã nhập"
    assert check["E14"].value == "BW0001D"
    print(output)


if __name__ == "__main__":
    main()
