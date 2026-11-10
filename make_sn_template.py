"""
실물 SN 입력 템플릿 생성 스크립트
────────────────────────────────────────────────────────────
사용법:
  python make_sn_template.py

동작:
  Board List XLS를 읽어서 용량별 시트가 만들어진
  실물SN_{제품군}.xlsx 템플릿 파일을 생성합니다.

  각 시트에 바코드 스캐너로 SN을 직접 입력하면 됩니다.
────────────────────────────────────────────────────────────
"""

import os
import sys
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(__file__))
from config import YEAR, QUARTER, ROOT_PATH, PRODUCT_GROUPS


def parse_xls_html(xls_path: str) -> pd.DataFrame:
    """Support Site XLS(HTML 위장) 파일을 DataFrame으로 반환"""
    with open(xls_path, "r", encoding="utf-8", errors="replace") as f:
        html = f.read()
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    rows = table.find_all("tr")
    headers = [td.get_text(strip=True) for td in rows[0].find_all("td")]
    data = []
    for row in rows[1:]:
        cells = [td.get_text(strip=True) for td in row.find_all("td")]
        if cells:
            data.append(cells)
    return pd.DataFrame(data, columns=headers)


def read_board_list_sizes(xls_path: str) -> list:
    """Board List에서 용량 목록만 추출"""
    df = parse_xls_html(xls_path)
    return sorted(df["Size"].dropna().unique().tolist())


def make_template(product_group: str, board_list_path: str):
    if not os.path.exists(board_list_path):
        print(f"  [스킵] Board List 없음: {board_list_path}")
        return

    sizes = read_board_list_sizes(board_list_path)
    print(f"  {product_group}: 용량 {sizes}")

    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill("solid", fgColor="2D4A8A")
    input_fill  = PatternFill("solid", fgColor="FFFDE7")
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for size in sizes:
        ws = wb.create_sheet(size)

        ws["A1"] = f"[{product_group} — {size}] 실물 SN 입력 시트"
        ws["A1"].font = Font(bold=True, size=11, color="2D4A8A")
        ws["A2"] = "바코드 스캐너 또는 수동으로 SN을 A열에 입력하세요. 중복 입력 무방합니다."
        ws["A2"].font = Font(size=9, color="888888")
        ws.append([])

        ws["A4"] = "SN"
        ws["B4"] = "비고 (선택)"
        for cell in [ws["A4"], ws["B4"]]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

        for i in range(5, 205):
            ws[f"A{i}"].fill = input_fill
            ws[f"A{i}"].border = border
            ws[f"B{i}"].border = border

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 20
        ws.freeze_panes = "A5"

    out_dir = os.path.join(ROOT_PATH, YEAR, QUARTER, product_group)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"실물SN_{product_group}.xlsx")
    wb.save(out_path)
    print(f"    → 템플릿 저장: {out_path}")


def main():
    print(f"\n실물 SN 입력 템플릿 생성 — {YEAR} {QUARTER}")
    print("=" * 50)
    for group, board_path in PRODUCT_GROUPS.items():
        make_template(group, board_path)
    print("\n완료! 각 제품군 폴더에 템플릿 파일이 생성되었습니다.")
    print("바코드 스캐너로 SN을 입력한 후 inventory_check.py를 실행하세요.\n")


if __name__ == "__main__":
    main()
