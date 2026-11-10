"""
재고조사 자동화 스크립트
────────────────────────────────────────────────────────────
사용법:
  1. config.py 에서 연도/분기/파일 경로 설정
  2. 각 제품군 board_list/ 폴더에 Support Site XLS 파일 넣기
  3. 실물 SN 목록 엑셀 준비 (make_sn_template.py 로 생성 가능)
  4. python inventory_check.py 실행

결과:
  재고조사/{YEAR}/{QUARTER}/{제품군}/재고조사_{제품군}.xlsx 자동 생성
────────────────────────────────────────────────────────────
"""

import os
import sys
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from config import YEAR, QUARTER, ROOT_PATH, PRODUCT_GROUPS, STATUS_IN_STOCK, STATUS_SHIPPED


# ── 색상 정의 ────────────────────────────────────────────────
COLOR = {
    "header":    "2D4A8A",
    "ok":        "D6F0D6",
    "missing":   "FFF0C0",
    "ship_leak": "FFD6D6",
    "ghost":     "F0D6FF",
    "summary":   "EEF2FF",
}


# ── Board List 읽기 ──────────────────────────────────────────
def read_board_list(xls_path: str) -> pd.DataFrame:
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
    df = pd.DataFrame(data, columns=headers)
    df["Serial"] = df["Serial"].astype(str).str.strip().str.upper()
    df["Board ID"] = df["Board ID"].astype(str).str.strip().str.upper()
    return df


# ── 실물 SN 목록 읽기 ────────────────────────────────────────
def read_actual_sns(xlsx_path: str, sheet_name: str) -> set:
    """실물 SN 입력 엑셀에서 해당 시트의 SN 집합 반환"""
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=3)
    except Exception as e:
        print(f"    [경고] 실물 SN 시트 '{sheet_name}' 읽기 실패: {e}")
        return set()

    col = None
    for c in df.columns:
        if str(c).strip().upper() in ["SN", "SERIAL", "SERIAL NUMBER", "시리얼"]:
            col = c
            break
    if col is None:
        col = df.columns[0]

    sns = df[col].dropna().astype(str).str.strip().str.upper()
    return set(sns[sns.str.len() > 0])


# ── 케이스 분류 ──────────────────────────────────────────────
def classify(board_df: pd.DataFrame, actual_sns: set) -> dict:
    board_sns = set(board_df["Serial"].str.upper())
    results = {"ok": [], "missing": [], "ship_leak": []}

    for _, row in board_df.iterrows():
        sn = str(row["Serial"]).upper()
        status = str(row.get("Status", "")).strip()
        in_actual = sn in actual_sns

        if status in STATUS_SHIPPED:
            if in_actual:
                results["ship_leak"].append(row)
        else:
            if in_actual:
                results["ok"].append(row)
            else:
                results["missing"].append(row)

    ghost_sns = {str(sn) for sn in (actual_sns - board_sns)
                 if str(sn).upper() not in ("NAN", "NONE", "")}

    return {
        "ok":        pd.DataFrame(results["ok"]),
        "missing":   pd.DataFrame(results["missing"]),
        "ship_leak": pd.DataFrame(results["ship_leak"]),
        "ghost":     sorted(ghost_sns),
    }


# ── 엑셀 스타일 헬퍼 ─────────────────────────────────────────
def _header_style(cell, color_hex: str):
    cell.fill = PatternFill("solid", fgColor=color_hex)
    cell.font = Font(bold=True, color="FFFFFF", size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _cell_border():
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _write_df_to_sheet(ws, df: pd.DataFrame, color_hex: str):
    if df.empty:
        ws.append(["데이터 없음"])
        return
    headers = list(df.columns)
    ws.append(headers)
    for cell in ws[1]:
        _header_style(cell, COLOR["header"])
        cell.border = _cell_border()
    fill = PatternFill("solid", fgColor=color_hex)
    for row_data in df.itertuples(index=False):
        ws.append(list(row_data))
        for cell in ws[ws.max_row]:
            cell.fill = fill
            cell.border = _cell_border()
            cell.alignment = Alignment(vertical="center")
    for col_idx, col_name in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(str(col_name)), 12) + 2


def _write_ghost_to_sheet(ws, ghost_sns: list):
    ws.append(["SN", "비고"])
    for cell in ws[1]:
        _header_style(cell, COLOR["header"])
        cell.border = _cell_border()
    fill = PatternFill("solid", fgColor=COLOR["ghost"])
    for sn in ghost_sns:
        ws.append([sn, "Board List 미등록"])
        for cell in ws[ws.max_row]:
            cell.fill = fill
            cell.border = _cell_border()
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 20


# ── 결과 엑셀 생성 ───────────────────────────────────────────
def write_result_excel(output_path: str, size_results: dict, product_group: str):
    wb = Workbook()
    wb.remove(wb.active)

    # 요약 시트
    ws_summary = wb.create_sheet("📊 요약")
    ws_summary.append([f"{YEAR} {QUARTER} 재고조사 — {product_group}"])
    ws_summary["A1"].font = Font(bold=True, size=13)
    ws_summary.append([f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M')}"])
    ws_summary.append([])
    summary_headers = ["용량/모델", "정상", "미확인", "출고누락", "유령재고", "실물합계"]
    ws_summary.append(summary_headers)
    for cell in ws_summary[4]:
        _header_style(cell, COLOR["header"])
        cell.border = _cell_border()

    summary_fill = PatternFill("solid", fgColor=COLOR["summary"])
    for size, res in size_results.items():
        row = [
            size,
            len(res["ok"]),
            len(res["missing"]),
            len(res["ship_leak"]),
            len(res["ghost"]),
            len(res["ok"]) + len(res["ship_leak"]) + len(res["ghost"]),
        ]
        ws_summary.append(row)
        for cell in ws_summary[ws_summary.max_row]:
            cell.fill = summary_fill
            cell.border = _cell_border()
            cell.alignment = Alignment(horizontal="center")
    for col in ["A", "B", "C", "D", "E", "F"]:
        ws_summary.column_dimensions[col].width = 16

    # 용량별 시트
    for size, res in size_results.items():
        s = size.replace("/", "-")
        ws_ok    = wb.create_sheet(f"✅{s}_정상")
        _write_df_to_sheet(ws_ok, res["ok"], COLOR["ok"])

        ws_miss  = wb.create_sheet(f"📋{s}_미확인")
        _write_df_to_sheet(ws_miss, res["missing"], COLOR["missing"])

        ws_leak  = wb.create_sheet(f"⚠{s}_출고누락")
        _write_df_to_sheet(ws_leak, res["ship_leak"], COLOR["ship_leak"])

        ws_ghost = wb.create_sheet(f"❌{s}_유령재고")
        _write_ghost_to_sheet(ws_ghost, res["ghost"])

    wb.save(output_path)
    print(f"    → 저장 완료: {output_path}")


# ── 메인 실행 ────────────────────────────────────────────────
def run_inventory_check(product_group: str, board_list_path: str, sn_xlsx_path: str):
    print(f"\n{'='*55}")
    print(f"  제품군: {product_group}")
    print(f"{'='*55}")

    if not os.path.exists(board_list_path):
        print(f"  [스킵] Board List 파일 없음: {board_list_path}")
        return

    board_df = read_board_list(board_list_path)
    sizes = sorted(board_df["Size"].dropna().unique())
    print(f"  Board List: {len(board_df)}행, 용량: {sizes}")

    if not os.path.exists(sn_xlsx_path):
        print(f"  [스킵] 실물 SN 파일 없음: {sn_xlsx_path}")
        print(f"         make_sn_template.py 로 템플릿을 먼저 생성하세요.")
        return

    size_results = {}
    for size in sizes:
        print(f"\n  [{size}] 처리 중...")
        df_size = board_df[board_df["Size"] == size].copy()
        actual_sns = read_actual_sns(sn_xlsx_path, sheet_name=size)
        print(f"    Board List: {len(df_size)}개 | 실물 SN: {len(actual_sns)}개")
        result = classify(df_size, actual_sns)
        size_results[size] = result
        print(f"    ✅ 정상:     {len(result['ok'])}개")
        print(f"    📋 미확인:   {len(result['missing'])}개")
        print(f"    ⚠️  출고누락: {len(result['ship_leak'])}개")
        print(f"    ❌ 유령재고: {len(result['ghost'])}개")

    out_dir = os.path.join(ROOT_PATH, YEAR, QUARTER, product_group)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"재고조사_{product_group}.xlsx")
    write_result_excel(out_path, size_results, product_group)


def main():
    print(f"\n{'='*55}")
    print(f"  Novachips 재고조사 자동화 스크립트")
    print(f"  대상: {YEAR} {QUARTER}")
    print(f"{'='*55}")

    for group, board_path in PRODUCT_GROUPS.items():
        sn_path = os.path.join(ROOT_PATH, YEAR, QUARTER, group, f"실물SN_{group}.xlsx")
        run_inventory_check(group, board_path, sn_path)

    print(f"\n{'='*55}")
    print("  전체 완료!")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
