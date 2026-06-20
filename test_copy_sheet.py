"""
테스트: 시재자금현황_2026.xlsx 의 6.18 시트를 복사해 6.19 신규 시트 생성
실제 거래 데이터 없이 시트 복사 + 전일잔액 이월 + 거래내역 표 초기화만 검증
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from daily_cash_report import (
    load_workbook,
    _latest_sheet,
    copy_sheet,
    scan_balance_table,
    carry_forward_balances,
    _clear_transaction_table,
    BAL_TABLE_START_ROW, BAL_TABLE_END_ROW,
    COL_ACCT_NO, COL_PREV_BAL, COL_CURR_BAL,
)

TARGET = Path("시재자금현황_2026.xlsx")
NEW_SHEET = "6.19"

if not TARGET.exists():
    print(f"[오류] {TARGET} 파일이 없습니다. 같은 폴더에 두고 실행하세요.")
    sys.exit(1)

wb = load_workbook(TARGET)
print(f"현재 시트 목록: {wb.sheetnames}")

if NEW_SHEET in wb.sheetnames:
    print(f"'{NEW_SHEET}' 시트가 이미 있습니다 — 삭제 후 재생성합니다.")
    del wb[NEW_SHEET]

prev_ws = _latest_sheet(wb)
print(f"복사 원본 시트: {prev_ws.title}")

new_ws = copy_sheet(wb, prev_ws, NEW_SHEET)
print(f"'{NEW_SHEET}' 시트 생성 완료")

balance_map = scan_balance_table(new_ws)
print(f"인식된 계좌 수: {len(balance_map)}")
for acct, info in list(balance_map.items())[:5]:
    prev = new_ws.cell(row=info['row'], column=COL_CURR_BAL).value
    print(f"  계좌 {acct} | {info['alias']} | 당일잔액(원본)={prev}")

carry_forward_balances(prev_ws, new_ws, balance_map)
print("전일잔액 이월 완료")

# 이월 결과 확인 (첫 5개)
for acct, info in list(balance_map.items())[:5]:
    prev_bal = new_ws.cell(row=info['row'], column=COL_PREV_BAL).value
    curr_bal = new_ws.cell(row=info['row'], column=COL_CURR_BAL).value
    print(f"  계좌 {acct} | 전일잔액={prev_bal} | 당일잔액={curr_bal}")

_clear_transaction_table(new_ws)
print("거래내역 표 초기화 완료")

wb.save(TARGET)
print(f"\n저장 완료: {TARGET}")
print(f"시트 목록: {wb.sheetnames}")
