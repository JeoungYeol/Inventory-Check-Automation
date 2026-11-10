# ============================================================
# 재고조사 설정 샘플 파일
# 실제 사용 시 이 파일을 config.py로 복사하여 수정하세요.
# ============================================================

import os

# 조사 기간
YEAR = "2026"
QUARTER = "1분기"

# 재고조사 루트 경로
# 실제 사용하는 PC의 저장 경로로 변경하세요.
ROOT_PATH = os.path.join(os.path.expanduser("~"), "InventoryData")


# 제품군별 Board List 파일 경로 설정
def get_board_list_path(product_group: str, filename: str) -> str:
    """board_list 폴더 기준 경로 자동 생성"""
    return os.path.join(
        ROOT_PATH,
        YEAR,
        QUARTER,
        product_group,
        "board_list",
        filename
    )


# 공개용 예시 제품군
PRODUCT_GROUPS = {
    "PRODUCT-A": get_board_list_path("PRODUCT-A", "board_a.xls"),
    "PRODUCT-B": get_board_list_path("PRODUCT-B", "board_b.xls"),
}


# 재고 상태 분류 기준
STATUS_IN_STOCK = ["입고", "테스트", "위탁재고"]
STATUS_SHIPPED = ["출고"]
