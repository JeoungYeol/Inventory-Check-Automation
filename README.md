# 재고조사 자동화 프로그램

## 1. 프로젝트 소개

Python을 활용하여 재고조사 업무를 자동화하는 프로그램입니다.

기존에는 제품별 Board List와 실제 보유한 제품의 시리얼 번호(SN)를 확인하고 재고 현황을 수작업으로 비교해야 했습니다.

본 프로그램은 두 데이터를 자동으로 비교하고, 재고 상태를 분류하여 Excel 보고서로 출력합니다.

## 2. 주요 기능

- Board List 데이터 자동 수집
- 실물 SN 엑셀 양식 자동 생성
- Board List와 실물 SN 데이터 비교
- 재고 상태별 자동 분류
- 제품군 및 용량별 Excel 보고서 생성
- 재고 현황 요약 및 결과 시각화

## 3. 재고 분류 기준

| 분류 | 설명 |
|---|---|
| 정상 재고 | Board List와 실물 SN이 일치하는 제품 |
| 미확인 재고 | Board List에는 있지만 실물 SN이 확인되지 않는 제품 |
| 출고 누락 의심 | 출고 처리되었지만 실물 SN이 확인되는 제품 |
| 미등록 재고 | 실물 SN에는 있지만 Board List에는 없는 제품 |

## 4. 개발 환경

- Language: Python
- Data Processing: Pandas
- HTML Parsing: BeautifulSoup4
- Excel Processing: OpenPyXL

## 5. 프로젝트 구조

```text
Inventory-Check-Automation/
├── config.example.py
├── inventory_check.py
├── make_sn_template.py
├── requirements.txt
├── .gitignore
└── README.md
```

## 6. 실행 방법

### 1) 저장소 복제

```bash
git clone 저장소_URL
```

### 2) 라이브러리 설치

```bash
pip install -r requirements.txt
```

### 3) 환경 설정

`config.example.py`를 복사하여 `config.py`로 변경한 뒤, 실제 작업 환경에 맞게 경로와 제품군 정보를 설정합니다.

### 4) 프로그램 실행

실물 SN 입력 양식 생성:

```bash
python make_sn_template.py
```

재고조사 실행:

```bash
python inventory_check.py
```

## 7. 활용 기술

- Python 기반 업무 자동화
- Pandas를 활용한 데이터 처리
- BeautifulSoup을 활용한 HTML 데이터 파싱
- Set 자료형을 활용한 시리얼 번호 비교
- OpenPyXL을 활용한 Excel 문서 생성 및 서식 적용
