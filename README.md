# Inventory Check Automation

Python을 활용한 실물 재고조사 및 시리얼 번호 비교 자동화 프로그램

## 1. 프로젝트 소개

### 개발 배경

기존 재고조사 업무에서는 제품별 Board List와 실제 보유 제품의 Serial Number(SN)를 수작업으로 대조하여 재고 현황을 확인해야 했습니다.

제품군과 용량에 따라 확인해야 할 데이터가 많아 반복적인 비교 작업이 발생했으며, 재고 상태를 분류하고 조사 결과를 정리하는 과정에서도 시간이 소요되었습니다.

이러한 반복 업무를 개선하기 위해 Python과 Excel을 연동하여 Board List와 실물 SN 데이터를 자동으로 비교하고, 재고조사 결과를 Excel 보고서로 생성하는 프로그램을 개발했습니다.

### 개발 목적

- 반복적인 시리얼 번호 대조 작업 자동화
- 제품군 및 용량별 재고 현황 분류
- 재고조사 결과의 Excel 보고서 자동 생성
- 수작업으로 발생할 수 있는 누락 및 확인 오류 감소

## 2. 주요 기능

| 기능 | 설명 |
|---|---|
| Board List 데이터 파싱 | HTML 형식으로 제공되는 XLS 파일의 테이블 데이터 추출 |
| 실물 SN 양식 생성 | Board List의 제품 용량을 기준으로 입력용 Excel 파일 자동 생성 |
| SN 데이터 비교 | Board List와 실물 SN 데이터를 비교하여 일치 여부 확인 |
| 재고 상태 분류 | 정상, 미확인, 출고 누락 의심, 미등록 재고로 분류 |
| 결과 보고서 생성 | 조사 결과와 요약 정보를 포함한 Excel 보고서 생성 |

## 3. 전체 동작 과정

```text
       Board List 파일
              │
              ▼
    make_sn_template.py
              │
              ▼
      실물 SN 입력 양식 생성
              │
              ▼
       바코드 스캔 및 입력
              │
              ▼
      inventory_check.py
              │
              ▼
      Board List와 실물 SN 비교
              │
              ▼
         재고 상태 분류
              │
              ▼
       Excel 결과 보고서 생성
```

## 4. 핵심 구현 내용

### ① Excel 및 HTML 데이터 처리

- BeautifulSoup을 활용하여 HTML 형식의 XLS 파일에서 테이블 데이터 추출
- Pandas DataFrame을 활용하여 제품 및 시리얼 번호 데이터 처리
- 제품 용량별 데이터를 구분하여 재고조사에 필요한 입력 양식 생성

### ② 시리얼 번호 비교 로직

- 문자열 앞뒤 공백 제거 및 대문자 변환을 통한 데이터 정규화
- Python Set 자료형을 활용하여 Board List와 실물 SN 데이터 비교
- 데이터의 존재 여부와 출고 상태를 기준으로 재고 유형 분류

### ③ Excel 보고서 자동 생성

- OpenPyXL을 활용하여 조사 결과 Excel 파일 생성
- 재고 상태별 데이터 시트 구성
- 요약 시트에 전체 조사 건수 및 분류별 결과 표시
- 셀 색상, 테두리, 정렬 등 보고서 서식 자동 적용

## 5. 기술 스택

| 구분 | 기술 |
|---|---|
| Language | Python |
| Data Processing | Pandas |
| HTML Parsing | BeautifulSoup4 |
| Excel Processing | OpenPyXL |
| Data Comparison | Python Set |

## 6. 프로젝트 구조

```text
Inventory-Check-Automation/
│
├── config.example.py
├── inventory_check.py
├── make_sn_template.py
├── requirements.txt
├── .gitignore
└── README.md
```

| 파일명 | 역할 |
|---|---|
| config.example.py | 조사 기간, 데이터 경로, 제품군 및 상태 설정 예시 |
| make_sn_template.py | Board List 기반 실물 SN 입력 양식 생성 |
| inventory_check.py | SN 비교, 재고 분류 및 결과 보고서 생성 |
| requirements.txt | 프로젝트 실행에 필요한 라이브러리 목록 |

## 7. 실행 방법

### 1) 저장소 복제

```bash
git clone 저장소_URL
```

### 2) 라이브러리 설치

```bash
pip install -r requirements.txt
```

### 3) 환경 설정

`config.example.py`를 복사하여 `config.py`로 변경합니다.

실제 사용 환경에 맞게 다음 항목을 설정합니다.

- 조사 연도 및 분기
- 재고조사 데이터 저장 경로
- 제품군별 Board List 파일 경로
- 재고 상태 분류 기준

### 
