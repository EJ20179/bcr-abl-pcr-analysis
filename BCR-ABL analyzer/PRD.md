# PRD — BCR-ABL RQ-PCR 분석 시스템

**문서 버전:** 1.3  
**작성일:** 2026-04-30  
**최종 수정:** 2026-06-08  
**작성자:** EJ20179  
**배포 URL:** https://bcr-abl-pcr-analysis.vercel.app  
**상태:** 확정

---

## 1. 제품 개요

BCR-ABL RQ-PCR 분석 시스템은 서버 없이 브라우저에서 단독 실행되는 **단일 HTML 파일** 웹 애플리케이션입니다.  
qPCR 원시 데이터를 업로드하면 BCR-ABL/ABL 정량 분석, QC 판정, IS 보정, Excel 보고서 출력까지 자동 수행합니다.

**기술 스택**
- Vanilla JavaScript (프레임워크 없음)
- xlsx-js-style 1.2.0 — Excel 출력 및 셀 스타일
- JSZip 3.10.1 — xlsx ZIP 조작 (이미지 삽입)
- Chart.js 4.4.1 — 표준곡선 시각화
- Tesseract.js — OCR (표준곡선 파라미터·Previous study 결과지 인식)

---

## 2. 기능 목록

### 2.1 파일 업로드 및 파싱

| ID | 기능 | 설명 |
|----|------|------|
| F-01 | CSV/Excel 업로드 | QuantStudio 등 qPCR 기기 출력 파일 지원 |
| F-02 | 타깃 자동 식별 | BCR-ABL, ABL 두 타깃 자동 구분 |
| F-03 | 샘플 분류 | 표준(SP1~SP6), 환자, 컨트롤(PC1/PC2/HC/IS-CAL/H2O) 자동 분류 |
| F-04 | 메타데이터 추출 | 실험일, 실험자, 기기 정보 자동 추출 |

### 2.2 계산 및 분석

| ID | 기능 | 설명 |
|----|------|------|
| F-10 | 표준곡선 계산 | 기울기(Slope), 절편(Y-Intercept), R², 효율(Eff%) 계산 |
| F-11 | Copy number 계산 | Quantity Mean 기반 BCR-ABL/ABL copy 수 계산 |
| F-12 | NCN 계산 | NCN(%) = (BCR-ABL / ABL) × 100 |
| F-13 | IS-NCN 계산 | IS-NCN = NCN × CF (보정계수 적용) |
| F-14 | CF 자동 계산 | IS-CAL NCN 기반 CF = IS-Cal Value / IS-CAL NCN |
| F-15 | 분자 반응 판정 | Complete / Major / No Major Molecular Response 자동 분류 |
| F-16 | AMR 판정 | IS-NCN이 정량 신뢰 범위(AMR) 내인지 표시 |

### 2.3 QC 자동 판정 (9개 항목)

| ID | QC 항목 | 기준 |
|----|---------|------|
| Q-01 | CT Variations | CT SD ≤ 2 (평균 Ct < 기준값) / ≤ 1.5 (≥ 기준값) |
| Q-02 | Slope | 설정 범위 내 (예: −3.1 ~ −3.9) |
| Q-03 | R² | ≥ 설정값 (예: > 0.990) |
| Q-04 | SP1 검출 | SP1 표준품 검출 및 표준곡선 포함 여부 |
| Q-05 | ABL CN | ABL copy number > 설정 하한 |
| Q-06 | Negative controls | PCR water / RT negative = 0 |
| Q-07 | IS-CAL NCN | IS-CAL NCN이 허용 범위 내 |
| Q-08 | HC 검출 | High positive control 검출 여부 |
| Q-09 | HC IS-NCN | HC IS-NCN ≥ 설정 하한 |

### 2.4 설정 관리

| ID | 설정 항목 | 기본값 |
|----|-----------|--------|
| S-01 | IS-Cal Value (CF 분자) | 0.101 |
| S-02 | AMR 최솟값 | 0.0023 |
| S-03 | AMR 최댓값 | 42.36 |
| S-04 | QC 기준값 (Slope, R², CT SD 임계 등) | 가이드라인 기본값 |
| S-05 | IS-CAL NCN 허용 범위 | 설정 가능 |
| S-06 | MMR / CMR 판정 기준 | 0.1% / 0.01% |

### 2.5 Excel 출력

#### 시트 구성

| 시트명 | 내용 |
|--------|------|
| 검사 | 검사 양식 (BCR-ABL 좌 / ABL 우 / QC·환자결과 하단) |
| 검사결과 | BCR-ABL RQ-PCR 검사 결과 보고서 |
| 환자결과 | 환자별 상세 결과 |
| QC결과 | QC 항목별 판정 결과 |
| 컨트롤 | 컨트롤 샘플 결과 |
| 표 값 | 숫자 정밀도 보존 결과 테이블 |
| work list 값 | LIS 복사용 2열 워크리스트 |
| Raw | 원시 데이터 전체 |

#### 검사 시트 세부 레이아웃

```
Cols A~L  : BCR-ABL 영역
  - SampleNo. / Sample Name / Target Name / Reporter / Quencher
  - Ct / Ct Mean / Ct SD / Quantity / Quantity Mean / Quantity SD
  - 표준곡선 이미지 (데이터 하단, JSZip으로 삽입)

Col M     : 갭

Cols N~AB : ABL 영역
  - SampleNo. / Sample Name / ... (동일 구조)
  - NCN (col Z, 소수점 3자리)
  - NCN(Quanti Mean) (col AA, 소수점 4자리)
  - IS-NCN (col AB, 소수점 4자리)

하단 섹션 (ABL 영역 내):
  - IS-Cal Value 행: W~Y 병합 레이블 + Z~AB 노란 배경 값
  - QC 테이블: Criteria / Acceptable values / Result / Detail (4열)
  - AMR IS-NCN 행: Y 레이블 + Z~AB 빨간 굵은 값
  - 환자 결과 테이블: No. / Sample / BCR-ABL1 copy / ABL1 copy / NCN / IS-NCN / Molecular
    - MR 범례 인라인 (첫 3행: No Major Molecular Response / Major Molecular Response / Complete Molecular Response 범위)
```

#### 셀 스타일 규격

| 항목 | 스타일 |
|------|--------|
| NCN/IS-NCN 헤더 | 분홍 배경 (#DDA0DD), 볼드 |
| Undetermined 행 (BCR-ABL) | 노란 배경 (#FFFF99) |
| PC1/PC2 첫 replicate | 노란 배경 |
| IS-CAL NCN / HC IS-NCN 값 | 노란 배경 |
| QC Detail 셀 | 노란 배경, 중앙 정렬 |
| AMR IS-NCN 값 | 노란 배경, 빨간 굵은 글씨 |
| 환자 헤더 | 회색 배경 (#D9D9D9), 볼드 |
| QC pass | 녹색 글씨 |
| QC FAIL | 빨간 글씨, 볼드 |

### 2.6 표준곡선 이미지 삽입

- Chart.js로 렌더링된 표준곡선 이미지를 Excel 검사 시트에 삽입
- JSZip으로 xlsx ZIP 파일을 직접 조작하여 `xl/drawings/drawing1.xml` 주입
- 검사 시트 BCR-ABL 영역 하단 (imageAnchorRow ~ lastBottomRow)에 위치

### 2.7 소프트웨어 값 직접 입력

- 표준곡선 Slope / Y-Intercept / R² / Eff%를 소프트웨어 계산값 대신 수동 입력 가능
- 수동 입력 시 Excel 검사 시트 표준곡선 텍스트에 반영

### 2.8 CF 이력 관리

- 날짜별 CF 분자, AMR 범위 이력 테이블 조회
- 이력 추가·수정·삭제

### 2.9 LIS Export

- **버튼**: 📤 LIS Export (Excel 다운로드 버튼 왼쪽)
- **출력 파일명**: `LIS_BCR-ABL_YYYYMMDD.xlsx`
- **시트명**: LIS
- **컬럼 구조**: 검체번호 / 검사명 / 결과값 (3열)
- **환자별 5행 고정**:

| 검사명 | 결과값 규칙 |
|--------|------------|
| BCR::ABL1 copy number | 정수 / LOD 미만·0이면 `<3` |
| ABL1 copy number | 정수 |
| NCN | 소수 4자리 / LOD 미만이면 `<0.0026` |
| BCR::ABL1 %IS | 소수 4자리 / LOD 미만이면 `<0.0026` |
| Molecular response | Complete / Major / No Major Molecular Response |

### 2.10 Previous study 보조 입력

Work List 탭의 각 환자 카드에서 직전 보고서 이미지를 붙여넣어 Previous study 텍스트를 자동 생성하는 기능.

| 항목 | 내용 |
|------|------|
| 날짜 입력 | 달력 picker (이전 접수 날짜, `YYYY-MM-DD` → `YYYY.MM.DD` 변환) |
| 날짜 상한 자동 설정 | 검체 ID 앞 6자리(YYMMDD) 파싱 → 해당 날짜를 `max`로 설정, 이후 날짜 선택 불가 |
| 검체 종류 | EDTA WB (기본값) / EDTA BM 토글 |
| 이미지 입력 | 영역 클릭 후 Ctrl+V (클립보드 이미지 붙여넣기) |
| OCR 엔진 | Tesseract.js 재사용 워커, 인식 시 문자 화이트리스트 해제 후 복원 |
| 추출 항목 | BCR::ABL1 %IS (LOD 미만 `<값` 포함), Molecular response |
| MR 약칭 변환 | Complete Molecular Response → CMR, Major → MMR, No Major → No MMR |
| 결과 구성 | 줄1: `(날짜)-{검체} : {%IS} / {약칭}` + 이전 보고서의 Previous study 항목 최대 2줄 |
| 편집 | OCR 결과를 textarea에 자동 입력 후 수동 수정 가능 |
| 개인정보 | 환자 정보 상단 제외, 결과표·Remark 영역만 캡쳐하여 사용 권장 |

**Previous study 출력 형식**
```
(YYYY.MM.DD)-EDTA WB : 0.7108 / No MMR
(2026.02.04)-EDTA WB : 0.3021 / No MMR
(2025.12.12)-EDTA WB : 0.0554 / MMR
```

---

## 3. 비기능 요구사항

| 항목 | 기준 |
|------|------|
| 실행 환경 | 최신 Chrome / Edge (서버 불필요) |
| 파일 크기 | 단일 HTML 파일 (외부 CDN 의존) |
| 보안 | 환자 데이터 외부 전송 없음 (로컬 처리) |
| 성능 | 10샘플 기준 Excel 생성 < 3초 |
| 오프라인 | CDN 없이도 기본 분석 가능 (xlsx 출력 불가) |

---

## 4. 개발 이력 (주요 변경사항)

### v1.0 (2026-04-30)

| 항목 | 내용 |
|------|------|
| Excel 레이아웃 | 검사 시트 하단 3섹션 가로 배치 (QC / 환자결과 / 표준곡선) |
| QC 영문화 | Criteria / Acceptable values / Result / Detail 4열 포맷 |
| 소수점 통일 | NCN 3자리, NCN(QM)/IS-NCN 4자리 고정 (trailing zero 포함) |
| CT Variations | 판정 텍스트 "pass/FAIL" 영문화 |
| R² Detail | 체크마크(✓✗) 제거 |
| AMR IS-NCN | Y열 단독 레이블, Z~AB 병합 값, 빨간 굵은 글씨 |
| Molecular 열 | IS-NCN 바로 옆(Y열)으로 이동 |
| 표준곡선 이미지 | JSZip으로 xlsx에 직접 삽입 |
| PC1/PC2 배경 | 첫 replicate 노란색 배경 추가 |
| NCN 0 처리 | 계산값 없을 때 빈칸 대신 0 표시 |
| SampleNo. | A/N열 헤더 변경, 환자:숫자/표준·컨트롤:이름 |
| Excel 수리 대화상자 | xmlns:r 네임스페이스 중복 버그 수정 |
| Content-Types 버그 | png 확장자 중복 등록 버그 수정 |
| IS-Cal Value 행 | 데이터와 Criteria 사이 추가, W~Y 병합 레이블 |
| 빈 셀 테두리 | 값 없는 셀 테두리 제거 (환자 MR 범례 4행+, AMR 빈 열) |
| MR 범례 인라인 | 환자 결과 첫 3행에 MR 판정 기준 인라인 배치 |

### v1.1 (2026-06-05)

| 항목 | 내용 |
|------|------|
| MR 전체 명칭 표기 | 모든 화면(배지·복사·Excel)에서 CMR → Complete Molecular Response, MMR → Major Molecular Response, No MMR → No Major Molecular Response로 변경 |
| Work List MR 표시 위치 | MR 배지를 카드 헤더에서 %IS 행 아래(MR: 행)로 이동 |
| Work List MR 복사 가능 | MR 배지 클릭 시 전체 명칭 클립보드 복사 지원 |
| 전체 복사 텍스트 | 📋 전체 버튼 복사 내용에 `MR: <전체 명칭>` 줄 추가 |

### v1.2 (2026-06-08)

| 항목 | 내용 |
|------|------|
| LIS Export 추가 | 자체 LIS import용 Excel 파일 출력 기능 추가 (§2.9) |
| LIS 파일 구조 | 검체번호 / 검사명 / 결과값 3열, 환자별 5행 고정 |
| BCR-ABL copy number 0 처리 | 값이 0이거나 LOD 미만이면 `<3` 표기 |

### v1.3 (2026-06-08)

| 항목 | 내용 |
|------|------|
| Previous study 보조 입력 추가 | Work List 카드에 날짜 picker·검체 종류 토글·이미지 붙여넣기+OCR 기능 추가 (§2.10) |
| OCR 파싱 개선 | 줄 단위 탐색으로 전환, LOD 미만 값(`<숫자`) 처리, MR이 다음 줄에 오는 경우 대응 |
| 날짜 상한 자동 설정 | 검체 ID 앞 6자리 YYMMDD 파싱 → 달력 picker max 자동 지정, 이후 날짜 선택 불가 |
| 기술 스택 추가 | Tesseract.js를 공식 의존 라이브러리로 명시 |

---

## 5. 오픈 이슈

| ID | 내용 | 우선순위 |
|----|------|----------|
| OI-01 | 2개 이상 파일 동시 업로드 지원 | Medium |
| OI-02 | PDF 보고서 직접 출력 | Low |
| OI-03 | 모바일 레이아웃 최적화 | Low |
