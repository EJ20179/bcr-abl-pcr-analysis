# BRD — BCR-ABL RQ-PCR 분석 시스템

**문서 버전:** 1.1  
**작성일:** 2026-04-30  
**최종 수정:** 2026-06-08  
**작성자:** EJ20179  
**상태:** 확정

---

## 1. 배경 및 목적

BCR-ABL 융합 유전자는 만성 골수성 백혈병(CML) 및 일부 급성 림프구성 백혈병(ALL)의 주요 진단·모니터링 지표입니다. 임상 검사실에서는 Real-Time Quantitative PCR(RQ-PCR)을 통해 BCR-ABL/ABL 비율(IS 척도)을 정기적으로 측정하여 치료 반응을 추적합니다.

기존에는 qPCR 기기에서 출력된 원시 데이터를 수기로 Excel에 정리하고, QC 판정·IS 보정계수 적용·결과 보고를 개별적으로 수행하였습니다. 이 과정은 시간이 많이 소요되고 계산 오류가 발생하기 쉬웠습니다.

본 시스템은 웹 브라우저 단독으로 동작하는 단일 HTML 파일 형태로, 데이터 업로드부터 최종 Excel 보고서 출력까지 전 과정을 자동화합니다.

---

## 2. 비즈니스 목표

| # | 목표 | 측정 기준 |
|---|------|-----------|
| B1 | 결과 처리 시간 단축 | 수기 대비 80% 이상 단축 |
| B2 | 계산 오류 제거 | NCN·IS-NCN 수작업 계산 오류 0건 |
| B3 | QC 판정 표준화 | 검사자 간 판정 불일치 제거 |
| B4 | 보고서 포맷 통일 | 국제표준(IS) 포맷 준수 |
| B5 | 서버 불필요 운영 | 별도 서버·DB 없이 브라우저에서 완결 |

---

## 3. 이해관계자

| 역할 | 설명 |
|------|------|
| 임상병리사 | 검사 수행, 결과 입력, 보고서 출력 |
| 검사실 관리자 | QC 기준·IS-CAL 보정계수 설정 관리 |
| 의사 (간접) | 최종 보고 결과 수신 |

---

## 4. 업무 요구사항

### BR-01. 데이터 입수
- qPCR 기기(QuantStudio 등)에서 출력된 CSV/Excel 원시 파일을 업로드
- 파일에서 BCR-ABL 및 ABL 두 타깃을 자동으로 식별

### BR-02. QC 자동 판정
- 국제 가이드라인 기준(슬로프, R², 효율, CT SD 등 9개 항목) 자동 판정
- 판정 결과(pass/FAIL) 및 상세값 자동 표시

### BR-03. IS 보정계수 적용
- IS-MMR Calibrator(IS-CAL)의 NCN을 기반으로 CF(보정계수) 자동 계산
- CF 분자(IS-Cal Value)를 수동으로도 설정 가능

### BR-04. 환자 결과 산출
- BCR-ABL/ABL copy 수, NCN(%), IS-NCN(%) 자동 계산
- 분자 반응 단계(Complete/Major/No Major Molecular Response) 자동 판정

### BR-05. 결과 보고서 출력
- 검사 양식에 맞는 Excel 파일 다운로드
- 시트 구성: 검사, 검사결과, 환자결과, QC결과, 컨트롤, 표 값, work list 값, Raw

### BR-07. LIS 인터페이스 파일 출력
- LIS import용 Excel 파일 별도 다운로드
- 환자별 5개 검사 항목을 검체번호·검사명·결과값 3열 구조로 출력
- 자체 개발 LIS에 직접 import하여 수동 입력 대체

### BR-06. 보정계수 이력 관리
- 날짜별 CF 이력 조회 및 관리

---

## 5. 제약사항

- 서버 없이 단일 HTML 파일로만 동작 (오프라인 사용 가능)
- 환자 데이터는 외부로 전송되지 않음 (브라우저 내에서만 처리)
- Excel 출력은 xlsx 포맷만 지원

---

## 6. 용어 정의

| 용어 | 설명 |
|------|------|
| IS | International Scale — BCR-ABL 측정의 국제 표준화 척도 |
| NCN | Normalized Copy Number — BCR-ABL/ABL × 100(%) |
| IS-NCN | IS 척도로 보정된 NCN |
| CF | Correction Factor — IS 보정계수 |
| AMR | Analytical Measurement Range — 정량 신뢰 범위 |
| CMR | Complete Molecular Response (IS-NCN < 0.01%) |
| MMR | Major Molecular Response (0.01% ≤ IS-NCN < 0.1%) |
| No MMR | No Major Molecular Response (IS-NCN ≥ 0.1%) |
| IS-CAL | IS-MMR Calibrator — CF 산출에 사용되는 표준물질 |
