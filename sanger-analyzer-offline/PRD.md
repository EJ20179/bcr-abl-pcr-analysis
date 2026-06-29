# PRD — Sanger Sequencing Analyzer (Offline)

## 1. 제품 개요

단일 HTML 파일로 배포되는 완전 오프라인 생거 시퀀싱 분석기.  
ABI 파일 업로드 → 레퍼런스 정렬 → 변이 감지 → F/R 쌍 비교까지 브라우저에서 처리한다.

**배포:** https://sanger-analyzer-offline.vercel.app  
**소스:** `sanger-analyzer-offline/index.html` (단일 파일, ~3.4 MB)  
**내장 라이브러리:** JSZip v3.10.1, Chart.js (UMD), Font Awesome (인라인)

---

## 2. 아키텍처 개요

```
index.html (단일 파일)
├── CSS (인라인, CSS 변수 기반 디자인 시스템)
├── HTML 구조
│   ├── [파일 업로드] 탭 — 레퍼런스 입력 + ABI 업로드
│   ├── [분석 결과] 탭 — 샘플별 / 프라이머 쌍별 결과
│   └── [Reference 관리] 탭 — 라이브러리 CRUD
└── JavaScript (인라인)
    ├── ABIParser          — ABIF 바이너리 파싱
    ├── SequenceAligner    — Smith-Waterman 정렬
    ├── ChromatogramRenderer — Canvas 크로마토그램
    ├── DocxParser         — Word .docx 파싱 (JSZip + XML)
    ├── SequenceView       — Canvas 정렬 시각화
    └── 앱 로직            — 탭·상태·렌더링·내보내기
```

---

## 3. 화면 및 탭 구조

### 3.1 [파일 업로드] 탭

#### 3.1.1 Reference Sequence 패널

| 기능 | 상세 |
|------|------|
| **입력 방법 전환** | `텍스트 직접 입력` / `FASTA 파일` / `Word 파일` 탭 전환 |
| **텍스트 입력** | ATCG 서열 붙여넣기, GC% · 길이 자동 표시 |
| **FASTA 파일** | 드래그&드롭 또는 클릭 업로드, 헤더 파싱 후 서열 추출 |
| **Word(.docx) 파일** | DocxParser로 프라이머/Exon 정보 자동 추출 |
| **유전자명 입력** | 예: `BRCA1`, `TP53`, `EGFR exon 19` |
| **저장** | 현재 레퍼런스를 라이브러리에 저장 (localStorage) |
| **라이브러리 불러오기** | 저장된 레퍼런스 목록에서 선택 |

**분석 설정 (사이드바):**
- 정렬 파라미터: Match Score, Mismatch Penalty, Gap Penalty
- 품질 임계값(QV threshold)
- 트리밍 윈도우 크기

#### 3.1.2 ABI 파일 업로드 패널

| 기능 | 상세 |
|------|------|
| **파일 업로드** | 드래그&드롭 / 클릭 선택, 복수 파일 동시 업로드 |
| **자동 파싱** | ABIF 바이너리 파싱 — 서열, 품질점수(Phred), 피크 위치, Trace 데이터 추출 |
| **F/R 자동 페어링** | 파일명 패턴으로 Forward/Reverse 쌍 자동 매칭 (6단계 Tier 알고리즘) |
| **파일 목록** | 파싱 상태(로딩/완료/오류), 선택 삭제, 전체 삭제 |
| **데모 데이터** | 샘플 데이터로 즉시 체험 가능 |

**F/R 자동 페어링 Tier 우선순위:**
1. 동일 Exon + 동일 샘플명 + 명시적 F/R
2. 동일 Exon(서브 포함) + 동일 베이스명
3. 동일 Exon(서브 포함), 베이스명 무관
4. 동일 베이스명 (Exon 무관)
5. 남은 F/R 아무거나 페어링
6. 방향 불명 — 베이스명 동일한 것끼리 묶기

---

### 3.2 [분석 결과] 탭

**3단계 탭 구조:** 샘플 탭 → 프라이머(Exon) 탭 → 결과 패널

#### 3.2.1 샘플 탭
- 샘플명 + F 개수/R 개수 배지
- Exon 개수 배지 (Word 파일에서 추출된 경우)

#### 3.2.2 프라이머 탭 (Exon 단위)
- 페어 그룹: Exon 배지 + `정방향프라이머 / 역방향프라이머`
- 단일 그룹: 방향 표시

#### 3.2.3 F/R 페어 결과 패널

**① 크로마토그램 (Canvas)**
- Forward / Reverse 각각 파형(G/A/T/C 채널) 렌더링
- 품질점수 히스토그램 (Chart.js)
- 이미지 저장(PNG) 버튼

**② 정렬 시각화 (Canvas — `SequenceView`)**
- 레퍼런스 행 + F 행 + R 행 3행 구성
- 컬럼 색상 의미:
  - **빨강 배경**: F와 R 모두 레퍼런스와 다르고 서로 일치 → **Bilateral confirmed variant**
  - **노랑 배경**: F와 R이 서로 다름 → Discordant (불일치, 아티팩트 의심)
  - **회색**: 매칭
- 프라이머/Exon 위치 annotation 바 (Word 파일 연동)
- 가로 스크롤 / 줌 지원

**③ 변이 요약 (Variant Summary)**

| 항목 | 내용 |
|------|------|
| **배지** | SNV 개수 / Insertion 개수 / Deletion 개수 / 전체 개수 |
| **변이 테이블** | Type, 위치(ref pos), RefBase, AltBase, Length, HGVS, Source(F/R), Confirmed 여부 |
| **HGVS 표기** | `g.123A>T` (SNV), `g.123_124insATG` (ins), `g.123_125del` (del) |
| **CSV 내보내기** | 변이 테이블 전체 다운로드 |
| **FASTA 내보내기** | 정렬된 서열 다운로드 |

**④ 품질 통계**
- 평균 QV, 고품질 염기(QV≥20) 비율, Trim 후 길이

---

### 3.3 [Reference 관리] 탭

| 기능 | 상세 |
|------|------|
| **라이브러리 목록** | 저장된 레퍼런스 카드 목록 |
| **카드 구성** | 유전자명, 저장일시, 서열 길이, 버전 이력 |
| **불러오기** | 업로드 탭에 해당 레퍼런스 적용 |
| **업데이트** | 새 파일로 레퍼런스 갱신, 이전 버전 보존 |
| **버전 복원** | 이전 버전으로 롤백 |
| **삭제** | 개별 버전 또는 전체 레퍼런스 삭제 |
| **JSON 내보내기** | 전체 라이브러리를 JSON으로 백업 |

**저장 위치:** `localStorage` 키 (`sanger-ref-library` 등)

---

## 4. 핵심 모듈 상세

### 4.1 ABIParser
- ABIF 매직 검증 (`ABIF`)
- 디렉토리 파싱 (28바이트 엔트리, 태그 기반)
- 추출 데이터: `PBAS.2` (서열), `PCON.2` (품질), `PLOC.2` (피크위치), `DATA.9~12` (G/A/T/C 채널), `SMPL.1` (샘플명), `RUND.1` (실행일)
- 품질 기반 자동 트리밍 (슬라이딩 윈도우)
- 실제 트레이스 데이터 없을 시 합성 데이터 생성

### 4.2 SequenceAligner
- **Smith-Waterman** 지역 정렬 (최대 3,000 bp)
- 파라미터: Match +2, Mismatch -1, Gap open -2
- 변이 호출: SNV / Insertion / Deletion 구분
- 정렬 통계: identity%, 매치수, gap수, 미스매치수

### 4.3 DocxParser
- JSZip으로 `.docx` 압축 해제 → `word/document.xml` 파싱
- 추출: 서열 블록, 프라이머(방향/Exon번호/위치), Exon 어노테이션
- Word 하이라이트 색상으로 F(파랑)/R(빨강) 프라이머 구분
- 앰플리콘 서열 자동 계산 (F 프라이머 시작 ~ R 프라이머 끝)
- 게놈 위치(genomic position) 매핑 지원

### 4.4 SequenceView (Canvas 정렬 렌더러)
- DPR(Device Pixel Ratio) 보정으로 레티나 디스플레이 대응
- 행 구성: Reference / Forward / Reverse
- Exon 어노테이션 바: 색상 사이클링, 프라이머 화살표 오버레이
- 룰러(위치 눈금자), 열 배경색, 베이스 문자 렌더링

---

## 5. 데이터 흐름

```
[ABI 파일] → ABIParser → entry(서열, 품질, 트레이스, 메타)
[레퍼런스]  → parseRef* → state.refSeq
                ↓
         SequenceAligner.align(query, refSeq)
                ↓
         entry.result = { alignedQuery, alignedRef, variants, identity, ... }
                ↓
    renderResults()
    ├── ChromatogramRenderer → <canvas> 크로마토그램
    ├── SequenceView → <canvas> 정렬 시각화
    ├── makeMergedVariantSection → 변이 테이블 HTML
    └── makePairQualitySection → 품질 통계 HTML
```

---

## 6. 상태 관리 (Global State)

```javascript
state = {
  refSeq: string,           // 현재 레퍼런스 서열
  refParsed: object,        // DocxParser 결과 (프라이머/Exon 포함)
  geneName: string,
  abiFiles: entry[],        // 업로드된 ABI 파일 배열
  analysisSettings: {
    matchScore, mismatchScore, gapOpen, qualThreshold, trimWindow
  }
}
```

---

## 7. 파일명 파싱 규칙 (F/R 자동 인식)

인식하는 방향 토큰 (대소문자 무관):

| 방향 | 토큰 예시 |
|------|----------|
| Forward | `_F`, `-F`, `_fwd`, `-fwd`, `_forward`, `F1`, `f1` |
| Reverse | `_R`, `-R`, `_rev`, `-rev`, `_reverse`, `R1`, `r1` |

Exon 인식 패턴: `exon3`, `ex3`, `e3`, `E3`, `EX3`

---

## 8. 내보내기 포맷

### FASTA
```
>SampleName_Forward
ATCGATCG...
>SampleName_Reverse
CGTAGCTA...
```

### 변이 CSV
```csv
Type,Position,RefBase,AltBase,Length,HGVS,Source,Confirmed
SNV,234,A,T,1,g.234A>T,F+R,Yes
INS,456,-,ATG,3,g.456_457insATG,F,No
```

### 레퍼런스 라이브러리 JSON
```json
[
  {
    "name": "BRCA1 exon11",
    "savedAt": "2025-01-01T00:00:00.000Z",
    "sequence": "ATCG...",
    "primers": [...],
    "exons": [...],
    "versions": [...]
  }
]
```

---

## 9. 향후 개선 고려사항

| 항목 | 내용 |
|------|------|
| 다중 파일 배치 분석 결과 ZIP 일괄 다운로드 | 현재 개별 다운로드만 가능 |
| 정렬 파라미터 프리셋 저장 | 유전자별 최적 파라미터 저장 |
| 변이 주석 DB 연동 | ClinVar 등 오프라인 스냅샷 |
| 결과 HTML 리포트 내보내기 | 그래프 포함 단일 HTML |
| 모바일 레이아웃 최적화 | 현재 900px 미만 1컬럼 지원 |

---

## 10. 개발 환경 및 배포

| 항목 | 내용 |
|------|------|
| 소스 파일 | `sanger-analyzer-offline/index.html` |
| Vercel 프로젝트 | `sanger-analyzer-offline` |
| Vercel 프로젝트 ID | `prj_szd9CpfOk3ZEhck6dqI91miOQuOL` |
| 조직 ID | `team_nyZ9Xks1RoZLg7MBa4cmtHro` |
| 배포 방법 | `vercel --prod` (index.html 단일 파일) |
| 로컬 실행 | 브라우저에서 index.html 직접 열기 또는 `npx serve .` |
