# 콘칼로리미터 반복시험 통계 (PUF 스킨층, Table 2 평균 ± 표준편차)

리뷰어 코멘트("triplicate 측정이므로 Table 2 주요 연소 인자에 표준편차 제시")에 대응하기 위해,
시료별 개별 시험 곡선에서 값을 구한 뒤 평균 ± 표준편차를 계산하는 엑셀 통계표를 만든다.
레이아웃은 연구실 STDEV.P 양식(항목별 블록 × 시료 × 반복 1~4회 → 표준편차·평균·±·평균−표준편차·
(평균−표준편차)/평균)을 따른다.

시험 데이터·원고·생성된 엑셀은 미공개 자료이므로 저장소에 올리지 않는다(`data/`, `output/`은 `.gitignore`).

## 입력 (`data/`)

| 파일 | 내용 |
| --- | --- |
| `141B_HRR.xlsx`, `141B_THR.xlsx`, `141B_SPR.xlsx`, `141B_CO.xlsx`, `141B_CO2.xlsx` | 시료별 시트(BSK, SKM, S1H, S2H, S1V, S2V), 2행 머리글(`HRR (kW/m²)` 등), 3–123행 = 0–600 s(5 s 간격) 시험 곡선 |
| `manuscript.docx` (선택) | 원고. `Sample`로 시작하는 Table 2를 읽어 현재 값과 비교 |

## 실행

```bash
pip install openpyxl numpy python-docx
python build_table2_stats.py --manuscript data/manuscript.docx
# 캐시 값을 채우려면 LibreOffice(Calc)로 한 번 재계산
```

## 출력 시트 (`output/141B_PUF_cone_replicate_stats.xlsx`)

- `STDEV(1) HRR·THR`, `STDEV(2) SPR`, `STDEV(3) CO·CO2` — 원본 양식 블록: 시험별 값(원자료 시트를 참조하는 수식),
  n, STDEV.P, 평균, ±, 평균−표준편차, (평균−표준편차)/평균(= 1 − CV), 93% 판정(= CV ≤ 7%), STDEV.S, 비고(피크 시각 등)
- `Table 2 (mean±SD)` — 논문용 "평균 ± SD" 문자열과 n(SD 종류·소수 자릿수 선택), CV(%) 표, 오차막대용 숫자 표
- `Table 2 비교` — 원고 Table 2 값과 개별 시험 기반 재계산 평균 비교(원고 자릿수로 반올림해 일치 여부 표시)
- `개별값 (Table S)` — 시험별 개별값
- `원자료_*` — 원본 시험 곡선(값 수정 없음)

## 지표 정의 (시험마다 계산 후 평균)

| 지표 | 정의 |
| --- | --- |
| HRRpeak, SPRpeak | 시험 곡선의 최대값 |
| HRRmean, SPRmean, COmean, CO₂mean | 0–600 s 산술평균(121점) |
| THRmean | THR 누적곡선의 0–600 s 시간평균(현재 원고 Table 2의 정의) |
| THR at 600 s | 시험 종료 시점 THR(참고 열) |

모든 통계값은 수식이며, 스크립트는 같은 값을 numpy로 따로 계산해 출력하므로 재계산 결과와 대조할 수 있다.
