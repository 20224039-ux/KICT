# KICT

- [`hollow_layer_fire/`](hollow_layer_fire/README.md) — 중공층(PIR) 화재시험(화원 상부 500–3000 mm): 온도이력·화재진행 단계, 화재확산 속도, 성장식(at²·지수·멱법칙), 높이별 화재경향성 분석
- [`cavity_fire_index/`](cavity_fire_index/README.md) — 중공층 100T 외단열 재료별(EPS 가연/준불연, PIR, PF 외장/내장) 화재확산 경향 비교와 3000 mm 기준 at²형 화재위험지수(α₆₀₀ = 600 K / t₆₀₀², 등급·심각도·강건성 검증), 원본 엑셀 데이터 품질점검, 열물성(k·ρ·c·착화온도·발열량) 기반 외벽 화재위험 평가식([RISK_MODEL.md](cavity_fire_index/RISK_MODEL.md))
- [`tga_dsc/`](tga_dsc/README.md) — 단열재 6종(EPS·PIR 가연/준불연, PF 외장/내장) TGA–DSC(N₂·산화 분위기) 열분해·산화 특성과 중공층 화재시험 연계(공기 중 최대 질량감소속도 순서 = 화재확산 순서)
- [`coupled_analysis/`](coupled_analysis/README.md) — 중공층 시험(30 kW 버너 20분) × TGA–DSC 연계: 열전도율을 직접 구할 수 없는 이유와 유효 열관성 역산(시스템 착화 저항), t₆₀₀의 착화·확산 단계 분해(확산 단계 = TGA 공기 r_max 순서), 버너 소화 후 자립 연소(= TGA 탄화층 산화 비율 f₄₀₀ 순서)
- [`integrated_model/`](integrated_model/README.md) — 통합 해석·평가식: 화재 진행 4메커니즘(착화=시스템, 확산=공기 중 분해속도, 소화 후 지속=탄화층 산화, EPS 용융·순간연소), TGA 예측식 t̂₆₀₀ = 220 s + 418 / r_max,air(교차검증 P 등급 5/5), TGA 선별 → 중규모 시험 확정의 2단계 위험성 평가(R = P × S + 플래그 A·B)
