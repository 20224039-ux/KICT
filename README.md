# KICT

- [`hollow_layer_fire/`](hollow_layer_fire/README.md) — 중공층(PIR) 화재시험(화원 상부 500–3000 mm): 온도이력·화재진행 단계, 화재확산 속도, 성장식(at²·지수·멱법칙), 높이별 화재경향성 분석
- [`cavity_fire_index/`](cavity_fire_index/README.md) — 중공층 100T 외단열 재료별(EPS·PIR(PUR) 가연/준불연, PF 준불연 외부용/내부용) 화재확산 경향 비교와 3000 mm 기준 at²형 화재위험지수(α₆₀₀ = 600 K / t₆₀₀², 등급·심각도·강건성 검증), 원본 엑셀 데이터 품질점검, 열물성(k·ρ·c·착화온도·발열량) 기반 외벽 화재위험 평가식([RISK_MODEL.md](cavity_fire_index/RISK_MODEL.md))
- [`tga_dsc/`](tga_dsc/README.md) — 단열재 6종 TGA–DSC(N₂·산화 분위기) 열분해·산화 특성과 중공층 화재시험 연계(준불연 4종: 공기 중 최대 질량감소속도 순서 = 화재확산 순서, 가연성 PIR(PUR)은 예외)
- [`coupled_analysis/`](coupled_analysis/README.md) — 중공층 시험(30 kW 버너 20분) × TGA–DSC 연계: 열전도율을 직접 구할 수 없는 이유와 유효 열관성 역산(시스템 착화 저항), t₆₀₀의 착화·확산 단계 분해, 가연성 PIR(PUR)의 30 s 전체 화염, 버너 소화 후 자립 연소
- [`integrated_model/`](integrated_model/README.md) — 통합 해석·평가식: 가연/준불연 두 영역과 준불연 4메커니즘, TGA 예측식 t̂₆₀₀ = 248 s + 423 / r_max,air(준불연, 교차검증 P 등급 6/6), 가연성 PIR(PUR) 블라인드 검증(예측 777 s → 실측 30 s, 실패), 등급 확인 → TGA 선별 → 중규모 시험의 3단계 위험성 평가(R = P × S + 플래그 A·B)
