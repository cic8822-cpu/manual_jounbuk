# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

교복 학교주관구매 업무를 안내하고 Excel·HWPX·PDF 서식 출력을 자동화하는 프로젝트임. `AGENTS.md`에도 같은 취지의 지침이 있으니 두 파일을 함께 갱신함(중복 서술은 피하고 최신 상태만 반영).

## 프로젝트 정보

- **우선 목표**: 원본 HWPX 서식을 근거로 Excel XLSM에서 기초자료를 한 번 입력해 서식별 반영·선택 출력·PDF 출력을 구현함. Excel에서 HWPX 출력이 곤란하면 대표 POC와 ADR 검증 후 웹앱에서 HWPX 생성·다운로드를 구현함.
- **주요 자료**: 교복 매뉴얼 HWPX, 기존 용역계약 길라잡이 XLSM, [계획.md](계획.md)
- **지원 환경**: Windows, Microsoft 365, 한컴오피스
- **원본 보호**: 제공된 HWPX·XLSM은 절대 덮어쓰지 않음. 사본·생성본은 `artifacts/`와 `_workspace/`에 둠.
- **데이터 경계**: 학생·학부모 개인정보와 직인 자동처리는 MVP에서 제외함. 문서 출력에 필요한 업무·계약 정보는 중앙 서버에 저장하지 않으며 로그·Git·대화에 원문을 노출하지 않음.

## 자주 사용하는 명령

- **Excel 안전 빌드(현재 활성 배포 경로)**: `powershell -ExecutionPolicy Bypass -File scripts/run_build_excel_v1_structure_utf8.ps1` — 내부적으로 구조 생성(`scripts/build_excel_v1_structure.ps1`) → VBA 주입(`scripts/build_excel_v1_vba.ps1`) → 사본 검증(`scripts/verify_excel_v1.ps1`)을 순서대로 UTF-8로 재해석해 실행함. Excel COM 자동화이므로 Microsoft 365가 설치된 Windows에서만 동작함. 검증을 통과하면 `artifacts/excel/교복구매_길라잡이_YYYYMMDD_vN.xlsm`으로 저장하며, 같은 날짜에 다시 통과할 때마다 `vN`을 증가시켜 기존 통과본을 보존함(기존 파일을 덮어쓰지 않음).
  - `build_excel_v1_structure.ps1`·`build_excel_v1_vba.ps1`·`verify_excel_v1.ps1`을 직접 실행하면 임시 `artifacts/excel/build.xlsm`을 대상으로 한 개별 단계만 수행하며, 배포 후보(`v N.xlsm`)를 갱신하지 않음.
  - `scripts/create_excel_mvp.ps1`/`run_excel_mvp_utf8.ps1`은 초기 8시트 구조 초안용 구버전 스크립트로, P3-01 진행 중 "구조 초안 수준"이라는 이유로 완료 상태가 철회됨([task.md](task.md) 차단 사유 참고). 더 이상 활성 경로가 아니므로 새 작업에는 사용하지 않음.
- **Phase 검증 Hook 수동 실행**: `powershell -ExecutionPolicy Bypass -File .claude/hooks/verify-phase.ps1` (Write/Edit·TaskCompleted·Stop 시 `.claude/settings.json` Hook으로 자동 실행되며, `.claude/quality-gate.json`의 `activePhase`·`phaseTestIds`·`sourceHashes` 기준으로 체크리스트·테스트 상태·원본 해시를 검사해 미완료 시 완료 처리를 차단함.)
- **원본 무결성 확인**: 원본 `.xlsm`·`.hwpx`의 SHA-256을 계산해 `.claude/quality-gate.json`의 `sourceHashes`와 대조함.
- **Markdown 공백 오류 검사(Q-02 게이트)**: `git diff --check`
- **단일 테스트 선택 실행 불가**: `verify_excel_v1.ps1`은 [test.md](test.md)의 골든 시나리오(E-01~E-06)와 회귀 검사를 하나의 Excel COM 세션·검증 사본 안에서 순서대로 모두 실행하는 단일 스크립트이며, 개별 시나리오만 골라 실행하는 옵션이나 파라미터는 없음. 특정 시나리오(예: E-03) 결과만 확인하려면 전체 스크립트를 실행한 뒤 [test.md](test.md)에서 해당 ID의 PASS/PARTIAL/NOT_RUN 행을 확인함.

## 아키텍처 개요

- **Phase 파이프라인**: P0(원본 무결성·Kordoc/XLSM 정적분석) → P1(Form ID 인벤토리·데이터 모델·POC 범위 확정) → P2(HWPX POC) → P3(Excel MVP) → P4(현장 검증) → P5(웹앱). 각 Phase의 세부 작업 ID·담당·선행·상태는 [task.md](task.md)에, 전체 로드맵과 근거는 [계획.md](계획.md)에 있음. 현재 활성 Phase는 `.claude/quality-gate.json`의 `activePhase` 값이 기준임(자주 바뀌므로 이 파일에 상태를 하드코딩하지 않음 — 항상 [task.md](task.md)·[test.md](test.md)·[로그.md](로그.md)를 직접 확인함).
- **에이전트 역할 분리**(`.claude/agents/`): `uniform-workflow-analyst`(업무·서식 분석 → `서식_인벤토리.md`/`입력데이터_사전.md`/`서식_매핑표.md` 작성) → `document-automation-engineer`(Kordoc 기반 HWPX 템플릿화·치환·검증) / `excel-automation-engineer`(XLSM 클린룸 재구현, 외부링크·매크로 원본 이식 금지) → `quality-compliance-reviewer`(독립 검토, Critical/High 결함 시 반려하며 완료 처리하지 않음). 각 담당자의 작업기록은 `_workspace/0N_*/`, 최종 산출물은 `artifacts/`에 분리해 둠.
- **스킬 오케스트레이션**(`.claude/skills/`): `uniform-purchase-orchestrator`가 분석(병렬)→문서자동화/Excel구현→검증(파이프라인) 순서로 에이전트 실행을 조율함. `uniform-purchase-goal`은 `/goal` 기반 목표 실행·상태 확인·재개·해제를 담당함.
- **품질 게이트**(`.claude/quality-gate.json` + `.claude/hooks/verify-phase.ps1`): Write/Edit·TaskCompleted·Stop 시점마다 활성 Phase의 [체크리스트.md](체크리스트.md)·[test.md](test.md)·원본 SHA-256을 검사함. FAIL·PENDING·미해결 BLOCKED·체크리스트 미완료가 있으면 완료를 차단함(Q-05).
- **핵심 문서 지도**: `서식_인벤토리.md`(Form ID F-001~F-057 목록, 개인정보 가능 문서 F-029/F-047/F-050 포함) · `입력데이터_사전.md`(공통/문서별/반복/계산/보호 필드 정의) · `서식_매핑표.md`(F-001~F-052 출력 매핑, F-053~F-057은 읽기 전용) · `결정사항.md`(범위·대표 POC 3종 확정 근거) · `엑셀_벤치마크_분석.md`(기존 XLSM 정적분석) · `kordoc_기술검증.md`(HWPX 파싱/치환 기술검증) · `로그.md`(세션 재개 기록, `/clear` 전 필수 갱신).
- **Excel 생성 경로(v1, 활성)**: `scripts/build_excel_v1_structure.ps1`이 Excel COM으로 시트를 순서대로 생성함 — `사용설명서`·`기초자료입력`·`DB`(및 `DB_품목`/`DB_업체`/`DB_위원`/`DB_평가` 반복행 정규화 시트)·`서식선택_출력`(F-001~F-052 선택 목록)·Form ID별 서식 시트(`F-007_구매요청기안문`/`F-024_단가비율표`/`F-014_평가항목배점기준` 등, 계속 추가 중)·`학교정보`·`학교검색`·`절차안내`·`계약방법안내`. 이어서 `build_excel_v1_vba.ps1`이 매크로(신규/초기화/저장/수정/불러오기, 저장 직전 전체 DB 재검증 등)를 주입하고 `verify_excel_v1.ps1`이 외부 링크·연결·`#REF!` 이름 정의 0건과 골든 시나리오를 사본에서 검사함. 미구현 Form ID·현재 통과/실패 상세는 [test.md](test.md)를 확인함(하드코딩하지 않음).
- **원본 자산(수정 금지)**: 루트의 `20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm`, `교복4차 (2)/*.hwpx`. 모든 파생 작업은 `_workspace/`와 `artifacts/`에서만 수행함.

## 하네스: 교복구매 길라잡이

- **실행 모드**: 분석은 병렬, 문서 자동화→Excel 구현→검증은 파이프라인으로 진행함.
- **오케스트레이터**: 교복구매 길라잡이의 분석·구현·검증·재개 요청에는 `.claude/skills/uniform-purchase-orchestrator/SKILL.md`를 사용함.
- **Goal 운영**: `/goal`, 목표 기반 자동 실행, 상태 확인·재개·해제 요청에는 `.claude/skills/uniform-purchase-goal/SKILL.md`를 사용함.
- **에이전트**: `uniform-workflow-analyst`, `document-automation-engineer`, `excel-automation-engineer`, `quality-compliance-reviewer` 역할을 사용함.
- **상태 기준**: [계획.md](계획.md)의 가장 높은 우선순위 `TODO` 또는 `IN_PROGRESS` 작업부터 수행하며, 품질 검토에서 Critical/High 결함이 남으면 완료로 보고하지 않음.

## 작업 원칙

- 항상 한글로 답변함.
- 작업 전 실제 대상 파일과 프로젝트 지침을 읽고, 요청 범위·가정·검증 가능한 완료 조건을 구분함. 결론을 바꾸는 불확실성만 한 가지 질문으로 확인함.
- 구현 전 사실·가정·선택지를 구분함. 불확실한 법령·계약 문구·도구 동작은 원문 또는 실제 실행으로 확인함.
- 한 작업에는 사용자 요청을 충족하는 최소 변경만 적용함. 요청하지 않은 리팩터링·추상화·기능 확장은 하지 않음.
- 기존 스타일과 공식 서식의 구조를 유지함. 원본보다 보기 좋다는 이유만으로 공문·계약서의 필수 구조를 변경하지 않음.
- 변경한 모든 줄은 사용자 요청과 완료조건에 연결되어야 함. 변경으로 생긴 미사용 요소만 제거함.
- 여러 단계 작업은 간단한 계획, 단계별 검증, 결과 기록의 순서를 지킴.
- 실패·미확정·차단 요인은 숨기지 않고 원인, 시도 결과, 가능한 다음 조치를 기록함.
- 보고서와 시각 자료에서는 핵심 메시지의 우선순위, 확인된 수치·근거, 남은 병목·최종 확인 항목을 구분하여 제시함. 막연한 형용사로 근거를 대체하지 않음.

## 검증과 완료 기준

- 수정 전 대상 파일을 실제로 읽음.
- Excel 산출물은 외부 링크·외부 연결·`#REF!` 이름 정의가 없는지 확인함.
- HWPX 산출물은 Kordoc 구조검증, 재열기, PDF 비교 및 한컴 수동 확인 항목을 분리해 기록함.
- 골든 테스트는 정상값, 필수값 누락, 최대 길이/특수문자, 복수 반복행을 포함함.
- 생성 문서는 행정실 직원이 사실관계·계약조건·서식 상태를 최종 확인한 뒤 출력함.
- 배포, 외부 전송, 강제 푸시, 대량 삭제, 시크릿 수정은 반드시 사용자 확인 후 진행함.

## /goal 사용 원칙

- `/goal`은 Claude Code의 세션 범위 내장 기능이며 하나의 측정 가능한 완료 조건에만 사용함.
- 목표 문장에는 완료조건, 검사 방법, 원본 보호와 같은 제약을 함께 적음.
- `/goal` 평가자는 대화에 나타난 정보로만 판정하므로 명령·테스트·검증 결과를 작업 기록에 남김.
- `/goal`은 권한을 바꾸지 않음. 자동 모드를 쓰더라도 배포·외부 전송 등 파급 작업은 별도 확인이 필요함.

## 변경 이력

| 날짜 | 변경 | 대상 | 사유 |
|---|---|---|---|
| 2026-09-15 | 교복구매 길라잡이 하네스 초기 구성 | `.claude/agents`, `.claude/skills`, `CLAUDE.md` | 서식 분석, HWPX 자동화, XLSM 구현, 품질 검증 역할을 분리하고 /goal 운영 기준을 등록함 |
| 2026-09-16 | `/init` 재점검: 표준 헤더 추가, Excel 명령·생성 경로를 v1 활성 스크립트(`build_excel_v1_structure.ps1` 등)로 갱신, Phase/기능 상태 하드코딩 제거 | `CLAUDE.md` | `create_excel_mvp.ps1` 기반 구버전 명령이 실제 활성 경로와 달라져 있었고, 상태 서술이 `task.md`/`test.md`와 어긋날 위험이 있어 최신화함 |
| 2026-09-17 | `/init` 재점검: 내용은 실제 저장소 상태(`task.md`/`test.md`/`quality-gate.json`, 스크립트 경로, 에이전트·스킬 구성)와 일치함을 확인. `verify_excel_v1.ps1`이 골든 시나리오를 개별 선택 실행할 수 없는 단일 스크립트라는 점만 보완 | `CLAUDE.md` | 향후 인스턴스가 "단일 테스트 실행" 방법을 찾다가 존재하지 않는 옵션을 가정하지 않도록 명시함 |
| 2026-09-18 | `/init` 재점검: `CLAUDE.md`는 실제 스크립트·quality-gate 구성과 일치해 변경 없음. `AGENTS.md`의 "구현 Form ID: F-007, F-024만" 서술이 `test.md`(F-007·F-014~F-020·F-024 구현·검증 완료, 2026-09-17 기준)와 어긋나 있어 하드코딩을 제거하고 `test.md` 참조로 교체함 | `AGENTS.md` | Form ID 구현 상태가 자주 바뀌어 두 안내 파일 중 하나에만 최신화가 반영되면 어긋남이 재발하므로, `CLAUDE.md`가 이미 채택한 "상태 하드코딩 금지" 원칙을 `AGENTS.md`에도 동일 적용함 |
