# 교복구매 길라잡이 프로젝트 지침

교복 학교주관구매 업무를 안내하고 Excel·HWPX·PDF 서식 출력을 자동화하는 프로젝트임.

## 프로젝트 정보

- **우선 목표**: 원본 HWPX 서식을 근거로 Excel XLSM에서 기초자료를 한 번 입력해 서식별 반영·선택 출력·PDF 출력을 구현함. Excel에서 HWPX 출력이 곤란하면 대표 POC와 ADR 검증 후 웹앱에서 HWPX 생성·다운로드를 구현함.
- **주요 자료**: 교복 매뉴얼 HWPX, 기존 용역계약 길라잡이 XLSM, [계획.md](계획.md)
- **지원 환경**: Windows, Microsoft 365, 한컴오피스
- **원본 보호**: 제공된 HWPX·XLSM은 절대 덮어쓰지 않음. 사본·생성본은 `artifacts/`와 `_workspace/`에 둠.
- **데이터 경계**: 학생·학부모 개인정보와 직인 자동처리는 MVP에서 제외함. 문서 출력에 필요한 업무·계약 정보는 중앙 서버에 저장하지 않으며 로그·Git·대화에 원문을 노출하지 않음.

## 자주 사용하는 명령

- **Excel MVP 초안 생성**: `powershell -ExecutionPolicy Bypass -File scripts/run_excel_mvp_utf8.ps1` (내부적으로 `scripts/create_excel_mvp.ps1`을 UTF-8로 재해석해 실행함. Excel COM 자동화이므로 Microsoft 365가 설치된 Windows에서만 동작하며, `서식_인벤토리.md`의 Form ID 목록(F-001~F-052, 52개)을 파싱해 `artifacts/excel/교복구매_길라잡이_MVP_초안.xlsm`을 새로 생성함.)
- **Phase 검증 Hook 수동 실행**: `powershell -ExecutionPolicy Bypass -File .Codex/hooks/verify-phase.ps1` (Write/Edit·TaskCompleted·Stop 시 자동 실행되며, `.Codex/quality-gate.json`의 `activePhase`·`phaseTestIds`·`sourceHashes` 기준으로 체크리스트·테스트 상태·원본 해시를 검사해 미완료 시 완료 처리를 차단함.)
- **원본 무결성 확인**: 원본 `.xlsm`·`.hwpx`의 SHA-256을 계산해 `.Codex/quality-gate.json`의 `sourceHashes`와 대조함.
- **Markdown 공백 오류 검사(Q-02 게이트)**: `git diff --check`

## 아키텍처 개요

- **Phase 파이프라인**: P0(원본 무결성·Kordoc/XLSM 정적분석) → P1(Form ID 인벤토리·데이터 모델·POC 범위 확정) → P2(HWPX POC) → P3(Excel MVP) → P4(현장 검증) → P5(웹앱). 각 Phase의 세부 작업 ID·담당·선행·상태는 [task.md](task.md)에, 전체 로드맵과 근거는 [계획.md](계획.md)에 있음. 현재 활성 Phase는 `.Codex/quality-gate.json`의 `activePhase`(`P3-01`)이며, P2-01·P3-01은 [task.md](task.md) 하단 차단 사유에 따라 BLOCKED 상태임.
- **에이전트 역할 분리**(`.Codex/agents/`): `uniform-workflow-analyst`(업무·서식 분석 → `서식_인벤토리.md`/`입력데이터_사전.md`/`서식_매핑표.md` 작성) → `document-automation-engineer`(Kordoc 기반 HWPX 템플릿화·치환·검증) / `excel-automation-engineer`(XLSM 클린룸 재구현, 외부링크·매크로 원본 이식 금지) → `quality-compliance-reviewer`(독립 검토, Critical/High 결함 시 반려하며 완료 처리하지 않음). 각 담당자의 작업기록은 `_workspace/0N_*/`, 최종 산출물은 `artifacts/`에 분리해 둠.
- **스킬 오케스트레이션**(`.Codex/skills/`): `uniform-purchase-orchestrator`가 분석(병렬)→문서자동화/Excel구현→검증(파이프라인) 순서로 에이전트 실행을 조율함. `uniform-purchase-goal`은 `/goal` 기반 목표 실행·상태 확인·재개·해제를 담당함.
- **품질 게이트**(`.Codex/quality-gate.json` + `.Codex/hooks/verify-phase.ps1`): Write/Edit·TaskCompleted·Stop 시점마다 활성 Phase의 [체크리스트.md](체크리스트.md)·[test.md](test.md)·원본 SHA-256을 검사함. FAIL·PENDING·미해결 BLOCKED·체크리스트 미완료가 있으면 완료를 차단함(Q-05).
- **핵심 문서 지도**: `서식_인벤토리.md`(Form ID F-001~F-057 목록, 개인정보 가능 문서 F-029/F-047/F-050 포함) · `입력데이터_사전.md`(공통/문서별/반복/계산/보호 필드 정의) · `서식_매핑표.md`(F-001~F-052 출력 매핑, F-053~F-057은 읽기 전용) · `결정사항.md`(범위·대표 POC 3종 확정 근거) · `엑셀_벤치마크_분석.md`(기존 XLSM 정적분석) · `kordoc_기술검증.md`(HWPX 파싱/치환 기술검증) · `로그.md`(세션 재개 기록, `/clear` 전 필수 갱신).
- **Excel 생성 경로**: `scripts/create_excel_mvp.ps1`이 인벤토리를 파싱해 Excel COM으로 8개 시트(`01_교복구매_워크플로우`/`02_기초자료_입력`/`03_서식선택_출력`/`DB`/`Ref_data`/`학교정보`/`서식Metadata`/`사용설명서`) 구조를 생성함. 현재는 구조 초안 수준이며, 기준 XLSM 수준의 실무 입력화면·자동반영·선택출력 기능은 미구현 상태임([test.md](test.md) X-06 FAIL 참고).
- **원본 자산(수정 금지)**: 루트의 `20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm`, `교복4차 (2)/*.hwpx`. 모든 파생 작업은 `_workspace/`와 `artifacts/`에서만 수행함.

## 하네스: 교복구매 길라잡이

- **실행 모드**: 분석은 병렬, 문서 자동화→Excel 구현→검증은 파이프라인으로 진행함.
- **오케스트레이터**: 교복구매 길라잡이의 분석·구현·검증·재개 요청에는 `.Codex/skills/uniform-purchase-orchestrator/SKILL.md`를 사용함.
- **Goal 운영**: `/goal`, 목표 기반 자동 실행, 상태 확인·재개·해제 요청에는 `.Codex/skills/uniform-purchase-goal/SKILL.md`를 사용함.
- **에이전트**: `uniform-workflow-analyst`, `document-automation-engineer`, `excel-automation-engineer`, `quality-compliance-reviewer` 역할을 사용함.
- **상태 기준**: [계획.md](계획.md)의 가장 높은 우선순위 `TODO` 또는 `IN_PROGRESS` 작업부터 수행하며, 품질 검토에서 Critical/High 결함이 남으면 완료로 보고하지 않음.

## 작업 원칙

- 항상 한글로 답변함.
- 구현 전 사실·가정·선택지를 구분함. 불확실한 법령·계약 문구·도구 동작은 원문 또는 실제 실행으로 확인함.
- 한 작업에는 사용자 요청을 충족하는 최소 변경만 적용함. 요청하지 않은 리팩터링·추상화·기능 확장은 하지 않음.
- 기존 스타일과 공식 서식의 구조를 유지함. 원본보다 보기 좋다는 이유만으로 공문·계약서의 필수 구조를 변경하지 않음.
- 변경한 모든 줄은 사용자 요청과 완료조건에 연결되어야 함. 변경으로 생긴 미사용 요소만 제거함.
- 여러 단계 작업은 간단한 계획, 단계별 검증, 결과 기록의 순서를 지킴.
- 실패·미확정·차단 요인은 숨기지 않고 원인, 시도 결과, 가능한 다음 조치를 기록함.

## 검증과 완료 기준

- 수정 전 대상 파일을 실제로 읽음.
- Excel 산출물은 외부 링크·외부 연결·`#REF!` 이름 정의가 없는지 확인함.
- HWPX 산출물은 Kordoc 구조검증, 재열기, PDF 비교 및 한컴 수동 확인 항목을 분리해 기록함.
- 골든 테스트는 정상값, 필수값 누락, 최대 길이/특수문자, 복수 반복행을 포함함.
- 생성 문서는 행정실 직원이 사실관계·계약조건·서식 상태를 최종 확인한 뒤 출력함.
- 배포, 외부 전송, 강제 푸시, 대량 삭제, 시크릿 수정은 반드시 사용자 확인 후 진행함.

## /goal 사용 원칙

- `/goal`은 Codex의 세션 범위 내장 기능이며 하나의 측정 가능한 완료 조건에만 사용함.
- 목표 문장에는 완료조건, 검사 방법, 원본 보호와 같은 제약을 함께 적음.
- `/goal` 평가자는 대화에 나타난 정보로만 판정하므로 명령·테스트·검증 결과를 작업 기록에 남김.
- `/goal`은 권한을 바꾸지 않음. 자동 모드를 쓰더라도 배포·외부 전송 등 파급 작업은 별도 확인이 필요함.

## 변경 이력

| 날짜 | 변경 | 대상 | 사유 |
|---|---|---|---|
| 2026-09-15 | 교복구매 길라잡이 하네스 초기 구성 | `.Codex/agents`, `.Codex/skills`, `AGENTS.md` | 서식 분석, HWPX 자동화, XLSM 구현, 품질 검증 역할을 분리하고 /goal 운영 기준을 등록함 |
