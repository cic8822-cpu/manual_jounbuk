# 테스트 계획 및 결과 — 교복구매 길라잡이

- **자동 검사 상태**: `PASS — P0 완료 게이트`
- **최종 실행일**: 2026-09-15
- **자동 검사 명령**: `.claude/hooks/verify-phase.ps1`

## 1. 공통 품질 게이트

| ID | 검사 | 통과 조건 | P0 결과 |
|---|---|---|---|
| Q-01 | Markdown 문서 구조 | 필수 문서가 존재하고 H1 제목이 있음 | PASS |
| Q-02 | Git 공백 오류 | `git diff --check` 성공 | PASS |
| Q-03 | 원본 XLSM 무결성 | 기준 SHA-256 일치 | PASS |
| Q-04 | Hook 설정 유효성 | `.claude/settings.json` JSON 및 Hook 경로 유효 | PASS |
| Q-05 | 완료 차단 | 실패·미해결 Critical/High·PENDING·BLOCKED·체크리스트 미완료 시 종료 Hook이 완료를 차단 | PASS |

## 2. Excel MVP 골든 테스트

| ID | 시나리오 | 기대 결과 | 상태 |
|---|---|---|---|
| E-01 | 정상 학교 1개 | 필수 공통값이 서식에 반영됨 | NOT_RUN |
| E-02 | 필수값 누락 | 출력 전 누락 항목이 명확히 표시됨 | NOT_RUN |
| E-03 | 최대 길이·특수문자 | 잘림·문자 깨짐 없이 출력됨 | NOT_RUN |
| E-04 | 복수 업체·위원·품목 | 반복행이 누락·오염 없이 출력됨 | NOT_RUN |
| E-05 | 외부 의존성 | 새 XLSM 외부 링크·연결·`#REF!` 이름 정의 0건 | NOT_RUN |
| E-06 | PDF 출력 | 결재란·병합표·페이지 잘림 없음 | NOT_RUN |

## 3. HWPX POC 테스트

P2에서 대표 3종 사본을 대상으로 값 치환, 재열기, `validate`, PDF 비교, 한컴 수동 열기 결과를 추가함.

## 4. 실패 처리

자동 검사 또는 수동 테스트 실패 시 상태를 `FAIL`로 바꾸고 [로그.md](로그.md)에 결함·원인·수정·재검증 결과를 기록함. `FAIL`, `PENDING`, 미해결 `BLOCKED`가 있는 Phase는 완료로 보고하지 않음.
