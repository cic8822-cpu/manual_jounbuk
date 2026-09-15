$ErrorActionPreference = 'Stop'

# Hook 호스트가 표준 입력을 전달한 경우에만 소비해 수동 실행 시 대기하지 않도록 함.
if ([Console]::IsInputRedirected) {
    [void][Console]::In.ReadToEnd()
}

$root = $env:CLAUDE_PROJECT_DIR
if ([string]::IsNullOrWhiteSpace($root)) {
    $root = (Get-Location).Path
}

$gatePath = Join-Path $root '.claude\quality-gate.json'
$gate = $null
$requiredGateProperties = @('activePhase', 'phaseTasks', 'phaseChecklistHeading', 'phaseTestIds', 'sourceHashes')
$failures = [System.Collections.Generic.List[string]]::new()
if (-not (Test-Path -LiteralPath $gatePath -PathType Leaf)) {
    $failures.Add('활성 Phase 품질 게이트 설정 누락: .claude/quality-gate.json')
} else {
    try {
        $gate = Get-Content -LiteralPath $gatePath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $requiredGateProperties) {
            if ($null -eq $gate.$property) {
                $failures.Add("품질 게이트 필수 속성 누락: $property")
            }
        }
    } catch {
        $failures.Add("품질 게이트 JSON 오류: $($_.Exception.Message)")
    }
}

$requiredDocuments = @(
    '계획.md',
    'prd.md',
    'task.md',
    '로드맵.md',
    'test.md',
    '체크리스트.md',
    '로그.md'
)
foreach ($relativePath in $requiredDocuments) {
    $path = Join-Path $root $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $failures.Add("필수 문서 누락: $relativePath")
        continue
    }

    $firstNonEmptyLine = Get-Content -LiteralPath $path | Where-Object { $_.Trim() } | Select-Object -First 1
    if ($firstNonEmptyLine -notmatch '^#\s+') {
        $failures.Add("문서 제목 누락: $relativePath")
    }
}

$settingsPath = Join-Path $root '.claude\settings.json'
if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) {
    $failures.Add('Hook 설정 파일 누락: .claude/settings.json')
} else {
    try {
        $settings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
        foreach ($eventName in @('PostToolUse', 'TaskCompleted', 'Stop')) {
            $eventHooks = $settings.hooks.$eventName
            $commandHook = $null
            foreach ($eventHook in @($eventHooks)) {
                foreach ($hook in @($eventHook.hooks)) {
                    $hookArguments = @($hook.args)
                    $fileArgumentIndex = [array]::IndexOf($hookArguments, '-File')
                    $hasExpectedScript = $fileArgumentIndex -ge 0 -and $fileArgumentIndex -lt ($hookArguments.Count - 1) -and
                        $hookArguments[$fileArgumentIndex + 1] -eq '${CLAUDE_PROJECT_DIR}/.claude/hooks/verify-phase.ps1'
                    if ($hook.type -eq 'command' -and $hook.command -eq 'powershell.exe' -and $hasExpectedScript) {
                        $commandHook = $hook
                        break
                    }
                }
                if ($null -ne $commandHook) { break }
            }
            if ($null -eq $commandHook) {
                $failures.Add("완료 차단 Hook 설정 누락 또는 경로 오류: $eventName")
            }
        }
        if (-not (Test-Path -LiteralPath (Join-Path $root '.claude\hooks\verify-phase.ps1') -PathType Leaf)) {
            $failures.Add('완료 차단 Hook 스크립트 누락: .claude/hooks/verify-phase.ps1')
        }
    } catch {
        $failures.Add("Hook 설정 JSON 오류: $($_.Exception.Message)")
    }
}

$testPath = Join-Path $root 'test.md'
if ($null -ne $gate -and (Test-Path -LiteralPath $testPath -PathType Leaf)) {
    $testContent = Get-Content -LiteralPath $testPath -Raw
    foreach ($testId in @($gate.phaseTestIds)) {
        $escapedId = [regex]::Escape([string]$testId)
        $match = [regex]::Match($testContent, "(?im)^\|\s*$escapedId\s*\|.*\|\s*(PASS|FAIL|PENDING|NOT_RUN)\s*\|\s*$")
        if (-not $match.Success) {
            $failures.Add("활성 Phase 검사 항목 누락 또는 상태 형식 오류: $testId")
        } elseif ($match.Groups[1].Value -ne 'PASS') {
            $failures.Add("활성 Phase 검사 미통과: $testId = $($match.Groups[1].Value)")
        }
    }
}

$taskPath = Join-Path $root 'task.md'
if ($null -ne $gate -and (Test-Path -LiteralPath $taskPath -PathType Leaf)) {
    $taskContent = Get-Content -LiteralPath $taskPath -Raw
    foreach ($taskId in @($gate.phaseTasks)) {
        $escapedId = [regex]::Escape([string]$taskId)
        $match = [regex]::Match($taskContent, "(?im)^\|\s*$escapedId\s*\|.*\|\s*(DONE|TODO|IN_PROGRESS|BLOCKED)\s*\|(?:.*\|)?\s*$")
        if (-not $match.Success) {
            $failures.Add("활성 Phase 작업 누락 또는 상태 형식 오류: $taskId")
        } elseif ($match.Groups[1].Value -ne 'DONE') {
            $failures.Add("활성 Phase 작업 미완료: $taskId = $($match.Groups[1].Value)")
        }
    }
}

$checklistPath = Join-Path $root '체크리스트.md'
if ($null -ne $gate -and (Test-Path -LiteralPath $checklistPath -PathType Leaf)) {
    $checklistContent = Get-Content -LiteralPath $checklistPath -Raw -Encoding UTF8
    $heading = [regex]::Escape([string]$gate.phaseChecklistHeading)
    $sectionMatch = [regex]::Match($checklistContent, "(?ms)^$heading\s*$\r?\n(.*?)(?=^##\s|\z)")
    if (-not $sectionMatch.Success) {
        $failures.Add("활성 Phase 체크리스트 구역 누락: $($gate.activePhase)")
    } elseif ($sectionMatch.Groups[1].Value -match '(?m)^- \[ \]') {
        $failures.Add("활성 Phase 체크리스트 미완료: $($gate.activePhase)")
    }
}

if ($null -ne $gate) {
    foreach ($sourceHash in @($gate.sourceHashes)) {
        $sourcePath = Join-Path $root ([string]$sourceHash.path)
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            $failures.Add("원본 무결성 검사 대상 누락: $($sourceHash.path)")
            continue
        }
        $actualHash = (Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
        if ($actualHash -ne ([string]$sourceHash.sha256).ToUpperInvariant()) {
            $failures.Add("원본 SHA-256 불일치: $($sourceHash.path)")
        }
    }
}

if ($null -ne $gate -and $gate.activePhase -eq 'P1-02') {
    $inventoryPath = Join-Path $root '서식_인벤토리.md'
    $rawStructurePath = Join-Path $root '_workspace\02_hwpx\P0-02_구조.json'
    if (-not (Test-Path -LiteralPath $inventoryPath -PathType Leaf)) {
        $failures.Add('P1-02 인벤토리 문서 누락: 서식_인벤토리.md')
    } elseif (-not (Test-Path -LiteralPath $rawStructurePath -PathType Leaf)) {
        $failures.Add('P1-02 원시 구조 JSON 누락')
    } else {
        $inventoryContent = Get-Content -LiteralPath $inventoryPath -Raw -Encoding UTF8
        $formSection = [regex]::Match($inventoryContent, '(?ms)^## 2\. Form ID 목록\s*$\r?\n(.*?)(?=^##\s|\z)')
        if (-not $formSection.Success) {
            $failures.Add('P1-02 Form ID 목록 구역 누락')
        } else {
            $formIds = [regex]::Matches($formSection.Groups[1].Value, '(?m)^\|\s*(F-\d{3})\s*\|') | ForEach-Object { $_.Groups[1].Value }
            $expectedFormIds = 1..57 | ForEach-Object { 'F-{0:D3}' -f $_ }
            if (@($formIds).Count -ne 57 -or (@($formIds | Select-Object -Unique).Count -ne 57) -or ([string]::Join(',', $formIds) -ne [string]::Join(',', $expectedFormIds))) {
                $failures.Add('P1-02 Form ID가 F-001~F-057의 고유·연속 목록이 아님')
            }
        }

        $rawStructure = Get-Content -LiteralPath $rawStructurePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $rawCandidates = @($rawStructure.blocks | Where-Object { $_.type -eq 'paragraph' -and $_.text -match '^\s*\[\s*\d+(?:[-_]\d+)?\s*\]' }).Count
        if ($rawCandidates -ne 58) {
            $failures.Add("P1-02 원시 서식 표제 후보 수 불일치: $rawCandidates")
        }

        foreach ($requiredText in @(
            '`[9-13] 참고 단가 비율표 기준`',
            'Form ID `F-018`, `F-019`로 분리',
            '개인식별정보·연락처·서명·설문 원자료의 수집·저장·자동치환·로그 기록을 금지',
            '## 3. 담당·난이도 적용표',
            '`F-001`~`F-057`의 모든 Form ID에 빠짐없이 적용함'
        )) {
            if (-not $inventoryContent.Contains($requiredText)) {
                $failures.Add("P1-02 인벤토리 필수 근거 누락: $requiredText")
            }
        }
    }
}

if ($null -ne $gate -and $gate.activePhase -eq 'P1-03') {
    $dictionaryPath = Join-Path $root '입력데이터_사전.md'
    $mappingPath = Join-Path $root '서식_매핑표.md'
    foreach ($requiredPath in @($dictionaryPath, $mappingPath)) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            $failures.Add("P1-03 필수 문서 누락: $(Split-Path -Leaf $requiredPath)")
        }
    }

    if ((Test-Path -LiteralPath $dictionaryPath -PathType Leaf) -and (Test-Path -LiteralPath $mappingPath -PathType Leaf)) {
        $dictionaryContent = Get-Content -LiteralPath $dictionaryPath -Raw -Encoding UTF8
        $mappingContent = Get-Content -LiteralPath $mappingPath -Raw -Encoding UTF8
        $mappingSection = [regex]::Match($mappingContent, '(?ms)^## 2\. Form ID별 매핑\s*$\r?\n(.*?)(?=^##\s|\z)')
        if (-not $mappingSection.Success) {
            $failures.Add('P1-03 Form ID별 매핑 구역 누락')
        } else {
            $mappedFormIds = [regex]::Matches($mappingSection.Groups[1].Value, '(?m)^\|\s*(F-\d{3})\s*\|') | ForEach-Object { $_.Groups[1].Value }
            $expectedMappedFormIds = 1..52 | ForEach-Object { 'F-{0:D3}' -f $_ }
            if (@($mappedFormIds).Count -ne 52 -or (@($mappedFormIds | Select-Object -Unique).Count -ne 52) -or ([string]::Join(',', $mappedFormIds) -ne [string]::Join(',', $expectedMappedFormIds))) {
                $failures.Add('P1-03 Form ID 매핑이 F-001~F-052의 고유·연속 목록이 아님')
            }

            $expandFormIds = {
                param([string]$value)
                $expanded = [System.Collections.Generic.List[string]]::new()
                foreach ($rangeMatch in [regex]::Matches($value, 'F-(\d{3})(?:~F-(\d{3}))?')) {
                    $start = [int]$rangeMatch.Groups[1].Value
                    $end = if ($rangeMatch.Groups[2].Success) { [int]$rangeMatch.Groups[2].Value } else { $start }
                    for ($number = $start; $number -le $end; $number++) {
                        $expanded.Add(('F-{0:D3}' -f $number))
                    }
                }
                return @($expanded)
            }

            $calculationAllowances = @{}
            foreach ($dictionaryRow in [regex]::Matches($dictionaryContent, '(?m)^\|\s*(K-\d{2})\s*\|.*\|\s*([^|]+)\s*\|\s*$')) {
                $calculationAllowances[$dictionaryRow.Groups[1].Value] = @(& $expandFormIds $dictionaryRow.Groups[2].Value | Sort-Object -Unique)
            }
            $calculationMappings = @{}
            foreach ($mappingRow in [regex]::Matches($mappingSection.Groups[1].Value, '(?m)^\|\s*(F-\d{3})\s*\|[^|]*\|[^|]*\|\s*([^|]+)\|')) {
                $formId = $mappingRow.Groups[1].Value
                foreach ($calculationId in [regex]::Matches($mappingRow.Groups[2].Value, 'K-\d{2}') | ForEach-Object { $_.Value }) {
                    if (-not $calculationMappings.ContainsKey($calculationId)) {
                        $calculationMappings[$calculationId] = [System.Collections.Generic.List[string]]::new()
                    }
                    $calculationMappings[$calculationId].Add($formId)
                }
            }
            foreach ($calculationId in @('K-01', 'K-02', 'K-03', 'K-04')) {
                $allowed = @($calculationAllowances[$calculationId] | Sort-Object -Unique)
                $used = if ($calculationMappings.ContainsKey($calculationId)) { @($calculationMappings[$calculationId] | Sort-Object -Unique) } else { @() }
                if ([string]::Join(',', $allowed) -ne [string]::Join(',', $used)) {
                    $failures.Add("P1-03 계산 필드 허용 대상 불일치: $calculationId")
                }
            }

            $inventoryPath = Join-Path $root '서식_인벤토리.md'
            if (-not (Test-Path -LiteralPath $inventoryPath -PathType Leaf)) {
                $failures.Add('P1-03 개인정보 경계 인벤토리 문서 누락')
            } else {
                $inventoryContent = Get-Content -LiteralPath $inventoryPath -Raw -Encoding UTF8
                $protectedSection = [regex]::Match($inventoryContent, '(?ms)^## 6\. 개인정보 문서의 허용 입력 경계\s*$\r?\n(.*?)(?=^##\s|\z)')
                if (-not $protectedSection.Success) {
                    $failures.Add('P1-03 개인정보 문서 허용 입력 경계 구역 누락')
                }
                foreach ($protectedFormId in @('F-029', 'F-047', 'F-050')) {
                    $inventoryRow = [regex]::Match($protectedSection.Groups[1].Value, "(?m)^\|\s*$protectedFormId\s*\|[^|]*\|\s*([^|]+)\|\s*([^|]+)\|")
                    $mappingRow = [regex]::Match($mappingSection.Groups[1].Value, "(?m)^\|\s*$protectedFormId\s*\|\s*([^|]+)\|[^|]*\|[^|]*\|\s*금지:\s*([^;|]+);")
                    if (-not $inventoryRow.Success -or -not $mappingRow.Success) {
                        $failures.Add("P1-03 개인정보 문서 허용 필드 행 누락: $protectedFormId")
                        continue
                    }
                    $inventoryAllowed = @([regex]::Matches($inventoryRow.Groups[1].Value, '[CBRD]-\d{2}') | ForEach-Object { $_.Value } | Sort-Object -Unique)
                    $mappingAllowed = @([regex]::Matches($mappingRow.Groups[1].Value, '[CBRD]-\d{2}') | ForEach-Object { $_.Value } | Sort-Object -Unique)
                    if ([string]::Join(',', $inventoryAllowed) -ne [string]::Join(',', $mappingAllowed)) {
                        $failures.Add("P1-03 개인정보 문서 허용 필드 불일치: $protectedFormId")
                    }
                    $inventoryForbidden = $inventoryRow.Groups[2].Value.Trim()
                    $mappingForbidden = $mappingRow.Groups[2].Value.Trim()
                    if ($inventoryForbidden -ne $mappingForbidden) {
                        $failures.Add("P1-03 개인정보 문서 금지 필드 불일치: $protectedFormId")
                    }
                }
            }
        }

        foreach ($requiredText in @(
            '## 1. 데이터 영역',
            '## 2. 필드 정의',
            '## 3. 보호 필드 및 금지 처리',
            '## 4. 검증 순서',
            '개인식별·연락처·서명·치수·개별 응답은 입력·저장·치환·로그·Git·웹 전송 대상에서 제외',
            'F-053~F-057',
            '| F-053~F-057 | 입력·치환·출력 선택 대상에서 제외 |',
            '실제 HWPX 치환 위치·반복 표 구조는 P2 사본 POC에서만 확정함.',
            '| F-029 | C-01~C-02, C-05 |',
            '금지: 동의자 성명·주소·연락처·서명·식별번호; 빈 양식',
            '| F-047 | C-01~C-02, B-07, D-02 |',
            '금지: 학생·학부모 성명·연락처·신청 여부·치수·개별 수량; 빈 응답란',
            '| F-050 | C-01~C-02, D-02, D-05 |',
            '금지: 응답자 성명·연락처·개별 응답·자유서술; 빈 설문지'
        )) {
            if (-not ($dictionaryContent + "`n" + $mappingContent).Contains($requiredText)) {
                $failures.Add("P1-03 데이터 모델 필수 근거 누락: $requiredText")
            }
        }
    }
}

if ($null -ne $gate -and $gate.activePhase -eq 'P1-04') {
    $decisionPath = Join-Path $root '결정사항.md'
    if (-not (Test-Path -LiteralPath $decisionPath -PathType Leaf)) {
        $failures.Add('P1-04 결정 문서 누락: 결정사항.md')
    } else {
        $decisionContent = Get-Content -LiteralPath $decisionPath -Raw -Encoding UTF8
        foreach ($requiredText in @(
            'F-001~F-052, 총 52종',
            'F-053~F-057, 총 5종',
            '총 57종',
            'F-007 교복구매 구매 요청',
            'F-024 품목별 단가 비율표',
            'F-013 교복 디자인 및 규격서',
            '개인정보 가능 문서 F-029, F-047, F-050은 대표 POC에서 제외',
            'Kordoc `validate`, 생성본 재열기, 한컴 수동 열기, PDF 비교, 담당자 확인 기록'
        )) {
            if (-not $decisionContent.Contains($requiredText)) {
                $failures.Add("P1-04 결정 근거 누락: $requiredText")
            }
        }

        $approvalRows = [regex]::Matches($decisionContent, '(?m)^\|\s*(?:52종 출력 \+ 5종 읽기 전용 범위|F-007 단순 기안문 POC|F-024 반복·계산 문서 POC|F-013 표·이미지 복합 문서 POC|P2 한컴·PDF 수동 검토 책임)\s*\|[^|]*\|\s*(APPROVED)\s*\|')
        if ($approvalRows.Count -ne 5) {
            $failures.Add('P1-04 업무 담당 승인 기록이 5건 모두 APPROVED가 아님')
        }
    }
}

# PowerShell 스크립트는 실행 전에 구문 오류를 차단한다. PSScriptAnalyzer가 없는 환경에서도
# 표준 PowerShell 파서를 사용하므로 린트 공백이 완료 판정으로 새지 않는다.
foreach ($relativeScriptPath in @(
    'scripts\\verify_excel_v1.ps1',
    'scripts\\build_excel_v1_structure.ps1',
    'scripts\\build_excel_v1_vba.ps1'
)) {
    $scriptPath = Join-Path $root $relativeScriptPath
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        $failures.Add("P3 스크립트 누락: $relativeScriptPath")
        continue
    }
    $parseTokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$parseTokens, [ref]$parseErrors)
    if (@($parseErrors).Count -gt 0) {
        $failures.Add("PowerShell 구문 검사 실패: $relativeScriptPath — $($parseErrors[0].Message)")
    }
}

$scriptAnalyzer = Get-Module -ListAvailable -Name PSScriptAnalyzer | Select-Object -First 1
if ($null -eq $scriptAnalyzer) {
    $failures.Add('PSScriptAnalyzer가 설치되지 않아 PowerShell 린트를 수행할 수 없음')
} else {
    Import-Module PSScriptAnalyzer -ErrorAction Stop
    foreach ($relativeScriptPath in @(
        'scripts\\verify_excel_v1.ps1',
        'scripts\\build_excel_v1_structure.ps1',
        'scripts\\build_excel_v1_vba.ps1',
        '.claude\\hooks\\verify-phase.ps1'
    )) {
        $scriptPath = Join-Path $root $relativeScriptPath
        if (Test-Path -LiteralPath $scriptPath -PathType Leaf) {
            $lintFindings = @(Invoke-ScriptAnalyzer -Path $scriptPath -Severity Error,Warning)
            if ($lintFindings.Count -gt 0) {
                $failures.Add("PowerShell 린트 실패: $relativeScriptPath — $($lintFindings[0].RuleName) (line $($lintFindings[0].Line))")
            }
        }
    }
}

# P3는 Markdown 상태만 PASS/DONE으로 바꿔서는 완료할 수 없다. 실제 XLSM 검증 스크립트가
# 사본에서 Excel COM·PDF·수식·외부 의존성을 재검증해 0으로 종료해야 한다.
if ($null -ne $gate -and $gate.activePhase -eq 'P3-01') {
    $excelVerifier = Join-Path $root 'scripts\\verify_excel_v1.ps1'
    if (Test-Path -LiteralPath $excelVerifier -PathType Leaf) {
        $excelVerifierOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $excelVerifier 2>&1
        if ($LASTEXITCODE -ne 0) {
            $failures.Add("P3 실제 XLSM 검증 실패: $($excelVerifierOutput -join ' ')")
        }
        $excelVerifierLog = Join-Path $root '_workspace\\03_excel\\verify_v1_log.txt'
        if (-not (Test-Path -LiteralPath $excelVerifierLog -PathType Leaf) -or -not ((Get-Content -LiteralPath $excelVerifierLog -Raw -Encoding UTF8) -match 'PASS: Excel v1 검증 전체 통과')) {
            $failures.Add('P3 실제 XLSM 검증 로그에 최종 PASS 증빙이 없음')
        }
    }
}

$logPath = Join-Path $root '로그.md'
if (Test-Path -LiteralPath $logPath -PathType Leaf) {
    $logContent = Get-Content -LiteralPath $logPath -Raw
    if ($null -ne $gate) {
        $phase = [regex]::Escape([string]$gate.activePhase)
        if ($logContent -match "(?im)^\|[^\r\n]*\|\s*$phase(?:[-\s]|\|).*\|\s*(Critical|High)\s*\|.*\|\s*(OPEN|IN_PROGRESS|BLOCKED)\s*\|\s*$") {
            $failures.Add("활성 Phase 로그에 미해결 Critical 또는 High 결함이 남아 있음: $($gate.activePhase)")
        }
    }
}

Push-Location -LiteralPath $root
try {
    # Git의 CRLF 안내는 stderr로만 출력되고 검사 실패가 아니다. PowerShell의 Stop 설정이
    # 이를 NativeCommandError로 승격하지 않도록 종료 코드를 명시적으로 판정한다.
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $diffCheck = & git diff --check 2>&1
    $gitDiffExitCode = $LASTEXITCODE
    $ErrorActionPreference = $previousErrorActionPreference
    if ($gitDiffExitCode -ne 0) {
        $failures.Add("git diff --check 실패: $($diffCheck -join ' ')")
    }
} finally {
    Pop-Location
}

if ($failures.Count -gt 0) {
    [Console]::Error.WriteLine(('품질 게이트 미통과: ' + ($failures -join '; ') + '. 오류를 수정하고 검사 상태를 PASS로 갱신한 뒤 다시 검증해야 함.'))
    exit 2
}

exit 0
