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
        $match = [regex]::Match($taskContent, "(?im)^\|\s*$escapedId\s*\|.*\|\s*(DONE|TODO|IN_PROGRESS|BLOCKED)\s*\|.*\|\s*$")
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
        $actualHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
        if ($actualHash -ne ([string]$sourceHash.sha256).ToUpperInvariant()) {
            $failures.Add("원본 SHA-256 불일치: $($sourceHash.path)")
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
    $diffCheck = & git diff --check 2>&1
    if ($LASTEXITCODE -ne 0) {
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
