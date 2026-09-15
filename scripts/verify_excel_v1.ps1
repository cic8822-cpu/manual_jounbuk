# P3-01 클린룸 재구현 — 검증: 외부 링크/연결/#REF! 0건, 매크로 실행, PDF 출력, A4 1쪽 재확인
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$targetPath = Join-Path $root 'artifacts\excel\교복구매_길라잡이_Excel_v1.xlsm'
$logPath = Join-Path $root '_workspace\03_excel\verify_v1_log.txt'
$validationDirectory = Join-Path $root 'artifacts\excel\_validation'
$validationPath = Join-Path $validationDirectory '교복구매_길라잡이_Excel_v1.xlsm'
$legacyValidationPath = Join-Path $root 'artifacts\excel\~검증용_교복구매_길라잡이_Excel_v1.xlsm'
$failures = [System.Collections.Generic.List[string]]::new()

if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
    throw "검증 대상 XLSM이 없습니다: $targetPath"
}

if (Test-Path -LiteralPath $logPath) { Remove-Item -LiteralPath $logPath -Force }
if (Test-Path -LiteralPath $legacyValidationPath) { Remove-Item -LiteralPath $legacyValidationPath -Force }
if (Test-Path -LiteralPath $validationDirectory) { Remove-Item -LiteralPath $validationDirectory -Recurse -Force }
New-Item -ItemType Directory -Path $validationDirectory -Force | Out-Null
function L($s) {
    [System.IO.File]::AppendAllText($logPath, "$s`r`n", [System.Text.UTF8Encoding]::new($false))
}
function Assert-Check([bool]$condition, [string]$message) {
    if ($condition) {
        L "PASS: $message"
    } else {
        L "FAIL: $message"
        $script:failures.Add($message)
    }
}
function Invoke-ValidationMacro([string]$macroName) {
    $excel.Run($macroName)
}

Copy-Item -LiteralPath $targetPath -Destination $validationPath -Force
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null
try {
    $wb = $excel.Workbooks.Open($validationPath, [Type]::Missing, $false)

    # 1. 외부 링크
    $links = $wb.LinkSources(1)  # xlExcelLinks
    if ($null -eq $links) {
        L "외부 링크(xlExcelLinks): 0건"
    } else {
        L "외부 링크(xlExcelLinks): $($links.Count)건 -- $($links -join '; ')"
    }
    Assert-Check ($null -eq $links -or $links.Count -eq 0) '외부 링크가 0건임'

    # 2. 외부 연결(Connections)
    $connCount = $wb.Connections.Count
    L "외부 연결(Connections): $connCount 건"
    Assert-Check ($connCount -eq 0) '외부 연결이 0건임'

    # 3. #REF! 이름정의
    $refCount = 0
    $refNames = @()
    foreach ($n in $wb.Names) {
        try {
            $r = $n.RefersTo
            if ($r -like "*#REF!*") { $refCount++; $refNames += $n.Name }
        } catch { $refCount++; $refNames += $n.Name }
    }
    L "#REF! 이름정의: $refCount 건 $(if($refNames.Count -gt 0){'-- ' + ($refNames -join ', ')})"
    Assert-Check ($refCount -eq 0) '#REF! 이름 정의가 0건임'

    # 4. 시트 목록
    L "시트 목록: $((@($wb.Worksheets) | ForEach-Object { $_.Name }) -join ', ')"
    $requiredSheets = @('사용설명서', '기초자료입력', 'DB', 'DB_품목', '서식선택_출력', 'F-007_구매요청기안문', 'F-024_단가비율표', '학교정보')
    $sheetNames = @($wb.Worksheets | ForEach-Object { $_.Name })
    Assert-Check ($sheetNames.Count -eq 8 -and @($requiredSheets | Where-Object { $_ -notin $sheetNames }).Count -eq 0) '필수 8개 시트가 모두 존재함'

    # 5. A4 1쪽 자연 충족 재확인 (F-007, F-024)
    foreach ($sn in @("F-007_구매요청기안문","F-024_단가비율표")) {
        $ws = $wb.Worksheets.Item($sn)
        $hb = $ws.HPageBreaks.Count
        $vb = $ws.VPageBreaks.Count
        $zoom = $ws.PageSetup.Zoom
        L "$sn : Zoom=$zoom, HPageBreaks=$hb, VPageBreaks=$vb => $(if($hb -eq 0 -and $vb -eq 0){'A4 1쪽 자연 충족 OK'}else{'FAIL - 1쪽 초과'})"
        Assert-Check ($zoom -eq 100 -and $hb -eq 0 -and $vb -eq 0) "$sn A4 1쪽 자연 배율 출력"
    }

    # 6. 매크로 실행 테스트 (마스킹 테스트값 사용 — 실제 업체/개인정보 아님)
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    $wsIn.Range("C4").Value2 = "테스트초등학교"
    $wsIn.Range("C5").Value2 = "2026"
    $wsIn.Range("C12").Value2 = "2026학년도 동복 구매"
    $wsIn.Range("C22").Value2 = "2026학년도 동복 학교주관구매 요청"
    $wsIn.Range("C21").Value2 = "교육지원청"
    $wsIn.Range("C24").Value2 = "구매계획서 1부"
    $wsIn.Range("B29").Value2 = "동복 상의"
    $wsIn.Range("C29").Value2 = 100
    $wsIn.Range("D29").Value2 = 50000
    $wsIn.Range("B30").Value2 = "동복 하의"
    $wsIn.Range("C30").Value2 = 100
    $wsIn.Range("D30").Value2 = 40000
    $excel.CalculateFullRebuild()

    try {
        Invoke-ValidationMacro -macroName '검증_저장하기'
        L "매크로 저장하기(): 정상 실행"
    } catch {
        $failures.Add("매크로 저장하기() 실패: $($_.Exception.Message)")
        L "FAIL: 매크로 저장하기() 실패: $($_.Exception.Message)"
    }

    $wsDB = $wb.Worksheets.Item("DB")
    $dbRow2A = $wsDB.Cells.Item(2,1).Value2
    $dbRow2B = $wsDB.Cells.Item(2,2).Value2
    L "DB 시트 2행 기록 확인: 순번=$dbRow2A, 학교명=$dbRow2B"
    Assert-Check ($dbRow2A -eq 1 -and $dbRow2B -eq '테스트초등학교') '저장하기()가 DB에 마스킹 테스트값을 기록함'
    $wsItems = $wb.Worksheets.Item('DB_품목')
    L "DB_품목 기록 확인: $($wsItems.Cells.Item(2, 1).Value2)/$($wsItems.Cells.Item(2, 3).Value2), $($wsItems.Cells.Item(3, 1).Value2)/$($wsItems.Cells.Item(3, 3).Value2)"
    Assert-Check ($wsItems.Cells.Item(2, 1).Value2 -eq 1 -and $wsItems.Cells.Item(2, 3).Value2 -eq '동복 상의' -and $wsItems.Cells.Item(3, 1).Value2 -eq 1 -and $wsItems.Cells.Item(3, 3).Value2 -eq '동복 하의') '저장하기()가 복수 품목 반복행을 DB_품목에 기록함'
    $wsIn.Range('B31:D31').ClearContents()
    $wsIn.Range('B31').Value2 = '불완전 품목'
    Assert-Check (-not [bool]$excel.Run('검증_품목행검증')) '수정 경로와 공유하는 품목 검증이 불완전 행을 거부함'
    Invoke-ValidationMacro -macroName '검증_저장하기'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsDB.Cells.Item(3, 1).Value2)) '불완전 품목 행은 DB 저장이 거부됨'
    $wsIn.Range('B31:D31').ClearContents()

    # F-007/F-024 수식이 HWPX 원문 대조 후 확정한 v2 배치에서 기초자료입력을 정상 참조하는지 확인
    $wsF7 = $wb.Worksheets.Item("F-007_구매요청기안문")
    L "F-007 제목 셀(C9) 계산값: $($wsF7.Range('C9').Value2)"
    Assert-Check ($wsF7.Range('C9').Value2 -eq '2026학년도 동복 학교주관구매 요청') 'F-007 공통 제목이 입력값을 참조함'
    $wsF24 = $wb.Worksheets.Item("F-024_단가비율표")
    L "F-024 수량 합계(D15) 계산값: $($wsF24.Range('D15').Value2), 비율(E9/E10): $($wsF24.Range('E9').Value2)/$($wsF24.Range('E10').Value2)"
    Assert-Check ($wsF24.Range('D15').Value2 -eq 204 -and $wsF24.Range('E9').Text -eq '55.6%' -and $wsF24.Range('E10').Text -eq '44.4%') 'F-024 수량 합계와 단가비율이 입력값을 참조함'
    Assert-Check ($wsF24.Range('C9').Value2 -eq '동복 상의' -and $wsF24.Range('C10').Value2 -eq '동복 하의') 'F-024 품목명이 입력 반복행을 참조함'

    # 7. 서식선택_출력 체크 + PDF 내보내기 매크로 테스트
    $wsSel = $wb.Worksheets.Item("서식선택_출력")
    # F-007, F-024 행 찾기
    $lastRow = $wsSel.Cells.Item($wsSel.Rows.Count, 2).End(-4162).Row  # xlUp
    for ($r = 5; $r -le $lastRow; $r++) {
        $fid = $wsSel.Cells.Item($r, 2).Value2
        if ($fid -eq "F-007" -or $fid -eq "F-024") {
            $wsSel.Cells.Item($r, 1).Value2 = $true
        }
    }
    $pdfStartTime = Get-Date
    try {
        Invoke-ValidationMacro -macroName '검증_선택서식_PDF저장'
        L "매크로 선택서식_PDF저장(): 정상 실행"
    } catch {
        $failures.Add("매크로 선택서식_PDF저장() 실패: $($_.Exception.Message)")
        L "FAIL: 매크로 선택서식_PDF저장() 실패: $($_.Exception.Message)"
    }

    $outputDir = Join-Path $validationDirectory 'output'
    if (Test-Path $outputDir) {
        $pdfs = Get-ChildItem $outputDir -Filter "*.pdf" | Where-Object { $_.LastWriteTime -ge $pdfStartTime } | Sort-Object LastWriteTime -Descending
        if ($pdfs.Count -gt 0) {
            L "PDF 생성 확인: $($pdfs[0].FullName) ($([math]::Round($pdfs[0].Length/1024,1)) KB)"
            Assert-Check ($pdfs[0].Length -gt 0) '선택 서식 PDF가 비어 있지 않음'
        } else {
            $failures.Add('PDF 생성 실패: output 폴더에 PDF 없음')
            L 'FAIL: PDF 생성 실패: output 폴더에 PDF 없음'
        }
    } else {
        $failures.Add('PDF 생성 실패: output 폴더 자체가 없음')
        L 'FAIL: PDF 생성 실패: output 폴더 자체가 없음'
    }

    # 8. 매크로 실행 후 테스트값 원복(초기화) — 배포본에 테스트 데이터가 남지 않도록
    Invoke-ValidationMacro -macroName '검증_초기화'
    $wsDB2 = $wb.Worksheets.Item("DB")
    $wsDB2.Range("A2:T2").ClearContents()
    $wb.Worksheets.Item('DB_품목').Range('A2:F100').ClearContents()
    $missingInputAccepted = [bool]$excel.Run('검증_필수값검증')
    Assert-Check (-not $missingInputAccepted) '필수값이 비어 있으면 저장 검증이 거부됨'
    L "테스트 데이터 정리(기초자료입력 초기화, DB 2행 삭제) 완료"

    L '테스트 데이터 정리 완료(검증 사본만 변경, 배포본 저장 없음)'

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $validationDirectory) { Remove-Item -LiteralPath $validationDirectory -Recurse -Force }
}

if ($failures.Count -gt 0) {
    throw ('Excel v1 검증 실패: ' + ($failures -join '; '))
}

L 'PASS: Excel v1 검증 전체 통과'
