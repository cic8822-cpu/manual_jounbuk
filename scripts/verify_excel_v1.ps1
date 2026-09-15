# P3-01 클린룸 재구현 — 검증: 외부 링크/연결/#REF! 0건, 매크로 실행, PDF 출력, A4 1쪽 재확인
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$targetPath = Join-Path $root 'artifacts\excel\교복구매_길라잡이_Excel_v1.xlsm'
$logPath = Join-Path $root '_workspace\03_excel\verify_v1_log.txt'

if (Test-Path $logPath) { Remove-Item $logPath }
function L($s) {
    [System.IO.File]::AppendAllText($logPath, "$s`r`n", [System.Text.UTF8Encoding]::new($false))
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null
try {
    $wb = $excel.Workbooks.Open($targetPath, [Type]::Missing, $false)

    # 1. 외부 링크
    $links = $wb.LinkSources(1)  # xlExcelLinks
    if ($links -eq $null) {
        L "외부 링크(xlExcelLinks): 0건"
    } else {
        L "외부 링크(xlExcelLinks): $($links.Count)건 -- $($links -join '; ')"
    }

    # 2. 외부 연결(Connections)
    $connCount = $wb.Connections.Count
    L "외부 연결(Connections): $connCount 건"

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

    # 4. 시트 목록
    L "시트 목록: $((@($wb.Worksheets) | ForEach-Object { $_.Name }) -join ', ')"

    # 5. A4 1쪽 자연 충족 재확인 (F-007, F-024)
    foreach ($sn in @("F-007_구매요청기안문","F-024_단가비율표")) {
        $ws = $wb.Worksheets.Item($sn)
        $hb = $ws.HPageBreaks.Count
        $vb = $ws.VPageBreaks.Count
        $zoom = $ws.PageSetup.Zoom
        L "$sn : Zoom=$zoom, HPageBreaks=$hb, VPageBreaks=$vb => $(if($hb -eq 0 -and $vb -eq 0){'A4 1쪽 자연 충족 OK'}else{'FAIL - 1쪽 초과'})"
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

    try {
        $excel.SendKeys("~", $false)
        $excel.Run("저장하기")
        L "매크로 저장하기(): 정상 실행"
    } catch {
        L "매크로 저장하기() 실패: $($_.Exception.Message)"
    }

    $wsDB = $wb.Worksheets.Item("DB")
    $dbRow2A = $wsDB.Cells.Item(2,1).Value2
    $dbRow2B = $wsDB.Cells.Item(2,2).Value2
    L "DB 시트 2행 기록 확인: 순번=$dbRow2A, 학교명=$dbRow2B"

    # F-007/F-024 수식이 기초자료입력을 정상 참조하는지 값 확인
    $wsF7 = $wb.Worksheets.Item("F-007_구매요청기안문")
    L "F-007 제목 셀(C8) 계산값: $($wsF7.Range('C8').Value2)"
    $wsF24 = $wb.Worksheets.Item("F-024_단가비율표")
    L "F-024 합계(F18) 계산값: $($wsF24.Range('F18').Value2)"

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
    try {
        $excel.SendKeys("~", $false)
        $excel.SendKeys("~", $false)
        $excel.Run("선택서식_PDF저장")
        L "매크로 선택서식_PDF저장(): 정상 실행"
    } catch {
        L "매크로 선택서식_PDF저장() 실패: $($_.Exception.Message)"
    }

    $outputDir = Join-Path $root 'artifacts\excel\output'
    if (Test-Path $outputDir) {
        $pdfs = Get-ChildItem $outputDir -Filter "*.pdf" | Sort-Object LastWriteTime -Descending
        if ($pdfs.Count -gt 0) {
            L "PDF 생성 확인: $($pdfs[0].FullName) ($([math]::Round($pdfs[0].Length/1024,1)) KB)"
        } else {
            L "PDF 생성 실패: output 폴더에 PDF 없음"
        }
    } else {
        L "PDF 생성 실패: output 폴더 자체가 없음"
    }

    # 8. 매크로 실행 후 테스트값 원복(초기화) — 배포본에 테스트 데이터가 남지 않도록
    $excel.SendKeys("~", $false)
    $excel.Run("초기화")
    $wsDB2 = $wb.Worksheets.Item("DB")
    $wsDB2.Range("A2:T2").ClearContents()
    L "테스트 데이터 정리(기초자료입력 초기화, DB 2행 삭제) 완료"

    $wb.Save()
    L "검증 후 저장 완료(테스트 데이터 제거 상태로 저장)"

} finally {
    if ($wb) { $wb.Close($true) }
    $excel.Quit()
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

[System.IO.File]::WriteAllText($logPath, $log.ToString(), [System.Text.UTF8Encoding]::new($false))
