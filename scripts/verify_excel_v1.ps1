# P3-01 클린룸 재구현 — 검증: 외부 링크/연결/#REF! 0건, 매크로 실행, PDF 출력, A4 1쪽 재확인
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$defaultTargetPath = Join-Path $root 'artifacts\excel\build.xlsm'
$targetPath = if ([string]::IsNullOrWhiteSpace($env:UNIFORM_EXCEL_BUILD_PATH)) { $defaultTargetPath } else { $env:UNIFORM_EXCEL_BUILD_PATH }
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
    $requiredSheets = @('사용설명서', '기초자료입력', 'DB', 'DB_품목', 'DB_업체', '서식선택_출력', 'F-007_구매요청기안문', 'F-024_단가비율표', '학교정보', '학교검색', '절차안내', '계약방법안내')
    $sheetNames = @($wb.Worksheets | ForEach-Object { $_.Name })
    Assert-Check ($sheetNames.Count -eq 12 -and @($requiredSheets | Where-Object { $_ -notin $sheetNames }).Count -eq 0) '필수 12개 시트가 모두 존재함'

    # 4a. 학교검색 및 절차안내 정적 구조
    $wsSchool = $wb.Worksheets.Item('학교정보')
    Assert-Check ($wsSchool.Range('A123').Value2 -eq '학교정보 표본 행 수' -and $wsSchool.Range('B123').Text -eq '120') '학교정보 공개 기관정보 표본이 정확히 120행임'
    $wsSearch = $wb.Worksheets.Item('학교검색')
    Assert-Check ($wsSearch.Range('A1').Value2 -match '학교정보 검색' -and $wsSearch.Range('B3,D3,F3').Interior.Color -eq 16777164) '학교명·지역·급별 검색 조건 UI가 존재함'
    $wsFlow = $wb.Worksheets.Item('절차안내')
    Assert-Check ($wsFlow.Range('A5:A13').Count -eq 9 -and $wsFlow.Range('D5').Value2 -eq 'F-001' -and $wsFlow.Range('D13').Value2 -eq 'F-044') '교복구매 9단계와 대표 Form ID 이동값이 존재함'
    $wsMethod = $wb.Worksheets.Item('계약방법안내')
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B5').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B6').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B7').Value2) -and $wsMethod.Buttons('btn계약방법반영').OnAction -eq '계약방법안내_기초자료반영') '계약방법안내가 빈 확인값으로 시작하며 반영 버튼이 존재함'

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
    $wsIn.Range("B44").Value2 = "검증업체가"
    $wsIn.Range("B45").Value2 = "검증업체나"
    $excel.CalculateFullRebuild()

    $wsIn.Range('C14').Value2 = '수동입력값'
    Invoke-ValidationMacro -macroName '검증_계약방법안내_기초자료반영'
    Assert-Check ($wsIn.Range('C14').Value2 -eq '수동입력값') '계약방법안내가 확인값이 비어 있으면 B-03을 덮어쓰지 않음'
    $wsMethod.Range('B5').Value2 = '아니오'
    $wsMethod.Range('B6').Value2 = '예'
    Invoke-ValidationMacro -macroName '검증_계약방법안내_기초자료반영'
    Assert-Check ($wsIn.Range('C14').Value2 -eq '수동입력값') '계약방법안내가 확인값 미충족 시 B-03을 덮어쓰지 않음'
    $wsMethod.Range('B5:B6').Value2 = '예'
    Invoke-ValidationMacro -macroName '검증_계약방법안내_기초자료반영'
    Assert-Check ($wsIn.Range('C14').Value2 -eq '2단계 입찰(규격·가격 동시)') '계약방법안내가 두 확인값을 명시적으로 충족 시 B-03에 예시를 반영함'

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
    $wsVendors = $wb.Worksheets.Item('DB_업체')
    L "DB_업체 기록 확인: $($wsVendors.Cells.Item(2, 1).Value2)/$($wsVendors.Cells.Item(2, 3).Value2), $($wsVendors.Cells.Item(3, 1).Value2)/$($wsVendors.Cells.Item(3, 3).Value2)"
    Assert-Check ($wsVendors.Cells.Item(2, 1).Value2 -eq 1 -and $wsVendors.Cells.Item(2, 2).Value2 -eq 1 -and $wsVendors.Cells.Item(2, 3).Value2 -eq '검증업체가' -and $wsVendors.Cells.Item(3, 1).Value2 -eq 1 -and $wsVendors.Cells.Item(3, 2).Value2 -eq 2 -and $wsVendors.Cells.Item(3, 3).Value2 -eq '검증업체나') '저장하기()가 복수 업체 반복행을 DB_업체에 기록함'
    $wsIn.Range('B44:B53').ClearContents()
    Invoke-ValidationMacro -macroName '검증_첫레코드불러오기'
    Assert-Check ($wsIn.Range('B44').Value2 -eq '검증업체가' -and $wsIn.Range('B45').Value2 -eq '검증업체나') '불러오기()가 업체 반복행을 원래 행에 복원함'
    $wsIn.Range('B45').ClearContents()
    $wsIn.Range('B47').Value2 = '검증업체다'
    Invoke-ValidationMacro -macroName '검증_수정하기'
    $wsIn.Range('B44:B53').ClearContents()
    Invoke-ValidationMacro -macroName '검증_첫레코드불러오기'
    Assert-Check ($wsIn.Range('B44').Value2 -eq '검증업체가' -and [string]::IsNullOrWhiteSpace([string]$wsIn.Range('B45').Value2) -and $wsIn.Range('B47').Value2 -eq '검증업체다' -and $wsVendors.Cells.Item(3, 2).Value2 -eq 4) '수정하기()가 기존 업체 반복행을 교체하고 저장 행번호를 보존함'
    $wsIn.Range('B46').Value2 = '검증업체가'
    Assert-Check ([bool]$excel.Run('검증_업체행검증') -and [bool]$excel.Run('검증_업체중복경고')) '중복 업체명은 저장 전 경고 신호를 반환함'
    $wsIn.Range('B46').Value2 = 'X'
    Assert-Check (-not [bool]$excel.Run('검증_업체행검증')) '1자 업체명 입력을 거부함'
    Invoke-ValidationMacro -macroName '검증_저장하기'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsDB.Cells.Item(3, 1).Value2)) '형식 오류 업체명은 DB 저장이 거부됨'
    $wsIn.Range('B46').Value2 = ('가' * 151)
    Assert-Check (-not [bool]$excel.Run('검증_업체행검증')) '151자 이상 업체명 입력을 거부함'
    $wsIn.Range('B46').ClearContents()
    foreach ($forbiddenVendorValue in @('010-1234-5678', '010.1234.5678', '123-45-67890', '123 45 67890', '대표자 홍길동')) {
        $wsIn.Range('B46').Value2 = $forbiddenVendorValue
        Assert-Check (-not [bool]$excel.Run('검증_업체행검증')) "업체명 입력란이 금지된 식별·연락처 표기($forbiddenVendorValue)를 거부함"
        Invoke-ValidationMacro -macroName '검증_저장하기'
        Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsDB.Cells.Item(3, 1).Value2)) "금지된 식별·연락처 표기($forbiddenVendorValue)가 DB에 저장되지 않음"
    }
    $wsIn.Range('B46').ClearContents()
    $wsIn.Range('B31:D31').ClearContents()
    $wsIn.Range('C31').Value2 = 1
    Assert-Check (-not [bool]$excel.Run('검증_품목행검증')) '품목명 없는 수량·단가 고아 입력행을 거부함'
    $wsIn.Range('B31:D31').ClearContents()
    $wsIn.Range('B31').Value2 = '불완전 품목'
    Assert-Check (-not [bool]$excel.Run('검증_품목행검증')) '수정 경로와 공유하는 품목 검증이 불완전 행을 거부함'
    Invoke-ValidationMacro -macroName '검증_저장하기'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsDB.Cells.Item(3, 1).Value2)) '불완전 품목 행은 DB 저장이 거부됨'
    $wsIn.Range('B31:D31').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_품목행검증') -and $wsIn.Application.WorksheetFunction.CountA($wsIn.Range('B29:B38')) -ge 1 -and $wsIn.Application.WorksheetFunction.CountA($wsIn.Range('B29:B38')) -le 6) 'F-024 출력용 품목 입력이 완전하고 1~6개임'
    Assert-Check ([bool]$excel.Run('검증_F024출력가능') -eq $true) 'F-024 선택 출력이 정상 품목 1~6개를 허용함'
    for ($r = 31; $r -le 35; $r++) {
        $wsIn.Cells.Item($r, 2).Value2 = "추가 품목 $r"
        $wsIn.Cells.Item($r, 3).Value2 = 1
        $wsIn.Cells.Item($r, 4).Value2 = 1000
    }
    Assert-Check ([bool]$excel.Run('검증_F024출력가능') -eq $false) 'F-024 선택 출력이 완전한 7개 품목을 거부함'
    $wsIn.Range('B31:D35').ClearContents()

    # F-007/F-024 수식이 HWPX 원문 대조 후 확정한 v2 배치에서 기초자료입력을 정상 참조하는지 확인
    $wsF7 = $wb.Worksheets.Item("F-007_구매요청기안문")
    L "F-007 제목 셀(C9) 계산값: $($wsF7.Range('C9').Value2)"
    Assert-Check ($wsF7.Range('C9').Value2 -eq '2026학년도 동복 학교주관구매 요청') 'F-007 공통 제목이 입력값을 참조함'
    $wsF24 = $wb.Worksheets.Item("F-024_단가비율표")
    L "F-024 수량 합계(D15) 계산값: $($wsF24.Range('D15').Value2), 비율(E9/E10): $($wsF24.Range('E9').Value2)/$($wsF24.Range('E10').Value2)"
    Assert-Check ($wsF24.Range('D15').Value2 -eq 204 -and $wsF24.Range('E9').Text -eq '55.6%' -and $wsF24.Range('E10').Text -eq '44.4%') 'F-024 수량 합계와 단가비율이 입력값을 참조함'
    Assert-Check ($wsF24.Range('C9').Value2 -eq '동복 상의' -and $wsF24.Range('C10').Value2 -eq '동복 하의') 'F-024 품목명이 입력 반복행을 참조함'

    # 6a. 학교명·지역·급별 검색 및 선택 결과 C4 반영 (공개 기관정보 표본, 검증 사본만 사용)
    # 원본 표본의 외부 쿼리 캐시에 검색 열 값이 비어 있을 수 있으므로, 검증 사본에만 마스킹 기관정보를 넣어
    # 학교명·지역·급별 조건의 복합 검색과 C4 반영 흐름을 독립적으로 재현함.
    $wsSchool.Cells.Item(2, 2).Value2 = '검증지역'
    $wsSchool.Cells.Item(2, 3).Value2 = '초등'
    $wsSchool.Cells.Item(2, 5).Value2 = '검증학교'
    $wsSearch.Range('B3').Value2 = '검증학교'
    $wsSearch.Range('D3').Value2 = '검증지역'
    $wsSearch.Range('F3').Value2 = '초등'
    Invoke-ValidationMacro -macroName '검증_학교검색_실행'
    Assert-Check ($wsSearch.Cells.Item(7, 4).Value2 -eq '검증학교') '학교명·지역·급별 복합 검색 결과가 생성됨'
    $wsSearch.Activate()
    $wsSearch.Cells.Item(7, 1).Select()
    Invoke-ValidationMacro -macroName '검증_선택학교_기초자료반영'
    Assert-Check ($wsIn.Range('C4').Value2 -eq '검증학교') '선택 학교명이 기초자료입력!C4에 반영됨'

    # 6b. 9단계 절차 안내에서 단계별 대표 Form ID로 이동
    $wsFlow.Activate()
    $wsFlow.Cells.Item(7, 1).Select()  # 3단계: F-007
    Invoke-ValidationMacro -macroName '검증_관련FormID로이동'
    Assert-Check ($excel.ActiveSheet.Name -eq '서식선택_출력' -and $excel.ActiveCell.Value2 -eq 'F-007') '3단계 절차안내가 F-007 선택 행으로 이동함'

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
    $wb.Worksheets.Item('DB_업체').Range('A2:C100').ClearContents()
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
