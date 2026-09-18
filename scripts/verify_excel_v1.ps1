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

# Phase 훅은 모든 Write/Edit마다 이 스크립트를 다시 실행하므로, 직전 실행이 정리한
# EXCEL.EXE의 파일 잠금이 완전히 풀리기 전에 다음 실행이 곧바로 이어질 수 있음.
# 잠금 해제를 잠깐 기다리는 재시도로 이 경합 때문에 최상단에서 통째로 실패하는 것을 막는다.
function RemovePathWithRetry([string]$path, [int]$maxAttempts = 5, [int]$delayMs = 500) {
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop }
            return
        } catch {
            if ($attempt -eq $maxAttempts) { throw }
            Start-Sleep -Milliseconds $delayMs
        }
    }
}

RemovePathWithRetry -path $logPath
RemovePathWithRetry -path $legacyValidationPath
RemovePathWithRetry -path $validationDirectory
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
# Quit()·ReleaseComObject·GC만으로는 남은 스크립트 지역변수가 COM 참조를 계속
# 살려 두어 EXCEL.EXE가 좀비로 남을 수 있음. 생성 전후 프로세스 목록을 비교해
# 새로 뜬 PID를 기록해 두고, finally에서 종료가 확인되지 않으면 이 PID만
# 강제 종료해 고아 프로세스가 다음 실행을 막지 않게 함.
$excelPidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel = New-Object -ComObject Excel.Application
Start-Sleep -Milliseconds 300
$excelProcessId = Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $excelPidsBefore -notcontains $_.Id } |
    Select-Object -First 1 -ExpandProperty Id
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
    $requiredSheets = @('사용설명서', '기초자료입력', 'DB', 'DB_품목', 'DB_업체', 'DB_위원', 'DB_평가', 'DB_정량평가', 'DB_정성평가', 'DB_자기평점', '서식선택_출력', 'F-001_위원회구성기안', 'F-002_위원수락확인서', 'F-003_위원청렴보안서약서', 'F-004_구매추진계획수립기안문', 'F-005_구매추진계획안', 'F-006_운영위원회심의안', 'F-007_구매요청기안문', 'F-014_평가항목배점기준', 'F-015_정량적평가', 'F-016_정성적평가', 'F-017_제출서류자기확인서', 'F-018_정량적평가자기평점표', 'F-019_입찰참가신청서', 'F-020_입찰참가신고서', 'F-021_교복납품제안서', 'F-022_교복납품실적표', 'F-023_교복제조사양서', 'F-024_단가비율표', 'F-025_교복AS계획서', 'F-041_낙찰자결정', 'F-042_낙찰자결정통보', 'F-043_계약체결', '학교정보', '학교검색', '절차안내', '계약방법안내')
    $sheetNames = @($wb.Worksheets | ForEach-Object { $_.Name })
    Assert-Check ($sheetNames.Count -eq 37 -and @($requiredSheets | Where-Object { $_ -notin $sheetNames }).Count -eq 0) 'F-041~F-043을 포함한 필수 37개 시트가 모두 존재함'
    $wsProtectedDB = $wb.Worksheets.Item('DB')
    $wsProtectedItems = $wb.Worksheets.Item('DB_품목')
    $wsProtectedVendors = $wb.Worksheets.Item('DB_업체')
    $wsProtectedCommittee = $wb.Worksheets.Item('DB_위원')
    $wsProtectedScore = $wb.Worksheets.Item('DB_평가')
    $wsProtectedQuant = $wb.Worksheets.Item('DB_정량평가')
    $wsProtectedQual = $wb.Worksheets.Item('DB_정성평가')
    $wsProtectedSelf = $wb.Worksheets.Item('DB_자기평점')
    Assert-Check ($wsProtectedDB.ProtectContents -and $wsProtectedItems.ProtectContents -and $wsProtectedVendors.ProtectContents -and $wsProtectedCommittee.ProtectContents -and $wsProtectedScore.ProtectContents -and $wsProtectedQuant.ProtectContents -and $wsProtectedQual.ProtectContents -and $wsProtectedSelf.ProtectContents -and $wsProtectedDB.Visible -eq 0 -and $wsProtectedItems.Visible -eq 2 -and $wsProtectedVendors.Visible -eq 2 -and $wsProtectedCommittee.Visible -eq 2 -and $wsProtectedScore.Visible -eq 2 -and $wsProtectedQuant.Visible -eq 2 -and $wsProtectedQual.Visible -eq 2 -and $wsProtectedSelf.Visible -eq 2) 'DB는 숨김·보호되고 반복 DB는 VeryHidden·보호됨'
    L 'DB 숨김·무암호 보호는 우발적 편집 방지 상태로만 확인함. 직접 DB 편집의 영구 저장 차단은 아래 Workbook_BeforeSave 저장 경계 시나리오에서 별도로 검증함.'
    # ProtectionMode=True여야 매크로가 UserInterfaceOnly로 쓸 수 있음(=False면 매크로도 막혀 저장 매크로가
    # 보호된 시트 오류로 멈춤). UserInterfaceOnly는 파일을 다시 열면 유지되지 않아 Workbook_Open이 재적용하며,
    # 새 DB 시트를 추가할 때 그 재적용 목록에 빠뜨리면 이 검사가 잡아낸다(F-015 DB_정량평가 무한 대기 재발 방지).
    Assert-Check ($wsProtectedDB.ProtectionMode -and $wsProtectedItems.ProtectionMode -and $wsProtectedVendors.ProtectionMode -and $wsProtectedCommittee.ProtectionMode -and $wsProtectedScore.ProtectionMode -and $wsProtectedQuant.ProtectionMode -and $wsProtectedQual.ProtectionMode -and $wsProtectedSelf.ProtectionMode) 'Workbook_Open이 재열기 후 모든 DB 시트의 UserInterfaceOnly 보호를 재적용함'

    # 4a. 학교검색 및 절차안내 정적 구조
    $wsSchool = $wb.Worksheets.Item('학교정보')
    Assert-Check ($wsSchool.Range('A123').Value2 -eq '학교정보 표본 행 수' -and $wsSchool.Range('B123').Text -eq '120') '학교정보 공개 기관정보 표본이 정확히 120행임'
    $wsSearch = $wb.Worksheets.Item('학교검색')
    Assert-Check ($wsSearch.Range('A1').Value2 -match '학교정보 검색' -and $wsSearch.Range('B3,D3,F3').Interior.Color -eq 16777164) '학교명·지역·급별 검색 조건 UI가 존재함'
    $wsFlow = $wb.Worksheets.Item('절차안내')
    Assert-Check ($wsFlow.Range('A5:A13').Count -eq 9 -and $wsFlow.Range('D5').Value2 -eq 'F-001' -and $wsFlow.Range('D13').Value2 -eq 'F-044') '교복구매 9단계와 대표 Form ID 이동값이 존재함'
    $wsMethod = $wb.Worksheets.Item('계약방법안내')
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B5').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B6').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsMethod.Range('B7').Value2) -and $wsMethod.Buttons('btn계약방법반영').OnAction -eq '계약방법안내_기초자료반영') '계약방법안내가 빈 확인값으로 시작하며 반영 버튼이 존재함'

    # 5. A4 1쪽 자연 충족 재확인 (F-007, F-014, F-024)
    foreach ($sn in @("F-001_위원회구성기안","F-002_위원수락확인서","F-003_위원청렴보안서약서","F-004_구매추진계획수립기안문","F-006_운영위원회심의안","F-007_구매요청기안문","F-014_평가항목배점기준","F-015_정량적평가","F-016_정성적평가","F-017_제출서류자기확인서","F-018_정량적평가자기평점표","F-019_입찰참가신청서","F-020_입찰참가신고서","F-021_교복납품제안서","F-022_교복납품실적표","F-023_교복제조사양서","F-024_단가비율표","F-025_교복AS계획서","F-041_낙찰자결정","F-042_낙찰자결정통보","F-043_계약체결")) {
        $ws = $wb.Worksheets.Item($sn)
        $hb = $ws.HPageBreaks.Count
        $vb = $ws.VPageBreaks.Count
        $zoom = $ws.PageSetup.Zoom
        L "$sn : Zoom=$zoom, HPageBreaks=$hb, VPageBreaks=$vb => $(if($hb -eq 0 -and $vb -eq 0){'A4 1쪽 자연 충족 OK'}else{'FAIL - 1쪽 초과'})"
        Assert-Check ($zoom -eq 100 -and $hb -eq 0 -and $vb -eq 0) "$sn A4 1쪽 자연 배율 출력"
    }

    # 5a. F-005는 원문이 6개 물리 페이지 분량이라 A4 1쪽 원칙의 예외로 문서화됨(F-005_구현검증.md).
    # 강제 축소(FitToPages)만 없으면 되고, 가로 폭 초과(VPageBreaks)만 결함으로 간주하며 세로 다중
    # 페이지(HPageBreaks>0)는 정상으로 허용한다.
    $wsF05chk = $wb.Worksheets.Item('F-005_구매추진계획안')
    $vbF05chk = $wsF05chk.VPageBreaks.Count
    $zoomF05chk = $wsF05chk.PageSetup.Zoom
    L "F-005_구매추진계획안 : Zoom=$zoomF05chk, HPageBreaks=$($wsF05chk.HPageBreaks.Count)(다중 페이지 허용, 예외), VPageBreaks=$vbF05chk"
    Assert-Check ($zoomF05chk -eq 100 -and $vbF05chk -eq 0 -and -not $wsF05chk.PageSetup.FitToPagesWide -and -not $wsF05chk.PageSetup.FitToPagesTall) 'F-005 자연 배율(강제 축소 없음)이며 가로 폭 초과가 없음(세로 다중 페이지는 원문이 6쪽이라 허용된 예외)'

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
    $wsIn.Range("C15").Value2 = 5000000
    $wsIn.Range("C16").Value2 = 100
    $wsIn.Range("C17").Formula = "=DATE(2026,3,1)"
    $wsIn.Range("C18").Formula = "=DATE(2026,2,20)+18/24"
    $wsIn.Range("B44").Value2 = "검증업체가"
    $wsIn.Range("C44").Value2 = 10
    $wsIn.Range("D44").Value2 = 10
    $wsIn.Range("E44").Value2 = 15
    $wsIn.Range("F44").Value2 = 15
    $wsIn.Range("B45").Value2 = "검증업체나"
    $wsIn.Range("C45").Value2 = 6
    $wsIn.Range("D45").Value2 = 8
    $wsIn.Range("E45").Value2 = 9
    $wsIn.Range("F45").Value2 = 0
    $wsIn.Range("N44").Value2 = 8
    $wsIn.Range("O44").Value2 = 8
    $wsIn.Range("P44").Value2 = 12
    $wsIn.Range("Q44").Value2 = 15
    $wsIn.Range("N45").Value2 = 10
    $wsIn.Range("O45").Value2 = 5
    $wsIn.Range("P45").Value2 = 5
    $wsIn.Range("Q45").Value2 = 0
    $wsIn.Range("H44").Value2 = 15
    $wsIn.Range("I44").Value2 = 10
    $wsIn.Range("J44").Value2 = 15
    $wsIn.Range("K44").Value2 = 10
    $wsIn.Range("L44").Value2 = 5
    $wsIn.Range("H45").Value2 = 9
    $wsIn.Range("I45").Value2 = 6
    $wsIn.Range("J45").Value2 = 12
    $wsIn.Range("K45").Value2 = 8
    $wsIn.Range("L45").Value2 = -2
    $wsIn.Range("B58").Value2 = "교원위원"
    $wsIn.Range("C58").Value2 = "위원 ○○"
    $wsIn.Range("B59").Value2 = "학부모위원"
    $wsIn.Range("C59").Value2 = "위원 **"
    $wsIn.Range("B72").Value2 = "제품 품질"
    $wsIn.Range("C72").Value2 = 40
    $wsIn.Range("D72").Value2 = 38.5
    $wsIn.Range("B73").Value2 = "사후 관리"
    $wsIn.Range("C73").Value2 = 60
    $wsIn.Range("D73").Value2 = 60
    $excel.CalculateFullRebuild()

    $wsF14 = $wb.Worksheets.Item('F-014_평가항목배점기준')
    Assert-Check ($wsF14.Range('B3').Value2 -eq '[붙임 3] 제안서 평가항목 및 배점기준' -and $wsF14.Range('B9').Value2 -eq '구분' -and $wsF14.Range('C9').Value2 -eq '평가항목' -and $wsF14.Range('D9').Value2 -eq '배점기준' -and $wsF14.Range('E9').Value2 -eq '배점' -and $wsF14.Range('F9').Value2 -eq '평가점수') 'F-014 원문 대조 제목과 평가기준 표 머리글이 존재함'
    Assert-Check ($wsF14.Range('C10').Value2 -eq '제품 품질' -and $wsF14.Range('E10').Value2 -eq 40 -and $wsF14.Range('F10').Value2 -eq 38.5 -and $wsF14.Range('C11').Value2 -eq '사후 관리' -and $wsF14.Range('F11').Value2 -eq 60 -and $wsF14.Range('F20').Value2 -eq 98.5) 'F-014가 R-07 복수 평가행과 K-03 총점을 자동 반영함'
    Assert-Check ([bool]$excel.Run('검증_F014출력가능')) 'F-014 선택 출력이 필수값과 완전한 복수 평가행을 허용함'

    $wsF15 = $wb.Worksheets.Item('F-015_정량적평가')
    Assert-Check ($wsF15.Range('B3').Value2 -eq '[9-4] [붙임 3_1] [1단계] 정량적 평가' -and $wsF15.Range('B8').Value2 -eq '구분' -and $wsF15.Range('C8').Value2 -eq '평가항목' -and $wsF15.Range('D8').Value2 -eq '배점기준' -and $wsF15.Range('E8').Value2 -eq '배점' -and $wsF15.Range('F8').Value2 -eq '평가점수') 'F-015 원문 대조 제목과 고정 배점표 머리글이 존재함'
    $wsF15.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF15.Range('G5').Value2 -eq '검증업체가' -and $wsF15.Range('F9').Value2 -eq 10 -and $wsF15.Range('F13').Value2 -eq 10 -and $wsF15.Range('F16').Value2 -eq 15 -and $wsF15.Range('F20').Value2 -eq 15 -and $wsF15.Range('F22').Value2 -eq 50) 'F-015가 선택 업체 순번(I3)의 4개 항목 점수와 합계 50점을 정확히 반영함'
    $wsF15.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF15.Range('G5').Value2 -eq '검증업체나' -and $wsF15.Range('F9').Value2 -eq 6 -and $wsF15.Range('F22').Value2 -eq 23) 'F-015가 선택 업체 순번을 바꾸면 다른 업체의 점수를 반영함'
    $wsF15.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F015출력가능')) 'F-015 선택 출력이 완전한 업체별 정량평가를 허용함'

    $wsF18 = $wb.Worksheets.Item('F-018_정량적평가자기평점표')
    Assert-Check ($wsF18.Range('B3').Value2 -eq '[9-7] [서식 1_1] 정량적 평가 자기 평점표' -and $wsF18.Range('F7').Value2 -eq '자기 평점') 'F-018 원문 대조 제목과 자기평점 머리글이 존재함'
    $wsF18.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF18.Range('C5').Value2 -eq '검증업체가' -and $wsF18.Range('F8').Value2 -eq 8 -and $wsF18.Range('F12').Value2 -eq 8 -and $wsF18.Range('F15').Value2 -eq 12 -and $wsF18.Range('F19').Value2 -eq 15 -and $wsF18.Range('F21').Value2 -eq 43 -and $wsF15.Range('F9').Value2 -eq 10) 'F-018이 F-015 학교평가(10점)와 다른 업체 자기평점(8점)만 반영하여 혼용하지 않음'
    $wsF18.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF18.Range('C5').Value2 -eq '검증업체나' -and $wsF18.Range('F21').Value2 -eq 20) 'F-018이 선택 업체별 자기평점 합계를 반영함'
    $wsF18.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F018출력가능')) 'F-018 선택 출력이 완전한 업체별 자기평점을 허용함'

    $wsF19 = $wb.Worksheets.Item('F-019_입찰참가신청서')
    Assert-Check ($wsF19.Range('B3').Value2 -match '입 찰 참 가 신 청 서' -and $wsF19.Range('D6').Value2 -match '상호' -and $wsF19.Range('I6').Value2 -eq '법인등록번호') 'F-019 원문 대조 제목과 신청인 표 머리글이 존재함'
    $wsF19.Range('N3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF19.Range('F6').Value2 -eq '검증업체가' -and $wsF19.Range('F10').Value2 -eq '테스트초등학교 공고 제20○○-00호' -and $wsF19.Range('F11').Value2 -eq '테스트초등학교 2026학년도 교복 학교주관 구매 입찰') 'F-019가 선택 업체명과 학교 공통정보만 반영함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF19.Range('J6').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF19.Range('F7').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF19.Range('J7').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF19.Range('F8').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF19.Range('J8').Value2) -and $wsF19.Range('D15').Value2 -match '성 명' -and $wsF19.Range('I15').Value2 -match '사용인감') 'F-019 법인·연락처·대표자·대리인·인감 개인정보는 빈칸 또는 수기 안내로 보존함'
    Assert-Check ([bool]$excel.Run('검증_F019출력가능')) 'F-019 선택 출력이 필수 학교정보와 업체명을 허용함'

    $wsF20 = $wb.Worksheets.Item('F-020_입찰참가신고서')
    Assert-Check ($wsF20.Range('B3').Value2 -eq '교복 학교주관구매 입찰 참가 신고서' -and $wsF20.Range('B7').Value2 -eq '사업자 상호' -and $wsF20.Range('B8').Value2 -eq '사업자 번호') 'F-020 원문 대조 제목과 사업자 정보 표 머리글이 존재함'
    $wsF20.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF20.Range('C7').Value2 -eq '검증업체가' -and $wsF20.Range('B11').Value2 -match '테스트초등학교' -and $wsF20.Range('B20').Value2 -eq '테스트초등학교장 귀하') 'F-020이 선택 업체명과 학교 공통정보만 반영함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF20.Range('C8').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF20.Range('C9').Value2) -and $wsF20.Range('B19').Value2 -match '○○○') 'F-020 사업자 번호·대표자·직인 개인정보는 빈칸 또는 수기 안내로 보존함'
    Assert-Check ([bool]$excel.Run('검증_F020출력가능')) 'F-020 선택 출력이 필수 학교정보와 업체명을 허용함'

    $wsF21 = $wb.Worksheets.Item('F-021_교복납품제안서')
    $wsF21.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF21.Range('B7').Value2 -eq '1. 일반현황 및 연혁' -and $wsF21.Range('B19').Value2 -eq '2. 제안(제출)서류' -and $wsF21.Range('C8').Value2 -eq '검증업체가' -and $wsF21.Range('B31').Value2 -eq '테스트초등학교장 귀하') 'F-021 원문 대조 표와 업체명·학교 공통정보 반영'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF21.Range('C9').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF21.Range('C11').Value2) -and $wsF21.Range('B30').Value2 -match '서명 또는 날인' -and [bool]$excel.Run('검증_F021출력가능')) 'F-021 개인정보 빈칸 및 선택 출력 경계'

    $wsF22 = $wb.Worksheets.Item('F-022_교복납품실적표')
    $wsF22.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF22.Range('B3').Value2 -eq '교복 납품 실적표' -and $wsF22.Range('C9').Value2 -eq '사업명' -and $wsF22.Range('D9').Value2 -eq '사업기간' -and $wsF22.Range('E9').Value2 -match '계약금액' -and $wsF22.Range('F9').Value2 -eq '발주처' -and $wsF22.Range('B5').Value2 -match '검증업체가') 'F-022 원문 대조 제목·실적표 머리글과 선택 업체명이 존재함'
    Assert-Check ($wsF22.Range('B7').Value2 -match '최근 3년' -and $wsF22.Range('B10').Value2 -eq 1 -and $wsF22.Range('B18').Value2 -eq 9 -and $wsF22.Range('B20').Value2 -match '실적증명서' -and $wsF22.Range('B24').Value2 -match '서명 또는 날인' -and [bool]$excel.Run('검증_F022출력가능')) 'F-022 실적증명서 9행·수기 서명 경계 및 선택 출력 허용'

    $wsF23 = $wb.Worksheets.Item('F-023_교복제조사양서')
    $wsF23.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF23.Range('B3').Value2 -eq '납품 교복 제조 사양서' -and $wsF23.Range('B5').Value2 -match '검증업체가' -and $wsF23.Range('B7').Value2 -eq '생활형 교복' -and $wsF23.Range('B9').Value2 -eq '구분' -and $wsF23.Range('C9').Value2 -eq '견본 품목' -and $wsF23.Range('D9').Value2 -eq '재질' -and $wsF23.Range('E9').Value2 -eq '설명 내용') 'F-023 원문 대조 제목·표 머리글과 선택 업체명이 존재함'
    Assert-Check ($wsF23.Range('B10').Value2 -eq '동복' -and $wsF23.Range('C10').Value2 -eq '후드 점퍼' -and $wsF23.Range('C11').Value2 -eq '집업티' -and $wsF23.Range('C12').Value2 -eq '맨투맨티' -and $wsF23.Range('C13').Value2 -eq '긴바지' -and $wsF23.Range('B14').Value2 -eq '하복' -and $wsF23.Range('C14').Value2 -eq '반팔티' -and $wsF23.Range('C15').Value2 -eq '반바지') 'F-023 동복 4행·하복 2행 견본 품목이 원문과 일치함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF23.Range('D10').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF23.Range('E10').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF23.Range('D15').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF23.Range('E15').Value2) -and $wsF23.Range('B17').Value2 -match '국산섬유원단' -and $wsF23.Range('B20').Value2 -match '서명 또는 날인' -and [bool]$excel.Run('검증_F023출력가능')) 'F-023 재질·설명 내용 수기 경계 및 선택 출력 허용'

    $wsF25 = $wb.Worksheets.Item('F-025_교복AS계획서')
    $wsF25.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF25.Range('B3').Value2 -eq '교복 A/S 계획서' -and $wsF25.Range('B6').Value2 -eq '1. A/S 기간' -and $wsF25.Range('B9').Value2 -eq '2. A/S 편의성' -and $wsF25.Range('B15').Value2 -match 'A/S지정업체' -and $wsF25.Range('B20').Value2 -eq '4. A/S 내용') 'F-025 원문 대조 제목과 4개 항목(A/S 기간·편의성·지정업체·내용) 머리글이 존재함'
    Assert-Check ($wsF25.Range('B23').Value2 -match '바지/치마 기장수선' -and $wsF25.Range('B30').Value2 -match '서명 또는 날인' -and $wsF25.Range('B31').Value2 -eq '테스트초등학교장 귀하' -and [bool]$excel.Run('검증_F025출력가능')) 'F-025 원문에 업체명 표시 자리가 없어 학교 공통정보만 반영하고 선택 출력을 허용함'
    $bodyText025 = ($wsF25.Range('B1:E31').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText025 -notmatch '검증업체가') 'F-025 본문 어디에도 업체명이 출력되지 않음(원문에 업체명 표시 자리 없음)'

    $wsF01 = $wb.Worksheets.Item('F-001_위원회구성기안')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF01.Range('B3').Value2 -eq '교복선정위원회 구성(안)' -and $wsF01.Range('B7').Value2 -eq '수  신' -and $wsF01.Range('C7').Value2 -eq '내부결재' -and $wsF01.Range('B9').Value2 -eq '제  목' -and $wsF01.Range('C9').Value2 -match '교복선정위원회 구성\(안\)') 'F-001 원문 대조 제목·수신·제목란이 존재함'
    Assert-Check ($wsF01.Range('B15').Value2 -match '위원회명：교복선정위원회' -and $wsF01.Range('B16').Value2 -match '구성인원' -and $wsF01.Range('B18').Value2 -eq '다. 역 할' -and $wsF01.Range('B24').Value2 -eq '라. 위원 명단' -and $wsF01.Range('B25').Value2 -eq '순' -and $wsF01.Range('C25').Value2 -eq '구분' -and $wsF01.Range('D25').Value2 -eq '직급(직위)' -and $wsF01.Range('E25').Value2 -eq '성명') 'F-001 위원회 구성 안내와 위원 명단 표 머리글이 원문과 일치함'
    Assert-Check ($wsF01.Range('C26').Value2 -eq '위원장' -and $wsF01.Range('D26').Value2 -eq '교원위원' -and $wsF01.Range('E26').Value2 -eq '위원 ○○' -and $wsF01.Range('C27').Value2 -eq '위원') 'F-001 위원 명단이 기초자료입력 위원 반복행(DB_위원 공용)에서 마스킹 식별표시로만 반영됨'
    $bodyText001 = ($wsF01.Range('B3:H43').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText001 -notmatch '홍길동' -and $bodyText001 -notmatch '위원 ○○○') 'F-001 본문 어디에도 실명 형태 문자열이 출력되지 않음(마스킹 식별표시만 허용)'
    Assert-Check ($wsF01.Range('B37').Value2 -match '위원 수락 및 확인서' -and $wsF01.Range('B38').Value2 -match '청렴 및 보안 서약서' -and [bool]$excel.Run('검증_F001출력가능')) 'F-001 붙임 문구가 원문과 일치하고 선택 출력을 허용함'

    $wsF02 = $wb.Worksheets.Item('F-002_위원수락확인서')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF02.Range('B3').Value2 -eq '교복선정위원회 위원 수락 및 확인서' -and $wsF02.Range('B6').Value2 -match '공정하게 수행할 것이며' -and $wsF02.Range('B7').Value2 -match '전혀 관련이 없음을 확인합니다' -and $wsF02.Range('B8').Value2 -match '즉시 위원직을 사퇴') 'F-002 원문 대조 제목과 3단락 서약 본문이 원문과 일치함'
    Assert-Check ($wsF02.Range('B14').Value2 -match '주\s*소' -and $wsF02.Range('B16').Value2 -match '성\s*명' -and $wsF02.Range('B18').Value2 -match '서\s*명' -and [bool]$excel.Run('검증_F002출력가능')) 'F-002 주소·성명·서명란이 존재하고 선택 출력을 허용함'
    $bodyText002 = ($wsF02.Range('B3:E21').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText002 -notmatch '홍길동') 'F-002 본문 어디에도 실명 형태 문자열이 출력되지 않음(주소·성명·서명 전부 수기 작성 공란)'

    $wsF03 = $wb.Worksheets.Item('F-003_위원청렴보안서약서')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF03.Range('B3').Value2 -eq '청렴 및 보안 서약서') 'F-003 B3 제목이 원문과 일치함'
    Assert-Check ($wsF03.Range('B6').Value2 -match '부패 없는 투명한 사회') 'F-003 B6 도입 문구에 부패 없는 투명한 사회가 포함됨'
    Assert-Check ($wsF03.Range('B6').Value2 -notmatch '○○학교' -and $wsF03.Range('B6').Value2 -match [regex]::Escape($wsIn.Range('C4').Value2)) 'F-003 B6 도입 문구에 학교 공통정보가 반영됨(기본값 미노출)'
    Assert-Check ($wsF03.Range('B8').Value2 -match '객관적이고 공정하게 심사' -and $wsF03.Range('B9').Value2 -match '금품' -and $wsF03.Range('B9').Value2 -match '향응' -and $wsF03.Range('B10').Value2 -match '보안관계 규정' -and $wsF03.Range('B11').Value2 -match '업무상 비밀을 타인에게 누설') 'F-003 4개 준수사항 항목이 원문과 일치함'
    Assert-Check ($wsF03.Range('B17').Value2 -match '서\s*약\s*자\s*성\s*명' -and [bool]$excel.Run('검증_F003출력가능')) 'F-003 서약자 성명·서명란이 존재하고 선택 출력을 허용함'
    $bodyText003 = ($wsF03.Range('B3:E20').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyText003 -notmatch '홍길동') 'F-003 본문 어디에도 실명 형태 문자열이 출력되지 않음(성명·서명·날짜 전부 수기 작성 공란)'

    $wsF04 = $wb.Worksheets.Item('F-004_구매추진계획수립기안문')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF04.Range('B3').Value2 -eq '교복 학교주관구매 추진 계획 수립(안)' -and $wsF04.Range('B7').Value2 -eq '수  신' -and $wsF04.Range('C7').Value2 -eq '내부결재' -and $wsF04.Range('C9').Value2 -match '교복 학교주관구매 추진 계획\(안\)') 'F-004 원문 대조 제목·수신·제목란이 존재함'
    Assert-Check ($wsF04.Range('B13').Value2 -match '공정하고 투명한 교복 선정' -and $wsF04.Range('B13').Value2 -notmatch '○○학교' -and $wsF04.Range('B15').Value2 -match '교복 학교주관구매 추진 계획\(안\) 1부' -and [bool]$excel.Run('검증_F004출력가능')) 'F-004 본문·붙임 문구가 원문과 일치하고 선택 출력을 허용함'

    $wsF05 = $wb.Worksheets.Item('F-005_구매추진계획안')
    $excel.CalculateFullRebuild()
    $bodyTextF05 = ($wsF05.UsedRange.Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF05 -match '교복 학교주관구매 추진 계획\(안\)' -and $bodyTextF05 -match '기본방향' -and $bodyTextF05 -match '추진 절차 개요' -and $bodyTextF05 -match '세부 추진 절차' -and $bodyTextF05 -match '기대효과') 'F-005 원문 대조 제목과 5개 장(개요/기본방향/추진절차개요/세부추진절차/기대효과) 표제가 모두 존재함'
    Assert-Check ($bodyTextF05 -match '단계 및 시기' -and $bodyTextF05 -match '학교운영위원회 심의' -and $bodyTextF05 -match '2단계 가격 경쟁') 'F-005 3장 추진 절차 개요 표가 원문과 일치함'
    Assert-Check ($bodyTextF05 -notmatch '○○학교' -and $bodyTextF05 -match [regex]::Escape($wsIn.Range('C4').Value2)) 'F-005 본문에 학교 공통정보가 반영되고 기본값이 노출되지 않음'
    Assert-Check ($bodyTextF05 -notmatch '홍길동' -and $bodyTextF05 -notmatch '위원 ○○○') 'F-005 위원 명단에 실명 형태 문자열이 출력되지 않음(마스킹 식별표시만 허용)'
    Assert-Check ([bool]$excel.Run('검증_F005출력가능')) 'F-005 선택 출력을 허용함'

    $wsF06 = $wb.Worksheets.Item('F-006_운영위원회심의안')
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF06.Range('B3').Value2 -match '교복 학교주관구매 계획\(안\)' -and $wsF06.Range('B6').Value2 -eq '안건번호' -and $wsF06.Range('B8').Value2 -eq '제 안 자') 'F-006 원문 대조 제목·안건 정보란이 존재함'
    Assert-Check ($wsF06.Range('B12').Value2 -match '초·중등교육법' -and $wsF06.Range('B21').Value2 -eq '다. 교복구매 상한가격' -and $wsF06.Range('B22').Value2 -match '하복' -and $wsF06.Range('B22').Value2 -match '동복') 'F-006 제안 이유·상한가격 문구가 원문과 일치함'
    Assert-Check ($wsF06.Range('B24').Value2 -match '교복선정위원회 구성' -and $wsF06.Range('B25').Value2 -eq '순' -and $wsF06.Range('C25').Value2 -eq '구분' -and $wsF06.Range('D25').Value2 -eq '직위' -and $wsF06.Range('E25').Value2 -eq '성명') 'F-006 협의사항 위원회 구성 표 머리글이 원문과 일치함'
    Assert-Check ($wsF06.Range('C26').Value2 -eq '위원장' -and $wsF06.Range('D26').Value2 -eq '교원위원' -and $wsF06.Range('E26').Value2 -eq '위원 ○○') 'F-006 위원 명단이 기초자료입력 위원 반복행(F-001과 동일 공용)에서 마스킹 식별표시로만 반영됨'
    $bodyTextF06 = ($wsF06.Range('B3:H40').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF06 -notmatch '홍길동' -and $bodyTextF06 -notmatch '위원 ○○○') 'F-006 본문 어디에도 실명 형태 문자열이 출력되지 않음(마스킹 식별표시만 허용)'
    Assert-Check ($wsF06.Range('B40').Value2 -match '참고 서식' -and [bool]$excel.Run('검증_F006출력가능')) 'F-006 참고 서식 안내문이 존재하고 선택 출력을 허용함'

    $wsF41 = $wb.Worksheets.Item('F-041_낙찰자결정')
    $wsF41.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF41.Range('B3').Value2 -eq '교복 학교주관구매 낙찰자 결정(안)' -and $wsF41.Range('B7').Value2 -eq '수  신' -and $wsF41.Range('C7').Value2 -eq '내부결재' -and $wsF41.Range('B19').Value2 -eq '라. 낙찰자') 'F-041 원문 대조 제목·수신·낙찰자 항목 라벨이 존재함'
    Assert-Check ($wsF41.Range('B20').Value2 -eq '순위' -and $wsF41.Range('C20').Value2 -eq '업체명' -and $wsF41.Range('D20').Value2 -eq '대표자' -and $wsF41.Range('E20').Value2 -eq '투찰금액(원)' -and $wsF41.Range('F20').Value2 -eq '낙찰률(%)' -and $wsF41.Range('G20').Value2 -eq '비고(주소)') 'F-041 낙찰자 표 머리글(순위/업체명/대표자/투찰금액/낙찰률/비고)이 원문과 일치함'
    Assert-Check ($wsF41.Range('C21').Value2 -eq '검증업체가' -and $wsF41.Range('B23').Value2 -match '지방자치단체를 당사자로 하는 계약에 관한 법률 시행령') 'F-041이 I3 순번 1의 낙찰 업체명을 반영하고 근거 법령 문구가 원문과 일치함'
    $wsF41.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF41.Range('C21').Value2 -eq '검증업체나') 'F-041이 선택 순번을 바꾸면 다른 낙찰 업체명을 반영함'
    $wsF41.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF41.Range('D21').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF41.Range('G21').Value2)) 'F-041 대표자·비고(주소)는 개인정보·업체 상세정보이므로 빈칸으로 보존됨'
    Assert-Check ([bool]$excel.Run('검증_F041출력가능')) 'F-041 선택 출력이 유효한 낙찰 업체 선택을 허용함'
    $bodyTextF41 = ($wsF41.Range('B3:H29').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF41 -notmatch '홍길동') 'F-041 본문 어디에도 대표자 실명 형태 문자열이 출력되지 않음'

    $wsF42 = $wb.Worksheets.Item('F-042_낙찰자결정통보')
    $wsF42.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF42.Range('B3').Value2 -eq '교복 학교주관구매 낙찰자 결정 통보' -and $wsF42.Range('B7').Value2 -eq '수  신' -and $wsF42.Range('C7').Value2 -eq '검증업체가' -and $wsF42.Range('B15').Value2 -eq '가. 낙찰자') 'F-042 원문 대조 제목이 존재하고 수신이 선택 낙찰 업체명으로 반영됨(통보문 유형)'
    Assert-Check ($wsF42.Range('B16').Value2 -eq '업체명' -and $wsF42.Range('C16').Value2 -eq '대표자' -and $wsF42.Range('D16').Value2 -eq '단가(원)' -and $wsF42.Range('E16').Value2 -eq '예정수량' -and $wsF42.Range('F16').Value2 -eq '계약체결금액(원)' -and $wsF42.Range('G16').Value2 -eq '비고') 'F-042 낙찰자 표 머리글이 원문과 일치함'
    Assert-Check ($wsF42.Range('B17').Value2 -eq '검증업체가' -and $wsF42.Range('E17').Value2 -eq 100) 'F-042가 선택 업체명과 예정수량(B-05)을 반영함'
    Assert-Check ($wsF42.Range('B13').Value2 -match '2026\.02\.20\.') 'F-042 제출기한(B-07) 안내문이 기초자료입력 값을 반영함'
    Assert-Check ($wsF42.Range('B20').Value2 -match '품목별 산출내역서' -and $wsF42.Range('B21').Value2 -match '원단 의류 시험성적서' -and $wsF42.Range('B22').Value2 -match '사용인장계' -and $wsF42.Range('B23').Value2 -match '법인 등기사항증명서') 'F-042 제출서류 4개 항목이 원문과 일치함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF42.Range('C17').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF42.Range('G17').Value2)) 'F-042 대표자·비고는 개인정보·업체 상세정보이므로 빈칸으로 보존됨'
    Assert-Check ([bool]$excel.Run('검증_F042출력가능')) 'F-042 선택 출력이 유효한 낙찰 업체 선택을 허용함'
    $bodyTextF42 = ($wsF42.Range('B3:H29').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF42 -notmatch '홍길동') 'F-042 본문 어디에도 대표자 실명 형태 문자열이 출력되지 않음'

    $wsF43 = $wb.Worksheets.Item('F-043_계약체결')
    $wsF43.Range('I3').Value2 = 1
    $wsF43.Range('C16').Value2 = 50000
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF43.Range('B3').Value2 -eq '교복 학교주관구매 계약 체결(안)' -and $wsF43.Range('B7').Value2 -eq '수  신' -and $wsF43.Range('C7').Value2 -eq '내부결재' -and $wsF43.Range('B20').Value2 -eq '3. 계약상대자') 'F-043 원문 대조 제목·수신·계약상대자 라벨이 존재함'
    Assert-Check ($wsF43.Range('C17').Value2 -eq '금5,000,000원[50,000원(단가)×100벌(예정수량)]') 'F-043이 담당자 입력 단가×예정수량(B-05)으로 계약금액을 자동 계산함(K-01/K-02)'
    Assert-Check ($wsF43.Range('B21').Value2 -eq '업체명' -and $wsF43.Range('C21').Value2 -eq '소재지' -and $wsF43.Range('D21').Value2 -eq '대표자' -and $wsF43.Range('E21').Value2 -eq '전화번호') 'F-043 계약상대자 표 머리글이 원문과 일치함'
    Assert-Check ($wsF43.Range('B22').Value2 -eq '검증업체가') 'F-043이 I3 순번 1의 계약상대자 업체명을 반영함'
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF43.Range('C22').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF43.Range('D22').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF43.Range('E22').Value2)) 'F-043 소재지·대표자·전화번호는 개인정보·업체 상세정보이므로 빈칸으로 보존됨'
    Assert-Check ($wsF43.Range('B25').Value2 -match '2026\.03\.01\.') 'F-043 납품기한(B-06)이 기초자료입력 값을 반영함'
    Assert-Check ($wsF43.Range('B28').Value2 -match '계약서\(초안\)' -and $wsF43.Range('B29').Value2 -match '교복 사양서' -and $wsF43.Range('B30').Value2 -match '교복 구매계약 특수조건' -and $wsF43.Range('B31').Value2 -match '입찰가격산출내역서' -and $wsF43.Range('B32').Value2 -match '기타서류') 'F-043 붙임 5개 항목이 원문과 일치함'
    Assert-Check ([bool]$excel.Run('검증_F043출력가능')) 'F-043 선택 출력이 유효한 계약상대자 선택을 허용함'
    $bodyTextF43 = ($wsF43.Range('B3:H37').Cells | ForEach-Object { [string]$_.Value2 }) -join ' '
    Assert-Check ($bodyTextF43 -notmatch '홍길동') 'F-043 본문 어디에도 대표자 실명 형태 문자열이 출력되지 않음'

    $wsF16 = $wb.Worksheets.Item('F-016_정성적평가')
    Assert-Check ($wsF16.Range('B3').Value2 -eq '[9-5] [붙임 3_2] [2단계] 정성적 평가' -and $wsF16.Range('B8').Value2 -eq '구분' -and $wsF16.Range('C8').Value2 -eq '평가항목' -and $wsF16.Range('F8').Value2 -eq '평가점수') 'F-016 원문 대조 제목과 고정 배점표 머리글이 존재함'
    $wsF16.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF16.Range('B6').Value2 -eq '검증업체가' -and $wsF16.Range('F9').Value2 -eq 15 -and $wsF16.Range('F13').Value2 -eq 5 -and $wsF16.Range('F14').Value2 -eq 55) 'F-016이 선택 업체 순번의 정성평가 점수와 가감점 합계를 정확히 반영함'
    $wsF16.Range('I3').Value2 = 2
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF16.Range('B6').Value2 -eq '검증업체나' -and $wsF16.Range('F9').Value2 -eq 9 -and $wsF16.Range('F14').Value2 -eq 33) 'F-016이 선택 업체 순번을 바꾸면 다른 업체의 점수를 반영함'
    $wsF16.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ([bool]$excel.Run('검증_F016출력가능')) 'F-016 선택 출력이 완전한 업체별 정성평가를 허용함'

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
    $wsQuant = $wb.Worksheets.Item('DB_정량평가')
    Assert-Check ($wsQuant.Cells.Item(2, 1).Value2 -eq 1 -and $wsQuant.Cells.Item(2, 2).Value2 -eq 1 -and $wsQuant.Cells.Item(2, 3).Value2 -eq 10 -and $wsQuant.Cells.Item(2, 4).Value2 -eq 10 -and $wsQuant.Cells.Item(2, 5).Value2 -eq 15 -and $wsQuant.Cells.Item(2, 6).Value2 -eq 15 -and $wsQuant.Cells.Item(3, 2).Value2 -eq 2 -and $wsQuant.Cells.Item(3, 3).Value2 -eq 6) '저장하기()가 업체별 F-015 정량평가 점수를 DB_정량평가에 정규화 저장함(Workbook_Open 재보호 목록에 DB_정량평가 포함되어 무한 대기 재발 없음)'
    $wsQual = $wb.Worksheets.Item('DB_정성평가')
    Assert-Check ($wsQual.Cells.Item(2, 1).Value2 -eq 1 -and $wsQual.Cells.Item(2, 2).Value2 -eq 1 -and $wsQual.Cells.Item(2, 3).Value2 -eq 15 -and $wsQual.Cells.Item(2, 4).Value2 -eq 10 -and $wsQual.Cells.Item(2, 5).Value2 -eq 15 -and $wsQual.Cells.Item(2, 6).Value2 -eq 10 -and $wsQual.Cells.Item(2, 7).Value2 -eq 5 -and $wsQual.Cells.Item(3, 2).Value2 -eq 2 -and $wsQual.Cells.Item(3, 7).Value2 -eq -2) '저장하기()가 업체별 F-016 정성평가 점수와 가감점을 DB_정성평가에 정규화 저장함'
    $wsSelf = $wb.Worksheets.Item('DB_자기평점')
    Assert-Check ($wsSelf.Cells.Item(2, 1).Value2 -eq 1 -and $wsSelf.Cells.Item(2, 2).Value2 -eq 1 -and $wsSelf.Cells.Item(2, 3).Value2 -eq 8 -and $wsSelf.Cells.Item(2, 4).Value2 -eq 8 -and $wsSelf.Cells.Item(2, 5).Value2 -eq 12 -and $wsSelf.Cells.Item(2, 6).Value2 -eq 15 -and $wsSelf.Cells.Item(3, 2).Value2 -eq 2 -and $wsSelf.Cells.Item(3, 3).Value2 -eq 10) '저장하기()가 F-018 업체 자기평점을 DB_자기평점에 학교평가와 분리하여 정규화 저장함'
    $wsCommittee = $wb.Worksheets.Item('DB_위원')
    Assert-Check ($wsCommittee.Cells.Item(2, 1).Value2 -eq 1 -and $wsCommittee.Cells.Item(2, 2).Value2 -eq 1 -and $wsCommittee.Cells.Item(2, 3).Value2 -eq '교원위원' -and $wsCommittee.Cells.Item(2, 4).Value2 -eq '위원 ○○' -and $wsCommittee.Cells.Item(3, 3).Value2 -eq '학부모위원' -and $wsCommittee.Cells.Item(3, 4).Value2 -eq '위원 **') '저장하기()가 복수 위원 역할·마스킹 식별표시를 DB_위원에 정규화 저장함'
    $wsScore = $wb.Worksheets.Item('DB_평가')
    Assert-Check ($wsScore.Cells.Item(2, 1).Value2 -eq 1 -and $wsScore.Cells.Item(2, 2).Value2 -eq 1 -and $wsScore.Cells.Item(2, 3).Value2 -eq '제품 품질' -and $wsScore.Cells.Item(2, 4).Value2 -eq 40 -and $wsScore.Cells.Item(2, 5).Value2 -eq 38.5 -and $wsScore.Cells.Item(3, 5).Value2 -eq 60 -and $wsIn.Range('D82').Value2 -eq 98.5) '저장하기()가 복수 평가점수와 K-03 총점을 DB_평가에 정규화 저장함'
    Assert-Check ([bool]$excel.Run('검증_위원행검증') -and [bool]$excel.Run('검증_평가행검증')) '위원 마스킹·평가점수 정상값과 경계값(점수=배점)을 허용함'

    # 6.0 저장 경계: 무암호 보호를 해제한 직접 DB 편집도 Workbook_BeforeSave에서 재검증해 영구 저장을 취소한다.
    foreach ($internalSheet in @($wsDB, $wsItems, $wsVendors, $wsCommittee, $wsScore, $wsQuant, $wsQual, $wsSelf)) { $internalSheet.Unprotect("") }
    Assert-Check ([bool]$excel.Run('검증_저장경계')) '정상 DB·반복행은 저장 경계 재검증을 통과함'
    $wb.Save()
    Assert-Check ($wb.Saved) '정상 DB·반복행은 Workbook_BeforeSave 경계를 거쳐 저장됨'

    $wsCommittee.Cells.Item(2, 4).Value2 = '홍길동'
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_위원의 실명 R-02 직접 입력을 저장 경계가 거부함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check (-not $wb.Saved) 'DB_위원 실명 R-02 직접 입력은 Workbook_BeforeSave에서 저장이 취소됨'
    $wsCommittee.Cells.Item(2, 4).Value2 = '위원 ○○'

    $wsCommittee.Cells.Item(2, 3).ClearContents()
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_위원 역할·식별표시 불완전 직접 입력을 저장 경계가 거부함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check (-not $wb.Saved) 'DB_위원 역할 누락 직접 입력은 Workbook_BeforeSave에서 저장이 취소됨'
    $wsCommittee.Cells.Item(2, 3).Value2 = '교원위원'

    $wsScore.Cells.Item(2, 5).Value2 = 40.01
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_평가 배점 초과 점수 직접 입력을 저장 경계가 거부함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check (-not $wb.Saved) 'DB_평가 배점 초과 직접 입력은 Workbook_BeforeSave에서 저장이 취소됨'
    $wsScore.Cells.Item(2, 5).Value2 = 38.5

    $wsQual.Cells.Item(2, 7).Value2 = 6
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_정성평가 가감점 범위 초과 직접 입력을 저장 경계가 거부함'
    $wsQual.Cells.Item(2, 7).Value2 = 5

    $wsVendors.Cells.Item(2, 3).Value2 = '010-1234-5678'
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB_업체 연락처성 직접 입력을 저장 경계가 거부함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check (-not $wb.Saved) 'DB_업체 연락처성 직접 입력은 Workbook_BeforeSave에서 저장이 취소됨'
    $wsVendors.Cells.Item(2, 3).Value2 = '검증업체가'

    $wsDB.Cells.Item(3, 1).Value2 = 2
    Assert-Check (-not [bool]$excel.Run('검증_저장경계')) 'DB 불완전 주 레코드 직접 입력을 저장 경계가 거부함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check (-not $wb.Saved) 'DB 불완전 주 레코드 직접 입력은 Workbook_BeforeSave에서 저장이 취소됨'
    $wsDB.Range('A3:T3').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_저장경계')) '직접 입력 오류를 원복하면 저장 경계가 정상 데이터만 허용함'
    $wb.Saved = $false; $wb.Save()
    Assert-Check ($wb.Saved) '직접 입력 오류를 원복한 정상 데이터는 Workbook_BeforeSave를 거쳐 저장됨'

    $wsIn.Range('B60').Value2 = '외부위원'
    $wsIn.Range('C60').Value2 = '마스킹 없음'
    Assert-Check (-not [bool]$excel.Run('검증_위원행검증')) '위원 식별표시에 마스킹 기호가 없으면 저장을 거부함'
    $wsIn.Range('C60').ClearContents()
    $wsIn.Range('B60').Value2 = '외부위원'
    Assert-Check (-not [bool]$excel.Run('검증_위원행검증')) '위원 역할만 있는 불완전 반복행을 거부함'
    $wsIn.Range('B60:C60').ClearContents()
    $wsIn.Range('C58').Value2 = '010-1234-5678 ○'
    Assert-Check (-not [bool]$excel.Run('검증_위원행검증')) '위원 식별표시에 연락처성 숫자열이 있으면 저장을 거부함'
    $wsIn.Range('C58').Value2 = '홍길동 ○'
    Assert-Check (-not [bool]$excel.Run('검증_위원행검증')) '실명 의심값과 마스킹 기호를 결합한 R-02 우회 입력을 거부함'
    Invoke-ValidationMacro -macroName '검증_수정하기'
    Assert-Check ($wsCommittee.Cells.Item(2, 4).Value2 -eq '위원 ○○') '실명 의심 R-02 우회 입력은 수정 저장을 차단함'
    $wsIn.Range('C58').Value2 = '위원 ○○'
    $wsIn.Range('D74').Value2 = 1
    Assert-Check (-not [bool]$excel.Run('검증_평가행검증')) '평가항목·배점 없는 점수 고아 입력행을 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F014출력가능')) 'F-014 선택 출력이 불완전 평가행을 차단함'
    $wsIn.Range('B74:D74').ClearContents()
    $wsIn.Range('B74').Value2 = '가격 경쟁력'
    $wsIn.Range('C74').Value2 = 20
    $wsIn.Range('D74').Value2 = 20.01
    Assert-Check (-not [bool]$excel.Run('검증_평가행검증')) '배점을 초과한 평가점수를 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F014출력가능')) 'F-014 선택 출력이 배점 초과 평가행을 차단함'
    $wsIn.Range('B74:D74').ClearContents()

    $wsIn.Range('C44').Value2 = 11
    Assert-Check (-not [bool]$excel.Run('검증_정량평가행검증')) 'F-015 배점 상한(수행경험 10점)을 초과한 점수를 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F015출력가능')) 'F-015 선택 출력이 배점 초과 정량평가를 차단함'
    $wsIn.Range('C44').Value2 = 10
    $wsIn.Range('D46').Value2 = 5
    Assert-Check (-not [bool]$excel.Run('검증_정량평가행검증')) '업체명 없이 정량평가 점수만 있는 F-015 고아 입력을 거부함'
    $wsIn.Range('D46').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_정량평가행검증')) 'F-015 정량평가 정상 복귀 후 저장 검증을 통과함'
    $wsIn.Range('N44').Value2 = 11
    Assert-Check (-not [bool]$excel.Run('검증_자기평점행검증')) 'F-018 자기 수행경험 배점 상한(10점)을 초과한 점수를 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F018출력가능')) 'F-018 선택 출력이 배점 초과 자기평점을 차단함'
    $wsIn.Range('N44').Value2 = 8
    $wsIn.Range('O46').Value2 = 5
    Assert-Check (-not [bool]$excel.Run('검증_자기평점행검증')) '업체명 없이 자기평점만 있는 F-018 고아 입력을 거부함'
    $wsIn.Range('O46').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_자기평점행검증')) 'F-018 자기평점 정상 복귀 후 저장 검증을 통과함'
    $wsIn.Range('H44').Value2 = 16
    Assert-Check (-not [bool]$excel.Run('검증_정성평가행검증')) 'F-016 재질 배점 상한을 초과한 점수를 거부함'
    Assert-Check (-not [bool]$excel.Run('검증_F016출력가능')) 'F-016 선택 출력이 배점 초과 정성평가를 차단함'
    $wsIn.Range('H44').Value2 = 15
    $wsIn.Range('L46').Value2 = 0
    Assert-Check (-not [bool]$excel.Run('검증_정성평가행검증')) '업체명 없이 정성평가 점수만 있는 F-016 고아 입력을 거부함'
    $wsIn.Range('L46').ClearContents()
    Assert-Check ([bool]$excel.Run('검증_정성평가행검증')) 'F-016 정성평가 정상 복귀 후 저장 검증을 통과함'
    $wsIn.Range('B58:C67').ClearContents()
    $wsIn.Range('B72:D81').ClearContents()
    $wsIn.Range('B44:B53').ClearContents()
    $wsIn.Range('H44:L53').ClearContents()
    $wsIn.Range('N44:Q53').ClearContents()
    Invoke-ValidationMacro -macroName '검증_첫레코드불러오기'
    Assert-Check ($wsIn.Range('B44').Value2 -eq '검증업체가' -and $wsIn.Range('B45').Value2 -eq '검증업체나') '불러오기()가 업체 반복행을 원래 행에 복원함'
    Assert-Check ($wsIn.Range('C44').Value2 -eq 10 -and $wsIn.Range('D44').Value2 -eq 10 -and $wsIn.Range('E44').Value2 -eq 15 -and $wsIn.Range('F44').Value2 -eq 15 -and $wsIn.Range('C45').Value2 -eq 6) '불러오기()가 F-015 업체별 정량평가 점수를 원래 행에 복원함'
    Assert-Check ($wsIn.Range('H44').Value2 -eq 15 -and $wsIn.Range('I44').Value2 -eq 10 -and $wsIn.Range('J44').Value2 -eq 15 -and $wsIn.Range('K44').Value2 -eq 10 -and $wsIn.Range('L44').Value2 -eq 5 -and $wsIn.Range('L45').Value2 -eq -2) '불러오기()가 F-016 업체별 정성평가 점수와 가감점을 원래 행에 복원함'
    Assert-Check ($wsIn.Range('N44').Value2 -eq 8 -and $wsIn.Range('O44').Value2 -eq 8 -and $wsIn.Range('P44').Value2 -eq 12 -and $wsIn.Range('Q44').Value2 -eq 15 -and $wsIn.Range('N45').Value2 -eq 10 -and $wsIn.Range('Q45').Value2 -eq 0) '불러오기()가 F-018 업체별 자기평점을 학교평가와 분리하여 원래 행에 복원함'
    Assert-Check ($wsIn.Range('B58').Value2 -eq '교원위원' -and $wsIn.Range('C58').Value2 -eq '위원 ○○' -and $wsIn.Range('B59').Value2 -eq '학부모위원' -and $wsIn.Range('D82').Value2 -eq 98.5) '불러오기()가 위원·평가 반복행과 K-03 총점을 복원함'
    $wsIn.Range('C59').Value2 = '위원 ○*'
    $wsIn.Range('D73').Value2 = 59.5
    $wsIn.Range('B60').Value2 = '교원위원'
    $wsIn.Range('C60').Value2 = '위원 ○○'
    Assert-Check ([bool]$excel.Run('검증_위원중복경고')) '수정 전 위원 역할 중복 경고를 감지함'
    $wsIn.Range('B60:C60').ClearContents()
    Invoke-ValidationMacro -macroName '검증_수정하기'
    $wsIn.Range('B58:C67').ClearContents()
    $wsIn.Range('B72:D81').ClearContents()
    Invoke-ValidationMacro -macroName '검증_첫레코드불러오기'
    Assert-Check ($wsIn.Range('C59').Value2 -eq '위원 ○*' -and $wsIn.Range('D73').Value2 -eq 59.5 -and $wsScore.Cells.Item(3, 5).Value2 -eq 59.5 -and $wsIn.Range('D82').Value2 -eq 98) '수정하기()가 위원·평가 반복행을 교체하고 총점을 재계산함'
    # 업체 2의 학교 정량·정성 평가와 F-018 자기평점을 함께 비워야 업체명 없는 고아행이 남지 않는다.
    $wsIn.Range('B45:F45').ClearContents()
    $wsIn.Range('H45:L45').ClearContents()
    $wsIn.Range('N45:Q45').ClearContents()
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

    # F-007/F-014/F-024 수식이 HWPX 원문 대조 후 확정한 배치에서 기초자료입력을 정상 참조하는지 확인
    $wsF7 = $wb.Worksheets.Item("F-007_구매요청기안문")
    L "F-007 제목 셀(C9) 계산값: $($wsF7.Range('C9').Value2)"
    Assert-Check ($wsF7.Range('C9').Value2 -eq '2026학년도 동복 학교주관구매 요청') 'F-007 공통 제목이 입력값을 참조함'
    $wsF24 = $wb.Worksheets.Item("F-024_단가비율표")
    L "F-024 수량 합계(D15) 계산값: $($wsF24.Range('D15').Value2), 비율(E9/E10): $($wsF24.Range('E9').Value2)/$($wsF24.Range('E10').Value2)"
    Assert-Check ($wsF24.Range('D15').Value2 -eq 204 -and $wsF24.Range('E9').Text -eq '55.6%' -and $wsF24.Range('E10').Text -eq '44.4%') 'F-024 수량 합계와 단가비율이 입력값을 참조함'
    Assert-Check ($wsF24.Range('C9').Value2 -eq '동복 상의' -and $wsF24.Range('C10').Value2 -eq '동복 하의') 'F-024 품목명이 입력 반복행을 참조함'
    $wsF17 = $wb.Worksheets.Item('F-017_제출서류자기확인서')
    $wsF17.Range('G3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF17.Range('B8').Value2 -eq '검증업체가' -and [string]::IsNullOrWhiteSpace([string]$wsF17.Range('C8').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF17.Range('D8').Value2) -and [string]::IsNullOrWhiteSpace([string]$wsF17.Range('E8').Value2) -and $wsF17.Range('B10').Value2 -eq '제출서류') 'F-017은 업체명만 반영하고 대표자·연락처·대리인 정보는 빈 양식으로 유지함'
    Assert-Check ([bool]$excel.Run('검증_F017출력가능')) 'F-017 선택 출력이 공통 필수값과 업체 반복행을 확인함'
    $longTitle = ('평가 기준 & 특수문자 <>()[]{} / ' * 4)
    $wsIn.Range('C4').Value2 = ('검증학교 & 특수문자 <>()[]{} ' * 3)
    $wsIn.Range('C22').Value2 = $longTitle
    # DB_정량평가 연동 제거로 F-015 점수는 불러오기() 이후 유지되지 않으므로(위 DB_정량평가 주석 참고),
    # 수식 검증을 위해 여기서 다시 직접 입력함.
    $wsIn.Range('B44').Value2 = '검증업체가'
    $wsIn.Range('C44').Value2 = 10
    $wsIn.Range('D44').Value2 = 10
    $wsIn.Range('E44').Value2 = 15
    $wsIn.Range('F44').Value2 = 15
    $wsF15.Range('I3').Value2 = 1
    $excel.CalculateFullRebuild()
    Assert-Check ($wsF14.Range('C5').Value2 -eq $wsIn.Range('C4').Value2 -and $wsF14.Range('B6').Value2 -eq $longTitle -and $wsF14.Range('F20').Value2 -eq 98) 'F-014 최대 길이·특수문자 공통값과 점수 합계가 수식 오류 없이 반영됨'
    Assert-Check ($wsF15.Range('C5').Value2 -eq $wsIn.Range('C4').Value2 -and $wsF15.Range('F22').Value2 -eq 50) 'F-015 최대 길이·특수문자 학교명과 정량평가 합계가 수식 오류 없이 반영됨(같은 세션 직접 입력값 기준)'
    $wsIn.Range('C4').Value2 = '테스트초등학교'
    $wsIn.Range('C22').Value2 = '2026학년도 동복 학교주관구매 요청'
    $excel.CalculateFullRebuild()

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
    # F-007, F-014, F-024 행 찾기
    $lastRow = $wsSel.Cells.Item($wsSel.Rows.Count, 2).End(-4162).Row  # xlUp
    for ($r = 5; $r -le $lastRow; $r++) {
        $fid = $wsSel.Cells.Item($r, 2).Value2
        if ($fid -eq "F-001" -or $fid -eq "F-002" -or $fid -eq "F-003" -or $fid -eq "F-004" -or $fid -eq "F-005" -or $fid -eq "F-006" -or $fid -eq "F-007" -or $fid -eq "F-014" -or $fid -eq "F-015" -or $fid -eq "F-016" -or $fid -eq "F-017" -or $fid -eq "F-018" -or $fid -eq "F-019" -or $fid -eq "F-020" -or $fid -eq "F-021" -or $fid -eq "F-022" -or $fid -eq "F-023" -or $fid -eq "F-024" -or $fid -eq "F-025") {
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
    $leftoverF015TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F015_임시_*' })
    Assert-Check ($leftoverF015TempSheets.Count -eq 0) 'F-015 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF016TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F016_임시_*' })
    Assert-Check ($leftoverF016TempSheets.Count -eq 0) 'F-016 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF017TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F017_임시_*' })
    Assert-Check ($leftoverF017TempSheets.Count -eq 0) 'F-017 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF018TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F018_임시_*' })
    Assert-Check ($leftoverF018TempSheets.Count -eq 0) 'F-018 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF019TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F019_임시_*' })
    Assert-Check ($leftoverF019TempSheets.Count -eq 0) 'F-019 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF020TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F020_임시_*' })
    Assert-Check ($leftoverF020TempSheets.Count -eq 0) 'F-020 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF021TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F021_임시_*' })
    Assert-Check ($leftoverF021TempSheets.Count -eq 0) 'F-021 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF022TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F022_임시_*' })
    Assert-Check ($leftoverF022TempSheets.Count -eq 0) 'F-022 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF023TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F023_임시_*' })
    Assert-Check ($leftoverF023TempSheets.Count -eq 0) 'F-023 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'
    $leftoverF025TempSheets = @($wb.Worksheets | Where-Object { $_.Name -like 'F025_임시_*' })
    Assert-Check ($leftoverF025TempSheets.Count -eq 0) 'F-025 업체별 임시 인쇄 시트가 PDF 저장 후 정리됨'

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
    $excel.CalculateFullRebuild()
    Assert-Check ([string]::IsNullOrWhiteSpace([string]$wsF17.Range('B8').Value2)) 'F-017 빈 배포본은 업체명 미선택 시 0 대신 빈 칸을 표시함'
    $wsDB2 = $wb.Worksheets.Item("DB")
    foreach ($internalSheet in @($wsDB2, $wb.Worksheets.Item('DB_품목'), $wb.Worksheets.Item('DB_업체'), $wb.Worksheets.Item('DB_위원'), $wb.Worksheets.Item('DB_평가'), $wb.Worksheets.Item('DB_정량평가'), $wb.Worksheets.Item('DB_정성평가'), $wb.Worksheets.Item('DB_자기평점'))) {
        $internalSheet.Unprotect("")
    }
    $wsDB2.Range("A2:T2").ClearContents()
    $wb.Worksheets.Item('DB_품목').Range('A2:F100').ClearContents()
    $wb.Worksheets.Item('DB_업체').Range('A2:C100').ClearContents()
    $wb.Worksheets.Item('DB_위원').Range('A2:D100').ClearContents()
    $wb.Worksheets.Item('DB_평가').Range('A2:E100').ClearContents()
    $wb.Worksheets.Item('DB_정량평가').Range('A2:F100').ClearContents()
    $wb.Worksheets.Item('DB_정성평가').Range('A2:G100').ClearContents()
    $wb.Worksheets.Item('DB_자기평점').Range('A2:F100').ClearContents()
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
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($excelProcessId) {
        Start-Sleep -Milliseconds 500
        if (Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue) {
            Stop-Process -Id $excelProcessId -Force -ErrorAction SilentlyContinue
            L "PID $excelProcessId Excel 프로세스가 Quit() 이후에도 남아 있어 강제 종료함"
        }
    }
    if (Test-Path -LiteralPath $validationDirectory) { Remove-Item -LiteralPath $validationDirectory -Recurse -Force }
}

if ($failures.Count -gt 0) {
    throw ('Excel v1 검증 실패: ' + ($failures -join '; '))
}

L 'PASS: Excel v1 검증 전체 통과'
