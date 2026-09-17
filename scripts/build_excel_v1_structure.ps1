# P3-01 클린룸 재구현 — 구조 빌더(1/2): 시트·필드·서식선택 UI·A4 페이지 설정
# 원본 파일은 읽기 전용으로만 열며(학교정보 표본 복사), 수정하지 않음.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $root 'artifacts\excel'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$defaultOutPath = Join-Path $outDir 'build.xlsm'
$outPath = if ([string]::IsNullOrWhiteSpace($env:UNIFORM_EXCEL_BUILD_PATH)) { $defaultOutPath } else { $env:UNIFORM_EXCEL_BUILD_PATH }
$outParent = Split-Path -Parent $outPath
if (-not (Test-Path $outParent)) { New-Item -ItemType Directory -Path $outParent -Force | Out-Null }
$sourcePath = Join-Path $root '20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm'
$logPath = Join-Path $root '_workspace\03_excel\build_structure_log.txt'

function CmToPt($cm) { return [double]$cm * 28.3465 }

# Quit()·ReleaseComObject·GC만으로는 남은 스크립트 지역변수(워크시트 등)가
# COM 참조를 계속 살려 두어 EXCEL.EXE가 좀비로 남을 수 있음. 생성 전후 프로세스
# 목록을 비교해 새로 뜬 PID를 기록해 두고, finally에서 종료가 확인되지 않으면
# 이 PID만 강제 종료해 고아 프로세스가 다음 실행을 막지 않게 함.
$excelPidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel = New-Object -ComObject Excel.Application
Start-Sleep -Milliseconds 300
$excelProcessId = Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $excelPidsBefore -notcontains $_.Id } |
    Select-Object -First 1 -ExpandProperty Id
$excel.Visible = $false
$excel.DisplayAlerts = $false
$log = New-Object System.Text.StringBuilder
function L($s) { [void]$log.AppendLine($s) }

$wbNew = $null
$wbSrc = $null
try {
    # 학교정보는 원본에서 읽기 전용으로 매번 추출함. 기존 생성본을 시드로 재사용하지 않아
    # 이전 생성 실패·누락이 다음 빌드에 전파되지 않게 함.
    $schoolSeedPath = $sourcePath
    $wbSrc = $excel.Workbooks.Open($schoolSeedPath, [Type]::Missing, $true)
    $wsSrcSchool = $wbSrc.Worksheets.Item("학교정보")
    L "학교정보 시드 통합문서를 읽기 전용으로 열기 완료: $schoolSeedPath"

    # ---- 1. 새 워크북 생성 ----
    $wbNew = $excel.Workbooks.Add()
    while ($wbNew.Worksheets.Count -gt 1) { $wbNew.Worksheets.Item($wbNew.Worksheets.Count).Delete() }
    $wsFirst = $wbNew.Worksheets.Item(1)
    $wsFirst.Name = "사용설명서"

    # ---- 2. 사용설명서 ----
    $ws = $wsFirst
    $ws.Range("A1").Value2 = "교복구매 학교주관구매 길라잡이 — Excel 기초자료·서식 선택 도구 (클린룸 재구현본, v1)"
    $ws.Range("A1").Font.Size = 14
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A3").Value2 = "1. 이 파일을 열면 Excel 상단에 '보안 경고 - 매크로 사용 안 함' 알림이 뜹니다. [콘텐츠 사용]을 눌러 매크로를 허용해야 입력·출력 기능이 동작합니다."
    $ws.Range("A4").Value2 = "2. '기초자료입력' 시트에 공통·사업·문서 정보와 품목·업체·위원·평가점수 반복행을 입력한 뒤 [저장하기] 버튼을 누르면 정규화 DB 시트에 기록됩니다."
    $ws.Range("A5").Value2 = "3. '서식선택_출력' 시트에서 원하는 서식에 체크(TRUE)한 뒤 [선택 서식 인쇄 미리보기]/[선택 서식 PDF 저장] 버튼을 누릅니다."
    $ws.Range("A6").Value2 = "4. '학교검색' 시트에서 학교명·지역·급별 조건을 입력하고 [검색] 후 결과 행을 선택해 [선택 학교를 기초자료에 반영]을 누르면 학교명이 기초자료입력!C4에 반영됩니다."
    $ws.Range("A7").Value2 = "5. '절차안내' 시트는 교복구매 9단계와 단계별 관련 Form ID를 제공합니다. 단계 행을 선택하고 [관련 Form ID로 이동]을 누르면 서식선택_출력의 해당 Form ID로 이동합니다."
    $ws.Range("A8").Value2 = "6. 이 v1 버전은 F-007, F-014, F-024 세 서식만 완전히 구현되어 있습니다. 나머지 서식은 '구현상태' 열에 표시된 대로 순차 추가 예정입니다."
    $ws.Range("A9").Value2 = "7. F-013(교복 디자인 및 규격서)은 표·이미지가 많은 다쪽(11쪽) 문서로, Excel보다 HWPX 경로가 적합하여 이번 버전에서는 보류하고 document-automation-engineer 협업 대상으로 남겼습니다."
    $ws.Range("A10").Value2 = "8. 업체 후보는 '기초자료입력' 시트의 업체 반복행에 상호명만 입력합니다. 대표자·연락처·전화번호·사업자번호는 입력·저장하지 않으며, 중복 상호는 실제 동일 업체인지 확인합니다."
    $ws.Range("A11").Value2 = "9. 위원 반복행에는 역할·직위와 ○ 또는 *를 포함한 마스킹 식별표시만 입력합니다. 실제 성명·연락처·서명은 입력·저장하지 않습니다. 평가점수는 항목별 배점을 초과할 수 없고 총점은 자동 계산됩니다."
    $ws.Range("A12").Value2 = "10. 내부 DB 시트의 숨김·무암호 보호는 우발적 편집 방지용이며 보안 경계가 아닙니다. 매크로 허용 상태에서는 저장 직전에 DB·반복행 전체를 재검증하고, 개인정보성 표기·불완전 위원·평가행 등 허용되지 않은 값이 있으면 저장을 취소합니다."
    $ws.Range("A13").Value2 = "원본 보호: 이 파일은 20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm 을 참고해 클린룸 방식으로 새로 작성한 사본이며, 원본 파일을 직접 열거나 수정하지 않습니다."
    $ws.Range("A14").Value2 = "개인정보 경계: 이 파일은 학교·계약 단위 업무 정보만 다루며, 학생·학부모 개인정보 및 서명·직인 자동처리는 포함하지 않습니다."
    $ws.Columns.Item("A").ColumnWidth = 110
    $ws.Range("A3:A14").WrapText = $false

    L "사용설명서 시트 작성 완료"

    # ---- 3. 기초자료입력 ----
    $wsBase = $wbNew.Worksheets.Add()
    $wsBase.Name = "기초자료입력"
    $ws = $wsBase
    $ws.Range("A1").Value2 = "기초자료 입력 (공통 C-01~C-06 / 사업 B-01~B-07 / 문서별 D-01~D-05 / 반복 R-01~R-07)"
    $ws.Range("A1:F1").Merge() | Out-Null
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true

    $ws.Range("B3").Value2 = "1. 공통 정보 (C)"
    $ws.Range("B3").Font.Bold = $true
    $fields1 = @(
        @{r=4; id="C-01"; label="학교명 *"; cell="C4"},
        @{r=5; id="C-02"; label="학년도 *"; cell="C5"},
        @{r=6; id="C-03"; label="담당부서"; cell="C6"},
        @{r=7; id="C-04"; label="담당자 직위"; cell="C7"},
        @{r=8; id="C-05"; label="문서 발행일"; cell="C8"},
        @{r=9; id="C-06"; label="문서번호"; cell="C9"}
    )
    foreach ($f in $fields1) {
        $ws.Range("B$($f.r)").Value2 = $f.label
        $ws.Range("A$($f.r)").Value2 = $f.id
        $ws.Range("A$($f.r)").Font.Size = 8
        $ws.Range("A$($f.r)").Font.Color = 10526880
    }
    $ws.Range("C5").NumberFormat = "0"
    $ws.Range("C8").NumberFormat = "yyyy.mm.dd."

    $ws.Range("B11").Value2 = "2. 사업 정보 (B)"
    $ws.Range("B11").Font.Bold = $true
    $fields2 = @(
        @{r=12; id="B-01"; label="구매명 *"},
        @{r=13; id="B-02"; label="구매 학년"},
        @{r=14; id="B-03"; label="계약방식 (계약방법안내에서 검토 후 반영)"},
        @{r=15; id="B-04"; label="기초금액(원)"},
        @{r=16; id="B-05"; label="예정수량"},
        @{r=17; id="B-06"; label="납품기한"},
        @{r=18; id="B-07"; label="제출기한"}
    )
    foreach ($f in $fields2) {
        $ws.Range("B$($f.r)").Value2 = $f.label
        $ws.Range("A$($f.r)").Value2 = $f.id
        $ws.Range("A$($f.r)").Font.Size = 8
        $ws.Range("A$($f.r)").Font.Color = 10526880
    }
    $ws.Range("C15").NumberFormat = "#,##0"
    $ws.Range("C16").NumberFormat = "0"
    $ws.Range("C17").NumberFormat = "yyyy.mm.dd."
    $ws.Range("C18").NumberFormat = "yyyy.mm.dd. hh:mm"

    $ws.Range("B20").Value2 = "3. 문서별 정보 (D) — 서식 공통 기본값"
    $ws.Range("B20").Font.Bold = $true
    $fields3 = @(
        @{r=21; id="D-01"; label="수신"},
        @{r=22; id="D-02"; label="제목 *"},
        @{r=23; id="D-03"; label="관련문서"},
        @{r=24; id="D-04"; label="붙임 목록"},
        @{r=25; id="D-05"; label="안내문 본문"}
    )
    foreach ($f in $fields3) {
        $ws.Range("B$($f.r)").Value2 = $f.label
        $ws.Range("A$($f.r)").Value2 = $f.id
        $ws.Range("A$($f.r)").Font.Size = 8
        $ws.Range("A$($f.r)").Font.Color = 10526880
    }
    $ws.Range("C25:F25").Merge() | Out-Null
    $ws.Rows.Item(25).RowHeight = 60
    $ws.Range("C25").VerticalAlignment = -4160  # xlTop
    $ws.Range("C25").WrapText = $true

    $ws.Range("B27").Value2 = "4. 품목 반복행 (R-04 품목명 / R-05 수량 / R-06 단가 / K-01 금액 자동계산) — F-024 등에서 사용"
    $ws.Range("B27").Font.Bold = $true
    $ws.Range("B28").Value2 = "품목명"
    $ws.Range("C28").Value2 = "수량"
    $ws.Range("D28").Value2 = "단가(원)"
    $ws.Range("E28").Value2 = "금액(자동, K-01)"
    $ws.Range("B28:E28").Font.Bold = $true
    for ($i = 0; $i -lt 10; $i++) {
        $r = 29 + $i
        $ws.Range("D$r").NumberFormat = "#,##0"
        $ws.Range("E$r").Formula = "=IF(AND(B$r<>`"`",C$r<>`"`",D$r<>`"`"),C$r*D$r,`"`")"
        $ws.Range("E$r").NumberFormat = "#,##0"
    }
    $ws.Range("B39").Value2 = "합계 (K-02)"
    $ws.Range("B39").Font.Bold = $true
    $ws.Range("E39").Formula = "=SUM(E29:E38)"
    $ws.Range("E39").NumberFormat = "#,##0"
    $ws.Range("E39").Font.Bold = $true

    $ws.Range("B42").Value2 = "5. 업체 반복행 (R-03 업체명) — C~F열은 F-015 학교 정량평가, H~L열은 F-016 학교 정성평가, N~Q열은 F-018 업체 자기평점이며 서로 혼용하지 않음"
    $ws.Range("B42").Font.Bold = $true
    $ws.Range("B43").Value2 = "업체명"
    $ws.Range("C43").Value2 = "수행경험(10)"
    $ws.Range("D43").Value2 = "공인인증(10)"
    $ws.Range("E43").Value2 = "거리적접근성(15)"
    $ws.Range("F43").Value2 = "상한가격(15)"
    $ws.Range("G43").Value2 = "F-015 상태"
    $ws.Range("H43").Value2 = "재질(15)"
    $ws.Range("I43").Value2 = "완성도(10)"
    $ws.Range("J43").Value2 = "A/S(15)"
    $ws.Range("K43").Value2 = "하자보상(10)"
    $ws.Range("L43").Value2 = "가감점(-15~5)"
    $ws.Range("M43").Value2 = "F-016 상태"
    $ws.Range("N43").Value2 = "자기 수행경험(10)"
    $ws.Range("O43").Value2 = "자기 공인인증(10)"
    $ws.Range("P43").Value2 = "자기 거리(15)"
    $ws.Range("Q43").Value2 = "자기 상한가격(15)"
    $ws.Range("R43").Value2 = "F-018 상태"
    $ws.Range("B43:R43").Font.Bold = $true
    for ($i = 0; $i -lt 10; $i++) {
        $r = 44 + $i
        $ws.Range("B$r:F$r").Interior.Color = 16777164
        $ws.Range("C$r:F$r").NumberFormat = "0"
        $ws.Range("G$r").Formula = "=IF(B$r=`"`",`"`",IF(AND(ISNUMBER(C$r),ISNUMBER(D$r),ISNUMBER(E$r),ISNUMBER(F$r),C$r>=0,C$r<=10,D$r>=0,D$r<=10,E$r>=0,E$r<=15,F$r>=0,F$r<=15),`"정상`",`"점수 확인`"))"
        $ws.Range("H$r:L$r").Interior.Color = 16777164
        $ws.Range("H$r:L$r").NumberFormat = "0"
        $ws.Range("M$r").Formula = "=IF(B$r=`"`",`"`",IF(AND(ISNUMBER(H$r),ISNUMBER(I$r),ISNUMBER(J$r),ISNUMBER(K$r),ISNUMBER(L$r),H$r>=0,H$r<=15,I$r>=0,I$r<=10,J$r>=0,J$r<=15,K$r>=0,K$r<=10,L$r>=-15,L$r<=5),`"정상`",`"점수 확인`"))"
        $ws.Range("N$r:Q$r").Interior.Color = 16777164
        $ws.Range("N$r:Q$r").NumberFormat = "0"
        $ws.Range("R$r").Formula = "=IF(B$r=`"`",`"`",IF(AND(ISNUMBER(N$r),ISNUMBER(O$r),ISNUMBER(P$r),ISNUMBER(Q$r),N$r>=0,N$r<=10,O$r>=0,O$r<=10,P$r>=0,P$r<=15,Q$r>=0,Q$r<=15),`"정상`",`"점수 확인`"))"
    }

    # R-01/R-02: 위원 역할과 마스킹 식별표시만 저장한다. 실제 성명·연락처·서명은 입력·저장하지 않는다.
    $ws.Range("B56").Value2 = "6. 위원 반복행 (R-01 역할·직위 / R-02 마스킹 식별표시) — 실제 성명·연락처·서명 입력 금지"
    $ws.Range("B56").Font.Bold = $true
    $ws.Range("B57").Value2 = "역할·직위"
    $ws.Range("C57").Value2 = "마스킹 식별표시 (위원 ○○/위원 **만 허용)"
    $ws.Range("B57:C57").Font.Bold = $true
    for ($i = 0; $i -lt 10; $i++) {
        $r = 58 + $i
        $ws.Range("B$r:C$r").Interior.Color = 16777164
    }

    # R-07/K-03: 평가항목별 배점과 점수는 분리 저장하고 총점은 수식으로만 계산한다.
    $ws.Range("B70").Value2 = "7. 평가점수 반복행 (R-07 점수 / K-03 평가 총점 자동계산) — 업체·위원 식별정보는 입력하지 않음"
    $ws.Range("B70").Font.Bold = $true
    $ws.Range("B71").Value2 = "평가항목"
    $ws.Range("C71").Value2 = "배점"
    $ws.Range("D71").Value2 = "점수"
    $ws.Range("E71").Value2 = "점수 상태"
    $ws.Range("B71:E71").Font.Bold = $true
    for ($i = 0; $i -lt 10; $i++) {
        $r = 72 + $i
        $ws.Range("B$r:D$r").Interior.Color = 16777164
        $ws.Range("C$r:D$r").NumberFormat = "0.00"
        $ws.Range("E$r").Formula = "=IF(AND(B$r<>`"`",C$r<>`"`",D$r<>`"`"),IF(AND(ISNUMBER(C$r),ISNUMBER(D$r),C$r>=0,D$r>=0,D$r<=C$r),`"정상`",`"점수/배점 확인`"),`"`")"
    }
    $ws.Range("B82").Value2 = "평가 총점 (K-03)"
    $ws.Range("B82").Font.Bold = $true
    $ws.Range("D82").Formula = "=SUM(D72:D81)"
    $ws.Range("D82").NumberFormat = "0.00"
    $ws.Range("D82").Font.Bold = $true

    $ws.Columns.Item("A").ColumnWidth = 6
    $ws.Columns.Item("B").ColumnWidth = 20
    $ws.Columns.Item("C").ColumnWidth = 20
    $ws.Columns.Item("D").ColumnWidth = 14
    $ws.Columns.Item("E").ColumnWidth = 16
    $ws.Columns.Item("F").ColumnWidth = 16
    $ws.Columns.Item("N").ColumnWidth = 16; $ws.Columns.Item("O").ColumnWidth = 16; $ws.Columns.Item("P").ColumnWidth = 15; $ws.Columns.Item("Q").ColumnWidth = 17; $ws.Columns.Item("R").ColumnWidth = 13
    $ws.Range("C4:C9,C12:C18,C21:C25,B29:D38").Interior.Color = 16777164  # 연노랑 입력영역 표시
    L "기초자료입력 시트 작성 완료"

    # ---- 4. DB ----
    $wsDB = $wbNew.Worksheets.Add()
    $wsDB.Name = "DB"
    $ws = $wsDB
    $dbHeaders = @("순번","학교명","학년도","담당부서","담당자직위","문서발행일","문서번호","구매명","구매학년","계약방식","기초금액","예정수량","납품기한","제출기한","수신","제목","관련문서","붙임목록","안내문본문","저장시각")
    for ($i = 0; $i -lt $dbHeaders.Count; $i++) {
        $ws.Cells.Item(1, $i + 1).Value2 = $dbHeaders[$i]
        $ws.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $ws.Rows.Item(1).AutoFilter() | Out-Null
    L "DB 시트 작성 완료 (열 수: $($dbHeaders.Count))"

    # ---- 5. DB_품목 (반복행 정규화 저장) ----
    $wsItems = $wbNew.Worksheets.Add()
    $wsItems.Name = "DB_품목"
    $itemHeaders = @("레코드순번", "행번호", "품목명", "수량", "단가", "금액")
    for ($i = 0; $i -lt $itemHeaders.Count; $i++) {
        $wsItems.Cells.Item(1, $i + 1).Value2 = $itemHeaders[$i]
        $wsItems.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $wsItems.Rows.Item(1).AutoFilter() | Out-Null
    L "DB_품목 시트 작성 완료 (반복행 저장용)"

    # ---- 5a. DB_업체 (R-03 반복행 정규화 저장, 연락처·사업자번호 등 제외) ----
    $wsVendors = $wbNew.Worksheets.Add()
    $wsVendors.Name = "DB_업체"
    $vendorHeaders = @("레코드순번", "행번호", "업체명")
    for ($i = 0; $i -lt $vendorHeaders.Count; $i++) {
        $wsVendors.Cells.Item(1, $i + 1).Value2 = $vendorHeaders[$i]
        $wsVendors.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $wsVendors.Rows.Item(1).AutoFilter() | Out-Null
    L "DB_업체 시트 작성 완료 (업체명 반복행 저장용, 개인정보·연락처 제외)"

    # ---- 5b. DB_위원 (R-01/R-02 정규화 저장, 실제 성명·연락처·서명 제외) ----
    $wsCommittee = $wbNew.Worksheets.Add()
    $wsCommittee.Name = "DB_위원"
    $committeeHeaders = @("레코드순번", "행번호", "역할직위", "마스킹식별표시")
    for ($i = 0; $i -lt $committeeHeaders.Count; $i++) {
        $wsCommittee.Cells.Item(1, $i + 1).Value2 = $committeeHeaders[$i]
        $wsCommittee.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $wsCommittee.Rows.Item(1).AutoFilter() | Out-Null
    L "DB_위원 시트 작성 완료 (역할·마스킹 식별표시만 저장, 개인정보 제외)"

    # ---- 5c. DB_평가 (R-07/K-03 정규화 저장) ----
    $wsScore = $wbNew.Worksheets.Add()
    $wsScore.Name = "DB_평가"
    $scoreHeaders = @("레코드순번", "행번호", "평가항목", "배점", "점수")
    for ($i = 0; $i -lt $scoreHeaders.Count; $i++) {
        $wsScore.Cells.Item(1, $i + 1).Value2 = $scoreHeaders[$i]
        $wsScore.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $wsScore.Rows.Item(1).AutoFilter() | Out-Null
    L "DB_평가 시트 작성 완료 (평가항목·배점·점수 반복행 저장용)"

    # ---- 5d. DB_정량평가 (F-015 업체별 정량평가 점수 정규화 저장) ----
    $wsQuant = $wbNew.Worksheets.Add()
    $wsQuant.Name = "DB_정량평가"
    $quantHeaders = @("레코드순번", "행번호", "수행경험점수", "공인인증점수", "거리적접근성점수", "상한가격점수")
    for ($i = 0; $i -lt $quantHeaders.Count; $i++) {
        $wsQuant.Cells.Item(1, $i + 1).Value2 = $quantHeaders[$i]
        $wsQuant.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $wsQuant.Rows.Item(1).AutoFilter() | Out-Null
    L "DB_정량평가 시트 작성 완료 (F-015 업체별 정량평가 점수 저장용)"

    # ---- 5e. DB_정성평가 (F-016 업체별 정성평가 점수 정규화 저장) ----
    $wsQual = $wbNew.Worksheets.Add()
    $wsQual.Name = "DB_정성평가"
    $qualHeaders = @("레코드순번", "행번호", "재질점수", "완성도점수", "AS점수", "하자보상점수", "가감점")
    for ($i = 0; $i -lt $qualHeaders.Count; $i++) {
        $wsQual.Cells.Item(1, $i + 1).Value2 = $qualHeaders[$i]
        $wsQual.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $wsQual.Rows.Item(1).AutoFilter() | Out-Null
    L "DB_정성평가 시트 작성 완료 (F-016 업체별 정성평가 점수 저장용)"

    # ---- 5f. DB_자기평점 (F-018 업체 자기평점, 학교 평가 DB와 분리) ----
    $wsSelf = $wbNew.Worksheets.Add()
    $wsSelf.Name = "DB_자기평점"
    $selfHeaders = @("레코드순번", "행번호", "수행경험자기점수", "공인인증자기점수", "거리적접근성자기점수", "상한가격자기점수")
    for ($i = 0; $i -lt $selfHeaders.Count; $i++) {
        $wsSelf.Cells.Item(1, $i + 1).Value2 = $selfHeaders[$i]
        $wsSelf.Cells.Item(1, $i + 1).Font.Bold = $true
    }
    $wsSelf.Rows.Item(1).AutoFilter() | Out-Null
    L "DB_자기평점 시트 작성 완료 (F-018 업체 자기평점 저장용, 학교 평가와 분리)"

    # ---- 6. 서식선택_출력 ----
    $wsSel = $wbNew.Worksheets.Add()
    $wsSel.Name = "서식선택_출력"
    $ws = $wsSel
    $ws.Range("A1").Value2 = "서식 선택 · 출력"
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A2").Value2 = "선택(TRUE) 열에 TRUE를 입력한 뒤 매크로 버튼을 눌러 인쇄 미리보기/PDF 저장을 실행합니다. 구현상태가 'N'인 서식은 아직 연결된 시트가 없어 선택해도 건너뜁니다."

    $headers = @("선택","Form ID","서식명","업무단계","구현상태","연결시트")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $ws.Cells.Item(4, $i + 1).Value2 = $headers[$i]
        $ws.Cells.Item(4, $i + 1).Font.Bold = $true
    }

    # 서식_인벤토리.md F-001~F-052 (F-053~F-057 참고자료 제외) 그대로 반영
    $forms = @(
        @("F-001","교복선정위원회 구성 기안","준비"), @("F-002","위원 수락 및 확인서","준비"),
        @("F-003","위원 청렴 및 보안 서약서","준비"), @("F-004","구매 추진 계획 수립","준비"),
        @("F-005","구매 추진 계획안","준비"), @("F-006","학교운영위원회 심의(안)","준비"),
        @("F-007","교복구매 구매 요청","입찰"), @("F-008","교복 사양서","입찰"),
        @("F-009","기초금액 및 계약방법 결정","입찰"), @("F-010","사전규격공개 기안문","입찰"),
        @("F-011","교복 학교주관구매 입찰 공고","입찰"), @("F-012","계약 특수조건","입찰"),
        @("F-013","교복 디자인 및 규격서","입찰"), @("F-014","제안서 평가항목 및 배점기준","평가"),
        @("F-015","정량적 평가","평가"), @("F-016","정성적 평가","평가"),
        @("F-017","제출서류 자기확인서","입찰"), @("F-018","정량적 평가 자기 평점표","입찰"),
        @("F-019","입찰참가신청서","입찰"), @("F-020","입찰 참가 신고서","입찰"),
        @("F-021","교복 납품 제안서","입찰"), @("F-022","교복 납품 실적표","입찰"),
        @("F-023","교복 제조 사양서","입찰"), @("F-024","품목별 단가 비율표","입찰"),
        @("F-025","교복 A/S 계획서","입찰"), @("F-026","소비자 불만 처리 계획","입찰"),
        @("F-027","위임장","입찰"), @("F-028","서약서","입찰"),
        @("F-029","개인정보제공 동의서","입찰"), @("F-030","청렴계약 이행서약서","입찰"),
        @("F-031","교복 품목별 금액표","입찰"), @("F-032","제안서 접수 결과","평가"),
        @("F-033","제안서 접수대장","평가"), @("F-034","제안서 평가위원회 개최","평가"),
        @("F-035","제안서 정량평가 결과","평가"), @("F-036","제안서 평가 결과","평가"),
        @("F-037","평가위원회 참석 등록부","평가"), @("F-038","업체 참가 등록부","평가"),
        @("F-039","업체별 제안서 평가표","평가"), @("F-040","위원 청렴 및 보안 서약서","평가"),
        @("F-041","낙찰자 결정","계약"), @("F-042","낙찰자 결정 통보","계약"),
        @("F-043","계약체결","계약"), @("F-044","사전 안내 가정통신문 안내","구매"),
        @("F-045","교복구매 사전 안내 가정통신문","구매"), @("F-046","수요조사 가정통신문 안내","구매"),
        @("F-047","교복구매 수요조사 가정통신문","구매"), @("F-048","교복 구매 안내(신청 수량 파악)","구매"),
        @("F-049","만족도 설문조사 실시","사후평가"), @("F-050","만족도 조사 설문지","사후평가"),
        @("F-051","만족도 설문조사 결과","사후평가"), @("F-052","만족도 조사 설문 결과 서식","사후평가")
    )
    $implemented = @{ "F-007" = "F-007_구매요청기안문"; "F-014" = "F-014_평가항목배점기준"; "F-015" = "F-015_정량적평가"; "F-016" = "F-016_정성적평가"; "F-017" = "F-017_제출서류자기확인서"; "F-018" = "F-018_정량적평가자기평점표"; "F-024" = "F-024_단가비율표" }
    $deferred = @{ "F-013" = "HWPX 우선순위 위임(표·이미지 복합조판)" }

    $row = 5
    foreach ($f in $forms) {
        $fid = $f[0]
        $ws.Cells.Item($row, 1).Value2 = $false
        $ws.Cells.Item($row, 2).Value2 = $fid
        $ws.Cells.Item($row, 3).Value2 = $f[1]
        $ws.Cells.Item($row, 4).Value2 = $f[2]
        if ($implemented.ContainsKey($fid)) {
            $ws.Cells.Item($row, 5).Value2 = "Y"
            $ws.Cells.Item($row, 6).Value2 = $implemented[$fid]
        } elseif ($deferred.ContainsKey($fid)) {
            $ws.Cells.Item($row, 5).Value2 = "D"
            $ws.Cells.Item($row, 6).Value2 = $deferred[$fid]
        } else {
            $ws.Cells.Item($row, 5).Value2 = "N"
            $ws.Cells.Item($row, 6).Value2 = ""
        }
        $row++
    }
    $lastRow = $row - 1
    $ws.Range("A4:F$lastRow").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 8
    $ws.Columns.Item("B").ColumnWidth = 10
    $ws.Columns.Item("C").ColumnWidth = 32
    $ws.Columns.Item("D").ColumnWidth = 10
    $ws.Columns.Item("E").ColumnWidth = 10
    $ws.Columns.Item("F").ColumnWidth = 26
    $ws.Range("A5:A$lastRow").HorizontalAlignment = -4108  # center
    L "서식선택_출력 시트 작성 완료 (행 수: $($forms.Count))"

    # ---- 6. F-007_구매요청기안문 (A4 1쪽, 기안문) ----
    $wsF7 = $wbNew.Worksheets.Add()
    $wsF7.Name = "F-007_구매요청기안문"
    $ws = $wsF7
    $ws.Range("A1").Value2 = "[구조 초안 — HWPX 원본 문안 대조 검증 필요]"
    $ws.Range("A1").Font.Size = 8
    $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H3").Merge() | Out-Null
    $ws.Range("B3").Value2 = "교복 학교주관구매 구매 요청(안)"
    $ws.Range("B3").Font.Size = 16
    $ws.Range("B3").Font.Bold = $true
    $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Rows.Item(3).RowHeight = 30

    $ws.Range("B5").Value2 = "문서번호"
    $ws.Range("C5").Formula = "=기초자료입력!C9"
    $ws.Range("E5").Value2 = "시행일자"
    $ws.Range("F5").Formula = "=기초자료입력!C8"
    $ws.Range("F5").NumberFormat = "yyyy.mm.dd."

    $ws.Range("B7").Value2 = "수  신"
    $ws.Range("C7:H7").Merge() | Out-Null
    $ws.Range("C7").Formula = "=IF(기초자료입력!C21<>`"`",기초자료입력!C21,`"내부결재`")"

    $ws.Range("B8").Value2 = "제  목"
    $ws.Range("C8:H8").Merge() | Out-Null
    $ws.Range("C8").Formula = "=IF(기초자료입력!C22<>`"`",기초자료입력!C22,기초자료입력!C12&`" 구매 요청`")"
    $ws.Range("C8").Font.Bold = $true

    $ws.Range("B10:H10").Merge() | Out-Null
    $ws.Range("B10").Formula = "=기초자료입력!C4&`" `"&기초자료입력!C5&`"학년도 `"&기초자료입력!C12&`"을(를) 다음과 같이 학교주관구매 방식으로 구매하고자 요청합니다.`""
    $ws.Range("B10").WrapText = $true
    $ws.Rows.Item(10).RowHeight = 40
    $ws.Range("B10").VerticalAlignment = -4160

    $ws.Range("B12").Value2 = "1. 구매명"
    $ws.Range("C12:H12").Merge() | Out-Null
    $ws.Range("C12").Formula = "=기초자료입력!C12"

    $ws.Range("B13").Value2 = "2. 구매 학년"
    $ws.Range("C13:H13").Merge() | Out-Null
    $ws.Range("C13").Formula = "=기초자료입력!C13"

    $ws.Range("B14").Value2 = "3. 기초금액"
    $ws.Range("C14:H14").Merge() | Out-Null
    $ws.Range("C14").Formula = "=IF(기초자료입력!C15<>`"`",TEXT(기초자료입력!C15,`"#,##0`")&`"원`",`"`")"

    $ws.Range("B15").Value2 = "4. 계약방식"
    $ws.Range("C15:H15").Merge() | Out-Null
    $ws.Range("C15").Formula = "=기초자료입력!C14"

    $ws.Range("B16").Value2 = "5. 납품기한"
    $ws.Range("C16:H16").Merge() | Out-Null
    $ws.Range("C16").Formula = "=기초자료입력!C17"

    $ws.Range("B18").Value2 = "붙임"
    $ws.Range("C18:H18").Merge() | Out-Null
    $ws.Range("C18").Formula = "=기초자료입력!C24"

    $ws.Range("B26").Value2 = "관련문서: "
    $ws.Range("C26:H26").Merge() | Out-Null
    $ws.Range("C26").Formula = "=기초자료입력!C23"

    # 결재란 (빈칸 유지 — 서명 자동처리 금지)
    $ws.Range("F28").Value2 = "담당"
    $ws.Range("G28").Value2 = "행정실장"
    $ws.Range("H28").Value2 = "교장"
    $ws.Range("F28:H28").Font.Bold = $true
    $ws.Range("F28:H28").HorizontalAlignment = -4108
    $ws.Range("F29:H33").Borders.LineStyle = 1
    $ws.Rows.Item(29).RowHeight = 5
    for ($r=30; $r -le 33; $r++) { $ws.Rows.Item($r).RowHeight = 18 }

    $ws.Columns.Item("A").ColumnWidth = 2.5
    $ws.Columns.Item("B").ColumnWidth = 10
    for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 8.5 }
    $ws.Range("B5:H33").Font.Size = 10
    $ws.Range("B5,E5").Font.Bold = $true
    $ws.Range("B7,B8,B12,B13,B14,B15,B16,B18,B26").Font.Bold = $true

    # PageSetup — A4, 여백, 100% 배율(1페이지 자연 확정 목표)
    $ps = $ws.PageSetup
    $ps.PaperSize = 9  # xlPaperA4
    $ps.Orientation = 1  # xlPortrait
    $ps.TopMargin = CmToPt 1.5
    $ps.BottomMargin = CmToPt 1.5
    $ps.LeftMargin = CmToPt 1.5
    $ps.RightMargin = CmToPt 1.5
    $ps.HeaderMargin = CmToPt 0.8
    $ps.FooterMargin = CmToPt 0.8
    $ps.Zoom = 100
    $ps.CenterHorizontally = $false
    $ws.PageSetup.PrintArea = "`$A`$1:`$H`$33"
    $hb = $ws.HPageBreaks.Count
    $vb = $ws.VPageBreaks.Count
    L "F-007 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hb, VPageBreaks=$vb (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-007_구매요청기안문 시트 작성 완료"

    # ---- 7. F-024_단가비율표 (A4 1쪽, 단가표) ----
    $wsF24 = $wbNew.Worksheets.Add()
    $wsF24.Name = "F-024_단가비율표"
    $ws = $wsF24
    $ws.Range("A1").Value2 = "[구조 초안 — HWPX 원본 문안 대조 검증 필요]"
    $ws.Range("A1").Font.Size = 8
    $ws.Range("A1").Font.Color = 255

    $ws.Range("B3:G3").Merge() | Out-Null
    $ws.Range("B3").Value2 = "품목별 단가 비율표"
    $ws.Range("B3").Font.Size = 16
    $ws.Range("B3").Font.Bold = $true
    $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Rows.Item(3).RowHeight = 30

    $ws.Range("B5").Value2 = "학교명"
    $ws.Range("C5").Formula = "=기초자료입력!C4"
    $ws.Range("D5").Value2 = "학년도"
    $ws.Range("E5").Formula = "=기초자료입력!C5"
    $ws.Range("F5").Value2 = "구매명"
    $ws.Range("G5").Formula = "=기초자료입력!C12"

    $hdrRow = 7
    $headers24 = @("연번","품목명","수량","단가(원)","금액(원)","비율(%)")
    $cols24 = @("B","C","D","E","F","G")
    for ($i = 0; $i -lt $headers24.Count; $i++) {
        $ws.Range("$($cols24[$i])$hdrRow").Value2 = $headers24[$i]
        $ws.Range("$($cols24[$i])$hdrRow").Font.Bold = $true
        $ws.Range("$($cols24[$i])$hdrRow").HorizontalAlignment = -4108
    }
    for ($i = 0; $i -lt 10; $i++) {
        $r = $hdrRow + 1 + $i
        $srcRow = 29 + $i
        $ws.Range("B$r").Value2 = [double]($i + 1)
        $ws.Range("C$r").Formula = "=기초자료입력!B$srcRow"
        $ws.Range("D$r").Formula = "=기초자료입력!C$srcRow"
        $ws.Range("E$r").Formula = "=기초자료입력!D$srcRow"
        $ws.Range("E$r").NumberFormat = "#,##0"
        $ws.Range("F$r").Formula = "=기초자료입력!E$srcRow"
        $ws.Range("F$r").NumberFormat = "#,##0"
        $ws.Range("G$r").Formula = "=IF(AND(F`$$($hdrRow+11)<>0,ISNUMBER(F$r)),F$r/F`$$($hdrRow+11)*100,`"`")"
        $ws.Range("G$r").NumberFormat = "0.0"
    }
    $sumRow = $hdrRow + 11
    $ws.Range("C$sumRow").Value2 = "합계"
    $ws.Range("C$sumRow").Font.Bold = $true
    $ws.Range("F$sumRow").Formula = "=SUM(F$($hdrRow+1):F$($hdrRow+10))"
    $ws.Range("F$sumRow").NumberFormat = "#,##0"
    $ws.Range("F$sumRow").Font.Bold = $true
    $ws.Range("G$sumRow").Value2 = [double]100
    $ws.Range("G$sumRow").Font.Bold = $true

    $ws.Range("B$($hdrRow):G$sumRow").Borders.LineStyle = 1
    $ws.Range("B$($hdrRow):G$($hdrRow)").Interior.Color = 15987699

    $ws.Columns.Item("A").ColumnWidth = 2.5
    $ws.Columns.Item("B").ColumnWidth = 6
    $ws.Columns.Item("C").ColumnWidth = 20
    $ws.Columns.Item("D").ColumnWidth = 8
    $ws.Columns.Item("E").ColumnWidth = 12
    $ws.Columns.Item("F").ColumnWidth = 12
    $ws.Columns.Item("G").ColumnWidth = 10
    $ws.Range("B5:G$sumRow").Font.Size = 10

    $ps = $ws.PageSetup
    $ps.PaperSize = 9
    $ps.Orientation = 1
    $ps.TopMargin = CmToPt 1.5
    $ps.BottomMargin = CmToPt 1.5
    $ps.LeftMargin = CmToPt 1.5
    $ps.RightMargin = CmToPt 1.5
    $ps.HeaderMargin = CmToPt 0.8
    $ps.FooterMargin = CmToPt 0.8
    $ps.Zoom = 100
    $ws.PageSetup.PrintArea = "`$A`$1:`$G`$$sumRow"
    $hb = $ws.HPageBreaks.Count
    $vb = $ws.VPageBreaks.Count
    L "F-024 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hb, VPageBreaks=$vb (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-024_단가비율표 시트 작성 완료"

    # ---- 7a. F-007/F-024 v2 원문 대조 반영 ----
    # 위의 초안은 P3-01 초기 구조이다. 아래에서 2026-09-15 HWPX 대조 결과를
    # 반영한 최종 v2 레이아웃으로 두 시트를 다시 구성한다. 이 단계가 없으면
    # 빌더를 재실행했을 때 배포본의 원문 대조 수식/표 구조가 되돌아간다.
    $notice = "[검토중 — 담당자 최종 확인 후 사용] 상세 변경 이력: _workspace/03_excel/F-007_F024_수정후_재대조.md"

    $ws = $wsF7
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = $notice
    $ws.Range("A1").Font.Size = 8
    $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H3").Merge() | Out-Null
    $ws.Range("B3").Value2 = "교복 학교주관구매 구매 요청(안)"
    $ws.Range("B3").Font.Size = 16; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Formula = '=IF(기초자료입력!C21<>"",기초자료입력!C21,"내부결재")'
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,기초자료입력!C5&"학년도 "&기초자료입력!C12&" 구매 요청")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("C11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 관련："; $ws.Range("C11").Formula = '=IF(기초자료입력!C23<>"",기초자료입력!C23,"")'
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&기초자료입력!C5&"학년도 "&기초자료입력!C12&IF(기초자료입력!C12="","을",IF(MOD(UNICODE(RIGHT(기초자료입력!C12,1))-44032,28)=0,"를","을"))&" 위하여 교복 구매 계획에 대한 학교운영위원회 심의가 완료되어, 아래와 같이 구매 요청을 하고자 합니다."'; $ws.Range("B13").WrapText = $true
    $labelsF7 = @(@("B15","가. 대상 및 수량"), @("B16","나. 예상단가"), @("B17","다. 사양내역"), @("B18","라. 납품기한"), @("B20","붙임"))
    foreach ($item in $labelsF7) { $ws.Range($item[0]).Value2 = $item[1] }
    foreach ($r in 15,16,17,18,20) { $ws.Range("C$r:H$r").Merge() | Out-Null }
    $ws.Range("C15").Formula = '=기초자료입력!C5&"학년도 신입생 중 학교주관구매 참여자 "&IF(기초자료입력!C16<>"",기초자료입력!C16,"○○")&"명"'
    $ws.Range("C16").Formula = '="동복 "&IF(SUM(기초자료입력!D29:D32)=0,"○○○,○○○",TEXT(SUM(기초자료입력!D29:D32),"#,##0"))&"원, 하복 "&IF(SUM(기초자료입력!D33:D34)=0,"○○,○○○",TEXT(SUM(기초자료입력!D33:D34),"#,##0"))&"원"'
    $ws.Range("C17").Value2 = "붙임 참조"
    $ws.Range("C18").Formula = '=IF(기초자료입력!C17<>"",IFERROR(TEXT(기초자료입력!C17,"yyyy-mm-dd"),기초자료입력!C17),"")'
    $ws.Range("C20").Formula = '=IF(기초자료입력!C24<>"",기초자료입력!C24,"교복 사양서 1부")&"  끝."'
    $ws.Range("F23").Value2 = "담당"; $ws.Range("G23").Value2 = "협조자"; $ws.Range("H23").Value2 = "교장"
    $ws.Range("B24").Value2 = "시행"; $ws.Range("C24:E24").Merge() | Out-Null; $ws.Range("C24").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F24").Value2 = "접수"; $ws.Range("G24:H24").Merge() | Out-Null
    $ws.Range("B25").Value2 = "우편번호"; $ws.Range("D25").Value2 = "주소"; $ws.Range("E25:H25").Merge() | Out-Null
    $ws.Range("B26").Value2 = "전화"; $ws.Range("D26").Value2 = "전송(팩스)"; $ws.Range("F26").Value2 = "이메일"; $ws.Range("G26:H26").Merge() | Out-Null
    $ws.Range("B5:H26").Font.Size = 10; $ws.Range("B5,E5,B7,B8,B9,B11,B15,B16,B17,B18,B20,B24,F24,B25,D25,B26,D26,F26").Font.Bold = $true
    $ws.Range("B5:H26").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 11; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$26"

    $ws = $wsF24
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = $notice
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:G3").Merge() | Out-Null; $ws.Range("B3").Value2 = "[서식 7] 교복 품목별 단가 비율표"; $ws.Range("B3").Font.Size = 16; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B4:G4").Merge() | Out-Null; $ws.Range("B4").Value2 = "(업체 제출양식)"; $ws.Range("B4").HorizontalAlignment = -4108
    $ws.Range("B6").Value2 = "학교명"; $ws.Range("C6").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"")'; $ws.Range("D6").Value2 = "학년도"; $ws.Range("E6").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"")'; $ws.Range("F6").Value2 = "구매명"; $ws.Range("G6").Formula = '=IF(기초자료입력!C12<>"",기초자료입력!C12,"")'
    $ws.Range("B8:C8").Merge() | Out-Null; $ws.Range("B8").Value2 = "품목"; $ws.Range("D8").Value2 = "수량"; $ws.Range("E8").Value2 = "단가비율(%)"; $ws.Range("F8:G8").Merge() | Out-Null; $ws.Range("F8").Value2 = "비고"
    $ws.Range("B9:B12").Merge() | Out-Null; $ws.Range("B9").Value2 = "동복"; $ws.Range("B13:B14").Merge() | Out-Null; $ws.Range("B13").Value2 = "하복"; $ws.Range("F9:G14").Merge() | Out-Null; $ws.Range("F9").Value2 = "품목별 단가 비율 기준과 ±3% 이상 차이가 나는 품목이 있을 시 기준 미충족"; $ws.Range("F9").WrapText = $true
    $itemsF24 = @("후드 점퍼","집업티","맨투맨티","긴바지(치마)","반팔티","반바지")
    for ($i = 0; $i -lt 6; $i++) { $r = 9 + $i; $src = 29 + $i; $defaultName = $itemsF24[$i]; $ws.Range("C$r").Formula = "=IF(기초자료입력!B$src<>`"`",기초자료입력!B$src,`"$defaultName`")"; $ws.Range("D$r").Formula = "=IF(기초자료입력!C$src<>`"`",기초자료입력!C$src,1)"; $ws.Range("E$r").Formula = "=IF(SUM(기초자료입력!`$D`$29:`$D`$34)=0,`"`",TEXT(기초자료입력!D$src/SUM(기초자료입력!`$D`$29:`$D`$34)*100,`"0.0`")&`"%`")" }
    $ws.Range("B15:C15").Merge() | Out-Null; $ws.Range("B15").Value2 = "합계"; $ws.Range("D15").Formula = "=SUM(D9:D14)"; $ws.Range("E15").Formula = '="100%"'; $ws.Range("F15:G15").Merge() | Out-Null
    $ws.Range("B17:G17").Merge() | Out-Null; $ws.Range("B17").Value2 = "20○○년    월    일"; $ws.Range("B18:G18").Merge() | Out-Null; $ws.Range("B18").Value2 = "제안자 성  명                    (서명 또는 날인)"; $ws.Range("B19:G19").Merge() | Out-Null; $ws.Range("B19").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○○○학교")&"장 귀하"'
    $ws.Range("B6:G6,B8:G15").Borders.LineStyle = 1; $ws.Range("B8:G8").Font.Bold = $true; $ws.Range("B8:G8").HorizontalAlignment = -4108; $ws.Range("B15:G15").Font.Bold = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 9; $ws.Columns.Item("C").ColumnWidth = 16; $ws.Columns.Item("D").ColumnWidth = 9; $ws.Columns.Item("E").ColumnWidth = 14; $ws.Columns.Item("F").ColumnWidth = 17; $ws.Columns.Item("G").ColumnWidth = 10; $ws.Range("B3:G19").Font.Size = 10
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$G`$19"
    # ---- 7b. F-014 제안서 평가항목 및 배점기준 ----
    # 원본 HWPX [9-3]의 제목·예시안 안내·평가기준 표·합계 구조를 대조하여 재구성한다.
    # 학교별 평가기준은 R-07 입력 반복행으로만 관리하며, 원본의 고정 예시 항목을 자동 확정하지 않는다.
    $wsF14 = $wbNew.Worksheets.Add()
    $wsF14.Name = "F-014_평가항목배점기준"
    $ws = $wsF14
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX [9-3]의 예시안 구조를 대조하여 구성함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:G3").Merge() | Out-Null
    $ws.Range("B3").Value2 = "[붙임 3] 제안서 평가항목 및 배점기준"
    $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B4:G4").Merge() | Out-Null; $ws.Range("B4").Value2 = "(예시안)"; $ws.Range("B4").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "학교명"; $ws.Range("C5").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"")'
    $ws.Range("D5").Value2 = "학년도"; $ws.Range("E5").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"")'
    $ws.Range("F5").Value2 = "평가 총점"; $ws.Range("G5").Formula = '=IF(COUNTA(기초자료입력!B72:B81)>0,기초자료입력!D82,"")'
    $ws.Range("B6:G6").Merge() | Out-Null; $ws.Range("B6").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,"제안서 평가항목 및 배점기준")'; $ws.Range("B6").Font.Bold = $true
    $ws.Range("B7:G7").Merge() | Out-Null; $ws.Range("B7").Value2 = "※ 제안서 평가기준은 예시안으로 학교 실정에 맞게 변경하여 사용합니다. 출력 전 담당자가 평가구분·배점기준·가감점 적용 여부를 최종 확인하십시오."; $ws.Range("B7").WrapText = $true; $ws.Range("B7").Font.Size = 8
    $headersF14 = @("구분", "평가항목", "배점기준", "배점", "평가점수", "비고")
    for ($i = 0; $i -lt $headersF14.Count; $i++) { $ws.Cells.Item(9, $i + 2).Value2 = $headersF14[$i] }
    for ($i = 0; $i -lt 10; $i++) {
        $r = 10 + $i; $src = 72 + $i
        $ws.Range("B$r").Formula = "=IF(기초자료입력!B$src<>`"`",`"평가`",`"`")"
        $ws.Range("C$r").Formula = "=IF(기초자료입력!B$src<>`"`",기초자료입력!B$src,`"`")"
        $ws.Range("D$r").Value2 = "학교별 평가기준 확인"
        $ws.Range("E$r").Formula = "=IF(기초자료입력!C$src<>`"`",기초자료입력!C$src,`"`")"
        $ws.Range("F$r").Formula = "=IF(기초자료입력!D$src<>`"`",기초자료입력!D$src,`"`")"
        $ws.Range("G$r").Formula = "=IF(AND(기초자료입력!B$src<>`"`",기초자료입력!D$src<=기초자료입력!C$src),`"정상`",`"`")"
        $ws.Range("C$r:D$r").WrapText = $true
    }
    $ws.Range("B20:D20").Merge() | Out-Null; $ws.Range("B20").Value2 = "합 계"; $ws.Range("E20").Formula = "=SUM(E10:E19)"; $ws.Range("F20").Formula = "=기초자료입력!D82"; $ws.Range("G20").Value2 = "K-03"
    $ws.Range("B22:G22").Merge() | Out-Null; $ws.Range("B22").Value2 = "■ 평가기준에 대한 상세 설명"; $ws.Range("B22").Font.Bold = $true
    $ws.Range("B23:G23").Merge() | Out-Null; $ws.Range("B23").Value2 = "• 원본 예시안은 정량적 평가와 정성적 평가를 각 50점으로 구분하고, 수행경험·품질인증·거리·상한가격 및 재질·완성도·A/S·하자보상 항목을 제시합니다."
    $ws.Range("B24:G24").Merge() | Out-Null; $ws.Range("B24").Value2 = "• 현재 출력은 기초자료입력의 평가항목·배점·점수를 반영합니다. 학교별 배점기준, 가점·감점 및 평가구분은 원본 예시안과 최신 지침을 확인하여 확정하십시오."
    $ws.Range("B25:G25").Merge() | Out-Null; $ws.Range("B25").Value2 = "■ 평가 시 유의사항: 객관적인 자료와 제출된 제품의 품질을 기준으로 공정하게 평가하며, 최종 출력 전 평가기준과 점수의 사실관계를 확인하십시오."
    $ws.Range("B22:G25").WrapText = $true
    $ws.Range("B5:G5,B9:G20").Borders.LineStyle = 1; $ws.Range("B9:G9").Font.Bold = $true; $ws.Range("B9:G9").HorizontalAlignment = -4108; $ws.Range("B20:G20").Font.Bold = $true; $ws.Range("E10:F20").NumberFormat = "0.00"
    $ws.Range("B5:G25").VerticalAlignment = -4108; $ws.Range("B5,G5,B9:B20,E9:G20").HorizontalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 7; $ws.Columns.Item("C").ColumnWidth = 18; $ws.Columns.Item("D").ColumnWidth = 18; $ws.Columns.Item("E").ColumnWidth = 7; $ws.Columns.Item("F").ColumnWidth = 9; $ws.Columns.Item("G").ColumnWidth = 7; $ws.Range("B3:G25").Font.Size = 9
    $ws.Rows.Item(7).RowHeight = 30; for ($r = 10; $r -le 19; $r++) { $ws.Rows.Item($r).RowHeight = 24 }; $ws.Rows.Item(23).RowHeight = 30; $ws.Rows.Item(24).RowHeight = 30; $ws.Rows.Item(25).RowHeight = 30
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.HeaderMargin = CmToPt 0.8; $ps.FooterMargin = CmToPt 0.8; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$G`$25"
    L "F-007/F-014/F-024 원문 대조 레이아웃 및 수식 적용 완료"

    # ---- 7c. F-015 정량적 평가 ----
    # 원본 HWPX [9-4]의 고정 4개 평가항목(수행경험/공인인증/거리적접근성/상한가격)·배점구간을 그대로 재현한다.
    # 업체별 점수는 기초자료입력!C44:F53(업체 반복행 옆 열)에 입력하고, 이 시트는 I3 선택 순번이 가리키는
    # 업체 1곳의 점수만 표시하는 템플릿이다. 업체별 별도 쪽 출력은 Module_출력에서 이 시트를 업체 수만큼
    # 복사해 I3만 바꾼 뒤 하나의 PDF로 결합한다(F-007/F-014/F-024와 동일한 다중 시트 결합 방식).
    $wsF15 = $wbNew.Worksheets.Add()
    $wsF15.Name = "F-015_정량적평가"
    $ws = $wsF15
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX [9-4]의 고정 배점표를 대조하여 구성함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("I2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("I2").Font.Size = 7
    $ws.Range("I3").Value2 = 1
    $ws.Range("B3:G3").Merge() | Out-Null; $ws.Range("B3").Value2 = "[9-4] [붙임 3_1] [1단계] 정량적 평가"; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B4:G4").Merge() | Out-Null; $ws.Range("B4").Value2 = "[1단계] 정량적 평가(50점): 사업담당 또는 계약부서에서 평가"; $ws.Range("B4").Font.Bold = $true; $ws.Range("B4").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "학교명"; $ws.Range("C5").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"")'
    $ws.Range("D5").Value2 = "학년도"; $ws.Range("E5").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"")'
    $ws.Range("F5").Value2 = "평가 대상 업체"; $ws.Range("G5").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'
    $ws.Range("B6:G6").Merge() | Out-Null; $ws.Range("B6").Value2 = "※ 평가항목은 학교별 상황에 따라 자율적으로 변경 적용 가능함"; $ws.Range("B6").WrapText = $true
    $headersF15 = @("구분", "평가항목", "배점기준", "배점", "평가점수")
    for ($i = 0; $i -lt $headersF15.Count; $i++) { $ws.Cells.Item(8, $i + 2).Value2 = $headersF15[$i] }
    $ws.Range("B9:B21").Merge() | Out-Null; $ws.Range("B9").Value2 = "정량적 평가`n(50점)"; $ws.Range("B9").WrapText = $true
    $ws.Range("C9:C12").Merge() | Out-Null; $ws.Range("C9").Value2 = "1. 수행경험 (10)`n- 입찰공고일 기준 최근 3년간 학교 주관구매 교복 납품실적(해당학교 실적 증명서 첨부)"
    $ws.Range("D9").Value2 = "5건 이상"; $ws.Range("E9").Value2 = 10
    $ws.Range("D10").Value2 = "3건 이상"; $ws.Range("E10").Value2 = 8
    $ws.Range("D11").Value2 = "1건 이상 또는 해당연도 신규업체"; $ws.Range("E11").Value2 = 6
    $ws.Range("D12").Value2 = "실적없음"; $ws.Range("E12").Value2 = 4
    $ws.Range("F9:F12").Merge() | Out-Null; $ws.Range("F9").Formula = '=IFERROR(INDEX(기초자료입력!$C$44:$C$53,I3),"")'
    $ws.Range("C13:C15").Merge() | Out-Null; $ws.Range("C13").Value2 = "2. 공인인증(시험)기관의 품질인증 여부(10)`n[한국의류시험연구원, FITI시험연구원, KOLAS기관의 시험성적서 등]`n※ Q 마크 등 검사 인증"
    $ws.Range("D13").Value2 = "교복 전품목 인증"; $ws.Range("E13").Value2 = 10
    $ws.Range("D14").Value2 = "일부 품목 인증"; $ws.Range("E14").Value2 = 8
    $ws.Range("D15").Value2 = "전품목 미인증"; $ws.Range("E15").Value2 = 5
    $ws.Range("F13:F15").Merge() | Out-Null; $ws.Range("F13").Formula = '=IFERROR(INDEX(기초자료입력!$D$44:$D$53,I3),"")'
    $ws.Range("C16:C19").Merge() | Out-Null; $ws.Range("C16").Value2 = "3. 거리적 접근성(15)`n- 학교와 교복업체(매장)와의 거리(사업자등록증의 사업장주소지 확인 및, 실제 매장 운영여부 반드시 현장 확인)"
    $ws.Range("D16").Value2 = "이동거리 5km 이내"; $ws.Range("E16").Value2 = 15
    $ws.Range("D17").Value2 = "이동거리 5km이상~10km미만"; $ws.Range("E17").Value2 = 12
    $ws.Range("D18").Value2 = "이동거리 10km이상~15km미만"; $ws.Range("E18").Value2 = 9
    $ws.Range("D19").Value2 = "이동거리 15km이상"; $ws.Range("E19").Value2 = 5
    $ws.Range("F16:F19").Merge() | Out-Null; $ws.Range("F16").Formula = '=IFERROR(INDEX(기초자료입력!$E$44:$E$53,I3),"")'
    $ws.Range("C20:C21").Merge() | Out-Null; $ws.Range("C20").Value2 = "4. 품목별 상한가격(15)"
    $ws.Range("D20").Value2 = "전품목 상한가격 기준 충족"; $ws.Range("E20").Value2 = 15
    $ws.Range("D21").Value2 = "기준 미충족 품목이 있을 시"; $ws.Range("E21").Value2 = 0
    $ws.Range("F20:F21").Merge() | Out-Null; $ws.Range("F20").Formula = '=IFERROR(INDEX(기초자료입력!$F$44:$F$53,I3),"")'
    $ws.Range("B22:C22").Merge() | Out-Null; $ws.Range("B22").Value2 = "합계"
    $ws.Range("D22:E22").Merge() | Out-Null; $ws.Range("D22").Value2 = "50점 만점"
    $ws.Range("F22").Formula = "=IF(OR(F9=`"`",F13=`"`",F16=`"`",F20=`"`"),`"`",F9+F13+F16+F20)"
    $ws.Range("B24:G24").Merge() | Out-Null; $ws.Range("B24").Formula = '="본인은 "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○○○학교")&" 교복 학교주관 교복 구매업체 선정을 위한 제시된 항목에 따라 객관적이고 공정하게 심사할 것을 약속합니다."'; $ws.Range("B24").WrapText = $true
    $ws.Range("B25:G25").Merge() | Out-Null; $ws.Range("B25").Value2 = "20○○년    월    일"; $ws.Range("B25").HorizontalAlignment = -4108
    $ws.Range("B26:G26").Merge() | Out-Null; $ws.Range("B26").Value2 = "○○학교 사업담당자(계약담당자) (인 또는 서명)"; $ws.Range("B26").HorizontalAlignment = -4108
    $ws.Range("B5:G5,B8:G22").Borders.LineStyle = 1; $ws.Range("B8:F8").Font.Bold = $true; $ws.Range("B8:F8").HorizontalAlignment = -4108; $ws.Range("B22:F22").Font.Bold = $true
    $ws.Range("B9:G22").VerticalAlignment = -4108; $ws.Range("B9:B22,E9:F22").HorizontalAlignment = -4108; $ws.Range("C9:D21").WrapText = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 7; $ws.Columns.Item("C").ColumnWidth = 20; $ws.Columns.Item("D").ColumnWidth = 18; $ws.Columns.Item("E").ColumnWidth = 6; $ws.Columns.Item("F").ColumnWidth = 8; $ws.Columns.Item("G").ColumnWidth = 6; $ws.Range("B3:G26").Font.Size = 9
    $ws.Rows.Item(6).RowHeight = 24; for ($r = 9; $r -le 21; $r++) { $ws.Rows.Item($r).RowHeight = 22 }; $ws.Rows.Item(24).RowHeight = 30
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.HeaderMargin = CmToPt 0.8; $ps.FooterMargin = CmToPt 0.8; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$G`$26"
    L "F-015 원문 고정 배점표 대조 레이아웃 및 수식 적용 완료"

    # ---- 7d. F-016 정성적 평가 ----
    # 원본 HWPX [9-5]의 고정 4개 평가항목(재질/완성도/A·S/하자보상)·5단계 배점(탁월~불량)과
    # 가점·감점 요인(담합 처분/만족도 조사)을 재현한다. 5단계는 담당자가 최종 점수만 직접 입력하는
    # 방식으로 단순화하고(항목별 허용 범위로 검증), 가점·감점은 단일 입력칸(-15~+5)으로 처리한다.
    # F-015와 동일하게 업체별 별도 쪽으로 출력하며, 평가 주체는 교복선정위원회임을 명시한다.
    $wsF16 = $wbNew.Worksheets.Add()
    $wsF16.Name = "F-016_정성적평가"
    $ws = $wsF16
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX [9-5]의 고정 배점표를 대조하여 구성함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("I2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("I2").Font.Size = 7
    $ws.Range("I3").Value2 = 1
    $ws.Range("B3:F3").Merge() | Out-Null; $ws.Range("B3").Value2 = "[9-5] [붙임 3_2] [2단계] 정성적 평가"; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B4:F4").Merge() | Out-Null; $ws.Range("B4").Value2 = "[2단계] 정성적 평가(50점): 교복선정위원회에서 평가"; $ws.Range("B4").Font.Bold = $true; $ws.Range("B4").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "학교명"; $ws.Range("C5").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"")'
    $ws.Range("D5").Value2 = "학년도"; $ws.Range("E5").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"")'
    $ws.Range("F5").Value2 = "대상 업체"
    $ws.Range("B6:F6").Merge() | Out-Null; $ws.Range("B6").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'; $ws.Range("B6").HorizontalAlignment = -4108
    $ws.Range("B7:F7").Merge() | Out-Null; $ws.Range("B7").Value2 = "※ 평가항목은 학교별 상황에 따라 자율적으로 변경 적용 가능함"; $ws.Range("B7").WrapText = $true
    $headersF16 = @("구분", "평가항목", "배점기준", "배점", "평가점수")
    for ($i = 0; $i -lt $headersF16.Count; $i++) { $ws.Cells.Item(8, $i + 2).Value2 = $headersF16[$i] }
    $ws.Range("B9:B13").Merge() | Out-Null; $ws.Range("B9").Value2 = "정성적 평가`n(50점)"; $ws.Range("B9").WrapText = $true
    $ws.Range("C9").Value2 = "재질(15점)"; $ws.Range("D9").Value2 = "1. 옷감의 촉감과 질감의 상태(10)`n2. 섬유조직의 세밀함과 부드러움(10)`n탁월15·우수12·보통9·미흡6·불량3"; $ws.Range("E9").Value2 = 15
    $ws.Range("F9").Formula = '=IFERROR(INDEX(기초자료입력!$H$44:$H$53,I3),"")'
    $ws.Range("C10").Value2 = "완성도(10점)"; $ws.Range("D10").Value2 = "1. 바느질의 꼼꼼함`n2. 단추·지퍼의 견고함`n3. 마감처리의 완성도`n4. 여유단의 정도`n탁월10·우수8·보통6·미흡4·불량2"; $ws.Range("E10").Value2 = 10
    $ws.Range("F10").Formula = '=IFERROR(INDEX(기초자료입력!$I$44:$I$53,I3),"")'
    $ws.Range("C11").Value2 = "A/S(15점)"; $ws.Range("D11").Value2 = "1. 무상 A/S 의무기간`n2. A/S의 편의성(신속성)`n3. 출장 A/S 가능 여부`n4. A/S 내용(바느질·단추·지퍼 등)`n탁월15·우수12·보통9·미흡6·불량3"; $ws.Range("E11").Value2 = 15
    $ws.Range("F11").Formula = '=IFERROR(INDEX(기초자료입력!$J$44:$J$53,I3),"")'
    $ws.Range("C12").Value2 = "하자보상(10점)"; $ws.Range("D12").Value2 = "1. 제품하자에 대한 교환 등 하자 보상 및 소비자 불만 사항 처리 방안의 적정성`n탁월10·우수8·보통6·미흡4·불량2"; $ws.Range("E12").Value2 = 10
    $ws.Range("F12").Formula = '=IFERROR(INDEX(기초자료입력!$K$44:$K$53,I3),"")'
    $ws.Range("C13").Value2 = "가점 및 감점 요인"; $ws.Range("D13").Value2 = "· 최근 1년 공정거래위원회 담합 행정처분: -10`n· 최근 3년 교복 만족도 조사: 80점 이상 +5 / 75점 이상 +2 / 65점 이상 0 / 60점 이상 -2 / 55점 미만 -5"; $ws.Range("E13").Value2 = "-15~+5"
    $ws.Range("F13").Formula = '=IFERROR(INDEX(기초자료입력!$L$44:$L$53,I3),"")'
    $ws.Range("B14:C14").Merge() | Out-Null; $ws.Range("B14").Value2 = "합계"
    $ws.Range("D14:E14").Merge() | Out-Null; $ws.Range("D14").Value2 = "50점 만점"
    $ws.Range("F14").Formula = "=IF(OR(F9=`"`",F10=`"`",F11=`"`",F12=`"`",F13=`"`"),`"`",F9+F10+F11+F12+F13)"
    $ws.Range("B16:F16").Merge() | Out-Null; $ws.Range("B16").Formula = '="본인은 "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○○○학교")&" 교복 학교주관 교복 구매업체 선정을 위한 제시된 항목에 따라 객관적이고 공정하게 심사할 것을 약속합니다."'; $ws.Range("B16").WrapText = $true
    $ws.Range("B17:F17").Merge() | Out-Null; $ws.Range("B17").Value2 = "20○○년    월    일"; $ws.Range("B17").HorizontalAlignment = -4108
    $ws.Range("B18:F18").Merge() | Out-Null; $ws.Range("B18").Value2 = "○○학교 교복선정위원회 평가위원 (인 또는 서명)"; $ws.Range("B18").HorizontalAlignment = -4108
    $ws.Range("B5:F5,B8:F14").Borders.LineStyle = 1; $ws.Range("B8:F8").Font.Bold = $true; $ws.Range("B8:F8").HorizontalAlignment = -4108; $ws.Range("B14:F14").Font.Bold = $true
    $ws.Range("B9:F14").VerticalAlignment = -4108; $ws.Range("B9:C13,E9:F14").HorizontalAlignment = -4108; $ws.Range("D9:D13").WrapText = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 12; $ws.Columns.Item("C").ColumnWidth = 14; $ws.Columns.Item("D").ColumnWidth = 34; $ws.Columns.Item("E").ColumnWidth = 8; $ws.Columns.Item("F").ColumnWidth = 8; $ws.Range("B3:F18").Font.Size = 9
    $ws.Rows.Item(7).RowHeight = 24; for ($r = 9; $r -le 13; $r++) { $ws.Rows.Item($r).RowHeight = 45 }; $ws.Rows.Item(16).RowHeight = 30
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.HeaderMargin = CmToPt 0.8; $ps.FooterMargin = CmToPt 0.8; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$F`$18"
    L "F-016 원문 고정 배점표 대조 레이아웃 및 수식 적용 완료"

    # ---- 7e. F-017 제출서류 자기확인서 ----
    # 원본 HWPX [9-6]의 업체별 제출서류 확인표를 재현한다. 업체명만 기존 R-03에서
    # 반영하며 대표자·연락처·대리인 정보는 개인정보 경계에 따라 빈 양식으로 둔다.
    $wsF17 = $wbNew.Worksheets.Add()
    $wsF17.Name = "F-017_제출서류자기확인서"
    $ws = $wsF17
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX [9-6] 대조. 대표자·연락처·대리인 정보는 자동 반영하지 않음"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("G2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("G2").Font.Size = 7
    $ws.Range("G3").Value2 = 1
    $ws.Range("B3:E3").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"년도 신입생 교복 제출서류 자기확인서"'; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5:E5").Merge() | Out-Null; $ws.Range("B5").Value2 = "※ 업체명만 기초자료의 반복행에서 반영합니다. 대표자·연락처·대리인 정보는 제출자가 작성하는 빈 칸으로 유지합니다."; $ws.Range("B5").Font.Size = 8; $ws.Range("B5").WrapText = $true
    $headersF17Vendor = @("업체명", "대표자", "연락처", "위임자(대리인 제출 시): 직위·이름·전화번호")
    for ($i = 0; $i -lt $headersF17Vendor.Count; $i++) { $ws.Cells.Item(7, $i + 2).Value2 = $headersF17Vendor[$i] }
    $ws.Range("B8").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$G$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$G$3)),"")'
    $ws.Range("C8:E8").ClearContents()
    $headersF17 = @("제출서류", "수량", "제출여부`n(○,✕)", "비고`n(서류 확인 사항)")
    for ($i = 0; $i -lt $headersF17.Count; $i++) { $ws.Cells.Item(10, $i + 2).Value2 = $headersF17[$i] }
    $documentsF17 = @(
        @("정량적 평가 자기 평점표[서식 1_1]", "1부", ""),
        @("입찰참가 신청서[서식 2]", "1부", ""),
        @("입찰 참가 신고서[서식 3]", "1부", ""),
        @("위임장[서식 10]", "1부", "대리인 재직증명서, 4대보험 가입증명서"),
        @("서약서[서식 11]", "1부", ""),
        @("개인정보제공 동의서[서식 12]", "1부", ""),
        @("청렴계약 이행서약서[서식 13]", "1부", ""),
        @("사업자등록증 사본", "1부", ""),
        @("중소기업확인서", "1부", ""),
        @("공정거래위원회 법위반사실확인서", "1부", ""),
        @("교복 납품 실적 증명서", "1부", ""),
        @("기타 평가기준에 따른 증명서류", "1부", "제안평가 반영 시 제출"),
        @("교복(동복·하복) 납품 제안서[서식 4]", "원본 1부, 사본 ○부", "사본 블라인드 처리"),
        @("교복 납품 실적표[서식 5]", "원본 1부, 사본 ○부", "사본 블라인드 처리"),
        @("교복 제조 사양서[서식 6]", "원본 1부, 사본 ○부", "사본 블라인드 처리"),
        @("품목별 단가 비율표[서식 7]", "원본 1부, 사본 ○부", "사본 블라인드 처리"),
        @("교복 A/S 계획서[서식 8]", "원본 1부, 사본 ○부", "사본 블라인드 처리"),
        @("소비자 불만 사항에 대한 처리방안[서식 9]", "원본 1부, 사본 ○부", "사본 블라인드 처리"),
        @("교복 제조 및 판매 시설 현황서", "원본 1부, 사본 ○부", "사본 블라인드 처리"),
        @("교복 샘플", "샘플 1부", "블라인드 처리")
    )
    for ($i = 0; $i -lt $documentsF17.Count; $i++) { $r = 11 + $i; $ws.Cells.Item($r, 2).Value2 = $documentsF17[$i][0]; $ws.Cells.Item($r, 3).Value2 = $documentsF17[$i][1]; $ws.Cells.Item($r, 4).Value2 = ""; $ws.Cells.Item($r, 5).Value2 = $documentsF17[$i][2] }
    $ws.Range("B7:E8,B10:E30").Borders.LineStyle = 1; $ws.Range("B7:E7,B10:E10").Font.Bold = $true; $ws.Range("B7:E7,B10:E10").HorizontalAlignment = -4108
    $ws.Range("B7:E30").VerticalAlignment = -4108; $ws.Range("C7:D30").HorizontalAlignment = -4108; $ws.Range("B7:E10").WrapText = $true; $ws.Range("B11:E30").WrapText = $false; $ws.Range("B11:E30").ShrinkToFit = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 24; $ws.Columns.Item("C").ColumnWidth = 10; $ws.Columns.Item("D").ColumnWidth = 10; $ws.Columns.Item("E").ColumnWidth = 25; $ws.Range("B3:E30").Font.Size = 8
    $ws.Rows.Item(5).RowHeight = 20; $ws.Rows.Item(7).RowHeight = 24; $ws.Rows.Item(8).RowHeight = 20; $ws.Rows.Item(10).RowHeight = 24; for ($r = 11; $r -le 30; $r++) { $ws.Rows.Item($r).RowHeight = 16 }; $ws.Rows.Item(22).RowHeight = 24
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 0.8; $ps.BottomMargin = CmToPt 0.8; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$E`$30"
    L "F-017 원문 제출서류 표 대조 레이아웃 및 개인정보 빈칸 경계 적용 완료"

    # ---- 7f. F-018 정량적 평가 자기 평점표 ----
    # 원본 HWPX [9-7]의 4개 정량 항목과 배점표를 재현한다. 학교의 F-015 평가는 참조하지 않고,
    # 업체가 입력한 N:Q 자기평점만 표시한다. 작성자·대표자·서명은 자동 반영하지 않는다.
    $wsF18 = $wbNew.Worksheets.Add()
    $wsF18.Name = "F-018_정량적평가자기평점표"
    $ws = $wsF18
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX [9-7] 대조. F-015 학교평가와 분리된 업체 자기평점이며 작성자·대표자·서명은 자동 반영하지 않음"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("I2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("I2").Font.Size = 7
    $ws.Range("I3").Value2 = 1
    $ws.Range("B3:F3").Merge() | Out-Null; $ws.Range("B3").Value2 = "[9-7] [서식 1_1] 정량적 평가 자기 평점표"; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B4:F4").Merge() | Out-Null; $ws.Range("B4").Value2 = "※ 업체 자기평점표입니다. 학교의 F-015 정량적 평가 점수와 별도로 입력·보관하며, 사실관계와 증빙은 제출자가 최종 확인합니다."; $ws.Range("B4").WrapText = $true
    $ws.Range("B5").Value2 = "업체명"; $ws.Range("C5:F5").Merge() | Out-Null; $ws.Range("C5").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$I$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$I$3)),"")'
    $headersF18 = @("구분", "평가항목", "배점기준", "배점", "자기 평점")
    for ($i = 0; $i -lt $headersF18.Count; $i++) { $ws.Cells.Item(7, $i + 2).Value2 = $headersF18[$i] }
    $ws.Range("B8:B20").Merge() | Out-Null; $ws.Range("B8").Value2 = "정량적 평가`n(50점)"; $ws.Range("B8").WrapText = $true
    $ws.Range("C8:C11").Merge() | Out-Null; $ws.Range("C8").Value2 = "1. 수행경험 (10)`n최근 3년간 학교 주관구매 교복 납품실적"; $ws.Range("D8").Value2 = "5건 이상"; $ws.Range("E8").Value2 = 10; $ws.Range("D9").Value2 = "3건 이상"; $ws.Range("E9").Value2 = 8; $ws.Range("D10").Value2 = "1건 이상 또는 신규업체"; $ws.Range("E10").Value2 = 6; $ws.Range("D11").Value2 = "실적없음"; $ws.Range("E11").Value2 = 4; $ws.Range("F8:F11").Merge() | Out-Null; $ws.Range("F8").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$I$3)=0,"",INDEX(기초자료입력!$N$44:$N$53,$I$3)),"")'
    $ws.Range("C12:C14").Merge() | Out-Null; $ws.Range("C12").Value2 = "2. 공인인증(시험)기관의 품질인증 여부(10)"; $ws.Range("D12").Value2 = "교복 전품목 인증"; $ws.Range("E12").Value2 = 10; $ws.Range("D13").Value2 = "일부 품목 인증"; $ws.Range("E13").Value2 = 8; $ws.Range("D14").Value2 = "전품목 미인증"; $ws.Range("E14").Value2 = 5; $ws.Range("F12:F14").Merge() | Out-Null; $ws.Range("F12").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$I$3)=0,"",INDEX(기초자료입력!$O$44:$O$53,$I$3)),"")'
    $ws.Range("C15:C18").Merge() | Out-Null; $ws.Range("C15").Value2 = "3. 거리적 접근성(15)`n학교와 교복업체(매장)와의 거리"; $ws.Range("D15").Value2 = "이동거리 5km 이내"; $ws.Range("E15").Value2 = 15; $ws.Range("D16").Value2 = "5km 이상~10km 미만"; $ws.Range("E16").Value2 = 12; $ws.Range("D17").Value2 = "10km 이상~15km 미만"; $ws.Range("E17").Value2 = 9; $ws.Range("D18").Value2 = "15km 이상"; $ws.Range("E18").Value2 = 5; $ws.Range("F15:F18").Merge() | Out-Null; $ws.Range("F15").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$I$3)=0,"",INDEX(기초자료입력!$P$44:$P$53,$I$3)),"")'
    $ws.Range("C19:C20").Merge() | Out-Null; $ws.Range("C19").Value2 = "4. 품목별 상한가격(15)"; $ws.Range("D19").Value2 = "전품목 상한가격 기준 충족"; $ws.Range("E19").Value2 = 15; $ws.Range("D20").Value2 = "기준 미충족 품목이 있을 시"; $ws.Range("E20").Value2 = 0; $ws.Range("F19:F20").Merge() | Out-Null; $ws.Range("F19").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$I$3)=0,"",INDEX(기초자료입력!$Q$44:$Q$53,$I$3)),"")'
    $ws.Range("B21:C21").Merge() | Out-Null; $ws.Range("B21").Value2 = "합계"; $ws.Range("D21:E21").Merge() | Out-Null; $ws.Range("D21").Value2 = "50점 만점"; $ws.Range("F21").Formula = '=IF(OR(F8="",F12="",F15="",F19=""),"",F8+F12+F15+F19)'
    $ws.Range("B23:F23").Merge() | Out-Null; $ws.Range("B23").Value2 = "업체명은 기초자료의 업체 반복행에서만 반영합니다. 작성자·대표자·서명은 제출자가 직접 작성하는 빈 칸으로 유지합니다."; $ws.Range("B23").WrapText = $true
    $ws.Range("B25:F25").Merge() | Out-Null; $ws.Range("B25").Formula = '="업체명: "&IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$I$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$I$3)),"")&"     작성자:                         대표자:                         (인)"'
    $ws.Range("B5:F5,B7:F21").Borders.LineStyle = 1; $ws.Range("B7:F7").Font.Bold = $true; $ws.Range("B7:F7").HorizontalAlignment = -4108; $ws.Range("B21:F21").Font.Bold = $true
    $ws.Range("B5:F25").VerticalAlignment = -4108; $ws.Range("B5,F5,B7:B21,E7:F21").HorizontalAlignment = -4108
    $ws.Range("C8:D20").WrapText = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 7; $ws.Columns.Item("C").ColumnWidth = 21; $ws.Columns.Item("D").ColumnWidth = 19; $ws.Columns.Item("E").ColumnWidth = 7; $ws.Columns.Item("F").ColumnWidth = 9; $ws.Range("B3:F25").Font.Size = 9
    $ws.Rows.Item(4).RowHeight = 28; for ($r = 8; $r -le 20; $r++) { $ws.Rows.Item($r).RowHeight = 22 }; $ws.Rows.Item(23).RowHeight = 28
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.HeaderMargin = CmToPt 0.8; $ps.FooterMargin = CmToPt 0.8; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$F`$25"
    L "F-018 원문 고정 배점표 대조 레이아웃 및 학교평가 분리 경계 적용 완료"

    # ---- 8. 학교정보 (원본 공개 데이터 표본 복사 — 학생·학부모 개인정보 아님) ----
    $wsSchool = $wbNew.Worksheets.Add()
    $wsSchool.Name = "학교정보"
    $ws = $wsSchool
    $sampleRows = 120  # 표본(전체 1,516행 중 일부) — 전체가 필요하면 후속 작업에서 확장
    # Excel 클립보드·배열 대입은 자동화 세션에서 대기하거나 형 변환에 실패할 수 있어
    # 값만 셀 단위로 복사함. 이 표본은 검색용 공개 기관정보이므로 원본의 스타일·외부 연결은 이식하지 않음.
    for ($sourceRow = 1; $sourceRow -le ($sampleRows + 1); $sourceRow++) {
        for ($sourceColumn = 1; $sourceColumn -le 7; $sourceColumn++) {
            $targetCell = $ws.Cells.Item($sourceRow, $sourceColumn)
            $targetCell.Value2 = [string]$wsSrcSchool.Cells.Item($sourceRow, $sourceColumn).Text
        }
    }
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.Columns.Item("A").ColumnWidth = 8
    $ws.Columns.Item("B").ColumnWidth = 10
    $ws.Columns.Item("C").ColumnWidth = 8
    $ws.Columns.Item("D").ColumnWidth = 8
    $ws.Columns.Item("E").ColumnWidth = 24
    $ws.Columns.Item("F").ColumnWidth = 34
    $ws.Columns.Item("G").ColumnWidth = 14
    $ws.Range("A123").Value2 = "학교정보 표본 행 수"
    $ws.Range("B123").Value2 = [string]$sampleRows
    $ws.Range("A123:B123").Font.Size = 8
    L "학교정보 시트 작성 완료 (원본 $sampleRows 행 표본 복사, 전체 1516행 중 일부 — 공개 기관정보, 개인정보 아님)"
    $wbSrc.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbSrc) | Out-Null
    $wbSrc = $null
    L "학교정보 시드 통합문서 닫기 완료"

    # ---- 9. 학교검색 (120행 공개 기관정보 표본의 VBA 검색 UI) ----
    $wsSearch = $wbNew.Worksheets.Add()
    $wsSearch.Name = "학교검색"
    $ws = $wsSearch
    $ws.Range("A1").Value2 = "학교정보 검색 (학교명 · 지역 · 급별)"
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A2").Value2 = "노란색 조건칸에 하나 이상을 입력하고 [검색]을 누르십시오. 결과 행 하나를 선택한 뒤 [선택 학교를 기초자료에 반영]을 누르면 학교명만 기초자료입력!C4에 반영됩니다."
    $ws.Range("A3").Value2 = "학교명"
    $ws.Range("C3").Value2 = "지역"
    $ws.Range("E3").Value2 = "급별"
    $ws.Range("B3,D3,F3").Interior.Color = 16777164
    $searchHeaders = @("번호", "지역", "급별", "학교명", "주소", "연락처")
    for ($i = 0; $i -lt $searchHeaders.Count; $i++) { $ws.Cells.Item(6, $i + 1).Value2 = $searchHeaders[$i] }
    $ws.Range("A6:F6").Font.Bold = $true
    $ws.Range("A6:F6").Interior.Color = 15987699
    $ws.Range("A6:F126").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 8
    $ws.Columns.Item("B").ColumnWidth = 12
    $ws.Columns.Item("C").ColumnWidth = 10
    $ws.Columns.Item("D").ColumnWidth = 26
    $ws.Columns.Item("E").ColumnWidth = 38
    $ws.Columns.Item("F").ColumnWidth = 16
    L "학교검색 시트 작성 완료 (학교명/지역/급별 VBA 검색 및 C4 반영 UI)"

    # ---- 10. 절차안내 (교복구매 9단계 → Form ID 이동 UI) ----
    $wsFlow = $wbNew.Worksheets.Add()
    $wsFlow.Name = "절차안내"
    $ws = $wsFlow
    $ws.Range("A1").Value2 = "교복 학교주관구매 9단계 절차 안내"
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A2").Value2 = "단계 행을 선택하고 [관련 Form ID로 이동]을 누르면 서식선택_출력 시트의 대표 Form ID 행으로 이동합니다. 미구현 서식은 구현상태를 확인하고, 출력 전에는 행정실 담당자가 사실관계·계약조건·서식 상태를 최종 확인해야 합니다."
    $flowHeaders = @("단계", "해야 할 일", "관련 Form ID", "이동 Form ID")
    for ($i = 0; $i -lt $flowHeaders.Count; $i++) {
        $ws.Cells.Item(4, $i + 1).Value2 = $flowHeaders[$i]
        $ws.Cells.Item(4, $i + 1).Font.Bold = $true
    }
    $flowRows = @(
        @("1. 준비", "교복선정위원회 구성 및 구매 추진계획 수립", "F-001~F-006", "F-001"),
        @("2. 심의", "학교운영위원회 심의안 상정 및 심의", "F-006", "F-006"),
        @("3. 구매요청·기초조사", "구매 요청, 사양·기초금액·계약방법을 확정", "F-007~F-010", "F-007"),
        @("4. 입찰공고", "사전규격공개, 입찰공고, 특수조건 및 규격서를 확인", "F-011~F-013", "F-011"),
        @("5. 제안·접수", "입찰참가 및 제안서 제출·접수 서류를 관리", "F-017~F-033", "F-017"),
        @("6. 평가", "평가위원회 운영 및 정량·정성 평가를 수행", "F-014~F-016, F-034~F-040", "F-014"),
        @("7. 낙찰", "낙찰자 결정 및 통보", "F-041~F-042", "F-041"),
        @("8. 계약·납품", "계약 체결 및 납품 조건을 확인", "F-043", "F-043"),
        @("9. 구매안내·사후평가", "가정통신문, 수요·만족도 조사 및 결과 정리", "F-044~F-052", "F-044")
    )
    $flowRow = 5
    foreach ($flow in $flowRows) {
        for ($i = 0; $i -lt $flow.Count; $i++) { $ws.Cells.Item($flowRow, $i + 1).Value2 = [string]$flow[$i] }
        $flowRow++
    }
    $ws.Range("A4:D13").Borders.LineStyle = 1
    $ws.Range("A5:A13,D5:D13").HorizontalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 22
    $ws.Columns.Item("B").ColumnWidth = 52
    $ws.Columns.Item("C").ColumnWidth = 30
    $ws.Columns.Item("D").ColumnWidth = 14
    $ws.Range("A5:D13").VerticalAlignment = -4160
    $ws.Range("B5:C13").WrapText = $true
    L "절차안내 시트 작성 완료 (9단계 및 관련 Form ID 이동 UI)"

    # ---- 11. 계약방법안내 (교복 매뉴얼의 2단계 입찰 예시를 확인 후 반영) ----
    $wsMethod = $wbNew.Worksheets.Add()
    $wsMethod.Name = "계약방법안내"
    $ws = $wsMethod
    $ws.Range("A1").Value2 = "교복 학교주관구매 계약방법 안내"
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A2").Value2 = "교복 매뉴얼의 2단계 입찰(규격·가격 동시) 예시를 안내합니다. 이 시트는 계약금액·예외사유·법령 적용에 따른 계약방법을 자동 판정하지 않으며, 행정실 담당자가 최신 법령·지침과 사실관계를 확인한 뒤에만 반영해야 합니다."
    $ws.Range("A2:E2").Merge() | Out-Null
    $ws.Range("A2").WrapText = $true
    $ws.Rows.Item(2).RowHeight = 42
    $ws.Range("A4").Value2 = "확인 항목"
    $ws.Range("B4").Value2 = "확인 값"
    $ws.Range("C4").Value2 = "안내"
    $ws.Range("A4:C4").Font.Bold = $true
    $ws.Range("A4:C4").Interior.Color = 15987699
    $ws.Range("A5").Value2 = "규격(제안서) 심사가 필요한가"
    $ws.Range("A6").Value2 = "가격 경쟁 절차가 필요한가"
    $ws.Range("A7").Value2 = "안내된 계약방법"
    $ws.Range("B5:B6").Interior.Color = 16777164
    $ws.Range("B5:B6").Validation.Delete()
    $ws.Range("B5:B6").Validation.Add(3, 1, 1, "예,아니오")
    $ws.Range("B7").Formula = '=IF(AND(B5="예",B6="예"),"2단계 입찰(규격·가격 동시)","")'
    $ws.Range("C5").Value2 = "교복의 디자인·재질·바느질·A/S 등 품질 기준을 확인함"
    $ws.Range("C6").Value2 = "품질 적격 업체를 대상으로 가격입찰을 실시함"
    $ws.Range("C7").Value2 = "매뉴얼 예시: 지방계약법 시행령 제18조 제3항에 따른 2단계 입찰"
    $ws.Range("A4:C7").Borders.LineStyle = 1
    $ws.Range("A5:C7").VerticalAlignment = -4160
    $ws.Range("C5:C7").WrapText = $true
    $ws.Columns.Item("A").ColumnWidth = 32
    $ws.Columns.Item("B").ColumnWidth = 28
    $ws.Columns.Item("C").ColumnWidth = 62
    $ws.Range("A9").Value2 = "반영 전 확인"
    $ws.Range("A9").Font.Bold = $true
    $ws.Range("A10").Value2 = "두 확인값은 비어 있는 상태로 시작합니다. 담당자가 최신 법령·지침과 사실관계를 확인한 뒤 각각 '예'를 명시적으로 선택한 경우에만 [계약방법을 기초자료에 반영] 버튼이 B-03에 안내값을 기록합니다. 그 밖의 경우 자동 추천하지 않습니다."
    $ws.Range("A10:E10").Merge() | Out-Null
    $ws.Range("A10").WrapText = $true
    $ws.Rows.Item(10).RowHeight = 36
    L "계약방법안내 시트 작성 완료 (2단계 입찰 예시 확인 및 수동 검토 경계)"

    # ---- 내부 DB 보호: 사용자 직접 편집을 막고, 매크로만 UserInterfaceOnly로 기록하게 한다. ----
    # 비밀번호를 지정하지 않으며, DB는 불러오기 순번 확인용으로 일반 숨김, 반복 DB는 VeryHidden 처리한다.
    $wsDB.Visible = 0  # xlSheetHidden
    foreach ($internalSheet in @($wsDB, $wsItems, $wsVendors, $wsCommittee, $wsScore, $wsQuant, $wsQual, $wsSelf)) {
        $internalSheet.Protect("", $true, $true, $true, $true)
    }
    foreach ($internalSheet in @($wsItems, $wsVendors, $wsCommittee, $wsScore, $wsQuant, $wsQual, $wsSelf)) {
        $internalSheet.Visible = 2  # xlSheetVeryHidden
    }
    L "DB 및 반복 DB 시트 보호 완료 (DB=숨김, 반복 DB=VeryHidden, UserInterfaceOnly 매크로 기록 허용, 비밀번호 없음)"

    # ---- 저장 ----
    $wsFirst.Activate()
    $wbNew.SaveAs($outPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    L "저장 완료: $outPath"

} finally {
    if ($wbSrc) { $wbSrc.Close($false) }
    if ($wbNew) { $wbNew.Close($false) }
    $excel.Quit()
    if ($wbSrc) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbSrc) | Out-Null }
    if ($wbNew) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbNew) | Out-Null }
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
}

[System.IO.File]::WriteAllText($logPath, $log.ToString(), [System.Text.UTF8Encoding]::new($false))
