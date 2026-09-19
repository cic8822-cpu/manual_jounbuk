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
# F-003의 "1./2./3./4." 순번 목록형 문구를 셀에 쓸 때 Excel의 셀 값 자동완성(AutoComplete)이
# 같은 열의 앞선 값과 유사한 접두 패턴을 감지해 값을 예기치 않게 바꿔치기하는 현상이 재현되어
# 비활성화한다(COM 자동화 세션 전체에 적용, 다른 시트 입력에는 영향 없음 — 이 속성은 사용자가
# 직접 타이핑할 때의 UI 동작을 제어하며 COM Value2/Formula 대입 자체를 막지 않음).
try { $excel.EnableAutoComplete = $false } catch { $null = $_ }
$log = New-Object System.Text.StringBuilder
function L($s) {
    [void]$log.AppendLine($s)
    [System.IO.File]::WriteAllText($logPath, $log.ToString(), [System.Text.UTF8Encoding]::new($false))
}

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
    $ws.Range("A8").Value2 = "6. 구현된 서식은 서식선택_출력의 '구현상태=Y'로 표시됩니다. 나머지 서식은 순차 추가 예정입니다."
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
    $implemented = @{ "F-001" = "F-001_위원회구성기안"; "F-002" = "F-002_위원수락확인서"; "F-003" = "F-003_위원청렴보안서약서"; "F-004" = "F-004_구매추진계획수립기안문"; "F-005" = "F-005_구매추진계획안"; "F-006" = "F-006_운영위원회심의안"; "F-007" = "F-007_구매요청기안문"; "F-014" = "F-014_평가항목배점기준"; "F-015" = "F-015_정량적평가"; "F-016" = "F-016_정성적평가"; "F-017" = "F-017_제출서류자기확인서"; "F-018" = "F-018_정량적평가자기평점표"; "F-019" = "F-019_입찰참가신청서"; "F-020" = "F-020_입찰참가신고서"; "F-021" = "F-021_교복납품제안서"; "F-022" = "F-022_교복납품실적표"; "F-023" = "F-023_교복제조사양서"; "F-024" = "F-024_단가비율표"; "F-025" = "F-025_교복AS계획서"; "F-032" = "F-032_제안서접수결과"; "F-033" = "F-033_제안서접수대장"; "F-034" = "F-034_평가위원회개최"; "F-035" = "F-035_정량평가결과"; "F-036" = "F-036_제안서평가결과"; "F-037" = "F-037_평가위원회참석등록부"; "F-038" = "F-038_업체참가등록부"; "F-039" = "F-039_업체별제안서평가표"; "F-040" = "F-040_위원청렴보안서약서"; "F-041" = "F-041_낙찰자결정"; "F-042" = "F-042_낙찰자결정통보"; "F-043" = "F-043_계약체결"; "F-044" = "F-044_사전안내가정통신문안내"; "F-045" = "F-045_사전안내가정통신문"; "F-046" = "F-046_수요조사가정통신문안내" }
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

    # ---- 7g. F-019 입찰참가신청서 ----
    # 원문 HWPX [9-7] [서식 2]를 대조한다. 업체명·학교 공통정보만 반영하고 법인번호·주소·전화·주민번호·대리인·인감은 빈칸으로 보존한다.
    $wsF19 = $wbNew.Worksheets.Add()
    $wsF19.Name = "F-019_입찰참가신청서"
    $ws = $wsF19
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX [9-7] [서식 2] 대조. 민감정보와 날인란은 자동 반영하지 않습니다."
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("N2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("N2").Font.Size = 7; $ws.Range("N3").Value2 = 1
    $ws.Range("B3:J4").Merge() | Out-Null; $ws.Range("B3").Value2 = "입 찰 참 가 신 청 서`n※ 아래 사항 중 해당되는 경우에만 기재하시기 바랍니다."; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108; $ws.Range("B3").WrapText = $true
    $ws.Range("K3:L4").Merge() | Out-Null; $ws.Range("K3").Value2 = "처리기간`n즉 시"; $ws.Range("K3").HorizontalAlignment = -4108; $ws.Range("K3").VerticalAlignment = -4108; $ws.Range("K3").WrapText = $true
    $ws.Range("B6:C8").Merge() | Out-Null; $ws.Range("B6").Value2 = "신`n청`n인"; $ws.Range("B6").HorizontalAlignment = -4108; $ws.Range("B6").VerticalAlignment = -4108
    $ws.Range("D6:E6").Merge() | Out-Null; $ws.Range("D6").Value2 = "상호 또는`n법인명칭"; $ws.Range("F6:H6").Merge() | Out-Null; $ws.Range("F6").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$N$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$N$3)),"")'; $ws.Range("I6").Value2 = "법인등록번호"; $ws.Range("J6:L6").Merge() | Out-Null
    $ws.Range("D7:E7").Merge() | Out-Null; $ws.Range("D7").Value2 = "주 소"; $ws.Range("F7:H7").Merge() | Out-Null; $ws.Range("I7").Value2 = "전화번호"; $ws.Range("J7:L7").Merge() | Out-Null
    $ws.Range("D8:E8").Merge() | Out-Null; $ws.Range("D8").Value2 = "대표자"; $ws.Range("F8:H8").Merge() | Out-Null; $ws.Range("I8").Value2 = "주민등록번호"; $ws.Range("J8:L8").Merge() | Out-Null
    $ws.Range("B10:C11").Merge() | Out-Null; $ws.Range("B10").Value2 = "입찰`n개요"; $ws.Range("B10").HorizontalAlignment = -4108; $ws.Range("B10").VerticalAlignment = -4108
    $ws.Range("D10:E10").Merge() | Out-Null; $ws.Range("D10").Value2 = "입찰공고`n(지명)번호"; $ws.Range("F10:H10").Merge() | Out-Null; $ws.Range("F10").Formula = '=IF(기초자료입력!$C$4="","",기초자료입력!$C$4&" 공고 제20○○-00호")'; $ws.Range("I10").Value2 = "입찰일자"; $ws.Range("J10:L10").Merge() | Out-Null
    $ws.Range("D11:E11").Merge() | Out-Null; $ws.Range("D11").Value2 = "입찰건명"; $ws.Range("F11:L11").Merge() | Out-Null; $ws.Range("F11").Formula = '=IF(OR(기초자료입력!$C$4="",기초자료입력!$C$5=""),"",기초자료입력!$C$4&" "&기초자료입력!$C$5&"학년도 교복 학교주관 구매 입찰")'
    $ws.Range("B13:C13").Merge() | Out-Null; $ws.Range("B13").Value2 = "입찰보증금"; $ws.Range("D13:L13").Merge() | Out-Null; $ws.Range("D13").Value2 = "귀 교에서 시행하는 위 건명에 대해 입찰보증금 납부를 면제받고자 하며, 낙찰자 결정 통지를 받은 후 10일 이내에 계약을 체결하지 않을 시 입찰금액의 5/100에 해당하는 입찰보증금을 납부할 것을 확약합니다."; $ws.Range("D13").WrapText = $true
    $ws.Range("B15:C18").Merge() | Out-Null; $ws.Range("B15").Value2 = "대리인·`n사용인감"; $ws.Range("B15").HorizontalAlignment = -4108; $ws.Range("B15").VerticalAlignment = -4108
    $ws.Range("D15:H18").Merge() | Out-Null; $ws.Range("D15").Value2 = "본 입찰에 관한 일체의 권한을 다음의 자에게 위임합니다.`n성 명 :`n주민등록번호 :"; $ws.Range("D15").WrapText = $true
    $ws.Range("I15:L18").Merge() | Out-Null; $ws.Range("I15").Value2 = "본 입찰에 사용할 인감을 다음과 같이 신고합니다.`n`n`n(사용인감 날인)"; $ws.Range("I15").WrapText = $true
    $ws.Range("B20:L20").Merge() | Out-Null; $ws.Range("B20").Formula = '="본인은 위의 번호로 공고한 "&IF(기초자료입력!$C$4="","○○○○학교",기초자료입력!$C$4)&"의 2단계 입찰(규격·가격 동시 입찰)에 참가하고자 지방자치단체 입찰 및 계약집행기준 제8장 입찰유의서, 제9장 계약일반조건 및 입찰공고 사항을 모두 승낙하고 별첨 서류를 첨부하여 입찰 참가 신청을 합니다."'; $ws.Range("B20").WrapText = $true
    $ws.Range("B21:L21").Merge() | Out-Null; $ws.Range("B21").Value2 = "붙임서류 : 공고에 정한 서류"
    $ws.Range("B23:L23").Merge() | Out-Null; $ws.Range("B23").Formula = '=IF(OR(기초자료입력!$C$4="",기초자료입력!$C$5=""),"","20"&RIGHT(기초자료입력!$C$5,2)&". . .")'
    $ws.Range("B24:L24").Merge() | Out-Null; $ws.Range("B24").Formula = '="업체명: "&IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$N$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$N$3)),"")&"                 대표자:                         (인)"'
    $ws.Range("B25:L25").Merge() | Out-Null; $ws.Range("B25").Formula = '=IF(기초자료입력!$C$4="","○○○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'
    $ws.Range("B3:L4,B6:L8,B10:L11,B13:L13,B15:L18").Borders.LineStyle = 1; $ws.Range("B3:L25").VerticalAlignment = -4108; $ws.Range("B6:L18").HorizontalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2; $ws.Columns.Item("B").ColumnWidth = 5; $ws.Columns.Item("C").ColumnWidth = 5; $ws.Columns.Item("D").ColumnWidth = 8; $ws.Columns.Item("E").ColumnWidth = 8; $ws.Columns.Item("F").ColumnWidth = 10; $ws.Columns.Item("G").ColumnWidth = 10; $ws.Columns.Item("H").ColumnWidth = 8; $ws.Columns.Item("I").ColumnWidth = 9; $ws.Columns.Item("J").ColumnWidth = 8; $ws.Columns.Item("K").ColumnWidth = 8; $ws.Columns.Item("L").ColumnWidth = 8; $ws.Range("B3:L25").Font.Size = 9
    $ws.Rows.Item(1).RowHeight = 10; $ws.Rows.Item(2).RowHeight = 10; $ws.Rows.Item(3).RowHeight = 30; $ws.Rows.Item(4).RowHeight = 24; for ($r = 6; $r -le 11; $r++) { $ws.Rows.Item($r).RowHeight = 23 }; $ws.Rows.Item(13).RowHeight = 40; for ($r = 15; $r -le 18; $r++) { $ws.Rows.Item($r).RowHeight = 22 }; $ws.Rows.Item(20).RowHeight = 42
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 2; $ps.TopMargin = CmToPt 0.8; $ps.BottomMargin = CmToPt 0.8; $ps.LeftMargin = CmToPt 0.8; $ps.RightMargin = CmToPt 0.8; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$L`$25"
    L "F-019 원문 입찰참가신청서 표 대조 레이아웃 및 개인정보 빈칸 경계 적용 완료"

    # ---- 7h. F-020 입찰 참가 신고서 ----
    # 원문 HWPX [9-8] [서식 3]을 대조한다. 사업자 상호(R-03)만 반영하고 사업자 번호·대표자 성명·직인은 빈칸으로 보존한다.
    $wsF20 = $wbNew.Worksheets.Add()
    $wsF20.Name = "F-020_입찰참가신고서"
    $ws = $wsF20
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX [9-8] [서식 3] 대조. 사업자 번호·대표자 성명·직인은 자동 반영하지 않습니다."
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("G2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("G2").Font.Size = 7; $ws.Range("G3").Value2 = 1
    $ws.Range("B3:E4").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복 학교주관구매 입찰 참가 신고서"; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5:E5").Merge() | Out-Null; $ws.Range("B5").Value2 = "※ 사업자 상호(업체명)만 기초자료의 반복행에서 반영합니다. 사업자 번호·대표자 성명·직인은 사업자가 직접 작성하는 빈 칸으로 유지합니다."; $ws.Range("B5").Font.Size = 8; $ws.Range("B5").WrapText = $true
    $ws.Range("B7").Value2 = "사업자 상호"; $ws.Range("B7").Font.Bold = $true; $ws.Range("C7:E7").Merge() | Out-Null; $ws.Range("C7").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$G$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$G$3)),"")'
    $ws.Range("B8").Value2 = "사업자 번호"; $ws.Range("B8").Font.Bold = $true; $ws.Range("C8:E8").Merge() | Out-Null
    $ws.Range("B9").Value2 = "사업자 대표"; $ws.Range("B9").Font.Bold = $true; $ws.Range("C9:E9").Merge() | Out-Null
    $lq = [char]0x2018; $rq = [char]0x2019  # PowerShell 5.1은 한글 문서의 둥근 작은따옴표(U+2018/U+2019)를 문자열 구분자로 오인하므로 문자코드로 조립함
    $ws.Range("B11:E13").Merge() | Out-Null; $ws.Range("B11").Formula = '="위 사업자는 전북특별자치도교육청의 교복 가격 안정화를 위한 교복 ' + $lq + '학교주관구매' + $rq + ' 실시에 따라, "&IF(기초자료입력!$C$4="","○○○○학교",기초자료입력!$C$4)&"(주소지 전북 )가 시행하는 교복 ' + $lq + '학교주관구매' + $rq + '가 원만히 추진될 수 있도록 20○○.00.00. 일자로 공고한 입찰에 참가하기에 신고합니다."'; $ws.Range("B11").WrapText = $true
    $ws.Range("B15:E16").Merge() | Out-Null; $ws.Range("B15").Value2 = "※ 위 신고 사업자는 해당 학교의 " + $lq + "학교주관구매" + $rq + "를 저해하는 행위를 하는 경우, 향후 입찰에서 제외되는 등 불이익을 받을 수 있음을 확인하였습니다."; $ws.Range("B15").WrapText = $true
    $ws.Range("B18:E18").Merge() | Out-Null; $ws.Range("B18").Formula = '=IF(기초자료입력!$C$5="","20○○. ○○. ○○","20"&RIGHT(기초자료입력!$C$5,2)&". ○○. ○○")'; $ws.Range("B18").HorizontalAlignment = -4108
    $ws.Range("B19:E19").Merge() | Out-Null; $ws.Range("B19").Value2 = "위 신고인 사업자 대표 성명 ○○○ 사업자 직인"; $ws.Range("B19").HorizontalAlignment = -4108
    $ws.Range("B20:E20").Merge() | Out-Null; $ws.Range("B20").Formula = '=IF(기초자료입력!$C$4="","○○○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'; $ws.Range("B20").HorizontalAlignment = -4108
    $ws.Range("B3:E5,B7:E9,B11:E13,B15:E16,B18:E20").Borders.LineStyle = 1
    $ws.Range("B7:E9").VerticalAlignment = -4108; $ws.Range("B11:E16").VerticalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 24; $ws.Columns.Item("C").ColumnWidth = 10; $ws.Columns.Item("D").ColumnWidth = 10; $ws.Columns.Item("E").ColumnWidth = 25; $ws.Range("B3:E20").Font.Size = 10
    $ws.Rows.Item(1).RowHeight = 10; $ws.Rows.Item(2).RowHeight = 10; $ws.Rows.Item(3).RowHeight = 22; $ws.Rows.Item(4).RowHeight = 22; $ws.Rows.Item(5).RowHeight = 28; $ws.Rows.Item(6).RowHeight = 8; $ws.Rows.Item(7).RowHeight = 20; $ws.Rows.Item(8).RowHeight = 20; $ws.Rows.Item(9).RowHeight = 20; $ws.Rows.Item(10).RowHeight = 8; $ws.Rows.Item(11).RowHeight = 18; $ws.Rows.Item(12).RowHeight = 18; $ws.Rows.Item(13).RowHeight = 18; $ws.Rows.Item(14).RowHeight = 8; $ws.Rows.Item(15).RowHeight = 22; $ws.Rows.Item(16).RowHeight = 22; $ws.Rows.Item(17).RowHeight = 8; $ws.Rows.Item(18).RowHeight = 20; $ws.Rows.Item(19).RowHeight = 20; $ws.Rows.Item(20).RowHeight = 20
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 0.8; $ps.BottomMargin = CmToPt 0.8; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$E`$20"
    L "F-020 원문 입찰 참가 신고서 대조 레이아웃 및 개인정보 빈칸 경계 적용 완료"

    # ---- 7i. F-021 교복(동·하복) 납품 제안서 ----
    # 원문 HWPX [9-9] [서식 4]을 대조한다. 업체명(R-03)과 학교 공통값만 반영하고 대표자·주소·연락처·서명은 빈칸으로 보존한다.
    $wsF21 = $wbNew.Worksheets.Add(); $wsF21.Name = "F-021_교복납품제안서"; $ws = $wsF21
    $ws.Range("B1:E1").Merge() | Out-Null; $ws.Range("B1").Value2 = "[검토용] 원본 HWPX [9-9] [서식 4] 대조. 업체명만 자동 반영하며 대표자·주소·연락처·서명은 직접 작성합니다."; $ws.Range("B1").Font.Size = 7; $ws.Range("B1").Font.Color = 255; $ws.Range("B1").WrapText = $false
    $ws.Range("G2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("G2").Font.Size = 7; $ws.Range("G3").Value2 = 1
    $ws.Range("B3:E4").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!$C$5="","20○○학년도","20"&RIGHT(기초자료입력!$C$5,2)&"학년도")&" "&IF(기초자료입력!$C$4="","○○○○학교",기초자료입력!$C$4)&" 교복(동·하복) 납품 제안서"'; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108; $ws.Range("B3").WrapText = $true
    $ws.Range("B5:E5").Merge() | Out-Null; $ws.Range("B5").Value2 = "※ 업체명만 기초자료의 반복행에서 반영합니다. 일반현황·연혁·제안자 성명은 직접 작성하십시오."; $ws.Range("B5").Font.Size = 8; $ws.Range("B5").WrapText = $true
    $ws.Range("B7:E7").Merge() | Out-Null; $ws.Range("B7").Value2 = "1. 일반현황 및 연혁"; $ws.Range("B7").Font.Bold = $true
    $labels = @("회사명","대표자","사업분야","주소","연락처","회사설립년도","직원현황","해당부문 사업기간"); for ($i = 0; $i -lt $labels.Count; $i++) { $r = 8 + $i; $ws.Range("B$r").Value2 = $labels[$i]; $ws.Range("B$r").Font.Bold = $true; $ws.Range("C$r`:E$r").Merge() | Out-Null }
    $ws.Range("C8").Formula = '=IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$G$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$G$3)),"")'; $ws.Range("C12").Value2 = "전화:                         FAX:"
    $ws.Range("B16:E17").Merge() | Out-Null; $ws.Range("B16").Value2 = "＜주요연혁＞"; $ws.Range("B16").VerticalAlignment = -4108
    $ws.Range("B19:E19").Merge() | Out-Null; $ws.Range("B19").Value2 = "2. 제안(제출)서류"; $ws.Range("B19").Font.Bold = $true
    $ws.Range("B20:E27").Merge() | Out-Null; $ws.Range("B20").Value2 = "① 교복 납품 실적표(F-022)" + [Environment]::NewLine + "② 교복 제조 사양서(F-023)" + [Environment]::NewLine + "③ 품목별 단가 비율표(F-024)" + [Environment]::NewLine + "④ 교복 A/S 계획서(F-025)" + [Environment]::NewLine + "⑤ 소비자 불만 처리 계획(F-026)" + [Environment]::NewLine + "⑥ 위임장(F-027), 서약서(F-028), 개인정보 동의서(F-029)" + [Environment]::NewLine + "⑦ 품질인증·시험성적 등 제안요청서에서 정한 증빙서류"; $ws.Range("B20").Font.Size = 9; $ws.Range("B20").WrapText = $true; $ws.Range("B20").VerticalAlignment = -4160
    $ws.Range("B29:E29").Merge() | Out-Null; $ws.Range("B29").Formula = '=IF(기초자료입력!$C$5="","20○○. ○○. ○○","20"&RIGHT(기초자료입력!$C$5,2)&". ○○. ○○")'; $ws.Range("B29").HorizontalAlignment = -4108
    $ws.Range("B30:E30").Merge() | Out-Null; $ws.Range("B30").Value2 = "제안자 성명                         (서명 또는 날인)"; $ws.Range("B30").HorizontalAlignment = -4108
    $ws.Range("B31:E31").Merge() | Out-Null; $ws.Range("B31").Formula = '=IF(기초자료입력!$C$4="","○○○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'; $ws.Range("B31").HorizontalAlignment = -4108
    $ws.Range("B7:E7,B8:E17,B19:E27,B29:E31").Borders.LineStyle = 1; $ws.Range("B7:E31").VerticalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 19; $ws.Columns.Item("C").ColumnWidth = 17; $ws.Columns.Item("D").ColumnWidth = 14; $ws.Columns.Item("E").ColumnWidth = 18; $ws.Range("B3:E31").Font.Size = 10; $ws.Range("B5").Font.Size = 8
    $ws.Rows.Item(1).RowHeight = 12; $ws.Rows.Item(2).RowHeight = 8; $ws.Rows.Item(3).RowHeight = 25; $ws.Rows.Item(4).RowHeight = 25; $ws.Rows.Item(5).RowHeight = 28; $ws.Rows.Item(6).RowHeight = 6; $ws.Rows.Item(7).RowHeight = 18; for ($r = 8; $r -le 15; $r++) { $ws.Rows.Item($r).RowHeight = 18 }; $ws.Rows.Item(16).RowHeight = 24; $ws.Rows.Item(17).RowHeight = 24; $ws.Rows.Item(18).RowHeight = 6; $ws.Rows.Item(19).RowHeight = 18; for ($r = 20; $r -le 27; $r++) { $ws.Rows.Item($r).RowHeight = 16 }; $ws.Rows.Item(28).RowHeight = 6; $ws.Rows.Item(29).RowHeight = 18; $ws.Rows.Item(30).RowHeight = 18; $ws.Rows.Item(31).RowHeight = 18
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 0.7; $ps.BottomMargin = CmToPt 0.7; $ps.LeftMargin = CmToPt 0.9; $ps.RightMargin = CmToPt 0.9; $ps.HeaderMargin = CmToPt 0.3; $ps.FooterMargin = CmToPt 0.3; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$E`$31"
    L "F-021 원문 교복 납품 제안서 대조 레이아웃 및 개인정보 빈칸 경계 적용 완료"

    # ---- 7j. F-022 교복 납품 실적표 ----
    # 원문 HWPX [9-10] [서식 5]을 대조한다. 업체명(R-03)만 자동 반영하며, 사업명·기간·금액·발주처·비고는
    # 업체별 실적증명서 원본과 대조하여 이 출력 양식에 수기로 기재한다. 공통 DB로 전파하거나 자동 저장하지 않는다.
    $wsF22 = $wbNew.Worksheets.Add(); $wsF22.Name = "F-022_교복납품실적표"; $ws = $wsF22
    $ws.Range("B1:G1").Merge() | Out-Null; $ws.Range("B1").Value2 = "[검토용] 원본 HWPX [9-10] [서식 5] 대조. 업체명만 자동 반영하며 실적행은 업체별 증명서와 대조하여 직접 작성합니다."; $ws.Range("B1").Font.Size = 7; $ws.Range("B1").Font.Color = 255; $ws.Range("B1").WrapText = $false
    $ws.Range("I2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("I2").Font.Size = 7; $ws.Range("I3").Value2 = 1
    $ws.Range("B3:G4").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복 납품 실적표"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5:G5").Merge() | Out-Null; $ws.Range("B5").Formula = '="(제안자: "&IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$I$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$I$3)),"")&")"'; $ws.Range("B5").HorizontalAlignment = -4108
    $ws.Range("B7:G7").Merge() | Out-Null; $ws.Range("B7").Value2 = "교복 납품 실적(입찰공고일 기준 최근 3년 이내 실적)"; $ws.Range("B7").Font.Bold = $true; $ws.Range("B7").HorizontalAlignment = -4108
    $ws.Range("B9").Value2 = "순번"; $ws.Range("C9").Value2 = "사업명"; $ws.Range("D9").Value2 = "사업기간"; $ws.Range("E9").Value2 = "계약금액`n(천원)"; $ws.Range("F9").Value2 = "발주처"; $ws.Range("G9").Value2 = "비고"
    $ws.Range("B9:G9").Font.Bold = $true; $ws.Range("B9:G9").HorizontalAlignment = -4108; $ws.Range("B9:G9").VerticalAlignment = -4108; $ws.Range("B9:G9").WrapText = $true
    $ws.Range("B10").Value2 = 1; $ws.Range("B11").Value2 = 2; $ws.Range("B12").Value2 = 3; $ws.Range("B13").Value2 = 4; $ws.Range("B14").Value2 = 5; $ws.Range("B15").Value2 = 6; $ws.Range("B16").Value2 = 7; $ws.Range("B17").Value2 = 8; $ws.Range("B18").Value2 = 9
    $ws.Range("B10:B18").HorizontalAlignment = -4108
    $ws.Range("B20:G20").Merge() | Out-Null; $ws.Range("B20").Value2 = "붙임  실적증명서 각 1부."; $ws.Range("B21:G21").Merge() | Out-Null; $ws.Range("B21").Value2 = "위와 같이 납품실적을 제출합니다."; $ws.Range("B23:G23").Merge() | Out-Null; $ws.Range("B23").Formula = '=IF(기초자료입력!$C$5="","20○○.  .  .","20"&RIGHT(기초자료입력!$C$5,2)&".  .  .")'; $ws.Range("B23").HorizontalAlignment = -4108
    $ws.Range("B24:G24").Merge() | Out-Null; $ws.Range("B24").Value2 = "제안자 성 명                         (서명 또는 날인)"; $ws.Range("B24").HorizontalAlignment = -4108
    $ws.Range("B25:G25").Merge() | Out-Null; $ws.Range("B25").Formula = '=IF(기초자료입력!$C$4="","○○○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'; $ws.Range("B25").HorizontalAlignment = -4108
    $ws.Range("B3:G5,B7:G7,B9:G18,B20:G21,B23:G25").Borders.LineStyle = 1; $ws.Range("B3:G25").VerticalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 6; $ws.Columns.Item("C").ColumnWidth = 20; $ws.Columns.Item("D").ColumnWidth = 14; $ws.Columns.Item("E").ColumnWidth = 12; $ws.Columns.Item("F").ColumnWidth = 15; $ws.Columns.Item("G").ColumnWidth = 13; $ws.Range("B3:G25").Font.Size = 10
    $ws.Rows.Item(1).RowHeight = 12; $ws.Rows.Item(2).RowHeight = 8; $ws.Rows.Item(3).RowHeight = 24; $ws.Rows.Item(4).RowHeight = 24; $ws.Rows.Item(5).RowHeight = 20; $ws.Rows.Item(6).RowHeight = 6; $ws.Rows.Item(7).RowHeight = 20; $ws.Rows.Item(8).RowHeight = 6; $ws.Rows.Item(9).RowHeight = 30; for ($r = 10; $r -le 18; $r++) { $ws.Rows.Item($r).RowHeight = 24 }; $ws.Rows.Item(19).RowHeight = 6; $ws.Rows.Item(20).RowHeight = 18; $ws.Rows.Item(21).RowHeight = 18; $ws.Rows.Item(22).RowHeight = 8; $ws.Rows.Item(23).RowHeight = 18; $ws.Rows.Item(24).RowHeight = 18; $ws.Rows.Item(25).RowHeight = 18
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 0.8; $ps.BottomMargin = CmToPt 0.8; $ps.LeftMargin = CmToPt 0.6; $ps.RightMargin = CmToPt 0.6; $ps.HeaderMargin = CmToPt 0.3; $ps.FooterMargin = CmToPt 0.3; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$G`$25"
    L "F-022 원문 교복 납품 실적표 대조 레이아웃(실적행 9행) 및 업체별 수기 실적행 경계 적용 완료"

    # ---- 7k. F-023 교복 제조 사양서 ----
    # 원문 HWPX [9-11] [서식 6]을 대조한다. 업체명(R-03)과 학교 공통값만 반영하고 재질·설명 내용은
    # 업체가 직접 작성하는 빈 칸으로 유지한다. 서식_매핑표.md의 F-023 행(R-04~R-06/K-01)은 원문에 없는
    # 품목·수량·단가·금액 필드를 가리켜 F-013 행과 동일한 문구로 남아 있는 것으로 보이며, 실제 원문은
    # 가격표가 아니라 F-017~F-022와 같은 업체명 전용 서식이므로 그 패턴을 따른다.
    $wsF23 = $wbNew.Worksheets.Add(); $wsF23.Name = "F-023_교복제조사양서"; $ws = $wsF23
    $ws.Range("B1:E1").Merge() | Out-Null; $ws.Range("B1").Value2 = "[검토용] 원본 HWPX [9-11] [서식 6] 대조. 업체명만 자동 반영하며 재질·설명 내용은 직접 작성합니다."; $ws.Range("B1").Font.Size = 7; $ws.Range("B1").Font.Color = 255; $ws.Range("B1").WrapText = $false
    $ws.Range("G2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("G2").Font.Size = 7; $ws.Range("G3").Value2 = 1
    $ws.Range("B3:E4").Merge() | Out-Null; $ws.Range("B3").Value2 = "납품 교복 제조 사양서"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5:E5").Merge() | Out-Null; $ws.Range("B5").Formula = '="(제안자: "&IFERROR(IF(INDEX(기초자료입력!$B$44:$B$53,$G$3)=0,"",INDEX(기초자료입력!$B$44:$B$53,$G$3)),"")&")"'; $ws.Range("B5").HorizontalAlignment = -4108
    $ws.Range("B6:E6").Merge() | Out-Null; $ws.Range("B6").Value2 = "*서식은 입찰 업체 상황에 맞게 변경 사용 가능(사진 등 첨부 가능)"; $ws.Range("B6").Font.Size = 8; $ws.Range("B6").HorizontalAlignment = -4108
    $ws.Range("B7:E7").Merge() | Out-Null; $ws.Range("B7").Value2 = "생활형 교복"; $ws.Range("B7").Font.Bold = $true
    $ws.Range("B9").Value2 = "구분"; $ws.Range("C9").Value2 = "견본 품목"; $ws.Range("D9").Value2 = "재질"; $ws.Range("E9").Value2 = "설명 내용"
    $ws.Range("B9:E9").Font.Bold = $true; $ws.Range("B9:E9").HorizontalAlignment = -4108; $ws.Range("B9:E9").VerticalAlignment = -4108; $ws.Range("B9:E9").WrapText = $true
    $ws.Range("B10:B13").Merge() | Out-Null; $ws.Range("B10").Value2 = "동복"
    $ws.Range("C10").Value2 = "후드 점퍼"; $ws.Range("C11").Value2 = "집업티"; $ws.Range("C12").Value2 = "맨투맨티"; $ws.Range("C13").Value2 = "긴바지"
    $ws.Range("B14:B15").Merge() | Out-Null; $ws.Range("B14").Value2 = "하복"
    $ws.Range("C14").Value2 = "반팔티"; $ws.Range("C15").Value2 = "반바지"
    $ws.Range("B10:E15").HorizontalAlignment = -4108; $ws.Range("B10:B15,C10:C15").HorizontalAlignment = -4108
    $ws.Range("B17:E17").Merge() | Out-Null; $ws.Range("B17").Value2 = "※ 국산섬유원단 사용 여부(한국섬유산업연합회 인증마크 획득) 등 상세히 기재"; $ws.Range("B17").Font.Size = 8; $ws.Range("B17").WrapText = $true
    $ws.Range("B19:E19").Merge() | Out-Null; $ws.Range("B19").Formula = '=IF(기초자료입력!$C$5="","20○○.  .  .","20"&RIGHT(기초자료입력!$C$5,2)&".  .  .")'; $ws.Range("B19").HorizontalAlignment = -4108
    $ws.Range("B20:E20").Merge() | Out-Null; $ws.Range("B20").Value2 = "제안자 성 명                         (서명 또는 날인)"; $ws.Range("B20").HorizontalAlignment = -4108
    $ws.Range("B21:E21").Merge() | Out-Null; $ws.Range("B21").Formula = '=IF(기초자료입력!$C$4="","○○○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'; $ws.Range("B21").HorizontalAlignment = -4108
    $ws.Range("B3:E6,B7:E7,B9:E15,B17:E17,B19:E21").Borders.LineStyle = 1; $ws.Range("B3:E21").VerticalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 10; $ws.Columns.Item("C").ColumnWidth = 16; $ws.Columns.Item("D").ColumnWidth = 16; $ws.Columns.Item("E").ColumnWidth = 24; $ws.Range("B3:E21").Font.Size = 10
    $ws.Rows.Item(1).RowHeight = 12; $ws.Rows.Item(2).RowHeight = 8; $ws.Rows.Item(3).RowHeight = 24; $ws.Rows.Item(4).RowHeight = 24; $ws.Rows.Item(5).RowHeight = 20; $ws.Rows.Item(6).RowHeight = 16; $ws.Rows.Item(7).RowHeight = 18; $ws.Rows.Item(8).RowHeight = 6; $ws.Rows.Item(9).RowHeight = 22; for ($r = 10; $r -le 15; $r++) { $ws.Rows.Item($r).RowHeight = 22 }; $ws.Rows.Item(16).RowHeight = 6; $ws.Rows.Item(17).RowHeight = 20; $ws.Rows.Item(18).RowHeight = 8; $ws.Rows.Item(19).RowHeight = 18; $ws.Rows.Item(20).RowHeight = 18; $ws.Rows.Item(21).RowHeight = 18
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 0.8; $ps.BottomMargin = CmToPt 0.8; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.3; $ps.FooterMargin = CmToPt 0.3; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$E`$21"
    L "F-023 원문 교복 제조 사양서 대조 레이아웃(동복 4행·하복 2행) 및 업체명만 반영·재질/설명 수기 경계 적용 완료"

    # ---- 7l. F-025 교복 A/S 계획서 ----
    # 원문 HWPX [9-14] [서식 8]을 대조한다. 표가 아닌 순수 텍스트 서식이며 업체명·회사명·상호를 표시하는
    # 자리가 원문에 없다(F-021의 "회사명", F-022/F-023의 "(제안자: 업체명)"과 다름). 서명란은
    # "제안자 성 명(서명 또는 날인)"과 "◯◯◯◯학교장 귀하"뿐이므로 학교 공통값(C-01·C-02)만 반영하고
    # 나머지 전 항목(A/S 기간·편의성·지정업체·내용)은 F-017~F-023과 같이 업체가 직접 작성하는 빈 칸으로 유지한다.
    # 서식_매핑표.md의 F-025 행은 F-026 행과 입력·출력 열이 완전히 동일해(비고만 다름) 복사 흔적으로 보이나,
    # 출력 열이 R-03을 제외하고 C-01/C-02/B-01/D-02만 적은 것은 원문에 업체명 표시 자리가 없다는 사실과 결과적으로
    # 부합한다. 업체별 제출 서류(F-021 첨부 목록 ④)이므로 F-017~F-023과 동일하게 업체 반복행 기준으로
    # 업체별 임시 인쇄 시트를 생성하되, 업체명 텍스트 자체는 출력하지 않는다.
    $wsF25 = $wbNew.Worksheets.Add(); $wsF25.Name = "F-025_교복AS계획서"; $ws = $wsF25
    $ws.Range("B1:E1").Merge() | Out-Null; $ws.Range("B1").Value2 = "[검토용] 원본 HWPX [9-14] [서식 8] 대조. 원문에 업체명 표시 자리가 없어 반영하지 않으며 A/S 계획 전 항목은 직접 작성합니다."; $ws.Range("B1").Font.Size = 7; $ws.Range("B1").Font.Color = 255; $ws.Range("B1").WrapText = $false
    $ws.Range("G2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("G2").Font.Size = 7; $ws.Range("G3").Value2 = 1
    $ws.Range("B3:E4").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복 A/S 계획서"; $ws.Range("B3").Font.Size = 16; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B6:E6").Merge() | Out-Null; $ws.Range("B6").Value2 = "1. A/S 기간"; $ws.Range("B6").Font.Bold = $true
    $ws.Range("B7:E7").Merge() | Out-Null; $ws.Range("B7").Value2 = "가. 무상 A/S 기간 (                부터                까지 )"; $ws.Range("B7").WrapText = $true
    $ws.Range("B9:E9").Merge() | Out-Null; $ws.Range("B9").Value2 = "2. A/S 편의성"; $ws.Range("B9").Font.Bold = $true
    $ws.Range("B10:E10").Merge() | Out-Null; $ws.Range("B10").Value2 = "가. 사업장 소재지 :"; $ws.Range("B10").WrapText = $true
    $ws.Range("B11:E11").Merge() | Out-Null; $ws.Range("B11").Value2 = "나. 신속성 : A/S 완료까지 소요되는 기간 (                일)"; $ws.Range("B11").WrapText = $true
    $ws.Range("B12:E12").Merge() | Out-Null; $ws.Range("B12").Value2 = "다. 출장 A/S 가능여부 : 가능한 경우 출장 A/S 최소 학생 수 기재 (                명)"; $ws.Range("B12").WrapText = $true
    $ws.Range("B13:E13").Merge() | Out-Null; $ws.Range("B13").Value2 = "라. 택배 A/S 가능여부 : 가능한 경우 택배비용 무상/유상 기재 (                )"; $ws.Range("B13").WrapText = $true
    $ws.Range("B15:E15").Merge() | Out-Null; $ws.Range("B15").Value2 = "3. A/S지정업체 (A/S지정업체가 별도 있을 경우 교복업체와 관계를 입증할 계약서 등 첨부)"; $ws.Range("B15").Font.Bold = $true; $ws.Range("B15").WrapText = $true
    $ws.Range("B16:E16").Merge() | Out-Null; $ws.Range("B16").Value2 = "가. 주 소 :"; $ws.Range("B16").WrapText = $true
    $ws.Range("B17:E17").Merge() | Out-Null; $ws.Range("B17").Value2 = "나. 연락처 :"; $ws.Range("B17").WrapText = $true
    $ws.Range("B18:E18").Merge() | Out-Null; $ws.Range("B18").Value2 = "다. 거리접근성 : 학교와 A/S 수선점과의 거리 기재 (                km)"; $ws.Range("B18").WrapText = $true
    $ws.Range("B20:E20").Merge() | Out-Null; $ws.Range("B20").Value2 = "4. A/S 내용"; $ws.Range("B20").Font.Bold = $true
    $ws.Range("B21:E21").Merge() | Out-Null; $ws.Range("B21").Value2 = "가. 시설현황 : A/S관련 시설 및 수량 기재(재단기, 오버로크 등)"; $ws.Range("B21").WrapText = $true
    $ws.Range("B22:E22").Merge() | Out-Null; $ws.Range("B22").Value2 = "나. 무상A/S 세부내용"; $ws.Range("B22").WrapText = $true
    $ws.Range("B23:E24").Merge() | Out-Null; $ws.Range("B23").Value2 = "- 바지/치마 기장수선  - 품·허리수선  - 헤진부위수선  - 단추  - 지퍼  - 바느질 등 세부적 사항 기재"; $ws.Range("B23").WrapText = $true; $ws.Range("B23").VerticalAlignment = -4108
    $ws.Range("B25:E25").Merge() | Out-Null; $ws.Range("B25").Value2 = "다. 유상A/S 세부내용"; $ws.Range("B25").WrapText = $true
    $ws.Range("B26:E26").Merge() | Out-Null; $ws.Range("B26").Value2 = "-"
    $ws.Range("B27:E27").Merge() | Out-Null; $ws.Range("B27").Value2 = "-"
    $ws.Range("B29:E29").Merge() | Out-Null; $ws.Range("B29").Formula = '=IF(기초자료입력!$C$5="","20○○년   월   일","20"&RIGHT(기초자료입력!$C$5,2)&"년   월   일")'; $ws.Range("B29").HorizontalAlignment = -4108
    $ws.Range("B30:E30").Merge() | Out-Null; $ws.Range("B30").Value2 = "제안자 성 명                         (서명 또는 날인)"; $ws.Range("B30").HorizontalAlignment = -4108
    $ws.Range("B31:E31").Merge() | Out-Null; $ws.Range("B31").Formula = '=IF(기초자료입력!$C$4="","○○○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'; $ws.Range("B31").HorizontalAlignment = -4108
    $ws.Range("B3:E4,B6:E7,B9:E13,B15:E18,B20:E27,B29:E31").Borders.LineStyle = 1; $ws.Range("B3:E31").VerticalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 20; $ws.Columns.Item("C").ColumnWidth = 20; $ws.Columns.Item("D").ColumnWidth = 20; $ws.Columns.Item("E").ColumnWidth = 20; $ws.Range("B3:E31").Font.Size = 10
    $ws.Rows.Item(1).RowHeight = 12; $ws.Rows.Item(2).RowHeight = 8; $ws.Rows.Item(3).RowHeight = 26; $ws.Rows.Item(4).RowHeight = 26; $ws.Rows.Item(5).RowHeight = 6; $ws.Rows.Item(6).RowHeight = 18; $ws.Rows.Item(7).RowHeight = 20; $ws.Rows.Item(8).RowHeight = 6; $ws.Rows.Item(9).RowHeight = 18; for ($r = 10; $r -le 13; $r++) { $ws.Rows.Item($r).RowHeight = 20 }; $ws.Rows.Item(14).RowHeight = 6; $ws.Rows.Item(15).RowHeight = 26; $ws.Rows.Item(16).RowHeight = 18; $ws.Rows.Item(17).RowHeight = 18; $ws.Rows.Item(18).RowHeight = 20; $ws.Rows.Item(19).RowHeight = 6; $ws.Rows.Item(20).RowHeight = 18; $ws.Rows.Item(21).RowHeight = 20; $ws.Rows.Item(22).RowHeight = 18; $ws.Rows.Item(23).RowHeight = 22; $ws.Rows.Item(24).RowHeight = 22; $ws.Rows.Item(25).RowHeight = 18; $ws.Rows.Item(26).RowHeight = 18; $ws.Rows.Item(27).RowHeight = 18; $ws.Rows.Item(28).RowHeight = 8; $ws.Rows.Item(29).RowHeight = 18; $ws.Rows.Item(30).RowHeight = 18; $ws.Rows.Item(31).RowHeight = 18
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 0.8; $ps.BottomMargin = CmToPt 0.8; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.3; $ps.FooterMargin = CmToPt 0.3; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$B`$1:`$E`$31"
    L "F-025 원문 교복 A/S 계획서 대조 레이아웃(업체명 표시 없음, 학교 공통값만 반영) 및 전 항목 수기 경계 적용 완료"

    # ---- 7m. F-001 교복선정위원회 구성 기안 ----
    # 원문 HWPX 3쪽 [1] 교복선정위원회 구성 기안(예시)을 대조한다. F-007과 동일한 기안문 틀(수신·경유·제목·
    # 결재란)에 "라. 위원 명단" 표(순/구분/직급(직위)/성명/비고 7행 예시)가 추가된 구조다. 위원 성명은
    # 개인정보이므로 자동 반영하지 않고, 기초자료입력!B58:C67(기존 "6. 위원 반복행" — 역할직위/마스킹식별표시,
    # F-014~F-016 설계 시 R-01/R-02 공용 저장소로 이미 마련돼 있었으나 지금까지 어떤 서식도 소비하지 않던 것)을
    # 그대로 재사용한다. "구분"(위원장/위원)은 원문처럼 첫 행만 위원장으로 고정 표시하고 나머지는 위원으로
    # 파생하며, 실제 성명 대신 마스킹 식별표시만 노출한다. 기초자료입력·기존 위원행검증 VBA는 변경하지 않는다.
    $wsF01 = $wbNew.Worksheets.Add(); $wsF01.Name = "F-001_위원회구성기안"; $ws = $wsF01
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 3쪽 [1] 대조. 위원 성명은 개인정보이므로 마스킹 식별표시만 반영함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H3").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복선정위원회 구성(안)"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5&"학년도 교복선정위원회 구성(안)","교복선정위원회 구성(안)")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("C11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 관련："; $ws.Range("C11").Formula = '=IF(기초자료입력!C23<>"",기초자료입력!C23,"")'
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" "&IF(기초자료입력!C5<>"",기초자료입력!C5,"○○")&"학년도 교복선정을 위한 교복선정위원회를 아래와 같이 구성하고자 합니다."'; $ws.Range("B13").WrapText = $true
    $ws.Range("B15:H15").Merge() | Out-Null; $ws.Range("B15").Value2 = "가. 위원회명：교복선정위원회"
    $ws.Range("B16:H16").Merge() | Out-Null; $ws.Range("B16").Value2 = "나. 구성인원：7인 이상(교원위원 3인, 학부모위원 2인, 학생위원 2인)"; $ws.Range("B16").WrapText = $true
    $ws.Range("B17:H17").Merge() | Out-Null; $ws.Range("B17").Value2 = "※ 교장 및 행정실장은 위원에서 제외, 교원을 제외한 위원이 1/2 이상 포함되도록 구성"; $ws.Range("B17").Font.Size = 8; $ws.Range("B17").WrapText = $true
    $ws.Range("B18:H18").Merge() | Out-Null; $ws.Range("B18").Value2 = "다. 역 할"
    $rolesF01 = @("1) 교복 디자인 및 사양 선정", "2) 교복 품질심사기준 설정 및 심사", "3) 교복 제작 감독", "4) 구매결과 평가 실시")
    for ($i = 0; $i -lt $rolesF01.Count; $i++) { $r = 19 + $i; $ws.Range("B${r}:H${r}").Merge() | Out-Null; $ws.Range("B$r").Value2 = $rolesF01[$i] }
    $ws.Range("B24:H24").Merge() | Out-Null; $ws.Range("B24").Value2 = "라. 위원 명단"
    $headersF01 = @("순","구분","직급(직위)","성명","비고")
    $colsF01 = @("B","C","D","E","F")
    for ($i = 0; $i -lt $headersF01.Count; $i++) { $ws.Range("$($colsF01[$i])25").Value2 = $headersF01[$i] }
    $ws.Range("F25:H25").Merge() | Out-Null
    for ($i = 0; $i -lt 10; $i++) {
        $r = 26 + $i; $src = 58 + $i
        $roleLabel = if ($i -eq 0) { "위원장" } else { "위원" }
        $ws.Range("B$r").Formula = "=IF(기초자료입력!B${src}<>`"`",$($i+1),`"`")"
        $ws.Range("C$r").Formula = "=IF(기초자료입력!B${src}<>`"`",`"$roleLabel`",`"`")"
        $ws.Range("D$r").Formula = "=IF(기초자료입력!B${src}<>`"`",기초자료입력!B${src},`"`")"
        $ws.Range("E$r").Formula = "=IF(기초자료입력!C${src}<>`"`",기초자료입력!C${src},`"`")"
        $ws.Range("F${r}:H${r}").Merge() | Out-Null
    }
    $ws.Range("B37:H37").Merge() | Out-Null; $ws.Range("B37").Value2 = "붙임 1. 교복선정위원회 위원 수락 및 확인서 1부."
    $ws.Range("B38:H38").Merge() | Out-Null; $ws.Range("B38").Value2 = "      2. 교복선정위원회 위원 청렴 및 보안 서약서 1부.  끝."
    $ws.Range("F40").Value2 = "담당"; $ws.Range("G40").Value2 = "협조자"; $ws.Range("H40").Value2 = "교장"
    $ws.Range("B41").Value2 = "시행"; $ws.Range("C41:E41").Merge() | Out-Null; $ws.Range("C41").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F41").Value2 = "접수"; $ws.Range("G41:H41").Merge() | Out-Null
    $ws.Range("B42").Value2 = "우편번호"; $ws.Range("D42").Value2 = "주소"; $ws.Range("E42:H42").Merge() | Out-Null
    $ws.Range("B43").Value2 = "전화"; $ws.Range("D43").Value2 = "전송(팩스)"; $ws.Range("F43").Value2 = "이메일"; $ws.Range("G43:H43").Merge() | Out-Null
    $ws.Range("B5:H43").Font.Size = 9
    $ws.Range("B5,E5,B7,B8,B9,B11,B15,B16,B18,B24,B25:F25,B37,F40:H40,B41,F41,B42,D42,B43,D43,F43").Font.Bold = $true
    $ws.Range("B25:F25").HorizontalAlignment = -4108
    $ws.Range("B5:H25,B26:H35,F40:H43").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 8; $ws.Columns.Item("C").ColumnWidth = 8; $ws.Columns.Item("D").ColumnWidth = 12; $ws.Columns.Item("E").ColumnWidth = 10; for ($c = 6; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 8 }
    for ($r = 5; $r -le 43; $r++) { $ws.Rows.Item($r).RowHeight = 14.5 }
    $ws.Rows.Item(3).RowHeight = 24; $ws.Rows.Item(16).RowHeight = 20; $ws.Rows.Item(17).RowHeight = 12; $ws.Rows.Item(13).RowHeight = 26
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 0.8; $ps.BottomMargin = CmToPt 0.8; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.3; $ps.FooterMargin = CmToPt 0.3; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$43"
    $hbF01 = $ws.HPageBreaks.Count; $vbF01 = $ws.VPageBreaks.Count
    L "F-001 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF01, VPageBreaks=$vbF01 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-001 원문 교복선정위원회 구성 기안 대조 레이아웃(위원 성명 마스킹, DB_위원 재사용) 적용 완료"

    # ---- 7n. F-002 교복선정위원회 위원 수락 및 확인서 ----
    # 원문 HWPX 4쪽 [1-1]을 대조한다. 표 없는 순수 서약 문서로, 위원 개인이 자필로 작성·서명하는
    # 빈 양식이다(F-025와 동일하게 업체명/위원명 표시 자리가 없고, 주소·성명·서명은 전부 공란).
    # 학교명(수신처)·학년도(날짜 연도부)만 공통값으로 반영하고 나머지 전 항목은 수기 작성 공란으로 유지한다.
    $wsF02 = $wbNew.Worksheets.Add(); $wsF02.Name = "F-002_위원수락확인서"; $ws = $wsF02
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 4쪽 [1-1] 대조. 위원이 자필로 작성·서명하는 빈 양식이며 학교 공통정보만 반영함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:E4").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복선정위원회 위원 수락 및 확인서"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    # 본문 3단락은 여러 행에 걸친 병합 셀 하나로 쓰지 않는다 — WrapText는 병합 셀의 행 높이를 자동으로
    # 늘려주지 않아 인쇄 시 글자가 잘리는 결함이 1차 PDF 육안 확인에서 실제로 발견됨. 단락마다 별도의
    # 단일 행(가로 병합만 사용)으로 나누고 AutoFit으로 각 행 높이를 내용에 맞게 자동 계산한다.
    $ws.Range("B6:E6").Merge() | Out-Null
    $ws.Range("B6").Formula = '="본인은 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복선정위원회 활동을 공정하게 수행할 것이며, 교복가격 안정화에 적극 노력하겠습니다."'
    $ws.Range("B7:E7").Merge() | Out-Null
    $ws.Range("B7").Formula = '="아울러 본인은 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 입찰에 참가한(또는 앞으로 참가할) 업체와 전혀 관련이 없음을 확인합니다."'
    $ws.Range("B8:E8").Merge() | Out-Null
    $ws.Range("B8").Value2 = "업체와의 관련이 있는 것으로 판명될 경우, 즉시 위원직을 사퇴할 것이며, 추후 어떠한 이의도 제기하지 않을 것을 서약합니다."
    # 주의: 서로 다른 단일행 병합 3개(B6:E6/B7:E7/B8:E8)에 걸친 "B6:B8" 같은 범위에 서식을 한 번에 적용하면
    # Excel이 이를 하나의 B6:E8 병합으로 합쳐버려 B7·B8 내용이 사라지는 결함이 실제로 발견됨(COM 자동화에서
    # 병합 경계를 넘나드는 범위에 속성을 설정할 때의 부작용). 행마다 개별적으로 서식을 적용해 회피한다.
    foreach ($r in 6,7,8) { $ws.Range("B$r").WrapText = $true; $ws.Range("B$r").VerticalAlignment = -4160 }
    $ws.Range("B12:E12").Merge() | Out-Null; $ws.Range("B12").Formula = '=IF(기초자료입력!$C$5="","20○○.  ○○  .  ○○  .","20"&RIGHT(기초자료입력!$C$5,2)&".  ○○  .  ○○  .")'; $ws.Range("B12").HorizontalAlignment = -4108
    $ws.Range("B14:E14").Merge() | Out-Null; $ws.Range("B14").Value2 = "주      소 :"
    $ws.Range("B16:E16").Merge() | Out-Null; $ws.Range("B16").Value2 = "성      명 :"
    $ws.Range("B18:E18").Merge() | Out-Null; $ws.Range("B18").Value2 = "서      명 :                              (인)"
    $ws.Range("B21:E21").Merge() | Out-Null; $ws.Range("B21").Formula = '=IF(기초자료입력!$C$4="","○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'; $ws.Range("B21").HorizontalAlignment = -4108
    $ws.Range("B3:E4,B6:E8,B12:E12,B14:E18,B21:E21").Borders.LineStyle = 1; $ws.Range("B3:E21").VerticalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 20; $ws.Columns.Item("C").ColumnWidth = 20; $ws.Columns.Item("D").ColumnWidth = 20; $ws.Columns.Item("E").ColumnWidth = 20; $ws.Range("B3:E21").Font.Size = 11
    $ws.Rows.Item(1).RowHeight = 12; $ws.Rows.Item(3).RowHeight = 26; $ws.Rows.Item(4).RowHeight = 26; $ws.Rows.Item(5).RowHeight = 10
    $ws.Rows.Item("6:8").AutoFit()
    for ($r = 6; $r -le 8; $r++) { if ($ws.Rows.Item($r).RowHeight -lt 28) { $ws.Rows.Item($r).RowHeight = 28 } }
    $ws.Rows.Item(9).RowHeight = 14; $ws.Rows.Item(10).RowHeight = 14; $ws.Rows.Item(11).RowHeight = 10; $ws.Rows.Item(12).RowHeight = 20; $ws.Rows.Item(13).RowHeight = 14; $ws.Rows.Item(14).RowHeight = 22; $ws.Rows.Item(15).RowHeight = 14; $ws.Rows.Item(16).RowHeight = 22; $ws.Rows.Item(17).RowHeight = 14; $ws.Rows.Item(18).RowHeight = 22; $ws.Rows.Item(19).RowHeight = 14; $ws.Rows.Item(20).RowHeight = 10; $ws.Rows.Item(21).RowHeight = 20
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$E`$21"
    $hbF02 = $ws.HPageBreaks.Count; $vbF02 = $ws.VPageBreaks.Count
    L "F-002 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF02, VPageBreaks=$vbF02 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-002 원문 위원 수락 및 확인서 대조 레이아웃(주소·성명·서명 전 항목 수기 경계, 학교 공통정보만 반영) 적용 완료"

    # ---- 7o. F-003 교복선정위원회 위원 청렴 및 보안 서약서 ----
    # 원문 HWPX 5쪽 [1-2]를 대조한다. F-002와 동일한 성격의 서약 문서(위원이 자필 작성·서명)이며, 4개 항목의
    # 준수사항이 나열된다. 성명·서명·날짜는 전부 공란으로 유지한다. F-002에서 발견한 COM 함정(병합 폭과
    # 일치하지 않는 부분 범위에 서식을 일괄 적용하면 Excel이 인접 병합을 자동 통합함)을 피하기 위해
    # 단락·항목마다 단일행 병합을 만들고 서식은 행별로 개별 적용한다.
    $wsF03 = $wbNew.Worksheets.Add(); $wsF03.Name = "F-003_위원청렴보안서약서"; $ws = $wsF03
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 5쪽 [1-2] 대조. 위원이 자필로 작성·서명하는 빈 양식이며 학교 공통정보만 반영함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:E4").Merge() | Out-Null; $ws.Range("B3").Value2 = "청렴 및 보안 서약서"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B6:E6").Merge() | Out-Null
    $ws.Range("B6").Formula = '="본인은      년      월      일 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&"에서 실시하는 교복 학교주관구매업체 선정을 위한 평가함에 있어 「부패 없는 투명한 사회」 구현 등을 위하여 다음 사항을 준수할 것을 서약합니다."'
    # 주의: @() 배열 리터럴의 첫 원소에 "a" + "b" 형태의 문자열 연결을 쓰면 [scriptblock]::Create()
    # 경로(이 빌더 전체가 이 방식으로 실행됨)에서 배열이 원소 1개로 붕괴하는 파서 결함이 실제로
    # 재현됨(격리된 최소 재현 스크립트로 확인). 원소는 전부 단일 리터럴 문자열로만 작성한다.
    $itemsF03 = @(
        "1. 20○○학년도 ○○학교 교복 학교주관구매업체 선정을 위해 제시된 항목에 따라 객관적이고 공정하게 심사할 것을 약속합니다.",
        "2. ○○학교 교복 학교 주관 교복선정 위원회 지위를 이용하여 관련업체로부터 어떠한 경우에도 금품·향응·편의 등을 수수하거나 제공받지 않을 것이며, 이러한 상황이 발생하면 사업부서에 통보하여 공정한 평가가 이루어지도록 하겠습니다.",
        "3. 업무상 취득한 비밀을 준수하고 보안관계 규정 및 지침을 성실히 수행하겠습니다.",
        "4. 평가와 관련하여 알게 된 업무상 비밀을 타인에게 누설하지 않겠으며, 업무상 취득한 비밀을 누설할 때에는 관계법규에 따라 처벌을 받는 것에 이의를 제기하지 않겠습니다."
    )
    for ($i = 0; $i -lt $itemsF03.Count; $i++) {
        $r = 8 + $i
        $ws.Range("B${r}:E${r}").Merge() | Out-Null
        if ($i -eq 0) {
            $ws.Cells.Item($r, 2).Formula = '="1. 20○○학년도 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 학교주관구매업체 선정을 위해 제시된 항목에 따라 객관적이고 공정하게 심사할 것을 약속합니다."'
        } else {
            $escaped = $itemsF03[$i].Replace('"', '""')
            $ws.Cells.Item($r, 2).Formula = '="' + $escaped + '"'
        }
    }
    foreach ($r in @(6,8,9,10,11)) { $ws.Range("B$r").WrapText = $true; $ws.Range("B$r").VerticalAlignment = -4160 }
    $ws.Range("B14:E14").Merge() | Out-Null; $ws.Cells.Item(14, 2).Formula = '="20      년      월      일"'; $ws.Range("B14").HorizontalAlignment = -4108
    $ws.Range("B17:E17").Merge() | Out-Null; $ws.Cells.Item(17, 2).Formula = '="서 약 자 성 명 :                              (인)"'
    $ws.Range("B20:E20").Merge() | Out-Null; $ws.Range("B20").Formula = '=IF(기초자료입력!$C$4="","○○○학교장 귀하",기초자료입력!$C$4&"장 귀하")'; $ws.Range("B20").HorizontalAlignment = -4108
    # 최종 확인: 모든 핵심 셀이 실제로 값을 가지는지 한 번 더 점검해 조용한 빈 칸 출력을 방지한다.
    foreach ($r in @(3,6,8,9,10,11,14,17,20)) {
        if ([string]::IsNullOrEmpty([string]$ws.Cells.Item($r, 2).Value2)) { throw "F-003 B$r 값이 비어 있습니다." }
    }
    foreach ($rng in @("B3:E4","B6:E6","B8:E8","B9:E9","B10:E10","B11:E11","B14:E14","B17:E17","B20:E20")) { $ws.Range($rng).Borders.LineStyle = 1 }
    $ws.Range("B3:E20").VerticalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 20; $ws.Columns.Item("C").ColumnWidth = 20; $ws.Columns.Item("D").ColumnWidth = 20; $ws.Columns.Item("E").ColumnWidth = 20; $ws.Range("B3:E20").Font.Size = 11
    $ws.Rows.Item(1).RowHeight = 12; $ws.Rows.Item(3).RowHeight = 26; $ws.Rows.Item(4).RowHeight = 26; $ws.Rows.Item(5).RowHeight = 10
    $ws.Rows.Item("6:6").AutoFit(); if ($ws.Rows.Item(6).RowHeight -lt 28) { $ws.Rows.Item(6).RowHeight = 28 }
    $ws.Rows.Item(7).RowHeight = 10
    $ws.Rows.Item("8:11").AutoFit()
    for ($r = 8; $r -le 11; $r++) { if ($ws.Rows.Item($r).RowHeight -lt 26) { $ws.Rows.Item($r).RowHeight = 26 } }
    $ws.Rows.Item(12).RowHeight = 10; $ws.Rows.Item(13).RowHeight = 14; $ws.Rows.Item(14).RowHeight = 20; $ws.Rows.Item(15).RowHeight = 14; $ws.Rows.Item(16).RowHeight = 14; $ws.Rows.Item(17).RowHeight = 22; $ws.Rows.Item(18).RowHeight = 14; $ws.Rows.Item(19).RowHeight = 10; $ws.Rows.Item(20).RowHeight = 20
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$E`$20"
    $hbF03 = $ws.HPageBreaks.Count; $vbF03 = $ws.VPageBreaks.Count
    L "F-003 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF03, VPageBreaks=$vbF03 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-003 원문 위원 청렴 및 보안 서약서 대조 레이아웃(성명·서명·날짜 전 항목 수기 경계, 학교 공통정보만 반영) 적용 완료"

    # ---- 7p. F-004 교복 학교주관구매 추진 계획 수립 기안문 ----
    # 원문 HWPX 6쪽 [2]를 대조한다. F-001과 동일한 기안문 틀(수신·경유·제목·결재란)이나 위원 명단 표가
    # 없는 단순 2단락 기안문이다. 붙임은 F-005(구매 추진 계획안)를 가리키는 고정 1건.
    $wsF04 = $wbNew.Worksheets.Add(); $wsF04.Name = "F-004_구매추진계획수립기안문"; $ws = $wsF04
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 6쪽 [2] 대조"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H3").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복 학교주관구매 추진 계획 수립(안)"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5&"학년도 교복 학교주관구매 추진 계획(안)","교복 학교주관구매 추진 계획(안)")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("C11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 관련："; $ws.Range("C11").Formula = '=IF(기초자료입력!C23<>"",기초자료입력!C23,"")'
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. 공정하고 투명한 교복 선정 및 합리적인 교복 구매를 위해 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" "&IF(기초자료입력!C5<>"",기초자료입력!C5,"○○")&"학년도 교복 학교주관구매 계획을 붙임과 같이 수립하여 업무 추진에 만전을 기하고자 합니다."'; $ws.Range("B13").WrapText = $true
    $ws.Range("B15:H15").Merge() | Out-Null; $ws.Range("B15").Formula = '="붙임  "&IF(기초자료입력!C5<>"",기초자료입력!C5,"○○")&"학년도 교복 학교주관구매 추진 계획(안) 1부.  끝."'
    $ws.Range("F17").Value2 = "담당"; $ws.Range("G17").Value2 = "협조자"; $ws.Range("H17").Value2 = "교장"
    $ws.Range("B18").Value2 = "시행"; $ws.Range("C18:E18").Merge() | Out-Null; $ws.Range("C18").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F18").Value2 = "접수"; $ws.Range("G18:H18").Merge() | Out-Null
    $ws.Range("B19").Value2 = "우편번호"; $ws.Range("D19").Value2 = "주소"; $ws.Range("E19:H19").Merge() | Out-Null
    $ws.Range("B20").Value2 = "전화"; $ws.Range("D20").Value2 = "전송(팩스)"; $ws.Range("F20").Value2 = "이메일"; $ws.Range("G20:H20").Merge() | Out-Null
    $ws.Range("B5:H20").Font.Size = 10
    $ws.Range("B5,E5,B7,B8,B9,B11,B15,B18,F18,B19,D19,B20,D20,F20").Font.Bold = $true
    $ws.Range("B5:H9,B17:H20").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 11; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    # F-002/F-004의 1차 PDF 육안 확인에서 실제로 발견된 결함(긴 줄바꿈 문단이 기본 행 높이에서 잘림)을
    # 재발시키지 않도록 문단 행을 명시적으로 AutoFit한다.
    $ws.Rows.Item("13:13").AutoFit(); if ($ws.Rows.Item(13).RowHeight -lt 28) { $ws.Rows.Item(13).RowHeight = 28 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$20"
    $hbF04 = $ws.HPageBreaks.Count; $vbF04 = $ws.VPageBreaks.Count
    L "F-004 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF04, VPageBreaks=$vbF04 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-004 원문 구매 추진 계획 수립 기안문 대조 레이아웃 적용 완료"

    # ---- 7q. F-005 교복 학교주관구매 추진 계획안 ----
    # 원문 HWPX 7~12쪽 [2-1]을 대조한다. P1-01 구조 집계(`_workspace/01_analysis/HWPX_원시구조_목록.md`
    # 2절, "6~12쪽 [2] 구매 추진계획 및 규격서 관련 본문")에 따르면 F-004([2], 6쪽) 뒤에 이어지는 F-005는
    # 원문에서만 6개 물리 페이지(7~12쪽)에 걸친 장문 정책 설명 문서다. F-001~F-004(단일 물리 페이지)와
    # 달리 F-005는 "A4 1쪽 자연 배율" 원칙을 그대로 적용할 수 없다 — 6쪽 분량을 강제로 1쪽에 압축하면
    # 글자를 읽을 수 없거나 내용을 임의로 잘라야 하므로 원문 충실성을 해친다. 따라서 이 서식만 예외적으로
    # 자연스러운 다중 페이지 출력을 허용하고(FitToPagesWide/Tall은 여전히 False로 유지해 강제 축소는
    # 하지 않음), 이 사실을 검증·문서에 명시적으로 기록한다.
    # 반영 범위: 학교명·학년도(C-01·C-02) 공통값과, F-001과 동일한 위원 반복행(기초자료입력 B58:C67)
    # 재사용, 교복 상한가격(F-007 C16과 동일한 D29:D34 합계 수식)만 자동 반영하고, 나머지 정책 설명
    # 본문은 학교마다 달라지지 않는 고정 안내문이므로 원문 그대로 반영한다.
    $wsF05 = $wbNew.Worksheets.Add(); $wsF05.Name = "F-005_구매추진계획안"; $ws = $wsF05
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 7~12쪽 [2-1] 대조. 원문이 6개 물리 페이지 분량이라 이 서식만 예외적으로 다중 페이지 자연 출력을 허용함(강제 압축 없음)"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $row = 3
    # 값이 "="으로 시작하면 수식으로, 그렇지 않으면 (앞뒤에 실수로 남은 " 문자를 제거하고) 리터럴
    # 텍스트로 셀에 쓴다. 이 섹션의 정적 안내문 다수가 애초에 Excel 수식 문자열 리터럴 표기("...")로
    # 작성됐으나 앞에 "="가 없어 수식으로 해석되지 않는 실수를 여기서 일괄 방어한다.
    function F005SetCell($cell, $v) {
        if ($null -eq $v) { return }
        $s = [string]$v
        if ($s.Length -ge 1 -and $s.Substring(0,1) -eq '=') {
            $cell.Formula = $s
        } else {
            if ($s.Length -ge 2 -and $s.Substring(0,1) -eq '"' -and $s.Substring($s.Length-1,1) -eq '"') {
                $s = $s.Substring(1, $s.Length - 2)
            }
            $cell.Value2 = $s
        }
    }
    # Excel의 Rows.AutoFit()은 한 행에 병합 셀이 여러 개 있거나 폭이 넓은 병합 셀의 줄바꿈 높이를
    # 정확히 계산하지 못하는 것으로 알려진 한계가 있으며(Microsoft 공식 문서에도 기재된 제약), 이번
    # F-005 1차 PDF 육안 확인에서 실제로 여러 줄로 감싼 문단·표 셀의 행 높이가 한 줄 높이(17.4pt)에
    # 머물러 뒷부분이 잘리는 결함으로 재현됨. AutoFit에 의존하지 않고 텍스트 길이·병합 폭 기준으로
    # 필요한 줄 수를 추정해 행 높이를 직접 계산한다(넉넉하게 잡아 잘림보다는 여백이 남는 쪽을 택함).
    function F005LineCount([string]$text, [int]$charsPerLine) {
        if ([string]::IsNullOrEmpty($text)) { return 1 }
        $n = [Math]::Ceiling($text.Length / [double]$charsPerLine)
        if ($n -lt 1) { $n = 1 }
        return [int]$n
    }
    function F005Row([string]$kind, $val1, $val2, $val3) {
        $script:row = $script:row + 1
        $r = $script:row
        switch ($kind) {
            'title' {
                $ws.Range("B${r}:G${r}").Merge() | Out-Null; F005SetCell $ws.Cells.Item($r, 2) $val1
                $ws.Range("B$r").Font.Size = 15; $ws.Range("B$r").Font.Bold = $true; $ws.Range("B$r").HorizontalAlignment = -4108
            }
            'h1' {
                $script:row = $script:row + 1; $r = $script:row
                $ws.Range("B${r}:G${r}").Merge() | Out-Null; F005SetCell $ws.Cells.Item($r, 2) $val1
                $ws.Range("B$r").Font.Bold = $true; $ws.Range("B$r").Font.Size = 12
                $ws.Range("B$r").Borders.Item(1).LineStyle = 1; $ws.Range("B$r").Borders.Item(4).LineStyle = 1
            }
            'h2' {
                $ws.Range("B${r}:G${r}").Merge() | Out-Null; F005SetCell $ws.Cells.Item($r, 2) $val1
                $ws.Range("B$r").Font.Bold = $true; $ws.Range("B$r").Font.Size = 10
            }
            'item' {
                $ws.Range("B${r}:G${r}").Merge() | Out-Null; F005SetCell $ws.Cells.Item($r, 2) $val1
                $ws.Range("B$r").WrapText = $true
                $txt = [string]$ws.Cells.Item($r, 2).Value2
                $lines = F005LineCount $txt 34
                $ws.Rows.Item($r).RowHeight = 13 * $lines + 4
            }
            'note' {
                $ws.Range("B${r}:G${r}").Merge() | Out-Null; F005SetCell $ws.Cells.Item($r, 2) $val1
                $ws.Range("B$r").Font.Size = 8; $ws.Range("B$r").WrapText = $true
                $txt = [string]$ws.Cells.Item($r, 2).Value2
                $lines = F005LineCount $txt 46
                $ws.Rows.Item($r).RowHeight = 11 * $lines + 4
            }
            'blank' { $ws.Rows.Item($r).RowHeight = 6 }
            'table3header' {
                $ws.Cells.Item($r, 2).Value2 = $val1; $ws.Cells.Item($r, 4).Value2 = $val2; $ws.Cells.Item($r, 7).Value2 = $val3
                $ws.Range("B${r}:C${r}").Merge() | Out-Null; $ws.Range("D${r}:F${r}").Merge() | Out-Null
                $ws.Range("B$r,D$r,G$r").Font.Bold = $true; $ws.Range("B$r,D$r,G$r").HorizontalAlignment = -4108
                $ws.Range("B${r}:G${r}").Borders.LineStyle = 1
            }
            'table3row' {
                F005SetCell $ws.Cells.Item($r, 2) $val1; F005SetCell $ws.Cells.Item($r, 4) $val2; F005SetCell $ws.Cells.Item($r, 7) $val3
                $ws.Range("B${r}:C${r}").Merge() | Out-Null; $ws.Range("D${r}:F${r}").Merge() | Out-Null
                $ws.Range("B$r:C$r").WrapText = $true; $ws.Range("D$r").WrapText = $true
                $ws.Range("B${r}:G${r}").Borders.LineStyle = 1
                $txtB = [string]$ws.Cells.Item($r, 2).Value2; $txtD = [string]$ws.Cells.Item($r, 4).Value2
                $linesB = F005LineCount $txtB 12; $linesD = F005LineCount $txtD 17
                $maxLines = [Math]::Max($linesB, $linesD)
                $ws.Rows.Item($r).RowHeight = 13 * $maxLines + 4
            }
        }
    }
    F005Row 'title' '="2026학년도 교복 학교주관구매 추진 계획(안)"' $null $null
    F005Row 'item' '=IF(기초자료입력!C4<>"",기초자료입력!C4&"장",  "○○○학교장")' $null $null
    $ws.Range("B$row").HorizontalAlignment = -4108
    F005Row 'blank' $null $null $null
    F005Row 'h1' '="1  개요"' $null $null
    F005Row 'item' '="❍ 사업명："&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매"' $null $null
    F005Row 'item' '"❍ 근 거：학교안전과-○○(20○○. ○○. ○○.) 「교복 학교주관구매 운영 요령」"' $null $null
    F005Row 'item' '"❍ 목 적：교복 학교주관구매 지원을 통한 학부모의 경제적 부담 경감"' $null $null
    F005Row 'item' '"❍ 시 기：20○○. ○○월 ~ 20○○. ○○월"' $null $null
    F005Row 'item' '"❍ 대 상：해당 학년도 모든 신입생"' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h1' '="2  기본방향"' $null $null
    F005Row 'item' '"❍ 학교장이 업체와 계약 체결 후 교복을 일괄 구매하는 학교주관구매 실시"' $null $null
    F005Row 'item' '"❍ 교복의 품질과 가격을 종합적으로 심사할 수 있는 2단계 입찰로 투명성·공정성·가격 경쟁력 확보"' $null $null
    F005Row 'item' '"❍ 교원, 학생대표, 학부모 대표 등으로 「교복선정위원회」를 구성·운영하여 명확하고 합리적인 심사기준 설정"' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h1' '="3  추진 절차 개요"' $null $null
    F005Row 'table3header' '단계 및 시기' '주요 추진 사항' '담당'
    F005Row 'table3row' '"교복 구매 계획 수립 [4월]"' '"교복 구매 계획 수립(추진일정, 디자인, 품질기준, 상한가격, 계약방식 등 포함)"' '"사업부서"'
    F005Row 'table3row' '"학교운영위원회 심의 [4월~5월]"' '"교복 구매 계획 학교운영위원회 심의"' '"사업부서"'
    F005Row 'table3row' '"입찰 공고 [6월~7월]"' '"구매물량 확정 단서 명시, A/S·반품·교환 요청사항 포함, 납품기한 명시"' '"계약부서"'
    F005Row 'table3row' '"1단계 품질심사(제안서 평가)"' '"교복선정위원회 품질(제안서) 심사"' '"사업부서(계약부서)"'
    F005Row 'table3row' '"2단계 가격 경쟁"' '"품질 적격업체 대상 가격 입찰"' '"계약부서"'
    F005Row 'table3row' '"사업자 선정·계약 체결 [7월~8월]"' '"낙찰통지 후 계약서 작성(단가계약 체결)"' '"사업부서(계약부서)"'
    F005Row 'table3row' '"학교주관구매 홍보 [12월~다음연도 1월]"' '"신입생 배정·합격자 통지 시 구매 방법 안내"' '"사업부서"'
    F005Row 'table3row' '"구성 인원 확정 및 치수 측정 [다음연도 1~2월]"' '"물량 확정 후 40일의 납품기한 보장"' '"사업부서"'
    F005Row 'table3row' '"교복 구매/납품 [다음연도 2월]"' '"교내 또는 업체 매장에서 검사·검수 실시"' '"사업부서(계약부서)"'
    F005Row 'table3row' '"대금 지급 [다음연도 3월]"' '"학교회계 절차에 따라 세입·세출 처리"' '"계약부서"'
    F005Row 'table3row' '"검사·검수 및 평가·환류 [다음연도 2~3월]"' '"계약 내용 확인 후 검수, 교복 만족도 조사"' '"사업부서(계약부서)"'
    F005Row 'blank' $null $null $null
    F005Row 'h1' '="4  세부 추진 절차"' $null $null
    F005Row 'h2' '"❍ 교복디자인 선정·공개"' $null $null
    F005Row 'item' '"1) 신입생 교복은 원칙적으로 전년도 교복 디자인으로 한다."' $null $null
    F005Row 'item' '"2) 교복 형태에 관한 사항을 학생 생활규정으로 명문화하고, 사양서와 사진, 규정 등을 학교 규정집과 학교 누리집 공지사항 등을 통해 공개한다."' $null $null
    F005Row 'note' '"※ 교복 디자인 변경 시 도교육청 누리집 등을 통해 공개(최소 입찰공고 1개월 전까지)"' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h2' '"❍ 교복선정위원회 구성 및 역할"' $null $null
    F005Row 'item' '"1) 교복선정위원회 구성 현황(위원 성명은 개인정보이므로 마스킹 식별표시만 표시함)"' $null $null
    $r = $script:row + 1; $script:row = $r
    $ws.Cells.Item($r, 2).Value2 = "순"; $ws.Cells.Item($r, 3).Value2 = "구분"; $ws.Cells.Item($r, 4).Value2 = "직급(직위)"; $ws.Cells.Item($r, 5).Value2 = "성명"; $ws.Cells.Item($r, 6).Value2 = "비고"
    $ws.Range("B${r}:G${r}").Font.Bold = $true; $ws.Range("B${r}:G${r}").HorizontalAlignment = -4108; $ws.Range("F${r}:G${r}").Merge() | Out-Null
    $ws.Range("B${r}:G${r}").Borders.LineStyle = 1
    for ($i = 0; $i -lt 10; $i++) {
        $r = $script:row + 1; $script:row = $r
        $src = 58 + $i
        $roleLabel = if ($i -eq 0) { "위원장" } else { "위원" }
        $ws.Cells.Item($r, 2).Formula = "=IF(기초자료입력!B${src}<>`"`",$($i+1),`"`")"
        $ws.Cells.Item($r, 3).Formula = "=IF(기초자료입력!B${src}<>`"`",`"$roleLabel`",`"`")"
        $ws.Cells.Item($r, 4).Formula = "=IF(기초자료입력!B${src}<>`"`",기초자료입력!B${src},`"`")"
        $ws.Cells.Item($r, 5).Formula = "=IF(기초자료입력!C${src}<>`"`",기초자료입력!C${src},`"`")"
        $ws.Range("F${r}:G${r}").Merge() | Out-Null
        $ws.Range("B${r}:G${r}").Borders.LineStyle = 1
    }
    F005Row 'item' '"2) 역할：(제안서 심사) 디자인 적합성·재질·바느질 상태·A/S(의무기간 1년)·하자 이행 등 품질심사 기준 설정 및 제안서 평가·적격업체 선정 / (제작 감독) 계약 체결 후 제작 상황 확인(샘플검사·원단검사 등) / (검수) 납품 시 수량·사양서 등 계약 내용 확인 후 검수, 품질검사 실시 / (평가) 착용 이후 학생·학부모 의견수렴을 거쳐 구매 결과 평가 실시"' $null $null
    F005Row 'note' '"※ 교복 업계 이해관계자는 위촉 불가, 위원 후보자로부터 확인서·서약서 징구. 필요시 교복 제조·판매 시설 실사 실시"' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h2' '"❍ 교복 구매 기준"' $null $null
    F005Row 'item' '"1) 교복 구매 시 품질기준, 디자인, 상한 가격, 구매절차 및 입찰방법, 참가자 신청 일정 및 비용 납부, A/S 방안 등에 관한 사항을 포함한다."' $null $null
    F005Row 'item' '="※ 상한가격(참고용 자동 계산)：동복(4pcs 기준)："&IF(SUM(기초자료입력!$D$29:$D$32)=0,"○○○,○○○",TEXT(SUM(기초자료입력!$D$29:$D$32),"#,##0"))&"원 / 하복(2pcs 기준)："&IF(SUM(기초자료입력!$D$33:$D$34)=0,"○○,○○○",TEXT(SUM(기초자료입력!$D$33:$D$34),"#,##0"))&"원 (계약일 기준 적용기간은 교육청 고시에 따름)"' $null $null
    F005Row 'item' '"2) 신입생 3월 동복 착용 시 교복 제작 일정을 감안하여 8월까지 사업자를 선정하여 추진한다."' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h2' '"❍ 입찰 공고"' $null $null
    F005Row 'item' '"1) 교복 구매 사업자 선정을 위한 입찰공고문을 작성하여 교복사양서와 함께 학교 누리집 및 지정 정보처리장치 등에 게재한다."' $null $null
    F005Row 'item' '"2) 입찰에 부치는 사항, 입찰·개찰의 장소 및 일시, 참가자격, 낙찰자 결정방법, 계약 방법, 이행예정 기간 등을 명시하되, 구매물량은 「학생 인원에 따라 구매 물량 확정」이라는 단서를 반드시 명시한다."' $null $null
    F005Row 'item' '"3) 물량변동 분쟁을 방지하도록 단가계약을 체결하되, 전입생 등 추가 구매가 가능하도록 충분한 수량을 확보한다."' $null $null
    F005Row 'item' '"4) 입찰 참가 사업자는 학교에 입찰 참가 신청서를 제출한다."' $null $null
    F005Row 'item' '"5) 기초가격은 교육청 상한 가격을 넘지 않는 선에서 교복사양(디자인·원단 등)을 고려한 거래실례가격·견적가격 등을 감안해 산출한다."' $null $null
    F005Row 'item' '"6) 입찰 공고 등에 관한 사항은 관련 법령과 규정을 준수하여 진행한다."' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h2' '"❍ 교복 납품 사업자 선정"' $null $null
    F005Row 'item' '"1) 교복선정위원회가 응찰 사업자를 대상으로 2단계 입찰(품질 → 가격)을 심사한다."' $null $null
    F005Row 'item' '"2) 품질 심사：가격 심사 참가를 위한 최소 품질기준 합격 여부를 심사하며, 명확한 기준에 따라 엄격히 심사하고 Q-MARK 검사 기준 등을 적극 활용한다."' $null $null
    F005Row 'item' '"3) 가격 경쟁：품질 합격 업체를 대상으로 입찰 관련 법령을 준수해 가격 경쟁을 실시하며, 상한 가격 초과·지역제한 미충족 업체는 탈락시킨다."' $null $null
    F005Row 'item' '"4) 사업자 선정：최종 낙찰자를 사업자로 선정하고 공급 단가만 계약하되 「구매물량은 신청 학생 인원에 따라 확정」이라는 단서를 반드시 명시한다(단가계약 체결)."' $null $null
    F005Row 'note' '"※ 계약 시 확인 사항：무상 A/S 기한(1년 의무기간)·처리과정, 납품 기일 설정 및 계약이행지급각서 등 제출, 원단 검사 시 의류 시험성적서 제출, 납품 시점 하자이행보증보험증권 제출, 계약 특수조건(전입생 추가 구매 물량 확보) 명시"' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h2' '"❍ 규모 확정 및 학교회계 처리"' $null $null
    F005Row 'item' '"1) 3월 입학 동복 착용의 경우 예비소집일 등을 통해 신속히 참여 규모를 확정하고 회계처리한다."' $null $null
    F005Row 'note' '"※ 신청 시 희망 치수와 실제측정 치수를 함께 조사한다."' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h2' '"❍ 검수 및 평가·환류"' $null $null
    F005Row 'item' '"1) 계약 체결 후 원단 구입 전 샘플 검사(1차), 원단 검사(2차) 등 제작 상황을 수시로 확인한다."' $null $null
    F005Row 'item' '"2) 물품 검수 시 수량·사양서 등 계약내용 및 제조 연월 표시 여부(라벨 불량·품질 표시 등)를 철저히 확인 후 검수한다."' $null $null
    F005Row 'item' '"3) 교복 품질 검사 실시 후 검수한다."' $null $null
    F005Row 'item' '"4) 착용 이후 학생·학부모 의견 수렴을 거쳐 「교복선정위원회」에서 구매 결과 평가를 실시한다."' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'h1' '="5  기대효과"' $null $null
    F005Row 'item' '"❍ 공정하고 투명한 입찰 방법을 통하여 저렴하고 높은 품질의 교복을 확보하고, 학교회계 관련 법령 및 절차에 따른 진행으로 교복구매의 합리성을 도모한다."' $null $null
    F005Row 'blank' $null $null $null
    F005Row 'item' '"붙임  교복 사양서 1부.  끝."' $null $null
    $lastRowF005 = $script:row
    $ws.Range("B3:G$lastRowF005").Font.Size = 9
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 13; $ws.Columns.Item("C").ColumnWidth = 11; $ws.Columns.Item("D").ColumnWidth = 13; $ws.Columns.Item("E").ColumnWidth = 11; $ws.Columns.Item("F").ColumnWidth = 11; $ws.Columns.Item("G").ColumnWidth = 13
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$G`$$lastRowF005"
    $hbF05 = $ws.HPageBreaks.Count; $vbF05 = $ws.VPageBreaks.Count
    L "F-005 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF05(다중 페이지 허용, 예외), VPageBreaks=$vbF05(0이어야 폭 초과 없음 — 세로 폭 초과만 결함)"
    L "F-005 원문 구매 추진 계획안 대조(6개 물리 페이지 원문, 다중 페이지 자연 출력 예외 적용) 완료"

    # ---- 7r. F-006 교복 학교주관구매 계획 학교운영위원회 심의(안) ----
    # 원문 HWPX 13쪽 [3]을 대조한다. 단일 물리 페이지. 안건번호/제출일자/제안자/제안설명자는 원문 자체가
    # 빈 칸(회의 소집 시 수기 기재)이라 자동 반영하지 않는다. "라. 협의사항"의 위원회 구성 표는 F-001과
    # 완전히 동일한 순/구분/직위/성명/비고 구조이므로 같은 위원 반복행(기초자료입력 B58:C67)을 재사용한다.
    # "다. 교복구매 상한가격"은 F-007/F-005와 동일한 D29:D34 합계 수식을 재사용한다. 문서 끝의
    # "※ 참고 서식이니 해당학교의 학교운영위원회 서식을 사용할 것" 안내문을 그대로 반영한다.
    $wsF06 = $wbNew.Worksheets.Add(); $wsF06.Name = "F-006_운영위원회심의안"; $ws = $wsF06
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 13쪽 [3] 대조. 참고 서식이며 학교운영위원회 자체 서식 사용을 권장함(원문 안내문 그대로 반영)"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 학교주관구매 계획(안)"'
    $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B6").Value2 = "안건번호"; $ws.Range("C6:H6").Merge() | Out-Null
    $ws.Range("B7").Value2 = "제출일자"; $ws.Range("C7:H7").Merge() | Out-Null
    $ws.Range("B8").Value2 = "제 안 자"; $ws.Range("C8:H8").Merge() | Out-Null
    $ws.Range("B9").Value2 = "제안설명자"; $ws.Range("C9:H9").Merge() | Out-Null
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 제안 이유"; $ws.Range("B11").Font.Bold = $true; $ws.Range("B11").Font.Size = 12
    $ws.Range("B12:H12").Merge() | Out-Null; $ws.Range("B12").Value2 = "❍ 근 거：「초·중등교육법」 제30조의2, 제32조 제1항"
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Formula = '="❍ 교복 가격의 안정화 및 교복 구매의 효율성 제고를 위한 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" "&IF(기초자료입력!C5<>"",기초자료입력!C5,"○○")&"학년도 교복 학교주관구매 계획(안)을 심의하고자 함."'
    $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 28
    $ws.Range("B15:H15").Merge() | Out-Null; $ws.Range("B15").Value2 = "2. 주요 내용"; $ws.Range("B15").Font.Bold = $true; $ws.Range("B15").Font.Size = 12
    $ws.Range("B16:H16").Merge() | Out-Null; $ws.Range("B16").Value2 = "가. 목적"; $ws.Range("B16").Font.Bold = $true
    $ws.Range("B17:H17").Merge() | Out-Null; $ws.Range("B17").Value2 = "❍ 투명하고 공정한 선정 절차로 우수한 품질과 가격 경쟁력 확보"
    $ws.Range("B18:H18").Merge() | Out-Null; $ws.Range("B18").Value2 = "❍ 학부모 교육비 부담을 경감하여 교복 가격 안정화 도모"
    $ws.Range("B19:H19").Merge() | Out-Null; $ws.Range("B19").Value2 = "나. 교복 납품 사업자 선정 방법"; $ws.Range("B19").Font.Bold = $true
    $ws.Range("B20:H20").Merge() | Out-Null; $ws.Range("B20").Value2 = "❍ 교복선정위원회를 구성하여 교복에 대한 2단계 경쟁(품질 → 가격)을 심사"
    $ws.Range("B21:H21").Merge() | Out-Null; $ws.Range("B21").Value2 = "다. 교복구매 상한가격"; $ws.Range("B21").Font.Bold = $true
    $ws.Range("B22:H22").Merge() | Out-Null
    $ws.Range("B22").Formula = '="❍ 하복："&IF(SUM(기초자료입력!$D$33:$D$34)=0,"○○,○○○",TEXT(SUM(기초자료입력!$D$33:$D$34),"#,##0"))&"원, 동복："&IF(SUM(기초자료입력!$D$29:$D$32)=0,"○○○,○○○",TEXT(SUM(기초자료입력!$D$29:$D$32),"#,##0"))&"원("&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"년 기준)"'
    $ws.Range("B23:H23").Merge() | Out-Null; $ws.Range("B23").Value2 = "라. 협의사항"; $ws.Range("B23").Font.Bold = $true
    $ws.Range("B24:H24").Merge() | Out-Null; $ws.Range("B24").Value2 = "❍ 교복선정위원회 구성"
    $headersF06 = @("순","구분","직위","성명","비고")
    $colsF06 = @("B","C","D","E","F")
    for ($i = 0; $i -lt $headersF06.Count; $i++) { $ws.Range("$($colsF06[$i])25").Value2 = $headersF06[$i] }
    $ws.Range("F25:H25").Merge() | Out-Null
    $ws.Range("B25:H25").Font.Bold = $true; $ws.Range("B25:F25").HorizontalAlignment = -4108
    for ($i = 0; $i -lt 10; $i++) {
        $r = 26 + $i; $src = 58 + $i
        $roleLabel = if ($i -eq 0) { "위원장" } else { "위원" }
        $ws.Range("B$r").Formula = "=IF(기초자료입력!B${src}<>`"`",$($i+1),`"`")"
        $ws.Range("C$r").Formula = "=IF(기초자료입력!B${src}<>`"`",`"$roleLabel`",`"`")"
        $ws.Range("D$r").Formula = "=IF(기초자료입력!B${src}<>`"`",기초자료입력!B${src},`"`")"
        $ws.Range("E$r").Formula = "=IF(기초자료입력!C${src}<>`"`",기초자료입력!C${src},`"`")"
        $ws.Range("F${r}:H${r}").Merge() | Out-Null
    }
    $ws.Range("B36:H36").Merge() | Out-Null; $ws.Range("B36").Value2 = "❍ 기타 의견 수렴"
    $ws.Range("B38:H38").Merge() | Out-Null
    $ws.Range("B38").Formula = '="붙임  "&IF(기초자료입력!C5<>"",기초자료입력!C5,"○○")&"학년도 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 학교주관구매 계획안 1부.  끝."'
    $ws.Range("B40:H40").Merge() | Out-Null; $ws.Range("B40").Value2 = "※ 참고 서식이니 해당학교의 학교운영위원회 서식을 사용할 것"; $ws.Range("B40").Font.Size = 8
    $ws.Range("B6:H9,B25:H35").Borders.LineStyle = 1
    $ws.Range("B3:H40").Font.Size = 10
    $ws.Range("B6,B7,B8,B9").Font.Bold = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 10; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$40"
    $hbF06 = $ws.HPageBreaks.Count; $vbF06 = $ws.VPageBreaks.Count
    L "F-006 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF06, VPageBreaks=$vbF06 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-006 원문 학교운영위원회 심의(안) 대조 레이아웃(위원 성명 마스킹, DB_위원 재사용, 참고 서식 안내문 반영) 적용 완료"

    # ---- 7s. F-041 낙찰자 결정 ----
    # 원문 HWPX 70쪽 [14]를 대조한다. F-007/F-001과 동일한 기안문 틀(수신=내부결재·경유·제목·결재란)에
    # "라. 낙찰자" 표(순위/업체명/대표자/투찰금액/낙찰률/비고(주소))가 추가된 구조다. 낙찰 업체는 기존
    # 업체 반복행(기초자료입력!B44:B53, DB_업체, F-015~F-025가 이미 쓰던 R-03 공용 저장소)에서 담당자가
    # I3 순번으로 1곳만 선택해 반영한다 — F-017~F-025처럼 업체마다 별도 문서를 만드는 것이 아니라 낙찰자는
    # 언제나 1개 업체이므로, F-015의 "선택 업체 순번" 단일 선택 패턴을 재사용한다(Module_출력에서 업체
    # 반복 루프를 돌리지 않고 F-001/F-006과 같은 단일 출력으로 취급함). 대표자·비고(주소)는 개인정보·업체
    # 상세정보이므로 F-019~F-023과 동일하게 빈칸(수기 작성)으로 유지한다. 예정가격·투찰금액·낙찰률은
    # `입력데이터_사전.md`에 정의되지 않은 낙찰 결과 고유값이라 기초자료입력에 새 필드를 추가하지 않고
    # 이 시트에 직접 입력영역(연노랑)을 둔다. `서식_매핑표.md`의 F-041 계산 필드(K-02 합계금액)는 원문에
    # 반복 품목·합계 구조가 없어 실제로는 적용되지 않는 매핑 불일치로 확인됨(기록만 하고 매핑표는 수정하지
    # 않음 — F-004~F-006·F-023·F-025와 같은 종류의 기존 발견 패턴).
    $wsF41 = $wbNew.Worksheets.Add(); $wsF41.Name = "F-041_낙찰자결정"; $ws = $wsF41
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 70쪽 [14] 대조. 낙찰 업체는 I3 순번으로 1곳만 선택 반영하며 대표자·비고(주소)는 수기 작성함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("I2").Value2 = "낙찰 업체 선택 순번(자동, 인쇄 전용)"; $ws.Range("I2").Font.Size = 7; $ws.Range("I3").Value2 = 1
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복 학교주관구매 낙찰자 결정(안)"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"년도 교복 학교주관구매 낙찰자 결정")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("C11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 관련："; $ws.Range("C11").Formula = '=IF(기초자료입력!C23<>"",기초자료입력!C23,"")'
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"년도 교복 학교주관구매 건에 대해 개찰 결과 예정가격 대비 최저가격으로 투찰하여 1순위가 된 아래 업체를 낙찰자로 결정하고자 합니다."'; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 28
    $ws.Range("B15").Value2 = "가. 건  명："; $ws.Range("C15:H15").Merge() | Out-Null; $ws.Range("C15").Formula = '=IF(기초자료입력!C12<>"",IF(기초자료입력!C5<>"",기초자료입력!C5&"학년도 ","")&기초자료입력!C12,"")'
    $ws.Range("B16").Value2 = "나. 기초금액："; $ws.Range("C16:H16").Merge() | Out-Null; $ws.Range("C16").Formula = '=IF(기초자료입력!C15<>"","금"&TEXT(기초자료입력!C15,"#,##0")&"원","")'
    $ws.Range("B17").Value2 = "다. 예정가격："; $ws.Range("C17:H17").Merge() | Out-Null; $ws.Range("C17").Interior.Color = 16777164
    $ws.Range("B19:H19").Merge() | Out-Null; $ws.Range("B19").Value2 = "라. 낙찰자"; $ws.Range("B19").Font.Bold = $true
    $headersF41 = @("순위","업체명","대표자","투찰금액(원)","낙찰률(%)","비고(주소)")
    $colsF41 = @("B","C","D","E","F","G")
    for ($i = 0; $i -lt $headersF41.Count; $i++) { $ws.Range("$($colsF41[$i])20").Value2 = $headersF41[$i] }
    $ws.Range("B21").Value2 = 1
    $ws.Range("C21").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'
    $ws.Range("E21").NumberFormat = "#,##0"; $ws.Range("F21").NumberFormat = "0.00"
    $ws.Range("E21:F21").Interior.Color = 16777164
    $ws.Range("B23:H23").Merge() | Out-Null; $ws.Range("B23").Value2 = "마. 근거：「지방자치단체를 당사자로 하는 계약에 관한 법률 시행령」 제42조.  끝."; $ws.Range("B23").WrapText = $true
    $ws.Range("F26").Value2 = "담당"; $ws.Range("G26").Value2 = "협조자"; $ws.Range("H26").Value2 = "교장"
    $ws.Range("B27").Value2 = "시행"; $ws.Range("C27:E27").Merge() | Out-Null; $ws.Range("C27").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F27").Value2 = "접수"; $ws.Range("G27:H27").Merge() | Out-Null
    $ws.Range("B28").Value2 = "우편번호"; $ws.Range("D28").Value2 = "주소"; $ws.Range("E28:H28").Merge() | Out-Null
    $ws.Range("B29").Value2 = "전화"; $ws.Range("D29").Value2 = "전송(팩스)"; $ws.Range("F29").Value2 = "이메일"; $ws.Range("G29:H29").Merge() | Out-Null
    $ws.Range("B5:H29").Font.Size = 9
    $ws.Range("B5,E5,B7,B8,B9,B11,B15,B16,B17,B19,B20:G20,B23,F26:H26,B27,F27,B28,D28,B29,D29,F29").Font.Bold = $true
    $ws.Range("B20:G20").HorizontalAlignment = -4108
    $ws.Range("B5:H9,B20:G21,F26:H29").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 9; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$29"
    $hbF41 = $ws.HPageBreaks.Count; $vbF41 = $ws.VPageBreaks.Count
    L "F-041 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF41, VPageBreaks=$vbF41 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-041 원문 낙찰자 결정 대조 레이아웃(낙찰 업체 1곳 선택 반영, 대표자·비고 수기 경계) 적용 완료"

    # ---- 7t. F-042 낙찰자 결정 통보 ----
    # 원문 HWPX 71쪽 [15]를 대조한다. F-041과 달리 수신이 "내부결재"가 아니라 낙찰 업체명 자체다(통보문
    # 유형). F-041과 동일한 I3 업체 선택 패턴을 재사용해 수신·가.낙찰자 표의 업체명을 반영한다. 제출기한은
    # 기존 기초자료입력!C18(B-07 제출기한, F-014~F-016 설계 시 이미 마련돼 있었으나 F-042가 처음 소비함)을
    # 재사용한다. 대표자·비고는 F-041과 동일하게 빈칸으로 유지하고, 단가·계약체결금액은
    # `서식_매핑표.md`가 F-042의 계산 필드를 "없음"으로 명시해 자동 계산하지 않고 담당자 직접 입력으로 둔다.
    $wsF42 = $wbNew.Worksheets.Add(); $wsF42.Name = "F-042_낙찰자결정통보"; $ws = $wsF42
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 71쪽 [15] 대조. 수신·업체명은 I3 순번으로 선택한 낙찰 업체를 반영하며 대표자·비고는 수기 작성함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("I2").Value2 = "낙찰 업체 선택 순번(자동, 인쇄 전용)"; $ws.Range("I2").Font.Size = 7; $ws.Range("I3").Value2 = 1
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복 학교주관구매 낙찰자 결정 통보"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"년도 교복 학교주관 구매 낙찰자 결정 통보")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 귀사의 무궁한 발전을 기원합니다."
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. 우리학교 "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 건에 대하여 귀사를 낙찰자로 결정하여 통보하오니 전자계약[나라장터(G2B)] 체결을 위해 아래 서류를 "&IF(기초자료입력!C18<>"",TEXT(기초자료입력!C18,"yyyy.mm.dd."),"20○○.○○.○○")&"까지 제출하여 주시기 바랍니다."'; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 40
    $ws.Range("B15:H15").Merge() | Out-Null; $ws.Range("B15").Value2 = "가. 낙찰자"; $ws.Range("B15").Font.Bold = $true
    $headersF42 = @("업체명","대표자","단가(원)","예정수량","계약체결금액(원)","비고")
    $colsF42 = @("B","C","D","E","F","G")
    for ($i = 0; $i -lt $headersF42.Count; $i++) { $ws.Range("$($colsF42[$i])16").Value2 = $headersF42[$i] }
    $ws.Range("B17").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'
    $ws.Range("D17").NumberFormat = "#,##0"; $ws.Range("D17").Interior.Color = 16777164
    $ws.Range("E17").Formula = '=IF(기초자료입력!C16<>"",기초자료입력!C16,"")'
    $ws.Range("F17").NumberFormat = "#,##0"; $ws.Range("F17").Interior.Color = 16777164
    $ws.Range("B19:H19").Merge() | Out-Null; $ws.Range("B19").Value2 = "나. 제출서류"; $ws.Range("B19").Font.Bold = $true
    $ws.Range("B20:H20").Merge() | Out-Null; $ws.Range("B20").Value2 = "1) 품목별 산출내역서 1부."
    $ws.Range("B21:H21").Merge() | Out-Null; $ws.Range("B21").Value2 = "2) 원단 의류 시험성적서 1부."
    $ws.Range("B22:H22").Merge() | Out-Null; $ws.Range("B22").Value2 = "3) 사용인장계(조달등록 인장 아닌 경우) 1부."
    $ws.Range("B23:H23").Merge() | Out-Null; $ws.Range("B23").Value2 = "4) 법인 등기사항증명서(법인) 또는 주민등록등본(개인) 1부.  끝."
    $ws.Range("F26").Value2 = "담당"; $ws.Range("G26").Value2 = "협조자"; $ws.Range("H26").Value2 = "교장"
    $ws.Range("B27").Value2 = "시행"; $ws.Range("C27:E27").Merge() | Out-Null; $ws.Range("C27").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F27").Value2 = "접수"; $ws.Range("G27:H27").Merge() | Out-Null
    $ws.Range("B28").Value2 = "우편번호"; $ws.Range("D28").Value2 = "주소"; $ws.Range("E28:H28").Merge() | Out-Null
    $ws.Range("B29").Value2 = "전화"; $ws.Range("D29").Value2 = "전송(팩스)"; $ws.Range("F29").Value2 = "이메일"; $ws.Range("G29:H29").Merge() | Out-Null
    $ws.Range("B5:H29").Font.Size = 9
    $ws.Range("B5,E5,B7,B8,B9,B11,B15,B16:G16,B19,F26:H26,B27,F27,B28,D28,B29,D29,F29").Font.Bold = $true
    $ws.Range("B16:G16").HorizontalAlignment = -4108
    $ws.Range("F16").Font.Size = 8
    $ws.Range("B5:H9,B16:G17,F26:H29").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 12; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ws.Columns.Item("F").ColumnWidth = 13; $ws.Columns.Item("C").ColumnWidth = 7
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$29"
    $hbF42 = $ws.HPageBreaks.Count; $vbF42 = $ws.VPageBreaks.Count
    L "F-042 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF42, VPageBreaks=$vbF42 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-042 원문 낙찰자 결정 통보 대조 레이아웃(수신·업체명 1곳 선택 반영, 제출기한 B-07 재사용, 대표자·비고 수기 경계) 적용 완료"

    # ---- 7u. F-043 계약체결 ----
    # 원문 HWPX 72쪽 [16]를 대조한다. F-041과 동일한 기안문 틀(수신=내부결재)에 "나. 계약금액：금00,000,000원
    # [000,000원(단가)×○○벌(예정수량)]" 산식이 있어, 이 서식만 유일하게 `서식_매핑표.md`의 K-01(품목별금액=
    # 수량×단가)·K-02(합계금액) 계산 필드를 실제로 구현한다 — 단가는 이 시트의 새 입력영역(연노랑, 담당자
    # 직접 입력)이고 예정수량은 기존 기초자료입력!C16(B-05)을 재사용한다. "3. 계약상대자" 표(업체명/소재지/
    # 대표자/전화번호)는 F-041/F-042와 동일한 I3 업체 선택 패턴으로 업체명만 반영하고 소재지·대표자·전화번호는
    # F-019~F-023과 같이 빈칸(수기 작성)으로 유지한다. "5. 납품기한"은 원문이 동복/하복 날짜를 따로 표기하나
    # 기초자료입력에는 단일 납품기한(B-06, C17)만 있어 하나의 값으로만 반영함을 한계로 기록한다.
    $wsF43 = $wbNew.Worksheets.Add(); $wsF43.Name = "F-043_계약체결"; $ws = $wsF43
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 72쪽 [16] 대조. 계약상대자 업체명은 I3 순번으로 선택 반영하며 소재지·대표자·전화번호는 수기 작성함. 납품기한은 기초자료입력의 단일 값만 반영(원문 동복/하복 구분 한계)"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("I2").Value2 = "계약상대자 선택 순번(자동, 인쇄 전용)"; $ws.Range("I2").Font.Size = 7; $ws.Range("I3").Value2 = 1
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복 학교주관구매 계약 체결(안)"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 계약 체결")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("C11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 관련："; $ws.Range("C11").Formula = '=IF(기초자료입력!C23<>"",기초자료입력!C23,"")'
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매를 위한 계약을 아래와 같이 체결하고자 합니다."'; $ws.Range("B13").WrapText = $true
    $ws.Range("B15").Value2 = "가. 계약명："; $ws.Range("C15:H15").Merge() | Out-Null; $ws.Range("C15").Formula = '=IF(기초자료입력!C12<>"",IF(기초자료입력!C5<>"",기초자료입력!C5&"학년도 ","")&기초자료입력!C12,"")'
    $ws.Range("B16").Value2 = "단가(원) 입력："; $ws.Range("C16:D16").Merge() | Out-Null; $ws.Range("C16").NumberFormat = "#,##0"; $ws.Range("C16").Interior.Color = 16777164
    $ws.Range("E16:H16").Merge() | Out-Null; $ws.Range("E16").Value2 = "(담당자가 계약 단가를 직접 입력함 — 기초자료입력 미연동, K-01)"; $ws.Range("E16").Font.Size = 8
    $ws.Range("B17").Value2 = "나. 계약금액："; $ws.Range("C17:H17").Merge() | Out-Null
    $ws.Range("C17").Formula = '=IF(AND(C16<>"",기초자료입력!C16<>""),"금"&TEXT(C16*기초자료입력!C16,"#,##0")&"원["&TEXT(C16,"#,##0")&"원(단가)×"&기초자료입력!C16&"벌(예정수량)]","")'
    $ws.Range("B18:H18").Merge() | Out-Null; $ws.Range("B18").Value2 = "※ 최종구매수량은 희망 학생 수에 따라 증감될 수 있음."; $ws.Range("B18").Font.Size = 8
    $ws.Range("B20:H20").Merge() | Out-Null; $ws.Range("B20").Value2 = "3. 계약상대자"; $ws.Range("B20").Font.Bold = $true
    $headersF43 = @("업체명","소재지","대표자","전화번호")
    $colsF43 = @("B","C","D","E")
    for ($i = 0; $i -lt $headersF43.Count; $i++) { $ws.Range("$($colsF43[$i])21").Value2 = $headersF43[$i] }
    $ws.Range("F21:H21").Merge() | Out-Null
    $ws.Range("B22").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'
    $ws.Range("F22:H22").Merge() | Out-Null
    $ws.Range("B24:H24").Merge() | Out-Null; $ws.Range("B24").Value2 = "4. 계약보증금：교복단가에 예정 수량을 곱한 금액의 10%(보증서 전자 접수)"; $ws.Range("B24").WrapText = $true
    $ws.Range("B25:H25").Merge() | Out-Null; $ws.Range("B25").Formula = '="5. 납품기한："&IF(기초자료입력!C17<>"",TEXT(기초자료입력!C17,"yyyy.mm.dd."),"")'
    $ws.Range("B26:H26").Merge() | Out-Null; $ws.Range("B26").Value2 = "6. 계약방법：나라장터(G2B)를 통한 전자계약"
    $ws.Range("B28:H28").Merge() | Out-Null; $ws.Range("B28").Value2 = "붙임 1. 계약서(초안) 1부."
    $ws.Range("B29:H29").Merge() | Out-Null; $ws.Range("B29").Value2 = "      2. 교복 사양서 1부."
    $ws.Range("B30:H30").Merge() | Out-Null; $ws.Range("B30").Value2 = "      3. 교복 구매계약 특수조건 1부."
    $ws.Range("B31:H31").Merge() | Out-Null; $ws.Range("B31").Value2 = "      4. 입찰가격산출내역서(품목별 단가표) 1부."
    $ws.Range("B32:H32").Merge() | Out-Null; $ws.Range("B32").Value2 = "      5. 기타서류(별첨).  끝."
    $ws.Range("F34").Value2 = "담당"; $ws.Range("G34").Value2 = "협조자"; $ws.Range("H34").Value2 = "교장"
    $ws.Range("B35").Value2 = "시행"; $ws.Range("C35:E35").Merge() | Out-Null; $ws.Range("C35").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F35").Value2 = "접수"; $ws.Range("G35:H35").Merge() | Out-Null
    $ws.Range("B36").Value2 = "우편번호"; $ws.Range("D36").Value2 = "주소"; $ws.Range("E36:H36").Merge() | Out-Null
    $ws.Range("B37").Value2 = "전화"; $ws.Range("D37").Value2 = "전송(팩스)"; $ws.Range("F37").Value2 = "이메일"; $ws.Range("G37:H37").Merge() | Out-Null
    $ws.Range("B5:H37").Font.Size = 9
    $ws.Range("B5,E5,B7,B8,B9,B11,B15,B16,B17,B20,B21:E21,B24,B26,F34:H34,B35,F35,B36,D36,B37,D37,F37").Font.Bold = $true
    $ws.Range("B21:E21").HorizontalAlignment = -4108
    $ws.Range("B5:H9,B21:E22,F34:H37").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 12; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$37"
    $hbF43 = $ws.HPageBreaks.Count; $vbF43 = $ws.VPageBreaks.Count
    L "F-043 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF43, VPageBreaks=$vbF43 (둘 다 0이면 A4 1쪽 자연 충족)"
    L "F-043 원문 계약체결 대조 레이아웃(계약상대자 업체명 1곳 선택 반영, 단가 입력×예정수량 K-01/K-02 계산, 소재지·대표자·전화번호 수기 경계) 적용 완료"

    # ---- 7v. F-044 사전 안내 가정통신문 안내 ----
    # 원본 HWPX 73쪽 [17] 대조. 학교명·학년도·문서번호·발행일·사업명만 공통값으로 반영한다.
    # 대상 학생·학부모의 성명·연락처 등 개인값과 학교별 배정 세부정보는 원문에도 없고 자동 처리하지 않는다.
    $wsF44 = $wbNew.Worksheets.Add(); $wsF44.Name = "F-044_사전안내가정통신문안내"; $ws = $wsF44
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 73쪽 [17] 대조. 학교·학년도 공통값만 반영하며 학생·학부모 개인정보는 입력·저장·출력하지 않음"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("G2:H2").Merge() | Out-Null; $ws.Range("G2").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○ ○ 학 교")'; $ws.Range("G2").Font.Bold = $true; $ws.Range("G2").HorizontalAlignment = -4108
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Value2 = "신입생 교복구매 사전 안내 가정통신문 발송"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재(사업부서)"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,"신입생 교복구매 사전 안내 가정통신문 발송")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 관련: 전북특별자치도교육청 학교안전과-0000(20○○.) 학교주관공동구매 요령"
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Value2 = "2. 상급학교 진학 이전에 교복 구매에 대한 사전 안내를 실시하기 위하여 붙임과 같이 가정통신문을 발송하고자 합니다."; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 30
    $ws.Range("B15").Value2 = "가. 대상:"; $ws.Range("C15:H15").Merge() | Out-Null; $ws.Range("C15").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 예비 중1(예비 고1) 학생 및 학부모"'
    $ws.Range("B16").Value2 = "나. 내용:"; $ws.Range("C16:H16").Merge() | Out-Null; $ws.Range("C16").Formula = '=IF(기초자료입력!C12<>"",기초자료입력!C12&" 관련 교복 학교주관구매 참여 및 지원 방법, 기타 교복관련 사항","교복 학교주관구매 참여 및 지원 방법, 기타 교복관련 사항")'
    $ws.Range("B18:H18").Merge() | Out-Null; $ws.Range("B18").Value2 = "붙임 교복구매 사전 안내 가정통신문 1부. 끝."
    $ws.Range("F22").Value2 = "담당"; $ws.Range("G22").Value2 = "협조자"; $ws.Range("H22").Value2 = "교장"
    $ws.Range("B23").Value2 = "시행"; $ws.Range("C23:E23").Merge() | Out-Null; $ws.Range("C23").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F23").Value2 = "접수"; $ws.Range("G23:H23").Merge() | Out-Null
    $ws.Range("B24").Value2 = "우편번호"; $ws.Range("D24").Value2 = "주소"; $ws.Range("E24:H24").Merge() | Out-Null
    $ws.Range("B25").Value2 = "전화"; $ws.Range("D25").Value2 = "전송(팩스)"; $ws.Range("F25").Value2 = "이메일"; $ws.Range("G25:H25").Merge() | Out-Null
    $ws.Range("B5:H25").Font.Size = 9; $ws.Range("B5,E5,B7,B8,B9,B11,B15,B16,B18,F22:H22,B23,F23,B24,D24,B25,D25,F25").Font.Bold = $true
    $ws.Range("B5:H9,F22:H25").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 10; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$25"
    $hbF44 = $ws.HPageBreaks.Count; $vbF44 = $ws.VPageBreaks.Count
    L "F-044 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF44, VPageBreaks=$vbF44 (둘 다 0이면 A4 1쪽 자연 충족)"

    # ---- 7w. F-045 교복구매 사전 안내 가정통신문 ----
    # 원문 HWPX 74쪽 [17-1] 대조. 학교·학년도·구매명·제목만 공통값으로 반영한다.
    $wsF45 = $wbNew.Worksheets.Add(); $wsF45.Name = "F-045_사전안내가정통신문"; $ws = $wsF45
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 74쪽 [17-1] 대조. 학생·학부모 개인정보와 개별 신청값은 입력·저장·출력하지 않음"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,"초 6(예비 중1), 중 3(예비 고1) 교복 구매 사전 안내")'; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B6:H6").Merge() | Out-Null; $ws.Range("B6").Value2 = "안녕하십니까? 졸업과 함께 상급학교로 진학함을 축하드립니다."
    $ws.Range("B8:H8").Merge() | Out-Null; $ws.Range("B8").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 중학교, 고등학교 신입생 배정을 앞두고 "&IF(기초자료입력!C12<>"",기초자료입력!C12,"교복 구매")&"에 대한 사항을 다음과 같이 안내해 드립니다."'; $ws.Range("B8").WrapText = $true; $ws.Rows.Item(8).RowHeight = 30
    $ws.Range("B10:H11").Merge() | Out-Null; $ws.Range("B10").Value2 = "정부의「교복 가격 안정화 방안」(2013.7.9)에 따라, 학교는 ‘학교주관구매’를 실시하고 있으며, 전북 도내 교복을 착용하는 대부분의 중·고등학교는 ‘교복 학교주관구매 제도’를 실시하고 있습니다."; $ws.Range("B10").WrapText = $true; $ws.Rows.Item(10).RowHeight = 42
    $ws.Range("B13:H14").Merge() | Out-Null; $ws.Range("B13").Value2 = "교복은 ‘학교주관구매’의 권고 상한가인 동복 000원, 하복 000원 이하로 지원하게 됩니다. 학교별 자체 구매계획에 따라 업체를 선정하며, 학생들은 진학하는 중학교(고등학교)에서 교복을 현물로 지원 받습니다."; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 42
    $ws.Range("B16:H17").Merge() | Out-Null; $ws.Range("B16").Value2 = "기타 교복 착용 여부와 구매 일정, 착용 시기 등 세부사항은 신입생 배정 발표 이후에 해당 학교의 안내를 통해 반드시 확인하시고, 학교의 안내 없이 교복을 개별적으로 구매하여 착오가 생기는 일이 없도록 유의하시기 바랍니다."; $ws.Range("B16").WrapText = $true; $ws.Rows.Item(16).RowHeight = 42
    $ws.Range("B19:H19").Merge() | Out-Null; $ws.Range("B19").Value2 = "감사합니다."
    $ws.Range("B21:H21").Merge() | Out-Null; $ws.Range("B21").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy. m. d."),"년  월  일")'; $ws.Range("B21").HorizontalAlignment = -4108
    $ws.Range("B22:H22").Merge() | Out-Null; $ws.Range("B22").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4&"장","○○학교장")'; $ws.Range("B22").HorizontalAlignment = -4108; $ws.Range("B22").Font.Bold = $true
    $ws.Range("B3:H22").Font.Size = 10; $ws.Columns.Item("A").ColumnWidth = 2.5; for ($c = 2; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 10 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.2; $ps.BottomMargin = CmToPt 1.2; $ps.LeftMargin = CmToPt 1.5; $ps.RightMargin = CmToPt 1.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$22"
    $hbF45 = $ws.HPageBreaks.Count; $vbF45 = $ws.VPageBreaks.Count; L "F-045 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF45, VPageBreaks=$vbF45"

    # ---- 7w2. F-046 수요조사 가정통신문 안내 ----
    # 원본 HWPX 75쪽 [18] 대조. F-044와 같은 기안문 구조(수요조사 발송을 내부결재하는 문서).
    # 학교명·학년도·문서번호·발행일·구매명만 공통값으로 반영한다.
    $wsF46 = $wbNew.Worksheets.Add(); $wsF46.Name = "F-046_수요조사가정통신문안내"; $ws = $wsF46
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 75쪽 [18] 대조. 학교·학년도·구매명 공통값만 반영하며 학생·학부모 개인정보는 입력·저장·출력하지 않음"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("G2:H2").Merge() | Out-Null; $ws.Range("G2").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○ ○ 학 교")'; $ws.Range("G2").Font.Bold = $true; $ws.Range("G2").HorizontalAlignment = -4108
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Value2 = "신입생 예비소집일 교복구매 수요조사 가정통신문 발송"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108; $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재(사업부서)"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,"신입생 예비소집일 교복구매 수요조사 가정통신문 발송")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Value2 = "1. 관련: 전북특별자치도교육청 학교안전과-0000(20○○.) 학교주관공동구매 요령"
    $ws.Range("B13:H13").Merge() | Out-Null; $ws.Range("B13").Value2 = "2. 정확한 교복 물량 파악을 위한 학교주관구매 수요조사를 위해 붙임과 같이 가정통신문을 발송하고자 합니다."; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 30
    $ws.Range("B15").Value2 = "가. 대상:"; $ws.Range("C15:H15").Merge() | Out-Null; $ws.Range("C15").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 신입생 및 학부모"'
    $ws.Range("B16").Value2 = "나. 내용:"; $ws.Range("C16:H16").Merge() | Out-Null; $ws.Range("C16").Formula = '=IF(기초자료입력!C12<>"",기초자료입력!C12&" 관련 교복 학교주관구매 참여 여부 조사","교복 학교주관구매 참여 여부 조사")'
    $ws.Range("B18:H18").Merge() | Out-Null; $ws.Range("B18").Value2 = "붙임 교복구매 수요조사 가정통신문 1부. 끝."
    $ws.Range("F22").Value2 = "담당"; $ws.Range("G22").Value2 = "협조자"; $ws.Range("H22").Value2 = "교장"
    $ws.Range("B23").Value2 = "시행"; $ws.Range("C23:E23").Merge() | Out-Null; $ws.Range("C23").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F23").Value2 = "접수"; $ws.Range("G23:H23").Merge() | Out-Null
    $ws.Range("B24").Value2 = "우편번호"; $ws.Range("D24").Value2 = "주소"; $ws.Range("E24:H24").Merge() | Out-Null
    $ws.Range("B25").Value2 = "전화"; $ws.Range("D25").Value2 = "전송(팩스)"; $ws.Range("F25").Value2 = "이메일"; $ws.Range("G25:H25").Merge() | Out-Null
    $ws.Range("B5:H25").Font.Size = 9; $ws.Range("B5,E5,B7,B8,B9,B11,B15,B16,B18,F22:H22,B23,F23,B24,D24,B25,D25,F25").Font.Bold = $true
    $ws.Range("B5:H9,F22:H25").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 10; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$25"
    $hbF46 = $ws.HPageBreaks.Count; $vbF46 = $ws.VPageBreaks.Count
    L "F-046 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF46, VPageBreaks=$vbF46 (둘 다 0이면 A4 1쪽 자연 충족)"

    # ---- 7x. F-032 제안서 접수 결과 ----
    # 원본 HWPX 61쪽 [10] 대조. 업체 반복행(R-03)의 업체명만 반영한다. 접수일자와
    # 대표자는 제안서를 실제 접수한 뒤 담당자가 수기 작성하는 값이므로 자동 반영하지 않는다.
    $wsF32 = $wbNew.Worksheets.Add(); $wsF32.Name = "F-032_제안서접수결과"; $ws = $wsF32
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 61쪽 [10] 대조. 업체명만 자동 반영하며 접수일자·대표자는 수기 작성함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복구매 제안서 접수 결과"'; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 제안서 접수 결과 보고")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Formula = '="1. "&IF(기초자료입력!C23<>"",기초자료입력!C23,"관련 문서를 확인하십시오.")'
    $ws.Range("B13:H14").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 "&IF(기초자료입력!C12<>"",기초자료입력!C12,"교복 학교주관구매")&" 2단계 입찰(규격·가격 동시) 공고와 관련하여 제안서 접수 결과를 아래와 같이 보고합니다."'; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 32
    $ws.Range("B16:H16").Merge() | Out-Null; $ws.Range("B16").Value2 = "□ 제안서 접수 결과 □"; $ws.Range("B16").Font.Bold = $true
    $ws.Range("B17").Value2 = "접수번호"; $ws.Range("C17").Value2 = "접수일자"; $ws.Range("D17:E17").Merge() | Out-Null; $ws.Range("D17").Value2 = "업체현황"; $ws.Range("D18").Value2 = "업체명"; $ws.Range("E18").Value2 = "대표자"; $ws.Range("B17:B18").Merge() | Out-Null; $ws.Range("C17:C18").Merge() | Out-Null
    for ($i = 0; $i -lt 10; $i++) { $r = 19 + $i; $srcRow = 44 + $i; $ws.Range("B$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",$($i+1),`"`")"; $ws.Range("D$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",기초자료입력!B$srcRow,`"`")"; $ws.Range("C$r,E$r").Interior.Color = 16777164 }
    $ws.Range("B30:H30").Merge() | Out-Null; $ws.Range("B30").Formula = '="붙임  "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 제안서 접수대장 1부.  끝."'
    $ws.Range("F33").Value2 = "담당"; $ws.Range("G33").Value2 = "협조자"; $ws.Range("H33").Value2 = "교장"
    $ws.Range("B34").Value2 = "시행"; $ws.Range("C34:E34").Merge() | Out-Null; $ws.Range("C34").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F34").Value2 = "접수"; $ws.Range("G34:H34").Merge() | Out-Null
    $ws.Range("B35").Value2 = "우편번호"; $ws.Range("D35").Value2 = "주소"; $ws.Range("E35:H35").Merge() | Out-Null
    $ws.Range("B36").Value2 = "전화"; $ws.Range("D36").Value2 = "전송(팩스)"; $ws.Range("F36").Value2 = "이메일"; $ws.Range("G36:H36").Merge() | Out-Null
    $ws.Range("B5:H36").Font.Size = 9; $ws.Range("B5,E5,B7,B8,B9,B11,B16,B17:E18,F33:H33,B34,F34,B35,D35,B36,D36,F36").Font.Bold = $true; $ws.Range("B17:E28").Borders.LineStyle = 1; $ws.Range("B17:E18").Interior.Color = 15987699; $ws.Range("B17:E28").HorizontalAlignment = -4108; $ws.Range("B5:H9,F33:H36").Borders.LineStyle = 1
    # D·E열(업체현황 하위 업체명·대표자) 폭이 넓어 1.0cm 여백에서도 A4 폭을 초과함
    # (2026-09-19 검증에서 VPageBreaks=1로 확인). 열 폭을 줄이고 여백을 0.75cm로 좁힘.
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 11; $ws.Columns.Item("C").ColumnWidth = 14; $ws.Columns.Item("D").ColumnWidth = 13; $ws.Columns.Item("E").ColumnWidth = 10; $ws.Columns.Item("F").ColumnWidth = 8; $ws.Columns.Item("G").ColumnWidth = 8; $ws.Columns.Item("H").ColumnWidth = 8
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 0.75; $ps.RightMargin = CmToPt 0.75; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$36"
    $hbF32 = $ws.HPageBreaks.Count; $vbF32 = $ws.VPageBreaks.Count; L "F-032 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF32, VPageBreaks=$vbF32"

    # ---- 7w. F-033 제안서 접수대장 ----
    # 원본 HWPX 62쪽 [10-1] 대조. 업체명만 반복행에서 자동 반영하고 대표자·연락처·대리인·서명은 공란이다.
    $wsF33 = $wbNew.Worksheets.Add(); $wsF33.Name = "F-033_제안서접수대장"; $ws = $wsF33
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 62쪽 [10-1] 대조. 업체명 외 대표자·연락처·대리인·수령인은 수기 공란임"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:L3").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 제안서 접수대장"'; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B4:L4").Merge() | Out-Null; $ws.Range("B4").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")'; $ws.Range("B4").HorizontalAlignment = -4108
    $ws.Range("B6").Value2 = "접수" + [Environment]::NewLine + "일자"; $ws.Range("C6").Value2 = "접수" + [Environment]::NewLine + "번호"; $ws.Range("D6:F6").Merge() | Out-Null; $ws.Range("D6").Value2 = "회사현황"; $ws.Range("G6:I6").Merge() | Out-Null; $ws.Range("G6").Value2 = "접수인(대리인)현황"; $ws.Range("J6:K6").Merge() | Out-Null; $ws.Range("J6").Value2 = "샘플 반환"; $ws.Range("L6:L7").Merge() | Out-Null; $ws.Range("L6").Value2 = "비고"
    $headersF33 = @("업체명","대표자","전화" + [Environment]::NewLine + "번호","직위","접수인" + [Environment]::NewLine + "(대리인)","전화" + [Environment]::NewLine + "번호","일자","수령인")
    for ($i = 0; $i -lt $headersF33.Count; $i++) { $ws.Cells.Item(7, 4 + $i).Value2 = $headersF33[$i] }
    $ws.Range("B6:B7").Merge() | Out-Null; $ws.Range("C6:C7").Merge() | Out-Null
    # "B$r,E$r:L$r"의 "$r:L"은 PowerShell이 드라이브/스코프 한정자 구문으로 오인해 ":L$r"이
    # 통째로 사라지는 결함이 있었음(F-040 서약 문단에서 처음 발견·확정한 것과 동일한 유형).
    # ${r}로 변수명 경계를 명시해 회피함. 영향은 서식 하이라이트 범위 축소뿐이었음(데이터 없음).
    for ($i = 0; $i -lt 10; $i++) { $r = 8 + $i; $srcRow = 44 + $i; $ws.Range("C$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",$($i+1),`"`")"; $ws.Range("D$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",기초자료입력!B$srcRow,`"`")"; $ws.Range("B${r},E${r}:L${r}").Interior.Color = 16777164 }
    $ws.Range("B20:L20").Merge() | Out-Null; $ws.Range("B20").Value2 = "※ 샘플은 낙찰업체를 제외하고, 낙찰자 결정 이후 7일 이내 업체에서 수거"; $ws.Range("B20").Font.Size = 8
    $ws.Range("B21:L21").Merge() | Out-Null; $ws.Range("B21").Value2 = "(탈락업체에서 샘플 미수거 시 폐기처분 할 수 있음)"; $ws.Range("B21").Font.Size = 8
    $ws.Range("B6:L17").Borders.LineStyle = 1; $ws.Range("B6:L7").Interior.Color = 15987699; $ws.Range("B6:L7").Font.Bold = $true; $ws.Range("B6:L17").HorizontalAlignment = -4108; $ws.Range("B6:L17").VerticalAlignment = -4108; $ws.Range("B6:L17").WrapText = $true; $ws.Rows.Item(6).RowHeight = 30; $ws.Rows.Item(7).RowHeight = 30
    $ws.Columns.Item("A").ColumnWidth = 2.5; foreach ($col in @("B","C","E","F","G","H","I","J","K","L")) { $ws.Columns.Item($col).ColumnWidth = 8 }; $ws.Columns.Item("D").ColumnWidth = 16; $ws.Range("B3:L21").Font.Size = 9
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 2; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$L`$21"
    $hbF33 = $ws.HPageBreaks.Count; $vbF33 = $ws.VPageBreaks.Count; L "F-033 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF33, VPageBreaks=$vbF33"

    # ---- 7x. F-034 제안서 평가위원회 개최 ----
    # 원본 HWPX 63쪽 [11] 대조. 위원 역할·마스킹 식별표시만 반영하며 일시·장소·서명은 수기 공란이다.
    $wsF34 = $wbNew.Worksheets.Add(); $wsF34.Name = "F-034_평가위원회개최"; $ws = $wsF34
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 63쪽 [11] 대조. 위원 성명은 마스킹 식별표시만 반영하며 일시·장소는 수기 작성함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 제안서 평가위원회 개최"'; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"; $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 제안서 평가위원회 개최")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Formula = '="1. "&IF(기초자료입력!C23<>"",기초자료입력!C23,"관련 문서를 확인하십시오.")'
    $ws.Range("B13:H14").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 제안서 평가위원회를 아래와 같이 개최하고자 합니다."'; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 28
    $ws.Range("B16").Value2 = "가. 일시："; $ws.Range("C16:H16").Merge() | Out-Null; $ws.Range("C16").Interior.Color = 16777164
    $ws.Range("B17").Value2 = "나. 장소："; $ws.Range("C17:H17").Merge() | Out-Null; $ws.Range("C17").Interior.Color = 16777164
    $ws.Range("B18").Value2 = "다. 대상："; $ws.Range("C18:H18").Merge() | Out-Null; $ws.Range("C18").Formula = '=IF(COUNTA(기초자료입력!B58:B67)>0,"평가위원 "&COUNTA(기초자료입력!B58:B67)&"명","")'
    $headersF34 = @("순","구분","성명","비고"); $colsF34 = @("B","C","D","E")
    for ($i = 0; $i -lt $headersF34.Count; $i++) { $ws.Range("$($colsF34[$i])20").Value2 = $headersF34[$i] }
    for ($i = 0; $i -lt 10; $i++) { $r = 21 + $i; $srcRow = 58 + $i; $role = if ($i -eq 0) { "위원장" } elseif ($i -eq 1) { "부위원장" } else { "위원" }; $ws.Range("B$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",$($i+1),`"`")"; $ws.Range("C$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",`"$role`",`"`")"; $ws.Range("D$r").Formula = "=IF(기초자료입력!C$srcRow<>`"`",기초자료입력!C$srcRow,`"`")" }
    $ws.Range("B20:E30").Borders.LineStyle = 1; $ws.Range("B20:E20").Interior.Color = 15987699; $ws.Range("B20:E20").Font.Bold = $true; $ws.Range("B20:E30").HorizontalAlignment = -4108
    $ws.Range("B32:H32").Merge() | Out-Null; $ws.Range("B32").Value2 = "라. 안내방법：개별통지. 끝."
    $ws.Range("F35").Value2 = "담당"; $ws.Range("G35").Value2 = "협조자"; $ws.Range("H35").Value2 = "교장"; $ws.Range("B36").Value2 = "시행"; $ws.Range("C36:E36").Merge() | Out-Null; $ws.Range("C36").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F36").Value2 = "접수"; $ws.Range("G36:H36").Merge() | Out-Null
    $ws.Range("B37").Value2 = "우편번호"; $ws.Range("D37").Value2 = "주소"; $ws.Range("E37:H37").Merge() | Out-Null; $ws.Range("B38").Value2 = "전화"; $ws.Range("D38").Value2 = "전송(팩스)"; $ws.Range("F38").Value2 = "이메일"; $ws.Range("G38:H38").Merge() | Out-Null
    $ws.Range("B5:H38").Font.Size = 9; $ws.Range("B5,E5,B7,B8,B9,B11,B16:B18,B20:E20,F35:H35,B36,F36,B37,D37,B38,D38,F38").Font.Bold = $true; $ws.Range("B5:H9,F35:H38").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 9; $ws.Columns.Item("C").ColumnWidth = 14; $ws.Columns.Item("D").ColumnWidth = 16; $ws.Columns.Item("E").ColumnWidth = 12; $ws.Columns.Item("F").ColumnWidth = 8; $ws.Columns.Item("G").ColumnWidth = 8; $ws.Columns.Item("H").ColumnWidth = 8
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$38"
    $hbF34 = $ws.HPageBreaks.Count; $vbF34 = $ws.VPageBreaks.Count; L "F-034 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF34, VPageBreaks=$vbF34"

    # ---- 7z. F-035 제안서 정량평가 결과 ----
    # 원본 HWPX 64쪽 [12] 대조. 업체별 정량평가 점수는 F-015가 이미 쓰는 학교 정량평가
    # 입력열(기초자료입력!C44:F53)을 재조회하는 집계 보고서이며, 업체명(R-03)만 반영하고
    # 접수일자·서명 등은 없다(원문 자체에 개인 서명란이 없음).
    $wsF35 = $wbNew.Worksheets.Add(); $wsF35.Name = "F-035_정량평가결과"; $ws = $wsF35
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 64쪽 [12] 대조. 업체별 점수는 F-015 학교 정량평가 입력값(기초자료입력 C:F열)을 재조회함"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 구매 제안서 정량평가 결과"'; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 학교주관구매 제안서 정량평가 결과 보고")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Formula = '="1. "&IF(기초자료입력!C23<>"",기초자료입력!C23,"관련 문서를 확인하십시오.")'
    $ws.Range("B13:H14").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 "&IF(기초자료입력!C12<>"",기초자료입력!C12,"교복 학교주관구매")&" 2단계 입찰(규격가격동시) 공고와 관련하여 접수된 제안서에 대한 정량평가를 실시하고, 그 결과를 아래와 같이 보고합니다."'; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 32
    $ws.Range("B16:H16").Merge() | Out-Null; $ws.Range("B16").Value2 = "□ 정량평가 □"; $ws.Range("B16").Font.Bold = $true
    $ws.Range("B17:B18").Merge() | Out-Null; $ws.Range("B17").Value2 = "순"
    $ws.Range("C17:C18").Merge() | Out-Null; $ws.Range("C17").Value2 = "업체명"
    $ws.Range("D17:H17").Merge() | Out-Null; $ws.Range("D17").Value2 = "평가점수"
    $headersF35 = @("수행경험","품질인증","접근성","상한가격","합계"); $colsF35 = @("D","E","F","G","H")
    for ($i = 0; $i -lt $headersF35.Count; $i++) { $ws.Range("$($colsF35[$i])18").Value2 = $headersF35[$i] }
    for ($i = 0; $i -lt 10; $i++) {
        $r = 19 + $i; $srcRow = 44 + $i
        $ws.Range("B$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",$($i+1),`"`")"
        $ws.Range("C$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",기초자료입력!B$srcRow,`"`")"
        # 소스 점수 셀(C:F열)이 진짜로 빈 셀일 때 IF(B<>"",소스셀,"")로 값을 그대로 반환하면
        # 빈 셀 참조가 0으로 강제되어(엑셀 사양) 점수 미입력 상태를 0점으로 잘못 표시하는
        # 결함이 있었음(2026-09-19 발견). D~G열 각각 소스 셀 자체의 공백 여부를 직접 확인해
        # 진짜 공백일 때만 공란을 반환하도록 수정함(F-036이 이미 쓰는 안전한 패턴과 동일).
        $ws.Range("D$r").Formula = "=IF(OR(기초자료입력!B$srcRow=`"`",기초자료입력!C$srcRow=`"`"),`"`",기초자료입력!C$srcRow)"
        $ws.Range("E$r").Formula = "=IF(OR(기초자료입력!B$srcRow=`"`",기초자료입력!D$srcRow=`"`"),`"`",기초자료입력!D$srcRow)"
        $ws.Range("F$r").Formula = "=IF(OR(기초자료입력!B$srcRow=`"`",기초자료입력!E$srcRow=`"`"),`"`",기초자료입력!E$srcRow)"
        $ws.Range("G$r").Formula = "=IF(OR(기초자료입력!B$srcRow=`"`",기초자료입력!F$srcRow=`"`"),`"`",기초자료입력!F$srcRow)"
        $ws.Range("H$r").Formula = "=IF(기초자료입력!B$srcRow=`"`",`"`",IF(OR(D$r=`"`",E$r=`"`",F$r=`"`",G$r=`"`"),`"`",D$r+E$r+F$r+G$r))"
    }
    $ws.Range("B30:H30").Merge() | Out-Null; $ws.Range("B30").Value2 = "붙임  1. 정량 평가표 각 1부.  2. 업체별 실적, 품질 인증, 상한 가격 자료(별첨) 각 1부.  끝."; $ws.Range("B30").WrapText = $true
    $ws.Range("F33").Value2 = "담당"; $ws.Range("G33").Value2 = "협조자"; $ws.Range("H33").Value2 = "교장"
    $ws.Range("B34").Value2 = "시행"; $ws.Range("C34:E34").Merge() | Out-Null; $ws.Range("C34").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F34").Value2 = "접수"; $ws.Range("G34:H34").Merge() | Out-Null
    $ws.Range("B35").Value2 = "우편번호"; $ws.Range("D35").Value2 = "주소"; $ws.Range("E35:H35").Merge() | Out-Null
    $ws.Range("B36").Value2 = "전화"; $ws.Range("D36").Value2 = "전송(팩스)"; $ws.Range("F36").Value2 = "이메일"; $ws.Range("G36:H36").Merge() | Out-Null
    $ws.Range("B5:H36").Font.Size = 9; $ws.Range("B5,E5,B7,B8,B9,B11,B16,B17:H18,F33:H33,B34,F34,B35,D35,B36,D36,F36").Font.Bold = $true; $ws.Range("B17:H28").Borders.LineStyle = 1; $ws.Range("B17:H18").Interior.Color = 15987699; $ws.Range("B17:H28").HorizontalAlignment = -4108; $ws.Range("B5:H9,F33:H36").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 6; $ws.Columns.Item("C").ColumnWidth = 16; foreach ($col in @("D","E","F","G","H")) { $ws.Columns.Item($col).ColumnWidth = 10 }
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$36"
    $hbF35 = $ws.HPageBreaks.Count; $vbF35 = $ws.VPageBreaks.Count; L "F-035 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF35, VPageBreaks=$vbF35"

    # ---- 7za. F-036 제안서 평가 결과 ----
    # 원본 HWPX 65쪽 [13] 대조. 정량평가 합계(F-015/F-035와 동일한 C:F열 합)와 정성평가
    # 합계(F-016과 동일한 H:L열 합)를 더해 적격 여부(80점 이상)를 계산한다. 원문 표 헤더
    # 문구가 "성 명"이지만 실제 셀 값은 "적격/부적격" 문자열이라 헤더-값 불일치가 있음을
    # 확인함(F-004~006·F-023·F-025·F-041과 같은 유형의 매핑 불일치, 매핑표는 수정하지
    # 않고 여기 기록만 함). 값의 의미에 맞춰 헤더를 "적부 판정"으로 표시한다.
    $wsF36 = $wbNew.Worksheets.Add(); $wsF36.Name = "F-036_제안서평가결과"; $ws = $wsF36
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 65쪽 [13] 대조. 원문 표 헤더가 `"성 명`"이나 실제 값은 적격/부적격 판정이라 `"적부 판정`"으로 표시함(매핑표 미수정, 기록만)"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 구매 제안서 평가 결과"'; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "문서번호"; $ws.Range("C5").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'
    $ws.Range("E5").Value2 = "시행일자"; $ws.Range("F5").Formula = '=IF(기초자료입력!C8<>"",TEXT(기초자료입력!C8,"yyyy-mm-dd"),"")'
    $ws.Range("C7:H7").Merge() | Out-Null; $ws.Range("B7").Value2 = "수  신"; $ws.Range("C7").Value2 = "내부결재"
    $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("B8").Value2 = "(경유)"
    $ws.Range("C9:H9").Merge() | Out-Null; $ws.Range("B9").Value2 = "제  목"; $ws.Range("C9").Formula = '=IF(기초자료입력!C22<>"",기초자료입력!C22,IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복 업체 선정을 위한 제안서 평가 집계결과 보고")'; $ws.Range("C9").Font.Bold = $true
    $ws.Range("B11:H11").Merge() | Out-Null; $ws.Range("B11").Formula = '="1. 관련："&IF(기초자료입력!C23<>"",기초자료입력!C23,"관련 문서를 확인하십시오.")'
    $ws.Range("B13:H14").Merge() | Out-Null; $ws.Range("B13").Formula = '="2. "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복업체 선정을 위한 제안서 평가위원회 실시 결과를 아래와 같이 보고합니다."'; $ws.Range("B13").WrapText = $true; $ws.Rows.Item(13).RowHeight = 28
    $ws.Range("B17:B18").Merge() | Out-Null; $ws.Range("B17").Value2 = "업체명"
    $ws.Range("C17:E17").Merge() | Out-Null; $ws.Range("C17").Value2 = "규격 평가(제안서 평가)"
    $ws.Range("C18").Value2 = "정량평가"; $ws.Range("D18").Value2 = "정성평가"; $ws.Range("E18").Value2 = "합계"
    $ws.Range("F17:F18").Merge() | Out-Null; $ws.Range("F17").Value2 = "적부 판정"
    $ws.Range("G17:G18").Merge() | Out-Null; $ws.Range("G17").Value2 = "비 고"
    for ($i = 0; $i -lt 10; $i++) {
        $r = 19 + $i; $srcRow = 44 + $i
        $ws.Range("B$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",기초자료입력!B$srcRow,`"`")"
        $ws.Range("C$r").Formula = "=IF(기초자료입력!B$srcRow=`"`",`"`",IF(OR(기초자료입력!C$srcRow=`"`",기초자료입력!D$srcRow=`"`",기초자료입력!E$srcRow=`"`",기초자료입력!F$srcRow=`"`"),`"`",기초자료입력!C$srcRow+기초자료입력!D$srcRow+기초자료입력!E$srcRow+기초자료입력!F$srcRow))"
        $ws.Range("D$r").Formula = "=IF(기초자료입력!B$srcRow=`"`",`"`",IF(OR(기초자료입력!H$srcRow=`"`",기초자료입력!I$srcRow=`"`",기초자료입력!J$srcRow=`"`",기초자료입력!K$srcRow=`"`",기초자료입력!L$srcRow=`"`"),`"`",기초자료입력!H$srcRow+기초자료입력!I$srcRow+기초자료입력!J$srcRow+기초자료입력!K$srcRow+기초자료입력!L$srcRow))"
        $ws.Range("E$r").Formula = "=IF(OR(C$r=`"`",D$r=`"`"),`"`",C$r+D$r)"
        $ws.Range("F$r").Formula = "=IF(E$r=`"`",`"`",IF(E$r>=80,`"적격`",`"부적격`"))"
        $ws.Range("G$r").Formula = "=IF(F$r=`"적격`",`"가격개찰대상`",`"`")"
    }
    $ws.Range("B29:G29").Merge() | Out-Null; $ws.Range("B29").Value2 = "※ 적격 여부 판정 점수：합계 80점 이상"; $ws.Range("B29").Font.Size = 8
    $ws.Range("B30:G34").Merge() | Out-Null; $ws.Range("B30").Value2 = "붙임  1. 제안서 평가위원 등록부 1부.  2. 제안서 설명 업체 참가자 등록부 1부.  3. 제안서 평가의결서 및 집계표 각 1부.  4. 업체별 평가표 00부.  5. 평가위원 청렴 및 보안각서 각 1부.  끝."; $ws.Range("B30").WrapText = $true
    $ws.Range("F37").Value2 = "담당"; $ws.Range("G37").Value2 = "교장"
    $ws.Range("B38").Value2 = "시행"; $ws.Range("C38:E38").Merge() | Out-Null; $ws.Range("C38").Formula = '=IF(기초자료입력!C9<>"",기초자료입력!C9,"")'; $ws.Range("F38").Value2 = "접수"; $ws.Range("G38").Value2 = ""
    $ws.Range("B5:G38").Font.Size = 9; $ws.Range("B5,E5,B7,B8,B9,B11,B17:G18,F37:G37,B38,F38").Font.Bold = $true; $ws.Range("B17:G28").Borders.LineStyle = 1; $ws.Range("B17:G18").Interior.Color = 15987699; $ws.Range("B17:G28").HorizontalAlignment = -4108; $ws.Range("B5:G9").Borders.LineStyle = 1
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 16; foreach ($col in @("C","D","E","F")) { $ws.Columns.Item($col).ColumnWidth = 10 }; $ws.Columns.Item("G").ColumnWidth = 14
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$G`$38"
    $hbF36 = $ws.HPageBreaks.Count; $vbF36 = $ws.VPageBreaks.Count; L "F-036 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF36, VPageBreaks=$vbF36"

    # ---- 7y. F-037 평가위원회 참석 등록부 ----
    # 원본 HWPX 66쪽 [13-1] 대조. 위원 역할(R-01)만 표시하고 성명·서명(R-02 및 자필란)은
    # 빈 양식으로 유지한다. C-01/C-02/D-02는 공통 문서 정보로만 반영한다.
    $wsF37 = $wbNew.Worksheets.Add(); $wsF37.Name = "F-037_평가위원회참석등록부"; $ws = $wsF37
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 66쪽 [13-1] 대조. 참석자 성명·서명은 자필 작성 공란임"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H3").Merge() | Out-Null; $ws.Range("B3").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5&"학년도 ","")&"교복선정 평가위원회 참석 등록부"'; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5:H5").Merge() | Out-Null; $ws.Range("B5").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"")&IF(기초자료입력!C22<>""," · "&기초자료입력!C22,"")'; $ws.Range("B5").HorizontalAlignment = -4108
    $ws.Range("B7").Value2 = "1. 일시："; $ws.Range("C7:H7").Merge() | Out-Null
    $ws.Range("B8").Value2 = "2. 장소："; $ws.Range("C8:H8").Merge() | Out-Null; $ws.Range("C8").Value2 = "회의실"
    $ws.Range("B10:H10").Merge() | Out-Null; $ws.Range("B10").Value2 = "3. 참석자 현황"; $ws.Range("B10").Font.Bold = $true
    $headersF37 = @("순","직위","성명","서명","비고"); $colsF37 = @("B","C","D","E","F")
    for ($i = 0; $i -lt $headersF37.Count; $i++) { $ws.Range("$($colsF37[$i])11").Value2 = $headersF37[$i] }
    for ($i = 0; $i -lt 10; $i++) {
        $r = 12 + $i; $srcRow = 58 + $i
        $ws.Range("B$r").Value2 = [double]($i + 1)
        $ws.Range("C$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",기초자료입력!B$srcRow,`"`")"
    }
    $ws.Range("B11:F21").Borders.LineStyle = 1; $ws.Range("B11:F11").Interior.Color = 15987699; $ws.Range("B11:F11").Font.Bold = $true; $ws.Range("B11:F21").HorizontalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 7; $ws.Columns.Item("C").ColumnWidth = 16; $ws.Columns.Item("D").ColumnWidth = 16; $ws.Columns.Item("E").ColumnWidth = 15; $ws.Columns.Item("F").ColumnWidth = 15; $ws.Columns.Item("G").ColumnWidth = 2; $ws.Columns.Item("H").ColumnWidth = 2
    $ws.Range("B3:H21").Font.Size = 10
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$21"
    $hbF37 = $ws.HPageBreaks.Count; $vbF37 = $ws.VPageBreaks.Count; L "F-037 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF37, VPageBreaks=$vbF37"

    # ---- 7w. F-038 제안서 설명 업체 참가 등록부 ----
    # 원본 HWPX 67쪽 [13-2] 대조. 업체명(R-03)만 반영하고 대표자·전화·대리인·서명 등은 빈칸이다.
    $wsF38 = $wbNew.Worksheets.Add(); $wsF38.Name = "F-038_업체참가등록부"; $ws = $wsF38
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 67쪽 [13-2] 대조. 업체명 외 개인정보·연락처·대리인 정보는 수기 공란임"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:I3").Merge() | Out-Null; $ws.Range("B3").Value2 = "제안서 설명 업체 참가 등록부"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5:I5").Merge() | Out-Null; $ws.Range("B5").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"")&IF(기초자료입력!C5<>""," ("&기초자료입력!C5&"학년도)","")&IF(기초자료입력!C22<>""," · "&기초자료입력!C22,"")'; $ws.Range("B5").HorizontalAlignment = -4108
    $ws.Range("B7:D7").Merge() | Out-Null; $ws.Range("B7").Value2 = "회사"; $ws.Range("E7:G7").Merge() | Out-Null; $ws.Range("E7").Value2 = "제안자"; $ws.Range("H7:H8").Merge() | Out-Null; $ws.Range("H7").Value2 = "제안심사" + [Environment]::NewLine + "관리번호"; $ws.Range("I7:I8").Merge() | Out-Null; $ws.Range("I7").Value2 = "비고"
    $headersF38 = @("업체명","대표자","전화번호","직위","접수인" + [Environment]::NewLine + "(대리인)","전화번호")
    for ($i = 0; $i -lt $headersF38.Count; $i++) { $ws.Cells.Item(8, 2 + $i).Value2 = $headersF38[$i] }
    for ($i = 0; $i -lt 10; $i++) { $r = 9 + $i; $srcRow = 44 + $i; $ws.Range("B$r").Formula = "=IF(기초자료입력!B$srcRow<>`"`",기초자료입력!B$srcRow,`"`")" }
    $ws.Range("B7:I18").Borders.LineStyle = 1; $ws.Range("B7:I8").Interior.Color = 15987699; $ws.Range("B7:I8").Font.Bold = $true; $ws.Range("B7:I18").HorizontalAlignment = -4108; $ws.Range("B7:I18").VerticalAlignment = -4108; $ws.Range("B7:I18").WrapText = $true
    $ws.Range("B20:I20").Merge() | Out-Null; $ws.Range("B20").Value2 = "※ 제안 심사 관리 번호"; $ws.Range("B20").Font.Bold = $true
    $ws.Range("B21:I21").Merge() | Out-Null; $ws.Range("B21").Value2 = "○ 제안서 블라인드 평가를 위한 관리번호"
    $ws.Range("B22:I22").Merge() | Out-Null; $ws.Range("B22").Value2 = "○ 접수 완료 후 제안평가 심사 실시(30분 이상 확보) 전 추첨하여 관리번호 부여"
    $ws.Range("B23:I23").Merge() | Out-Null; $ws.Range("B23").Value2 = "- 관리번호는 교복 샘플과 평가위원용 제안서에 표기하여 관리한다."
    $ws.Columns.Item("A").ColumnWidth = 2.5; foreach ($col in @("B","C","D","E","F","G","H","I")) { $ws.Columns.Item($col).ColumnWidth = 8 }
    $ws.Rows.Item(7).RowHeight = 28; $ws.Rows.Item(8).RowHeight = 30; $ws.Range("B3:I23").Font.Size = 9
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.0; $ps.BottomMargin = CmToPt 1.0; $ps.LeftMargin = CmToPt 1.0; $ps.RightMargin = CmToPt 1.0; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$I`$23"
    $hbF38 = $ws.HPageBreaks.Count; $vbF38 = $ws.VPageBreaks.Count; L "F-038 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF38, VPageBreaks=$vbF38"

    # ---- 7wa. F-039 업체별 제안서 평가표 ----
    # 원본 HWPX 68쪽 [13-3] 대조. F-041~F-043과 같은 I3 단일 업체 선택 패턴을 재사용한다
    # (업체 반복행 전체를 순회하지 않음). 위원 8명의 원점수는 이 문서 전용 입력칸
    # (E8:H15)이며 다른 서식과 공유하지 않는다. 위원명(D열)은 청탁방지·익명성 원칙에 따라
    # 자동 반영하지 않고 자필 공란으로 둔다. 총계·평균은 원문 안내 "위원별 평가 점수 중
    # 최고점 및 최저점을 제외"에 따라 8명 점수 중 최댓값·최솟값을 제외하고 계산한다.
    $wsF39 = $wbNew.Worksheets.Add(); $wsF39.Name = "F-039_업체별제안서평가표"; $ws = $wsF39
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 68쪽 [13-3] 대조. 위원명은 자필 공란이며 옷감·완성도·A/S·하자 점수만 입력함(이 문서 전용 입력칸)"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("K2").Value2 = "선택 업체 순번(자동, 인쇄 전용)"; $ws.Range("K2").Font.Size = 7
    $ws.Range("I3").Value2 = 1
    $ws.Range("B3:H4").Merge() | Out-Null; $ws.Range("B3").Formula = '="[13-3] "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 교복(동복·하복) 구매 업체별 제안서 평가표"'; $ws.Range("B3").Font.Size = 14; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5").Value2 = "학교명"; $ws.Range("C5").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"")'
    $ws.Range("D5").Value2 = "학년도"; $ws.Range("E5").Formula = '=IF(기초자료입력!C5<>"",기초자료입력!C5,"")'
    $ws.Range("F5").Value2 = "평가 대상 업체"; $ws.Range("G5:J5").Merge() | Out-Null; $ws.Range("G5").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'
    $headersF39 = @("번호","업체명","위원명","옷감의재질`n(15점)","교복의완성도`n(10점)","A/S`n(15점)","하자`n(10점)","점수합계`n(50점)","비고"); $colsF39 = @("B","C","D","E","F","G","H","I","J")
    for ($i = 0; $i -lt $headersF39.Count; $i++) { $ws.Range("$($colsF39[$i])7").Value2 = $headersF39[$i] }
    $ws.Range("B8:B15").Merge() | Out-Null; $ws.Range("B8").Formula = "=I3"
    $ws.Range("C8:C15").Merge() | Out-Null; $ws.Range("C8").Formula = '=IFERROR(INDEX(기초자료입력!$B$44:$B$53,I3),"")'
    $ws.Range("J8:J15").Merge() | Out-Null; $ws.Range("J8").Value2 = "위원별 평가 점수 중`n최고점 및 최저점을 제외"; $ws.Range("J8").WrapText = $true
    for ($i = 0; $i -lt 8; $i++) {
        $r = 8 + $i
        $ws.Range("I$r").Formula = "=IF(COUNT(E`$r:H`$r)=4,SUM(E`$r:H`$r),`"`")" -replace '\$r', $r
        $ws.Range("E$r,F$r,G$r,H$r").Interior.Color = 16777164
    }
    $ws.Range("B16:D16").Merge() | Out-Null; $ws.Range("B16").Value2 = "총계"
    $ws.Range("B17:D17").Merge() | Out-Null; $ws.Range("B17").Value2 = "평균"
    foreach ($col in @("E","F","G","H","I")) {
        $ws.Range("$($col)16").Formula = "=IF(COUNT($col`8:$col`15)=8,SUM($col`8:$col`15)-MAX($col`8:$col`15)-MIN($col`8:$col`15),`"`")"
        $ws.Range("$($col)17").Formula = "=IF($($col)16=`"`",`"`",$($col)16/6)"
        $ws.Range("$($col)17").NumberFormat = "0.0"
    }
    $ws.Range("B19:J19").Merge() | Out-Null; $ws.Range("B19").Value2 = "※ 학교의 제안서 평가항목에 맞게 변경 가능"; $ws.Range("B19").Font.Size = 8
    $ws.Range("B7:J17").Borders.LineStyle = 1; $ws.Range("B7:J7").Interior.Color = 15987699; $ws.Range("B7:J7").Font.Bold = $true; $ws.Range("B16:D17").Font.Bold = $true
    $ws.Range("B7:J17").HorizontalAlignment = -4108; $ws.Range("B7:J17").VerticalAlignment = -4108; $ws.Range("B7:H7").WrapText = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 6; $ws.Columns.Item("C").ColumnWidth = 14; $ws.Columns.Item("D").ColumnWidth = 12; foreach ($col in @("E","F","G","H","I")) { $ws.Columns.Item($col).ColumnWidth = 9 }; $ws.Columns.Item("J").ColumnWidth = 14
    $ws.Rows.Item(7).RowHeight = 28; for ($r = 8; $r -le 15; $r++) { $ws.Rows.Item($r).RowHeight = 18 }
    $ws.Range("B3:J19").Font.Size = 9
    # 9개 열(번호~비고) 표는 세로 방향 1.27cm 여백에서 폭이 A4를 초과함(2026-09-19 검증에서
    # VPageBreaks=1로 확인). F-033(접수대장, 12개 열)과 같은 이유로 가로 방향으로 전환함.
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 2; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$J`$19"
    $hbF39 = $ws.HPageBreaks.Count; $vbF39 = $ws.VPageBreaks.Count; L "F-039 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF39, VPageBreaks=$vbF39"

    # ---- 7x. F-040 위원 청렴 및 보안 서약서 ----
    # 원본 HWPX 69쪽 [13-4] 대조. 서약자 성명·서명은 자동 반영하지 않는 단일 빈 양식이다.
    $wsF40 = $wbNew.Worksheets.Add(); $wsF40.Name = "F-040_위원청렴보안서약서"; $ws = $wsF40
    $ws.Range("A1").Value2 = "[검토중 — 담당자 최종 확인 후 사용] 원본 HWPX 69쪽 [13-4] 대조. 서약자 성명·서명은 자필 작성 공란임"
    $ws.Range("A1").Font.Size = 8; $ws.Range("A1").Font.Color = 255
    $ws.Range("B3:H3").Merge() | Out-Null; $ws.Range("B3").Value2 = "교복선정위원회 위원 청렴 및 보안 서약서"; $ws.Range("B3").Font.Size = 15; $ws.Range("B3").Font.Bold = $true; $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B5:H5").Merge() | Out-Null; $ws.Range("B5").Formula = '=IF(기초자료입력!C22<>"","평가 안건: "&기초자료입력!C22,"")'; $ws.Range("B5").HorizontalAlignment = -4108
    $ws.Range("B7:H8").Merge() | Out-Null; $ws.Range("B7").Formula = '="본인은 년 월 일 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&"에서 실시하는 "&IF(기초자료입력!C5<>"",기초자료입력!C5&"학년도 ","")&"교복 학교주관구매업체 선정을 위한 평가함에 있어 「부패 없는 투명한 사회」 구현 등을 위하여 다음 사항을 준수할 것을 서약합니다."'; $ws.Range("B7").WrapText = $true
    # 2026-09-19 발견(F-039 QA 육안 확인 중): 1번 항목은 "1. " 리터럴과 "="로 시작하는 수식을
    # 그냥 문자열 결합해 만든 탓에 결과 문자열이 "1. =IF(...)"로 "="가 아닌 문자로 시작해
    # .Formula가 이를 수식이 아닌 있는 그대로의 텍스트로 저장해버려 수식 원문이 그대로
    # 인쇄되는 결함이 있었음. 2~4번 항목도 같은 이유로 문자열 앞뒤에 불필요한 큰따옴표가
    # 그대로 남아 인쇄되는 결함이 있었음. 1번만 실제 수식으로, 2~4번은 순수 텍스트 값으로
    # 수정함(고정 텍스트라 수식이 필요 없음).
    $pledgesF40 = @(
        '="1. "&IF(기초자료입력!C5<>"",기초자료입력!C5,"20○○")&"학년도 "&IF(기초자료입력!C4<>"",기초자료입력!C4,"○○학교")&" 교복 학교주관구매업체 선정을 위해 제시된 항목에 따라 객관적이고 공정하게 심사할 것을 약속합니다."',
        "2. ○○학교 교복 학교 주관 교복선정 위원회 지위를 이용하여 관련업체로부터 어떠한 경우에도 금품·향응·편의 등을 수수하거나 제공받지 않을 것이며, 이러한 상황이 발생하면 사업부서에 통보하여 공정한 평가가 이루어지도록 하겠습니다.",
        "3. 업무상 취득한 비밀을 준수하고 보안관계 규정 및 지침을 성실히 수행하겠습니다.",
        "4. 평가와 관련하여 알게 된 업무상 비밀을 타인에게 누설하지 않겠으며, 업무상 취득한 비밀을 누설할 때에는 관계법규에 따라 처벌을 받는 것에 이의를 제기하지 않겠습니다."
    )
    # 2026-09-19 발견: "B$r:H$($r+1)" 형태 문자열은 PowerShell이 "$r:"를 드라이브/스코프
    # 한정자 구문으로 잘못 해석해(예: $env:PATH 같은 패턴과 혼동) $r 값이 통째로 사라지고
    # "B11"처럼 깨진 주소가 만들어지는 결함이 있었음. 실제로 이 병합이 전부 실패해(단일 셀
    # 취급) 서약 문단이 B열 폭(13유닛)만으로 줄바꿈되어 육안 확인 시 심하게 잘려 보였음.
    # ${r}처럼 변수명 경계를 명시해 회피함(F-039 작성 시 이미 다른 곳에서 적용한 원칙과 동일).
    $pledgesF40IsFormula = @($true, $false, $false, $false)
    for ($i = 0; $i -lt $pledgesF40.Count; $i++) {
        $r = 10 + ($i * 2)
        $ws.Range("B${r}:H$($r+1)").Merge() | Out-Null
        if ($pledgesF40IsFormula[$i]) { $ws.Range("B$r").Formula = $pledgesF40[$i] } else { $ws.Range("B$r").Value2 = $pledgesF40[$i] }
        $ws.Range("B$r").WrapText = $true
        # 긴 문단(2·4번)이 병합 셀에서 잘리지 않도록 여유 있게 높이를 늘림(Rows.AutoFit()은
        # 병합+줄바꿈 셀에서 정확히 계산되지 않는 기존에 확립된 한계 — F-002/F-004에서 이미 확인).
        $ws.Rows.Item($r).RowHeight = if ($i -eq 1) { 60 } elseif ($i -eq 3) { 50 } else { 34 }
    }
    $ws.Range("B19:H19").Merge() | Out-Null; $ws.Range("B19").Value2 = "20 년 월 일"; $ws.Range("B19").HorizontalAlignment = -4108
    $ws.Range("B21:H21").Merge() | Out-Null; $ws.Range("B21").Formula = '=IF(기초자료입력!B58<>"","위원 구분: "&기초자료입력!B58,"")'; $ws.Range("B21").HorizontalAlignment = -4108
    $ws.Range("B23:H23").Merge() | Out-Null; $ws.Range("B23").Value2 = "서 약 자  성 명 :                         (인)"; $ws.Range("B23").HorizontalAlignment = -4108
    $ws.Range("B25:H25").Merge() | Out-Null; $ws.Range("B25").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○○학교")&"장 귀하"'; $ws.Range("B25").HorizontalAlignment = -4108
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 13; for ($c = 3; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 9 }
    $ws.Range("B3:H25").Font.Size = 10
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.HeaderMargin = CmToPt 0.5; $ps.FooterMargin = CmToPt 0.5; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$H`$25"
    $hbF40 = $ws.HPageBreaks.Count; $vbF40 = $ws.VPageBreaks.Count; L "F-040 시트: 자연 배율(Zoom=100) 기준 HPageBreaks=$hbF40, VPageBreaks=$vbF40"

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

} catch {
    $buildFailure = $_
    L "ERROR: $($_.Exception.Message)"
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
if ($buildFailure) { throw $buildFailure }
