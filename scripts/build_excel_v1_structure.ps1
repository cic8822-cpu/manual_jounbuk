# P3-01 클린룸 재구현 — 구조 빌더(1/2): 시트·필드·서식선택 UI·A4 페이지 설정
# 원본 파일은 읽기 전용으로만 열며(학교정보 표본 복사), 수정하지 않음.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $root 'artifacts\excel'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$outPath = Join-Path $outDir '교복구매_길라잡이_Excel_v1.xlsm'
$sourcePath = Join-Path $root '20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm'
$logPath = Join-Path $root '_workspace\03_excel\build_structure_log.txt'

function CmToPt($cm) { return [double]$cm * 28.3465 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$log = New-Object System.Text.StringBuilder
function L($s) { [void]$log.AppendLine($s) }

$wbNew = $null
$wbSrc = $null
try {
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
    $ws.Range("A4").Value2 = "2. '기초자료입력' 시트에 공통·사업·문서 정보와 품목을 입력한 뒤 [저장하기] 버튼을 누르면 'DB' 시트에 기록됩니다."
    $ws.Range("A5").Value2 = "3. '서식선택_출력' 시트에서 원하는 서식에 체크(TRUE)한 뒤 [선택 서식 인쇄 미리보기]/[선택 서식 PDF 저장] 버튼을 누릅니다."
    $ws.Range("A6").Value2 = "4. 이 v1 버전은 F-007, F-024 두 서식만 완전히 구현되어 있습니다. 나머지 서식은 '구현상태' 열에 표시된 대로 순차 추가 예정입니다."
    $ws.Range("A7").Value2 = "5. F-013(교복 디자인 및 규격서)은 표·이미지가 많은 다쪽(11쪽) 문서로, Excel보다 HWPX 경로가 적합하여 이번 버전에서는 보류하고 document-automation-engineer 협업 대상으로 남겼습니다."
    $ws.Range("A9").Value2 = "원본 보호: 이 파일은 20230808_용역계약갈라잡이(디깅모멘텀)_이행원.xlsm 을 참고해 클린룸 방식으로 새로 작성한 사본이며, 원본 파일을 직접 열거나 수정하지 않습니다."
    $ws.Range("A10").Value2 = "개인정보 경계: 이 파일은 학교·계약 단위 업무 정보만 다루며, 학생·학부모 개인정보 및 서명·직인 자동처리는 포함하지 않습니다."
    $ws.Columns.Item("A").ColumnWidth = 110
    $ws.Range("A3:A10").WrapText = $false

    L "사용설명서 시트 작성 완료"

    # ---- 3. 기초자료입력 ----
    $wsBase = $wbNew.Worksheets.Add()
    $wsBase.Name = "기초자료입력"
    $ws = $wsBase
    $ws.Range("A1").Value2 = "기초자료 입력 (공통 C-01~C-06 / 사업 B-01~B-07 / 문서별 D-01~D-05 / 품목 R-04~R-06)"
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
        @{r=14; id="B-03"; label="계약방식 (목록 확정 전 — 자유 입력)"},
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

    $ws.Columns.Item("A").ColumnWidth = 6
    $ws.Columns.Item("B").ColumnWidth = 20
    $ws.Columns.Item("C").ColumnWidth = 20
    $ws.Columns.Item("D").ColumnWidth = 14
    $ws.Columns.Item("E").ColumnWidth = 16
    $ws.Columns.Item("F").ColumnWidth = 16
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

    # ---- 5. 서식선택_출력 ----
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
    $implemented = @{ "F-007" = "F-007_구매요청기안문"; "F-024" = "F-024_단가비율표" }
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
    for ($i = 0; $i -lt 6; $i++) { $r = 9 + $i; $src = 29 + $i; $ws.Range("C$r").Value2 = $itemsF24[$i]; $ws.Range("D$r").Formula = "=IF(기초자료입력!C$src<>`"`",기초자료입력!C$src,1)"; $ws.Range("E$r").Formula = "=IF(SUM(기초자료입력!`$D`$29:`$D`$34)=0,`"`",TEXT(기초자료입력!D$src/SUM(기초자료입력!`$D`$29:`$D`$34)*100,`"0.0`")&`"%`")" }
    $ws.Range("B15:C15").Merge() | Out-Null; $ws.Range("B15").Value2 = "합계"; $ws.Range("D15").Formula = "=SUM(D9:D14)"; $ws.Range("E15").Formula = '="100%"'; $ws.Range("F15:G15").Merge() | Out-Null
    $ws.Range("B17:G17").Merge() | Out-Null; $ws.Range("B17").Value2 = "20○○년    월    일"; $ws.Range("B18:G18").Merge() | Out-Null; $ws.Range("B18").Value2 = "제안자 성  명                    (서명 또는 날인)"; $ws.Range("B19:G19").Merge() | Out-Null; $ws.Range("B19").Formula = '=IF(기초자료입력!C4<>"",기초자료입력!C4,"○○○○학교")&"장 귀하"'
    $ws.Range("B6:G6,B8:G15").Borders.LineStyle = 1; $ws.Range("B8:G8").Font.Bold = $true; $ws.Range("B8:G8").HorizontalAlignment = -4108; $ws.Range("B15:G15").Font.Bold = $true
    $ws.Columns.Item("A").ColumnWidth = 2.5; $ws.Columns.Item("B").ColumnWidth = 9; $ws.Columns.Item("C").ColumnWidth = 16; $ws.Columns.Item("D").ColumnWidth = 9; $ws.Columns.Item("E").ColumnWidth = 14; $ws.Columns.Item("F").ColumnWidth = 17; $ws.Columns.Item("G").ColumnWidth = 10; $ws.Range("B3:G19").Font.Size = 10
    $ps = $ws.PageSetup; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.TopMargin = CmToPt 1.27; $ps.BottomMargin = CmToPt 1.27; $ps.LeftMargin = CmToPt 1.27; $ps.RightMargin = CmToPt 1.27; $ps.Zoom = 100; $ps.FitToPagesWide = $false; $ps.FitToPagesTall = $false; $ps.PrintArea = "`$A`$1:`$G`$19"
    L "F-007/F-024 v2 원문 대조 레이아웃 및 수식 재적용 완료"

    # ---- 8. 학교정보 (원본 공개 데이터 표본 복사 — 학생·학부모 개인정보 아님) ----
    $wbSrc = $excel.Workbooks.Open($sourcePath, [Type]::Missing, $true)
    $wsSrcSchool = $wbSrc.Worksheets.Item("학교정보")
    $wsSchool = $wbNew.Worksheets.Add()
    $wsSchool.Name = "학교정보"
    $ws = $wsSchool
    $sampleRows = 120  # 표본(전체 1,516행 중 일부) — 전체가 필요하면 후속 작업에서 확장
    $srcRange = $wsSrcSchool.Range("A1:G$($sampleRows + 1)")
    $destRange = $ws.Range("A1:G$($sampleRows + 1)")
    $srcRange.Copy($destRange)
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.Columns.Item("A").ColumnWidth = 8
    $ws.Columns.Item("B").ColumnWidth = 10
    $ws.Columns.Item("C").ColumnWidth = 8
    $ws.Columns.Item("D").ColumnWidth = 8
    $ws.Columns.Item("E").ColumnWidth = 24
    $ws.Columns.Item("F").ColumnWidth = 34
    $ws.Columns.Item("G").ColumnWidth = 14
    L "학교정보 시트 작성 완료 (원본 $sampleRows 행 표본 복사, 전체 1516행 중 일부 — 공개 기관정보, 개인정보 아님)"

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
}

[System.IO.File]::WriteAllText($logPath, $log.ToString(), [System.Text.UTF8Encoding]::new($false))
