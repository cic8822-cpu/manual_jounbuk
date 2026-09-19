# P3-01 클린룸 재구현 — VBA 빌더(2/2): 기초자료 신규/저장/수정/불러오기, 서식선택 출력(PDF) 매크로 추가
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$defaultTargetPath = Join-Path $root 'artifacts\excel\build.xlsm'
$targetPath = if ([string]::IsNullOrWhiteSpace($env:UNIFORM_EXCEL_BUILD_PATH)) { $defaultTargetPath } else { $env:UNIFORM_EXCEL_BUILD_PATH }
$logPath = Join-Path $root '_workspace\03_excel\build_vba_log.txt'

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
$log = New-Object System.Text.StringBuilder
function L($s) { [void]$log.AppendLine($s) }

$wb = $null
$saveSucceeded = $false
try {
    $wb = $excel.Workbooks.Open($targetPath, [Type]::Missing, $false)  # ReadOnly=False (편집)

    # ---- 레코드 추적 셀 추가 (기초자료입력!G1/H1) ----
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    $wsIn.Range("G1").Value2 = "불러온 레코드 순번"
    $wsIn.Range("G1").Font.Size = 8
    $wsIn.Range("H1").Value2 = ""
    L "기초자료입력!G1/H1 레코드 추적 셀 추가 완료"

    # 반복행 저장 시트는 구조 빌더가 중단된 경우에도 VBA 빌더 재실행만으로 복구할 수 있게 보장함.
    try {
        $wsItems = $wb.Worksheets.Item("DB_품목")
    } catch {
        $wsItems = $wb.Worksheets.Add()
        $wsItems.Name = "DB_품목"
        $itemHeaders = @("레코드순번", "행번호", "품목명", "수량", "단가", "금액")
        for ($i = 0; $i -lt $itemHeaders.Count; $i++) {
            $wsItems.Cells.Item(1, $i + 1).Value2 = $itemHeaders[$i]
            $wsItems.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsItems.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_품목 시트 복구 생성 완료"
    }
    try {
        $wsVendors = $wb.Worksheets.Item("DB_업체")
    } catch {
        $wsVendors = $wb.Worksheets.Add()
        $wsVendors.Name = "DB_업체"
        $vendorHeaders = @("레코드순번", "행번호", "업체명")
        for ($i = 0; $i -lt $vendorHeaders.Count; $i++) {
            $wsVendors.Cells.Item(1, $i + 1).Value2 = $vendorHeaders[$i]
            $wsVendors.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsVendors.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_업체 시트 복구 생성 완료"
    }
    try {
        $wsCommittee = $wb.Worksheets.Item("DB_위원")
    } catch {
        $wsCommittee = $wb.Worksheets.Add()
        $wsCommittee.Name = "DB_위원"
        $committeeHeaders = @("레코드순번", "행번호", "역할직위", "마스킹식별표시")
        for ($i = 0; $i -lt $committeeHeaders.Count; $i++) {
            $wsCommittee.Cells.Item(1, $i + 1).Value2 = $committeeHeaders[$i]
            $wsCommittee.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsCommittee.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_위원 시트 복구 생성 완료"
    }
    try {
        $wsScore = $wb.Worksheets.Item("DB_평가")
    } catch {
        $wsScore = $wb.Worksheets.Add()
        $wsScore.Name = "DB_평가"
        $scoreHeaders = @("레코드순번", "행번호", "평가항목", "배점", "점수")
        for ($i = 0; $i -lt $scoreHeaders.Count; $i++) {
            $wsScore.Cells.Item(1, $i + 1).Value2 = $scoreHeaders[$i]
            $wsScore.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsScore.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_평가 시트 복구 생성 완료"
    }
    try {
        $blankSheet = $wb.Worksheets.Item("Sheet2")
        if ($blankSheet.UsedRange.CountLarge -eq 1 -and [string]::IsNullOrWhiteSpace([string]$blankSheet.Range("A1").Value2)) {
            $excel.DisplayAlerts = $false
            $blankSheet.Delete()
            L "빈 기본 시트 Sheet2 제거 완료"
        }
    } catch { L "제거할 빈 기본 시트 Sheet2 없음" }

    # ---- VBA 모듈 추가 (재실행 대비: 동일 이름 기존 모듈 제거 후 추가) ----
    $vbproj = $wb.VBProject
    foreach ($nm in @("Module_기초자료","Module_출력","Module_검색_절차")) {
        try {
            $existing = $vbproj.VBComponents.Item($nm)
            $vbproj.VBComponents.Remove($existing)
            L "기존 모듈 $nm 제거 후 재생성"
        } catch { L "기존 모듈 $nm 없음" }
    }

    $modInput = $vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    $modInput.Name = "Module_기초자료"
    $codeInput = @'
Option Explicit

Sub 초기화()
    Call 초기화_실행(True)
End Sub

Public Sub 검증_초기화()
    Call 초기화_실행(False)
End Sub

Private Sub 초기화_실행(ByVal showMessage As Boolean)
    With Sheets("기초자료입력")
        .Range("C4").Value = ""
        .Range("C5").Value = ""
        .Range("C6").Value = ""
        .Range("C7").Value = ""
        .Range("C8").Value = ""
        .Range("C9").Value = ""
        .Range("C12").Value = ""
        .Range("C13").Value = ""
        .Range("C14").Value = ""
        .Range("C15").Value = ""
        .Range("C16").Value = ""
        .Range("C17").Value = ""
        .Range("C18").Value = ""
        .Range("C21").Value = ""
        .Range("C22").Value = ""
        .Range("C23").Value = ""
        .Range("C24").Value = ""
        .Range("C25").Value = ""
        .Range("B29:D38").ClearContents
        .Range("B44:B53").ClearContents
        .Range("C44:F53").ClearContents
        .Range("H44:L53").ClearContents
        .Range("N44:Q53").ClearContents
        .Range("B58:C67").ClearContents
        .Range("B72:D81").ClearContents
        .Range("H1").Value = ""
    End With
    If showMessage Then MsgBox "기초자료가 초기화되었습니다.", vbInformation
End Sub

Function 필수값검증() As Boolean
    Dim ws As Worksheet
    Set ws = Sheets("기초자료입력")
    필수값검증 = True
    If Trim(ws.Range("C4").Value & "") = "" Then
        MsgBox "학교명(C-01)을 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
    If Trim(ws.Range("C5").Value & "") = "" Then
        MsgBox "학년도(C-02)를 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
    If Trim(ws.Range("C12").Value & "") = "" Then
        MsgBox "구매명(B-01)을 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
    If Trim(ws.Range("C22").Value & "") = "" Then
        MsgBox "제목(D-02)을 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
End Function

Public Function 검증_필수값검증() As Boolean
    Dim ws As Worksheet
    Set ws = Sheets("기초자료입력")
    검증_필수값검증 = Trim(ws.Range("C4").Value & "") <> "" And _
        Trim(ws.Range("C5").Value & "") <> "" And _
        Trim(ws.Range("C12").Value & "") <> "" And _
        Trim(ws.Range("C22").Value & "") <> ""
End Function

Private Sub 필드복사_기초자료_DB(wsIn As Worksheet, wsDB As Worksheet, r As Long)
    wsDB.Cells(r, 2).Value = wsIn.Range("C4").Value
    wsDB.Cells(r, 3).Value = wsIn.Range("C5").Value
    wsDB.Cells(r, 4).Value = wsIn.Range("C6").Value
    wsDB.Cells(r, 5).Value = wsIn.Range("C7").Value
    wsDB.Cells(r, 6).Value = wsIn.Range("C8").Value
    wsDB.Cells(r, 7).Value = wsIn.Range("C9").Value
    wsDB.Cells(r, 8).Value = wsIn.Range("C12").Value
    wsDB.Cells(r, 9).Value = wsIn.Range("C13").Value
    wsDB.Cells(r, 10).Value = wsIn.Range("C14").Value
    wsDB.Cells(r, 11).Value = wsIn.Range("C15").Value
    wsDB.Cells(r, 12).Value = wsIn.Range("C16").Value
    wsDB.Cells(r, 13).Value = wsIn.Range("C17").Value
    wsDB.Cells(r, 14).Value = wsIn.Range("C18").Value
    wsDB.Cells(r, 15).Value = wsIn.Range("C21").Value
    wsDB.Cells(r, 16).Value = wsIn.Range("C22").Value
    wsDB.Cells(r, 17).Value = wsIn.Range("C23").Value
    wsDB.Cells(r, 18).Value = wsIn.Range("C24").Value
    wsDB.Cells(r, 19).Value = wsIn.Range("C25").Value
    wsDB.Cells(r, 20).Value = Now
End Sub

Private Function 품목행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long
    Dim hasName As Boolean, hasQuantity As Boolean, hasPrice As Boolean
    품목행검증 = True
    For sourceRow = 29 To 38
        hasName = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> ""
        hasQuantity = Trim(wsIn.Cells(sourceRow, 3).Value & "") <> ""
        hasPrice = Trim(wsIn.Cells(sourceRow, 4).Value & "") <> ""
        If hasName Or hasQuantity Or hasPrice Then
            If Not hasName Or Not hasQuantity Or Not hasPrice Or _
                Not IsNumeric(wsIn.Cells(sourceRow, 3).Value) Or Not IsNumeric(wsIn.Cells(sourceRow, 4).Value) Or _
                CDbl(wsIn.Cells(sourceRow, 3).Value) <= 0 Or CDbl(wsIn.Cells(sourceRow, 4).Value) <= 0 Then
                If showMessage Then MsgBox "품목 " & (sourceRow - 28) & "행은 품목명·수량·단가를 모두 입력하고, 수량과 단가는 0보다 큰 숫자로 입력하세요.", vbExclamation
                품목행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_품목행검증() As Boolean
    검증_품목행검증 = 품목행검증(Sheets("기초자료입력"), False)
End Function

Private Function 숫자개수(ByVal textValue As String) As Long
    Dim i As Long, count As Long
    For i = 1 To Len(textValue)
        If Mid(textValue, i, 1) Like "#" Then count = count + 1
    Next i
    숫자개수 = count
End Function

Private Function 업체행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, vendorName As String
    업체행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" Then
            If Len(vendorName) < 2 Or Len(vendorName) > 150 Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행은 2~150자로 입력하세요.", vbExclamation
                업체행검증 = False
                Exit Function
            End If
            If InStr(vendorName, "대표자") > 0 Or InStr(vendorName, "대표") > 0 Or _
               InStr(vendorName, "연락처") > 0 Or InStr(vendorName, "전화") > 0 Or _
               InStr(vendorName, "휴대폰") > 0 Or InStr(vendorName, "사업자번호") > 0 Or _
               숫자개수(vendorName) >= 8 Or _
               vendorName Like "*[0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*0[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*01[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Then
                If showMessage Then MsgBox "업체명만 입력하세요. 대표자·연락처·전화번호·사업자번호는 입력·저장하지 않습니다.", vbExclamation
                업체행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_업체행검증() As Boolean
    검증_업체행검증 = 업체행검증(Sheets("기초자료입력"), False)
End Function

Public Function 검증_업체중복경고() As Boolean
    Dim wsIn As Worksheet, sourceRow As Long, compareRow As Long
    Dim vendorName As String, compareName As String
    Set wsIn = Sheets("기초자료입력")
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" Then
            For compareRow = sourceRow + 1 To 53
                compareName = Trim(wsIn.Cells(compareRow, 2).Value & "")
                If compareName <> "" And StrComp(vendorName, compareName, vbTextCompare) = 0 Then
                    검증_업체중복경고 = True
                    Exit Function
                End If
            Next compareRow
        End If
    Next sourceRow
    검증_업체중복경고 = False
End Function

Private Function 정량평가행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    ' 업체명만 먼저 등록하고 정량평가 점수는 평가 이후 입력하는 것이 정상 흐름이므로,
    ' 점수 4칸이 모두 비어 있으면 통과시키고, 하나라도 있으면 4칸 모두·범위까지 완전해야 저장을 허용한다.
    ' (F-015 출력 가능 여부는 별도의 F015업체행완전한가에서 더 엄격하게 판단한다.)
    Dim sourceRow As Long, vendorName As String
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    Dim hasAnyScore As Boolean, hasAllScore As Boolean
    정량평가행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        v1 = wsIn.Cells(sourceRow, 3).Value
        v2 = wsIn.Cells(sourceRow, 4).Value
        v3 = wsIn.Cells(sourceRow, 5).Value
        v4 = wsIn.Cells(sourceRow, 6).Value
        hasAnyScore = Trim(v1 & "") <> "" Or Trim(v2 & "") <> "" Or Trim(v3 & "") <> "" Or Trim(v4 & "") <> ""
        hasAllScore = Trim(v1 & "") <> "" And Trim(v2 & "") <> "" And Trim(v3 & "") <> "" And Trim(v4 & "") <> ""
        If hasAnyScore And vendorName = "" Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-015)은 업체명 없이 정량평가 점수만 입력할 수 없습니다.", vbExclamation
            정량평가행검증 = False
            Exit Function
        ElseIf hasAnyScore And Not hasAllScore Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-015)은 수행경험·공인인증·거리적접근성·상한가격 점수를 모두 입력하거나 모두 비워 두세요.", vbExclamation
            정량평가행검증 = False
            Exit Function
        ElseIf hasAllScore Then
            If Not 정수범위(v1, 0, 10) Or Not 정수범위(v2, 0, 10) Or Not 정수범위(v3, 0, 15) Or Not 정수범위(v4, 0, 15) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-015)은 수행경험(0~10)·공인인증(0~10)·거리적접근성(0~15)·상한가격(0~15) 범위의 정수만 입력하세요.", vbExclamation
                정량평가행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_정량평가행검증() As Boolean
    검증_정량평가행검증 = 정량평가행검증(Sheets("기초자료입력"), False)
End Function

Private Function 정성평가행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, vendorName As String
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant, adjustment As Variant
    Dim hasAnyScore As Boolean, hasAllScore As Boolean
    정성평가행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        v1 = wsIn.Cells(sourceRow, 8).Value
        v2 = wsIn.Cells(sourceRow, 9).Value
        v3 = wsIn.Cells(sourceRow, 10).Value
        v4 = wsIn.Cells(sourceRow, 11).Value
        adjustment = wsIn.Cells(sourceRow, 12).Value
        hasAnyScore = Trim(v1 & "") <> "" Or Trim(v2 & "") <> "" Or Trim(v3 & "") <> "" Or Trim(v4 & "") <> "" Or Trim(adjustment & "") <> ""
        hasAllScore = Trim(v1 & "") <> "" And Trim(v2 & "") <> "" And Trim(v3 & "") <> "" And Trim(v4 & "") <> "" And Trim(adjustment & "") <> ""
        If hasAnyScore And vendorName = "" Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-016)은 업체명 없이 정성평가 점수만 입력할 수 없습니다.", vbExclamation
            정성평가행검증 = False
            Exit Function
        ElseIf hasAnyScore And Not hasAllScore Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-016)은 재질·완성도·A/S·하자보상·가감점을 모두 입력하거나 모두 비워 두세요.", vbExclamation
            정성평가행검증 = False
            Exit Function
        ElseIf hasAllScore Then
            If Not 정수범위(v1, 0, 15) Or Not 정수범위(v2, 0, 10) Or Not 정수범위(v3, 0, 15) Or Not 정수범위(v4, 0, 10) Or Not 정수범위(adjustment, -15, 5) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-016)은 재질(0~15)·완성도(0~10)·A/S(0~15)·하자보상(0~10)·가감점(-15~5) 범위의 정수만 입력하세요.", vbExclamation
                정성평가행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_정성평가행검증() As Boolean
    검증_정성평가행검증 = 정성평가행검증(Sheets("기초자료입력"), False)
End Function

Private Function 자기평점행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, vendorName As String, v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    Dim hasAnyScore As Boolean, hasAllScore As Boolean
    자기평점행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        v1 = wsIn.Cells(sourceRow, 14).Value: v2 = wsIn.Cells(sourceRow, 15).Value: v3 = wsIn.Cells(sourceRow, 16).Value: v4 = wsIn.Cells(sourceRow, 17).Value
        hasAnyScore = Trim(v1 & "") <> "" Or Trim(v2 & "") <> "" Or Trim(v3 & "") <> "" Or Trim(v4 & "") <> ""
        hasAllScore = Trim(v1 & "") <> "" And Trim(v2 & "") <> "" And Trim(v3 & "") <> "" And Trim(v4 & "") <> ""
        If hasAnyScore And vendorName = "" Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-018)은 업체명 없이 자기평점만 입력할 수 없습니다.", vbExclamation
            자기평점행검증 = False: Exit Function
        ElseIf hasAnyScore And Not hasAllScore Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-018)은 수행경험·공인인증·거리적접근성·상한가격 자기평점을 모두 입력하거나 모두 비워 두세요.", vbExclamation
            자기평점행검증 = False: Exit Function
        ElseIf hasAllScore Then
            If Not 정수범위(v1, 0, 10) Or Not 정수범위(v2, 0, 10) Or Not 정수범위(v3, 0, 15) Or Not 정수범위(v4, 0, 15) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-018)은 수행경험(0~10)·공인인증(0~10)·거리적접근성(0~15)·상한가격(0~15) 범위의 정수만 입력하세요.", vbExclamation
                자기평점행검증 = False: Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_자기평점행검증() As Boolean
    검증_자기평점행검증 = 자기평점행검증(Sheets("기초자료입력"), False)
End Function

Private Function 위원행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, roleName As String, maskedLabel As String
    Dim hasRole As Boolean, hasMask As Boolean
    위원행검증 = True
    For sourceRow = 58 To 67
        roleName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        maskedLabel = Trim(wsIn.Cells(sourceRow, 3).Value & "")
        hasRole = roleName <> ""
        hasMask = maskedLabel <> ""
        If hasRole Or hasMask Then
            If Not hasRole Or Not hasMask Or Len(roleName) < 2 Or Len(roleName) > 50 Or Len(maskedLabel) < 2 Or Len(maskedLabel) > 50 Then
                If showMessage Then MsgBox "위원 " & (sourceRow - 57) & "행은 역할·직위와 마스킹 식별표시를 각각 2~50자로 입력하세요.", vbExclamation
                위원행검증 = False
                Exit Function
            End If
            If Not 마스킹식별표시허용(maskedLabel) Then
                If showMessage Then MsgBox "위원 식별표시는 실제 성명 대신 '위원 ○○', '위원 **', '평가위원 ○○', '평가위원 **' 형식만 사용하세요. 연락처·서명은 입력·저장하지 않습니다.", vbExclamation
                위원행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_위원행검증() As Boolean
    검증_위원행검증 = 위원행검증(Sheets("기초자료입력"), False)
End Function

Public Function 검증_위원중복경고() As Boolean
    Dim wsIn As Worksheet, sourceRow As Long, compareRow As Long
    Dim roleName As String, compareRole As String
    Set wsIn = Sheets("기초자료입력")
    For sourceRow = 58 To 67
        roleName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If roleName <> "" Then
            For compareRow = sourceRow + 1 To 67
                compareRole = Trim(wsIn.Cells(compareRow, 2).Value & "")
                If compareRole <> "" And StrComp(roleName, compareRole, vbTextCompare) = 0 Then
                    검증_위원중복경고 = True
                    Exit Function
                End If
            Next compareRow
        End If
    Next sourceRow
    검증_위원중복경고 = False
End Function

Private Function 평가행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, itemName As String
    Dim hasItem As Boolean, hasAllocation As Boolean, hasScore As Boolean
    평가행검증 = True
    For sourceRow = 72 To 81
        itemName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        hasItem = itemName <> ""
        hasAllocation = Trim(wsIn.Cells(sourceRow, 3).Value & "") <> ""
        hasScore = Trim(wsIn.Cells(sourceRow, 4).Value & "") <> ""
        If hasItem Or hasAllocation Or hasScore Then
            If Not hasItem Or Not hasAllocation Or Not hasScore Or Len(itemName) < 2 Or Len(itemName) > 100 Or _
                Not IsNumeric(wsIn.Cells(sourceRow, 3).Value) Or Not IsNumeric(wsIn.Cells(sourceRow, 4).Value) Or _
                CDbl(wsIn.Cells(sourceRow, 3).Value) < 0 Or CDbl(wsIn.Cells(sourceRow, 4).Value) < 0 Or _
                CDbl(wsIn.Cells(sourceRow, 4).Value) > CDbl(wsIn.Cells(sourceRow, 3).Value) Then
                If showMessage Then MsgBox "평가 " & (sourceRow - 71) & "행은 평가항목(2~100자), 0 이상 배점, 0 이상이며 배점 이하인 점수를 모두 입력하세요.", vbExclamation
                평가행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Private Function 마스킹식별표시허용(ByVal maskedLabel As String) As Boolean
    마스킹식별표시허용 = maskedLabel = "위원 ○○" Or maskedLabel = "위원 **" Or _
        maskedLabel = "위원 ○*" Or maskedLabel = "위원 *○" Or _
        maskedLabel = "평가위원 ○○" Or maskedLabel = "평가위원 **" Or _
        maskedLabel = "평가위원 ○*" Or maskedLabel = "평가위원 *○"
End Function

Public Function 정수범위(ByVal value As Variant, ByVal minimum As Long, ByVal maximum As Long) As Boolean
    If Not IsNumeric(value) Then Exit Function
    If CDbl(value) <> Fix(CDbl(value)) Then Exit Function
    정수범위 = CDbl(value) >= minimum And CDbl(value) <= maximum
End Function

Private Function 양수숫자(ByVal value As Variant) As Boolean
    양수숫자 = IsNumeric(value)
    If 양수숫자 Then 양수숫자 = CDbl(value) > 0
End Function

Private Function 음이아닌숫자(ByVal value As Variant) As Boolean
    음이아닌숫자 = IsNumeric(value)
    If 음이아닌숫자 Then 음이아닌숫자 = CDbl(value) >= 0
End Function

Private Function 품목금액일치(ByVal quantity As Variant, ByVal unitPrice As Variant, ByVal amount As Variant) As Boolean
    If Not 양수숫자(quantity) Or Not 양수숫자(unitPrice) Or Not IsNumeric(amount) Then Exit Function
    품목금액일치 = CDbl(amount) = CDbl(quantity) * CDbl(unitPrice)
End Function

Private Function 평가점수범위(ByVal allocation As Variant, ByVal score As Variant) As Boolean
    If Not 음이아닌숫자(allocation) Or Not 음이아닌숫자(score) Then Exit Function
    평가점수범위 = CDbl(score) <= CDbl(allocation)
End Function

Private Function 마지막사용행(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        마지막사용행 = 1
    Else
        마지막사용행 = lastCell.Row
    End If
End Function

Private Function DB레코드존재(ByVal wsDB As Worksheet, ByVal seq As Long) As Boolean
    Dim rowIndex As Long
    For rowIndex = 2 To 마지막사용행(wsDB)
        If wsDB.Cells(rowIndex, 1).Value = seq Then
            DB레코드존재 = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function 저장경계검증(ByRef reason As String) As Boolean
    Dim wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Dim rowIndex As Long, lastRow As Long, seq As Long
    Dim vendorName As String, roleName As String, maskedLabel As String, itemName As String
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")

    lastRow = 마지막사용행(wsDB)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsDB.Range("A" & rowIndex & ":T" & rowIndex)) > 0 Then
            If Not 정수범위(wsDB.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Trim(wsDB.Cells(rowIndex, 2).Value & "") = "" Or Trim(wsDB.Cells(rowIndex, 3).Value & "") = "" Or _
               Trim(wsDB.Cells(rowIndex, 8).Value & "") = "" Or Trim(wsDB.Cells(rowIndex, 16).Value & "") = "" Then
                reason = "DB " & rowIndex & "행은 순번과 학교명·학년도·구매명·제목을 모두 갖춘 업무 레코드여야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsItems)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsItems.Range("A" & rowIndex & ":F" & rowIndex)) > 0 Then
            If Not 정수범위(wsItems.Cells(rowIndex, 1).Value, 1, 2147483647) Then
                reason = "DB_품목 " & rowIndex & "행의 레코드순번이 올바르지 않습니다."
                Exit Function
            End If
            seq = CLng(wsItems.Cells(rowIndex, 1).Value)
            If Not DB레코드존재(wsDB, seq) Or Not 정수범위(wsItems.Cells(rowIndex, 2).Value, 1, 10) Or _
               Trim(wsItems.Cells(rowIndex, 3).Value & "") = "" Or _
               Not 품목금액일치(wsItems.Cells(rowIndex, 4).Value, wsItems.Cells(rowIndex, 5).Value, wsItems.Cells(rowIndex, 6).Value) Then
                reason = "DB_품목 " & rowIndex & "행은 연결된 레코드와 완전한 품목·수량·단가·금액을 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsVendors)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsVendors.Range("A" & rowIndex & ":C" & rowIndex)) > 0 Then
            vendorName = Trim(wsVendors.Cells(rowIndex, 3).Value & "")
            If Not 정수범위(wsVendors.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsVendors.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsVendors.Cells(rowIndex, 2).Value, 1, 10) Or Len(vendorName) < 2 Or Len(vendorName) > 150 Or _
               InStr(vendorName, "대표자") > 0 Or InStr(vendorName, "대표") > 0 Or InStr(vendorName, "연락처") > 0 Or _
               InStr(vendorName, "전화") > 0 Or InStr(vendorName, "휴대폰") > 0 Or InStr(vendorName, "사업자번호") > 0 Or _
               숫자개수(vendorName) >= 8 Or vendorName Like "*[0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*0[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*01[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Then
                reason = "DB_업체 " & rowIndex & "행에는 업체명만 허용되며 대표자·연락처·전화번호·사업자번호는 저장할 수 없습니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsCommittee)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsCommittee.Range("A" & rowIndex & ":D" & rowIndex)) > 0 Then
            roleName = Trim(wsCommittee.Cells(rowIndex, 3).Value & "")
            maskedLabel = Trim(wsCommittee.Cells(rowIndex, 4).Value & "")
            If Not 정수범위(wsCommittee.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsCommittee.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsCommittee.Cells(rowIndex, 2).Value, 1, 10) Or Len(roleName) < 2 Or Len(roleName) > 50 Or _
               Len(maskedLabel) < 2 Or Len(maskedLabel) > 50 Or Not 마스킹식별표시허용(maskedLabel) Then
                reason = "DB_위원 " & rowIndex & "행은 역할·직위와 허용된 마스킹 식별표시를 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsScore)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsScore.Range("A" & rowIndex & ":E" & rowIndex)) > 0 Then
            itemName = Trim(wsScore.Cells(rowIndex, 3).Value & "")
            If Not 정수범위(wsScore.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsScore.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsScore.Cells(rowIndex, 2).Value, 1, 10) Or Len(itemName) < 2 Or Len(itemName) > 100 Or _
               Not 평가점수범위(wsScore.Cells(rowIndex, 4).Value, wsScore.Cells(rowIndex, 5).Value) Then
                reason = "DB_평가 " & rowIndex & "행은 평가항목·배점·점수가 완전하고 점수가 배점 이하이어야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsQuant)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsQuant.Range("A" & rowIndex & ":F" & rowIndex)) > 0 Then
            If Not 정수범위(wsQuant.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsQuant.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsQuant.Cells(rowIndex, 2).Value, 1, 10) Or _
               Not 정수범위(wsQuant.Cells(rowIndex, 3).Value, 0, 10) Or Not 정수범위(wsQuant.Cells(rowIndex, 4).Value, 0, 10) Or _
               Not 정수범위(wsQuant.Cells(rowIndex, 5).Value, 0, 15) Or Not 정수범위(wsQuant.Cells(rowIndex, 6).Value, 0, 15) Then
                reason = "DB_정량평가 " & rowIndex & "행은 연결된 레코드와 정상 범위의 수행경험·공인인증·거리적접근성·상한가격 점수를 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex
    lastRow = 마지막사용행(wsQual)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsQual.Range("A" & rowIndex & ":G" & rowIndex)) > 0 Then
            If Not 정수범위(wsQual.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsQual.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsQual.Cells(rowIndex, 2).Value, 1, 10) Or _
               Not 정수범위(wsQual.Cells(rowIndex, 3).Value, 0, 15) Or Not 정수범위(wsQual.Cells(rowIndex, 4).Value, 0, 10) Or _
               Not 정수범위(wsQual.Cells(rowIndex, 5).Value, 0, 15) Or Not 정수범위(wsQual.Cells(rowIndex, 6).Value, 0, 10) Or _
               Not 정수범위(wsQual.Cells(rowIndex, 7).Value, -15, 5) Then
                reason = "DB_정성평가 " & rowIndex & "행은 연결된 레코드와 정상 범위의 재질·완성도·A/S·하자보상·가감점을 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex
    lastRow = 마지막사용행(wsSelf)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsSelf.Range("A" & rowIndex & ":F" & rowIndex)) > 0 Then
            If Not 정수범위(wsSelf.Cells(rowIndex, 1).Value, 1, 2147483647) Or Not DB레코드존재(wsDB, CLng(Val(wsSelf.Cells(rowIndex, 1).Value))) Or Not 정수범위(wsSelf.Cells(rowIndex, 2).Value, 1, 10) Or Not 정수범위(wsSelf.Cells(rowIndex, 3).Value, 0, 10) Or Not 정수범위(wsSelf.Cells(rowIndex, 4).Value, 0, 10) Or Not 정수범위(wsSelf.Cells(rowIndex, 5).Value, 0, 15) Or Not 정수범위(wsSelf.Cells(rowIndex, 6).Value, 0, 15) Then
                reason = "DB_자기평점 " & rowIndex & "행은 연결된 레코드와 정상 범위의 업체 자기평점을 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex
    저장경계검증 = True
End Function

Public Function 검증_저장경계() As Boolean
    Dim reason As String
    검증_저장경계 = 저장경계검증(reason)
End Function

Public Function 검증_저장경계_메시지() As String
    Dim reason As String
    If Not 저장경계검증(reason) Then 검증_저장경계_메시지 = reason
End Function

Public Function 검증_평가행검증() As Boolean
    검증_평가행검증 = 평가행검증(Sheets("기초자료입력"), False)
End Function

Private Sub 품목복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsItems As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsItems.Cells(rowIndex, 1).Value = seq Then wsItems.Rows(rowIndex).Delete
    Next rowIndex

    targetRow = wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 29 To 38
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" Then
            wsItems.Cells(targetRow, 1).Value = seq
            wsItems.Cells(targetRow, 2).Value = sourceRow - 28
            wsItems.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 2).Value
            wsItems.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 3).Value
            wsItems.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 4).Value
            wsItems.Cells(targetRow, 6).Value = wsIn.Cells(sourceRow, 5).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 업체복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsVendors As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long, vendorName As String
    For rowIndex = wsVendors.Cells(wsVendors.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsVendors.Cells(rowIndex, 1).Value = seq Then wsVendors.Rows(rowIndex).Delete
    Next rowIndex

    targetRow = wsVendors.Cells(wsVendors.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" Then
            wsVendors.Cells(targetRow, 1).Value = seq
            wsVendors.Cells(targetRow, 2).Value = sourceRow - 43
            wsVendors.Cells(targetRow, 3).Value = vendorName
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 정량평가복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsQuant As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long, vendorName As String
    For rowIndex = wsQuant.Cells(wsQuant.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsQuant.Cells(rowIndex, 1).Value = seq Then wsQuant.Rows(rowIndex).Delete
    Next rowIndex

    targetRow = wsQuant.Cells(wsQuant.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" And Trim(wsIn.Cells(sourceRow, 3).Value & "") <> "" Then
            wsQuant.Cells(targetRow, 1).Value = seq
            wsQuant.Cells(targetRow, 2).Value = sourceRow - 43
            wsQuant.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 3).Value
            wsQuant.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 4).Value
            wsQuant.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 5).Value
            wsQuant.Cells(targetRow, 6).Value = wsIn.Cells(sourceRow, 6).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 정성평가복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsQual As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long, vendorName As String
    For rowIndex = wsQual.Cells(wsQual.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsQual.Cells(rowIndex, 1).Value = seq Then wsQual.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsQual.Cells(wsQual.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" And Trim(wsIn.Cells(sourceRow, 8).Value & "") <> "" Then
            wsQual.Cells(targetRow, 1).Value = seq
            wsQual.Cells(targetRow, 2).Value = sourceRow - 43
            wsQual.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 8).Value
            wsQual.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 9).Value
            wsQual.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 10).Value
            wsQual.Cells(targetRow, 6).Value = wsIn.Cells(sourceRow, 11).Value
            wsQual.Cells(targetRow, 7).Value = wsIn.Cells(sourceRow, 12).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 자기평점복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsSelf As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsSelf.Cells(wsSelf.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsSelf.Cells(rowIndex, 1).Value = seq Then wsSelf.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsSelf.Cells(wsSelf.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 14).Value & "") <> "" Then
            wsSelf.Cells(targetRow, 1).Value = seq: wsSelf.Cells(targetRow, 2).Value = sourceRow - 43
            wsSelf.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 14).Value: wsSelf.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 15).Value
            wsSelf.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 16).Value: wsSelf.Cells(targetRow, 6).Value = wsIn.Cells(sourceRow, 17).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 위원복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsCommittee As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsCommittee.Cells(wsCommittee.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsCommittee.Cells(rowIndex, 1).Value = seq Then wsCommittee.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsCommittee.Cells(wsCommittee.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 58 To 67
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" Then
            wsCommittee.Cells(targetRow, 1).Value = seq
            wsCommittee.Cells(targetRow, 2).Value = sourceRow - 57
            wsCommittee.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 2).Value
            wsCommittee.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 3).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 평가복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsScore As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsScore.Cells(wsScore.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsScore.Cells(rowIndex, 1).Value = seq Then wsScore.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsScore.Cells(wsScore.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 72 To 81
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" Then
            wsScore.Cells(targetRow, 1).Value = seq
            wsScore.Cells(targetRow, 2).Value = sourceRow - 71
            wsScore.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 2).Value
            wsScore.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 3).Value
            wsScore.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 4).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Sub 저장하기()
    Call 저장하기_실행(True)
End Sub

Public Sub 검증_저장하기()
    Call 저장하기_실행(False)
End Sub

Private Sub 저장하기_실행(ByVal showMessage As Boolean)
    If Not 필수값검증() Then Exit Sub
    Dim wsIn As Worksheet, wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")
    If Not 품목행검증(wsIn, showMessage) Then Exit Sub
    If Not 업체행검증(wsIn, showMessage) Then Exit Sub
    If Not 정량평가행검증(wsIn, showMessage) Then Exit Sub
    If Not 정성평가행검증(wsIn, showMessage) Then Exit Sub
    If Not 자기평점행검증(wsIn, showMessage) Then Exit Sub
    If Not 위원행검증(wsIn, showMessage) Then Exit Sub
    If Not 평가행검증(wsIn, showMessage) Then Exit Sub
    If showMessage And 검증_업체중복경고() Then MsgBox "중복된 업체명이 있습니다. 실제 동일 업체인지 확인한 뒤 저장하세요.", vbExclamation
    If showMessage And 검증_위원중복경고() Then MsgBox "중복된 위원 역할·직위가 있습니다. 실제 구성과 역할을 확인하세요.", vbExclamation

    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    Dim newRow As Long
    If lastRow <= 1 And wsDB.Cells(2, 1).Value = "" Then
        newRow = 2
    Else
        newRow = lastRow + 1
    End If
    Dim seq As Long
    seq = newRow - 1

    wsDB.Cells(newRow, 1).Value = seq
    Call 필드복사_기초자료_DB(wsIn, wsDB, newRow)
    Call 품목복사_기초자료_DB(wsIn, wsItems, seq)
    Call 업체복사_기초자료_DB(wsIn, wsVendors, seq)
    Call 정량평가복사_기초자료_DB(wsIn, wsQuant, seq)
    Call 정성평가복사_기초자료_DB(wsIn, wsQual, seq)
    Call 자기평점복사_기초자료_DB(wsIn, wsSelf, seq)
    Call 위원복사_기초자료_DB(wsIn, wsCommittee, seq)
    Call 평가복사_기초자료_DB(wsIn, wsScore, seq)

    wsIn.Range("H1").Value = seq
    If showMessage Then MsgBox "저장되었습니다. (순번 " & seq & ")", vbInformation
End Sub

Sub 수정하기()
    Call 수정하기_실행(True)
End Sub

Public Sub 검증_수정하기()
    Call 수정하기_실행(False)
End Sub

Private Sub 수정하기_실행(ByVal showMessage As Boolean)
    If Not 필수값검증() Then Exit Sub
    Dim wsIn As Worksheet, wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")
    If Not 품목행검증(wsIn, showMessage) Then Exit Sub
    If Not 업체행검증(wsIn, showMessage) Then Exit Sub
    If Not 정량평가행검증(wsIn, showMessage) Then Exit Sub
    If Not 정성평가행검증(wsIn, showMessage) Then Exit Sub
    If Not 자기평점행검증(wsIn, showMessage) Then Exit Sub
    If Not 위원행검증(wsIn, showMessage) Then Exit Sub
    If Not 평가행검증(wsIn, showMessage) Then Exit Sub
    If showMessage And 검증_업체중복경고() Then MsgBox "중복된 업체명이 있습니다. 실제 동일 업체인지 확인한 뒤 수정하세요.", vbExclamation
    If showMessage And 검증_위원중복경고() Then MsgBox "중복된 위원 역할·직위가 있습니다. 실제 구성과 역할을 확인하세요.", vbExclamation

    Dim seq As Variant
    seq = wsIn.Range("H1").Value
    If seq = "" Then
        MsgBox "먼저 [불러오기]로 수정할 레코드를 선택하거나, 신규 레코드는 [저장하기]를 사용하세요.", vbExclamation
        Exit Sub
    End If
    Dim r As Long
    r = CLng(seq) + 1
    If wsDB.Cells(r, 1).Value <> CLng(seq) Then
        MsgBox "DB에서 해당 레코드를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If
    Call 필드복사_기초자료_DB(wsIn, wsDB, r)
    Call 품목복사_기초자료_DB(wsIn, wsItems, CLng(seq))
    Call 업체복사_기초자료_DB(wsIn, wsVendors, CLng(seq))
    Call 정량평가복사_기초자료_DB(wsIn, wsQuant, CLng(seq))
    Call 정성평가복사_기초자료_DB(wsIn, wsQual, CLng(seq))
    Call 자기평점복사_기초자료_DB(wsIn, wsSelf, CLng(seq))
    Call 위원복사_기초자료_DB(wsIn, wsCommittee, CLng(seq))
    Call 평가복사_기초자료_DB(wsIn, wsScore, CLng(seq))
    If showMessage Then MsgBox "수정되었습니다. (순번 " & seq & ")", vbInformation
End Sub

Sub 불러오기()
    Dim seqStr As String
    seqStr = InputBox("불러올 레코드의 순번을 입력하세요. (DB 시트 A열 참고)", "레코드 불러오기")
    If seqStr = "" Then Exit Sub
    If Not IsNumeric(seqStr) Then
        MsgBox "숫자를 입력하세요.", vbExclamation
        Exit Sub
    End If
    Call 불러오기_순번(CLng(seqStr), True)
End Sub

Public Sub 검증_첫레코드불러오기()
    Call 불러오기_순번(1, False)
End Sub

Private Sub 불러오기_순번(ByVal seq As Long, ByVal showMessage As Boolean)
    Dim wsIn As Worksheet, wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")
    Dim r As Long
    r = seq + 1
    If wsDB.Cells(r, 1).Value <> seq Then
        If showMessage Then MsgBox "해당 순번의 레코드를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If
    wsIn.Range("C4").Value = wsDB.Cells(r, 2).Value
    wsIn.Range("C5").Value = wsDB.Cells(r, 3).Value
    wsIn.Range("C6").Value = wsDB.Cells(r, 4).Value
    wsIn.Range("C7").Value = wsDB.Cells(r, 5).Value
    wsIn.Range("C8").Value = wsDB.Cells(r, 6).Value
    wsIn.Range("C9").Value = wsDB.Cells(r, 7).Value
    wsIn.Range("C12").Value = wsDB.Cells(r, 8).Value
    wsIn.Range("C13").Value = wsDB.Cells(r, 9).Value
    wsIn.Range("C14").Value = wsDB.Cells(r, 10).Value
    wsIn.Range("C15").Value = wsDB.Cells(r, 11).Value
    wsIn.Range("C16").Value = wsDB.Cells(r, 12).Value
    wsIn.Range("C17").Value = wsDB.Cells(r, 13).Value
    wsIn.Range("C18").Value = wsDB.Cells(r, 14).Value
    wsIn.Range("C21").Value = wsDB.Cells(r, 15).Value
    wsIn.Range("C22").Value = wsDB.Cells(r, 16).Value
    wsIn.Range("C23").Value = wsDB.Cells(r, 17).Value
    wsIn.Range("C24").Value = wsDB.Cells(r, 18).Value
    wsIn.Range("C25").Value = wsDB.Cells(r, 19).Value
    wsIn.Range("B29:D38").ClearContents
    Dim itemRow As Long, inputRow As Long
    For itemRow = 2 To wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row
        If wsItems.Cells(itemRow, 1).Value = seq Then
            inputRow = 28 + CLng(wsItems.Cells(itemRow, 2).Value)
            If inputRow >= 29 And inputRow <= 38 Then
                wsIn.Cells(inputRow, 2).Value = wsItems.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 3).Value = wsItems.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 4).Value = wsItems.Cells(itemRow, 5).Value
            End If
        End If
    Next itemRow
    wsIn.Range("B58:C67").ClearContents
    For itemRow = 2 To wsCommittee.Cells(wsCommittee.Rows.Count, 1).End(xlUp).Row
        If wsCommittee.Cells(itemRow, 1).Value = seq Then
            inputRow = 57 + CLng(wsCommittee.Cells(itemRow, 2).Value)
            If inputRow >= 58 And inputRow <= 67 Then
                wsIn.Cells(inputRow, 2).Value = wsCommittee.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 3).Value = wsCommittee.Cells(itemRow, 4).Value
            End If
        End If
    Next itemRow
    wsIn.Range("B72:D81").ClearContents
    For itemRow = 2 To wsScore.Cells(wsScore.Rows.Count, 1).End(xlUp).Row
        If wsScore.Cells(itemRow, 1).Value = seq Then
            inputRow = 71 + CLng(wsScore.Cells(itemRow, 2).Value)
            If inputRow >= 72 And inputRow <= 81 Then
                wsIn.Cells(inputRow, 2).Value = wsScore.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 3).Value = wsScore.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 4).Value = wsScore.Cells(itemRow, 5).Value
            End If
        End If
    Next itemRow
    wsIn.Range("B44:F53").ClearContents
    wsIn.Range("H44:L53").ClearContents
    wsIn.Range("N44:Q53").ClearContents
    For itemRow = 2 To wsVendors.Cells(wsVendors.Rows.Count, 1).End(xlUp).Row
        If wsVendors.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsVendors.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then wsIn.Cells(inputRow, 2).Value = wsVendors.Cells(itemRow, 3).Value
        End If
    Next itemRow
    For itemRow = 2 To wsSelf.Cells(wsSelf.Rows.Count, 1).End(xlUp).Row
        If wsSelf.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsSelf.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then
                wsIn.Cells(inputRow, 14).Value = wsSelf.Cells(itemRow, 3).Value: wsIn.Cells(inputRow, 15).Value = wsSelf.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 16).Value = wsSelf.Cells(itemRow, 5).Value: wsIn.Cells(inputRow, 17).Value = wsSelf.Cells(itemRow, 6).Value
            End If
        End If
    Next itemRow
    For itemRow = 2 To wsQual.Cells(wsQual.Rows.Count, 1).End(xlUp).Row
        If wsQual.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsQual.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then
                wsIn.Cells(inputRow, 8).Value = wsQual.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 9).Value = wsQual.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 10).Value = wsQual.Cells(itemRow, 5).Value
                wsIn.Cells(inputRow, 11).Value = wsQual.Cells(itemRow, 6).Value
                wsIn.Cells(inputRow, 12).Value = wsQual.Cells(itemRow, 7).Value
            End If
        End If
    Next itemRow
    For itemRow = 2 To wsQuant.Cells(wsQuant.Rows.Count, 1).End(xlUp).Row
        If wsQuant.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsQuant.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then
                wsIn.Cells(inputRow, 3).Value = wsQuant.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 4).Value = wsQuant.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 5).Value = wsQuant.Cells(itemRow, 5).Value
                wsIn.Cells(inputRow, 6).Value = wsQuant.Cells(itemRow, 6).Value
            End If
        End If
    Next itemRow
    wsIn.Range("H1").Value = seq
    If showMessage Then MsgBox "불러왔습니다. (순번 " & seq & ")", vbInformation
End Sub
'@

    $codeInput = $codeInput -replace "`r`n", "`r" -replace "`n", "`r"
    $modInput.CodeModule.AddFromString($codeInput)
    L "Module_기초자료 추가 완료 (줄 수: $($modInput.CodeModule.CountOfLines))"
    try { $wb.Save(); L "DIAG_CHECKPOINT1_SAVE_OK" } catch { L "DIAG_CHECKPOINT1_SAVE_FAIL: $($_.Exception.Message)" }

    $modOutput = $vbproj.VBComponents.Add(1)
    $modOutput.Name = "Module_출력"
    $codeOutput = @'
Option Explicit

' 2026-09-19 진단: 이 프로시저가 Form ID마다 개별 skippedInvalidFxxx 카운터 변수와
' 반복되는 "wsSel.Cells(r,2).Value = "F-XXX"" 비교, 그리고 그 변수들을 전부 나열한
' 초장문 skipDetail 연결식·OR 조건식을 갖고 있어, Form ID를 추가할 때마다 이 프로시저
' 하나의 컴파일 크기가 커지다가 VBA "프로시저가 너무 큼" 한계를 실제로 넘어 저장(컴파일)
' 자체가 실패하는 사고가 재현됨(이분 탐색으로 확정). 재발 방지를 위해 카운터를 Dictionary
' 하나로, "단순 단일 출력" 그룹(업체 반복 없이 검증 함수 1개만 통과하면 서식선택_출력의
' F열 시트를 그대로 출력하는 F-001~006·014·024·032~043)의 20개 ElseIf 문자열 비교를
' Select Case 조회 함수(단일출력검증함수명)로 옮겨 이 프로시저 자체의 크기를 줄인다.
' 업체 반복 출력 그룹(F-015~023·025)의 동작 로직은 전혀 바꾸지 않았다(캐시된 fid 변수
' 사용과 IncSkip 호출로만 치환).
Private Sub IncSkip(ByRef d As Object, ByVal key As String)
    If d.Exists(key) Then
        d(key) = d(key) + 1
    Else
        d.Add key, 1
    End If
End Sub

Private Function 단일출력검증함수명(ByVal fid As String) As String
    Select Case fid
        Case "F-001": 단일출력검증함수명 = "검증_F001출력가능"
        Case "F-002": 단일출력검증함수명 = "검증_F002출력가능"
        Case "F-003": 단일출력검증함수명 = "검증_F003출력가능"
        Case "F-004": 단일출력검증함수명 = "검증_F004출력가능"
        Case "F-005": 단일출력검증함수명 = "검증_F005출력가능"
        Case "F-006": 단일출력검증함수명 = "검증_F006출력가능"
        Case "F-014": 단일출력검증함수명 = "검증_F014출력가능"
        Case "F-024": 단일출력검증함수명 = "검증_F024출력가능"
        Case "F-032": 단일출력검증함수명 = "검증_F032출력가능"
        Case "F-033": 단일출력검증함수명 = "검증_F033출력가능"
        Case "F-034": 단일출력검증함수명 = "검증_F034출력가능"
        Case "F-035": 단일출력검증함수명 = "검증_F035출력가능"
        Case "F-036": 단일출력검증함수명 = "검증_F036출력가능"
        Case "F-037": 단일출력검증함수명 = "검증_F037출력가능"
        Case "F-038": 단일출력검증함수명 = "검증_F038출력가능"
        Case "F-039": 단일출력검증함수명 = "검증_F039출력가능"
        Case "F-040": 단일출력검증함수명 = "검증_F040출력가능"
        Case "F-041": 단일출력검증함수명 = "검증_F041출력가능"
        Case "F-042": 단일출력검증함수명 = "검증_F042출력가능"
        Case "F-043": 단일출력검증함수명 = "검증_F043출력가능"
        Case "F-044": 단일출력검증함수명 = "검증_F044출력가능"
        Case "F-045": 단일출력검증함수명 = "검증_F045출력가능"
        Case "F-046": 단일출력검증함수명 = "검증_F046출력가능"
        Case "F-047": 단일출력검증함수명 = "검증_F047출력가능"
        Case "F-048": 단일출력검증함수명 = "검증_F048출력가능"
        Case Else: 단일출력검증함수명 = ""
    End Select
End Function

Private Function 선택된시트목록(ByRef outNames() As String) As Long
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim lastRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row

    Dim tmp() As String
    ReDim tmp(1 To lastRow + 50)   ' 업체별 F-015~F-020 출력에 대비해 여유를 둠
    Dim cnt As Long
    cnt = 0
    Dim skipCounts As Object
    Set skipCounts = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 5 To lastRow
        If wsSel.Cells(r, 1).Value = True Then
            Dim status As String, fid As String
            status = wsSel.Cells(r, 5).Value
            fid = wsSel.Cells(r, 2).Value

            Dim simpleValidator As String
            simpleValidator = 단일출력검증함수명(fid)

            If simpleValidator <> "" Then
                If CBool(Application.Run(simpleValidator)) Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    IncSkip skipCounts, fid
                End If
            ElseIf fid = "F-015" Then
                If Not 검증_F015출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim vendorRow As Long, addedAny As Boolean
                    addedAny = False
                    For vendorRow = 44 To 53
                        If F015업체행완전한가(vendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F015임시시트생성(vendorRow - 43)
                            addedAny = True
                        End If
                    Next vendorRow
                    If Not addedAny Then IncSkip skipCounts, fid
                End If
            ElseIf fid = "F-016" Then
                If Not 검증_F016출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim qualitativeVendorRow As Long, addedQualitative As Boolean
                    addedQualitative = False
                    For qualitativeVendorRow = 44 To 53
                        If F016업체행완전한가(qualitativeVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F016임시시트생성(qualitativeVendorRow - 43)
                            addedQualitative = True
                        End If
                    Next qualitativeVendorRow
                    If Not addedQualitative Then IncSkip skipCounts, fid
                End If
            ElseIf fid = "F-017" Then
                If Not 검증_F017출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim confirmationVendorRow As Long
                    For confirmationVendorRow = 44 To 53
                        If F017업체행완전한가(confirmationVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F017임시시트생성(confirmationVendorRow - 43)
                        End If
                    Next confirmationVendorRow
                End If
            ElseIf fid = "F-018" Then
                If Not 검증_F018출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim selfVendorRow As Long
                    For selfVendorRow = 44 To 53
                        If F018업체행완전한가(selfVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F018임시시트생성(selfVendorRow - 43)
                        End If
                    Next selfVendorRow
                End If
            ElseIf fid = "F-019" Then
                If Not 검증_F019출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim applicationVendorRow As Long
                    For applicationVendorRow = 44 To 53
                        If F019업체행완전한가(applicationVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F019임시시트생성(applicationVendorRow - 43)
                        End If
                    Next applicationVendorRow
                End If
            ElseIf fid = "F-020" Then
                If Not 검증_F020출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim reportVendorRow As Long
                    For reportVendorRow = 44 To 53
                        If F020업체행완전한가(reportVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F020임시시트생성(reportVendorRow - 43)
                        End If
                    Next reportVendorRow
                End If
            ElseIf fid = "F-021" Then
                If Not 검증_F021출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim proposalVendorRow As Long
                    For proposalVendorRow = 44 To 53
                        If F021업체행완전한가(proposalVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F021임시시트생성(proposalVendorRow - 43)
                        End If
                    Next proposalVendorRow
                End If
            ElseIf fid = "F-022" Then
                If Not 검증_F022출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim performanceVendorRow As Long
                    For performanceVendorRow = 44 To 53
                        If F022업체행완전한가(performanceVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F022임시시트생성(performanceVendorRow - 43)
                        End If
                    Next performanceVendorRow
                End If
            ElseIf fid = "F-023" Then
                If Not 검증_F023출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim specVendorRow As Long
                    For specVendorRow = 44 To 53
                        If F023업체행완전한가(specVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F023임시시트생성(specVendorRow - 43)
                        End If
                    Next specVendorRow
                End If
            ElseIf fid = "F-025" Then
                If Not 검증_F025출력가능() Then
                    IncSkip skipCounts, fid
                Else
                    Dim asVendorRow As Long
                    For asVendorRow = 44 To 53
                        If F025업체행완전한가(asVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F025임시시트생성(asVendorRow - 43)
                        End If
                    Next asVendorRow
                End If
            ElseIf status = "Y" Then
                cnt = cnt + 1
                tmp(cnt) = wsSel.Cells(r, 6).Value
            ElseIf status = "D" Then
                IncSkip skipCounts, "HWPX 우선순위 위임"
            Else
                IncSkip skipCounts, "미구현"
            End If
        End If
    Next r

    Dim skipDetail As String
    skipDetail = ""
    Dim skipKey As Variant
    For Each skipKey In skipCounts.Keys
        skipDetail = skipDetail & skipKey & " 제외: " & skipCounts(skipKey) & "건, "
    Next skipKey
    If Len(skipDetail) >= 2 Then skipDetail = Left(skipDetail, Len(skipDetail) - 2)

    If cnt = 0 Then
        MsgBox "구현된 서식이 선택되지 않았습니다." & vbCrLf & skipDetail, vbExclamation
        선택된시트목록 = 0
        Exit Function
    End If

    If skipCounts.Count > 0 Then
        MsgBox "일부 선택 서식은 아직 준비되지 않아 제외합니다." & vbCrLf & skipDetail, vbInformation
    End If

    ReDim outNames(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        outNames(i) = tmp(i)
    Next i
    선택된시트목록 = cnt
End Function

Public Function 검증_F024출력가능() As Boolean
    Dim itemCount As Long
    itemCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B29:B38"))
    검증_F024출력가능 = 검증_품목행검증() And itemCount >= 1 And itemCount <= 6
End Function

Public Function 검증_F014출력가능() As Boolean
    Dim scoreCount As Long
    scoreCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B72:B81"))
    검증_F014출력가능 = 검증_필수값검증() And 검증_평가행검증() And scoreCount >= 1 And scoreCount <= 10
End Function

Public Function 검증_F015출력가능() As Boolean
    Dim vendorRow As Long, hasCompleteVendor As Boolean
    If Not 검증_정량평가행검증() Then Exit Function
    hasCompleteVendor = False
    For vendorRow = 44 To 53
        If F015업체행완전한가(vendorRow) Then
            hasCompleteVendor = True
            Exit For
        End If
    Next vendorRow
    검증_F015출력가능 = 검증_필수값검증() And hasCompleteVendor
End Function

Public Function 검증_F016출력가능() As Boolean
    Dim vendorRow As Long, hasCompleteVendor As Boolean
    If Not 검증_정성평가행검증() Then Exit Function
    For vendorRow = 44 To 53
        If F016업체행완전한가(vendorRow) Then
            hasCompleteVendor = True
            Exit For
        End If
    Next vendorRow
    검증_F016출력가능 = 검증_필수값검증() And hasCompleteVendor
End Function

Public Function 검증_F017출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F017업체행완전한가(vendorRow) Then
            검증_F017출력가능 = True
            Exit Function
        End If
    Next vendorRow
End Function

Public Function 검증_F018출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Or Not 검증_자기평점행검증() Then Exit Function
    For vendorRow = 44 To 53
        If F018업체행완전한가(vendorRow) Then 검증_F018출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F019출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F019업체행완전한가(vendorRow) Then 검증_F019출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F020출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F020업체행완전한가(vendorRow) Then 검증_F020출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F021출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F021업체행완전한가(vendorRow) Then 검증_F021출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F022출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F022업체행완전한가(vendorRow) Then 검증_F022출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F023출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F023업체행완전한가(vendorRow) Then 검증_F023출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F025출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F025업체행완전한가(vendorRow) Then 검증_F025출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F001출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F001출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

Public Function 검증_F002출력가능() As Boolean
    검증_F002출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F003출력가능() As Boolean
    검증_F003출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F004출력가능() As Boolean
    검증_F004출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F005출력가능() As Boolean
    검증_F005출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F006출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F006출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

' F-041~F-043은 F-017~F-025처럼 업체마다 별도 문서를 만드는 것이 아니라, 낙찰자·계약상대자는
' 언제나 1개 업체이므로 각 시트의 I3 선택 순번이 가리키는 업체행 1곳만 완전한지(상호명이 비어있지
' 않은지) 확인한다. Module_출력에서 업체 반복 루프를 돌리지 않고 F-001/F-006과 같은 단일 출력으로
' 취급하는 이유가 이 검증 함수에도 그대로 반영된다.
Private Function 낙찰업체선택완전한가(ByVal sheetName As String) As Boolean
    Dim wsForm As Worksheet, vendorRow As Long
    Set wsForm = Sheets(sheetName)
    vendorRow = 43 + CLng(Val(wsForm.Range("I3").Value))
    If vendorRow < 44 Or vendorRow > 53 Then Exit Function
    낙찰업체선택완전한가 = Trim(Sheets("기초자료입력").Cells(vendorRow, 2).Value & "") <> ""
End Function

Public Function 검증_F041출력가능() As Boolean
    With Sheets("F-041_낙찰자결정")
        검증_F041출력가능 = 검증_필수값검증() And 낙찰업체선택완전한가("F-041_낙찰자결정") And _
            양수값(.Range("C17").Value) And 양수값(.Range("E21").Value) And 양수값(.Range("F21").Value)
    End With
End Function

Public Function 검증_F042출력가능() As Boolean
    With Sheets("F-042_낙찰자결정통보")
        검증_F042출력가능 = 검증_필수값검증() And 낙찰업체선택완전한가("F-042_낙찰자결정통보") And _
            양수값(.Range("D17").Value) And 양수값(.Range("E17").Value) And 양수값(.Range("F17").Value) And _
            계약금액일치(.Range("D17").Value, .Range("E17").Value, .Range("F17").Value) And _
            Trim(Sheets("기초자료입력").Range("C18").Value & "") <> ""
    End With
End Function

Public Function 검증_F043출력가능() As Boolean
    With Sheets("F-043_계약체결")
        검증_F043출력가능 = 검증_필수값검증() And 낙찰업체선택완전한가("F-043_계약체결") And _
            양수값(.Range("C16").Value) And 양수값(Sheets("기초자료입력").Range("C16").Value) And _
            Trim(Sheets("기초자료입력").Range("C17").Value & "") <> ""
    End With
End Function

Public Function 검증_F044출력가능() As Boolean
    검증_F044출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F045출력가능() As Boolean
    검증_F045출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F046출력가능() As Boolean
    검증_F046출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F048출력가능() As Boolean
    검증_F048출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F047출력가능() As Boolean
    검증_F047출력가능 = 검증_필수값검증()
End Function

Private Function 양수값(ByVal value As Variant) As Boolean
    양수값 = IsNumeric(value) And CDbl(value) > 0
End Function

Private Function 계약금액일치(ByVal unitPrice As Variant, ByVal quantity As Variant, ByVal contractAmount As Variant) As Boolean
    If Not 양수값(unitPrice) Or Not 양수값(quantity) Or Not 양수값(contractAmount) Then Exit Function
    계약금액일치 = Abs(CDbl(contractAmount) - WorksheetFunction.Round(CDbl(unitPrice) * CDbl(quantity), 0)) < 0.000001
End Function

Public Function 검증_F037출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F037출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

Public Function 검증_F038출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F038출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

' F-039는 F-041~F-043과 같은 I3 단일 업체 선택 패턴을 재사용한다(업체 반복 루프 없음).
' 위원별 원점수(옷감·완성도·A/S·하자)는 이 문서 전용 입력칸(E8:H15)이며 다른 서식과
' 공유하는 DB가 아니다. 위원 성명(D열)은 청탁방지·익명성 원칙에 따라 자동 반영하지 않고
' 자필 공란으로 유지하므로 검증 대상에서 제외한다. 총계·평균은 원문의 "위원별 평가 점수
' 중 최고점 및 최저점을 제외" 안내에 따라 8명 점수 중 최댓값·최솟값을 뺀 뒤 합산·평균한다.
Public Function 검증_F039출력가능() As Boolean
    With Sheets("F-039_업체별제안서평가표")
        Dim ok As Boolean
        ok = 검증_필수값검증() And 낙찰업체선택완전한가("F-039_업체별제안서평가표")
        If ok Then
            ok = Application.WorksheetFunction.Count(.Range("E8:H15")) = 32 And _
                Application.WorksheetFunction.Min(.Range("E8:E15")) >= 0 And Application.WorksheetFunction.Max(.Range("E8:E15")) <= 15 And _
                Application.WorksheetFunction.Min(.Range("F8:F15")) >= 0 And Application.WorksheetFunction.Max(.Range("F8:F15")) <= 10 And _
                Application.WorksheetFunction.Min(.Range("G8:G15")) >= 0 And Application.WorksheetFunction.Max(.Range("G8:G15")) <= 15 And _
                Application.WorksheetFunction.Min(.Range("H8:H15")) >= 0 And Application.WorksheetFunction.Max(.Range("H8:H15")) <= 10
        End If
        검증_F039출력가능 = ok
    End With
End Function

Public Function 검증_F040출력가능() As Boolean
    검증_F040출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F032출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F032출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Public Function 검증_F033출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F033출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Public Function 검증_F034출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F034출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

' F-035·F-036은 F-032·F-033과 같이 업체 반복행(기초자료입력!B44:B53) 전체를 한 표에
' 보여주는 집계 보고서이며, 점수는 F-015/F-016이 이미 쓰는 학교 평가 점수 열(C:F 정량,
' H:L 정성)을 그대로 재조회한다. 업체 존재만 확인하고 점수 입력 여부는 강제하지 않는다
' (F-032/F-033과 동일 수준 — 평가 진행 중에도 접수 현황·중간 보고서를 출력할 수 있어야 함).
Public Function 검증_F035출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F035출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Public Function 검증_F036출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F036출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Private Function F015업체행완전한가(ByVal sourceRow As Long) As Boolean
    Dim wsIn As Worksheet
    Set wsIn = Sheets("기초자료입력")
    F015업체행완전한가 = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 3).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 4).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 5).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 6).Value & "") <> "" And _
        정수범위(wsIn.Cells(sourceRow, 3).Value, 0, 10) And 정수범위(wsIn.Cells(sourceRow, 4).Value, 0, 10) And _
        정수범위(wsIn.Cells(sourceRow, 5).Value, 0, 15) And 정수범위(wsIn.Cells(sourceRow, 6).Value, 0, 15)
End Function

Private Function F016업체행완전한가(ByVal sourceRow As Long) As Boolean
    Dim wsIn As Worksheet
    Set wsIn = Sheets("기초자료입력")
    F016업체행완전한가 = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 8).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 9).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 10).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 11).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 12).Value & "") <> "" And _
        정수범위(wsIn.Cells(sourceRow, 8).Value, 0, 15) And 정수범위(wsIn.Cells(sourceRow, 9).Value, 0, 10) And _
        정수범위(wsIn.Cells(sourceRow, 10).Value, 0, 15) And 정수범위(wsIn.Cells(sourceRow, 11).Value, 0, 10) And _
        정수범위(wsIn.Cells(sourceRow, 12).Value, -15, 5)
End Function

Private Function F017업체행완전한가(ByVal sourceRow As Long) As Boolean
    F017업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F018업체행완전한가(ByVal sourceRow As Long) As Boolean
    Dim wsIn As Worksheet
    Set wsIn = Sheets("기초자료입력")
    F018업체행완전한가 = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 14).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 15).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 16).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 17).Value & "") <> "" And 정수범위(wsIn.Cells(sourceRow, 14).Value, 0, 10) And 정수범위(wsIn.Cells(sourceRow, 15).Value, 0, 10) And 정수범위(wsIn.Cells(sourceRow, 16).Value, 0, 15) And 정수범위(wsIn.Cells(sourceRow, 17).Value, 0, 15)
End Function

Private Function F019업체행완전한가(ByVal sourceRow As Long) As Boolean
    F019업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F020업체행완전한가(ByVal sourceRow As Long) As Boolean
    F020업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F021업체행완전한가(ByVal sourceRow As Long) As Boolean
    F021업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F022업체행완전한가(ByVal sourceRow As Long) As Boolean
    F022업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F023업체행완전한가(ByVal sourceRow As Long) As Boolean
    F023업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F025업체행완전한가(ByVal sourceRow As Long) As Boolean
    F025업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F015임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF15 As Worksheet, wsTemp As Worksheet
    Set wsF15 = Sheets("F-015_정량적평가")
    wsF15.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F015_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F015임시시트생성 = wsTemp.Name
End Function

Private Function F016임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF16 As Worksheet, wsTemp As Worksheet
    Set wsF16 = Sheets("F-016_정성적평가")
    wsF16.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F016_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F016임시시트생성 = wsTemp.Name
End Function

Private Function F017임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF17 As Worksheet, wsTemp As Worksheet
    Set wsF17 = Sheets("F-017_제출서류자기확인서")
    wsF17.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F017_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F017임시시트생성 = wsTemp.Name
End Function

Private Function F018임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF18 As Worksheet, wsTemp As Worksheet
    Set wsF18 = Sheets("F-018_정량적평가자기평점표")
    wsF18.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F018_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F018임시시트생성 = wsTemp.Name
End Function

Private Function F019임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF19 As Worksheet, wsTemp As Worksheet
    Set wsF19 = Sheets("F-019_입찰참가신청서")
    wsF19.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F019_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("N3").Value = vendorSlot
    F019임시시트생성 = wsTemp.Name
End Function

Private Function F020임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF20 As Worksheet, wsTemp As Worksheet
    Set wsF20 = Sheets("F-020_입찰참가신고서")
    wsF20.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F020_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F020임시시트생성 = wsTemp.Name
End Function

Private Function F021임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF21 As Worksheet, wsTemp As Worksheet
    Set wsF21 = Sheets("F-021_교복납품제안서")
    wsF21.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F021_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F021임시시트생성 = wsTemp.Name
End Function

Private Function F022임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF22 As Worksheet, wsTemp As Worksheet
    Set wsF22 = Sheets("F-022_교복납품실적표")
    wsF22.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F022_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F022임시시트생성 = wsTemp.Name
End Function

Private Function F023임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF23 As Worksheet, wsTemp As Worksheet
    Set wsF23 = Sheets("F-023_교복제조사양서")
    wsF23.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F023_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F023임시시트생성 = wsTemp.Name
End Function

Private Function F025임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF25 As Worksheet, wsTemp As Worksheet
    Set wsF25 = Sheets("F-025_교복AS계획서")
    wsF25.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = Sheets(Sheets.Count)
    wsTemp.Name = "F025_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F025임시시트생성 = wsTemp.Name
End Function

Private Sub F015임시시트정리(ByRef names() As String)
    Dim i As Long
    On Error Resume Next
    Application.DisplayAlerts = False
    For i = LBound(names) To UBound(names)
        If Left(names(i), 5) = "F015_" Or Left(names(i), 5) = "F016_" Or Left(names(i), 5) = "F017_" Or Left(names(i), 5) = "F018_" Or Left(names(i), 5) = "F019_" Or Left(names(i), 5) = "F020_" Or Left(names(i), 5) = "F021_" Or Left(names(i), 5) = "F022_" Or Left(names(i), 5) = "F023_" Or Left(names(i), 5) = "F025_" Then Sheets(names(i)).Delete
    Next i
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Sub 선택서식_인쇄미리보기()
    Dim names() As String
    Dim n As Long
    n = 선택된시트목록(names)
    If n = 0 Then Exit Sub
    ThisWorkbook.Sheets(names).Select
    ActiveWindow.SelectedSheets.PrintPreview
    Call F015임시시트정리(names)
End Sub

Sub 선택서식_PDF저장()
    Call 선택서식_PDF저장_실행(True)
End Sub

Public Sub 검증_선택서식_PDF저장()
    Call 선택서식_PDF저장_실행(False)
End Sub

Private Sub 선택서식_PDF저장_실행(ByVal showMessage As Boolean)
    Dim names() As String
    Dim n As Long
    n = 선택된시트목록(names)
    If n = 0 Then Exit Sub

    Dim folderPath As String
    folderPath = ThisWorkbook.Path & "\output\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    Dim fileName As String
    fileName = folderPath & "선택서식_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    ThisWorkbook.Sheets(names).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Call F015임시시트정리(names)

    If showMessage Then MsgBox "PDF로 저장되었습니다:" & vbCrLf & fileName, vbInformation
End Sub
'@
    $codeOutput = $codeOutput -replace "`r`n", "`r" -replace "`n", "`r"
    $modOutput.CodeModule.AddFromString($codeOutput)
    L "Module_출력 추가 완료 (줄 수: $($modOutput.CodeModule.CountOfLines))"
    try { $wb.Save(); L "DIAG_CHECKPOINT2_SAVE_OK" } catch { L "DIAG_CHECKPOINT2_SAVE_FAIL: $($_.Exception.Message)" }

    $modNavigation = $vbproj.VBComponents.Add(1)
    $modNavigation.Name = "Module_검색_절차"
    $codeNavigation = @'
Option Explicit

Sub 계약방법안내_기초자료반영()
    Call 계약방법안내_기초자료반영_내부(True)
End Sub

Public Sub 검증_계약방법안내_기초자료반영()
    Call 계약방법안내_기초자료반영_내부(False)
End Sub

Private Sub 계약방법안내_기초자료반영_내부(ByVal showMessage As Boolean)
    Dim wsMethod As Worksheet, selectedMethod As String
    Set wsMethod = Sheets("계약방법안내")
    If Trim(wsMethod.Range("B5").Value & "") <> "예" Or Trim(wsMethod.Range("B6").Value & "") <> "예" Then
        If showMessage Then MsgBox "두 확인 항목이 모두 '예'인 경우에만 매뉴얼 예시를 반영합니다. 그 밖의 계약방법은 최신 법령·지침과 사실관계를 행정실 담당자가 검토하세요.", vbExclamation
        Exit Sub
    End If
    selectedMethod = Trim(wsMethod.Range("B7").Value & "")
    If selectedMethod = "" Then
        If showMessage Then MsgBox "반영할 계약방법 안내값이 없습니다.", vbExclamation
        Exit Sub
    End If
    Sheets("기초자료입력").Range("C14").Value = selectedMethod
    If showMessage Then MsgBox "계약방식(B-03)에 매뉴얼 예시를 반영했습니다. 출력 전 최신 법령·지침과 사실관계를 최종 확인하세요.", vbInformation
End Sub

Sub 학교검색_실행()
    Call 학교검색_실행_내부(True)
End Sub

Public Sub 검증_학교검색_실행()
    Call 학교검색_실행_내부(False)
End Sub

Private Sub 학교검색_실행_내부(ByVal showMessage As Boolean)
    Dim wsSearch As Worksheet, wsSchool As Worksheet
    Dim schoolName As String, regionName As String, schoolLevel As String
    Dim lastRow As Long, sourceRow As Long, resultRow As Long, count As Long

    Set wsSearch = Sheets("학교검색")
    Set wsSchool = Sheets("학교정보")
    schoolName = Trim(wsSearch.Range("B3").Value & "")
    regionName = Trim(wsSearch.Range("D3").Value & "")
    schoolLevel = Trim(wsSearch.Range("F3").Value & "")
    wsSearch.Range("A7:F126").ClearContents

    lastRow = wsSchool.Cells(wsSchool.Rows.Count, 5).End(xlUp).Row
    resultRow = 7
    count = 0
    For sourceRow = 2 To lastRow
        If (schoolName = "" Or InStr(1, wsSchool.Cells(sourceRow, 5).Value & "", schoolName, vbTextCompare) > 0) And _
           (regionName = "" Or InStr(1, wsSchool.Cells(sourceRow, 2).Value & "", regionName, vbTextCompare) > 0) And _
           (schoolLevel = "" Or InStr(1, wsSchool.Cells(sourceRow, 3).Value & "", schoolLevel, vbTextCompare) > 0) Then
            wsSearch.Cells(resultRow, 1).Value = wsSchool.Cells(sourceRow, 1).Value
            wsSearch.Cells(resultRow, 2).Value = wsSchool.Cells(sourceRow, 2).Value
            wsSearch.Cells(resultRow, 3).Value = wsSchool.Cells(sourceRow, 3).Value
            wsSearch.Cells(resultRow, 4).Value = wsSchool.Cells(sourceRow, 5).Value
            wsSearch.Cells(resultRow, 5).Value = wsSchool.Cells(sourceRow, 6).Value
            wsSearch.Cells(resultRow, 6).Value = wsSchool.Cells(sourceRow, 7).Value
            resultRow = resultRow + 1
            count = count + 1
        End If
    Next sourceRow
    If showMessage Then MsgBox count & "건을 찾았습니다. 결과 행을 선택한 뒤 [선택 학교를 기초자료에 반영]을 누르세요.", vbInformation
End Sub

Sub 선택학교_기초자료반영()
    Call 선택학교_기초자료반영_내부(True)
End Sub

Public Sub 검증_선택학교_기초자료반영()
    Call 선택학교_기초자료반영_내부(False)
End Sub

Private Sub 선택학교_기초자료반영_내부(ByVal showMessage As Boolean)
    Dim wsSearch As Worksheet
    Dim resultRow As Long, selectedSchool As String
    Set wsSearch = Sheets("학교검색")
    If ActiveSheet.Name <> wsSearch.Name Then
        If showMessage Then MsgBox "학교검색 시트의 검색 결과 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    resultRow = ActiveCell.Row
    If resultRow < 7 Or resultRow > 126 Then
        If showMessage Then MsgBox "검색 결과의 학교 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    selectedSchool = Trim(wsSearch.Cells(resultRow, 4).Value & "")
    If selectedSchool = "" Then
        If showMessage Then MsgBox "선택한 행에 학교명이 없습니다. 먼저 검색하세요.", vbExclamation
        Exit Sub
    End If
    Sheets("기초자료입력").Range("C4").Value = selectedSchool
    If showMessage Then MsgBox "학교명(C-01)에 '" & selectedSchool & "'을(를) 반영했습니다.", vbInformation
End Sub

Sub 관련FormID로이동()
    Call 관련FormID로이동_내부(True)
End Sub

Public Sub 검증_관련FormID로이동()
    Call 관련FormID로이동_내부(False)
End Sub

Private Sub 관련FormID로이동_내부(ByVal showMessage As Boolean)
    Dim wsFlow As Worksheet, wsSelect As Worksheet
    Dim flowRow As Long, formId As String, lastRow As Long, r As Long
    Set wsFlow = Sheets("절차안내")
    Set wsSelect = Sheets("서식선택_출력")
    If ActiveSheet.Name <> wsFlow.Name Then
        If showMessage Then MsgBox "절차안내 시트의 단계 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    flowRow = ActiveCell.Row
    If flowRow < 5 Or flowRow > 13 Then
        If showMessage Then MsgBox "1~9단계의 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    formId = Trim(wsFlow.Cells(flowRow, 4).Value & "")
    lastRow = wsSelect.Cells(wsSelect.Rows.Count, 2).End(xlUp).Row
    For r = 5 To lastRow
        If wsSelect.Cells(r, 2).Value = formId Then
            wsSelect.Activate
            wsSelect.Cells(r, 2).Select
            Exit Sub
        End If
    Next r
    If showMessage Then MsgBox "이동할 Form ID를 찾지 못했습니다: " & formId, vbExclamation
End Sub
'@
    $codeNavigation = $codeNavigation -replace "`r`n", "`r" -replace "`n", "`r"
    $modNavigation.CodeModule.AddFromString($codeNavigation)
    L "Module_검색_절차 추가 완료 (줄 수: $($modNavigation.CodeModule.CountOfLines))"
    try { $wb.Save(); L "DIAG_CHECKPOINT3_SAVE_OK" } catch { L "DIAG_CHECKPOINT3_SAVE_FAIL: $($_.Exception.Message)" }

    # ---- ThisWorkbook: Workbook_Open (서식선택 체크박스 초기화) ----
    # 참고: 한글 Office에서는 ThisWorkbook 문서모듈의 기본 컴포넌트 이름이 "ThisWorkbook"이 아니라
    # 로컬라이즈된 이름(예: 이_통합_문서)으로 생성됨. Type=100(문서모듈)이면서 워크시트 코드네임
    # 패턴(Sheet숫자)이 아닌 컴포넌트를 찾아 식별함.
    $thisWb = $null
    foreach ($c in $vbproj.VBComponents) {
        if ($c.Type -eq 100 -and $c.Name -notmatch '^Sheet\d+$') {
            $thisWb = $c
            break
        }
    }
    if (-not $thisWb) { throw "ThisWorkbook 문서모듈을 찾지 못함" }
    L "ThisWorkbook 문서모듈 식별: $($thisWb.Name)"

    # 재실행 대비: 기존 Workbook_Open 있으면 프로시저 전체 삭제 후 재추가
    $cm = $thisWb.CodeModule
    try {
        $procName = "Workbook_Open"
        $startLine = $cm.ProcStartLine($procName, 0)
        $procCount = $cm.ProcCountLines($procName, 0)
        $cm.DeleteLines($startLine, $procCount)
        L "기존 Workbook_Open 프로시저 제거 후 재생성"
    } catch { L '기존 Workbook_Open 프로시저 없음' }

    $codeThisWb = @'

Private Sub Workbook_Open()
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim lastRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row
    Dim r As Long
    For r = 5 To lastRow
        wsSel.Cells(r, 1).Value = False
    Next r

    ' UserInterfaceOnly 보호 설정은 파일을 다시 열면 유지되지 않으므로, 매크로 허용 후 다시 적용한다.
    ' 비밀번호는 사용하지 않으며 일반 사용자의 직접 편집만 제한한다.
    Dim internalName As Variant
    For Each internalName In Array("DB", "DB_품목", "DB_업체", "DB_위원", "DB_평가", "DB_정량평가", "DB_정성평가", "DB_자기평점")
        Sheets(CStr(internalName)).Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Next internalName
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim reason As String
    reason = 검증_저장경계_메시지()
    If reason <> "" Then
        Cancel = True
        If Application.Visible And Application.UserControl Then
            MsgBox "저장이 취소되었습니다. " & reason, vbExclamation
        End If
    End If
End Sub
'@
    $codeThisWb = $codeThisWb -replace "`r`n", "`r" -replace "`n", "`r"
    $thisWb.CodeModule.AddFromString($codeThisWb)
    L "ThisWorkbook.Workbook_Open 추가 완료"
    try { $wb.Save(); L "DIAG_CHECKPOINT4_SAVE_OK" } catch { L "DIAG_CHECKPOINT4_SAVE_FAIL: $($_.Exception.Message)" }

    # ---- Form 컨트롤 버튼 배치 (재실행 대비: 동일 이름 기존 버튼 제거) ----
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    foreach ($nm in @("btn초기화","btn저장하기","btn수정하기","btn불러오기")) {
        try { $wsIn.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    $btnDefs = @(
        @{name="btn초기화"; caption="초기화"; macro="초기화"; left=520; top=10},
        @{name="btn저장하기"; caption="저장하기"; macro="저장하기"; left=610; top=10},
        @{name="btn수정하기"; caption="수정하기"; macro="수정하기"; left=700; top=10},
        @{name="btn불러오기"; caption="불러오기"; macro="불러오기"; left=790; top=10}
    )
    foreach ($b in $btnDefs) {
        $btn = $wsIn.Buttons().Add($b.left, $b.top, 80, 24)
        $btn.Caption = $b.caption
        $btn.OnAction = $b.macro
        $btn.Name = $b.name
    }
    L "기초자료입력 버튼 4개 배치 완료"

    $wsSel = $wb.Worksheets.Item("서식선택_출력")
    foreach ($nm in @("btn인쇄미리보기","btnPDF저장")) {
        try { $wsSel.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    $btn1 = $wsSel.Buttons().Add(300, 10, 160, 24)
    $btn1.Caption = "선택 서식 인쇄 미리보기"
    $btn1.OnAction = "선택서식_인쇄미리보기"
    $btn1.Name = "btn인쇄미리보기"

    $btn2 = $wsSel.Buttons().Add(470, 10, 160, 24)
    $btn2.Caption = "선택 서식 PDF 저장"
    $btn2.OnAction = "선택서식_PDF저장"
    $btn2.Name = "btnPDF저장"
    L "서식선택_출력 버튼 2개 배치 완료"

    $wsSearch = $wb.Worksheets.Item("학교검색")
    foreach ($nm in @("btn학교검색", "btn선택학교반영")) {
        try { $wsSearch.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    $btnSearch = $wsSearch.Buttons().Add(520, 35, 80, 24)
    $btnSearch.Caption = "검색"
    $btnSearch.OnAction = "학교검색_실행"
    $btnSearch.Name = "btn학교검색"
    $btnApplySchool = $wsSearch.Buttons().Add(610, 35, 190, 24)
    $btnApplySchool.Caption = "선택 학교를 기초자료에 반영"
    $btnApplySchool.OnAction = "선택학교_기초자료반영"
    $btnApplySchool.Name = "btn선택학교반영"
    L "학교검색 버튼 2개 배치 완료"

    $wsFlow = $wb.Worksheets.Item("절차안내")
    try { $wsFlow.Buttons("btn관련FormID이동").Delete() } catch { L "기존 버튼 btn관련FormID이동 없음" }
    $btnFlow = $wsFlow.Buttons().Add(650, 35, 150, 24)
    $btnFlow.Caption = "관련 Form ID로 이동"
    $btnFlow.OnAction = "관련FormID로이동"
    $btnFlow.Name = "btn관련FormID이동"
    L "절차안내 Form ID 이동 버튼 배치 완료"

    $wsMethod = $wb.Worksheets.Item("계약방법안내")
    try { $wsMethod.Buttons("btn계약방법반영").Delete() } catch { L "기존 버튼 btn계약방법반영 없음" }
    $btnMethod = $wsMethod.Buttons().Add(520, 58, 200, 24)
    $btnMethod.Caption = "계약방법을 기초자료에 반영"
    $btnMethod.OnAction = "계약방법안내_기초자료반영"
    $btnMethod.Name = "btn계약방법반영"
    L "계약방법안내 반영 버튼 배치 완료"

    # ---- 저장 ----
    # 2026-09-19 진단: Save()가 일반 COMException으로 실패해 SaveAs(같은 경로, 52)로
    # 바꿔봤으나, 이미 열려 있는 파일을 자기 자신에 SaveAs로 덮어쓰는 시도가 오히려 더
    # 구체적인 "저장할 수 없습니다" 잠금류 오류로 실패함(자기 자신을 여는 새 쓰기 핸들과
    # 기존 열림 핸들이 충돌하는 것으로 추정). Save()는 이 코드베이스 전체 이력에서 수백 회
    # 성공한 방식이므로 되돌리되, 이 프로젝트에서 반복 관찰된 일시적 COM/훅 레이스 현상에
    # 대비해 짧은 재시도(지수 백오프 + GC)를 추가한다.
    $saveAttempt = 0
    $saveLastError = $null
    while ($saveAttempt -lt 3 -and -not $saveSucceeded) {
        $saveAttempt++
        try {
            $wb.Save()
            $saveSucceeded = $true
        } catch {
            $saveLastError = $_
            L "저장 시도 $saveAttempt/3 실패: $($_.Exception.Message)"
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Start-Sleep -Seconds ([Math]::Pow(2, $saveAttempt))
        }
    }
    if (-not $saveSucceeded) { throw $saveLastError }
    L "VBA 매크로 및 버튼 추가 후 저장 완료(SaveAs): $targetPath"

} finally {
    # 정리 단계 자체의 COM 오류(예: 앞선 오류로 Excel 프로세스가 이미 비정상 상태가 된 경우의
    # Close() 실패)가 try 블록의 원래 예외를 가리거나 로그 flush·프로세스 강제 종료를 막지
    # 않도록 각 정리 호출을 개별적으로 감싼다.
    if ($wb) {
        try { $wb.Close($saveSucceeded) } catch { L "Workbook Close 실패(정리 계속 진행): $($_.Exception.Message)" }
    }
    try { $excel.Quit() } catch { L "Excel Quit 실패(정리 계속 진행): $($_.Exception.Message)" }
    if ($wb) {
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { L "Workbook ReleaseComObject 실패(정리 계속 진행): $($_.Exception.Message)" }
    }
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { L "Excel ReleaseComObject 실패(정리 계속 진행): $($_.Exception.Message)" }
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
    # try 블록에서 예외가 발생했더라도(원본 예외는 finally 종료 후 그대로 재전파됨) 여기까지
    # 수집된 로그는 항상 디스크에 남겨 원인 진단이 가능하게 한다.
    [System.IO.File]::WriteAllText($logPath, $log.ToString(), [System.Text.UTF8Encoding]::new($false))
}
