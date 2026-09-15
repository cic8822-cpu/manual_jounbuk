# P3-01 클린룸 재구현 — VBA 빌더(2/2): 기초자료 신규/저장/수정/불러오기, 서식선택 출력(PDF) 매크로 추가
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$targetPath = Join-Path $root 'artifacts\excel\교복구매_길라잡이_Excel_v1.xlsm'
$logPath = Join-Path $root '_workspace\03_excel\build_vba_log.txt'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$log = New-Object System.Text.StringBuilder
function L($s) { [void]$log.AppendLine($s) }

$wb = $null
try {
    $wb = $excel.Workbooks.Open($targetPath, [Type]::Missing, $false)  # ReadOnly=False (편집)

    # ---- 레코드 추적 셀 추가 (기초자료입력!G1/H1) ----
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    $wsIn.Range("G1").Value2 = "불러온 레코드 순번"
    $wsIn.Range("G1").Font.Size = 8
    $wsIn.Range("H1").Value2 = ""
    L "기초자료입력!G1/H1 레코드 추적 셀 추가 완료"

    # ---- VBA 모듈 추가 (재실행 대비: 동일 이름 기존 모듈 제거 후 추가) ----
    $vbproj = $wb.VBProject
    foreach ($nm in @("Module_기초자료","Module_출력")) {
        try {
            $existing = $vbproj.VBComponents.Item($nm)
            $vbproj.VBComponents.Remove($existing)
            L "기존 모듈 $nm 제거 후 재생성"
        } catch { }
    }

    $modInput = $vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    $modInput.Name = "Module_기초자료"
    $codeInput = @'
Option Explicit

Sub 초기화()
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
        .Range("H1").Value = ""
    End With
    MsgBox "기초자료가 초기화되었습니다.", vbInformation
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

Sub 저장하기()
    If Not 필수값검증() Then Exit Sub
    Dim wsIn As Worksheet, wsDB As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")

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

    wsIn.Range("H1").Value = seq
    MsgBox "저장되었습니다. (순번 " & seq & ")", vbInformation
End Sub

Sub 수정하기()
    If Not 필수값검증() Then Exit Sub
    Dim wsIn As Worksheet, wsDB As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")

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
    MsgBox "수정되었습니다. (순번 " & seq & ")", vbInformation
End Sub

Sub 불러오기()
    Dim seqStr As String
    seqStr = InputBox("불러올 레코드의 순번을 입력하세요. (DB 시트 A열 참고)", "레코드 불러오기")
    If seqStr = "" Then Exit Sub
    If Not IsNumeric(seqStr) Then
        MsgBox "숫자를 입력하세요.", vbExclamation
        Exit Sub
    End If
    Dim wsIn As Worksheet, wsDB As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")
    Dim r As Long
    r = CLng(seqStr) + 1
    If wsDB.Cells(r, 1).Value <> CLng(seqStr) Then
        MsgBox "해당 순번의 레코드를 찾을 수 없습니다.", vbCritical
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
    wsIn.Range("H1").Value = CLng(seqStr)
    MsgBox "불러왔습니다. (순번 " & seqStr & ")", vbInformation
End Sub
'@

    $codeInput = $codeInput -replace "`r`n", "`r" -replace "`n", "`r"
    $modInput.CodeModule.AddFromString($codeInput)
    L "Module_기초자료 추가 완료 (줄 수: $($modInput.CodeModule.CountOfLines))"

    $modOutput = $vbproj.VBComponents.Add(1)
    $modOutput.Name = "Module_출력"
    $codeOutput = @'
Option Explicit

Private Function 선택된시트목록(ByRef outNames() As String) As Long
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim lastRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row

    Dim tmp() As String
    ReDim tmp(1 To lastRow)
    Dim cnt As Long
    cnt = 0
    Dim skippedNotReady As Long, skippedDeferred As Long
    skippedNotReady = 0
    skippedDeferred = 0

    Dim r As Long
    For r = 5 To lastRow
        If wsSel.Cells(r, 1).Value = True Then
            Dim status As String
            status = wsSel.Cells(r, 5).Value
            If status = "Y" Then
                cnt = cnt + 1
                tmp(cnt) = wsSel.Cells(r, 6).Value
            ElseIf status = "D" Then
                skippedDeferred = skippedDeferred + 1
            Else
                skippedNotReady = skippedNotReady + 1
            End If
        End If
    Next r

    If cnt = 0 Then
        MsgBox "구현된 서식이 선택되지 않았습니다." & vbCrLf & _
               "미구현: " & skippedNotReady & "건, HWPX 우선순위 위임: " & skippedDeferred & "건", vbExclamation
        선택된시트목록 = 0
        Exit Function
    End If

    If skippedNotReady > 0 Or skippedDeferred > 0 Then
        MsgBox "일부 선택 서식은 아직 준비되지 않아 제외합니다." & vbCrLf & _
               "미구현: " & skippedNotReady & "건, HWPX 우선순위 위임: " & skippedDeferred & "건", vbInformation
    End If

    ReDim outNames(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        outNames(i) = tmp(i)
    Next i
    선택된시트목록 = cnt
End Function

Sub 선택서식_인쇄미리보기()
    Dim names() As String
    Dim n As Long
    n = 선택된시트목록(names)
    If n = 0 Then Exit Sub
    ThisWorkbook.Sheets(names).Select
    ActiveWindow.SelectedSheets.PrintPreview
End Sub

Sub 선택서식_PDF저장()
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

    MsgBox "PDF로 저장되었습니다:" & vbCrLf & fileName, vbInformation
End Sub
'@
    $codeOutput = $codeOutput -replace "`r`n", "`r" -replace "`n", "`r"
    $modOutput.CodeModule.AddFromString($codeOutput)
    L "Module_출력 추가 완료 (줄 수: $($modOutput.CodeModule.CountOfLines))"

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
    } catch { }

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
End Sub
'@
    $codeThisWb = $codeThisWb -replace "`r`n", "`r" -replace "`n", "`r"
    $thisWb.CodeModule.AddFromString($codeThisWb)
    L "ThisWorkbook.Workbook_Open 추가 완료"

    # ---- Form 컨트롤 버튼 배치 (재실행 대비: 동일 이름 기존 버튼 제거) ----
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    foreach ($nm in @("btn초기화","btn저장하기","btn수정하기","btn불러오기")) {
        try { $wsIn.Buttons($nm).Delete() } catch { }
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
        try { $wsSel.Buttons($nm).Delete() } catch { }
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

    # ---- 저장 ----
    $wb.Save()
    L "VBA 매크로 및 버튼 추가 후 저장 완료: $targetPath"

} finally {
    if ($wb) { $wb.Close($true) }
    $excel.Quit()
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

[System.IO.File]::WriteAllText($logPath, $log.ToString(), [System.Text.UTF8Encoding]::new($false))
