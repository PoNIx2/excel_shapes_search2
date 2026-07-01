Attribute VB_Name = "ShapeSearch"
Option Explicit

Private Type TSampleStyle
    HasFill As Boolean
    FillColor As Long
    HasLine As Boolean
    LineColor As Long
    LineDash As Long
End Type

Private Const SHEET_SETTINGS As String = "設定"
Private Const SHEET_SAMPLE As String = "判例"
Private Const SHEET_RESULT As String = "検索結果"

Private Const CELL_ROOT As String = "B2"
Private Const CELL_TEXT As String = "B3"
Private Const CELL_FLAG As String = "B4"

Public Sub SearchShapesInExcelFiles()
    On Error GoTo EH

    Dim wsSettings As Worksheet
    Dim wsSample As Worksheet
    Dim wsResult As Worksheet

    Set wsSettings = GetSheetOrError(ThisWorkbook, SHEET_SETTINGS)
    Set wsSample = GetSheetOrError(ThisWorkbook, SHEET_SAMPLE)
    Set wsResult = GetSheetOrError(ThisWorkbook, SHEET_RESULT)

    Dim rootPath As String
    Dim keyword As String
    Dim caseFlag As String

    rootPath = Trim$(CStr(wsSettings.Range(CELL_ROOT).Value))
    keyword = Trim$(CStr(wsSettings.Range(CELL_TEXT).Value))
    caseFlag = NormalizeFlag(Trim$(CStr(wsSettings.Range(CELL_FLAG).Value)))

    ValidateInputs rootPath, keyword, caseFlag, wsSample

    Dim sampleStyle As TSampleStyle
    Dim useSampleStyle As Boolean

    If caseFlag = "オン" Or caseFlag = "オフ" Then
        sampleStyle = GetSampleStyle(wsSample)
        useSampleStyle = True
    End If

    PrepareResultSheet wsResult

    Dim rowOut As Long
    rowOut = 2

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ScanFolderRecursive rootPath, keyword, caseFlag, useSampleStyle, sampleStyle, wsResult, rowOut

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "検索が完了しました。ヒット件数: " & (rowOut - 2), vbInformation
    Exit Sub

EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

Private Sub ValidateInputs(ByVal rootPath As String, ByVal keyword As String, ByVal caseFlag As String, ByVal wsSample As Worksheet)
    If Len(rootPath) = 0 Then
        Err.Raise vbObjectError + 1001, , "ルートパス(B2)は必須です。"
    End If

    If Dir(rootPath, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 1002, , "指定されたルートパスが存在しません: " & rootPath
    End If

    If caseFlag <> "オン" And caseFlag <> "オフ" And caseFlag <> "無効" And caseFlag <> vbNullString Then
        Err.Raise vbObjectError + 1003, , "判例フラグ(B4)は「オン」「オフ」「無効」のいずれかを指定してください。"
    End If

    If caseFlag = vbNullString Then caseFlag = "無効"

    If caseFlag = "無効" And Len(keyword) = 0 Then
        Err.Raise vbObjectError + 1004, , "判例フラグが無効の場合、図形文字(B3)は必須です。"
    End If

    If caseFlag = "オン" Or caseFlag = "オフ" Then
        If wsSample.Shapes.Count = 0 Then
            Err.Raise vbObjectError + 1005, , "判例フラグがオン/オフの場合、判例シートに図形が必要です。"
        End If
    End If
End Sub

Private Function NormalizeFlag(ByVal s As String) As String
    Select Case s
        Case "ON", "On", "on", "オン"
            NormalizeFlag = "オン"
        Case "OFF", "Off", "off", "オフ"
            NormalizeFlag = "オフ"
        Case "無効", "MUKOU", "mukou", "none", "NONE", ""
            NormalizeFlag = "無効"
        Case Else
            NormalizeFlag = s
    End Select
End Function

Private Sub PrepareResultSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "ファイル名"
    ws.Range("B1").Value = "シート名"
    ws.Range("C1").Value = "図形の文字"
    ws.Rows(1).Font.Bold = True
End Sub

Private Sub ScanFolderRecursive(ByVal folderPath As String, ByVal keyword As String, ByVal caseFlag As String, _
                                ByVal useSampleStyle As Boolean, ByRef sampleStyle As TSampleStyle, _
                                ByVal wsResult As Worksheet, ByRef rowOut As Long)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If IsExcelFile(CStr(file.Path)) Then
            If LCase$(CStr(file.Path)) <> LCase$(ThisWorkbook.FullName) Then
                ScanWorkbook CStr(file.Path), keyword, caseFlag, useSampleStyle, sampleStyle, wsResult, rowOut
            End If
        End If
    Next file

    For Each subFolder In folder.SubFolders
        ScanFolderRecursive CStr(subFolder.Path), keyword, caseFlag, useSampleStyle, sampleStyle, wsResult, rowOut
    Next subFolder
End Sub

Private Sub ScanWorkbook(ByVal wbPath As String, ByVal keyword As String, ByVal caseFlag As String, _
                         ByVal useSampleStyle As Boolean, ByRef sampleStyle As TSampleStyle, _
                         ByVal wsResult As Worksheet, ByRef rowOut As Long)
    On Error GoTo SAFE_EXIT

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim shpText As String

    Set wb = Workbooks.Open(Filename:=wbPath, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)

    For Each ws In wb.Worksheets
        For Each shp In ws.Shapes
            shpText = GetShapeText(shp)
            If ShapeMatches(shp, shpText, keyword, caseFlag, useSampleStyle, sampleStyle) Then
                wsResult.Cells(rowOut, 1).Value = wb.Name
                wsResult.Cells(rowOut, 2).Value = ws.Name
                wsResult.Cells(rowOut, 3).Value = shpText
                rowOut = rowOut + 1
            End If
        Next shp
    Next ws

SAFE_EXIT:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function ShapeMatches(ByVal shp As Shape, ByVal shpText As String, ByVal keyword As String, ByVal caseFlag As String, _
                              ByVal useSampleStyle As Boolean, ByRef sampleStyle As TSampleStyle) As Boolean
    Dim textOk As Boolean
    If Len(keyword) = 0 Then
        textOk = True
    Else
        textOk = (InStr(1, shpText, keyword, vbTextCompare) > 0)
    End If

    If Not textOk Then
        ShapeMatches = False
        Exit Function
    End If

    If Not useSampleStyle Then
        ShapeMatches = True
        Exit Function
    End If

    Dim currentStyle As TSampleStyle
    currentStyle = GetShapeStyle(shp)

    Dim styleEq As Boolean
    styleEq = CompareStyle(sampleStyle, currentStyle)

    Select Case caseFlag
        Case "オン"
            ShapeMatches = styleEq
        Case "オフ"
            ShapeMatches = Not styleEq
        Case Else
            ShapeMatches = textOk
    End Select
End Function

Private Function GetSampleStyle(ByVal wsSample As Worksheet) As TSampleStyle
    Dim shp As Shape
    Set shp = wsSample.Shapes(1)
    GetSampleStyle = GetShapeStyle(shp)
End Function

Private Function GetShapeStyle(ByVal shp As Shape) As TSampleStyle
    Dim st As TSampleStyle

    On Error Resume Next
    st.HasFill = (shp.Fill.Visible <> msoFalse)
    If st.HasFill Then st.FillColor = shp.Fill.ForeColor.RGB

    st.HasLine = (shp.Line.Visible <> msoFalse)
    If st.HasLine Then
        st.LineColor = shp.Line.ForeColor.RGB
        st.LineDash = shp.Line.DashStyle
    End If
    On Error GoTo 0

    GetShapeStyle = st
End Function

Private Function CompareStyle(ByRef a As TSampleStyle, ByRef b As TSampleStyle) As Boolean
    If a.HasFill <> b.HasFill Then Exit Function
    If a.HasLine <> b.HasLine Then Exit Function

    If a.HasFill Then
        If a.FillColor <> b.FillColor Then Exit Function
    End If

    If a.HasLine Then
        If a.LineColor <> b.LineColor Then Exit Function
        If a.LineDash <> b.LineDash Then Exit Function
    End If

    CompareStyle = True
End Function

Private Function GetShapeText(ByVal shp As Shape) As String
    On Error Resume Next
    If shp.TextFrame2.HasText Then
        GetShapeText = CStr(shp.TextFrame2.TextRange.Text)
        Exit Function
    End If

    If shp.TextFrame.HasText Then
        GetShapeText = CStr(shp.TextFrame.Characters.Text)
        Exit Function
    End If

    GetShapeText = ""
End Function

Private Function IsExcelFile(ByVal filePath As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    IsExcelFile = (ext = "xlsx" Or ext = "xlsm" Or ext = "xls")
End Function

Private Function GetSheetOrError(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error GoTo EH
    Set GetSheetOrError = wb.Worksheets(sheetName)
    Exit Function
EH:
    Err.Raise vbObjectError + 1010, , "シートが見つかりません: " & sheetName
End Function
