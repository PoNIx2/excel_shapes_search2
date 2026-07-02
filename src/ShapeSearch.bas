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

' 設定シートの条件をもとにフォルダ配下のExcelファイルを再帰的に走査し、
' 図形の文字列またはスタイル一致条件に合う図形を検索結果シートへ出力するメイン処理
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

    Dim sampleStyles() As TSampleStyle

    If caseFlag = "オン" Or caseFlag = "オフ" Then
        sampleStyles = GetSampleStyles(wsSample)
    End If

    PrepareResultSheet wsResult

    Dim rowOut As Long
    rowOut = 2

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ScanFolderRecursive rootPath, keyword, caseFlag, sampleStyles, wsResult, rowOut

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "検索が完了しました。ヒット件数: " & (rowOut - 2), vbInformation
    Exit Sub

EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

' 入力値（ルートパス・判例フラグ・必須条件）を検証し、
' 条件不備があれば原因別のエラーを送出する
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

' 判例フラグの入力ゆれ（ON/OFF/空文字など）を正規化して
' 「オン」「オフ」「無効」のいずれかへ変換する
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

' 検索結果シートを初期化し、ヘッダー行を設定する
Private Sub PrepareResultSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "ファイル名"
    ws.Range("B1").Value = "シート名"
    ws.Range("C1").Value = "図形の文字"
    ws.Rows(1).Font.Bold = True
End Sub

' 指定フォルダ以下を再帰的にたどり、
' Excelファイルを見つけるたびにブック単位の走査処理を実行する
Private Sub ScanFolderRecursive(ByVal folderPath As String, ByVal keyword As String, ByVal caseFlag As String, _
                                ByRef sampleStyles() As TSampleStyle, _
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
                ScanWorkbook CStr(file.Path), keyword, caseFlag, sampleStyles, wsResult, rowOut
            End If
        End If
    Next file

    For Each subFolder In folder.SubFolders
        ScanFolderRecursive CStr(subFolder.Path), keyword, caseFlag, sampleStyles, wsResult, rowOut
    Next subFolder
End Sub

' 1つのExcelブックを開き、全シート・全図形を走査して
' 条件に一致した図形情報を検索結果シートへ1行ずつ追加する
Private Sub ScanWorkbook(ByVal wbPath As String, ByVal keyword As String, ByVal caseFlag As String, _
                         ByRef sampleStyles() As TSampleStyle, _
                         ByVal wsResult As Worksheet, ByRef rowOut As Long)
    On Error GoTo SAFE_EXIT

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim shpText As String

    Set wb = Workbooks.Open(Filename:=wbPath, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)

    For Each ws In wb.Worksheets
        ' 非表示シート（Hidden / VeryHidden）は検索対象外とする
        If ws.Visible = xlSheetVisible Then
            For Each shp In ws.Shapes
                shpText = GetShapeText(shp)
                If ShapeMatches(shp, shpText, keyword, caseFlag, sampleStyles) Then
                    wsResult.Cells(rowOut, 1).Value = wb.Name
                    wsResult.Cells(rowOut, 2).Value = ws.Name
                    wsResult.Cells(rowOut, 3).Value = shpText
                    rowOut = rowOut + 1
                End If
            Next shp
        End If
    Next ws

SAFE_EXIT:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' 図形が検索条件に一致するかを判定する
' （判例フラグがオン/オフならスタイル一致判定、無効なら文字列部分一致判定）
Private Function ShapeMatches(ByVal shp As Shape, ByVal shpText As String, ByVal keyword As String, ByVal caseFlag As String, _
                              ByRef sampleStyles() As TSampleStyle) As Boolean
    If caseFlag = "オン" Or caseFlag = "オフ" Then
        Dim currentStyle As TSampleStyle
        Dim styleEq As Boolean

        currentStyle = GetShapeStyle(shp)
        styleEq = MatchesAnySampleStyle(sampleStyles, currentStyle)

        Select Case caseFlag
            Case "オン"
                ShapeMatches = styleEq
            Case "オフ"
                ShapeMatches = Not styleEq
        End Select
        Exit Function
    End If

    If Len(keyword) = 0 Then
        ShapeMatches = True
    Else
        ShapeMatches = (InStr(1, shpText, keyword, vbTextCompare) > 0)
    End If
End Function

' 判例シートの全図形をサンプルとして取得し、
' 比較用のスタイル情報（塗り・線）の配列を返す
Private Function GetSampleStyles(ByVal wsSample As Worksheet) As TSampleStyle()
    Dim styles() As TSampleStyle
    Dim i As Long

    ReDim styles(1 To wsSample.Shapes.Count)

    For i = 1 To wsSample.Shapes.Count
        styles(i) = GetShapeStyle(wsSample.Shapes(i))
    Next i

    GetSampleStyles = styles
End Function

' 候補図形のスタイルが判例シート上のいずれかのサンプルスタイルに一致するかを判定する
Private Function MatchesAnySampleStyle(ByRef sampleStyles() As TSampleStyle, ByRef currentStyle As TSampleStyle) As Boolean
    Dim i As Long

    For i = LBound(sampleStyles) To UBound(sampleStyles)
        If CompareStyle(sampleStyles(i), currentStyle) Then
            MatchesAnySampleStyle = True
            Exit Function
        End If
    Next i
End Function

' 図形からスタイル情報（塗り有無/色、線有無/色/破線種別）を抽出して返す
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

' 2つのスタイル情報を比較し、
' 塗り・線の有無と各属性がすべて一致する場合のみTrueを返す
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

' 図形内の文字列を取得する
' （TextFrame2を優先し、取得できなければTextFrameを参照）
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

' ファイル拡張子からExcelファイル（xls/xlsx/xlsm）かどうかを判定する
Private Function IsExcelFile(ByVal filePath As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    IsExcelFile = (ext = "xlsx" Or ext = "xlsm" Or ext = "xls")
End Function

' 指定名のシートをブックから取得し、存在しない場合はエラーを送出する
Private Function GetSheetOrError(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error GoTo EH
    Set GetSheetOrError = wb.Worksheets(sheetName)
    Exit Function
EH:
    Err.Raise vbObjectError + 1010, , "シートが見つかりません: " & sheetName
End Function
