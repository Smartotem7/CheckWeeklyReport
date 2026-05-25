Option Explicit

Sub CheckAllXlsxFiles_MultiPairs()
    Dim folderPath As String
    Dim fileList As Collection
    Dim filePath As Variant
    Dim wsOut As Worksheet
    Dim nextRow As Long

    ' 比較対象のセルペアを定義（左側と右側を1対1で対応）
    Dim cellPairs As Variant
    cellPairs = Array( _
        Array("AW24", "BA25", "B22", "月曜日"), _
        Array("AW34", "BA35", "B32", "火曜日"), _
        Array("AW44", "BA45", "B42", "水曜日"), _
        Array("AW54", "BA55", "B52", "木曜日"), _
        Array("AW64", "BA65", "B62", "金曜日"), _
        Array("AW74", "BA75", "B72", "土曜日"), _
        Array("AW84", "BA85", "B82", "日曜日"))

    Set wsOut = ThisWorkbook.Sheets(1)
    wsOut.Range("A4:D200").ClearContents
    nextRow = 4

    folderPath = wsOut.Range("A2").Value

    Set fileList = New Collection
    GetAllXlsxFiles folderPath, fileList

    If fileList.Count = 0 Then
        MsgBox "xlsxファイルが見つかりませんでした", vbExclamation
        Exit Sub
    End If

    Dim oldCalc As XlCalculation
    Dim oldSecurity As MsoAutomationSecurity
    Dim resultData() As Variant
    Dim r As Long

    ReDim resultData(1 To fileList.Count, 1 To 3)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    oldSecurity = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    On Error GoTo SafeExit

    ' 各xlsxファイルを順番に検査
    r = 1
    For Each filePath In fileList
        EvaluateFile CStr(filePath), cellPairs, resultData(r, 2), resultData(r, 3)
        resultData(r, 1) = filePath
        r = r + 1
    Next filePath

    If r > 1 Then
        wsOut.Cells(nextRow, 1).Resize(r - 1, 3).Value = resultData
    End If

SafeExit:
    Application.AutomationSecurity = oldSecurity
    Application.Calculation = oldCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "実行時エラー: " & Err.Description, vbExclamation
    Else
        MsgBox "チェック完了", vbInformation
    End If
End Sub

Public Sub CheckXlsxFiles_FromList(ByVal listFilePath As String, ByVal outputCsvPath As String)
    Dim files As Collection
    Dim fso As Object
    Dim ts As Object
    Dim onePath As String
    Dim cellPairs As Variant
    Dim resultData() As Variant
    Dim i As Long

    cellPairs = Array( _
        Array("AW24", "BA25", "B22", "月曜日"), _
        Array("AW34", "BA35", "B32", "火曜日"), _
        Array("AW44", "BA45", "B42", "水曜日"), _
        Array("AW54", "BA55", "B52", "木曜日"), _
        Array("AW64", "BA65", "B62", "金曜日"), _
        Array("AW74", "BA75", "B72", "土曜日"), _
        Array("AW84", "BA85", "B82", "日曜日"))

    Set files = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(listFilePath) Then Exit Sub

    Set ts = fso.OpenTextFile(listFilePath, 1)
    Do While Not ts.AtEndOfStream
        onePath = Trim$(ts.ReadLine)
        If LenB(onePath) > 0 Then files.Add onePath
    Loop
    ts.Close

    If files.Count = 0 Then Exit Sub

    ReDim resultData(1 To files.Count, 1 To 3)
    For i = 1 To files.Count
        resultData(i, 1) = files(i)
        EvaluateFile CStr(files(i)), cellPairs, resultData(i, 2), resultData(i, 3)
    Next i

    Set ts = fso.CreateTextFile(outputCsvPath, True)
    For i = 1 To UBound(resultData, 1)
        ts.WriteLine CsvEscape(CStr(resultData(i, 1))) & "," & _
                     CsvEscape(CStr(resultData(i, 2))) & "," & _
                     CsvEscape(CStr(resultData(i, 3)))
    Next i
    ts.Close
End Sub

Private Sub EvaluateFile(ByVal filePath As String, ByVal cellPairs As Variant, ByRef resultStatus As Variant, ByRef resultMessage As Variant)
    Dim wb As Workbook, ws As Worksheet
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    Dim i As Long, minuteForCompare As Long
    Dim allOK As Boolean, isMinuteChange As Boolean
    Dim diffInfo As String

    resultStatus = "NG"
    resultMessage = "ファイルを開けませんでした。"

    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=CStr(filePath), ReadOnly:=True, _
                            UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, _
                            AddToMru:=False, Notify:=False)
    On Error GoTo 0

    If wb Is Nothing Then Exit Sub

    Set ws = wb.Sheets(1)
    allOK = True
    diffInfo = vbNullString
    minuteForCompare = 0
    isMinuteChange = False

    For i = LBound(cellPairs) To UBound(cellPairs)
        v1 = ToLongOrZero(ws.Range(cellPairs(i)(0)).Value2)
        v2 = Round(ToDoubleOrZero(ws.Range(cellPairs(i)(1)).Value2) * 24 * 60, 0)
        v3 = ws.Range(cellPairs(i)(2)).Value2
        v4 = cellPairs(i)(3)

        If v1 <> v2 Then
            allOK = False
            diffInfo = diffInfo & cellPairs(i)(0) & "=" & v1 & "；" & _
                               cellPairs(i)(1) & "=" & v2 & "；"
        End If

        If i < 5 And v1 = 0 And IsEmpty(v3) Then
            allOK = False
            diffInfo = diffInfo & v4 & "が休みの場合、" & cellPairs(i)(2) & _
                               "に「休み or 祝日」を入れてください。"
            Exit For
        End If

        If minuteForCompare <> 0 And v2 <> 0 And minuteForCompare <> v2 Then
            isMinuteChange = True
        End If

        If v2 <> 0 Then minuteForCompare = v2
    Next i

    If allOK Then
        If Not isMinuteChange Then
            resultStatus = "Warning"
            resultMessage = "作業時間は毎日同じです。事実ですか。"
        Else
            resultStatus = "OK"
            resultMessage = vbNullString
        End If
    Else
        resultStatus = "NG"
        resultMessage = diffInfo
    End If

    wb.Close SaveChanges:=False
End Sub

Private Function CsvEscape(ByVal s As String) As String
    CsvEscape = """" & Replace(s, """", """"") & """"
End Function

Private Function ToLongOrZero(ByVal v As Variant) As Long
    If IsError(v) Or IsEmpty(v) Or LenB(vbNullString & v) = 0 Or v = 0 Then
        ToLongOrZero = 0
    Else
        ToLongOrZero = CLng(v)
    End If
End Function

Private Function ToDoubleOrZero(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or LenB(vbNullString & v) = 0 Or v = 0 Then
        ToDoubleOrZero = 0
    Else
        ToDoubleOrZero = CDbl(v)
    End If
End Function

'=== すべてのxlsxファイルを再帰的に取得 ===
Sub GetAllXlsxFiles(ByVal folderPath As String, ByRef fileList As Collection)
    Dim fso As Object
    Dim rootFolder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then Exit Sub

    Set rootFolder = fso.GetFolder(folderPath)
    CollectXlsxFiles rootFolder, fileList
End Sub

Private Sub CollectXlsxFiles(ByVal currentFolder As Object, ByRef fileList As Collection)
    Dim f As Object
    Dim subFolder As Object

    For Each f In currentFolder.Files
        If LCase$(Right$(f.Name, 5)) = ".xlsx" Then
            fileList.Add f.Path
        End If
    Next f

    For Each subFolder In currentFolder.SubFolders
        CollectXlsxFiles subFolder, fileList
    Next subFolder
End Sub
