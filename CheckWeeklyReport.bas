Option Explicit

Sub CheckAllXlsxFiles_MultiPairs()
    Dim folderPath As String
    Dim fileList As Collection
    Dim filePath As Variant
    Dim wsOut As Worksheet
    Dim nextRow As Long

    ' 定义要对比的单元格对（左边和右边一一对应）
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
        MsgBox "未找到xlsx文件", vbExclamation
        Exit Sub
    End If

    Dim wb As Workbook, ws As Worksheet
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    Dim i As Long, minuteForCompare As Long
    Dim allOK As Boolean, isMinuteChange As Boolean
    Dim diffInfo As String
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

    r = 1
    For Each filePath In fileList
        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=CStr(filePath), ReadOnly:=True, _
                                UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, _
                                AddToMru:=False, Notify:=False)
        On Error GoTo SafeExit

        If Not wb Is Nothing Then
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

                If v2 <> 0 Then
                    minuteForCompare = v2
                End If
            Next i

            resultData(r, 1) = filePath
            If allOK Then
                If Not isMinuteChange Then
                    resultData(r, 2) = "Warning"
                    resultData(r, 3) = "作業時間は毎日同じです。事実ですか。"
                Else
                    resultData(r, 2) = "OK"
                    resultData(r, 3) = vbNullString
                End If
            Else
                resultData(r, 2) = "NG"
                resultData(r, 3) = diffInfo
            End If

            r = r + 1
            wb.Close SaveChanges:=False
        End If
    Next filePath

    If r > 1 Then
        wsOut.Cells(nextRow, 1).Resize(r - 1, 3).Value = resultData
    End If

SafeExit:
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If

    Application.AutomationSecurity = oldSecurity
    Application.Calculation = oldCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "执行异常: " & Err.Description, vbExclamation
    Else
        MsgBox "チェック完了", vbInformation
    End If
End Sub

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

'=== 递归取所有xlsx文件 ===
Sub GetAllXlsxFiles(ByVal folderPath As String, ByRef fileList As Collection)
    Dim fileName As String
    Dim subFolder As String

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    fileName = Dir$(folderPath & "*.xlsx", vbNormal)
    Do While LenB(fileName) > 0
        fileList.Add folderPath & fileName
        fileName = Dir$()
    Loop

    subFolder = Dir$(folderPath & "*", vbDirectory)
    Do While LenB(subFolder) > 0
        If subFolder <> "." And subFolder <> ".." Then
            If (GetAttr(folderPath & subFolder) And vbDirectory) = vbDirectory Then
                GetAllXlsxFiles folderPath & subFolder, fileList
            End If
        End If
        subFolder = Dir$()
    Loop
End Sub
