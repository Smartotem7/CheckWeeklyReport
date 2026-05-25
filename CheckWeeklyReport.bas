Option Explicit

Sub CheckAllXlsxFiles_MultiPairs()
    Dim folderPath As String
    Dim fileList As Collection
    Dim filePath As Variant
    Dim wsOut As Worksheet
    Dim nextRow As Long
    
    ' 定?要?比的?元格?（左?和右?一一??）
    Dim cellPairs As Variant
    cellPairs = Array( _
        Array("AW24", "BA25", "B22", "月曜日"), _
        Array("AW34", "BA35", "B32", "火曜日"), _
        Array("AW44", "BA45", "B42", "水曜日"), _
        Array("AW54", "BA55", "B52", "木曜日"), _
        Array("AW64", "BA65", "B62", "金曜日"), _
        Array("AW74", "BA75", "B72", "土曜日"), _
        Array("AW84", "BA85", "B82", "日曜日"))  ' ← 可根据需要修改
    
    ' ======================
    
    Set wsOut = ThisWorkbook.Sheets(1)
    ' === 新增：清空上次?果 ===
    wsOut.Range("A4:D200").ClearContents
    'wsOut.Range("A3:D3").Value = Array("ファイル名", "結果", "セル", "値")
    nextRow = 4
    
    ' ==== 用??置区域 ====
    folderPath = wsOut.Range("A2").Value
    
    Set fileList = New Collection
    Call GetAllXlsxFiles(folderPath, fileList)
    
    Dim wb As Workbook
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    Dim i As Integer, minuteForCompare As Integer
    Dim allOK As Boolean, isMinuteChange As Boolean
    Dim diffInfo As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    For Each filePath In fileList
        On Error Resume Next
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, _
                                UpdateLinks:=False, AddToMru:=False)
        wb.Windows(1).Visible = False
        On Error GoTo 0
        
        If Not wb Is Nothing Then
            allOK = True
            diffInfo = ""
            minuteForCompare = 0
            isMinuteChange = False
            
            With wb.Sheets(1)
                For i = LBound(cellPairs) To UBound(cellPairs)
                    v1 = .Range(cellPairs(i)(0)).Value
                    v2 = Round(.Range(cellPairs(i)(1)).Value * 24 * 60, 0)
                    v3 = .Range(cellPairs(i)(2)).Value
                    v4 = cellPairs(i)(3)
                    
                    If IsEmpty(v1) Or v1 = 0 Then v1 = 0
                    If IsEmpty(v2) Or v2 = 0 Then v2 = 0
                    
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
            End With
            
            wsOut.Cells(nextRow, 1).Value = filePath
            If allOK Then
                wsOut.Cells(nextRow, 2).Value = "OK"
                If Not isMinuteChange Then
                    wsOut.Cells(nextRow, 2).Value = "Warning"
                    wsOut.Cells(nextRow, 3).Value = "作業時間は毎日同じです。事実ですか。"
                End If
            Else
                wsOut.Cells(nextRow, 2).Value = "NG"
                wsOut.Cells(nextRow, 3).Value = diffInfo
            End If
            
            nextRow = nextRow + 1
            wb.Close SaveChanges:=False
        End If
    Next filePath
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "チェック完了", vbInformation
End Sub


'=== ???取所有xlsx文件 ===
Sub GetAllXlsxFiles(ByVal folderPath As String, ByRef fileList As Collection)
    Dim fso As Object, folder As Object, subFolder As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
            fileList.Add file.Path
        End If
    Next file
    
    For Each subFolder In folder.SubFolders
        Call GetAllXlsxFiles(subFolder.Path, fileList)
    Next subFolder
End Sub


