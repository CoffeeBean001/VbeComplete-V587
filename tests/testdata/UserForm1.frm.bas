Option Explicit

Private Sub CommandButton1_Click()
    Dim sourceFolderPath$, targetFolderPath$
    sourceFolderPath = Label2.Caption
    targetFolderPath = Label4.Caption

    If sourceFolderPath = "" Then
        MsgBox "请选择数据源文件夹！", vbExclamation
        Exit Sub
    End If
    If targetFolderPath = "" Then
        MsgBox "请选择结果文件夹！", vbExclamation
        Exit Sub
    End If

    Application.DisplayAlerts = False
    Dim flagStr$, flagDic As Object, key$, i&, flagArr
    flagStr = TextBox1.Text
    Set flagDic = CreateObject("Scripting.Dictionary")
    If flagStr <> "" Then
        flagStr = UCase(flagStr)
        flagStr = Replace(flagStr, "，", ",")
        flagArr = Split(flagStr, ",")
        For i = 0 To UBound(flagArr)
            key = flagArr(i)
            flagDic(key) = ""
        Next
    End If

    Dim fs As Object, fd As Object, f As Object
    Dim wb As Workbook, ws As Worksheet, sortedDic As New Dictionary
    Dim fileKind$, num$, cellAddress$, oneCellAddress$, cellAddressArr, j&, m&
    Dim newBook As Workbook, newSheet As Worksheet
    Dim fileKindArr, itemsArr, numArr, subItemsArr
    Dim pr&, pc&, arr, targetSavePath$

    Set fs = CreateObject("Scripting.FileSystemObject")
    For Each fd In fs.GetFolder(sourceFolderPath).SubFolders
        sortedDic.RemoveAll
        sortedDic.Add "绿色一", New Dictionary
        sortedDic.Add "绿色二", New Dictionary
        sortedDic.Add "绿色三", New Dictionary
        sortedDic.Add "绿色四", New Dictionary
        sortedDic.Add "绿色五", New Dictionary
        sortedDic.Add "绿色六", New Dictionary
        sortedDic.Add "红色", New Dictionary
        sortedDic.Add "绿色", New Dictionary
        sortedDic.Add "所有", New Dictionary
        '汇总所有得，去重复
        For Each f In fd.Files
            fileKind = getFileKind(f.Name)
            If sortedDic.Exists(fileKind) Then
                Set wb = Workbooks.Open(f.Path)
                Set ws = wb.Worksheets(1)
                For i = 1 To ws.Cells(ws.Cells.Rows.Count, "a").End(xlUp).Row
                    num = ws.Cells(i, "a")
                    cellAddress = ws.Cells(i, "b")
                    If Not sortedDic.Item(fileKind).Exists(num) Then sortedDic.Item(fileKind).Add num, New Dictionary
                    cellAddressArr = Split(cellAddress, Chr(10))
                    For j = 0 To UBound(cellAddressArr)
                        If InStr(1, cellAddressArr(j), "    ", vbTextCompare) > 0 And _
                           InStr(1, cellAddressArr(j), ",", vbTextCompare) > 0 Then
                            oneCellAddress = Split(Split(cellAddressArr(j), "    ")(1), ",")(0)
                            If flagDic.Exists(oneCellAddress) Then
                                oneCellAddress = oneCellAddress & "@@@@"
                            End If
                            If Not sortedDic.Item(fileKind).Item(num).Exists(oneCellAddress) Then
                                sortedDic.Item(fileKind).Item(num).Add oneCellAddress, ""
                            End If
                        End If
                    Next
                Next
                wb.Close False
            Else
                MsgBox "未知文件名称：" & fileKind, vbExclamation
                Exit Sub
            End If
        Next
        '填充结果
        arr = Split(fd.Path, "\")
        arr = Split(arr(UBound(arr)), ".")
        If Dir(targetFolderPath & "\" & arr(1), vbDirectory) = "" Then
            MkDir targetFolderPath & "\" & arr(1)
        End If
        targetSavePath = targetFolderPath & arr(1) & "\" & fd.Name & ".xlsx"
        If Dir(targetSavePath, vbNormal) = "" Then
            Set newBook = Workbooks.Add()
            Set newSheet = newBook.Worksheets(1)
        Else
            Set newBook = Workbooks.Open(targetSavePath)
            Set newSheet = newBook.Worksheets(1)
            newSheet.Cells.Clear
        End If
        
        fileKindArr = sortedDic.Keys
        itemsArr = sortedDic.Items
        pc = 1
        For i = 0 To UBound(fileKindArr)
            pr = 2
            newSheet.Cells(1, pc) = fd.Name
            newSheet.Cells(1, pc + 1) = fileKindArr(i)
            If itemsArr(i).Count > 0 Then
                numArr = itemsArr(i).Keys
                subItemsArr = itemsArr(i).Items
                For j = 0 To UBound(numArr)
                    cellAddressArr = subItemsArr(j).Keys
                    For m = 0 To UBound(cellAddressArr)
                        newSheet.Cells(pr, pc) = numArr(j)
                        newSheet.Cells(pr, pc + 1) = cellAddressArr(m)
                        pr = pr + 1
                    Next
                Next
            End If
            pc = pc + 2
        Next

        newBook.SaveAs targetSavePath, xlWorkbookDefault
        newBook.Close False
    Next

    Application.DisplayAlerts = True
    Unload Me
    MsgBox "OK", vbInformation
End Sub

Private Sub Label2_Click()
    Label2.Caption = pub2.getFolderPath()
End Sub

Private Sub Label4_Click()
    Label4.Caption = pub2.getFolderPath()
End Sub

Private Function getFileKind(ByVal fileName As String) As String
    Dim result$
    If InStr(1, fileName, "绿色一", vbTextCompare) > 0 Then
        result = "绿色一"
    ElseIf InStr(1, fileName, "绿色二", vbTextCompare) > 0 Then
        result = "绿色二"
    ElseIf InStr(1, fileName, "绿色三", vbTextCompare) > 0 Then
        result = "绿色三"
    ElseIf InStr(1, fileName, "绿色四", vbTextCompare) > 0 Then
        result = "绿色四"
    ElseIf InStr(1, fileName, "绿色五", vbTextCompare) > 0 Then
        result = "绿色五"
    ElseIf InStr(1, fileName, "绿色六", vbTextCompare) > 0 Then
        result = "绿色六"
    ElseIf InStr(1, fileName, "红色", vbTextCompare) > 0 Then
        result = "红色"
    ElseIf InStr(1, fileName, "绿色", vbTextCompare) > 0 Then
        result = "绿色"
    ElseIf InStr(1, fileName, "所有", vbTextCompare) > 0 Then
        result = "所有"
    End If
    getFileKind = result
End Function

