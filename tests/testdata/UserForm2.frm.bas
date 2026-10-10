'Option Explicit
'
'Private Sub CommandButton1_Click()
'    Dim sourceFolderPath$, targetFolderPath$
'    Dim fs As Object, fd As Object, f As Object
'    Dim dataBook As Workbook, dataSheet As Worksheet
'    Dim targetBook As Workbook, targetSheet As Worksheet, targetSheet2 As Worksheet
'    Dim targetPointersDic As Object, dataPointersDic As Object, orderDic As Dictionary
'    Dim i&, lastRow&, j&, lastCol&, arr, num$, targetCol&, orderArr, itemsArr
'    Dim p&, kind$, targetLastCol&, dataLastRow&
'
'    sourceFolderPath = Label2.Caption
'    targetFolderPath = Label4.Caption
'
'    If sourceFolderPath = "" Then
'        MsgBox "请选择数据源文件夹！", vbExclamation
'        Exit Sub
'    End If
'    If targetFolderPath = "" Then
'        MsgBox "请选择结果文件夹！", vbExclamation
'        Exit Sub
'    End If
'
'    Set orderDic = getOrderDic() '按照绿色一，绿色二，这样顺序往下排
'    orderArr = orderDic.Keys
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    For Each fd In fs.GetFolder(sourceFolderPath).SubFolders
'        For Each f In fd.Files
'            Set dataBook = Workbooks.Open(f.Path)
'            Set dataSheet = dataBook.Worksheets(1)
'            Set dataPointersDic = getDataKindPointers(dataSheet)    '原始数据的列号
'
'            If Dir(targetFolderPath & fd.Name, vbDirectory) = "" Then
'                MkDir targetFolderPath & fd.Name
'            End If
'            If Dir(targetFolderPath & fd.Name & "\" & "data_" & fd.Name & ".xlsx", vbNormal) = "" Then
'                Set targetBook = Workbooks.Add()
'            Else
'                Set targetBook = Workbooks.Open(targetFolderPath & fd.Name & "\" & "data_" & fd.Name & ".xlsx")
'            End If
'            Set targetSheet = targetBook.Worksheets(1)
'            targetSheet.Copy after:=targetSheet
'            Set targetSheet2 = targetBook.Worksheets(2)
'            targetSheet.Cells.Clear
'
'            Set targetPointersDic = getTargetKindPointers(targetSheet)    '结果表里的kind标记，方便扩充。
'            targetSheet.Cells.ClearContents
'
'            arr = Split(dataSheet.Cells(1, 1), ".")
'            num = arr(0) & "." & arr(1)
'
'            targetLastCol = targetSheet2.Cells(1, targetSheet2.Cells.Columns.Count).End(xlToLeft).Column
'            targetCol = getCol(targetSheet, num)
'            targetSheet.Cells.Clear
'            If targetCol > 0 Then   '修改已经填的数据
'                targetSheet.Cells(1, targetCol).EntireColumn.Clear
'                targetSheet.Cells(1, targetCol + 1).EntireColumn.Clear
'                For i = 0 To UBound(orderArr)
'                    kind = orderArr(i)  '绿色一，绿色二等
'                    If targetPointersDic.Exists(kind) And dataPointersDic.Exists(kind) Then
'                        arr = targetSheet2.Range(targetSheet2.Cells(itemsArr(i).开始行, 1), targetSheet2.Cells(itemsArr(i).结束行, targetLastCol))
'                        p = pub2.getLastRow(targetSheet, 1)
'                        If p > 1 Then p = p + 7
'                        targetSheet.Cells(p, 1).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
'                        targetSheet.Cells(p, targetCol) = "'" & num
'                        targetSheet.Cells(p, targetCol + 1) = kind
'                        targetSheet.Cells(p, targetCol).Font.Bold = True
'                        targetSheet.Cells(p, targetCol + 1).Font.Bold = True
'                        For j = 2 To dataSheet.Cells(dataSheet.Cells.Rows.Count, dataPointersDic(kind)).End(xlUp).Row
'                            p = p + 1
'                            targetSheet.Cells(p, targetCol) = dataSheet.Cells(j, dataPointersDic(kind) - 1)
'                            targetSheet.Cells(p, targetCol + 1) = dataSheet.Cells(j, dataPointersDic(kind))
'                        Next
'                    ElseIf (Not targetPointersDic.Exists(kind)) And dataPointersDic.Exists(kind) Then
'                        p = pub2.getLastRow(targetSheet, 1)
'                        If p > 1 Then p = p + 7
'                        targetSheet.Cells(p, 1).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
'                        targetSheet.Cells(p, targetCol) = "'" & num
'                        targetSheet.Cells(p, targetCol + 1) = kind
'                        targetSheet.Cells(p, targetCol).Font.Bold = True
'                        targetSheet.Cells(p, targetCol + 1).Font.Bold = True
'                        For j = 2 To dataSheet.Cells(dataSheet.Cells.Rows.Count, dataPointersDic(kind)).End(xlUp).Row
'                            p = p + 1
'                            targetSheet.Cells(p, targetCol) = dataSheet.Cells(j, dataPointersDic(kind) - 1)
'                            targetSheet.Cells(p, targetCol + 1) = dataSheet.Cells(j, dataPointersDic(kind))
'                        Next
'                    ElseIf targetPointersDic.Exists(kind) And (Not dataPointersDic.Exists(kind)) Then
'                        p = pub2.getLastRow(targetSheet, 1)
'                        If p > 1 Then p = p + 7
'                        targetSheet.Cells(p, 1).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
'                        targetSheet.Cells(p, targetCol) = "'" & num
'                        targetSheet.Cells(p, targetCol + 1) = kind
'                        targetSheet.Cells(p, targetCol).Font.Bold = True
'                        targetSheet.Cells(p, targetCol + 1).Font.Bold = True
'                        For j = targetPointersDic.开始行 + 1 To targetPointersDic.结束行
'                            p = p + 1
'                            targetSheet.Cells(p, targetCol) = targetSheet2.Cells(j, targetCol)
'                            targetSheet.Cells(p, targetCol + 1) = targetSheet2.Cells(j, targetCol + 1)
'                        Next
'                    Else    '两个都没有，就什么都不做。
'
'                    End If
'                Next
'            Else    '填充新的数据，targetLastCol=-1
'                targetLastCol = targetSheet2.Cells(1, targetSheet2.Cells.Columns.Count).End(xlToLeft).Column + 1
'                If targetLastCol = 2 Then targetLastCol = 1
'                For i = 0 To UBound(orderArr)
'                    kind = orderArr(i)  '绿色一，绿色二等
'                    If targetPointersDic.Exists(kind) And dataPointersDic.Exists(kind) Then
'                        arr = targetSheet2.Range(targetSheet2.Cells(itemsArr(i).开始行, 1), targetSheet2.Cells(itemsArr(i).结束行, targetLastCol))
'                        p = pub2.getLastRow(targetSheet, 1)
'                        If p > 1 Then p = p + 7
'                        targetSheet.Cells(p, 1).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
'                        targetSheet.Cells(p, targetCol) = "'" & num
'                        targetSheet.Cells(p, targetCol + 1) = kind
'                        targetSheet.Cells(p, targetCol).Font.Bold = True
'                        targetSheet.Cells(p, targetCol + 1).Font.Bold = True
'                        For j = 2 To dataSheet.Cells(dataSheet.Cells.Rows.Count, dataPointersDic(kind)).End(xlUp).Row
'                            p = p + 1
'                            targetSheet.Cells(p, targetCol) = dataSheet.Cells(j, dataPointersDic(kind) - 1)
'                            targetSheet.Cells(p, targetCol + 1) = dataSheet.Cells(j, dataPointersDic(kind))
'                        Next
'                    ElseIf (Not targetPointersDic.Exists(kind)) And dataPointersDic.Exists(kind) Then
'                        p = pub2.getLastRow(targetSheet, 1)
'                        If p > 1 Then p = p + 7
'                        targetSheet.Cells(p, 1).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
'                        targetSheet.Cells(p, targetCol) = "'" & num
'                        targetSheet.Cells(p, targetCol + 1) = kind
'                        targetSheet.Cells(p, targetCol).Font.Bold = True
'                        targetSheet.Cells(p, targetCol + 1).Font.Bold = True
'                        For j = 2 To dataSheet.Cells(dataSheet.Cells.Rows.Count, dataPointersDic(kind)).End(xlUp).Row
'                            p = p + 1
'                            targetSheet.Cells(p, targetCol) = dataSheet.Cells(j, dataPointersDic(kind) - 1)
'                            targetSheet.Cells(p, targetCol + 1) = dataSheet.Cells(j, dataPointersDic(kind))
'                        Next
'                    ElseIf targetPointersDic.Exists(kind) And (Not dataPointersDic.Exists(kind)) Then
'                        p = pub2.getLastRow(targetSheet, 1)
'                        If p > 1 Then p = p + 7
'                        targetSheet.Cells(p, 1).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
'                        targetSheet.Cells(p, targetCol) = "'" & num
'                        targetSheet.Cells(p, targetCol + 1) = kind
'                        targetSheet.Cells(p, targetCol).Font.Bold = True
'                        targetSheet.Cells(p, targetCol + 1).Font.Bold = True
'                        For j = targetPointersDic.开始行 + 1 To targetPointersDic.结束行
'                            p = p + 1
'                            targetSheet.Cells(p, targetCol) = targetSheet2.Cells(j, targetCol)
'                            targetSheet.Cells(p, targetCol + 1) = targetSheet2.Cells(j, targetCol + 1)
'                        Next
'                    Else    '两个都没有，就什么都不做。
'
'                    End If
'                Next
'            End If
'        Next
'    Next
'
'    Unload Me
'    MsgBox "OK", vbInformation
'End Sub
'
'Private Sub Label2_Click()
'    Label2.Caption = pub2.getFolderPath()
'End Sub
'
'Private Sub Label4_Click()
'    Label4.Caption = pub2.getFolderPath()
'End Sub
'
'''按照绿色一，绿色二这个顺序往下排列
'Private Function getOrderDic() As Dictionary
'    Dim dic As New Dictionary
'    dic.Add "绿色一", ""
'    dic.Add "绿色二", ""
'    dic.Add "绿色三", ""
'    dic.Add "绿色四", ""
'    dic.Add "绿色五", ""
'    dic.Add "绿色六", ""
'    dic.Add "红色", ""
'    dic.Add "绿色", ""
'    dic.Add "所有", ""
'    Set getOrderDic = dic
'End Function
'
'''获取绿色一，绿色二等的列号
'Private Function getDataKindPointers(ByVal ws As Worksheet) As Object
'    Dim i&, key$, lastCol&, dic As Object
'    Set dic = CreateObject("Scripting.Dictionary")
'    lastCol = ws.Cells(1, ws.Cells.Columns.Count).End(xlToLeft).Column
'    For i = 2 To lastCol Step 2
'        key = ws.Cells(1, i)
'        dic.Add key, i
'    Next
'    Set getDataKindPointers = dic
'End Function
'
'''获取绿色一，绿色二等的开始结束位置
'Private Function getTargetKindPointers(ByVal ws As Worksheet) As Object
'    Dim pointersDic As Object, i, curKind$, preKind$, lastRow&
'    Set pointersDic = CreateObject("Scripting.Dictionary")
'    lastRow = ws.Cells(ws.Cells.Rows.Count, "b").End(xlUp).Row
'    For i = 1 To lastRow
'        curKind = ws.Cells(i, "b")
'        If curKind = "绿色一" Or curKind = "绿色二" Or _
'           curKind = "绿色三" Or curKind = "绿色四" Or _
'           curKind = "绿色五" Or curKind = "绿色六" Or _
'           curKind = "红色" Or curKind = "绿色" Or _
'           curKind = "所有" Then
'            If Not pointersDic.Exists(curKind) Then pointersDic.Add curKind, New 类1
'            pointersDic(curKind).开始行 = i
'            If preKind <> "" Then    '上一个标记的最后一行
'                pointersDic(preKind).结束行 = i - 1
'            End If
'            If i = lastRow Then
'                pointersDic(curKind).结束行 = i
'            End If
'            preKind = curKind
'        End If
'    Next
'    Set getTargetKindPointers = pointersDic
'End Function
'
'Private Function getCol(ByVal ws As Worksheet, ByVal title As String) As Long
'    Dim i&, lastCol&
'    lastCol = ws.Cells(1, ws.Cells.Columns.Count).End(xlToLeft).Column
'    For i = 1 To lastCol
'        If StrComp(ws.Cells(1, i), title, vbTextCompare) = 0 Then
'            getCol = i
'            Exit Function
'        End If
'    Next
'    getCol = -1
'End Function
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'

