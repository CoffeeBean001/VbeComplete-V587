Private Sub CommandButton1_Click()
    sourceFolderPath = CStr(Label2.Caption)
    targetFolderPath = CStr(Label2.Caption)
    If sourceFolderPath = "" Then MsgBox "请选择数据源文件夹！", vbExclamation: Exit Sub
    If targetFolderPath = "" Then MsgBox "请选择结果文件夹！", vbExclamation: Exit Sub
    
    orderArr = Array("绿色一", "绿色二", "绿色三", "绿色四", "绿色五", "绿色六", "红色", "绿色", "所有")
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    For Each fd In fs.GetFolder(sourceFolderPath).SubFolders
        For Each f In fd.Files
            Set dataBook = Workbooks.Open(f.Path)
            Set dataSheet = dataBook.Worksheets(1)
            num = CStr(dataSheet.Cells(1, 1))
            If num = "" Then GoTo 1
            arr = Split(num, ".")
            num = arr(0) & "." & arr(1) 'targetsheet里的号码，锁定列需要
            
            Set dataColsDic = getDataKindCols(dataSheet)    '绿色一，绿色二等列号
            
            If Dir(targetFolderPath & fd.Name, vbDirectory) = "" Then
                MkDir targetFolderPath & fd.Name
            End If
            If Dir(targetFolderPath & fd.Name & "\data_" & fd.Name & ".xlsx", vbNormal) = "" Then
                Set targetBook = Workbooks.Add()
            Else
                Set targetBook = Workbooks.Open(targetFolderPath & fd.Name & "\data_" & fd.Name & ".xlsx")
            End If
            
            Set targetSheet = targetBook.Worksheets(1)
            targetSheet.Copy after:=targetSheet
            Set targetSheet2 = targetBook.Worksheets(2)
            targetSheet.Cells.Clear
            Set targetPointsDic = getTargetKindPoints(targetSheet2)
            numColInTargetSheet = getNumColInTargetSheet(targetSheet2, num)
            
            dataLastCol = dataSheet.Cells(1, dataSheet.Cells.Columns.Count).End(xlToLeft).Column
            targetLastCol = targetSheet2.Cells(1, targetSheet2.Cells.Columns).End(xlToLeft).Column
            
            ''按照规定的顺序，填写相应数据。
            ''从targetsheet2表，往targetsheet表里填写。
            If numColInTargetSheet > 0 Then ''填充的列号
                pc = numColInTargetSheet
            Else
                pc = targetLastCol + 1
            End If
            targetSheet2.Cells(1, pc).EntireColumn.Clear
            targetSheet2.Cells(1, pc + 1).EntireColumn.Clear
            
            For i = 0 To UBound(orderArr)
                kind = CStr(orderArr(i))
                If dataColsDic.Exists(kind) And targetPointsDic.Exists(kind) Then
                    pr = pub3.getLastRow(targetSheet, 1)  '填充的行号
                    If pr > 1 Then pr = pr + 7
                    Set targetDataArea2 = targetSheet2.Range(targetSheet2.Cells(targetPointsDic(kind).开始行, "a"), _
                                                            targetSheet2.Cells(targetPointsDic(kind).结束行, targetLastCol))
                    targetSheet.Cells(pr, "a").Resize(targetDataArea2.Rows.Count, targetDataArea2.Columns.Count).value = targetDataArea2.value
                    targetSheet.Cells(pr, pc) = "'" & num
                    dataLastRow = dataSheet.Cells(dataSheet.Cells.Rows.Count, dataColsDic(kind)).End(xlUp).Row
                    
                    targetSheet.Cells (pr,pc)="'" &
                ElseIf dataColsDic.Exists(kind) And (Not targetPointsDic.Exists(kind)) Then
                    
                End If
            Next
            
1:
        Next
    Next
End Sub

Private Sub Label2_Click()
    Label2.Caption = pub3.getFolderPath()
End Sub

Private Sub Label4_Click()
    Label4.Caption = pub3.getFolderPath()
End Sub

Private Function getDataKindCols(ws)
    Set dic = CreateObject("Scripting.Dictionary")
    For i = 2 To ws.Cells(1, ws.Cells.Columns.Count).End(xlToLeft).Column Step 2
        key = CStr(ws.Cells(1, i))
        If key <> "" Then dic.Add key, i
    Next
    Set getDataKindCols = dic
End Function

''从目标表里，是上下竖着放的
''获取到每种kind的首行和尾行位置
Private Function getTargetKindPoints(ws)
    Set pointsDic = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Cells.Rows.Count, "b").End(xlUp).Row
    preKind = ""
    For i = 1 To lastRow
        curKind = CStr(ws.Cells(i, "b"))
        If curKind = "绿色一" Or curKind = "绿色二" Or curKind = "绿色三" Or curKind = "绿色四" Or _
           curKind = "绿色五" Or curKind = "绿色六" Or curKind = "红色" Or curKind = "绿色" Or curKind = "所有" Then
            If Not pointsDic.Exists(curKind) Then pointsDic.Add curKind, New 类1
            pointsDic(curKind).开始行 = i
            If preKind <> "" Then
                pointsDic(preKind).结束行 = i - 1
            End If
            If i = lastRow Then
                pointsDic(curKind).结束行 = i
            End If
            preKind = curKind
        End If
    Next
    Set getTargetKindPoints = pointsDic
End Function

'获取num在targetsheet中的列号
Private Function getNumColInTargetSheet(ws, num)
    For i = 1 To ws.Cells(1, ws.Cells.Columns.Count).End(xlToLeft).Column
        num2 = CStr(ws.Cells(1, i))
        If StrComp(num, num2, vbTextCompare) = 0 Then
            getNumColInTargetSheet = i
            Exit Function
        End If
    Next
    getNumColInTargetSheet = -1
End Function























