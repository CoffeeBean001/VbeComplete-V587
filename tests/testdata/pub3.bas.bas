'Public Const SP As String = "$$$$$"
'
'Public Enum E
'    A_ = 1: B_: C_: D_: E_: F_: G_: H_: I_: J_: K_: L_: M_: N_: O_: P_: Q_: R_: S_: T_: U_: V_: W_: X_: Y_: Z_
'    AA_ = 27: AB_: AC_: AD_: AE_: AF_: AG_: AH_: AI_: AJ_: AK_: AL_: AM_: AN_: AO_: AP_: AQ_: AR_: AS_: AT_: AU_: AV_: AW_: AX_: AY_: AZ_
'    BA_ = 53: BB_: BC_: BD_: BE_: BF_: BG_: BH_: BI_: BJ_: BK_: BL_: BM_: BN_: BO_: BP_: BQ_: BR_: BS_: BT_: BU_: BV_: BW_: BX_: BY_: BZ_
'    CA_ = 79: CB_: CC_: CD_: CE_: CF_: CG_: CH_: CI_: CJ_: CK_: CL_: CM_: CN_: CO_: CP_: CQ_: CR_: CS_: CT_: CU_: CV_: CW_: CX_: CY_: CZ_
'    DA_ = 105: DB_: DC_: DD_: DE_: DF_: DG_: DH_: DI_: DJ_: DK_: DL_: DM_: DN_: DO_: DP_: DQ_: DR_: DS_: DT_: DU_: DV_: DW_: DX_: DY_: DZ_
'    EA_ = 131: EB_: EC_: ED_: EE_: EF_: EG_: EH_: EI_: EJ_: EK_: EL_: EM_: EN_: EO_: EP_: EQ_: ER_: ES_: ET_: EU_: EV_: EW_: EX_: EY_: EZ_
'End Enum

'获取一个文件
Function getOneFile()
    s = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then
            s = .SelectedItems(1)
        End If
    End With
    getOneFile = s
End Function

'获取一个文件夹
Function getFolderPath()
    s = ""
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then
            s = .SelectedItems(1) & "\"
        End If
    End With
    getFolderPath = s
End Function

'获取工作表最后一行
Function getLastRow(ws, titleRow)
    maxRow = 0
    lastRow = 0
    For i = 1 To ws.Cells(titleRow, ws.Cells.Columns.Count).End(xlToLeft).Column + 100
        lastRow = ws.Cells(ws.Cells.Rows.Count, i).End(xlUp).Row
        If lastRow > maxRow Then maxRow = lastRow
    Next
    getLastRow = maxRow
End Function

'获取工作表表头标题
Function getTitle(ws, r1, r2)
    Set dic = CreateObject("Scripting.Dictionary")
    For i = r1 To r2
        For j = 1 To ws.Cells(i, ws.Cells.Columns.Count).End(xlToLeft).Column
            s = CStr(ws.Cells(i, j))
            If Not dic.Exists(s) Then dic.Add s, j
        Next
    Next
    Set getTitle = dic
End Function

''获取所有文件
Sub getAllFiles(filesDic, folderPath)
    Set fs = CreateObject("Scripting.FileSystemObject")
    For Each fd In fs.GetFolder(folderPath).SubFolders
        getAllFiles filesDic, fd.Path
    Next
    For Each f In fs.GetFolder(folderPath).Files
        If Left(f.Name, 1) <> "~" Then
            filesDic.Add f.Path, f.Name
        End If
    Next
End Sub

'从所有工作簿中查找某个工作簿
Function getOneWorkbook(bookName)
    For Each b In Workbooks
        If b.Name = bookName Then
            Set getOneWorkbook = b
            Exit Function
        End If
    Next
    Set getOneWorkbook = Nothing
End Function

'从某个excel文件里查找某个工作表
Function getOneSheet(wb, wsName)
    For Each w In wb.Worksheets
        If StrComp(w.Name, wsName, vbTextCompare) = 0 Then
            Set getOneSheet = w
            Exit Function
        End If
    Next
    Set getOneSheet = Nothing
End Function

''判断一个单词是否在数组中
Function isInArray(arr, s)
    For Each a In arr
        If InStr(1, CStr(a), CStr(s), vbTextCompare) > 0 Then
            isInArray = True
            Exit Function
        End If
    Next
    isInArray = False
End Function

Sub deleteAllFiles(folderPath)
    oneFile = Dir(folderPath & "*.*")
    Do While oneFile <> ""
        Kill folderPath & oneFile
        oneFile = Dir
    Loop
End Sub

''调整行高
Sub judgeRowsHeight(sourceArea, targetArea)
    For i = 1 To sourceArea.Rows.Count
        targetArea.Rows(i).RowHeight = sourceArea.Rows(i).RowHeight
    Next
End Sub

''调整列宽
Sub judgeColumnsWidth(sourceArea, targetArea)
    For i = 1 To sourceArea.Columns.Count
        targetArea.Columns(i).ColumnWidth = sourceArea.Columns(i).ColumnWidth
    Next
End Sub

Function getFiles()
    s = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = ThisWorkbook.Path
        .AllowMultiSelect = True
        If .Show = -1 Then
            For i = 1 To .SelectedItems.Count
                If s = "" Then
                    s = .SelectedItems(i)
                Else
                    s = s & Chr(10) & .SelectedItems(i)
                End If
            Next
        End If
    End With
    getFiles = s
End Function

Function getColsDic()
    Dim codesArr(1 To 2)
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    codesArr(2) = 65
    For i = 1 To 2888
        key = turnCodesToString(codesArr)
        dic.Add key, i
        codesArrAddOne codesArr
    Next
    Set getColsDic = dic
End Function

Private Sub codesArrAddOne(codesArr)
    codesArr(2) = codesArr(2) + 1
    If codesArr(2) > 90 Then
        codesArr(1) = codesArr(1) + 1
        If codesArr(1) < 65 Then codesArr(1) = 65
        codesArr(2) = 65
        If codesArr(1) > 90 Then
            codesArr(0) = codesArr(0) + 1
            If codesArr(0) < 65 Then codesArr(0) = 65
            codesArr(1) = 65
        End If
    End If
End Sub

Private Function turnCodesToString(codesArr)
    s1 = ""
    s2 = ""
    s3 = ""
    If codesArr(0) = 0 Then
        s1 = ""
    Else
        s1 = Chr(codesArr(0))
    End If
    If codesArr(1) = 0 Then
        s2 = ""
    Else
        s2 = Chr(codesArr(1))
    End If
    If codesArr(2) = 0 Then
        s3 = ""
    Else
        s3 = Chr(codesArr(2))
    End If
    turnCodesToString = s1 & s2 & s3
End Function


















