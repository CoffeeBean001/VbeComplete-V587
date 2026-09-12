Option Explicit

Public Const SP As String = "$$$$$"

Public Enum E
    A_ = 1: B_: C_: D_: E_: F_: G_: H_: I_: J_: K_: L_: M_: N_: O_: P_: Q_: R_: S_: T_: U_: V_: W_: X_: Y_: Z_
    AA_ = 27: AB_: AC_: AD_: AE_: AF_: AG_: AH_: AI_: AJ_: AK_: AL_: AM_: AN_: AO_: AP_: AQ_: AR_: AS_: AT_: AU_: AV_: AW_: AX_: AY_: AZ_
    BA_ = 53: BB_: BC_: BD_: BE_: BF_: BG_: BH_: BI_: BJ_: BK_: BL_: BM_: BN_: BO_: BP_: BQ_: BR_: BS_: BT_: BU_: BV_: BW_: BX_: BY_: BZ_
    CA_ = 79: CB_: CC_: CD_: CE_: CF_: CG_: CH_: CI_: CJ_: CK_: CL_: CM_: CN_: CO_: CP_: CQ_: CR_: CS_: CT_: CU_: CV_: CW_: CX_: CY_: CZ_
    DA_ = 105: DB_: DC_: DD_: DE_: DF_: DG_: DH_: DI_: DJ_: DK_: DL_: DM_: DN_: DO_: DP_: DQ_: DR_: DS_: DT_: DU_: DV_: DW_: DX_: DY_: DZ_
    EA_ = 131: EB_: EC_: ED_: EE_: EF_: EG_: EH_: EI_: EJ_: EK_: EL_: EM_: EN_: EO_: EP_: EQ_: ER_: ES_: ET_: EU_: EV_: EW_: EX_: EY_: EZ_
End Enum

'获取一个文件
Function getOneFile() As String
    Dim s$
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
    Dim s$
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then
            s = .SelectedItems(1) & "\"
        End If
    End With
    getFolderPath = s
End Function

'获取工作表最后一行
Function getLastRow(ws As Worksheet, titleRow As Long) As Long
    Dim maxRow&, i&, lastRow&
    For i = 1 To ws.Cells(titleRow, ws.Cells.Columns.Count).End(xlToLeft).Column
        lastRow = ws.Cells(ws.Cells.Rows.Count, i).End(xlUp).Row
        If lastRow > maxRow Then maxRow = lastRow
    Next
    getLastRow = maxRow
End Function

'获取工作表表头标题
Function getTitle(ws As Worksheet, r1 As Long, r2 As Long) As Object
    Dim dic As Object
    Dim i&, j&, s$
    Set dic = CreateObject("Scripting.Dictionary")
    For i = r1 To r2
        For j = 1 To ws.Cells(i, ws.Cells.Columns.Count).End(xlToLeft).Column
            s = ws.Cells(i, j)
            If Not dic.Exists(s) Then dic.Add s, j
        Next
    Next
    Set getTitle = dic
End Function

''获取所有文件
Sub getAllFiles(filesDic As Object, folderPath As String)
    Dim fs As Object, fd As Object, f As Object
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
Function getOneWorkbook(bookName As String) As Workbook
    Dim b As Workbook
    For Each b In Workbooks
        If b.Name = bookName Then
            Set getOneWorkbook = b
            Exit Function
        End If
    Next
    Set getOneWorkbook = Nothing
End Function

'从某个excel文件里查找某个工作表
Function getOneSheet(wb As Workbook, wsName As String) As Worksheet
    Dim w As Worksheet
    For Each w In wb.Worksheets
        If StrComp(w.Name, wsName, vbTextCompare) = 0 Then
            Set getOneSheet = w
            Exit Function
        End If
    Next
    Set getOneSheet = Nothing
End Function

''判断一个单词是否在数组中
Function isInArray(arr, s As String) As Boolean
    Dim a
    For Each a In arr
        If InStr(1, a, s, vbTextCompare) > 0 Then
            isInArray = True
            Exit Function
        End If
    Next
    isInArray = False
End Function

Sub deleteAllFiles(folderPath)
    Dim oneFile$
    oneFile = Dir(folderPath & "*.*")
    Do While oneFile <> ""
        Kill folderPath & oneFile
        oneFile = Dir
    Loop
End Sub

''调整行高
Sub judgeRowsHeight(sourceArea As Range, targetArea As Range)
    Dim i&
    For i = 1 To sourceArea.Rows.Count
        targetArea.Rows(i).RowHeight = sourceArea.Rows(i).RowHeight
    Next
End Sub

''调整列宽
Sub judgeColumnsWidth(sourceArea As Range, targetArea As Range)
    Dim i&
    For i = 1 To sourceArea.Columns.Count
        targetArea.Columns(i).ColumnWidth = sourceArea.Columns(i).ColumnWidth
    Next
End Sub

Function getFiles() As String
    Dim s$, i&
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




















