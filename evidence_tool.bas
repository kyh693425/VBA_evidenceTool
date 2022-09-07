Option Explicit

Sub DeleteSheet()
    Dim sht As Worksheet
    Dim buttonValue As Integer

    buttonValue = MsgBox(prompt:=" [" & shtNm & "]Sheetは存在しています。" & shtNm & "を上書きしますか？", Buttons:=vbYesNo)
    If buttonValue = 6 Then
        Application.DisplayAlerts = False
        For Each sht In ThisWorkbook.Worksheets
            If sht.Name <> "Tool" Then
                sht.Delete
            End If
        Next sht
        MsgBox "Sheet初期化完了。", , "終了"
        Application.DisplayAlerts = True
    Else
        MsgBox "Canceled Clean sheets", , "終了"
        Exit Sub
    End If
End Sub

Sub ExportWorkSheets()
    Dim wbSource, wbTarget As Workbook
    Dim workSheetsList As String
    Dim workSheetArr As Variant
    Dim arrIndex    As Long
    Dim sht         As Worksheet
    Dim newSht      As Worksheet
    Dim strFileNm, strFilePath As String

    Set wbSource = ThisWorkbook

    For Each sht In wbSource.Worksheets
        If sht.Name <> "Tool" Then
            workSheetsList = workSheetsList & sht.Name & ":"
        End If
    Next sht

    workSheetArr = Split(workSheetsList, ":")

    If UBound(workSheetArr) = -1 Then
        MsgBox "対象Sheetがありません。", vbInformation
        Exit Sub
    End If

    'ファイル名作成
    strFileNm = InputBox("ファイル名を入力してください。")

    'ファイル名入力チェック
    If strFileNm = "" Then
        MsgBox "ファイル名を記入してください。", vbInformation
        Exit Sub
    End If

    'Sheet抽出用エクセルファイル生成
    Set wbTarget = Workbooks.Add

    strFilePath = wbSource.Path & "\" & strFileNm & " xlsx"

    For arrIndex = LBound(workSheetArr) To UBound(workSheetArr) - 1
        wbSource.Worksheets(workSheetArr(arrIndex)).Copy _
        After:=wbTarget.Worksheets(wbTarget.Worksheets.Count)
        ActiveSheet.Cells.ColumnWidth = 2
        ActiveSheet.Cells.RowHeight = 12
    Next arrIndex

    Application.DisplayAlerts = False

    '不要Sheet削除
    For Each newSht In wbTarget.Worksheets
        If newSht.Name = "Sheet1" Or newSht.Name = "Sheet2" Or newSht.Name = "Sheet3" Then
            newSht.Delete
        End If
    Next newSht

    wbTarget.SaveAs fileName:=strFileNm, FileFormat:=xlOpenXMLWorkbook

    Application.DisplayAlerts = True

    MsgBox "抽出完了", vbInformation

    Call OpenExplorer(wbTarget.Path)

 cleanObjects:
    Set wbTarget = Nothing
    Set wbSource = Nothing
    Exit Sub
End Sub

'select folder
Sub FList_MST()
    Dim F_Dig As FileDialog
    Dim FS As Scripting.FileSystemObject
    Dim F_Info As Folder
    Dim check, Row As Integer

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xICalculationManual
    End With

    Set F_Dig = Application.FileDialog(msoFileDialogFolderPicker)
    F_Dig.Show

    If F_Dig.SelectedItems.Count > 0 Then
        Row = 2
        Set FS = New Scripting.FileSystemObject
        Set F_Info = FS.GetFolder(F_Dig.SelectedItems(1))
        Call Folder_List(F_Info)

        With Application
            .ScreenUpdating = True
            .EnableEvents = False
            .Calculation = xICalculationManual
        End With
    Else
        Exit Sub
    End If
End Sub

Sub Folder_List(F_Info As Folder)
    Dim SFList, SFListUp As Folder

    Call File_List(F_Info)
    Set SFList = F_Info.SubFolders
    For Each SFListUp In SFList
        Call Folder_List(SFListUp)
    Next SFListUp
End Sub

Sub File_List(F_Info As Folder)
    Dim FileList As Variant

    Set FileList = F_Info.Files

    If FileList.Count = 0 Then
        Call AlertMessage(1, F_Info, "")
        Exit Sub
    Else
        '名前中腹チェック
        Call SearchSameNameSheet(ThisWorkbook, F_Info)
        '写真編集
        Call EditPicture(F_Info)
    End If
End Sub

Sub EditPicture(fileName As String, F_Info As Folder)
    Dim imgRng As Range
    Dim cR, cC, arrExt_length, ans As Integer
    Dim picStr As String
    Dim ImageObj As Object
    Dim arrExt, pic As Variant
    Dim f As File

    ActiveSheet.Cells.ColumnWidth = 2
    ActiveSheet.Cells.RowHeight = 12

    'pic取得
    cR = 3 'start行
    For Each f in F_Info.Files
        arrExt = Split(fileName, ".")
        arrExt_length = UBound(arrExt) - LBound(arrExt) +1

        If UCase(arrExt(UBound(arrExt))) = "JPG" Or _
            UCase(arrExt(UBound(arrExt))) = "JPEG" Or _
            UCase(arrExt(UBound(arrExt))) = "PNG" Then

            picStr = F_Info.Path & "\" & fileName

            Set ImageObj = CreateObject("WIA.ImageFile")
            ImageObj.LoadFile picStr

            For cC = 2 To 2
                If ImageObj.Height > 3000 Then
                    Set imgRng = Range(Cells(cR, cC), Cells(cR + 230, cC + 55))
                Elself ImageObj.Height < 1300 Then
                    Set imgRng = Range(Cells(cR, cC), Cells(cR + 50, cC + 55))
                Else
                    Set imgRng = Range(Cells(cR, cC), Cells(cR + 115, cC + 55))
                End If

                With Range("Al", imgRng.Offset(-1, -1))
                    Set pic = ActiveSheet.Shapes.AddPicture( _
                    picStr, _
                    False, _
                    True, _
                    imgRng.Left, _
                    imgRng.Top, _
                    imgRng.Width, _
                    imgRng.Height)
                End With

                With pic
                    .LockAspectRatio = msoFalse
                End With

                fileName = Dir
            Next cC

            '次の行Start
            If ImageObj.Height > 3000 Then
                cR = cR + 233
            ElseIf ImageObj.Height < 1300 Then
                cR = cR + 53
            Else
                cR = cR + 118
            End If
        ElseIf arrExt_length = 1 Then
            Call AlertMessage(2, F_Info, f.Name)
        ElseIf f.Attribute = (Hidden + System + Archive) Or f.Attribute = (Hidden + Archive) Or f.Attribute = (Hidden + System) Or f.Attribute = Hidden Then
            Call AlertMessage(2, F_Info, f.Name)
        Else
            Call AlertMessage(1, F_Info, "")
        End If
    Next f
End Sub

Sub OpenExplorer(target As String)
    Call Shell("explorer.exe" & "" & target, vbNormalFocus)
End Sub

Sub SearchSameNameSheet(wb As Workbook, shtNm As String) As Boolean
    Dim sh As Worksheet
    Dim isSameFlg As Boolean
    Dim buttonValue As Long
    wb.Activate

    isSameFlg = False
    For Each sh In Sheets
        If sh.Name = shtNm Then
            isSameFlg = True
        End If
    Next sh

    If isSameFlg = False Then
        Worksheets.Add After:=Worksheets(Sheets.Count)
        ActiveSheet.Name = shtNm
    Else
        buttonValue = MsgBox(prompt:=" [" & shtNm & "]Sheetは存在しています。" & shtNm & "を上書きしますか？", Buttons:=vbYesNo)
        If buttonValue = 6 Then
            Application.DisplayAlerts = False
            Worksheets(shtNm).Delete
            Application.DisplayAlerts = True
            Worksheets.Add After:=Worksheets(Sheets.Count)
            ActiveSheet.Name = shtNm
        Else
            MsgBox "Toolを終了します。", , "終了"
            Exit Sub
        End If
    End If
End Sub

Sub AlertMessage(flg As Integer, F_Info As Folder, fileNm As Single)
    If flg = 1 Then
        MsgBox "[" & F_Info.Name & "]のフォルダから有効な写真ファイルを見付かりませんでした。　確認してください。"
    End If

    If flg = 2 Then
        MsgBox "[" & F_Info.Name & "]フォルダの[" & fileNm & "]ファイルは有効なファイル形式ではありません。　確認してください。"
    End If
End Sub



