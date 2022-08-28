'Sheet削除
Sub DeleteSheet()
    Dim sht         As Worksheet

    Application.DisplayAlerts = False
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name <> "Tool" Then
            sht.Delete
        End If
    Next sht
    MsgBox "Sheet初期化完了。", , "終了"
    Application.DisplayAlerts = True
End Sub

'Sheet抽出
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
        MsgBox "抽出?象Sheetがありません。", vbInformation
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
        If newSht.Name = "Sheetl" Or newSht.Name = "Sheet2" Or newSht.Name = "Sheet3" Then
            newSht.Delete
        End If
    Next newSht

    wbTarget.SaveAs fileName:=strFileNm, _
    FileFormat:=xlOpenXMLWorkbook

    Application.DisplayAlerts = True

    MsgBox "抽出完了", vbInformation

    OpenExplorer (wbTarget.Path)

 cleanObjects:
    Set wbTarget = Nothing
    Set wbSource = Nothing
    Exit Sub

End Sub

Sub FList_MST()
    Dim F_Dig As FileDialog
    Dim FS As Scripting.FileSystemObject
    Dim F_Info As Folder
    Dim check As Integer

    With Application
        .ScreenUpdating = False
        EnableEvents = False
        Calculation = xICalculationManual
    End With

    Set F_Dig = Application.FileDialog(msoFileDialogFolderPicker)
    F_Dig.Show

    If F_Dig.SelectedItems.Count > 0 Then
        Row = 2
        Set FS = New Scripting.FileSystemObject
        Set F_Info = FS.GetFolder(F_Dig.SelectedItems(1))
        Call Folder_List(F_Info)

        With Application
            Screen Updating = True
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
    Dim FleList, FileListUp As File
    Dim totalSheets As Integer
    Dim imgRng As Range
    Dim cR, cC, i As Integer
    Dim ans As Long
    Dim fileName, picStr, tempNm As String
    Dim allFiles As Variant
    Dim ImageObj As Object

    Set FileList = F_Info.Files

    If FileList.Count = 0 Then
        fileCheck = MsgBox(prompt:="[" & F_Info.Name & "]のフォルダから有?な??ファイルを見付かりませんでした。", Buttons:=vbOKOnly)

        If fileCheck = 1 Then
            Exit Sub
        End If

    Else
        fileName = Dir(F_Info.Path & "¥" & "*.*")

        If FileList.Count > 0 And fileName = "" Then
            fileCheck = MsgBox(prompt:="[" & F_Info.Name & "]のフォルダから有?な??ファイルを見付かりませんでした。", Buttons:=vbOKOnly)

            If fileCheck = 1 Then
                Exit Sub
            End If

        End If

        'Sheet名重複チェック
        If SameSheetNmSearch(ThisWorkbook, F_Info.Name) = False Then
            Worksheets.Add After:=Worksheets(Sheets.Count)
            ActiveSheet.Name = F_Info.Name
        Else
            ans = MsgBox(prompt:=" [" & F_Info.Name & "]Sheetは存在しています。" & F_Info.Name & "を上書きしますか？", Buttons:=vbYesNo)
            If ans = 6 Then
                Application.DisplayAlerts = False
                Worksheets(F_Info.Name).Delete
                Application.DisplayAlerts = True
                Worksheets.Add After:=Worksheets(Sheets.Count)
                ActiveSheet.Name = F_Info.Name
            Else
                MsgBox "Toolを終了します。", , "終了"
                Exit Sub
            End If
        End If

        ActiveSheet.Cells.ColumnWidth = 2
        ActiveSheet.Cells.RowHeight = 12

        'pic取得
        cR = 3 'start行
        Do While fileName <> ""
            arrExt = Split(fileName, ".")

            If UCase(arrExt(UBound(arrExt))) = "JPG" Or _
                UCase(arrExt(UBound(arrExt))) = "JPEG" Or _
                UCase(arrExt(UBound(arrExt))) = "PNG" Then

                picStr = F_Info.Path & "\" & fileName

                Set ImageObj = CreateObject("WIA.ImageFile")
                ImageObj.LoadFile picStr

                For cC = 2 To 2
                    If ImageObj.Height > 3000 Then
                        Set imgRng = Range(Cells(cR, cC), Cells(cR + 230, cC + 55))
                    Elself ImageObj.Height > 3000 Then
                        Set imgRng = Range(Cells(cR, cC), Cells(cR + 50, cC + 55))
                    Else
                        Set imgRng = Range(Cells(cR, cC), Cells(cR + 115, cC + 55))

                        With Range("Al", imgRng.Offset(-1, -1))
                            Set pic = ActiveSheet.Shapes.AddPicture(picStr, False, True, imgRng.Left, imgRng.Top, imgRng.Width, imgRng.Height)
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
                Else
                    fileCheck = MsgBox(prompt:=F_Info.Name & "ファイルがありません。確認してください。", Buttons:=vbOKOnly)
                    If fileCheck = 1 Then
                        Exit Sub
                    End If
                End If
            Loop
        End If
End Sub

'Sheet名重複チェック
Function SameSheetNmSearch(wb As Workbook, shtNm As String) As Boolean
    Dim i As Long
    Dim sh As Worksheet
    SameSheetNmSearch = False
    wb.Activate

    For Each sh In Sheets
        If sh.Name = shtNm Then
            SameSheetNmSearch = True
        End If
        Next
End Function

Sub OpenExplorer(target As String)
    Call Shell("explorer.exe" & "" & target, vbNormalFocus)
End Sub

