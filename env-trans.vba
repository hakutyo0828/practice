Option Explicit

Public Sub 環境名変換ツール_XLSX()
    Dim wsTool As Worksheet, wsTbl As Worksheet
    Set wsTool = ThisWorkbook.Worksheets("変換ツール")
    Set wsTbl = ThisWorkbook.Worksheets("テーブル")
    
    Dim srcEnv As String, dstEnv As String
    srcEnv = Trim(wsTool.Range("B3").Value)
    dstEnv = Trim(wsTool.Range("D3").Value)
    If Len(srcEnv) = 0 Or Len(dstEnv) = 0 Then
        MsgBox "B3に変換元、D3に変換後の環境名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Dim lastCol As Long, srcCol As Long, dstCol As Long, c As Long
    lastCol = wsTbl.Cells(1, wsTbl.Columns.Count).End(xlToLeft).Column
    For c = 2 To lastCol
        If wsTbl.Cells(1, c).Value = srcEnv Then srcCol = c
        If wsTbl.Cells(1, c).Value = dstEnv Then dstCol = c
    Next c
    If srcCol = 0 Or dstCol = 0 Then
        MsgBox "テーブルの1行目に指定した環境名の列が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsTbl.Cells(wsTbl.Rows.Count, srcCol).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "テーブルに変換データがありません。", vbExclamation
        Exit Sub
    End If
    
    ' xlsxファイル選択
    Dim fd As Object, fPath As String
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .AllowMultiSelect = False
        .Title = "変換対象のxlsxを選択してください"
        .Filters.Clear
        .Filters.Add "Excelブック", "*.xlsx;*.xlsm;*.xlsb"
        If .Show <> -1 Then Exit Sub
        fPath = .SelectedItems(1)
    End With
    
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=fPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    
    Dim fromStr As String, toStr As String, ws As Worksheet, r As Long
    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        For r = lastRow To 2 Step -1 ' 下から上へ
            fromStr = CStr(wsTbl.Cells(r, srcCol).Value)
            toStr = CStr(wsTbl.Cells(r, dstCol).Value)
            If Len(fromStr) > 0 Then
                ws.Cells.Replace What:=fromStr, Replacement:=toStr, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True
            End If
        Next r
    Next ws
    Application.ScreenUpdating = True
    
    ' 保存先選択（名前の加工なし）
    Dim saveName As Variant
    saveName = Application.GetSaveAsFilename(InitialFileName:=fPath, _
                                             FileFilter:="Excelブック,*.xlsx,マクロ有効ブック,*.xlsm,バイナリブック,*.xlsb", _
                                             Title:="変換後ファイルを保存")
    If saveName = False Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=saveName, FileFormat:=wb.FileFormat
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    
    MsgBox "変換が完了しました。", vbInformation
End Sub
