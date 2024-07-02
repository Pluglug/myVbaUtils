' Xtra_CopyPallet.bas
' セルの内容をクリップボードにコピーするボタンを作成します。
' テンプレートなど、頻繁に使用するテキストを管理するのに便利です。

' 使い方:
' 1. Xtra_CopyPallet.bas(このファイル)をブックにインポートする
' 2. ボタンを作成したいシートのシートモジュールに以下のプロシージャを追加する
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    ' テキスト列を2列目の2行目から開始。(ボタンは固定でテキストの右隣に配置される)
'    CreateCopyButtons Me, 2, 2
'End Sub
' 3. 管理したいテキストをシートに入力すると、右隣にCopyボタンが作成される
' 4. Copyボタンをクリックすると、その行のテキストがクリップボードにコピーされる

Option Explicit

Declare PtrSafe Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Sub CreateCopyButtons(ByVal targetSheet As Worksheet, _
                      ByVal textColumnIndex As Integer, _
                      Optional ByVal startRow As Long = 1, _
                      Optional ByVal DeleteExistingButtons As Boolean = True)
    Dim lastRow As Long
    Dim btn As Button
    Dim row_i As Long
    
    Dim buttonColumnIndex As Integer
    buttonColumnIndex = textColumnIndex + 1  ' テキストの右隣にボタンを配置
    
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, textColumnIndex).End(xlUp).Row
    
    If DeleteExistingButtons Then
        On Error Resume Next
        targetSheet.Buttons.Delete
        On Error GoTo 0
    End If

    For row_i = startRow To lastRow
        If targetSheet.Cells(row_i, textColumnIndex).Value <> "" Then
            Set btn = targetSheet.Buttons.Add( _
                targetSheet.Cells(row_i, buttonColumnIndex).Left, _
                targetSheet.Cells(row_i, buttonColumnIndex).Top, _
                targetSheet.Cells(row_i, buttonColumnIndex).Width, _
                targetSheet.Cells(row_i, buttonColumnIndex).Height)
            With btn
                .OnAction = "CopyToClipboard"
                .Caption = "Copy"
                ' HACK: 同一シートに複数のテキスト管理列を作成する際の重複回避措置
                .Name = "CopyButton" & row_i + textColumnIndex * 100
            End With
        End If
    Next row_i
End Sub

Sub CopyToClipboard()
    Dim btn As Button
    Dim trgRow As Long
    Dim trgCol As Long
    Dim targetSheet As Worksheet
    Dim textCell As Range
    
    Set btn = ActiveSheet.Buttons(Application.Caller)
    Set targetSheet = btn.TopLeftCell.Worksheet
    trgRow = btn.TopLeftCell.Row
    trgCol = btn.TopLeftCell.Column - 1  ' ボタンの左隣のセル
    
    Set textCell = targetSheet.Cells(trgRow, trgCol)
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  ' MSForms.DataObject
        .SetText textCell.Value
        .PutInClipboard
    End With
    
    Call Beep(441, 100)  ' フィードバック音 (A4, 100ms)
End Sub
