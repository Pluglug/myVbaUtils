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

Private Sub CreateButtons(ByVal targetSheet As Worksheet, _
                  ByVal textColumnIndex As Integer, _
                  ByVal buttonOffset As Integer, _
                  ByVal buttonCaption As String, _
                  ByVal buttonAction As String, _
                  ByVal buttonPrefix As String, _
                  Optional ByVal startRow As Long = 1, _
                  Optional ByVal DeleteExistingButtons As Boolean = True)
    Dim lastRow As Long
    Dim btn As Button
    Dim row_i As Long
    Dim buttonColumnIndex As Integer
    
    buttonColumnIndex = textColumnIndex + buttonOffset
    
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
                .OnAction = buttonAction
                .Caption = buttonCaption
                .Name = buttonPrefix & "Button" & row_i + textColumnIndex * 100
            End With
        End If
    Next row_i
End Sub

' --- Public
Public Sub CreateCopyButtons(ByVal targetSheet As Worksheet, _
                      ByVal textColumnIndex As Integer, _
                      Optional ByVal startRow As Long = 1, _
                      Optional ByVal DeleteExistingButtons As Boolean = True)
    Call CreateButtons(targetSheet, textColumnIndex, 1, "Copy", "CopyToClipboard", "CopyToClipboard", startRow, DeleteExistingButtons)
End Sub

Public Sub CreateAppendButtons(ByVal targetSheet As Worksheet, _
                        ByVal textColumnIndex As Integer, _
                        Optional ByVal startRow As Long = 1, _
                        Optional ByVal DeleteExistingButtons As Boolean = True)
    Call CreateButtons(targetSheet, textColumnIndex, 2, "Append", "AppendToClipboard", "AppendToClipboard", startRow, DeleteExistingButtons)
End Sub

' --- Internal
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
    
    Call Beep(441, 100)  ' フィードバック音 (A4, 100ms)  ' 行ごとに半音ずつ変えると楽しいかも
End Sub

Sub AppendToClipboard()
    Dim btn As Button
    Dim trgRow As Long
    Dim trgCol As Long
    Dim targetSheet As Worksheet
    Dim textCell As Range
    Dim clipboardText As String
    Dim dataObj As Object
    
    Set btn = ActiveSheet.Buttons(Application.Caller)
    Set targetSheet = btn.TopLeftCell.Worksheet
    trgRow = btn.TopLeftCell.Row
    trgCol = btn.TopLeftCell.Column - 2  ' ボタンの2つ左隣のセル
    
    Set textCell = targetSheet.Cells(trgRow, trgCol)
    
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.GetFromClipboard
    clipboardText = dataObj.GetText
    
    clipboardText = clipboardText & vbCrLf & vbCrLf & textCell.Value
    
    dataObj.SetText clipboardText
    dataObj.PutInClipboard
    
    Call Beep(523, 100)  ' フィードバック音 (C5, 100ms)
End Sub
