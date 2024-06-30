Option Explicit

Private Const DBG_GS As Boolean = False
' ユーザーが選択した行を取得し、処理に利用できます。
' テーブル等の特定の行データを取得する場合に便利です。

Private Const ERR_ONLY_ONE_ROW As Long = vbObjectError + 1000
Private Const ERR_OUT_OF_TABLE_RANGE As Long = vbObjectError + 1010
Private Const ERR_NO_TABLE_DATA As Long = vbObjectError + 1011
Private Const MSG_ONLY_ONE_ROW As String = "1行だけ選択してください。"
Private Const MSG_OUT_OF_TABLE_RANGE As String = _
                        "選択された行がテーブルのデータ範囲外です。" & vbCrLf & _
                        "テーブル範囲を選択して、やり直してください。"
Private Const MSG_NO_TABLE_DATA As String = "テーブルにデータがありません。"

Private Sub HandleError(Err As ErrObject)
    Select Case Err.Number
        Case ERR_ONLY_ONE_ROW
            MsgBox MSG_ONLY_ONE_ROW, vbExclamation, "参照エラー"
        Case ERR_OUT_OF_TABLE_RANGE
            MsgBox MSG_OUT_OF_TABLE_RANGE, vbExclamation, "参照エラー"
        Case ERR_NO_TABLE_DATA
            MsgBox MSG_NO_TABLE_DATA, vbExclamation, "参照エラー"
        Case Else
            If Err.Number <> 0 Then
                Log.Error "Error: " & Err.Number & " " & Err.Description
                MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
            End If
    End Select
End Sub


''' 単一のセルの行番号を取得する

' 指定したワークシートで単一セルが選択されているか確認し、そのシートの絶対行番号を返します。
Public Function GetSelectedRow(ByRef ws As Worksheet) As Long
    If DBG_GS Then Log.Info "GetSelectedRow(ws=" & ws.Name & ")"
    On Error GoTo catch_err

    Dim selectedRange As Range
    Set selectedRange = ws.Application.Selection

    If selectedRange.Rows.Count > 1 Then
        Err.Raise ERR_ONLY_ONE_ROW, "GetSelectedRow", MSG_ONLY_ONE_ROW
    End If

    GetSelectedRow = selectedRange.Cells(1, 1).row
    If DBG_GS Then Log.Info "Row: " & GetSelectedRow
    Exit Function
catch_err:
    HandleError Err
    GetSelectedRow = -1
End Function

' 指定したワークシートで選択されたセルに基づいてテーブルの相対行番号を返します。
Public Function GetSelectedTableRow(ws As Worksheet, tbl As ListObject) As Long
    If DBG_GS Then Log.Info "GetSelectedTableRow(ws=" & ws.Name & ", tbl=" & tbl.Name & ")"
    On Error GoTo catch_err

    Dim selectedRow As Long
    Dim tblBody As Range

    selectedRow = GetSelectedRow(ws)
    If selectedRow = -1 Then
        GetSelectedTableRow = -1
        Exit Function
    End If

    Set tblBody = tbl.DataBodyRange  ' Chenged tbl.Range -> tbl.DataBodyRange

    If selectedRow < tblBody.row Or selectedRow >= tblBody.row + tblBody.Rows.Count Then
        Err.Raise ERR_OUT_OF_TABLE_RANGE, "GetSelectedTableRow", "Selected cell is outside the table range."
    End If

    GetSelectedTableRow = GetTableRelativeRow(tbl, selectedRow)
    If DBG_GS Then Log.Info "Relative Row: " & GetSelectedTableRow
    Exit Function
catch_err:
    HandleError Err
    GetSelectedTableRow = -1
End Function


''' 複数行の選択されたセルの行番号を取得する

' 指定したワークシートで選択されているセルのシート上の絶対行番号の配列を返します。
' 選択されているセルがない場合は、空の配列を返します。ユーザーが選択した順に行番号が格納されます。
Public Function GetSelectedSheetRows(ByRef ws As Worksheet) As Variant  ' or Collection
    If DBG_GS Then Log.Info "GetSelectedSheetRows(ws=" & ws.Name & ")"
    On Error GoTo catch_err

    Dim selectedRange As Range
    Dim arrRows() As Long
    Dim cnt_i As Long
    Dim cell As Range
    Dim uniqueRows As Collection

    Set uniqueRows = New Collection
    Set selectedRange = ws.Application.Selection

    If selectedRange Is Nothing Then
        Log.Warn "No selected cells."
        GetSelectedSheetRows = Array()
        Exit Function
    End If

    For Each cell In selectedRange
        On Error Resume Next
        uniqueRows.Add cell.row, CStr(cell.row)
        On Error GoTo 0
    Next cell

    ReDim arrRows(1 To uniqueRows.Count)
    For cnt_i = 1 To uniqueRows.Count
        arrRows(cnt_i) = uniqueRows(cnt_i)  ' 1次元配列に変換
        If DBG_GS Then Log.Info "Row: " & arrRows(cnt_i)
    Next cnt_i

    GetSelectedSheetRows = arrRows

    Exit Function
catch_err:
    HandleError Err
    GetSelectedSheetRows = Array()
End Function


' 指定したテーブル範囲内で選択されているセルのテーブル上の相対行番号の配列を返します。
Public Function GetSelectedTableRows(ByRef ws As Worksheet, ByRef tbl As ListObject) As Variant
    If DBG_GS Then Log.Info "GetSelectedTableRows(ws=" & ws.Name & ", tbl=" & tbl.Name & ")"
    On Error GoTo catch_err

    Dim selectedRows() As Long
    Dim relativeRows() As Long
    Dim cnt_i As Long, cnt_j As Long
    Dim tblBody As Range
    Dim startRow As Long, endRow As Long

    selectedRows = GetSelectedSheetRows(ws)

    If UBound(selectedRows) = 0 Then
        GetSelectedTableRows = Array()
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        Err.Raise ERR_NO_TABLE_DATA, "GetSelectedTableRows", MSG_NO_TABLE_DATA
    End If

    Set tblBody = tbl.DataBodyRange
    startRow = tblBody.row
    endRow = tblBody.row + tblBody.Rows.Count - 1
    If DBG_GS Then Log.Info "DataBodyRange: " & startRow & " - " & endRow

    ' 選択された行がテーブルのDataBodyRange内にあるかどうかを確認
    ReDim relativeRows(LBound(selectedRows) To UBound(selectedRows))
    cnt_j = LBound(relativeRows)
    For cnt_i = LBound(selectedRows) To UBound(selectedRows)
        If selectedRows(cnt_i) >= startRow And selectedRows(cnt_i) <= endRow Then
            relativeRows(cnt_j) = GetTableRelativeRow(tbl, selectedRows(cnt_i))
            If DBG_GS Then Log.Info "Relative Row: " & relativeRows(cnt_j)
            cnt_j = cnt_j + 1
        Else
            Log.Warn "Out of DataBodyRange: " & selectedRows(cnt_i)
            Err.Raise ERR_OUT_OF_TABLE_RANGE, "GetSelectedTableRows", MSG_OUT_OF_TABLE_RANGE
        End If
    Next cnt_i

    ReDim Preserve relativeRows(LBound(relativeRows) To cnt_j - 1)
    GetSelectedTableRows = relativeRows
    Exit Function

catch_err:
    HandleError Err
    GetSelectedTableRows = Array()
End Function

' シートの絶対行番号からテーブルの相対行番号を計算する。
Private Function GetTableRelativeRow(tbl As ListObject, sheetRow As Long) As Long
    GetTableRelativeRow = sheetRow - tbl.HeaderRowRange.row
End Function


Sub testGS()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(1)

    Log.Info "=== GS TEST ==="
    GetSelectedRow ws
    Stop
    GetSelectedTableRow ws, tbl
    Stop
    GetSelectedSheetRows ws
    Stop
    GetSelectedTableRows ws, tbl
    Stop
End Sub
