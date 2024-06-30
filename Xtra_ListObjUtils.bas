Option Explicit


' テーブルに設定された関数、書式設定等を維持してテーブルを初期化
Sub ClearTable(ByRef Table As ListObject)
    Dim ans As Long
    ans = MsgBox("テーブルを初期化します" & vbCrLf & _
              "(この操作は取り消せません)", _
              vbOKCancel + vbExclamation + vbDefaultButton2, "テーブル初期化")
    If ans = vbOK Then
        ' テーブルの最後に空の行を追加
        Table.ListRows.Add
        ' 最後の行を除いて、他のすべての行を削除
        Table.DataBodyRange.Rows(1).Resize(Table.ListRows.Count - 1).Delete
    End If
End Sub


' マッピングに便利
Public Function GetDictFromTable(ByRef targetTable As ListObject, _
                                 ByVal keyCol As String, _
                                 ByVal valueCol As String) ' As Dictionary
    ' 指定されたキーと値の列を使って辞書を作成する
    ' targetTable: 対象のテーブル
    ' keyCol: キーとなる列の見出し名
    ' valueCol: 値となる列の見出し名

    Dim cnt_i As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For cnt_i = 1 To targetTable.ListRows.Count
        Dim key As Variant
        key = targetTable.ListColumns(keyCol).DataBodyRange(cnt_i).value
        If IsEmpty(key) Then GoTo Continue

        If dict.exists(key) Then
            Err.Raise vbObjectError + 513, _
                "GetDictFromTable", "重複する見出しが存在します: " & key
            Exit Function
        Else
            dict.Add key, targetTable.ListColumns(valueCol).DataBodyRange(cnt_i).value
        End If
Continue:
    Next cnt_i

    Set GetDictFromTable = dict
End Function


' #############################legacy_DataCopyUtils(非推奨)#############################
' TODO: フィルタ考慮 既存のデータに上書きするか、追加するかを選択できるようにする 非表示行列に注意

' 一致する見出し名をすべてコピー
Sub legacy_CopyMatchingColumns(sourceTable As ListObject, targetTable As ListObject)
    Dim col As ListColumn

    For Each col In sourceTable.ListColumns
        If ColumnExists(targetTable, col.name) Then
            CopyColumnData sourceTable, targetTable, col.name, col.name
        End If
    Next col
End Sub


' 一致する見出し名かつ、指定された列のみをコピー
' usage:
'Dim colNames() As String
'colNames = Split("列1,列2,列3", ",")
'CopySpecifiedColumns SourceTable, TargetTable, colNames
Sub legacy_CopySpecifiedColumns(sourceTable As ListObject, targetTable As ListObject, ColumnNames() As String)
    Dim colName As Variant

    For Each colName In ColumnNames
        If ColumnExists(sourceTable, colName) And ColumnExists(targetTable, colName) Then
            CopyColumnData sourceTable, targetTable, colName, colName
        End If
    Next colName
End Sub

' 異なる見出し名を持っている列間でのコピー
' usage:
'Dim colMapping As Dictionary: Set colMapping = New Dictionary
''' "列1"から"ColumnA"にコピー
'colMapping.Add "列1", "ColumnA"
'colMapping.Add "列2", "ColumnB"
''CopyMappedColumns SourceTable, TargetTable, colMapping
Sub legacy_CopyMappedColumns(sourceTable As ListObject, targetTable As ListObject, ColumnMapping As Dictionary)
    Dim SourceColName As Variant
    Dim TargetColName As String

    For Each SourceColName In ColumnMapping.Keys
        TargetColName = ColumnMapping(SourceColName)

        If ColumnExists(sourceTable, SourceColName) And ColumnExists(targetTable, TargetColName) Then
            CopyColumnData sourceTable, targetTable, SourceColName, TargetColName
        End If
    Next SourceColName
End Sub


' 列データのコピー処理
Private Sub CopyColumnData(sourceTable As ListObject, targetTable As ListObject, SourceColName As Variant, TargetColName As Variant)
    sourceTable.ListColumns(SourceColName).DataBodyRange.Copy targetTable.ListColumns(TargetColName).DataBodyRange
End Sub

' 列が存在するかをチェック
Private Function ColumnExists(Table As ListObject, ColumnName As Variant) As Boolean
    On Error Resume Next
    ColumnExists = Not IsEmpty(Table.ListColumns(ColumnName))
    On Error GoTo 0
End Function
