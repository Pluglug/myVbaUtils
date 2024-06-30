
Option Explicit

Private WithEvents wb As Excel.Workbook
Private WithEvents ws As Excel.Worksheet
Private pTblTemplates As Excel.ListObject

Private Property Get ActiveWs() As Excel.Worksheet
    Set ActiveWs = ws
End Property

Private Property Let ActiveWs(ByRef Sh As Excel.Worksheet)
    Set ws = Sh
End Property

Private Property Get tblMailData() As ListObject
    Set tblMailData = ActiveWs.ListObjects(1)
End Property

Private Property Get tblTemplates() As ListObject
    Set tblTemplates = ThisWorkbook.Sheets(ActiveWs.name).TemplateTable
End Property

Private Sub UserForm_Initialize()
    Set wb = Excel.ActiveWorkbook
    Set ws = Excel.ActiveSheet
    Me.chkMultiSelect.value = ThisWorkbook.mtsCheckBoxState
    Me.Caption = "メールテンプレート追加 【" & ws.name & "】"
    LoadTemplateData
End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.mtsCheckBoxState = Me.chkMultiSelect.value
End Sub

Private Sub wb_SheetActivate(ByVal Sh As Object)
    Dim sheetCodeName As String
    sheetCodeName = Sh.CodeName
    
    If Not Sh Is ActiveWs And Left(sheetCodeName, 11) = "shtMailForm" Then
        ActiveWs = Sh
        Me.Caption = "メールテンプレート追加 【" & ActiveWs.name & "】"
        LoadTemplateData
    End If
End Sub

Private Sub LoadTemplateData()

    Dim rngCategory As range, rngID As range, rngSubID As range, rngTemplate As range

    ' 各列のデータを取得
    Set rngCategory = tblTemplates.ListColumns("追加先").DataBodyRange
    Set rngID = tblTemplates.ListColumns("ID").DataBodyRange
    Set rngSubID = tblTemplates.ListColumns("サブID").DataBodyRange
    Set rngTemplate = tblTemplates.ListColumns("テンプレート本文").DataBodyRange

    ' リストボックスにデータを反映
    Me.lstCategory.List = UniqueValues(rngCategory)
    ' カテゴリ以外は初期状態では空にしておく
    Me.lstID.Clear
    Me.lstSubID.Clear
    Me.lstTemplate.Clear

End Sub

Private Sub lstCategory_Change()
    Dim rngIDs As range, cell As range
    Dim dict As Object
    Dim key As Variant
    Dim colCategory As Integer
    Dim colID As Integer

    colCategory = GetColumnIndexByHeader(tblTemplates, "追加先")
    colID = GetColumnIndexByHeader(tblTemplates, "ID")
    Set rngIDs = tblTemplates.ListColumns("ID").DataBodyRange
    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In rngIDs
        If cell.Offset(0, colCategory - colID).value = Me.lstCategory.value And Not dict.exists(cell.value) Then
            dict(cell.value) = 1
        End If
    Next cell

    Me.lstID.Clear
    Me.lstSubID.Clear
    Me.lstTemplate.Clear
    For Each key In dict.Keys
        Me.lstID.AddItem key
    Next key
End Sub

Private Sub lstID_Change()
    Dim rngSubIDs As range, cell As range
    Dim dict As Object
    Dim key As Variant
    Dim colCategory As Integer
    Dim colID As Integer
    Dim colSubID As Integer

    colCategory = GetColumnIndexByHeader(tblTemplates, "追加先")
    colID = GetColumnIndexByHeader(tblTemplates, "ID")
    colSubID = GetColumnIndexByHeader(tblTemplates, "サブID")

    Set rngSubIDs = tblTemplates.ListColumns("サブID").DataBodyRange
    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In rngSubIDs
        If cell.Offset(0, colCategory - colSubID).value = Me.lstCategory.value And _
           cell.Offset(0, colID - colSubID).value = Me.lstID.value And _
           Not dict.exists(cell.value) Then
            dict(cell.value) = 1
        End If
    Next cell

    Me.lstSubID.Clear
    Me.lstTemplate.Clear

    For Each key In dict.Keys
        If Trim(key) = "" Then
            Me.lstSubID.AddItem "(サブIDなし)"
            Me.lstSubID.ListIndex = 0
        Else
            Me.lstSubID.AddItem key
        End If
    Next key
End Sub

Private Sub lstSubID_Change()

    Dim rngTemplates As range, cell As range

    Set rngTemplates = tblTemplates.ListColumns("テンプレート本文").DataBodyRange

    Dim colNumber As Long
    Dim colCategory As Integer
    Dim colID As Integer
    Dim colSubID As Integer
    Dim colTemplate As Integer

    colNumber = GetColumnIndexByHeader(tblTemplates, "項番")
    
    colCategory = GetColumnIndexByHeader(tblTemplates, "追加先")
    colID = GetColumnIndexByHeader(tblTemplates, "ID")
    colSubID = GetColumnIndexByHeader(tblTemplates, "サブID")
    colTemplate = GetColumnIndexByHeader(tblTemplates, "テンプレート本文")

    Me.lstTemplate.Clear

    For Each cell In rngTemplates
        ' "(サブIDなし)"の場合、空欄のサブIDとマッチさせる
        Dim matchSubID As Variant
        If Me.lstSubID.value = "(サブIDなし)" Then
            matchSubID = ""
        Else
            matchSubID = Me.lstSubID.value
        End If

        Dim itemID As String
        Dim itemTemplate As String

        If cell.Offset(0, colCategory - colTemplate).value = Me.lstCategory.value _
            And cell.Offset(0, colID - colTemplate).value = Me.lstID.value _
            And cell.Offset(0, colSubID - colTemplate).value = matchSubID Then

            itemID = cell.Offset(0, colNumber - colTemplate).value
            itemTemplate = cell.value

            ' テンプレート本文と項目IDをリストボックスに追加
            Me.lstTemplate.AddItem itemTemplate
            Me.lstTemplate.List(Me.lstTemplate.ListCount - 1, 1) = itemID
        End If
    Next cell
End Sub


Private Sub AddTemplate_Click()
    Dim targetTblRows As Variant
    Dim i As Long

    On Error GoTo GetRowErrorHandler
    targetTblRows = GetSelectedTableRows(ws, tblMailData)

    If Not Me.chkMultiSelect And UBound(targetTblRows) - LBound(targetTblRows) > 0 Then
        MsgBox "複数行への追加は許可されていません", vbInformation, "Error"
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Dim template As String
    Dim header As String
    Dim columnIndex As Integer

    ' DataBodyRangeの存在を確認
    If tblMailData.DataBodyRange Is Nothing Then
        MsgBox "テーブルにデータが存在しません。", vbExclamation, "エラー"
        Exit Sub
    End If

    With Me.lstTemplate
        If .ListIndex = -1 Then
            MsgBox "テンプレートが未選択です"
            Exit Sub
        End If

        template = .List(.ListIndex, 0)
        header = Me.lstCategory.List(Me.lstCategory.ListIndex, 0)

        columnIndex = GetColumnIndexByHeader(tblMailData, header)
        If columnIndex = 0 Then
            MsgBox "指定されたヘッダー名が存在しません。"
            Exit Sub
        End If
    End With

    On Error GoTo TableErrorHandler

    ' 選択された行ごとにテンプレートを追加
    For i = LBound(targetTblRows) To UBound(targetTblRows)
        tblMailData.ListColumns(columnIndex).DataBodyRange(targetTblRows(i)).value = _
            tblMailData.ListColumns(columnIndex).DataBodyRange(targetTblRows(i)).value & template & vbNewLine
    Next i

    Exit Sub

GetRowErrorHandler:
    Select Case Err.Number
        Case vbObjectError + 1000:
            MsgBox "単一のセルを選択してください", vbInformation, "GetRowError"
        Case vbObjectError + 1010:
            MsgBox "テーブルの範囲内を選択してください", vbInformation, "GetRowError"
        Case vbObjectError + 1011:
            MsgBox "テーブルにデータが存在しません。", vbInformation, "GetRowError"
        Case vbObjectError + 1012:
            MsgBox "テーブルの範囲外もしくは見出し行が選択されています", vbInformation, "GetRowError"
        Case Else:
            MsgBox "予期しないエラーが発生しました: " & Err.Description, vbCritical, "GetRowError"
    End Select

    Exit Sub

TableErrorHandler:
    MsgBox "予期しないエラーが発生しました: " & Err.Description, vbCritical, "TableError"
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "Error"
End Sub



' 重複を削除して一意の値の配列を返す関数
Private Function UniqueValues(rng As range) As Variant
    Dim cell As range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In rng
        dict(cell.value) = 1
    Next cell

    UniqueValues = Application.Transpose(dict.Keys)
End Function

Private Function GetColumnIndexByHeader(tbl As ListObject, headerName As String) As Integer
    Dim i As Integer
    For i = 1 To tbl.ListColumns.Count
        If tbl.ListColumns(i).name = headerName Then
            GetColumnIndexByHeader = i
            Exit Function
        End If
    Next i
    GetColumnIndexByHeader = 0 ' 見出し名が見つからない場合は0を返す
End Function
