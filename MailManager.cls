Option Explicit

Private pTblMailData As ListObject
Private pTblAddress As ListObject

' tblMailData用の見出し名定数
Private Const HEADER_CREATION_HISTORY As String = "作成履歴"
Private Const HEADER_DESTINATION_GROUP As String = "宛先グループ"
Private Const HEADER_INDIVIDUAL_TO As String = "個別To"
Private Const HEADER_INDIVIDUAL_CC As String = "個別Cc"
Private Const HEADER_INDIVIDUAL_BCC As String = "個別Bcc"
Private Const HEADER_ATTACHMENT_PATH As String = "添付ファイルPath"

' tblAddress用の見出し名定数
Private Const HEADER_EMAIL_ADDRESS As String = "メールアドレス"


Private Property Get tblMailData() As ListObject
    '''作成するメール個々の情報を格納したテーブル
    ' 必須条件: 以下の見出し名を持つ6列が存在する
    ' "作成履歴": 過去に作成したメールの再作成を防ぐ 値の有無で識別
    ' "宛先グループ": tblAddressで設定したメールグループ名(見出し名)を指定
    ' "個別To", "個別Cc", "個別Bcc":
    '     宛先グループに加えて指定するアドレス 複数アドレスは`;`(セミコロン)で区切る
    ' "添付ファイルPath": 絶対パス 詳細はmodAddAttachmentPathsを参照されたい
    Set tblMailData = pTblMailData
End Property

Private Property Let tblMailData(val As ListObject)
    If Not CheckRequiredHeaders(val, Array(HEADER_CREATION_HISTORY, _
                                           HEADER_DESTINATION_GROUP, _
                                           HEADER_INDIVIDUAL_TO, _
                                           HEADER_INDIVIDUAL_CC, _
                                           HEADER_INDIVIDUAL_BCC, _
                                           HEADER_ATTACHMENT_PATH)) Then
        Err.Raise Number:=vbObjectError + 513, _
                  Description:="tblMailDataに必要な見出し名がありません。"
    End If
    Set pTblMailData = val
End Property

Private Property Get tblAddress() As ListObject
    '''メール送信先をグループ管理するテーブル
    ' メールグループを設定した列と、対応するメールアドレスを格納した行を持つ
    ' 必須条件: 見出し名"メールアドレス"列が存在し、
    '           それより右に１列以上データ(グループ)が存在する
    ' 詳細はGetAddressesByGroupNameを参照されたい
    Set tblAddress = pTblAddress
End Property

Private Property Let tblAddress(val As ListObject)
    If Not CheckRequiredHeaders(val, Array(HEADER_EMAIL_ADDRESS)) Then
        Err.Raise Number:=vbObjectError + 514, _
                  Description:="tblAddressに必須の見出し名がありません。"
    End If
    Set pTblAddress = val
End Property

Private Property Get fso() As Object
    Static obj As Object
    If obj Is Nothing Then Set obj = _
            CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property

Private Function olCreateItem() As Object
    On Error Resume Next
    Dim oApp As Object
    Set oApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    If oApp Is Nothing Then
        Set oApp = CreateObject("Outlook.Application")
        ' 未立ち上げの場合は、Inboxフォルダを選択してアプリを開く
        oApp.GetNamespace("MAPI").GetDefaultFolder(6).Display
    End If
    
    Set olCreateItem = oApp.CreateItem(0)
End Function


Public Sub Initialize(MailDataTbl As ListObject, _
                      AddressGpTbl As ListObject)
    tblMailData = MailDataTbl
    tblAddress = AddressGpTbl
End Sub


' 処理対象の行のリストを作成する
Private Function CreateProcessList(tbl As ListObject) As Collection
    Dim rowList As Collection
    Dim rowCounter As Long
    Dim colCounter As Long

    ' DataBodyRangeの存在を確認
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "テーブルにデータが存在しません。", vbExclamation, "エラー"
        Set CreateProcessList = Nothing
        Exit Function
    End If

    ' cells.textを使うと非表示のセルの値を取得できないため、
    ' すべての列を一時的に表示する
    Dim su As Boolean: su = Application.ScreenUpdating
    Application.ScreenUpdating = False  ' 画面更新を一時的にオフ
    Dim columnsToHideAgain As Variant
    columnsToHideAgain = TemporarilyShowHiddenColumns(tbl)

    ' 処理対象のリストを作成
    Set rowList = New Collection

    ' 処理対象の行をrowListに追加
    For rowCounter = 1 To tbl.DataBodyRange.Rows.Count
        ' デフォルトは処理対象としてフラグをTrueに設定
        Dim toProcess As Boolean: toProcess = True
        
        ' 行がフィルタで非表示の場合、スキップする
        If tbl.DataBodyRange.Rows(rowCounter).EntireRow.Hidden Then
            toProcess = False
        End If
    
        ' "作成履歴"が空欄でない場合、スキップする
        If tbl.ListColumns(HEADER_CREATION_HISTORY). _
                    DataBodyRange.Cells(rowCounter).value <> "" Then
            toProcess = False
        End If

        ' その他のスキップ条件を追加する場合はここに追加

        ' toProcessフラグがTrueの場合、データをrowListコレクションに追加
        If toProcess Then
            Dim rowData As Object
            Set rowData = CreateObject("Scripting.Dictionary")
            For colCounter = 1 To tbl.DataBodyRange.Columns.Count
                rowData.Add tbl.ListColumns(colCounter).name, _
                            tbl.DataBodyRange.Cells(rowCounter, colCounter).text
            Next colCounter

            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            item.Add "Data", rowData
            item.Add "RowNumber", rowCounter
            rowList.Add item
        End If
    Next rowCounter

    ' 非表示にした列を再度非表示にする
    HideColumns tbl, columnsToHideAgain
    Application.ScreenUpdating = su
    
    ' rowListに項目が追加されていない場合、エラーメッセージを表示して処理を中止
    If rowList.Count = 0 Then
        MsgBox "処理対象のデータが存在しません。", vbExclamation, "エラー"
        Set CreateProcessList = Nothing
        Exit Function
    End If
    
    ' Debug
'    For Each item In rowList
'        Dim data As Object
'        Set data = item("Data")
'        Debug.Print data("#店名#")
'    Next item

    Set CreateProcessList = rowList
End Function

Function TemporarilyShowHiddenColumns(tbl As ListObject) As Variant
    Dim hiddenColumns() As Variant
    Dim colIndex As Long
    Dim hiddenCount As Long
    hiddenCount = 0
    
    ' 非表示列のインデックスを格納する配列の初期サイズを設定
    ReDim hiddenColumns(1 To tbl.ListColumns.Count)
    
    For colIndex = 1 To tbl.ListColumns.Count
        With tbl.ListColumns(colIndex)
            If .range.EntireColumn.Hidden Then
                ' 非表示の列を一時的に表示
                .range.EntireColumn.Hidden = False
                
                ' 非表示だった列のインデックスを記録
                hiddenCount = hiddenCount + 1
                hiddenColumns(hiddenCount) = colIndex
            End If
        End With
    Next colIndex
    
    ' 非表示だった列があれば配列サイズを調整
    If hiddenCount > 0 Then
        ReDim Preserve hiddenColumns(1 To hiddenCount)
    Else
        ' 非表示列がなければ空の配列を返す
        ReDim hiddenColumns(0 To 0)
    End If

    TemporarilyShowHiddenColumns = hiddenColumns
End Function

Sub HideColumns(tbl As ListObject, columnsToHide As Variant)
    If UBound(columnsToHide) = 0 Then Exit Sub

    Dim colIndex As Variant
    For Each colIndex In columnsToHide
        tbl.ListColumns(colIndex).range.EntireColumn.Hidden = True
    Next colIndex
End Sub


' Pathの分割
Private Function ExtractPathsFromQuotedString(path As String) As Collection
    If path = "" Then
        Set ExtractPathsFromQuotedString = Nothing
        Exit Function
    End If

    Dim paths As Collection
    Dim individualPaths() As String
    Dim i As Long
    
    Set paths = New Collection
    individualPaths = Split(path, ",")
    For i = LBound(individualPaths) To UBound(individualPaths)
        individualPaths(i) = Replace(Trim(individualPaths(i)), """", "")
        paths.Add individualPaths(i)
    Next i

    Set ExtractPathsFromQuotedString = paths
End Function

' 添付ファイルの存在確認
Private Function CheckAttachFileExists(cll As Collection) As Collection
    Dim i As Long
    Dim missingFiles As Collection
    Dim warnMessage As String
    Dim currentRowData As Variant
    Dim attachmentPaths As String
    Dim individualPaths As Collection
    
    Set missingFiles = New Collection

    For Each currentRowData In cll
        ' attachmentPaths = currentRowData("Data")(pathIdx)
        attachmentPaths = currentRowData("Data")(HEADER_ATTACHMENT_PATH)
        If attachmentPaths <> "" Then
            Set individualPaths = ExtractPathsFromQuotedString(attachmentPaths)
            For i = 1 To individualPaths.Count
                If Not fso.FileExists(individualPaths(i)) Then
                    On Error Resume Next
                    missingFiles.Add individualPaths(i), individualPaths(i)
                    On Error GoTo 0
                    warnMessage = warnMessage & "行 " & _
                        currentRowData("RowNumber") & " のファイルが存在しません: " & _
                        individualPaths(i) & vbCrLf
                End If
            Next i
        End If
    Next currentRowData

    ' 存在しないファイルがあった場合、ユーザーに警告
    ' ユーザーが継続を選択した場合、存在しないファイルは無視される
    If warnMessage <> "" Then
        Dim response As VbMsgBoxResult
        response = MsgBox("以下のファイルが見つかりません。継続しますか?" & vbCrLf & _
                          "注意: 不在ファイルの添付はスキップされます" & vbCrLf & _
                          vbCrLf & _
                          warnMessage, vbExclamation + vbYesNo)
        If response = vbNo Then
            Err.Raise Number:=vbObjectError + 515, _
                      Description:="添付ファイルが存在しません。"
            Exit Function
        End If
    End If
    
    Set CheckAttachFileExists = missingFiles
End Function


Function IsTag(text) As Boolean
    If Left(text, 1) = "#" And Right(text, 1) = "#" Then
        IsTag = True
    Else
        IsTag = False
    End If
End Function

Function ReplaceTagsInData(dataDictionary As Object) As Object
    Dim key As Variant
    Dim replaceKey As Variant
    
    ' dataDictionary内のタグを置き換え
    For Each key In dataDictionary.Keys
        If IsTag(key) Then
            For Each replaceKey In dataDictionary.Keys
                If replaceKey <> key Then
                    dataDictionary(replaceKey) = Replace(dataDictionary(replaceKey), key, dataDictionary(key))
                End If
            Next replaceKey
        End If
    Next key
    
    Set ReplaceTagsInData = dataDictionary
End Function

Function ReplaceTagsInTemplates(ByVal targetTemplate As String, dataDictionary As Object) As String
    Dim key As Variant
    
    ' テンプレート内のタグを置き換え
    For Each key In dataDictionary.Keys
        If IsTag(key) Then
            ' Debug.Print "Key: " & key & " Value: " & dataDictionary(key)
            targetTemplate = Replace(targetTemplate, key, dataDictionary(key))
        End If
    Next key

    ReplaceTagsInTemplates = targetTemplate
End Function


Public Sub CreateMails(subjectTemplate As String, bodyTemplate As String)
    Dim i As Long
    
    Dim mailList As Collection
    Dim currentMailData As Object

'    On Error GoTo ErrHandler

    ' 処理データを取得
    Set mailList = CreateProcessList(tblMailData)
    If mailList Is Nothing Then Exit Sub

    ' 添付ファイルの存在確認
    Dim missingFiles As Collection
    Set missingFiles = CheckAttachFileExists(mailList)

    '---メールの作成
    For Each currentMailData In mailList
       
        Dim replacedData As Object
        Set replacedData = ReplaceTagsInData(currentMailData("Data"))

        Dim replacedSubject As String
        Dim replacedBody As String
        replacedSubject = ReplaceTagsInTemplates((subjectTemplate), replacedData)
        replacedBody = ReplaceTagsInTemplates((bodyTemplate), replacedData)  ' 値渡しを明示

        ' 新しいメールオブジェクトの取得
        Dim newMail As Object
        Set newMail = olCreateItem()

        Dim mailTo As String, mailCc As String, mailBcc As String
        If currentMailData("Data")(HEADER_DESTINATION_GROUP) <> "" Then
            ' グループ名が指定されている場合、グループ名を基にメールアドレスを取得
            mailTo = GetAddressesByGroupName( _
                CStr(currentMailData("Data")(HEADER_DESTINATION_GROUP)), tblAddress, "To")

            mailCc = GetAddressesByGroupName( _
                CStr(currentMailData("Data")(HEADER_DESTINATION_GROUP)), tblAddress, "Cc")

            mailBcc = GetAddressesByGroupName( _
                CStr(currentMailData("Data")(HEADER_DESTINATION_GROUP)), tblAddress, "Bcc")

            ' 個別宛先が指定されている場合、それを追加
            mailTo = mailTo & "; " & currentMailData("Data")(HEADER_INDIVIDUAL_TO)
            mailCc = mailCc & "; " & currentMailData("Data")(HEADER_INDIVIDUAL_CC)
            mailBcc = mailBcc & "; " & currentMailData("Data")(HEADER_INDIVIDUAL_BCC)
        Else
            ' グループ名が指定されていない場合、個別宛先をそのまま使用
            mailTo = currentMailData("Data")(HEADER_INDIVIDUAL_TO)
            mailCc = currentMailData("Data")(HEADER_INDIVIDUAL_CC)
            mailBcc = currentMailData("Data")(HEADER_INDIVIDUAL_BCC)
        End If

        ' メールの設定
        With newMail
            .BodyFormat = olFormatPlain  ' プレーンテキスト
            .To = mailTo
            .CC = mailCc
            .BCC = mailBcc
            .Subject = replacedSubject
            .Body = replacedBody

            ' 添付ファイルの追加
            Dim attachmentPaths As String
            Dim individualPaths As Collection
            attachmentPaths = currentMailData("Data")(HEADER_ATTACHMENT_PATH)
            If attachmentPaths <> "" Then
                Set individualPaths = ExtractPathsFromQuotedString(attachmentPaths)
                For i = 1 To individualPaths.Count
                    ' 存在しないファイルはスキップされる
                    If Not IsInCollection(missingFiles, individualPaths(i)) Then
                        .Attachments.Add individualPaths(i)
                    End If
                Next i
            End If
            
            .Display
        End With
        ' メールを作成したので"作成履歴"に現在の日時を入れる
        tblMailData.ListColumns(HEADER_CREATION_HISTORY) _
            .DataBodyRange.Cells(currentMailData("RowNumber")).value _
                = Format(Now, "yyyy/mm/dd hh:mm")
    Next currentMailData

'Exit Sub
'
'ErrHandler:
'    If Err.Number = vbObjectError + 515 Then
'        ' MsgBox "添付ファイルが存在しません。", vbExclamation, "エラー"
'        Exit Sub
'    Else
'        MsgBox "エラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "エラー"
'    End If
End Sub

' 事前に不正な添付ファイルパスをコレクションに格納し、メール作成時にその情報を使用するため
Private Function IsInCollection(cll As Collection, value As Variant) As Boolean
    Dim item As Variant
    On Error Resume Next
    item = cll(value)
    IsInCollection = Err.Number = 0
    On Error GoTo 0
End Function

Private Function GetAddressesByGroupName(Groupname As String, tblAddress As ListObject, addressType As String) As String
    Dim ResultAddresses As String
    Dim iCol As Integer
    Dim iRow As Long
    Dim SelectedGroupColumn As Integer
    Dim MailAddressColumn As Integer

    ' グループ名が空の場合、空の文字列を返す
    If Groupname = "" Then
        GetAddressesByGroupName = ""
        Exit Function
    End If
    
    ' グループ名の列とメールアドレス列を特定
    For iCol = 1 To tblAddress.ListColumns.Count
        If tblAddress.ListColumns(iCol).name = Groupname Then
            SelectedGroupColumn = iCol
        End If
        If tblAddress.ListColumns(iCol).name = HEADER_EMAIL_ADDRESS Then
            MailAddressColumn = iCol
        End If
    Next iCol

    ' 選択されたグループ名の列を基にメールアドレスを取得
    For iRow = 1 To tblAddress.ListRows.Count
        Dim addressSymbol As String
        addressSymbol = tblAddress.DataBodyRange(iRow, SelectedGroupColumn).value

        If (addressType = "To" And addressSymbol = "●") _
        Or (addressType = "Cc" And addressSymbol = "〇") _
        Or (addressType = "Bcc" And addressSymbol = "◇") Then
            ResultAddresses = ResultAddresses & ";" & tblAddress.DataBodyRange(iRow, MailAddressColumn).value
        End If
    Next iRow

    ' 望む文字列を返す
    If ResultAddresses <> "" Then
        GetAddressesByGroupName = Right(ResultAddresses, Len(ResultAddresses) - 1) '最初のセミコロンを取り除く
    Else
        GetAddressesByGroupName = ""
    End If

End Function

Private Function CheckRequiredHeaders(tbl As ListObject, headers As Variant) As Boolean
    ' テーブルに必要な見出し名がすべて存在するかチェックする
    Dim header As Variant
    Dim found As Boolean
    CheckRequiredHeaders = True
    
    For Each header In headers
        found = False
        Dim col As ListColumn
        For Each col In tbl.ListColumns
            If StrComp(col.name, header, vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next col
        If Not found Then
            CheckRequiredHeaders = False
            Exit Function
        End If
    Next header
End Function


