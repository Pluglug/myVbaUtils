Option Explicit

' Private fso As Object

' // 標準モジュール用
Private Property Get fso() 'As FileSystemObject
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property

' Private Sub Class_Initialize()
'     Set fso = CreateObject("Scripting.FileSystemObject")
' End Sub


' // ファイル選択ダイアログを表示し、選択されたファイルのパスを返すメソッド
' // multiSelect: マルチセレクトを許可するかどうか (Boolean)
' // fileFilter: ファイルフィルタ (例: "Excel Files,*.xlsx;*.xls")
' // dialogTitle: ダイアログのタイトル (省略可)
' // initialFolder: 初期フォルダ (省略可)
Public Function GetSelectedFiles(Optional ByVal multiSelect As Boolean = True, _
                                 Optional ByVal fileFilter As String = "All Files,*.*", _
                                 Optional ByVal dialogTitle As String = "ファイルを選択してください", _
                                 Optional ByVal initialFolder As String = "") As Variant
    Dim fd As FileDialog
    Dim selectedFiles() As String
    Dim cnt_i As Long
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = dialogTitle
        .AllowMultiSelect = multiSelect
        .Filters.Clear
        .Filters.Add Split(fileFilter, ",")(0), Split(fileFilter, ",")(1)
        If initialFolder <> "" Then .InitialFileName = initialFolder
        
        If .Show = -1 Then
            ReDim selectedFiles(1 To .SelectedItems.Count)
            For cnt_i = 1 To .SelectedItems.Count
                selectedFiles(cnt_i) = .SelectedItems(cnt_i)
            Next cnt_i
            GetSelectedFiles = selectedFiles
        Else
            GetSelectedFiles = Array()
        End If
    End With
End Function


' ' // ファイルの存在チェック
' Public Function FileExists(ByVal filePath As String) As Boolean
'     FileExists = fso.FileExists(filePath)
' End Function


'// パスに含まれる全てのフォルダの存在確認をしてフォルダを作る関数
'// 使用例
'// CreateFolderEx "C:\a\b\c\d\e\"
Public Sub CreateFolderEx(path_folder As String)
    '// 親フォルダが遡れなくなるところまで再帰で辿る
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    '// 途中の存在しないフォルダを作成しながら降りてくる
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub

'// ファイルの存在確認をしてファイルをコピーする関数
'// 使用例
'// CopyFileEx "C:\a\b\c\d\e\f.txt", "C:\a\b\c\d\e\g.txt"
Public Sub CopyFileEx(path_src As String, path_dst As String)
    If fso.FileExists(path_src) Then
        CreateFolderEx fso.GetParentFolderName(path_dst)
        fso.CopyFile path_src, path_dst
    End If
End Sub



' 任意のフォルダ内の作成日が古いファイルを削除
Sub DeleteOldFilesInFolder(folderPath As String, Optional days As Long = 7)
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    Dim file As Object
    For Each file In folder.Files
        If DateDiff("d", file.DateCreated, Now) > days Then
            Debug.Print "Delete Old File: " & file.path
            file.Delete
        End If
    Next file
End Sub
