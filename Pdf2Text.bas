Option Explicit

Private Property Get fso() 'As FileSystemObject
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property

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

Function GetOrDownloadToolPath(ByVal ToolName As String, ByVal ServerToolFolderPath As String) As String
    Dim LocalBinFolderPath As String
    Dim LocalToolPath As String
    Dim ServerToolExePath As String

    LocalBinFolderPath = ThisWorkbook.path & Application.PathSeparator & "bin"
    LocalToolPath = LocalBinFolderPath & Application.PathSeparator & ToolName
    ServerToolExePath = ServerToolFolderPath & Application.PathSeparator & ToolName

    If Not fso.FileExists(LocalToolPath) Then
        If fso.FileExists(ServerToolExePath) Then
            CopyTool ServerToolExePath, LocalToolPath, ToolName, GetOrDownloadToolPath
        Else
            MsgBox ToolName & " could not be found at the following server location: " & ServerToolExePath & vbCrLf & _
                   "Please manually select the file.", vbInformation, "Error: " & ToolName & " not found"
            With Application.FileDialog(msoFileDialogFilePicker)
                .title = "Select " & ToolName
                .Filters.Clear
                .Filters.Add "Executable Files", "*.exe"
                .AllowMultiSelect = False
                If .Show = -1 Then
                    CopyTool .SelectedItems(1), LocalToolPath, ToolName, GetOrDownloadToolPath
                Else
                    GetOrDownloadToolPath = ""
                    Exit Function
                End If
            End With
        End If
    Else
        GetOrDownloadToolPath = LocalToolPath
    End If
End Function


Sub CopyTool(ByVal source As String, ByVal Destination As String, ByVal ToolName As String, ByRef ToolExePath As String)
    Dim LocalBinFolderPath As String
    LocalBinFolderPath = Left(Destination, InStrRev(Destination, Application.PathSeparator) - 1)

    If fso.GetFileName(source) = ToolName Then
        CreateFolderEx LocalBinFolderPath

        On Error Resume Next
        fso.CopyFile source, Destination, False
        If Err.Number <> 0 Then
            MsgBox "Error copying " & ToolName & " from source to destination: " & Err.Description, vbCritical, "Error: Copying Tool"
            ToolExePath = ""
        Else
            ToolExePath = Destination
        End If
        On Error GoTo 0
    Else
        MsgBox "The selected file is not " & ToolName & ". Please select the correct file.", vbCritical, "Error: Invalid File"
        ToolExePath = ""
    End If
End Sub


Function GetPDFFilePath(ByVal TempFolderPath As String) As String
    Dim OpenFileDialogName As Variant
    Dim TempPDFPath As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "依頼表を選択してください"
        .Filters.Clear
        .Filters.Add "PDF Files", "*.pdf"
        .AllowMultiSelect = False
        If .Show = -1 Then
            OpenFileDialogName = .SelectedItems(1)
        Else
            GetPDFFilePath = ""
            Exit Function
        End If
    End With

    TempPDFPath = TempFolderPath & Application.PathSeparator & fso.GetFileName(OpenFileDialogName)
    fso.CopyFile OpenFileDialogName, TempPDFPath, False

    GetPDFFilePath = TempPDFPath
End Function

' TODO カレントディレクトリに付属フォルダを作成するように作り直す。つまり"temp"を引数に取るようにする。
' CreateNewはデフォルト=Falseで
' ファイルの作成に再帰処理を導入
Function GetTempFolderPath(ByVal CreateNew As Boolean) As String
    Dim TempFolderPath As String

    TempFolderPath = ThisWorkbook.path & Application.PathSeparator & "temp"

    If CreateNew Then
        If fso.FolderExists(TempFolderPath) Then
            fso.DeleteFolder TempFolderPath
        End If
        fso.CreateFolder TempFolderPath
    Else
        If Not fso.FolderExists(TempFolderPath) Then
            MsgBox "Error: Temp folder not found.", vbCritical, "Error"
            TempFolderPath = ""
            Exit Function
        End If
    End If

    GetTempFolderPath = TempFolderPath
End Function


Function ConvertPDFtoText(ByVal ws As Worksheet, _
                          ByVal PDFFilePath As String, _
                          ByVal PdftotextExePath As String, _
                          ByVal TempFolderPath As String) As String
    On Error GoTo ErrorHandler

    Dim TempOutputFile As String      ' テキストファイルを出力する一時フォルダ
    Dim InputFileName As String       ' 目的のPDFファイル名
    Dim PdftotextCommand As String    ' pdftotextツールの実行コマンド
    Dim ExitCode As Integer           ' pdftotextツールの終了コード
    Dim WshShell As Object            ' コマンド実行のためのWindows Script Host Shellオブジェクト
    Dim WshExec As Object             ' コマンド管理用Windows Script Host Execオブジェクト

    ' PDFのパスからファイル名を抽出する
    InputFileName = Mid(PDFFilePath, InStrRev(PDFFilePath, Application.PathSeparator) + 1)
    InputFileName = Left(InputFileName, InStrRev(InputFileName, ".") - 1)

    ' D1セルにPDFファイル名を書き込む TODO: お前の居場所はここじゃない
    ws.range("D1").value = InputFileName & ".pdf"
    ' 呼び出し元で、TempOutputFileからファイル名を取得すればよい

    ' テキストファイルを作成
    TempOutputFile = TempFolderPath & Application.PathSeparator & InputFileName & ".txt"

    ' 実行コマンドを作成
    PdftotextCommand = """" & PdftotextExePath & """ -enc UTF-8 -raw """ & PDFFilePath & """ """ & TempOutputFile & """"

    ' コマンドを実行し、終了コードを取得する
    Set WshShell = CreateObject("WScript.Shell")
    Set WshExec = WshShell.Exec(PdftotextCommand)
    WshExec.StdOut.ReadAll ' コマンドが完了するのを待つ
    ExitCode = WshExec.ExitCode

    ' エラーの有無を確認し、適切なメッセージを表示する
    Select Case ExitCode
        Case 1
            MsgBox "Error opening the PDF file."
            Exit Function
        Case 2
            MsgBox "Error opening the output file."
            Exit Function
        Case 3
            MsgBox "Error related to PDF permissions."
            Exit Function
        Case 99
            MsgBox "An unknown error occurred."
            Exit Function
        Case Else
            If ExitCode <> 0 Then
                If MsgBox("An unexpected error occurred. Exit code: " & ExitCode & vbCrLf & _
                          "Do you want to continue processing?", vbYesNo) = vbNo Then
                    Exit Function
                End If
            End If
    End Select

    ' 変換が完了したテキストファイルのパスを返す
    ConvertPDFtoText = TempOutputFile

    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    Exit Function
End Function
