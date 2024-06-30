Option Explicit

' // 導入方法
' // 1. Xtra_Log.cls(このファイル)をプロジェクトに追加
' // 2. Xtra_DebugUtils.basをプロジェクトに追加
' // 3. プロジェクト内で共通のフラグとLogが使えるようになります

' // 使い方
' // 1. Log.Info "シンプルな出力"
' // 2. If DBG_XXX Then Log.Warn "指定したフラグのみ出力"
' // 3. If DBG_XXX Then Log.Sep: Log.Info "区切り線を挿入して出力"

Private Function createLogMsg(ByVal message As String, ByVal level As String) As String
    Dim line As String
    ' line = Format(Now(), "yyyy/mm/dd") & " "  ' // 不要ならばコメントアウト
    line = Format(Now(), "hh:mm:ss") & " - "
    line = line & level & " - "
    createLogMsg = line & message
End Function

Private Sub outputLogMsg(ByVal message As String)
    On Error Resume Next
    Debug.Print message
    On Error GoTo 0
End Sub

Private Sub LogMsg(ByVal message As String, ByVal level As String)
    outputLogMsg createLogMsg(message, level)
End Sub

' // ログ出力メソッド

Public Sub Test(ByVal message As String)
    LogMsg message, "TEST"
End Sub

Public Sub Info(ByVal message As String)
    LogMsg message, "INFO"
End Sub

Public Sub Warn(ByVal message As String)
    LogMsg message, "WARN"
End Sub

Public Sub Error(ByVal message As String)
    LogMsg message, "ERROR"
End Sub

Public Sub Fatal(ByVal message As String)
    LogMsg message, "FATAL"
End Sub

Public Sub Sep(Optional ByVal length As Long = 50, Optional ByVal sepChar As String = "-")
    outputLogMsg vbCrLf & String(length, sepChar)
End Sub

' // 初期化処理

Public Sub init(Optional ByVal DoNotUse As Boolean = False)
    ' // ファイル出力等の初期化処理を記述
    outputLogMsg vbCrLf & "=== Init Log ===" & vbCrLf
End Sub

Private Sub Class_Initialize()
    Me.init
End Sub
