Option Explicit

' // 導入方法
' // 1. Xtra_Log.clsをプロジェクトに追加
' // 2. Xtra_DebugUtils.bas(このファイル)をプロジェクトに追加
' // 3. プロジェクト内で共通のフラグとLogが使えるようになります

' // 使い方
' // 1. Log.Info "シンプルな出力"
' // 2. If DBG_XXX Then Log.Warn "指定したフラグのみ出力"
' // 3. If DBG_XXX Then Log.Sep: Log.Info "区切り線を挿入して出力"

Public Const DBG_INIT As Boolean = False
Public Const DBG_RENAME As Boolean = False
Public Const DBG_OPS As Boolean = False
Public Const DBG_JSON As Boolean = False
Public Const DBG_PREFS As Boolean = False
Public Const DBG_PROP As Boolean = False
Public Const DBG_SHEET_A As Boolean = True

Public Property Get Log() As Xtra_Log
    Static obj As Xtra_Log
    If obj Is Nothing Then
        Set obj = New Xtra_Log
        ' obj.init  ' // 機能追加時に必要であれば
    End If
    Set Log = obj
End Property
