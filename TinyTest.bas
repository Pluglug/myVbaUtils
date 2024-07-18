' TinySeleniumVBA
' A tiny Selenium wrapper written in pure VBA
'
' (c)2021 uezo
'
' Mail: uezo@uezo.net
' Twitter: @uezochan
' https://github.com/uezo/TinySeleniumVBA
'
' ==========================================================================
' セットアップ
'
' 1. ツール＞参照設定で`Microsoft Scripting Runtime`をオンにする
'
' 2. WebDriver.cls, WebElement.cls JsonConverter.bas をプロジェクトに追加
'
' 3. WebDriverをダウンロード（ブラウザのメジャーバージョンと同じもの）
'   - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'   - Chrome: https://chromedriver.chromium.org/downloads
'
' 使い方
'    `WebDriver`のインスタンスをダウンロードしたWebDriverを使って生成します。
'    そこから先は下のExampleを参照ください。
' ==========================================================================

' ==========================================================================
' Setup
'
' 1. Set reference to `Microsoft Scripting Runtime`
'
' 2. Add WebDriver.cls, WebElement.cls and JsonConverter.bas to your VBA Project
'
' 3. Download WebDriver (driver and browser should be the same version)
'   - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'   - Chrome: https://chromedriver.chromium.org/downloads
'
' Usase
'    Create instance of `WebDriver` with the path to the driver you download.
'    See also the example below.
' ==========================================================================


' ==========================================================================
' Example
' ==========================================================================
Option Explicit

Const EDGE_DRIVER_PASS As String = "C:\Users\upwar\Documents\WebDriver\edgedriver.exe"
Private mSessionId As String

Public Sub main()
    ' Start WebDriver (Edge)
    Dim Driver As New WebDriver
    SafeOpen Driver, Edge
    
    rBoxLogin Driver, 

    Stop
    
    ' Get search textbox
    Dim searchInput
    Set searchInput = Driver.FindElement(by.Name, "q")
    
    ' Get value from textbox
    Debug.Print searchInput.GetValue
    
    ' Set value to textbox
    searchInput.SetValue "yomoda soba"
    
    ' Click search button
    Driver.FindElement(by.Name, "btnK").Click
    
    ' Refresh - you can use Execute with driver command even if the method is not provided
    Driver.Execute Driver.CMD_REFRESH
End Sub


Sub headlessEdgeTest()
    Dim Driver As New WebDriver
    Driver.Edge EDGE_DRIVER_PASS

    ' Capabilities を作成
    Dim cap As Capabilities
    Set cap = Driver.CreateCapabilities()
    
    ' ヘッドレスモードを設定
    cap.AddArgument "--headless"
    
    ' デバッグ用に Capabilities を JSON で表示
    Debug.Print cap.ToJson()

    ' ブラウザを開く
    Dim sessionId As String
    sessionId = Driver.OpenBrowser(cap)

    ' ナビゲーション
    Driver.Navigate "https://www.example.com", sessionId
    Stop
    Driver.CloseBrowser
End Sub

Sub multiTest()
    Dim Driver As New WebDriver
    Driver.Edge EDGE_DRIVER_PASS  ' EdgeDriverのパスを指定

    ' 1つ目のブラウザ用の Capabilities
    Dim cap1 As Capabilities
    Set cap1 = Driver.CreateCapabilities()
    cap1.AddArgument "--start-maximized"

    ' 2つ目のブラウザ用の Capabilities
    Dim cap2 As Capabilities
    Set cap2 = Driver.CreateCapabilities()
    cap2.AddArgument "--headless"

    ' ブラウザを開く
    Dim session1 As String, session2 As String
    session1 = Driver.OpenBrowser(cap1)
    session2 = Driver.OpenBrowser(cap2)

    ' 各セッションでナビゲーション
    Driver.Navigate "https://www.example.com", session1
    Driver.Navigate "https://www.example.org", session2
    
End Sub

Sub OpenBrowserAndNavigate()
    Dim Driver As New WebDriver
    Driver.Edge EDGE_DRIVER_PASS  ' ChromeDriverのパスを適切に設定してください

    ' Capabilities を作成
    Dim cap As Capabilities
    Set cap = Driver.CreateCapabilities()
    
    ' ブラウザを開く
    mSessionId = Driver.OpenBrowser(cap)

    ' 最初のページに遷移
    Driver.Navigate "https://www.example.com", mSessionId
    
    MsgBox "最初のページを開きました。次のプロシージャを実行してください。"
End Sub

Sub NavigateAndClose()
    If mSessionId = "" Then
        MsgBox "セッションIDが見つかりません。最初のプロシージャを実行してください。"
        Exit Sub
    End If

    Dim Driver As New WebDriver
    Driver.Edge EDGE_DRIVER_PASS  ' ChromeDriverのパスを適切に設定してください

    ' 別のページに遷移
    Driver.Navigate "https://www.example.org", mSessionId
    
    MsgBox "2つ目のページに遷移しました。OKをクリックするとブラウザを閉じます。"

    ' ブラウザを閉じる
    Driver.Execute Driver.CMD_QUIT, CreateParams("sessionId", mSessionId)

    ' セッションIDをクリア
    mSessionId = ""
    
    MsgBox "ブラウザを閉じました。"
End Sub

Public Function CreateParams(ParamArray args() As Variant) As Dictionary
'USAGE
'Dim params As Dictionary
'Set params = CreateParams("sessionId", mSessionId, "timeout", 30000, "url", "https://www.example.com")
    Dim params As New Dictionary
    Dim i As Long
    
    For i = LBound(args) To UBound(args) Step 2
        If i + 1 <= UBound(args) Then
            params.Add CStr(args(i)), args(i + 1)
        End If
    Next i
    
    Set CreateParams = params
End Function









Private Sub rBoxLogin(ByRef Driver As WebDriver, ByVal mail As String, ByVal pass As String)
    Driver.Navigate "https://rak.account.box.com/login"
    ' Log.Info driver.Execute(driver.CMD_GET_TITLE)  ' "Box | ログイン"
    Driver.FindElement(by.XPath, "/html/body/div/div[2]/div/div[1]/div/div[1]/form/button").Click

    ' Log.Info driver.Execute(driver.CMD_GET_TITLE) ' "Rakuten Global - サインイン"
    Dim mailInput As WebElement
    Set mailInput = WaitForElement(Driver, by.ID, "input43", 10)
    mailInput.SetValue mail
    Driver.FindElement(by.XPath, "//*[@id=""form35""]/div[2]/input").Click
    
    ' Log.Info driver.Execute(driver.CMD_GET_TITLE) ' "Rakuten Global - サインイン"
    Dim passInput As WebElement
    Set passInput = WaitForElement(by.Name, "credentials.passcode")
    passInput.SetValue pass
    Driver.FindElement(by.XPath, "//*[@id=""form21""]/div[2]/input").Click
End Sub

Public Function WaitForElement(Driver As WebDriver, by As by, value As String, Optional timeoutSeconds As Long = 10) As WebElement
    Dim startTime As Date
    startTime = Now
    
    Do While Now < startTime + TimeSerial(0, 0, timeoutSeconds)
        On Error Resume Next
        Set WaitForElement = Driver.FindElement(by, value)
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
        
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    
    Err.Raise 513, "WaitForElement", "Timeout waiting for element: " & value
End Function
