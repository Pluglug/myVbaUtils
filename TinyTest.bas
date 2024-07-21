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
    Dim driver As New WebDriver
    SafeOpen driver, Edge
    
    rBoxLogin driver, "ts-ko.kitajima@rakuten.com", "Fack_0ff!!!!"

    Stop
    
    ' Get search textbox
    Dim searchInput
    Set searchInput = driver.FindElement(by.Name, "q")
    
    ' Get value from textbox
    Debug.Print searchInput.GetValue
    
    ' Set value to textbox
    searchInput.SetValue "yomoda soba"
    
    ' Click search button
    driver.FindElement(by.Name, "btnK").Click
    
    ' Refresh - you can use Execute with driver command even if the method is not provided
    driver.Execute driver.CMD_REFRESH
End Sub


Sub headlessEdgeTest()
    Dim driver As New WebDriver
    driver.Edge EDGE_DRIVER_PASS

    ' Capabilities を作成
    Dim cap As Capabilities
    Set cap = driver.CreateCapabilities()
    
    ' ヘッドレスモードを設定
    cap.AddArgument "--headless"
    
    ' デバッグ用に Capabilities を JSON で表示
    Debug.Print cap.ToJson()

    ' ブラウザを開く
    Dim sessionId As String
    sessionId = driver.OpenBrowser(cap)

    ' ナビゲーション
    driver.Navigate "https://www.example.com", sessionId
    Stop
    driver.CloseBrowser
End Sub

Sub multiTest()
    Dim driver As New WebDriver
    driver.Edge EDGE_DRIVER_PASS  ' EdgeDriverのパスを指定

    ' 1つ目のブラウザ用の Capabilities
    Dim cap1 As Capabilities
    Set cap1 = driver.CreateCapabilities()
    cap1.AddArgument "--start-maximized"

    ' 2つ目のブラウザ用の Capabilities
    Dim cap2 As Capabilities
    Set cap2 = driver.CreateCapabilities()
    cap2.AddArgument "--headless"

    ' ブラウザを開く
    Dim session1 As String, session2 As String
    session1 = driver.OpenBrowser(cap1)
    session2 = driver.OpenBrowser(cap2)

    ' 各セッションでナビゲーション
    driver.Navigate "https://www.example.com", session1
    driver.Navigate "https://www.example.org", session2
    
End Sub

Sub OpenBrowserAndNavigate()
    Dim driver As New WebDriver
    driver.Edge EDGE_DRIVER_PASS  ' ChromeDriverのパスを適切に設定してください

    ' Capabilities を作成
    Dim cap As Capabilities
    Set cap = driver.CreateCapabilities()
    
    ' ブラウザを開く
    mSessionId = driver.OpenBrowser(cap)

    ' 最初のページに遷移
    driver.Navigate "https://www.example.com", mSessionId
    
    MsgBox "最初のページを開きました。次のプロシージャを実行してください。"
End Sub

Sub NavigateAndClose()
    If mSessionId = "" Then
        MsgBox "セッションIDが見つかりません。最初のプロシージャを実行してください。"
        Exit Sub
    End If

    Dim driver As New WebDriver
    driver.Edge EDGE_DRIVER_PASS  ' ChromeDriverのパスを適切に設定してください

    ' 別のページに遷移
    driver.Navigate "https://www.example.org", mSessionId
    
    MsgBox "2つ目のページに遷移しました。OKをクリックするとブラウザを閉じます。"

    ' ブラウザを閉じる
    driver.Execute driver.CMD_QUIT, CreateParams("sessionId", mSessionId)

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









Private Sub rBoxLogin(ByRef driver As WebDriver, ByVal mail As String, ByVal pass As String)
    driver.Navigate "https://rak.account.box.com/login"
    ' Log.Info driver.Execute(driver.CMD_GET_TITLE)  ' "Box | ログイン"
    driver.FindElement(by.XPath, "/html/body/div/div[2]/div/div[1]/div/div[1]/form/button").Click

    ' Log.Info driver.Execute(driver.CMD_GET_TITLE) ' "Rakuten Global - サインイン"
    Dim mailInput As WebElement
    Set mailInput = WaitForElement(driver, by.ID, "input43", 10)
    mailInput.SetValue mail
    driver.FindElement(by.XPath, "//*[@id=""form35""]/div[2]/input").Click
    
    ' Log.Info driver.Execute(driver.CMD_GET_TITLE) ' "Rakuten Global - サインイン"
    Dim passInput As WebElement
    Set passInput = WaitForElement(by.Name, "credentials.passcode")
    passInput.SetValue pass
    driver.FindElement(by.XPath, "//*[@id=""form21""]/div[2]/input").Click
End Sub

Public Function WaitForElement(driver As WebDriver, by As by, value As String, Optional timeoutSeconds As Long = 10) As WebElement
    Dim startTime As Date
    startTime = Now
    
    Do While Now < startTime + TimeSerial(0, 0, timeoutSeconds)
        On Error Resume Next
        Set WaitForElement = driver.FindElement(by, value)
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


Sub SearchTest()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("TEST1").ListObjects("テーブル3")
    
    Dim driver As New WebDriver
    SafeOpen driver, Edge
    
    driver.Navigate "https://trackings.post.japanpost.jp/services/srv/sequenceNoSearch/input"
    
    driver.FindElement(by.Name, "requestNo").SetValue "49428413064"
    driver.FindElement(by.Name, "count").SetValue "10"
    driver.FindElement(by.Name, "sequenceNoSearch").Click

    Dim r As Collection
    Set r = GetComplexTableData2(driver, "//*[@id=""content""]/form/div/table/tbody")
    
    Dim item As Dictionary
    Dim newRow As ListRow
    
    For Each item In r
        Set newRow = tbl.ListRows.Add
        
        With newRow
            If item("type") = "normal" Then
                .Range(1) = item("inquiryNumber")
                .Range(2) = item("itemType")
                .Range(3) = item("latestDate")
                .Range(4) = item("latestStatus")
                .Range(5) = item("latestOffice")
                .Range(6) = item("postalCode")
                .Range(7) = item("prefecture")
            Else
                .Range(1) = item("inquiryNumber")
                .Range(2) = "エラー"
                .Range(3) = item("errorType")
                .Range(4) = item("errorMessage")
                .Range(5) = "N/A"
                .Range(6) = "N/A"
                .Range(7) = "N/A"
            End If
        End With
    Next item
    
    driver.Shutdown
    
    MsgBox "検索が完了しました。", vbInformation
End Sub

Function GetComplexTableData(driver As WebDriver, tableLocator As String) As Collection
    Dim tableElement As WebElement
    Set tableElement = driver.FindElement(by.XPath, tableLocator)

    Dim rows() As WebElement
    rows = tableElement.FindElements(by.TagName, "tr")

    Dim data As New Collection
    Dim i As Long
    Dim skipNextRow As Boolean

    For i = 2 To UBound(rows) ' ヘッダー行をスキップ
        If skipNextRow Then
            skipNextRow = False
            GoTo ContinueLoop
        End If

        Dim firstCell As WebElement
        Set firstCell = rows(i).FindElements(by.TagName, "td")(0)

        Dim item As New Dictionary

        If firstCell.GetText Like "*-*-*-*" Then ' 問い合わせ番号のパターンをチェック
            ' 通常のデータ行
            item("type") = "normal"
            item("inquiryNumber") = firstCell.GetText
            item("itemType") = rows(i).FindElements(by.TagName, "td")(1).GetText
            item("latestDate") = rows(i).FindElements(by.TagName, "td")(2).GetText
            item("latestStatus") = rows(i).FindElements(by.TagName, "td")(3).GetText
            item("latestOffice") = rows(i).FindElements(by.TagName, "td")(4).GetText
            item("prefecture") = rows(i).FindElements(by.TagName, "td")(5).GetText
            item("postalCode") = rows(i + 1).FindElements(by.TagName, "td")(0).GetText

            skipNextRow = True
        Else
            ' エラーメッセージ行
            item("type") = "error"
            item("inquiryNumber") = firstCell.GetText
            item("errorMessage") = rows(i).FindElements(by.TagName, "td")(1).GetText
            item("itemType") = "N/A"
            item("latestDate") = "N/A"
            item("latestStatus") = "N/A"
            item("latestOffice") = "N/A"
            item("prefecture") = "N/A"
            item("postalCode") = "N/A"
        End If

        data.Add item

ContinueLoop:
    Next i

    Set GetComplexTableData = data
End Function


Function GetComplexTableData2(driver As WebDriver, tableLocator As String) As Collection
    Dim tableElement As WebElement
    Set tableElement = driver.FindElement(by.XPath, tableLocator)
    
    Dim rows() As WebElement
    rows = tableElement.FindElements(by.TagName, "tr")
    
    Dim data As New Collection
    Dim i As Long
    Dim skipNextRow As Boolean
    
    For i = 2 To UBound(rows) ' ヘッダー行をスキップ
        If skipNextRow Then
            skipNextRow = False
            GoTo ContinueLoop
        End If
        
        Dim cells() As WebElement
        cells = rows(i).FindElements(by.TagName, "td")
        
        Dim item As New Dictionary
        
        If UBound(cells) + 1 = 6 Then ' 通常のデータ行は6つのtdを持つ
            ' 通常のデータ行
            item("type") = "normal"
            item("inquiryNumber") = cells(0).GetText
            item("itemType") = cells(1).GetText
            item("latestDate") = cells(2).GetText
            item("latestStatus") = cells(3).GetText
            item("latestOffice") = cells(4).GetText
            item("prefecture") = cells(5).GetText
            
            ' 次の行の郵便番号を取得
            Dim nextRowCells() As WebElement
            nextRowCells = rows(i + 1).FindElements(by.TagName, "td")
            item("postalCode") = nextRowCells(0).GetText
            
            skipNextRow = True
        Else
            ' エラーメッセージ行
            item("type") = "error"
            item("inquiryNumber") = cells(0).GetText
            
            ' エラーメッセージを取得
            Dim errorCell As WebElement
            Set errorCell = cells(1)
            Dim errorMessage As String
            errorMessage = errorCell.GetText
            
            ' エラーの種類を判断
            If InStr(1, errorMessage, "見つかりません") > 0 Then
                item("errorType") = "NotFound"
            ElseIf InStr(1, errorMessage, "同じお問い合わせ番号が入力されています") > 0 Then
                item("errorType") = "Duplicate"
            Else
                item("errorType") = "Unknown"
            End If
            
            item("errorMessage") = errorMessage
            item("itemType") = "N/A"
            item("latestDate") = "N/A"
            item("latestStatus") = "N/A"
            item("latestOffice") = "N/A"
            item("prefecture") = "N/A"
            item("postalCode") = "N/A"
        End If
        
        data.Add item
        
ContinueLoop:
    Next i
    
    Set GetComplexTableData2 = data
End Function



Private Function params(ParamArray keysAndValues()) As Dictionary
    Dim dict As New Dictionary
    Dim i As Integer
    For i = 0 To UBound(keysAndValues) - 1 Step 2
        dict.Add keysAndValues(i), keysAndValues(i + 1)
    Next i
    Set params = dict
End Function
