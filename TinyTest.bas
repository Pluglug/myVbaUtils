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




' Option Explicit

' Private Sub UserForm_Initialize()
' With ProgressBar1
'     .Min = 0
'     .Max = 100
'     .Value = 0
' End With
' txtMessage.Caption = ""
' End Sub

' Private Sub UserForm_Terminate()
'     Unload Me
' End Sub

Option Explicit

'd8888b.  .d8b.  d888888b  .d8b.       d88888b d8888b. d888888b d888888b
'88  `8D d8' `8b `~~88~~' d8' `8b      88'     88  `8D   `88'   `~~88~~'
'88   88 88ooo88    88    88ooo88      88ooooo 88   88    88       88
'88   88 88~~~88    88    88~~~88      88~~~~~ 88   88    88       88
'88  .8D 88   88    88    88   88      88.     88  .8D   .88.      88
'Y8888D' YP   YP    YP    YP   YP      Y88888P Y8888D' Y888888P    YP

Public Sub EditingTableData(Table As ListObject)

    Const maxOrder As Long = 4
    
    ' 例外的な入力機能のON/OFF
    Const EnableExceptionalInput As Boolean = True
   
    ' 各チームの店舗の順序と行番号を格納する
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 各チームの最初の店舗情報を格納する
    Dim firstStoreDict As Object
    Set firstStoreDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, lastRow As Long
    lastRow = Table.ListRows.Count
    
    Dim team As String, order As Long, storeName As String, storeAddress As String
    Dim firstStoreName As String, firstStoreAddress As String
    Dim nextOrder As Long, nextStoreName As String, nextStoreAddress As String
    
    For i = 1 To lastRow
        ' 各行について、チーム名、順番、店舗名、店舗住所を読み取る
        team = Table.ListColumns("チーム名").DataBodyRange(i).Value
        order = Val(Table.ListColumns("順番").DataBodyRange(i).Value)
        
        storeName = Table.ListColumns("店舗名").DataBodyRange(i).Value
        storeAddress = Table.ListColumns("店舗住所").DataBodyRange(i).Value
        
        ' チームの最初の店を記憶
        If EnableExceptionalInput And Not firstStoreDict.Exists(team) Then
            firstStoreDict.Add team, Array(storeName, storeAddress)
        End If
        
        If dict.Exists(team & "_" & order) Then
            ' キーに重複、つまり同じチームに同じ順番がある場合、該当セルの背景色を赤にする
            Table.ListColumns("順番").DataBodyRange(dict(team & "_" & order)).Interior.Color = RGB(255, 0, 0)
        Else
            ' 辞書オブジェクトdictに、キーとして"チーム名_順序"を使用し
            ' それに対応する値として現在の行番号iを格納する
            ' これにより、特定のチームの特定の順番の店舗がどの行にあるかが追跡できる
            dict.Add team & "_" & order, i
        End If
    Next i
    
    ' 次の店舗名と次の店舗住所の設定
    For i = 1 To Table.ListRows.Count
        team = Table.ListColumns("チーム名").DataBodyRange(i).Value
        order = Table.ListColumns("順番").DataBodyRange(i).Value
        
        ' 次の順番を探す
        nextOrder = order
        Do
            nextOrder = nextOrder + 1
            If nextOrder > maxOrder Then Exit Do
        Loop Until dict.Exists(team & "_" & nextOrder) Or nextOrder > maxOrder
        
        ' `dict`に次の順番の店舗情報が存在すれば、その情報をセット
        If dict.Exists(team & "_" & nextOrder) Then
            nextStoreName = Table.ListColumns("店舗名").DataBodyRange(dict(team & "_" & nextOrder)).Value
            nextStoreAddress = Table.ListColumns("店舗住所").DataBodyRange(dict(team & "_" & nextOrder)).Value
            Table.ListColumns("次の店舗名").DataBodyRange(i).Value = nextStoreName
            Table.ListColumns("次の店舗住所").DataBodyRange(i).Value = nextStoreAddress
        Else
        ' 存在しなければ最後の店舗とみなす
            If EnableExceptionalInput And firstStoreDict.Exists(team) Then
                ' 例外的な入力が有効かつ最初の店舗情報が存在する場合
                ' 該当セルに最初の店舗情報をセットし、背景色を変更する
                firstStoreName = firstStoreDict(team)(0)
                firstStoreAddress = firstStoreDict(team)(1)
                Table.ListColumns("次の店舗名").DataBodyRange(i).Value = firstStoreName
                Table.ListColumns("次の店舗住所").DataBodyRange(i).Value = firstStoreAddress
                Table.ListColumns("次の店舗名").DataBodyRange(i).Interior.Color = RGB(226, 239, 218)
                Table.ListColumns("次の店舗住所").DataBodyRange(i).Interior.Color = RGB(226, 239, 218)
            Else
                Table.ListColumns("次の店舗名").DataBodyRange(i).Value = "なし"
                Table.ListColumns("次の店舗住所").DataBodyRange(i).Value = "なし"
            End If
        End If
    Next i

End Sub


'     .o88b.  .d88b.  db      db      d88888b  .o88b. d888888b      d888888b d8b   db d88888b  .d88b.
'    d8P  Y8 .8P  Y8. 88      88      88'     d8P  Y8 `~~88~~'        `88'   888o  88 88'     .8P  Y8.
'    8P      88    88 88      88      88ooooo 8P         88            88    88V8o 88 88ooo   88    88
'    8b      88    88 88      88      88~~~~~ 8b         88            88    88 V8o88 88~~~   88    88
'    Y8b  d8 `8b  d8' 88booo. 88booo. 88.     Y8b  d8    88           .88.   88  V888 88      `8b  d8'
'     `Y88P'  `Y88P'  Y88888P Y88888P Y88888P  `Y88P'    YP         Y888888P VP   V8P YP       `Y88P'

' # 事前準備
'
'SeleniumBasicをダウンロードしインストールしてください。
'http://florentbr.github.io/SeleniumBasic/
'当スクリプト作成時点では下記バージョンを使用しています。
'Version:2.0.9.0

'Microsoft Edgeを操作するためのWebDriverをダウンロードしてください。
'EdgeのバージョンとWebDriverのバージョンは同じにしないと動作しません。
'https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'当スクリプト作成時点では下記バージョンを使用しています。
'Version:109.0.1518.78:x64

'解凍したら"msedgedriver.exe"を"edgedriver.exe"とリネームします。
'同ファイルをSeleniumBasic同梱の"edgedriver.exe"と入れ替えます。
'C:\Users%username%AppData\Local\SeleniumBasic

'VBE>ツール>参照設定からMicrosoft Scripting Runtimeを有効にしてください。
'VBE>ツール>参照設定からSelenium Type Libraryを有効にしてください。

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'参考：https://excel-ubara.com/excelvba4/EXCEL_VBA_401.html

Public Sub CollectInfoWithGoogleMaps(Table As ListObject, startRow As Long, endRow As Long, headless As Boolean)
    
    formProgress.Caption = "Google Mapから情報を取得"
    formProgress.Show vbModeless
    progress_Repaint 0, "接続準備をしています"
    
    Dim i As Long, lastRow As Long
   
    lastRow = Table.ListRows.Count

    Dim driver As WebDriver
    Set driver = New EdgeDriver
    Dim Keys As New Keys
    
    If headless Then
        driver.SetCapability "ms:edgeOptions", _
            "{" & _
                """args"":[""headless""]" & _
            "}"
    End If
    
    On Error GoTo DriverStartErrorHandler
    progress_Repaint 0, "接続を開始します"
    driver.Start
    On Error GoTo 0

ContinueExecution:

    Dim url As String
    url = "https://www.google.co.jp/maps/dir///@33.4189985,131.5800935,13z/data=!4m6!4m5!2m3!1b1!2b1!3b1!3e0?hl=ja&entry=ttu"
    
    Dim fromAddress As String, toAddress As String
    Dim travelTime As String
    Dim distance As String
    
    Dim totalRows As Long
    totalRows = WorksheetFunction.Min(endRow, lastRow) - startRow + 1
    
    For i = startRow To WorksheetFunction.Min(endRow, lastRow)
        ' 現在の進捗度をパーセントで計算
        Dim progress As Long
        progress = ((i - startRow) / totalRows) * 100
        
        ' パーセントとメッセージを更新
        progress_Repaint progress, "処理中です: " & i - startRow + 1 & " / " & totalRows
        
        fromAddress = Table.ListColumns("店舗住所").DataBodyRange(i).Value
        toAddress = Table.ListColumns("次の店舗住所").DataBodyRange(i).Value
        
        ' 初期化
        travelTime = ""
        distance = ""
        
        If toAddress = "なし" Then
            Table.ListColumns("所要時間").DataBodyRange(i).Value = "なし"
        Else
            'ページの読み込みに失敗した場合、2回まで再試行する
            If Not TryGetURL(driver, url, 2) Then
                Unload formProgress
                'MsgBox "ページの読み込みに失敗しました。", vbCritical, "Error"
                Exit Sub
            End If
            
            driver.FindElementByXPath("//*[@id=""sb_ifc50""]/input").SendKeys fromAddress
            driver.FindElementByXPath("//*[@id=""sb_ifc51""]/input").SendKeys toAddress
            driver.FindElementByXPath("//*[@id=""sb_ifc51""]/input").SendKeys Keys.Enter
            
            ' ページが完全に読み込まれるまで待機
            Application.Wait (Now + timeValue("0:00:05"))
            
            On Error Resume Next
            travelTime = driver.FindElementByXPath("//*[@id=""section-directions-trip-0""]/div[1]/div/div[1]/div[1]").text
            distance = driver.FindElementByXPath("//*[@id=""section-directions-trip-0""]/div[1]/div/div[1]/div[2]/div").text
            On Error GoTo 0
            
            If travelTime = "" Then
                'NOTE: ブラウザの応答速度の影響で、travelTimeが空になる可能性もあります
                Table.ListColumns("所要時間").DataBodyRange(i).Value = "Error"
            Else
                Table.ListColumns("所要時間").DataBodyRange(i).Value = travelTime
                Table.ListColumns("移動距離").DataBodyRange(i).Value = distance
            End If
        End If
    Next i
    
    progress_Repaint 100, "終了しています"
    driver.Quit
    Unload formProgress
    'MsgBox "おわった", vbInformation, "Google Mapから情報を取得"
    
Exit Sub
DriverStartErrorHandler:
    If Err.Number = 33 Or Err.Number = 0 Then ' 21 TimeOut発生
        progress_Repaint 0, "ドライバを更新中です"
        SafeOpen driver, Edge
        progress_Repaint 0, "接続を開始します"
        Resume ContinueExecution
    Else
        Unload formProgress
        MsgBox "予期せぬErrorが発生しました。Error Number: " & Err.Number & ", Description: " & Err.Description, vbCritical, "DriverStart Error"
        Exit Sub
    End If
End Sub

Private Function TryGetURL(driver As WebDriver, url As String, maxRetries As Long) As Boolean
    Dim retries As Integer
    On Error GoTo ErrorHandler
    
    For retries = 1 To maxRetries
        driver.Get url
        TryGetURL = True
        Exit Function
ErrorHandler:
        If Err.Number = 21 Then ' TimeoutErrorの場合
            If retries < maxRetries Then
                driver.Navigate.Refresh ' ページをリフレッシュ
            Else
                MsgBox "ページの読み込みがタイムアウトしました。", vbCritical, "TryGetURL Error"
                TryGetURL = False
            End If
        Else ' それ以外のエラーの場合
            MsgBox "予期せぬErrorが発生しました。Error Number: " & Err.Number & ", Description: " & Err.Description, vbCritical, "TryGetURL Error"
            TryGetURL = False
        End If
    Next retries
End Function

Private Sub progress_Repaint(progress As Long, Message As String)
    With formProgress
        .ProgressBar1.Value = progress
        .txtMessage.Caption = Message
        .Repaint
    End With
End Sub


'     .o88b.  .d88b.  d8b   db db    db d88888b d8888b. d888888b
'    d8P  Y8 .8P  Y8. 888o  88 88    88 88'     88  `8D `~~88~~'
'    8P      88    88 88V8o 88 Y8    8P 88ooooo 88oobY'    88
'    8b      88    88 88 V8o88 `8b  d8' 88~~~~~ 88`8b      88
'    Y8b  d8 `8b  d8' 88  V888  `8bd8'  88.     88 `88.    88
'     `Y88P'  `Y88P'  VP   V8P    YP    Y88888P 88   YD    YP

Public Sub ConvertDataFormats(tbl As ListObject)
    
    ' 時間の変換
    Call ConvertTimeColumnFormat(tbl.ListColumns("所要時間"))
    
    ' 移動距離の変換
    Call ConvertDistanceColumnFormat(tbl.ListColumns("移動距離"))
    
End Sub

Private Sub ConvertTimeColumnFormat(timeColumn As ListColumn)
    Dim i As Long
    Dim timeValue As Variant
    Dim hours As Long
    Dim minutes As Long
    Dim customTimeString As String
    
    For i = 1 To timeColumn.DataBodyRange.Rows.Count
        timeValue = timeColumn.DataBodyRange.Cells(i, 1).Value
        
        ' セルが空、または"なし"、"Error"、または既に時間形式の場合はスキップ
        If IsEmpty(timeValue) Or timeValue = "なし" Or timeValue = "Error" Or IsNumeric(timeValue) Then
            GoTo Continue
        End If
        
        ' 時間と分を抽出して変換
        hours = 0
        minutes = 0
        If InStr(1, timeValue, "時間") > 0 Then
            hours = Split(timeValue, "時間")(0)
            timeValue = Split(timeValue, "時間")(1)
        End If
        If InStr(1, timeValue, "分") > 0 Then
            minutes = Split(timeValue, "分")(0)
        End If
        
        ' 時間形式に変換
        timeValue = TimeSerial(hours, minutes, 0)
       
        ' セルに時間形式で書き込み
        timeColumn.DataBodyRange.Cells(i, 1).Value = timeValue
        timeColumn.DataBodyRange.Cells(i, 1).NumberFormat = "[h]:mm"
Continue:
    Next i
End Sub

Private Sub ConvertDistanceColumnFormat(distanceColumn As ListColumn)
    Dim i As Long
    Dim distanceString As String
    Dim distanceValue As Double ' LongからDoubleに変更
    
    For i = 1 To distanceColumn.DataBodyRange.Rows.Count
        distanceString = distanceColumn.DataBodyRange.Cells(i, 1).Value
        
        ' セルが空、または"なし"、"Error"、または既に数値形式の場合はスキップ
        If distanceString = "" Or distanceString = "なし" Or distanceString = "Error" Or IsNumeric(distanceString) Then
            GoTo Continue
        End If
        
        ' キロメートルとメートルの両方を考慮して数値に変換
        If InStr(distanceString, " km") > 0 Then
            distanceValue = CDbl(Replace(distanceString, " km", ""))
        ElseIf InStr(distanceString, " m") > 0 Then
            distanceValue = CDbl(Replace(distanceString, " m", "")) / 1000
        Else
            GoTo Continue ' 未知の単位の場合はスキップ
        End If
        
        ' セルに数値を書き込み、形式を設定
        distanceColumn.DataBodyRange.Cells(i, 1).Value = distanceValue
        distanceColumn.DataBodyRange.Cells(i, 1).NumberFormat = "0.0 ""km"""
Continue:
    Next i
End Sub



'     .d8b.  d8888b. d8888b.      d888888b  .d88b.       .88b  d88.  .d8b.  .d8888. d888888b d88888b d8888b.
'    d8' `8b 88  `8D 88  `8D      `~~88~~' .8P  Y8.      88'YbdP`88 d8' `8b 88'  YP `~~88~~' 88'     88  `8D
'    88ooo88 88   88 88   88         88    88    88      88  88  88 88ooo88 `8bo.      88    88ooooo 88oobY'
'    88~~~88 88   88 88   88         88    88    88      88  88  88 88~~~88   `Y8b.    88    88~~~~~ 88`8b
'    88   88 88  .8D 88  .8D         88    `8b  d8'      88  88  88 88   88 db   8D    88    88.     88 `88.
'    YP   YP Y8888D' Y8888D'         YP     `Y88P'       YP  YP  YP YP   YP `8888Y'    YP    Y88888P 88   YD

Public Sub AddToMaster(sourceSheet As Worksheet)
    ' txtAlreadyAddedを確認し、これが空欄でない場合は、警告し、処理を中止
    Dim txtBox As Object
    Set txtBox = sourceSheet.OLEObjects("txtAlreadyAdded").Object
    If txtBox.Value <> "" Then
        MsgBox "既に追加済みです", vbExclamation, "警告"
        Exit Sub
    End If
    
    ' sourceTableからtargetTableへ一致する列のデータをコピー
    CopyMatchingColumns sourceSheet.ListObjects(1), shMasterData1.ListObjects(1), sheet.Name
    
    ' 処理が完了したら、txtAlreadyAddedに現在の日時"mm/dd hh:nn"を追加
    txtBox.Value = Format(Now, "mm/dd hh:nn")
End Sub

' 一致する見出し名をすべてコピー
Private Sub CopyMatchingColumns(sourceTable As ListObject, targetTable As ListObject, sourceName As String)
    Dim col As ListColumn
    Dim targetLastRow As Long
    
    ' targetTableにデータがない場合、新しい行を追加
    If targetTable.ListRows.Count = 0 Then
        targetTable.ListRows.Add
        targetLastRow = 0
    Else
        ' targetTableの最終行を取得
        targetLastRow = targetTable.ListRows.Count
    End If

    For Each col In sourceTable.ListColumns
        If ColumnExists(targetTable, col.Name) Then
            ' 新しい行を追加し、データをコピー
            Dim sourceRange As Range, targetRange As Range
            Set sourceRange = col.DataBodyRange
            Set targetRange = targetTable.ListColumns(col.Name).DataBodyRange.Resize(targetLastRow + sourceRange.Rows.Count).Offset(targetLastRow)
            sourceRange.Copy Destination:=targetRange
        End If
    Next col
    
    ' "追加元"列に履歴を追加
    If ColumnExists(targetTable, "追加元") Then
        Dim historyRange As Range
        Set historyRange = targetTable.ListColumns("追加元").DataBodyRange.Resize(targetLastRow + sourceTable.ListRows.Count).Offset(targetLastRow)
        historyRange.Value = sourceName
    End If
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
