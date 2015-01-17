Option Explicit

'===== ライセンス =====
'PaperTester
'Copyright (c) 2014 nezuq
'This software is released under the MIT License.
'https://github.com/nezuQ/PaperTester/blob/master/LICENSE.txt

Class PaperTester

  '===== 設定値 =====
  
  '証跡記録用EXCELブックのパス
  Public EvidenceBookPath

  'スクリーンショットを貼り付けるEXCELシート
  Public ScreenshotSheetName

  'スクリーンショットを貼り付ける開始セル
  Public ScreenshotPrintCellAddress

  'スクリーンショットを貼り付ける行間隔
  Public ScreenshotPageRows

  'スクロール時の対画面での縦幅比
  Public VerticalScrollRate

  'データベースの値を貼り付けるEXCELシート
  Public DatabaseSheetName

  'データベースの値を貼り付ける開始セル
  Public DataPrintCellAddress

  'データベースの値を貼り付ける行間隔
  Public DataIntervalRows

  '接続文字列
  Public ConnectionString

  '===== 固定値 =====
  
  '引数行区切りのキーワード
  Public OptionRowSeperateKey

  '入力値指定のキーワード
  Public SpecifyInputValueKey

  '要素指定のキーワード
  Public SpecifyElementKey

  'インデックス指定のキーワード
  Public SpecifyIndexKey

  '画面のアクティベーション処理の最大待機秒
  Public WindowActivationMaxWaitSeconds

  'ページ遷移失敗時のリフレッシュ処理間隔
  Public RefreshIntervalSeconds

  '===== 前処理 =====
  
  Private i, j, k
  Private wsh, shl, fs
  Private excel, wbk, shtSS, shtDB, rng
  Private con
  Private ies(), idxIes(), ie
  Private doc, elm
  Private wLoc, wSvc, wEnu, wIns
  Private idxPasteArea, idxSetArea
  
  'オブジェクト作成イベント
  Private Sub Class_Initialize
    Set wsh = WScript.CreateObject("WScript.Shell")
    Set shl = CreateObject("Shell.Application")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set excel = WScript.CreateObject("Excel.Application")
    excel.Application.DisplayAlerts = False

    Redim ies(0)
    Set ies(0) = Nothing
    Redim idxIes(0)
    idxIes(0) = 0
    Set ie = Nothing
    Set doc = Nothing
    Set elm = Nothing

    Set wLoc = CreateObject("WbemScripting.SWbemLocator")
    Set wSvc = wLoc.ConnectServer
    Set wEnu = wSvc.InstancesOf("Win32_Process")

    idxPasteArea = 0
    idxSetArea = 0
    
    '設定値にデフォルト値を入力する
    EvidenceBookPath = ""
    ScreenshotSheetName = "Sheet1"
    ScreenshotPrintCellAddress = "A1"
    ScreenshotPageRows = 62
    VerticalScrollRate = 1.00
    DatabaseSheetName = "Sheet2"
    DataPrintCellAddress = "A1"
    DataIntervalRows = 2
    ConnectionString = ""
    
    '固有値にデフォルト値を入力する
    OptionRowSeperateKey = " %|% "
    SpecifyInputValueKey = "<-"
    SpecifyElementKey = "="
    SpecifyIndexKey = "#"
    WindowActivationMaxWaitSeconds = 3
    RefreshIntervalSeconds = 30
  End Sub
  
  '初期化処理
  Public Sub Initialize()
    If (EvidenceBookPath <> "") Then
      '証跡記録用EXCELブックのパスが指定されている時
      excel.Application.Visible = True
      Set wbk = excel.Application.Workbooks.Open(fs.GetAbsolutePathName(EvidenceBookPath), 2, True)
      Set shtSS = excel.Worksheets(ScreenshotSheetName)
      If (ConnectionString <> "") Then
        '接続文字列が指定されている時
        Set shtDB = excel.Worksheets(DatabaseSheetName)
        Set con = CreateObject("ADODB.Connection")
        con.Open ConnectionString
      End If
    End If
  End Sub

  '===== 共通関数 =====
  
  'IEの遷移を待つ
  Private Sub IEWait(ie)
    Dim hmsLimit
    hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
    Do While (ie.Busy = True Or ie.readyState <> 4)
      Wscript.Sleep 100
      If (hmsLimit < Now) Then
        ie.Refresh
        hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
      End If
    Loop
    hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
    Do Until (ie.document.ReadyState = "complete")
      Wscript.Sleep 100
      If (hmsLimit < Now) Then
        ie.Refresh
        hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
      End If
    Loop
    Set doc = ie.document
  End Sub

  '指定ウィンドウをアクティブにする
  Private Sub ActivateWindow(processId)
    Dim cnt, maxCnt
    maxCnt = WindowActivationMaxWaitSeconds * 10
    cnt = 1
    Do While not wsh.AppActivate(processId)
      If (maxCnt <= cnt) Then Exit Do
      cnt = cnt + 1
      Wscript.Sleep 100
    Loop
  End Sub

  'IEをアクティブにする
  Private Function ActivateIE(isFirst)
    Dim pId
    pId = -1
    For Each wIns in wEnu
      If (Not IsEmpty(wIns.ProcessId)) _
        And (wIns.Description = "iexplore.exe") Then
        pId = wIns.ProcessId
        If (isFirst) Then Exit For
      End If
    Next
    ActivateWindow pId
    ActivateIE = pId
  End Function

  '入力する（SendKeys/Value共通）
  Private Sub Input(expOptsSet, useSendKeys)
    Dim aryExpOpts, aryOpt, expOpts, expOpt, idxSep, lenSep, valInput
    aryExpOpts = Split(expOptsSet, OptionRowSeperateKey)
    For Each expOpts in aryExpOpts
      idxSep = InStr(expOpts, SpecifyInputValueKey)
      lenSep = Len(SpecifyInputValueKey)
      Set elm = GetElement(Left(expOpts, idxSep - 1))
      elm.Focus
      valInput = Trim(Right(expOpts, Len(expOpts) - idxSep - (lenSep - 1)))
      valInput = Mid(valInput, 2, Len(valInput) - 2)
      Select Case useSendKeys 
        Case 0
          elm.Value = valInput
        Case 1
          CopyAndPaste valInput
        Case 2
          wsh.SendKeys valInput
      End Select
    Next
  End Sub

  '特殊キーを入力する
  Private Sub KeybdEvent(bVk, bScan, dwFlags, dwExtraInfo)
    Call excel.ExecuteExcel4Macro(Replace(Replace(Replace(Replace("CALL(""user32"",""keybd_event"",""JJJJJ"", %0, %1, %2, %3)", "%0", bVk), "%1", bScan), "%2", dwFlags), "%3", dwExtraInfo))
  End Sub

  '文字列をクリップボードに記録する
  Private Sub CopyText(str)
    Dim cmd
    cmd = "cmd /c ""echo " & str & "| clip"""
    wsh.Run cmd, 0
  End Sub

  '要素を取得する
  Private Function GetElement(expElm)
    Dim elmTgt
    Set elmTgt = Nothing
    Dim aryExpElm, aryExpElm2
    aryExpElm = Split(expElm, SpecifyElementKey)
    Dim keyElm, valElm, idxElm
    keyElm = Trim(aryExpElm(0))
    valElm = Trim(aryExpElm(1))
    If (0 < InStr(valElm, SpecifyIndexKey)) Then
      aryExpElm2 = Split(valElm, SpecifyIndexKey)
      valElm = Trim(aryExpElm2(0))
      idxElm = Trim(aryExpElm2(1))
    End If
    Select Case LCase(keyElm)
      Case "id"
        Set elmTgt = doc.getElementById(valElm)
      Case "name"
        Set elmTgt = doc.getElementsByName(valElm)(idxElm)
      Case "tag"
        Set elmTgt = doc.getElementsByTagName(valElm)(idxElm)
      Case "class"
        Set elmTgt = doc.getElementsByClassName(valElm)(idxElm)
    End Select
    Set GetElement = elmTgt
  End Function

  'スクロールする
  Private Function Scroll(numHeight)
    Dim numNextHeight
    numNextHeight = numHeight
    If (ie.document.body.ScrollHeight < numNextHeight) Then
      numNextHeight = ie.document.body.ScrollHeight
    End If
    ie.Navigate "javascript:scroll(0, " & numNextHeight & ")"
    Wscript.Sleep 1000
    Scroll = numNextHeight
  End Function

  '数値を切り上げする
  Private Function Ceil(Number)
    Ceil = Int(Number)
    if Ceil <> Number then
      Ceil = Ceil + 1
    end if
  end function

  '繰り返しスクリーンショットを撮る
  Private Sub RepeatScreenShot(isFull, msg)
    '1枚目のスクリーンショットを撮る
    If (isFull) Then
      FullScreenShot4VisibleArea msg
    Else
      ScreenShot4VisibleArea msg
    End If
    '縦幅からスクリーンショット回数を算出する
    Dim cntPage, numHeight, numPageHeight
    numHeight = 0
    numPageHeight = ie.Height * VerticalScrollRate
    If (numPageHeight < ie.document.body.ScrollHeight) Then
      cntPage = Ceil(ie.document.body.ScrollHeight / numPageHeight)
    Else
      cntPage = 1
    End If
    '2枚目以降のスクリーンショットを撮る
    numHeight = numPageHeight
    Dim i
    For i = 2 To cntPage
      Scroll (numHeight)
      If (isFull) Then
        FullScreenShot4VisibleArea ""
      Else
        ScreenShot4VisibleArea ""
      End If
      numHeight = numHeight + numPageHeight
    Next
  End Sub

  '===== 操作用関数 =====
  
  'InternetExplorerを開く
  Public Sub Open()
    Set ies(0) = CreateObject("InternetExplorer.Application")
    Set ie = ies(0)
    ie.Visible = True
    idxIes(0) = ActivateIE(False)
  End Sub

  'InternetExplorerを取得する
  Public Sub GetIE(isFirst)
    Dim win
    For Each win In shl.Windows
      If TypeName(win.document) = "HTMLDocument" Then
        'HTMLDocument型の場合
        Set ies(0) = win
        Set ie = ies(0)
        If (isFirst) Then Exit For
      End If
    Next
    ie.Visible = True
    idxIes(0) = ActivateIE(isFirst)
  End Sub

  'InternetExplorerを閉じる
  Public Sub Close()
    ie.Quit
    If (0 < Ubound(ies)) Then
      ActivateParentWindow
    Else
      Set ie = Nothing
    End If
  End Sub

  '戻る
  Public Sub GoBack()
    ie.GoBack
  End Sub

  '全画面表示を行う
  Public Sub FullScreen()
    ie.FullScreen = True
  End Sub

  '全画面表示を止める
  Sub NormalScreen()
    ie.FullScreen = False
  End Sub

  '最大化する
  Public Sub MaximumWindow()
    excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & ie.Hwnd & ", 3)"
  End Sub

  '最小化する
  Public Sub MinimumWindow()
    excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & ie.Hwnd & ", 2)"
  End Sub

  '標準表示にする
  Public Sub NormalWindow()
    excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & ie.Hwnd & ", 1)"
  End Sub

  '待機する
  Public Sub Sleep(sec)
    WScript.Sleep(sec * 1000)
  End Sub

  'URLで遷移する
  Public Sub Navigate(url)
    ie.Navigate url
    IEWait(ie)
  End Sub

  '子画面をアクティブにする
  Public Sub ActivateChildWindow()
    Redim Preserve ies(Ubound(ies) + 1)
    Redim Preserve idxIes(Ubound(idxIes) + 1)
    WScript.Sleep 1000
    Set ies(Ubound(ies)) = shl.Windows(shl.Windows.Count - 1)
    Set ie = ies(Ubound(ies))
    idxIes(Ubound(idxIes)) = ActivateIE(False)
    IEWait(ie)
  End Sub

  '親画面をアクティブにする
  Public Sub ActivateParentWindow()
    Redim Preserve ies(Ubound(ies) - 1)
    Redim Preserve idxIes(Ubound(idxIes) - 1)
    Set ie = ies(Ubound(ies))
    ActivateWindow idxIes(Ubound(ies))
    IEWait(ie)
  End Sub

  '指定フレームをアクティブにする
  Public Sub ActivateFrame(idxFrame)
    Set doc = doc.frames(idxFrame).document
  End Sub

  '元ドキュメントをアクティブにする
  Public Sub ActivateDocument()
    Set doc = ie.document
  End Sub

  'フォーカスを当てる
  Public Sub Focus(expElm)
    Set elm = GetElement(expElm)
    elm.Focus
  End Sub

  '入力する（Value）
  Public Sub ValueInput(expOptsSet)
    Input expOptsSet, 0
  End Sub

  '入力する（Copy&Paste）
  Public Sub PasteInput(expOptsSet)
    Input expOptsSet, 1
  End Sub

  '入力する（SendKeys）
  Public Sub KeyInput(expOptsSet)
    Input expOptsSet, 2
  End Sub

  'クリックする
  Public Sub Click(expElm)
    Set elm = GetElement(expElm)
    elm.Focus
    elm.Click
    IEWait(ie)
  End Sub

  '文字列をコピー&ペーストする。
  Public Sub CopyAndPaste(str)
    CopyText str
    Wscript.Sleep 750
    wsh.SendKeys "^(v)", True
    Wscript.Sleep 750
  End Sub

  'キーを押す
  Public Sub SendKeys(key)
    wsh.SendKeys key, True
  End Sub

  'スクリーンショットを撮る（画面全体, 表示箇所のみ）
  Public Sub FullScreenShot4VisibleArea(msg)
    WScript.Sleep 1000
    Call KeybdEvent(&H2C, 0, 1, 0)
    Call KeybdEvent(&H2C, 0, 3, 0)
    WScript.Sleep 1000
    shtSS.Activate
    Set rng = shtSS.Range( _
      ScreenshotPrintCellAddress _
        ).Offset(ScreenshotPageRows * idxPasteArea, 0)
    rng.Value = msg
    rng.Offset(1, 1).Select
    shtSS.Paste
    Set rng = Nothing
    idxPasteArea = idxPasteArea + 1
  End Sub

  'スクリーンショットを撮る（画面全体）
  Public Sub FullScreenShot(msg)
    RepeatScreenShot True, msg
  End Sub

  'スクリーンショットを撮る（アクティブ画面, 表示箇所のみ）
  Public Sub ScreenShot4VisibleArea(msg)
    Call KeybdEvent(&H12, 0, 1, 0)
    FullScreenShot4VisibleArea msg
    Call KeybdEvent(&H12, 0, 3, 0)
  End Sub

  'スクリーンショットを撮る（アクティブ画面）
  Public Sub ScreenShot(msg)
    RepeatScreenShot False, msg
  End Sub

  'SQL文を発行する
  Public Sub ExecuteSQL(sql)
    Dim rs, fld
    Set rs = CreateObject("ADODB.Recordset")
    Dim cmd
    cmd = Replace(sql, OptionRowSeperateKey, vbCrLf)
    rs.Open cmd , con, 1, 1
    Dim cntClm
    cntClm = 1
    ' SQL文を記録
    Set rng = shtDB.Range(DataPrintCellAddress)
    rng.Offset(idxSetArea, 0).Value = cmd
    idxSetArea = idxSetArea + 1
    ' 列名を記録
    For each fld in rs.Fields
      rng.Offset(idxSetArea, cntClm).Value = fld.Name
      cntClm = cntClm + 1
    Next
    idxSetArea = idxSetArea + 1
    ' 値を記録
    Do Until rs.EOF
      cntClm = 1
      For each fld in rs.Fields
        rng.Offset(idxSetArea, cntClm).Value = fld.Value
        cntClm = cntClm + 1
      Next
      idxSetArea = idxSetArea + 1
      rs.MoveNext
    Loop
    idxSetArea = idxSetArea + DataIntervalRows
    rs.Close
    Set rng = Nothing
    Set fld = Nothing
    Set rs = Nothing
  End Sub

  '===== 後処理 =====
  
  '終了処理
  Public Sub Terminate
    Set wLoc = Nothing
    Set wEnu = Nothing
    Set wSvc = Nothing
    Set wIns = Nothing
    Set elm = Nothing
    Set doc = Nothing
    If (Not(ie is Nothing)) Then
      ie.FullScreen = False
      Set ie = Nothing
    End If
    For i = LBound(ies) to UBound(ies)
      Set ies(i) = Nothing
    Next
    Set rng = Nothing
    Set shtSS = Nothing
    Set excel = Nothing
    Set wsh = Nothing
    Set shl = Nothing
    Set fs = Nothing
    Set con = Nothing
  End Sub

  'オブジェクト破棄時のイベント
  Private Sub Class_Terminate
    Terminate
  End Sub
End Class
