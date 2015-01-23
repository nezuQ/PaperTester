Option Explicit

'===== ライセンス =====
'PaperTester
'Copyright (c) 2014 nezuq
'This software is released under the MIT License.
'https://github.com/nezuQ/PaperTester/blob/master/LICENSE.txt


Class PaperTester_EXE
  
  Private exe, wsh, shl, excel
  Private m_adr, m_fullScreen
  
  'プロセスID
  Public ProcessID
  
  'プロセス名（実行ファイル名）
  Public ProcessName
  
  'Windowのコントローラ
  Public Hwnd

  'オブジェクト作成イベント
  Private Sub Class_Initialize()
    Set wsh = WScript.CreateObject("WScript.Shell")
    Set excel = WScript.CreateObject("Excel.Application")
    ProcessID = ""
    ProcessName = ""
    Hwnd = 0
    m_adr = ""
  End Sub

  '実行ファイルのパス
  Public Property Get Path
    Path = m_adr
  End Property
  
  'フルスクリーン状態のON/OFF
  Public Property Let FullScreen(doFullScreen)
    If (m_fullScreen = doFullScreen) Then
      If (doFullScreen) Then
        wsh.SendKeys "% X"
      Else
        wsh.SendKeys "% R"
      End If
    End If
    m_fullScreen = doFullScreen
  End Property
  
  Public Property Get FullScreen
    FullScreen = m_fullScreen 
  End Property
  
  'HWNDからプロセスIDを取得する
  Private Function GetPIDByHWND(hwnd)
    GetPIDByHWND = excel.ExecuteExcel4Macro("CALL(""user32"", ""GetWindowThreadProcessId"", ""2JN"", " & CStr(hwnd) & ", 0)")
  End Function
  
  'HWNDをプロセスIDで取得する
  Private Function GetHWNDByPID(pId)
    Dim hwnd, pIdLast
    hwnd = excel.ExecuteExcel4Macro("CALL(""user32"", ""GetDesktopWindow"", ""J"")")
    hwnd = excel.ExecuteExcel4Macro("CALL(""user32"", ""GetWindow"", ""JJJ"", " & CStr(hwnd) & ", 5)")
    GetHWNDByPID = 0
    Do While (0 <> hwnd)
      pIdLast = GetPIDByHWND(hwnd)
      If (pId = pIdLast) Then
        GetHWNDByPID = hwnd
        Exit Do
      End If
      hwnd = excel.ExecuteExcel4Macro("CALL(""user32"", ""GetWindow"", ""JJJ"", " & CStr(hwnd) & ", 2)")
    Loop
  End Function

  '起動する
  Public Sub Run(adr)
    m_adr = adr
    Set exe = wsh.Exec(adr)
    m_fullScreen = False
    Wscript.Sleep 1000
    Dim cmd
    cmd = Split(Trim(adr) & " ", " ")(0)
    ProcessID = exe.ProcessID
    ProcessName = Trim(Right(cmd, Len(cmd) - InStrRev(cmd, "\")))
    Hwnd = GetHWNDByPID(ProcessID)
  End Sub
  
  '戻る（BackSpace）
  Public Sub GoBack()
    wsh.SendKeys("{BS}")
  End Sub

  '終了する
  Public Sub Quit()
    exe.Terminate
    Set exe = Nothing
  End Sub
  
  '終了処理
  Public Sub Terminate()
    If (exe is Nothing) Then 
      ' 処理なし
    Else
      Quit
    End If
    
    Set wsh = Nothing
    Set shl = Nothing
    Set excel = Nothing
  End Sub
  
  'オブジェクト作成イベント
  Private Sub Class_Terminate()
    Terminate
  End Sub
  
End Class


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

  '検証結果を貼り付けた後の行間隔
  Public AfterValidationLogRows

  'データベースの値を貼り付けるEXCELシート
  Public DatabaseSheetName

  'データベースの値を貼り付ける開始セル
  Public DataPrintCellAddress

  'データベースの値を貼り付ける行間隔
  Public DataIntervalRows

  '接続文字列
  Public ConnectionString

  '===== 固定値 =====
  
  '既定で起動するEXEの名前（IE）
  Public DefaultExeName
  
  '引数行区切りのキーワード
  Public OptionRowSeperateKey

  '入力値指定のキーワード
  Public SpecifyInputKey

  '属性指定のキーワード
  Public SpecifyAttributeKey

  'インデックス指定のキーワード
  Public SpecifyIndexKey

  'テキスト包括のキーワード
  Public TextWrapKey

  '画面のアクティベーション処理の最大待機秒
  Public WindowActivationMaxWaitSeconds

  'ページ遷移失敗時のリフレッシュ処理間隔
  Public RefreshIntervalSeconds
  
  'EXEクラスの型名
  Public ExeTypeName

  '===== 前処理 =====
  
  Private i, j, k
  Private wsh, shl, fs
  Private excel, wbk, shtSS, shtDB, rng
  Private con
  Private exes(), idxExes(), nmeExes(), exe
  Private doc
  Private wLoc, wSvc, wEnu, wIns
  Private idxPasteArea, idxSetArea
  
  'オブジェクト作成イベント
  Private Sub Class_Initialize
    Set wsh = WScript.CreateObject("WScript.Shell")
    Set shl = CreateObject("Shell.Application")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set excel = WScript.CreateObject("Excel.Application")
    excel.Application.DisplayAlerts = False

    Redim exes(0)
    Set exes(0) = Nothing
    Redim idxExes(0)
    Redim nmeExes(0)
    idxExes(0) = 0
    nmeExes(0) = ""
    Set exe = Nothing
    Set doc = Nothing

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
    AfterValidationLogRows = 1
    VerticalScrollRate = 1.00
    DatabaseSheetName = "Sheet2"
    DataPrintCellAddress = "A1"
    DataIntervalRows = 2
    ConnectionString = ""
    
    '固有値にデフォルト値を入力する
    DefaultExeName = "iexplore.exe"
    OptionRowSeperateKey = " %|% "
    SpecifyInputKey = "<-"
    SpecifyAttributeKey = "="
    SpecifyIndexKey = "#"
    TextWrapKey = "'"
    WindowActivationMaxWaitSeconds = 3
    RefreshIntervalSeconds = 30
    ExeTypeName = "PaperTester_EXE"
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
  
  '数値を切り上げする
  Private Function Ceil(Number)
    Ceil = Int(Number)
    if Ceil <> Number then
      Ceil = Ceil + 1
    end if
  end function
  
  'Windowを最前面に移動する
  Private Sub BringWindowToTop(hwnd)
    excel.ExecuteExcel4Macro("CALL(""user32"",""SetWindowPos"",""JJJJJJJJ""," & hwnd & ",-1,0,0,0,0,3)")
    excel.ExecuteExcel4Macro("CALL(""user32"",""SetWindowPos"",""JJJJJJJJ""," & hwnd & ",-2,0,0,0,0,3)")
  End Sub  
  
  'Windowを最後面に移動する
  Private Sub BringWindowToBottom(hwnd)
    excel.ExecuteExcel4Macro("CALL(""user32"",""SetWindowPos"",""JJJJJJJJ""," & hwnd & ",1,0,0,0,0,3)")
  End Sub  
  
  'アクティブWindowのHWNDを取得する
  Private Function GetActiveWindow()
    GetActiveWindow = excel.ExecuteExcel4Macro("CALL(""user32"",""GetActiveWindow"",""J"")")
  End Function  
  
  'EXEを配列に保存する
  Private Function ShiftExesArray(cntShift)
    Dim idxExe
    idxExe = Ubound(exes) + cntShift
    If (exes(0) is Nothing) Then
      idxExe = idxExe - 1
    End If
    If (idxExe < 0) Then
      idxExe = 0
      Set exes(0) = Nothing
      idxExes(0) = 0
      nmeExes(0) = ""
    End If
    Redim Preserve exes(idxExe)
    Redim Preserve nmeExes(idxExe)
    Redim Preserve idxExes(idxExe)
    ShiftExesArray = idxExe
  End Function
  
  '特殊キーを入力する
  Private Sub KeybdEvent(bVk, bScan, dwFlags, dwExtraInfo)
    Call excel.ExecuteExcel4Macro("CALL(""user32"",""keybd_event"",""JJJJJ"", " & bVk & ", " & bScan & ", " & dwFlags & ", " & dwExtraInfo & ")")
  End Sub

  '文字列をクリップボードに記録する
  Private Sub CopyText(str)
    Dim cmd
    cmd = "cmd /c ""echo " & str & "| clip"""
    wsh.Run cmd, 0
  End Sub

  'IEの遷移を待つ
  Private Sub IEWait(exe)
    Dim hmsLimit
    hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
    Do While (exe.Busy = True Or exe.readyState <> 4)
      Wscript.Sleep 100
      If (hmsLimit < Now) Then
        exe.Refresh
        hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
      End If
    Loop
    hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
    Do Until (exe.document.ReadyState = "complete")
      Wscript.Sleep 100
      If (hmsLimit < Now) Then
        exe.Refresh
        hmsLimit = Now + TimeSerial(0, 0, RefreshIntervalSeconds)
      End If
    Loop
    Set doc = exe.document
  End Sub

  '指定ウィンドウをアクティブにする
  Private Function ActivateWindow(processId)
    Dim cnt, maxCnt
    maxCnt = WindowActivationMaxWaitSeconds * 10
    cnt = 1
    Do While not wsh.AppActivate(processId)
      If (maxCnt <= cnt) Then Exit Do
      cnt = cnt + 1
      Wscript.Sleep 100
    Loop
    ActivateWindow = (cnt < maxCnt)
  End Function

  'プロセスIDをプロセス名で取得する
  Private Function GetProcIDByName(nmeExe, idxExe)
    Dim pId
    pId = -1
    Dim cntExe
    cntExe = 0
    For Each wIns in wEnu
      If (Not IsEmpty(wIns.ProcessId)) _
        And (LCase(wIns.Description) = LCase(nmeExe)) Then
        pId = wIns.ProcessId
        If (cntExe = idxExe) Then Exit For
        cntExe = cntExe + 1
      End If
    Next
    GetProcIDByName = pId
  End Function
  
  'プロセス名をインデックスで取得する
  Private Function GetProcNameByIndex(idxExe)
    GetProcNameByIndex = ""
    Dim pId
    pId = -1
    Dim cntExe
    cntExe = 0
    For Each wIns in wEnu
      If (Not IsEmpty(wIns.ProcessId)) Then
        GetProcNameByIndex = wIns.Description
        If (cntExe = idxExe) Then Exit For
        cntExe = cntExe + 1
      End If
    Next
  End Function

  'テキスト包括キーワードを削除する
  Private Function Unwrap(exp)
    Dim expUnwrap
    expUnwrap = Trim(exp)
    If (Left(expUnwrap, 1) = TextWrapKey) And (Right(expUnwrap, 1) = TextWrapKey) Then
      expUnwrap = Right(expUnwrap, Len(expUnwrap) - 1)
      expUnwrap = Left(expUnwrap, Len(expUnwrap) - 1)
    End If
    Unwrap = expUnwrap
  End Function

  '属性指定表現を評価する
  Private Function EvalAtrSpecExp(exp)
    Dim expTrim, expValueTrim
    expTrim = Trim(exp)
    Dim aryAtrExp(2)
    aryAtrExp(0) = "value"
    aryAtrExp(1) = Unwrap(exp)
    aryAtrExp(2) = 0
    Dim idxSAKey, idxSIKey, idxTWKey, idxLastTWKey
    idxSAKey = InStr(expTrim, SpecifyAttributeKey)
    idxTWKey = InStr(expTrim, TextWrapKey)
    If ((0 < idxSAKey) And ((idxTWKey = 0) Or (idxSAKey < idxTWKey))) Then
      aryAtrExp(0) = Trim(Left(expTrim, idxSAKey - 1))
      expValueTrim = Trim(Right(expTrim, Len(expTrim) - idxSAKey - (Len(SpecifyAttributeKey) - 1)))
      idxSIKey = InStrRev(expValueTrim, SpecifyIndexKey)
      idxLastTWKey = InStrRev(expValueTrim, TextWrapKey)
      If ((0 < idxSIKey) And (idxLastTWKey < idxSIKey)) Then
        aryAtrExp(1) = Unwrap(Left(expValueTrim, idxSIKey - 1))
        aryAtrExp(2) = Trim(Right(expValueTrim, Len(expValueTrim) - idxSIKey - (Len(SpecifyIndexKey) - 1)))
      Else
        aryAtrExp(1) = Unwrap(expValueTrim)
      End If
    End If
    EvalAtrSpecExp = aryAtrExp
  End Function
  
  '入力値指定表現を評価する
  Private Function EvalInputSpecExp(exp)
    Dim aryInputExp(5), aryAtrExp
    aryInputExp(0) = ""
    aryInputExp(1) = ""
    aryInputExp(2) = ""
    aryInputExp(3) = ""
    aryInputExp(4) = ""
    aryInputExp(5) = ""
    Dim idxSIPKey
    idxSIPKey = InStr(exp, SpecifyInputKey)
    If (0 < idxSIPKey) Then
      aryAtrExp = EvalAtrSpecExp(Left(exp, idxSIPKey - 1))
      aryInputExp(0) = aryAtrExp(0)
      aryInputExp(1) = aryAtrExp(1)
      aryInputExp(2) = aryAtrExp(2)
      aryAtrExp = EvalAtrSpecExp(Right(exp, Len(exp) - idxSIPKey - (Len(SpecifyInputKey) - 1)))
      aryInputExp(3) = aryAtrExp(0)
      aryInputExp(4) = aryAtrExp(1)
      aryInputExp(5) = aryAtrExp(2)
    Else
      aryAtrExp = EvalAtrSpecExp(exp)
      aryInputExp(0) = aryAtrExp(0)
      aryInputExp(1) = aryAtrExp(1)
      aryInputExp(2) = aryAtrExp(2)
    End If
    EvalInputSpecExp = aryInputExp
  End Function

  '要素を取得する
  Private Function GetElement(aryExp)
    Dim elm
    Set elm = Nothing
    Select Case LCase(aryExp(0))
      Case "id"
        Set elm = doc.getElementById(aryExp(1))
      Case "name"
        Set elm = doc.getElementsByName(aryExp(1))(aryExp(2))
      Case "tag"
        Set elm = doc.getElementsByTagName(aryExp(1))(aryExp(2))
      Case "class"
        Set elm = doc.getElementsByClassName(aryExp(1))(aryExp(2))
    End Select
    Set GetElement = elm
    Set elm = Nothing
  End Function

  '入力する（SendKeys/Value共通）
  Private Sub Input(exp, useSendKeys)
    Dim elm
    Dim aryExpOpts, expOpts, aryOpt
    aryExpOpts = Split(exp, OptionRowSeperateKey)
    For Each expOpts in aryExpOpts
      aryOpt = EvalInputSpecExp(expOpts)
      Set elm = GetElement(aryOpt)
      elm.Focus
      Select Case useSendKeys 
        Case 0
          elm.Value = aryOpt(4)
        Case 1
          Paste aryOpt(4)
        Case 2
          wsh.SendKeys aryOpt(4)
      End Select
    Next
    Set elm = Nothing
  End Sub
  
  '属性を取得する
  Private Function GetAttribute(elm, nmeAtr)
    Dim atr, nmeAtrLCase
    nmeAtrLCase = LCase(nmeAtr)
    Select Case nmeAtrLCase 
      Case "value"
        atr = elm.Value
      Case Else
        atr = elm.getAttribute(nmeAtr, 2)
    End Select
    GetAttribute = atr
  End Function

  'スクロールする
  Private Function Scroll(numHeight)
    Dim numNextHeight
    numNextHeight = numHeight
    If (exe.document.body.ScrollHeight < numNextHeight) Then
      numNextHeight = exe.document.body.ScrollHeight
    End If
    exe.Navigate "javascript:scroll(0, " & numNextHeight & ")"
    Wscript.Sleep 1000
    Scroll = numNextHeight
  End Function

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
    numPageHeight = exe.Height * VerticalScrollRate
    If (numPageHeight < exe.document.body.ScrollHeight) Then
      cntPage = Ceil(exe.document.body.ScrollHeight / numPageHeight)
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
  
  '検証失敗時のエラーメッセージを取得する
  Private Function GetValidationMessage(exp, valSpec, valReal)
    Dim strResult
    If (valSpec = valReal) Then
      strResult = "OK"
    Else
      strResult = "NG"
    End If
    GetValidationMessage = "【" & strResult & "】" & exp & "|" & TextWrapKey & valReal & TextWrapKey
  End Function

  '===== 操作用関数 =====
  
  'InternetExplorerを開く
  Public Sub OpenIE()
    Dim idxNextExe
    idxNextExe = ShiftExesArray(1)
    Set exe = CreateObject("InternetExplorer.Application")
    Set exes(idxNextExe) = exe
    exe.Visible = True
    WScript.Sleep 1000
    nmeExes(idxNextExe) = DefaultExeName
    idxExes(idxNextExe) = GetProcIDByName(nmeExes(0), -1)
    BringWindowToTop exe.Hwnd
    ActivateWindow idxExes(idxNextExe)
  End Sub
  
  'InternetExplorerを取得する
  Public Sub GetIE(idxExe)
    Dim idxNextExe
    idxNextExe = ShiftExesArray(1)
    Dim win
    Dim cntExe
    cntExe = 0
    For Each win In shl.Windows
      If TypeName(win.document) = "HTMLDocument" Then
        'HTMLDocument型の場合
        Set exes(idxNextExe) = win
        Set exe = exes(idxNextExe)
        If (idxExe = cntExe) Then Exit For
        cntExe = cntExe + 1
      End If
    Next
    exe.Visible = True
    nmeExes(idxNextExe) = DefaultExeName
    idxExes(idxNextExe) = GetProcIDByName(DefaultExeName, idxExe)
    BringWindowToTop exe.Hwnd
    ActivateWindow idxExes(idxNextExe)
  End Sub
  
  'EXEを起動する
  Public Sub Run(adrExe)
    Dim idxNextExe
    idxNextExe = ShiftExesArray(1)
    Set exe = new PaperTester_EXE
    exe.Run adrExe
    Set exes(idxNextExe) = exe
    nmeExes(idxNextExe) = exe.ProcessName
    idxExes(idxNextExe) = exe.ProcessID
    BringWindowToTop exe.Hwnd
    ActivateWindow exe.ProcessID
    If (exe.Hwnd <> GetActiveWindow()) Then
      wsh.SendKeys("%({TAB})")
    End If
  End Sub
  
  'InternetExplorer/EXEを閉じる
  Public Sub Quit()
    exe.Quit
    If (0 < Ubound(exes)) Then
      ActivateBeforeIE
    Else
      ShiftExesArray -1
    End If
  End Sub

  '戻る
  Public Sub GoBack()
    exe.GoBack
  End Sub

  '全画面表示を行う
  Public Sub FullScreen()
    exe.FullScreen = True
  End Sub

  '全画面表示を止める
  Sub NormalScreen()
    exe.FullScreen = False
  End Sub

  '最大化する
  Public Sub MaximumWindow()
    If (TypeName(exe) = ExeTypeName) Then
      wsh.SendKeys("% X")
    Else
      excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & exe.Hwnd & ", 3)"
    End If
  End Sub

  '最小化する
  Public Sub MinimumWindow()
    If (TypeName(exe) = ExeTypeName) Then
      wsh.SendKeys("% N")
    Else
      excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & exe.Hwnd & ", 2)"
    End If
  End Sub

  '標準表示にする
  Public Sub NormalWindow()
    If (TypeName(exe) = ExeTypeName) Then
      wsh.SendKeys("% R")
    Else
      excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & exe.Hwnd & ", 1)"
    End If
  End Sub

  '待機する
  Public Sub Sleep(sec)
    WScript.Sleep(sec * 1000)
  End Sub

  'URLで遷移する
  Public Sub Navigate(url)
    exe.Navigate url
    IEWait(exe)
  End Sub

  '次のInternetExolorerをアクティブにする
  Public Sub ActivateNextIE()
    Dim idxNextExe
    idxNextExe = ShiftExesArray(1)
    WScript.Sleep 1000
    Set exes(idxNextExe) = shl.Windows(shl.Windows.Count - 1)
    Set exe = exes(idxNextExe)
    nmeExes(idxNextExe) = GetProcNameByIndex(-1)
    idxExes(idxNextExe) = GetProcIDByName(nmeExes(idxNextExe), -1)
    IEWait(exe)
    ActivateWindow idxExes(idxNextExe)
    BringWindowToTop exe.Hwnd
  End Sub

  '前のInternetExolorerをアクティブにする
  Public Sub ActivateBeforeIE()
    Dim idxNextExe
    idxNextExe = ShiftExesArray(-1)
    Set exe = exes(idxNextExe)
    ActivateWindow idxExes(idxNextExe)
    BringWindowToTop exe.Hwnd
  End Sub

  '指定フレームをアクティブにする
  Public Sub ActivateFrame(idxFrame)
    Set doc = doc.frames(idxFrame).document
  End Sub

  '元ドキュメントをアクティブにする
  Public Sub ActivateDocument()
    Set doc = exe.document
  End Sub

  'フォーカスを当てる
  Public Sub Focus(exp)
    Dim elm
    Set elm = GetElement(EvalAtrSpecExp(exp))
    elm.Focus
    Set elm = Nothing
  End Sub

  '入力する（Value）
  Public Sub ValueInput(exp)
    Input exp, 0
  End Sub

  '入力する（Copy&Paste）
  Public Sub PasteInput(exp)
    Input exp, 1
  End Sub

  '入力する（SendKeys）
  Public Sub KeyInput(exp)
    Input exp, 2
  End Sub

  'クリックする
  Public Sub Click(exp)
    Dim elm
    Set elm = GetElement(EvalAtrSpecExp(exp))
    elm.Focus
    elm.Click
    IEWait(exe)
    Set elm = Nothing
  End Sub

  '文字列をコピー&ペーストする。
  Public Sub Paste(str)
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
        ).Offset(idxPasteArea, 0)
    rng.Value = msg
    rng.Offset(1, 1).Select
    shtSS.Paste
    Set rng = Nothing
    idxPasteArea = idxPasteArea + ScreenshotPageRows
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
  
  '画面項目を検証する（検証NG時は処理中断）
  Public Sub ValidateAttribute(exp)
    Dim msg
    msg = Record2ValidateAttribute(exp)
    If (msg <> "") Then
      Err.Raise 9999, "PaperTester", "検証NG。" & msg
    End If
  End Sub

  '画面項目を検証する（検証NG時は処理続行）
  Public Function Record2ValidateAttribute(exp)
    Const keySepMsg = ", "
    Dim aryExpOpts, expOpts, msgAll
    msgAll = ""
    aryExpOpts = Split(exp, OptionRowSeperateKey)
    For Each expOpts in aryExpOpts
      Dim aryInputExp
      aryInputExp = EvalInputSpecExp(expOpts)
      Dim elm
      Set elm = GetElement(aryInputExp)
      Dim atr
      atr = GetAttribute(elm, aryInputExp(3))
      Dim msg
      msg = GetValidationMessage(expOpts, aryInputExp(4), atr)
      Dim rng
      shtSS.Activate
      Set rng = shtSS.Range( _
        ScreenshotPrintCellAddress _
          ).Offset(idxPasteArea, 0)
      rng.Offset(0, 1).Value = msg
      Set rng = Nothing
      idxPasteArea = idxPasteArea + 1
      If (aryInputExp(4) = atr) Then
        '処理なし
      Else
        msgAll = msgAll & msg & keySepMsg
      End If
    Next
    idxPasteArea = idxPasteArea + AfterValidationLogRows
    If (msgAll = "") Then
      Record2ValidateAttribute = "" 
    Else
      Record2ValidateAttribute = Left(msgAll, Len(msgAll) - Len(keySepMsg))
    End If
  End Function

  '画面タイトルを検証する（検証NG時は処理中断）
  Public Sub ValidateTitle(title)
    Dim msg
    msg = Record2ValidateTitle(title)
    If (msg <> "") Then
      Err.Raise 9999, "PaperTester", "検証NG。" & msg
    End If
  End Sub

  '画面タイトルを検証する（検証NG時は処理続行）
  Public Function Record2ValidateTitle(title)
    Dim msg
    msg = GetValidationMessage(title, title, doc.Title)
    Dim rng
    shtSS.Activate
    Set rng = shtSS.Range( _
      ScreenshotPrintCellAddress _
        ).Offset(idxPasteArea, 0)
    rng.Offset(0, 1).Value = msg
    Set rng = Nothing
    idxPasteArea = idxPasteArea + 1
    If (msg = "") Then
      Record2ValidateTitle = "" 
    Else
      Record2ValidateTitle = msg
    End If
  End Function

  'Javascriptを実行する。
  Public Sub ExecuteJS(cmd)
    exe.Navigate "javascript:" & cmd
  End Sub

  '===== 後処理 =====
  
  '終了処理
  Public Sub Terminate
    Set wLoc = Nothing
    Set wEnu = Nothing
    Set wSvc = Nothing
    Set wIns = Nothing
    Set doc = Nothing
    If (Not(exe is Nothing)) Then
      Set exe = Nothing
    End If
    For i = LBound(exes) to UBound(exes)
      Set exes(i) = Nothing
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
