Option Explicit

'=====  ライセンス  =====
'Copyright (c) 2014 nezuq
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php

'=====  設定  =====
'証跡記録用のEXCELブック
Dim EXCEL_EVIDENCE_BOOKPATH
EXCEL_EVIDENCE_BOOKPATH = ".\EvidenceTemplate.xlsx"

'スクリーンショットを貼り付けるEXCELシート
Dim EXCEL_SCREENSHOT_SHEETNAME
EXCEL_SCREENSHOT_SHEETNAME = "Screenshot"

'スクリーンショットを貼り付ける開始セル
Dim EXCEL_STARTPRINT_CELLADDRESS
EXCEL_STARTPRINT_CELLADDRESS = "B3"

'スクリーンショットを貼り付ける行間隔
Dim EXCEL_ONEPAGE_ROWS
EXCEL_ONEPAGE_ROWS = 62

'IE全体におけるスクリーン高の比率
Dim SCREEN_HEIGHTRATE
SCREEN_HEIGHTRATE = 0.85

'データベースの値を貼り付けるEXCELシート
Dim EXCEL_DATABASE_SHEETNAME
EXCEL_DATABASE_SHEETNAME = "Database"

'データベースの値を貼り付ける開始セル
Dim EXCEL_STARTSET_CELLADDRESS
EXCEL_STARTSET_CELLADDRESS = "B4"

'データベースの値を貼り付ける行間隔
Dim EXCEL_INTERVAL_ROWS
EXCEL_INTERVAL_ROWS = 2

'接続文字列（任意）
Dim CONNECTION_STRING
CONNECTION_STRING = ""

'=====  固定値  =====
'引数行区切りのキーワード
Dim OPTIONROW_SEPERATE_KEYWORD
OPTIONROW_SEPERATE_KEYWORD = " %|% "

'引数区切りのキーワード
Dim OPTION_SEPERATE_KEYWORD
OPTION_SEPERATE_KEYWORD = "<-"

'要素指定のキーワード
Dim ELEMENT_SPECIFY_KEYWORD
ELEMENT_SPECIFY_KEYWORD = "="

'インデックス指定のキーワード
Dim INDEX_SPECIFY_KEYWORD
INDEX_SPECIFY_KEYWORD = "#"

'===== 前処理 =====
Dim i, j, k
Dim wsh
Set wsh = WScript.CreateObject("WScript.Shell")
Dim shl
Set shl = CreateObject("Shell.Application")
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
Dim excel, wbk, shtSS, shtDB, rng
Set excel = WScript.CreateObject("Excel.Application")
excel.Application.Visible = True
excel.Application.DisplayAlerts = False
Set wbk = excel.Application.Workbooks.Open(fs.GetAbsolutePathName(EXCEL_EVIDENCE_BOOKPATH), 2, True)
Set shtSS = excel.Worksheets(EXCEL_SCREENSHOT_SHEETNAME)
Dim con
If(CONNECTION_STRING <> "") Then
  '接続文字列を指定した時
  Set shtDB = excel.Worksheets(EXCEL_DATABASE_SHEETNAME)
  Set con = CreateObject("ADODB.Connection")
  con.Open CONNECTION_STRING
End If

Dim ies(), idxIes(), ie
Redim ies(0)
Set ies(0) = Nothing
Redim idxIes(0)
idxIes(0) = 0
Set ie = Nothing
Dim doc, elm
Set doc = Nothing
Set elm = Nothing

Dim wLoc, wSvc, wEnu, wIns
Set wLoc = CreateObject("WbemScripting.SWbemLocator")
Set wSvc = wLoc.ConnectServer
Set wEnu = wSvc.InstancesOf("Win32_Process")

Dim idxPasteArea, idxSetArea
idxPasteArea = 0
idxSetArea = 0

'===== 共通関数 =====
'IEの遷移を待つ
Sub IEWait(ie)
  Do While ie.Busy = True Or ie.readyState <> 4
  Loop
  Wscript.Sleep 1000
  Set doc = ie.document
End Sub

'指定ウィンドウを活性にする
Sub ActivateWindow(processId)
  While not wsh.AppActivate(processId)
    Wscript.Sleep 100
  Wend
End Sub

'最後に起動したIEを活性にする
Function ActivateLastIE()
  Dim pId
  pId = -1
  For Each wIns in wEnu
    If Not IsEmpty(wIns.ProcessId) _
      And wIns.Description = "iexplore.exe" Then
        pId = wIns.ProcessId
    End If
  Next
  ActivateWindow pId
  ActivateLastIE = pId
End Function

'入力する（SendKeys/Value共通）
Sub Input(expOptsSet, useSendKeys)
  Dim aryExpOpts, aryOpt, expOpts, expOpt, idxSep, lenSep, valInput
  aryExpOpts = Split(expOptsSet, OPTIONROW_SEPERATE_KEYWORD)
  For Each expOpts in aryExpOpts
    idxSep = InStr(expOpts, OPTION_SEPERATE_KEYWORD)
    lenSep = Len(OPTION_SEPERATE_KEYWORD)
    Set elm = GetElement(Left(expOpts, idxSep - 1))
    elm.Focus
    valInput = Trim(Right(expOpts, Len(expOpts) - idxSep - (lenSep - 1)))
    valInput = Mid(valInput, 2, Len(valInput) - 2)
    Select Case useSendKeys 
      Case 0
        elm.Value = valInput
      Case 1
        CopyText valInput
        Wscript.Sleep 500
        wsh.SendKeys "^(v)"
      Case 2
        wsh.SendKeys valInput
    End Select
  Next
End Sub

'特殊キーを入力する
Sub KeybdEvent(bVk, bScan, dwFlags, dwExtraInfo)
  Call excel.ExecuteExcel4Macro(Replace(Replace(Replace(Replace("CALL(""user32"",""keybd_event"",""JJJJJ"", %0, %1, %2, %3)", "%0", bVk), "%1", bScan), "%2", dwFlags), "%3", dwExtraInfo))
End Sub

'文字列をクリップボードに記録する
Sub CopyText(str)
  Dim cmd
  cmd = "cmd /c ""echo " & str & "| clip"""
  wsh.Run cmd, 0
End Sub

'要素を取得する
Function GetElement(expElm)
  Dim elmTgt
  Set elmTgt = Nothing
  Dim aryExpElm, aryExpElm2
  aryExpElm = Split(expElm, ELEMENT_SPECIFY_KEYWORD)
  Dim keyElm, valElm, idxElm
  keyElm = Trim(aryExpElm(0))
  valElm = Trim(aryExpElm(1))
  If (0 < InStr(valElm, INDEX_SPECIFY_KEYWORD)) Then
    aryExpElm2 = Split(valElm, INDEX_SPECIFY_KEYWORD)
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
Function Scroll(numHeight)
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
function Ceil(Number)
  Ceil = Int(Number)
  if Ceil <> Number then
    Ceil = Ceil + 1
  end if
end function

'繰り返しスクリーンショットを撮る
Sub RepeatScreenShot(isFull, msg)
  '1枚目のスクリーンショットを撮る
  If (isFull) Then
    FullScreenShot4VisibleArea msg
  Else
    ScreenShot4VisibleArea msg
  End If
  '縦幅からスクリーンショット回数を算出する
  Dim cntPage, numHeight, numPageHeight
  numHeight = 0
  numPageHeight = ie.Height * SCREEN_HEIGHTRATE
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
Sub Open()
  Set ies(0) = CreateObject("InternetExplorer.Application")
  Set ie = ies(0)
  ie.Visible = True
  idxIes(0) = ActivateLastIE
End Sub

'InternetExplorerを閉じる
Sub Close()
  ie.Quit
  If (0 < Ubound(ies)) Then
    ActivateParentWindow
  Else
    Set ie = Nothing
  End If
End Sub

'戻る
Sub GoBack()
  ie.GoBack
End Sub

'全画面表示を行う
Sub FullScreen()
  ie.FullScreen = True
End Sub

'全画面表示を止める
Sub NormalScreen()
  ie.FullScreen = False
End Sub

'最大化する
Sub MaximumWindow()
  excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & ie.Hwnd & ", 3)"
End Sub

'最小化する
Sub MinimumWindow()
  excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & ie.Hwnd & ", 2)"
End Sub

'標準表示にする
Sub NormalWindow()
  excel.ExecuteExcel4Macro "CALL(""user32"", ""ShowWindow"", ""JJJ"", " & ie.Hwnd & ", 1)"
End Sub

'待機する
Sub Sleep(sec)
  WScript.Sleep(sec * 1000)
End Sub

'URLで遷移する
Sub Navigate(url)
  ie.Navigate url
  IEWait(ie)
End Sub

'子画面をアクティブにする
Sub ActivateChildWindow()
  Redim Preserve ies(Ubound(ies) + 1)
  Redim Preserve idxIes(Ubound(idxIes) + 1)
  Set ies(Ubound(ies)) = shl.Windows(shl.Windows.Count - 1)
  Set ie = ies(Ubound(ies))
  idxIes(Ubound(idxIes)) = ActivateLastIE
  IEWait(ie)
End Sub

'親画面をアクティブにする
Sub ActivateParentWindow()
  Redim Preserve ies(Ubound(ies) - 1)
  Redim Preserve idxIes(Ubound(idxIes) - 1)
  Set ie = ies(Ubound(ies))
  ActivateWindow idxIes(Ubound(ies))
  IEWait(ie)
End Sub

'指定フレームを活性にする
Sub ActivateFrame(idxFrame)
  Set doc = ie.document.frames(idxFrame).document
End Sub

'フォーカスを当てる
Sub Focus(expElm)
  Set elm = GetElement(expElm)
  elm.Focus
End Sub

'入力する（Value）
Sub ValueInput(expOptsSet)
  Input expOptsSet, 0
End Sub

'入力する（Copy&Paste）
Sub ValuePaste(expOptsSet)
  Input expOptsSet, 1
End Sub

'入力する（SendKeys）
Sub KeyInput(expOptsSet)
  Input expOptsSet, 2
End Sub

'クリックする
Sub Click(expElm)
  Set elm = GetElement(expElm)
  elm.Focus
  elm.Click
  IEWait(ie)
End Sub

'文字列をコピー&ペーストする。
Sub CopyAndPaste(str)
  CopyText str
  Wscript.Sleep 500
  wsh.SendKeys "^(v)", True
  IEWait(ie)
End Sub

'キーを押す
Sub SendKeys(key)
  wsh.SendKeys key, True
  IEWait(ie)
End Sub

'スクリーンショットを撮る（画面全体, 表示箇所のみ）
Sub FullScreenShot4VisibleArea(msg)
  WScript.Sleep(1000)
  Call KeybdEvent(&H2C, 0, 1, 0)
  Call KeybdEvent(&H2C, 0, 3, 0)
  WScript.Sleep(1000)
  shtSS.Activate
  Set rng = shtSS.Range( _
    EXCEL_STARTPRINT_CELLADDRESS _
      ).Offset(EXCEL_ONEPAGE_ROWS * idxPasteArea, 0)
  rng.Value = msg
  rng.Offset(1, 1).Select
  shtSS.Paste
  Set rng = Nothing
  idxPasteArea = idxPasteArea + 1
End Sub

'スクリーンショットを撮る（画面全体）
Sub FullScreenShot(msg)
  RepeatScreenShot True, msg
End Sub

'スクリーンショットを撮る（アクティブ画面, 表示箇所のみ）
Sub ScreenShot4VisibleArea(msg)
  Call KeybdEvent(&H12, 0, 1, 0)
  FullScreenShot4VisibleArea msg
  Call KeybdEvent(&H12, 0, 3, 0)
End Sub

'スクリーンショットを撮る（アクティブ画面）
Sub ScreenShot(msg)
  RepeatScreenShot False, msg
End Sub

'SQL文を発行する
Sub ExecuteSQL(sql)
  Dim rs, fld
  Set rs = CreateObject("ADODB.Recordset")
  Dim cmd
  cmd = Replace(sql, OPTIONROW_SEPERATE_KEYWORD, vbCrLf)
  rs.Open cmd , con, 1, 1
  Dim cntClm
  cntClm = 1
  ' SQL文を記録
  Set rng = shtDB.Range(EXCEL_STARTSET_CELLADDRESS)
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
  idxSetArea = idxSetArea + EXCEL_INTERVAL_ROWS
  rs.Close
  Set rng = Nothing
  Set fld = Nothing
  Set rs = Nothing
End Sub

'===== 本処理 =====
'【PaperTester.xlsxで生成されたVBScriptコマンドをここに貼り付ける。】


'===== 後処理 =====
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

Msgbox "処理が正常終了しました。"
