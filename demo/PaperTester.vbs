'=====  ライセンス  =====
'The MIT License (MIT)
'
'Copyright (c) 2014 nezuq
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

'=====  設定  =====
'スクリーンショットを貼り付けるEXCELブック
Dim EXCEL_PRINT_BOOKPATH
EXCEL_PRINT_BOOKPATH = ".\evidence.xlsx"
'スクリーンショットを貼り付ける開始セル
Dim EXCEL_STARTPRINT_CELLADDRESS
EXCEL_STARTPRINT_CELLADDRESS = "C4"
'スクリーンショットを貼り付ける行間隔
Dim EXCEL_ONEPAGE_ROWS
EXCEL_ONEPAGE_ROWS = 61

'=====  固定値  =====
'引数行区切りのキーワード
Dim OPTIONROW_SEPERATE_KEYWORD
OPTIONROW_SEPERATE_KEYWORD = "|"
'引数区切りのキーワード
Dim OPTION_SEPERATE_KEYWORD
OPTION_SEPERATE_KEYWORD = ","
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
Dim excel, wbk, sht, rng
Set excel = WScript.CreateObject("Excel.Application")
excel.Application.Visible = True
excel.Application.DisplayAlerts = False
Set wbk = excel.Application.Workbooks.Open(fs.GetAbsolutePathName(EXCEL_PRINT_BOOKPATH))
Set sht = excel.Worksheets(1)

Dim ies(), idxIes(), ie
Redim ies(0)
Redim idxIes(0)
Dim doc
Dim elm

Dim wLoc, wSvc, wEnu, wIns
Set wLoc = CreateObject("WbemScripting.SWbemLocator")
Set wSvc = wLoc.ConnectServer
Set wEnu = wSvc.InstancesOf("Win32_Process")

Dim cntScroll, idxPasteArea
cntScroll = 0
idxPasteArea = 0

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
  Dim aryExpOpts, aryOpt, expOpts, expOpt
  aryExpOpts = Split(expOptsSet, OPTIONROW_SEPERATE_KEYWORD)
  For Each expOpts in aryExpOpts
    aryOpt = Split(expOpts, OPTION_SEPERATE_KEYWORD)
    Set elm = GetElement(aryOpt(0))
    elm.Focus
    If (useSendKeys) Then
      wsh.SendKeys aryOpt(1)
    Else
      elm.Value = aryOpt(1)
    End If
  Next
End Sub

'特殊キーを入力する
Sub KeybdEvent(bVk, bScan, dwFlags, dwExtraInfo)
  Call excel.ExecuteExcel4Macro(Replace(Replace(Replace(Replace("CALL(""user32"",""keybd_event"",""JJJJJ"", %0, %1, %2, %3)", "%0", bVk), "%1", bScan), "%2", dwFlags), "%3", dwExtraInfo))
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
Function Scroll(goToEnd)
  If (goToEnd) Then
    ie.Navigate "javascript:scroll(0," & ie.document.body.ScrollHeight & ")"
  Else
    ie.Navigate "javascript:scrollTo(0," & ie.Height & ")"
  End If
  Wscript.Sleep 1000
End Function

'数値を切り上げする
function Ceil(Number)
  Ceil = Int(Number)
  if Ceil <> Number then
    Ceil = Ceil + 1
  end if
end function

'===== 操作用関数 =====
'InternetExplorerを開く
Sub Open()
  Set ies(0) = CreateObject("InternetExplorer.Application")
  Set ie = ies(0)
  ie.Visible = True
  ie.FullScreen = True
  idxIes(0) = ActivateLastIE
End Sub

'InternetExplorerを閉じる
Sub Close()
  ie.Quit
  Set ie = Nothing
End Sub

'最大表示にする
Sub FullScreen()
  ie.FullScreen = True
End Sub

'標準表示にする
Sub NormalScreen()
  ie.FullScreen = False
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
  Redim Preserve idxIes(Ubound(ies) - 1)
  Set ie = ies(Ubound(ies))
  ActivateWindow idxIes(Ubound(ies))
  IEWait(ie)
End Sub

'指定フレームを活性にする
Sub ActivateFrame(idxFrame)
  Set doc = ie.document.frames(idxFrame)
End Sub

'フォーカスを当てる
Sub Focus(expElm)
  Set elm = GetElement(expElm)
  elm.Focus
End Sub

'入力する（SendKeys）
Sub KeyInput(expOptsSet)
  Input expOptsSet, True
End Sub

'入力する（Value）
Sub ValueInput(expOptsSet)
  Input expOptsSet, False
End Sub

'クリックする
Sub Click(expElm)
  Set elm = GetElement(expElm)
  elm.Focus
  elm.Click
  IEWait(ie)
End Sub

'キーを押す
Sub SendKeys(key)
  wsh.SendKeys key, True
  IEWait(ie)
End Sub

'スクリーンショットを撮る（画面全体, 表示箇所のみ）
Sub FullScreenShot4VisibleArea()
  Call KeybdEvent(&H2C, 0, 1, 0)
  Call KeybdEvent(&H2C, 0, 3, 0)
  WScript.Sleep(2 * 1000)
  sht.Activate
  Set rng = sht.Range( _
    EXCEL_STARTPRINT_CELLADDRESS _
      ).Offset(EXCEL_ONEPAGE_ROWS * idxPasteArea, 0)
  rng.Select
  sht.Paste
  Set rng = Nothing
  idxPasteArea = idxPasteArea + 1
End Sub

'スクリーンショットを撮る（画面全体）
Sub FullScreenShot()
  FullScreenShot4VisibleArea
  cntScroll = Ceil(ie.document.body.ScrollHeight / ie.Height)
  Dim i
  For i = 2 To cntScroll
    Scroll (i = cntScroll)
    FullScreenShot4VisibleArea
  Next
End Sub

'スクリーンショットを撮る（アクティブ画面, 表示箇所のみ）
Sub ScreenShot4VisibleArea()
  Call KeybdEvent(&H12, 0, 1, 0)
  ScreenShot
  Call KeybdEvent(&H12, 0, 3, 0)
End Sub

'スクリーンショットを撮る（アクティブ画面）
Sub ScreenShot()
  ScreenShot4VisibleArea
  cntScroll = Ceil(ie.document.body.ScrollHeight / ie.Height)
  For i = 2 To cntScroll
    Scroll (i = cntScroll)
    ScreenShot4VisibleArea
  Next
End Sub

'===== 本処理 =====
'【テスト仕様書で生成された操作コマンドをここに記入する。 】
Open
Navigate "http://bl.ocks.org/nezuQ/raw/9719897/"
FullScreenShot4VisibleArea

ValueInput "id=ddlEndpoint,1|id=ddlSearchType,1|id=txtQuery,百合"
FullScreenShot
Click "tag=input#4"
ActivateChildWindow
FullScreenShot

'===== 後処理 =====
Set wLoc = Nothing
Set wEnu = Nothing
Set wSvc = Nothing
Set wIns = Nothing
Set elm = Nothing
Set doc = Nothing
Set sht = Nothing
Set ie = Nothing
For i = LBound(ies) to UBound(ies)
  ies(i).Quit
  Set ies(i) = Nothing
Next
Set rng = Nothing
Set sht = Nothing
Set excel = Nothing
Set wsh = Nothing
Set shl = Nothing
Set fs = Nothing

Msgbox "処理が正常終了しました。"
