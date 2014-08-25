'=====  設定  =====
'スクリーンショットをEXCELへ貼り付ける行間隔
Dim EXCEL_ONEPAGE_ROWS
EXCEL_ONEPAGE_ROWS = 61

'===== 前処理 =====
Dim i, j, k
Dim wsh
Set wsh = WScript.CreateObject("WScript.Shell")
Dim shl
Set shl = CreateObject("Shell.Application")
Dim xls, sht
Set xls = WScript.CreateObject("Excel.Application")
xls.Application.Visible = True
xls.Application.DisplayAlerts = False
xls.Application.Workbooks.Add
Set sht = xls.Worksheets(1)

Dim ies(), idxIes(), ie
Redim ies(0)
Redim idxIes(0)
Dim doc
Dim elm

Dim wLoc, wSvc, wEnu, wIns
Set wLoc = CreateObject("WbemScripting.SWbemLocator")
Set wSvc = wLoc.ConnectServer
Set wEnu = wSvc.InstancesOf("Win32_Process")

Dim idxPasteSS
idxPasteSS = 1

'IEの遷移待ち
Sub IEWait(ie)
  Do While ie.Busy = True Or ie.readyState <> 4
  Loop
  Set doc = ie.document
End Sub

'指定ウィンドウのアクティブ化
Sub ActivateWindow(processId)
  While not wsh.AppActivate(processId) 
    Wscript.Sleep 100 
  Wend 
End Sub

'最後に起動したIEのアクティブ化
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

'キーボード入力
Sub KeybdEvent(bVk, bScan, dwFlags, dwExtraInfo)
  Call xls.ExecuteExcel4Macro(Replace(Replace(Replace(Replace("CALL(""user32"",""keybd_event"",""JJJJJ"", %0, %1, %2, %3)", "%0", bVk), "%1", bScan), "%2", dwFlags), "%3", dwExtraInfo))
End Sub

'スクリーンショット
Sub ScreenShot
  Call KeybdEvent(&H2C, 0, 1, 0)
  Call KeybdEvent(&H2C, 0, 3, 0)
  WScript.Sleep(2 * 1000)
  sht.Activate
  sht.Range("A" & idxPasteSS).Select
  sht.Paste
  idxPasteSS = idxPasteSS + EXCEL_ONEPAGE_ROWS
End Sub

'スクリーンショット（アクティブ画面）
Sub ActiveScreenShot
  Call KeybdEvent(&H12, 0, 1, 0)
  ScreenShot
  Call KeybdEvent(&H12, 0, 3, 0)
End Sub

'=====個別処理=====

'【テスト仕様書で生成された操作コマンドをここに記入する。 】


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
  Set ies(i) = Nothing
Next
Set sht = Nothing
Set xls = Nothing
Set wsh = Nothing
Set shl = Nothing

Msgbox "処理が正常終了しました。"