Option Explicit

'===== ライセンス =====
'PaperTester
'Copyright (c) 2014 nezuq
'This software is released under the MIT License.
'https://github.com/nezuQ/PaperTester/blob/master/LICENSE.txt

'===== 前処理 =====
Dim hmsStart
hmsStart = Now
Dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(".\PaperTester.vbs", 1, False).ReadAll()
Set fso = Nothing
Dim pt
Set pt = New PaperTester

'終了メッセージの取得
Private Function getEndMsg()
  Dim hmsEnd
  hmsEnd = Now
  Dim mntDiff
  mntDiff = DateDiff("n", hmsStart, hmsEnd)
  getEndMsg = _
    "開始日時=" & FormatDateTime(hmsStart, 4) & _
      ", 終了日時=" & FormatDateTime(hmsEnd, 4) & _
      ", 経過時間=" & mntDiff & "分" 
End Function

'例外処理
Private Sub onErrorExit(msg)
  Dim msgErr
  If (Err.Number <> 0) Then
    msgErr = _
      "【異常終了】" & getEndMsg() & vbCrLf _
      & "例外番号 : " & Err.Number & vbCrLf _
      & "例外説明 : " & Err.Description & vbCrLf _
      & "追加説明 : " & msg
    pt.Terminate
    WScript.Echo msgErr
    WScript.Quit
  End If
End Sub

'===== 設定値 =====
pt.EvidenceBookPath = ".\EvidenceTemplate.xlsx"
pt.ScreenshotSheetName = "Screenshot"
pt.ScreenshotPrintCellAddress = "B3"
pt.ScreenshotPageRows = 62
pt.AfterValidationLogRows = 2
pt.VerticalScrollRate = 0.80
pt.DatabaseSheetName = "Database"
pt.DataPrintCellAddress = "B3"
pt.DataIntervalRows = 2
pt.ConnectionString = ""

pt.Initialize

On Error Resume Next

'===== 本処理 =====
'※PaperTester.xlsxの操作コマンド列のVBScriptコマンドをここに貼り付ける。


'===== 後処理 =====
On Error Goto 0
Set pt = Nothing
WScript.Echo "【正常終了】" & getEndMsg()
