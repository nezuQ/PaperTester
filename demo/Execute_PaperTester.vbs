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
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
pt.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & fs.GetAbsolutePathName(".\_database.xlsx") & "; Extended Properties=""Excel 8.0;HDR=Yes; [IMEX=1;]"";"
Set fs = Nothing

pt.Initialize

On Error Resume Next

'===== 本処理 =====
'※PaperTester.xlsxの操作コマンド列のVBScriptコマンドをここに貼り付ける。

pt.OpenIE : onErrorExit "テストケース = 1, Excel行 = 2"
pt.Navigate "http://bl.ocks.org/nezuQ/raw/9719897/" : onErrorExit "テストケース = 1, Excel行 = 3"
pt.MaximumWindow : onErrorExit "テストケース = 1, Excel行 = 4"
pt.FullScreenShot4VisibleArea "1" : onErrorExit "テストケース = 1, Excel行 = 5"
pt.Record2ValidateTitle "PxCSV ～Pixiv検索結果をCSV形式で～" : onErrorExit "テストケース = 1, Excel行 = 6"
pt.Record2ValidateAttribute "id=ddlEndpoint <- '0' %|% id=txtQuery <- 'あああああ '" : onErrorExit "テストケース = 1, Excel行 = 7"
pt.ExecuteSQL "SELECT * FROM [Sheet1$] " : onErrorExit "テストケース = , Excel行 = 8"

pt.ValueInput "id=ddlEndpoint <- '1' %|% id=ddlSearchType <- '1' %|% id=txtQuery <- '艦隊これくしょん' %|% id=txtPHPSessID <- ''" : onErrorExit "テストケース = 2-1, Excel行 = 10"
pt.FullScreenShot "2-1" : onErrorExit "テストケース = 2-1, Excel行 = 11"
pt.Click "tag=input#4" : onErrorExit "テストケース = 2-1, Excel行 = 12"
pt.ActivateNextIE : onErrorExit "テストケース = 2-1, Excel行 = 13"
pt.FullScreenShot "" : onErrorExit "テストケース = 2-1, Excel行 = 14"
pt.ExecuteSQL "SELECT * FROM [Sheet1$] WHERE 列名1 = 2" : onErrorExit "テストケース = 2-1, Excel行 = 15"

pt.Quit : onErrorExit "テストケース = 3-1, Excel行 = 17"
pt.Run "notepad.exe" : onErrorExit "テストケース = 3-1, Excel行 = 18"
pt.MaximumWindow : onErrorExit "テストケース = 3-1, Excel行 = 19"
pt.Paste "【開始】任意のEXEをキー操作できます。" : onErrorExit "テストケース = 3-1, Excel行 = 20"
pt.FullScreenShot4VisibleArea "3-1" : onErrorExit "テストケース = 3-1, Excel行 = 21"
pt.Sleep 1 : onErrorExit "テストケース = 3-1, Excel行 = 22"
pt.SendKeys "%(FNN)" : onErrorExit "テストケース = 3-1, Excel行 = 23"
pt.Paste "【終了】メモ帳を開き直しました。" : onErrorExit "テストケース = 3-1, Excel行 = 24"
pt.FullScreenShot4VisibleArea "" : onErrorExit "テストケース = 3-1, Excel行 = 25"
pt.Sleep 1 : onErrorExit "テストケース = 3-1, Excel行 = 26"
pt.Quit : onErrorExit "テストケース = 3-1, Excel行 = 27"
pt.Sleep 1 : onErrorExit "テストケース = 3-1, Excel行 = 28"
pt.ExecuteJS "alert('任意のJavascriptを実行できます。')" : onErrorExit "テストケース = 3-1, Excel行 = 29"
pt.Sleep 1 : onErrorExit "テストケース = 3-1, Excel行 = 30"
pt.SendKeys "{ENTER}" : onErrorExit "テストケース = 3-1, Excel行 = 31"
pt.Quit : onErrorExit "テストケース = 3-1, Excel行 = 32"

'===== 後処理 =====
On Error Goto 0
Set pt = Nothing
WScript.Echo "【正常終了】" & getEndMsg()
