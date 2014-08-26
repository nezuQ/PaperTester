PaperTester
===========

#概要
Excel製のInternetExplorer向けUIテストツール。  
http://qiita.com/nezuq/items/d2ff540cdba00d41bfda

#操作方法
PaperTester.xlsxのテスト仕様書シートの操作名列に、dataシートの操作名列の値を入力する。  
又、必要に応じて引数(%N)の値を引数列に入力する。  
操作コマンド列にUIテスト用のVBScriptコードが、操作内容列にそのコードの動作説明が、表示される。  
先のVBScriptコードをPaperTester.vbsの個別処理欄に入力・保存し、PaperTester.vbsをダブルクリックする。  
IEとExcelが自動操作され、IEのスクリーンショットがExcelに自動で貼り付けられる。  
