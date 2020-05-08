Option Explicit

Dim objIE

'IEオブジェクトを作成します
Set objIE = CreateObject("InternetExplorer.Application")

'ウィンドウの大きさを変更します
objIE.Width = 800
objIE.Height = 600

'表示位置を変更します
objIE.Left = 0
objIE.Top = 0

'ステータスバーとツールバーを非表示にします
objIE.Statusbar = False
objIE.ToolBar = False

'インターネットエクスプローラ画面を表示します
objIE.Visible = True

'①指定したURLを表示します
objIE.Navigate "http://mint/newsdc/tanamake.php"

'②ページの読み込みが終わるまでココでグルグル回る
Do Until objIE.Busy = False
   '空ループだと無駄にCPUを使うので250ミリ秒のインターバルを置く
   WScript.sleep(250)
Loop

'ステータスバーとツールバーを表示します
objIE.Statusbar = True
objIE.ToolBar = True

WScript.Sleep 5000

objIE.Quit
Set objIE = Nothing
