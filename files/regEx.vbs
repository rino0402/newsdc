Set regEx = New RegExp
 
regEx.Global		= True	' 全て検索
regEx.IgnoreCase	= True	' 大文字と小文字を区別しない
 
'***********************************************************
' 一括置換
'***********************************************************
' 開始文字 + 終了文字で無い文字集合 + 終了文字
regEx.Pattern = """" & "[^""]+" & """"
 
strText = _
"--abcd""efgh""hijk""lmno""pqrs---" & vbCrLf & _
"--1234""漢字""5678""表示""9ABC---"
 
strResult = regEx.Replace( strText, """置換されました""" )	' 置換します。
 
WScript.echo strResult
WScript.echo
