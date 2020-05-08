'* ショートカットの参照先を一括変更するスクリプト
'* http://pnpk.net/cms/archives/2231
'* 
'* ■使い方
'* NEW_TARGET_UNCで指定した参照先にPRECEDENT_TARGET_UNCを書き換えるスクリプトです。
'* ファイルサーバ移行後にファイルサーバのUNCが変更になった場合、
'* 各クライアント上に残っているショートカットの参照先を変更するのが非常に面倒だったので作成しました。
'* 
'* このスクリプトをテキストファイルに保存し、拡張子を".vbs"に変更してください。
'* ショートカットファイルをドラッグ＆ドロップで新しいUNCパスに変更します。
'* 一度に複数のファイルをドラッグ＆ドロップで一括更新する事が可能ですが、あまり多いとエラーになります。
'* 
'* ■注意
'* このスクリプトは実行すると、ユーザの同意無しに対象ファイルのリンク先を変更します。
'* 実行前には必ずバックアップを取ってから実行してください。
'* 
'* 読みとり専用のファイルを指定した場合や、
'* 更新権限の無いファイルに対して操作を行うとエラーになります。
'* また、存在しないパスを入力すると処理が非常に遅くなります。
Option Explicit

'* -----------------設定ここから--------------------
'* PRECEDENT_TARGET_UNCに書き換えたい元のUNCパス、もしくはリンクを登録して、
'* NEW_TARGET_UNCに書き換えた後のUNCパス、もしくはリンクを登録します。
'* 対象ファイルは拡張子が".lnk"、もしくは".url"のファイルのみです。
'* 
'* 誤動作を極力回避させるために、先頭からのパス変更のみを一回だけ実行します。
'* 　※例えば"\\FileServer01\hogehoge\FileServer01"というパスに対して
'* 　　"FileServer01"を書き換える場合でも、あったとしても書き変わるのは最初の
'* 　　"FileServer01"だけです。
'* また、パスにはドライブ名も利用する事が出来ます。
Const PRECEDENT_TARGET_UNC = "\\w1"
Const NEW_TARGET_UNC       = "\\w5"

'* -----------------設定ここまで--------------------

Call Main()

Private Sub Main()
	Dim objFS
	'Dim strPATH 2012/08/30 コメントアウトしました
	Dim strFile
	
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	'strPATH = objFS.GetParentFolderName(WScript.Arguments(0)) 2012/08/30 コメントアウトしました

	For Each strFile In WScript.Arguments
		'JAGDE_SHORTCUTの呼び出し
		If JAGDE_SHORTCUT(strFile) = True Then
			'REWRITE_SHORTCUTの呼び出し
			Call REWRITE_SHORTCUT(strFile)
		Else
			
		End If
	Next
	
'	Wscript.Echo "ショートカット書き換え処理が完了しました。"
	
End Sub

'ショートカットかどうかの判定
Function JAGDE_SHORTCUT(strFile)
	If UCase(Right(strFile, 4)) = ".LNK" OR UCase(Right(strFile, 4)) = ".URL" Then
		JAGDE_SHORTCUT = True
	Else
		JAGDE_SHORTCUT = False
	End If
End Function

'ショートカットの書き換え
Function REWRITE_SHORTCUT(strFile)
	On Error Resume Next

	Dim WshShell
	Dim objShellLink
	Dim strSHORTCUT_PATH
	
	set WshShell = WScript.CreateObject("WScript.Shell")
	set objShellLink = WshShell.CreateShortcut(strFile)
	
	strSHORTCUT_PATH = objShellLink.TargetPath
	'先頭から文字列を評価しています。
	If UCase(Left(strSHORTCUT_PATH,Len(PRECEDENT_TARGET_UNC))) = UCase(PRECEDENT_TARGET_UNC) Then
		objShellLink.TargetPath = Replace(strSHORTCUT_PATH,PRECEDENT_TARGET_UNC,NEW_TARGET_UNC,1,1,1)
		objShellLink.Save
	End If
	WScript.Echo "TargetPath=" & objShellLink.TargetPath
	'エラー処理
	If Err <> 0 Then
		If Err = -2147024891 Then
			WScript.Echo "エラーが発生しました。" & vbCrLf & vbCrLf &_
						 "以下のファイルに対する書き込み権限が不足しているため" & vbCrLf &_
						 "ファイルを更新する事が出来ませんでした。" & vbCrLf & vbCrLf &_
						 strFile & vbCrLf & vbCrLf &_
						 "ファイルのアクセス権限、または読み取り属性を確認してください。"& vbCrLf &_
						 "このファイルに対する処理はスキップします。"
		Else
			WScript.Echo Err.Number & " : " & Err.Description
		End If
	End If

	
End Function
