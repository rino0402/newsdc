' --------------------------------------------------------
' メールを受信するサンプル(VBS)
' Basp21.dllとBsmtp.dllをC:\Windowsにコピーしています
' [Regsvr32.exe Basp21.dll]を実行しています

' メール送受信APIの宣言
Set bobj = CreateObject("Basp21")
    
svname	= "ns"			' POP3サーバマシン名
user	= "newsdc9"		' メールボックス名
pass	= "123daa@Z"	' パスワード
dirname = "rcvtemp"		' 保存ディレクトリ名
outarray = bobj.RcvMail(svname,user,pass,"SAVD 1-10",dirname)
if IsArray(outarray) then	' OK ?
   for each file in outarray
      array2 = bobj.ReadMail(file,"subject:from:date:",">" & dirname)
      if IsArray(array2) then	' OK ?
        for each data in array2
			if Left(data,5) <> "Body:" then
		           wscript.echo data
			end if
        next
      end if
   next
end if
