' --------------------------------------------------------
' ���[������M����T���v��(VBS)
' Basp21.dll��Bsmtp.dll��C:\Windows�ɃR�s�[���Ă��܂�
' [Regsvr32.exe Basp21.dll]�����s���Ă��܂�

' ���[������MAPI�̐錾
Set bobj = CreateObject("Basp21")
    
svname	= "ns"			' POP3�T�[�o�}�V����
user	= "newsdc9"		' ���[���{�b�N�X��
pass	= "123daa@Z"	' �p�X���[�h
dirname = "rcvtemp"		' �ۑ��f�B���N�g����
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
