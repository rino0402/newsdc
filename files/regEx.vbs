Set regEx = New RegExp
 
regEx.Global		= True	' �S�Č���
regEx.IgnoreCase	= True	' �啶���Ə���������ʂ��Ȃ�
 
'***********************************************************
' �ꊇ�u��
'***********************************************************
' �J�n���� + �I�������Ŗ��������W�� + �I������
regEx.Pattern = """" & "[^""]+" & """"
 
strText = _
"--abcd""efgh""hijk""lmno""pqrs---" & vbCrLf & _
"--1234""����""5678""�\��""9ABC---"
 
strResult = regEx.Replace( strText, """�u������܂���""" )	' �u�����܂��B
 
WScript.echo strResult
WScript.echo
