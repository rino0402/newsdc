' --------------------------------------------------------
' ���[������M����T���v��(VBS)
' Basp21.dll��Bsmtp.dll��C:\Windows�ɃR�s�[���Ă��܂�
' [Regsvr32.exe Basp21.dll]�����s���Ă��܂�

' ���[������MAPI�̐錾
Set BASP21 = CreateObject("Basp21")
    
' ���[����M����у��[���{�b�N�X����폜
outary = BASP21.RcvMail( _
    "ns", "newsdc9", "123daa@Z", _
    "SAVD 1-1", ".")

' ��M���[���`�F�b�N
If IsArray(outary) Then
    outary2 = BASP21.ReadMail( _
        outary(0), "subject:from:date:", ".")
    Wscript.Echo "���[���L��:" & outary2(1)
Else
    Wscript.Echo "���[������"
End If
