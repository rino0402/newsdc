WScript.Echo "AAA"
dim	strRetVal
dim	strHTML
strHTML = "http://thira.plavox.info/transport/api/?t=fukutsu&no=41103160350"
strHTML = "http://yahoo.co.jp/"
WScript.Echo GetHtmlSource(strHTML,strRetVal,0,"","")
WScript.Echo strRetVal

'####################################################################################
'#
'# �֐����FGetHtmlSource
'#-----------------------------------------------------------------------------------
'# �@�\  �F�w���URL����HTML�\�[�X���擾����
'# ����  �FstrURL       I URL
'#         strRetVal    O �擾����������
'#         isSJIS       I �\�[�X�� Shift-JIS �̏ꍇ True
'#         strID        I �h���C���F�؂��K�v�ȏꍇ�̃��[�U�[ID
'#         strPass      I �h���C���F�؂��K�v�ȏꍇ�̃p�X���[�h
'# �߂�l�FTrue ����AFalse ���s
'#
'####################################################################################
Private Function GetHtmlSource(ByVal strURL, _
                               ByRef strRetVal, _
                      		   ByVal isSJIS, _
                      		   ByVal strID, _
                               ByVal strPass)

    Dim oHttp

    '�I�u�W�F�N�g�ϐ��ɎQ�Ƃ��Z�b�g���܂�
On Error Resume Next
    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    If (Err.Number <> 0) Then
        Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
		WScript.Echo "MSXML.XMLHTTPRequest"
    End If
On Error GoTo 0
    If oHttp Is Nothing Then
        MsgBox "XMLHTTP �I�u�W�F�N�g���쐬�ł��܂���ł����B", vbCritical
        Exit Function
    End If

    '�h���C���F�؂��K�v�ȏꍇ
    If strID <> "" Then
        oHttp.Open "GET", strURL, False, strID, strPass
    Else
        oHttp.Open "GET", strURL, False
    End If
    Call oHttp.Send(Null)

	do
		WScript.Echo "oHttp.readyState=" & oHttp.readyState
'		WScript.Echo "oHttp.Status=" & oHttp.Status
'		WScript.Echo "oHttp.statusText=" & oHttp.statusText

'		if oHttp.readyState = 1 then
			exit do
'		end if
	loop


    '���s�����ꍇ�͊֐����I�����܂��B
    If (oHttp.Status < 200 Or oHttp.Status >= 300) Then Exit Function

    '�\�[�X���i�[���܂�
    If isSJIS Then
        '�\�[�X�� Shift-JIS �̏ꍇ
        strRetVal = StrConv(oHttp.responseBody, vbUnicode)
    Else
        '�\�[�X�� Unicode �̏ꍇ
        strRetVal = oHttp.responseText
    End If

    '�I�u�W�F�N�g�ϐ��̎Q�Ƃ�������܂�
    Set oHttp = Nothing

    '�߂�l���Z�b�g���܂�
    GetHtmlSource = True

End Function

