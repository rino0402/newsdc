'* �V���[�g�J�b�g�̎Q�Ɛ���ꊇ�ύX����X�N���v�g
'* http://pnpk.net/cms/archives/2231
'* 
'* ���g����
'* NEW_TARGET_UNC�Ŏw�肵���Q�Ɛ��PRECEDENT_TARGET_UNC������������X�N���v�g�ł��B
'* �t�@�C���T�[�o�ڍs��Ƀt�@�C���T�[�o��UNC���ύX�ɂȂ����ꍇ�A
'* �e�N���C�A���g��Ɏc���Ă���V���[�g�J�b�g�̎Q�Ɛ��ύX����̂����ɖʓ|�������̂ō쐬���܂����B
'* 
'* ���̃X�N���v�g���e�L�X�g�t�@�C���ɕۑ����A�g���q��".vbs"�ɕύX���Ă��������B
'* �V���[�g�J�b�g�t�@�C�����h���b�O���h���b�v�ŐV����UNC�p�X�ɕύX���܂��B
'* ��x�ɕ����̃t�@�C�����h���b�O���h���b�v�ňꊇ�X�V���鎖���\�ł����A���܂葽���ƃG���[�ɂȂ�܂��B
'* 
'* ������
'* ���̃X�N���v�g�͎��s����ƁA���[�U�̓��Ӗ����ɑΏۃt�@�C���̃����N���ύX���܂��B
'* ���s�O�ɂ͕K���o�b�N�A�b�v������Ă�����s���Ă��������B
'* 
'* �ǂ݂Ƃ��p�̃t�@�C�����w�肵���ꍇ��A
'* �X�V�����̖����t�@�C���ɑ΂��đ�����s���ƃG���[�ɂȂ�܂��B
'* �܂��A���݂��Ȃ��p�X����͂���Ə��������ɒx���Ȃ�܂��B
Option Explicit

'* -----------------�ݒ肱������--------------------
'* PRECEDENT_TARGET_UNC�ɏ���������������UNC�p�X�A�������̓����N��o�^���āA
'* NEW_TARGET_UNC�ɏ������������UNC�p�X�A�������̓����N��o�^���܂��B
'* �Ώۃt�@�C���͊g���q��".lnk"�A��������".url"�̃t�@�C���݂̂ł��B
'* 
'* �듮����ɗ͉�������邽�߂ɁA�擪����̃p�X�ύX�݂̂���񂾂����s���܂��B
'* �@���Ⴆ��"\\FileServer01\hogehoge\FileServer01"�Ƃ����p�X�ɑ΂���
'* �@�@"FileServer01"������������ꍇ�ł��A�������Ƃ��Ă������ς��͍̂ŏ���
'* �@�@"FileServer01"�����ł��B
'* �܂��A�p�X�ɂ̓h���C�u�������p���鎖���o���܂��B
Const PRECEDENT_TARGET_UNC = "\\w1"
Const NEW_TARGET_UNC       = "\\w5"

'* -----------------�ݒ肱���܂�--------------------

Call Main()

Private Sub Main()
	Dim objFS
	'Dim strPATH 2012/08/30 �R�����g�A�E�g���܂���
	Dim strFile
	
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	'strPATH = objFS.GetParentFolderName(WScript.Arguments(0)) 2012/08/30 �R�����g�A�E�g���܂���

	For Each strFile In WScript.Arguments
		'JAGDE_SHORTCUT�̌Ăяo��
		If JAGDE_SHORTCUT(strFile) = True Then
			'REWRITE_SHORTCUT�̌Ăяo��
			Call REWRITE_SHORTCUT(strFile)
		Else
			
		End If
	Next
	
'	Wscript.Echo "�V���[�g�J�b�g���������������������܂����B"
	
End Sub

'�V���[�g�J�b�g���ǂ����̔���
Function JAGDE_SHORTCUT(strFile)
	If UCase(Right(strFile, 4)) = ".LNK" OR UCase(Right(strFile, 4)) = ".URL" Then
		JAGDE_SHORTCUT = True
	Else
		JAGDE_SHORTCUT = False
	End If
End Function

'�V���[�g�J�b�g�̏�������
Function REWRITE_SHORTCUT(strFile)
	On Error Resume Next

	Dim WshShell
	Dim objShellLink
	Dim strSHORTCUT_PATH
	
	set WshShell = WScript.CreateObject("WScript.Shell")
	set objShellLink = WshShell.CreateShortcut(strFile)
	
	strSHORTCUT_PATH = objShellLink.TargetPath
	'�擪���當�����]�����Ă��܂��B
	If UCase(Left(strSHORTCUT_PATH,Len(PRECEDENT_TARGET_UNC))) = UCase(PRECEDENT_TARGET_UNC) Then
		objShellLink.TargetPath = Replace(strSHORTCUT_PATH,PRECEDENT_TARGET_UNC,NEW_TARGET_UNC,1,1,1)
		objShellLink.Save
	End If
	WScript.Echo "TargetPath=" & objShellLink.TargetPath
	'�G���[����
	If Err <> 0 Then
		If Err = -2147024891 Then
			WScript.Echo "�G���[���������܂����B" & vbCrLf & vbCrLf &_
						 "�ȉ��̃t�@�C���ɑ΂��鏑�����݌������s�����Ă��邽��" & vbCrLf &_
						 "�t�@�C�����X�V���鎖���o���܂���ł����B" & vbCrLf & vbCrLf &_
						 strFile & vbCrLf & vbCrLf &_
						 "�t�@�C���̃A�N�Z�X�����A�܂��͓ǂݎ�葮�����m�F���Ă��������B"& vbCrLf &_
						 "���̃t�@�C���ɑ΂��鏈���̓X�L�b�v���܂��B"
		Else
			WScript.Echo Err.Number & " : " & Err.Description
		End If
	End If

	
End Function
