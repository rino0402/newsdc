' http://qiita.com/RelaxTools/items/84c9f362465a40275fe3
'-------------------------------------------------------------------------------
' Excel�t�@�C���̉E�N���b�N�u�ǂݎ���p�ŊJ���v��L���ɂ���X�N���v�g
' 
' ExcelReadOnly.vbs
' 
' Copyright (c) 2015 Y.Watanabe
' 
' This software is released under the MIT License.
' http://opensource.org/licenses/mit-license.php
'-------------------------------------------------------------------------------
' ����m�F : Windows 7 + Excel 2010 / Windows 8 + Excel 2013
'-------------------------------------------------------------------------------
' �ȉ��Q�l�T�C�g

' ���� - �E�N���b�N���j���[�u�ǂݎ���p�ŊJ���v��\������(Excel&Word) 
' https://sites.google.com/site/universeof/tips/openasreadonly'
'-------------------------------------------------------------------------------
Option Explicit

On Error Resume Next

If WScript.Arguments.Count = 0 Then

    '�������g���Ǘ��Ҍ����Ŏ��s
    With CreateObject("Shell.Application")
        .ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ dummy", "", "runas"
    End With

    WScript.Quit

End If

If MsgBox("�G�N�X�v���[���E�N���b�N(Excel�̓ǂݎ���p)��L���ɂ��܂����H", vbYesNo + vbQuestion, "�ǂݎ���p�L����") = vbNo Then 
    WScript.Quit 
End IF

With WScript.CreateObject("WScript.Shell")

    '�V�t�g�������Ȃ��Ă����j���[���\�������悤�ɂ���悤�ɁuExtended�v�L�[���폜
    .RegDelete "HKCR\Excel.Sheet.8\shell\OpenAsReadOnly\Extended"
    .RegDelete "HKCR\Excel.Sheet.12\shell\OpenAsReadOnly\Extended"
    .RegDelete "HKCR\Excel.SheetMacroEnabled.12\shell\OpenAsReadOnly\Extended"

    Err.Clear

    '�ǂݎ���p��L���ɂ���
    .RegWrite "HKCR\Excel.Sheet.8\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.Sheet.12\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"
    .RegWrite "HKCR\Excel.SheetMacroEnabled.12\shell\OpenAsReadOnly\ddeexec\","[open(""%1"",,1,,,,,,,,,,,,1,,1)]", "REG_SZ"

End With

If Err.Number = 0 Then
    MsgBox "����Ɋ֘A�t����ύX���܂����B", vbInformation + vbOkOnly, "�ǂݎ���p�L����"
Else
    MsgBox "�G���[���������܂����B", vbCritical + vbOkOnly, "�ǂݎ���p�L����"
End IF
