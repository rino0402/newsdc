Option Explicit

Dim objIE

'IE�I�u�W�F�N�g���쐬���܂�
Set objIE = CreateObject("InternetExplorer.Application")

'�E�B���h�E�̑傫����ύX���܂�
objIE.Width = 800
objIE.Height = 600

'�\���ʒu��ύX���܂�
objIE.Left = 0
objIE.Top = 0

'�X�e�[�^�X�o�[�ƃc�[���o�[���\���ɂ��܂�
objIE.Statusbar = False
objIE.ToolBar = False

'�C���^�[�l�b�g�G�N�X�v���[����ʂ�\�����܂�
objIE.Visible = True

'�@�w�肵��URL��\�����܂�
objIE.Navigate "http://mint/newsdc/tanamake.php"

'�A�y�[�W�̓ǂݍ��݂��I���܂ŃR�R�ŃO���O�����
Do Until objIE.Busy = False
   '�󃋁[�v���Ɩ��ʂ�CPU���g���̂�250�~���b�̃C���^�[�o����u��
   WScript.sleep(250)
Loop

'�X�e�[�^�X�o�[�ƃc�[���o�[��\�����܂�
objIE.Statusbar = True
objIE.ToolBar = True

WScript.Sleep 5000

objIE.Quit
Set objIE = Nothing
