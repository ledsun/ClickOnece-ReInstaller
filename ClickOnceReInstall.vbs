Option Explicit

''�C���X�g�[���p�X
const INSTALL_PATH = "http://hoge.example.com/oreoreApply/ClickOneceSample.application"

'' �A���C���X�g�[���R�}���h�̓����Ă��郌�W�X�g��
const REGISTRY_PATH = "Software\Microsoft\Windows\CurrentVersion\Uninstall"

'' ����
UnInstallAll()
'' PC�̃X�y�b�N�ɂ���ăA���C���X�g�[���Ɏ��Ԃ��W��̂ł�����Ƒ҂�
Wscript.Sleep 500
Install()

'' �������� �֐� ��������
''���W�X�g������C���X�g�[������Ă���ClickOnece�A�v�����������ăA���C���X�g�[��
Sub UnInstallAll
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")

	''�N�����̃A�v���������KILL
	objShell.Run "taskkill /f /im ClickOneceSample.exe*"

	'' ���W�X�g������UninstallString���擾���Ď��s���܂��B
	Dim subkey
	For Each subkey In GetSubKeyForUninstall()
		Dim registryKey
		registryKey = "HKEY_CURRENT_USER\" & REGISTRY_PATH & "\" & subkey &"\"
                  '' ���Ђ̃A�v�����A���C���X�g�[�����Ȃ��悤�ɁA���s�Җ��ōi�荞��
		If objShell.RegRead(registryKey & "Publisher") = "LEDSUN" Then
			Uninstall(registryKey)
		END IF
	Next
End Sub

''REGISTRY_PATH�z���̃T�u�L�[���擾
Function GetSubKeyForUninstall
	const HKEY_CURRENT_USER  = &H80000001
	const HKEY_LOCAL_MACHINE = &H80000002

	dim strComputer
	strComputer = "."
	Dim oReg
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")  
	Dim arrSubKeys
	IF oReg.EnumKey(HKEY_CURRENT_USER, REGISTRY_PATH, arrSubKeys) = 2 THEN
		GetSubKeyForUninstall = Nothing
	ELSE
		GetSubKeyForUninstall = arrSubKeys
	END IF
End Function

''�w��̃��W�X�g���L�[����ClickOnece�A�v���P�[�V�������A���C���X�g�[��
Sub Uninstall(registryKey)
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run objShell.RegRead(registryKey & "UninstallString")

	''�A���C���X�g�[���_�C�A���O�̕\���҂�
	Dim Success
	Do Until Success = True
	    Success = objShell.AppActivate("ClickOneceSample �̕ێ�")
	    Wscript.Sleep 200
	Loop

	''OK���������A���C���X�g�[�������s����B
	''�ȑO�̃o�[�W�����ɖ߂����I���ł���ꍇ�ɔ����ă^�u�A������͂���B
	objShell.SendKeys "{TAB}"
	objShell.SendKeys "{DOWN}"
	objShell.SendKeys "OK"
End Sub

''�C���X�g�[��
''IE���N������MASCOT OfflineTool���C���X�g�[�����܂��B
Sub Install
	Dim objIE
	Set objIE = CreateObject("InternetExplorer.Application")
	objIE.Visible = True
	objIE.Navigate INSTALL_PATH

	''�C���X�g�[���̃_�C�A���O�̕\���҂�
	Dim objShell2
	Set objShell2 = WScript.CreateObject("WScript.Shell")
	Dim Success2
	Do Until Success2 = True
	    Success2 = objShell2.AppActivate("�A�v���P�[�V�����̃C���X�g�[�� - �Z�L�����e�B�̌x��")
	    Wscript.Sleep 200
	Loop

	''�C���X�g�[�������������s����B
	objShell2.SendKeys "Install"
End Sub
