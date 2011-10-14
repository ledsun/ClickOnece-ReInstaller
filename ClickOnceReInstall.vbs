Option Explicit

''インストールパス
const INSTALL_PATH = "http://hoge.example.com/oreoreApply/ClickOneceSample.application"

'' アンインストールコマンドの入っているレジストリ
const REGISTRY_PATH = "Software\Microsoft\Windows\CurrentVersion\Uninstall"

'' 処理
UnInstallAll()
'' PCのスペックによってアンインストールに時間が係るのでちょっと待つ
Wscript.Sleep 500
Install()

'' ＊＊＊＊ 関数 ＊＊＊＊
''レジストリからインストールされているClickOneceアプリを検索してアンインストール
Sub UnInstallAll
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")

	''起動中のアプリがあればKILL
	objShell.Run "taskkill /f /im ClickOneceSample.exe*"

	'' レジストリからUninstallStringを取得して実行します。
	Dim subkey
	For Each subkey In GetSubKeyForUninstall()
		Dim registryKey
		registryKey = "HKEY_CURRENT_USER\" & REGISTRY_PATH & "\" & subkey &"\"
                  '' 他社のアプリをアンインストールしないように、発行者名で絞り込み
		If objShell.RegRead(registryKey & "Publisher") = "LEDSUN" Then
			Uninstall(registryKey)
		END IF
	Next
End Sub

''REGISTRY_PATH配下のサブキーを取得
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

''指定のレジストリキーからClickOneceアプリケーションをアンインストール
Sub Uninstall(registryKey)
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run objShell.RegRead(registryKey & "UninstallString")

	''アンインストールダイアログの表示待ち
	Dim Success
	Do Until Success = True
	    Success = objShell.AppActivate("ClickOneceSample の保守")
	    Wscript.Sleep 200
	Loop

	''OKを押下しアンインストールを実行する。
	''以前のバージョンに戻すが選択できる場合に備えてタブ、下を入力する。
	objShell.SendKeys "{TAB}"
	objShell.SendKeys "{DOWN}"
	objShell.SendKeys "OK"
End Sub

''インストール
''IEを起動してMASCOT OfflineToolをインストールします。
Sub Install
	Dim objIE
	Set objIE = CreateObject("InternetExplorer.Application")
	objIE.Visible = True
	objIE.Navigate INSTALL_PATH

	''インストールのダイアログの表示待ち
	Dim objShell2
	Set objShell2 = WScript.CreateObject("WScript.Shell")
	Dim Success2
	Do Until Success2 = True
	    Success2 = objShell2.AppActivate("アプリケーションのインストール - セキュリティの警告")
	    Wscript.Sleep 200
	Loop

	''インストールを押下し実行する。
	objShell2.SendKeys "Install"
End Sub
