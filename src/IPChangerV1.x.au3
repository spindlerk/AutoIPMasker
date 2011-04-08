#include <GuiConstants.au3>

Opt("TrayMenuMode",1)

Global $version = "Version 1.7"
Global $WinTitle = "Fast IP Changer (" & $version & ") - Timmo - written in AutoIT"
Global $SettingsFile = @ScriptDir & "\Settings.ini"
Global $DefaultWelcome = 0
Global $WelcomeMSG = "Fast IP Changer (" & $version & ")" & @CRLF & "written and designed by Timmo" & @CRLF & @CRLF & "special thanks go to anyone and everyone who has used this and supplied input."
Global $WelcomeWarning = @CRLF & @CRLF & "*****WARNING*****" & @CRLF & "This program is designed to be used by personel who administrate LAN's and have a fair understanding of the Network addressing sceme's of their networks." & @CRLF & "Using this program incorrectly may stop your computer from accessing or being accessed by the network." & @CRLF & "This also means that you may loose internet connection."
Global $WelcomeExtra = @CRLF & @CRLF & 'This message can be removed in the settings.ini file and changing the "Welcome MSG" option to "0"'
Global $DefaultDebug = 0

Global $DefautTab1Name = "Config 1"
Global $DefautTab2Name = "Config 2"
Global $DefautTab3Name = "Config 3"
Global $DefautTab4Name = "Config 4"

Global $Adapters, $Adapter3
Global $DefaultGW = ""
Global $IPaddress = ""
Global $Subnet = ""
Global $DHCP = ""
Global $Mark = 0

$check1 = 0
$check2 = 0
$check3 = 0
$check4 = 0

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems1 = ""
$colItems2 = ""

$objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")

$GUI = GuiCreate ($WinTitle, 700, 300, -1, -1)
$statuslabel = GUICtrlCreateLabel ("",1,280,698,19,BitOr($SS_SIMPLE,$SS_SUNKEN))

GUICtrlCreateLabel("Network Adapter: ", 10, 15)

If IniRead($SettingsFile, "Program Options", "Welcome MSG" , $DefaultWelcome) = 1 Then
	Global $WelcomeEnable = 1
Else
	Global $WelcomeEnable = 0
EndIf

If IniRead($SettingsFile, "Program Options", "Debug", $DefaultDebug) = 1 Then
	Global $show = @SW_SHOW
Else
	Global $show = @SW_HIDE
EndIf

$NoteXPos = 8
$NoteYPos = 257
$NoteString1 = "Note: IP Change Failed, Please Check settings (and cable?)"
$NoteString2 = "Note: No IP detected please check cable/connection"
$NoteString3 = "Note: No adapters detected, please check system config"
$NoteString4 = "IP Changed successfully"
$NoteString5 = "Please make sure you have selected your adapter"
$NoteString6 = "Failed to get IP via DHCP, make sure you are connected to a DHCP server."
$NoteString7 = "Failed to set DNS server (settings might still work)"

If $WelcomeEnable = 1 Then MsgBox(0,$WinTitle,$WelcomeMSG & $WelcomeWarning & $WelcomeExtra)

$Lab1Pos1 = 105 ;left position of the stats label
$Lab1Pos2 = 55 ;top position of the stats label
GUICtrlCreateGroup("", $Lab1Pos1, $Lab1Pos2, 190, 156)
$Lab1Pos1 = $Lab1Pos1 + 8
$Lab1Pos2 = $Lab1Pos2 + 13
GUICtrlCreateLabel("IP Address: ", $Lab1Pos1, $Lab1Pos2)
$Lab1Pos2 = $Lab1Pos2 + 20
GUICtrlCreateLabel("Subnet Mask: ", $Lab1Pos1, $Lab1Pos2)
$Lab1Pos2 = $Lab1Pos2 + 20
GUICtrlCreateLabel("Default GW: ", $Lab1Pos1, $Lab1Pos2)
$Lab1Pos2 = $Lab1Pos2 + 20
GUICtrlCreateLabel("Pri DNS Server: ", $Lab1Pos1, $Lab1Pos2)
$Lab1Pos2 = $Lab1Pos2 + 20
GUICtrlCreateLabel("Pri WINS Server: ", $Lab1Pos1, $Lab1Pos2)
$Lab1Pos2 = $Lab1Pos2 + 20
GUICtrlCreateLabel("DHCP Enabled: ", $Lab1Pos1, $Lab1Pos2)
$Lab1Pos2 = $Lab1Pos2 + 20
GUICtrlCreateLabel("DHCP Server: ", $Lab1Pos1, $Lab1Pos2)
$Lab1Pos1 = $Lab1Pos1 + 85
$Lab1Pos2 = $Lab1Pos2 - 120
$label1 = GUICtrlCreateLabel("", $Lab1Pos1, $Lab1Pos2, 70)
$Lab1Pos2 = $Lab1Pos2 + 20
$label2 = GUICtrlCreateLabel("", $Lab1Pos1, $Lab1Pos2, 70)
$Lab1Pos2 = $Lab1Pos2 + 20
$label3 = GUICtrlCreateLabel("", $Lab1Pos1, $Lab1Pos2, 70)
$Lab1Pos2 = $Lab1Pos2 + 20
$label4 = GUICtrlCreateLabel("", $Lab1Pos1, $Lab1Pos2, 70)
$Lab1Pos2 = $Lab1Pos2 + 20
$label5 = GUICtrlCreateLabel("", $Lab1Pos1, $Lab1Pos2, 70)
$Lab1Pos2 = $Lab1Pos2 + 20
$label6 = GUICtrlCreateLabel("", $Lab1Pos1, $Lab1Pos2, 70)
$Lab1Pos2 = $Lab1Pos2 + 20
$label7 = GUICtrlCreateLabel("", $Lab1Pos1, $Lab1Pos2, 70)
$Lab1Pos1 = $Lab1Pos1 - 38
$Lab1Pos2 = $Lab1Pos2 + 32
$ButtonRefresh = GUICtrlCreateButton(" Refresh ", $Lab1Pos1, $Lab1Pos2, 80)
$Lab1Pos1 = $Lab1Pos1 + 38
$Lab1Pos2 = $Lab1Pos2 - 32

$Adapters = _Get_Adapters()
$Adapter = GUICtrlCreateCombo("", 100, 12, 200)
GUICtrlSetData($Adapter,$Adapters , IniRead($SettingsFile, "Last Used Adapter", "Name", $Adapter3))
if GUICtrlRead($Adapter) <> "" Then _Refresh_Config()
$GetAdp = GUICtrlCreateButton(" Get Adapters ", 310, 10, 80)

$hidewin = GUICtrlCreateButton("Tray", 665, 5, 30, 17)

$trayabout = TrayCreateItem("About")
$trayshowhidewin = TrayCreateItem("Hide")
$hide = 0
TrayCreateItem("")
$TraySubMenu = TrayCreateMenu("Enable Config")
$TrayEnableCon1 = TrayCreateItem(IniRead($SettingsFile, "IPConfig1", "Name", $DefautTab1Name), $TraySubMenu)
$TrayEnableCon2 = TrayCreateItem(IniRead($SettingsFile, "IPConfig2", "Name", $DefautTab2Name), $TraySubMenu)
$TrayEnableCon3 = TrayCreateItem(IniRead($SettingsFile, "IPConfig3", "Name", $DefautTab3Name), $TraySubMenu)
$TrayEnableCon4 = TrayCreateItem(IniRead($SettingsFile, "IPConfig4", "Name", $DefautTab4Name), $TraySubMenu)
TrayCreateItem("")
$trayexit = TrayCreateItem("Exit")

GUICtrlCreateTab(415, 15, 240, 260)
$Con1Pos1 = 425 ; lefthand row
$Con1Pos2 = 40 ; top row
$Con2Pos1 = $Con1Pos1
$Con3Pos2 = $Con1Pos2
$Con2Pos2 = $Con1Pos2
$Con3Pos1 = $Con1Pos1
$Con4Pos1 = $Con2Pos1
$Con4Pos2 = $Con3Pos2
$tab1 = GUICtrlCreateTabItem(IniRead($SettingsFile, "IPConfig1", "Name", $DefautTab1Name))
GUICtrlCreateGroup ("", $Con1Pos1, $Con1Pos2, 220, 220)
$Con1Pos1 = $Con1Pos1 + 18
$Con1Pos2 = $Con1Pos2 + 23
GUICtrlCreateLabel("Name: ", $Con1Pos1, $Con1Pos2)
$Con1Pos2 = $Con1Pos2 + 30
GUICtrlCreateLabel("IP Address: ", $Con1Pos1, $Con1Pos2)
$Con1Pos2 = $Con1Pos2 + 25
GUICtrlCreateLabel("Subnet Mask: ", $Con1Pos1, $Con1Pos2)
$Con1Pos2 = $Con1Pos2 + 25
GUICtrlCreateLabel("Default GW: ", $Con1Pos1, $Con1Pos2)
$Con1Pos2 = $Con1Pos2 + 25
GUICtrlCreateLabel("DNS Server: ", $Con1Pos1, $Con1Pos2)
$Con1Pos2 = $Con1Pos2 + 25
GUICtrlCreateLabel("WINS Server: ", $Con1Pos1, $Con1Pos2)
$Con1Pos1 = $Con1Pos1 + 40
$Con1Pos2 = $Con1Pos2 - 135
$Config1Name = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig1", "Name", $DefautTab1Name), $Con1Pos1, $Con1Pos2, 90)
$Con1Pos1 = $Con1Pos1 + 30
$Con1Pos2 = $Con1Pos2 + 30
$IPaddressSet1 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig1", "IPAddress", "0.0.0.0"), $Con1Pos1, $Con1Pos2, 90)
$Con1Pos2 = $Con1Pos2 + 25
$SubnetSet1 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig1","SubnetMask","0.0.0.0"), $Con1Pos1, $Con1Pos2, 90)
$Con1Pos2 = $Con1Pos2 + 25
$DefaultGWSet1 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig1","DefaultGW","0.0.0.0"), $Con1Pos1, $Con1Pos2, 90)
$Con1Pos2 = $Con1Pos2 + 25
$DNSSet1 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig1","DNS","0.0.0.0"), $Con1Pos1, $Con1Pos2, 90)
$Con1Pos2 = $Con1Pos2 + 25
$WINSSet1 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig1","WINS","0.0.0.0"), $Con1Pos1, $Con1Pos2, 90)
$Con1Pos1 = $Con1Pos1 - 45
$Con1Pos2 = $Con1Pos2 + 25
$DHCPCheck1 = GUICtrlCreateCheckbox("DHCP ", $Con1Pos1, $Con1Pos2)
If IniRead($SettingsFile, "IPConfig1", "Use DHCP", "0") = "1" Then GUICtrlSetState($DHCPCheck1, $GUI_CHECKED)
$Con1Pos1 = $Con1Pos1 + 60
$Con1Pos2 = $Con1Pos2 + 5
$SetConfig1 = GUICtrlCreateButton(" Set ", $Con1Pos1, $Con1Pos2, 35, 20)
$Con1Pos1 = $Con1Pos1 + 50
$ButtonSave1 = GUICtrlCreateButton(" Save ", $Con1Pos1, $Con1Pos2, 40, 20)


$tab2 = GUICtrlCreateTabItem(IniRead($SettingsFile, "IPConfig2", "Name", $DefautTab2Name))
GUICtrlCreateGroup ("", $Con2Pos1, $Con2Pos2, 220, 220)
$Con2Pos1 = $Con2Pos1 + 18
$Con2Pos2 = $Con2Pos2 + 23
GUICtrlCreateLabel("Name: ", $Con2Pos1, $Con2Pos2)
$Con2Pos2 = $Con2Pos2 + 30
GUICtrlCreateLabel("IP Address: ", $Con2Pos1, $Con2Pos2)
$Con2Pos2 = $Con2Pos2 + 25
GUICtrlCreateLabel("Subnet Mask: ", $Con2Pos1, $Con2Pos2)
$Con2Pos2 = $Con2Pos2 + 25
GUICtrlCreateLabel("Default GW: ", $Con2Pos1, $Con2Pos2)
$Con2Pos2 = $Con2Pos2 + 25
GUICtrlCreateLabel("DNS Server: ", $Con2Pos1, $Con2Pos2)
$Con2Pos2 = $Con2Pos2 + 25
GUICtrlCreateLabel("WINS Server: ", $Con2Pos1, $Con2Pos2)
$Con2Pos1 = $Con2Pos1 + 40
$Con2Pos2 = $Con2Pos2 - 135
$Config2Name = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig2", "Name", $DefautTab2Name), $Con2Pos1, $Con2Pos2, 90)
$Con2Pos1 = $Con2Pos1 + 30
$Con2Pos2 = $Con2Pos2 + 30
$IPaddressSet2 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig2","IPAddress","0.0.0.0"), $Con2Pos1, $Con2Pos2, 90)
$Con2Pos2 = $Con2Pos2 + 25
$SubnetSet2 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig2","SubnetMask","0.0.0.0"), $Con2Pos1, $Con2Pos2, 90)
$Con2Pos2 = $Con2Pos2 + 25
$DefaultGWSet2 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig2","DefaultGW","0.0.0.0"), $Con2Pos1, $Con2Pos2, 90)
$Con2Pos2 = $Con2Pos2 + 25
$DNSSet2 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig2","DNS","0.0.0.0"), $Con2Pos1, $Con2Pos2, 90)
$Con2Pos2 = $Con2Pos2 + 25
$WINSSet2 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig2","WINS","0.0.0.0"), $Con2Pos1, $Con2Pos2, 90)
$Con2Pos1 = $Con2Pos1 - 45
$Con2Pos2 = $Con2Pos2 + 25
$DHCPCheck2 = GUICtrlCreateCheckbox("DHCP ", $Con2Pos1, $Con2Pos2)
If IniRead($SettingsFile, "IPConfig2", "Use DHCP", "0") = "1" Then GUICtrlSetState($DHCPCheck2, $GUI_CHECKED)
$Con2Pos1 = $Con2Pos1 + 60
$Con2Pos2 = $Con2Pos2 + 5
$SetConfig2 = GUICtrlCreateButton(" Set ", $Con2Pos1, $Con2Pos2, 35, 20)
$Con2Pos1 = $Con2Pos1 + 50
$ButtonSave2 = GUICtrlCreateButton(" Save ", $Con2Pos1, $Con2Pos2, 40, 20)

$tab3 = GUICtrlCreateTabItem(IniRead($SettingsFile, "IPConfig3", "Name", $DefautTab3Name))
GUICtrlCreateGroup ("", $Con3Pos1, $Con3Pos2, 220, 220)
$Con3Pos1 = $Con3Pos1 + 18
$Con3Pos2 = $Con3Pos2 + 23
GUICtrlCreateLabel("Name: ", $Con3Pos1, $Con3Pos2)
$Con3Pos2 = $Con3Pos2 + 30
GUICtrlCreateLabel("IP Address: ", $Con3Pos1, $Con3Pos2)
$Con3Pos2 = $Con3Pos2 + 25
GUICtrlCreateLabel("Subnet Mask: ", $Con3Pos1, $Con3Pos2)
$Con3Pos2 = $Con3Pos2 + 25
GUICtrlCreateLabel("Default GW: ", $Con3Pos1, $Con3Pos2)
$Con3Pos2 = $Con3Pos2 + 25
GUICtrlCreateLabel("DNS Server: ", $Con3Pos1, $Con3Pos2)
$Con3Pos2 = $Con3Pos2 + 25
GUICtrlCreateLabel("WINS Server: ", $Con3Pos1, $Con3Pos2)
$Con3Pos1 = $Con3Pos1 + 40
$Con3Pos2 = $Con3Pos2 - 135
$Config3Name = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig3", "Name", $DefautTab3Name), $Con3Pos1, $Con3Pos2, 90)
$Con3Pos1 = $Con3Pos1 + 30
$Con3Pos2 = $Con3Pos2 + 30
$IPaddressSet3 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig3","IPAddress","0.0.0.0"), $Con3Pos1, $Con3Pos2, 90)
$Con3Pos2 = $Con3Pos2 + 25
$SubnetSet3 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig3","SubnetMask","0.0.0.0"), $Con3Pos1, $Con3Pos2, 90)
$Con3Pos2 = $Con3Pos2 + 25
$DefaultGWSet3 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig3","DefaultGW","0.0.0.0"), $Con3Pos1, $Con3Pos2, 90)
$Con3Pos2 = $Con3Pos2 + 25
$DNSSet3 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig3","DNS","0.0.0.0"), $Con3Pos1, $Con3Pos2, 90)
$Con3Pos2 = $Con3Pos2 + 25
$WINSSet3 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig3","WINS","0.0.0.0"), $Con3Pos1, $Con3Pos2, 90)
$Con3Pos1 = $Con3Pos1 - 45
$Con3Pos2 = $Con3Pos2 + 25
$DHCPCheck3 = GUICtrlCreateCheckbox("DHCP ", $Con3Pos1, $Con3Pos2)
If IniRead($SettingsFile, "IPConfig3", "Use DHCP", "0") = "1" Then GUICtrlSetState($DHCPCheck3, $GUI_CHECKED)
$Con3Pos1 = $Con3Pos1 + 60
$Con3Pos2 = $Con3Pos2 + 5
$SetConfig3 = GUICtrlCreateButton(" Set ", $Con3Pos1, $Con3Pos2, 35, 20)
$Con3Pos1 = $Con3Pos1 + 50
$ButtonSave3 = GUICtrlCreateButton(" Save ", $Con3Pos1, $Con3Pos2, 40, 20)

$tab4 = GUICtrlCreateTabItem(IniRead($SettingsFile, "IPConfig4", "Name", $DefautTab4Name))
GUICtrlCreateGroup ("", $Con4Pos1, $Con4Pos2, 220, 220)
$Con4Pos1 = $Con4Pos1 + 18
$Con4Pos2 = $Con4Pos2 + 23
GUICtrlCreateLabel("Name: ", $Con4Pos1, $Con4Pos2)
$Con4Pos2 = $Con4Pos2 + 30
GUICtrlCreateLabel("IP Address:", $Con4Pos1, $Con4Pos2)
$Con4Pos2 = $Con4Pos2 + 25
GUICtrlCreateLabel("Subnet Mask:", $Con4Pos1, $Con4Pos2)
$Con4Pos2 = $Con4Pos2 + 25
GUICtrlCreateLabel("Default GW:", $Con4Pos1, $Con4Pos2)
$Con4Pos2 = $Con4Pos2 + 25
GUICtrlCreateLabel("DNS Server: ", $Con4Pos1, $Con4Pos2)
$Con4Pos2 = $Con4Pos2 + 25
GUICtrlCreateLabel("WINS Server: ", $Con4Pos1, $Con4Pos2)
$Con4Pos1 = $Con4Pos1 + 40
$Con4Pos2 = $Con4Pos2 - 135
$Config4Name = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig4", "Name", $DefautTab4Name), $Con4Pos1, $Con4Pos2, 90)
$Con4Pos1 = $Con4Pos1 + 30
$Con4Pos2 = $Con4Pos2 + 30
$IPaddressSet4 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig4","IPAddress","0.0.0.0"), $Con4Pos1, $Con4Pos2, 90)
$Con4Pos2 = $Con4Pos2 + 25
$SubnetSet4 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig4","SubnetMask","0.0.0.0"), $Con4Pos1, $Con4Pos2, 90)
$Con4Pos2 = $Con4Pos2 + 25
$DefaultGWSet4 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig4","DefaultGW","0.0.0.0"), $Con4Pos1, $Con4Pos2, 90)
$Con4Pos2 = $Con4Pos2 + 25
$DNSSet4 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig4","DNS","0.0.0.0"), $Con4Pos1, $Con4Pos2, 90)
$Con4Pos2 = $Con4Pos2 + 25
$WINSSet4 = GUICtrlCreateInput(IniRead($SettingsFile, "IPConfig4","WINS","0.0.0.0"), $Con4Pos1, $Con4Pos2, 90)
$Con4Pos1 = $Con4Pos1 - 45
$Con4Pos2 = $Con4Pos2 + 25
$DHCPCheck4 = GUICtrlCreateCheckbox("DHCP ", $Con4Pos1, $Con4Pos2)
If IniRead($SettingsFile, "IPConfig4", "Use DHCP", "0") = "1" Then GUICtrlSetState($DHCPCheck4, $GUI_CHECKED)
$Con4Pos1 = $Con4Pos1 + 60
$Con4Pos2 = $Con4Pos2 + 5
$SetConfig4 = GUICtrlCreateButton(" Set ", $Con4Pos1, $Con4Pos2, 35, 20)
$Con4Pos1 = $Con4Pos1 + 50
$ButtonSave4 = GUICtrlCreateButton(" Save ", $Con4Pos1, $Con4Pos2, 40, 20)

If $CmdLine[0] = 0 Then
	GUISetState()
Else
	If $CmdLine[1] = "-p" Then
		If Not $CmdLine[2] = "" Then
			$adapter3 = IniRead($SettingsFile, "Last Used Adapter", "Name", $Adapter3) ;
			Select
				Case $CmdLine[2] = 1
					_Apply_Button("Button1")
				Case $CmdLine[2] = 2
					_Apply_Button("Button2")
				Case $CmdLine[2] = 3
					_Apply_Button("Button3")
				Case $CmdLine[2] = 4
					_Apply_Button("Button4")
				Case Else
					CLIERROR()
			EndSelect
		Else
			CLIERROR()
		EndIf
	Else
		CLIERROR()
	EndIf
	Exit
EndIf

Func CLIERROR()
	MsgBox(0, "", "Command Line Options:" & @CRLF & @CRLF & "ipchanger.exe -p <preset config(1-4)>" & @CRLF & "(Note uses the last used adapter as specified in the ini)")
	Exit
EndFunc

while 1
	_CheckInput($WinTitle, $Config1Name)
	_CheckInput($WinTitle, $Config2Name)
	_CheckInput($WinTitle, $Config3Name)
	_CheckInput($WinTitle, $Config4Name)
	
	$msg = GUIGetMsg()
	$traymsg = TrayGetMsg()
	Select
		Case $msg = $ButtonSave1 Or $msg = $ButtonSave2 Or $msg = $ButtonSave3 Or $msg = $ButtonSave4
			_Set_Gui("Disable")
			If MsgBox(4, $WinTitle, "This will overwrite the settings file. Continue?") = 6 Then _SaveIniFile()
			_Set_Gui("Enable")
			
		Case $msg = $SetConfig1
			_Set_Gui("Disable")
			If GUICtrlRead($Adapter) = "" Then 
				MsgBox(0, $WinTitle, $NoteString5)
				GUICtrlSetData($statuslabel, $NoteString5)
			Else
				_Apply_Button("Button1")
			EndIf
			_Set_Gui("Enable")
			
		Case $msg = $SetConfig2
			_Set_Gui("Disable")
			If GUICtrlRead($Adapter) = "" Then 
				MsgBox(0, $WinTitle, $NoteString5)
				GUICtrlSetData($statuslabel, $NoteString5)
			Else
				_Apply_Button("Button2")
			EndIf
			_Set_Gui("Enable")
			
		Case $msg = $SetConfig3
			_Set_Gui("Disable")
			If GUICtrlRead($Adapter) = "" Then 
				MsgBox(0, $WinTitle, $NoteString5)
				GUICtrlSetData($statuslabel, $NoteString5)
			Else
				_Apply_Button("Button3")
			EndIf
			_Set_Gui("Enable")
			
		Case $msg = $SetConfig4
			_Set_Gui("Disable")
			If GUICtrlRead($Adapter) = "" Then 
				MsgBox(0, $WinTitle, $NoteString5)
				GUICtrlSetData($statuslabel, $NoteString5)
			Else
				_Apply_Button("Button4")
			EndIf
			_Set_Gui("Enable")
			
		Case $msg = $GUI_EVENT_CLOSE
			_Set_Gui("Disable")
			_Exit()
			
		Case $msg = $ButtonRefresh
			_Set_Gui("Disable")
			If GUICtrlRead($Adapter) = "" Then
				MsgBox(0, $WinTitle, $NoteString5)
				GUICtrlSetData($statuslabel, $NoteString5)
			Else
				_Refresh_Config()
			EndIf
			_Set_Gui("Enable")
			
		Case $msg = $Adapter
			_Set_Gui("Disable")
			_Refresh_Config()
			_Set_Gui("Enable")
			
		Case $msg = $GetAdp
			_Set_Gui("Disable")
			GUICtrlSetData($Adapter, _Get_Adapters(), GUICtrlRead($Adapter))
			_Refresh_Config()
			_Set_Gui("Enable")
			
		Case $msg = $hidewin
			show_hide_win()
			
		Case $traymsg = $trayexit
			_Exit()
			
		Case $traymsg = $trayabout
			MsgBox(0,$WinTitle,$WelcomeMSG & $WelcomeWarning)
			
		Case $traymsg = $trayshowhidewin
			show_hide_win()
			
		Case $traymsg = $TrayEnableCon1
			$adapter3 = IniRead($SettingsFile, "Last Used Adapter", "Name", $Adapter3)
			_Apply_Button("Button1")
			
		Case $traymsg = $TrayEnableCon2
			$adapter3 = IniRead($SettingsFile, "Last Used Adapter", "Name", $Adapter3)
			_Apply_Button("Button2")
			
		Case $traymsg = $TrayEnableCon3
			$adapter3 = IniRead($SettingsFile, "Last Used Adapter", "Name", $Adapter3)
			_Apply_Button("Button3")
			
		Case $traymsg = $TrayEnableCon4
			$adapter3 = IniRead($SettingsFile, "Last Used Adapter", "Name", $Adapter3)
			_Apply_Button("Button4")
			
	EndSelect
	If GUICtrlRead($DHCPCheck1) = $GUI_CHECKED and $check1 = 0 Then
		GUICtrlSetState($IPaddressSet1, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet1, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet1, $GUI_DISABLE)
		GUICtrlSetState($DNSSet1, $GUI_DISABLE)
		GUICtrlSetState($WINSSet1, $GUI_DISABLE)
		$check1 = 1
	EndIf
	If GUICtrlRead($DHCPCheck1) = $GUI_UNCHECKED and $check1 = 1 Then
		GUICtrlSetState($IPaddressSet1, $GUI_ENABLE)
		GUICtrlSetState($SubnetSet1, $GUI_ENABLE)
		GUICtrlSetState($DefaultGWSet1, $GUI_ENABLE)
		GUICtrlSetState($DNSSet1, $GUI_ENABLE)
		GUICtrlSetState($WINSSet1, $GUI_ENABLE)
		$check1 = 0
	EndIf
	
	If GUICtrlRead($DHCPCheck2) = $GUI_CHECKED and $check2 = 0 Then
		GUICtrlSetState($IPaddressSet2, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet2, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet2, $GUI_DISABLE)
		GUICtrlSetState($DNSSet2, $GUI_DISABLE)
		GUICtrlSetState($WINSSet2, $GUI_DISABLE)
		$check2 = 1
	EndIf
	If GUICtrlRead($DHCPCheck2) = $GUI_UNCHECKED and $check2 = 1 Then
		GUICtrlSetState($IPaddressSet2, $GUI_ENABLE)
		GUICtrlSetState($SubnetSet2, $GUI_ENABLE)
		GUICtrlSetState($DefaultGWSet2, $GUI_ENABLE)
		GUICtrlSetState($DNSSet2, $GUI_ENABLE)
		GUICtrlSetState($WINSSet2, $GUI_ENABLE)
		$check2 = 0
	EndIf
	
	If GUICtrlRead($DHCPCheck3) = $GUI_CHECKED and $check3 = 0 Then
		GUICtrlSetState($IPaddressSet3, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet3, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet3, $GUI_DISABLE)
		GUICtrlSetState($DNSSet3, $GUI_DISABLE)
		GUICtrlSetState($WINSSet3, $GUI_DISABLE)
		$check3 = 1
	EndIf
	If GUICtrlRead($DHCPCheck3) = $GUI_UNCHECKED and $check3 = 1 Then
		GUICtrlSetState($IPaddressSet3, $GUI_ENABLE)
		GUICtrlSetState($SubnetSet3, $GUI_ENABLE)
		GUICtrlSetState($DefaultGWSet3, $GUI_ENABLE)
		GUICtrlSetState($DNSSet3, $GUI_ENABLE)
		GUICtrlSetState($WINSSet3, $GUI_ENABLE)
		$check3 = 0
	EndIf
	
	If GUICtrlRead($DHCPCheck4) = $GUI_CHECKED and $check4 = 0 Then
		GUICtrlSetState($IPaddressSet4, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet4, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet4, $GUI_DISABLE)
		GUICtrlSetState($DNSSet4, $GUI_DISABLE)
		GUICtrlSetState($WINSSet4, $GUI_DISABLE)
		$check4 = 1
	EndIf
	If GUICtrlRead($DHCPCheck4) = $GUI_UNCHECKED and $check4 = 1 Then
		GUICtrlSetState($IPaddressSet4, $GUI_ENABLE)
		GUICtrlSetState($SubnetSet4, $GUI_ENABLE)
		GUICtrlSetState($DefaultGWSet4, $GUI_ENABLE)
		GUICtrlSetState($DNSSet4, $GUI_ENABLE)
		GUICtrlSetState($WINSSet4, $GUI_ENABLE)
		$check4 = 0
	EndIf
WEnd

Func ChangeConfigName()
	GUICtrlSetData($tab1, GUICtrlRead($Config1Name))
	GUICtrlSetData($tab2, GUICtrlRead($Config2Name))
	GUICtrlSetData($tab3, GUICtrlRead($Config3Name))
	GUICtrlSetData($tab4, GUICtrlRead($Config4Name))
	TrayItemDelete($TrayEnableCon1)
	TrayItemDelete($TrayEnableCon2)
	TrayItemDelete($TrayEnableCon3)
	TrayItemDelete($TrayEnableCon4)
	$TrayEnableCon1 = TrayCreateItem(GUICtrlRead($Config1Name), $TraySubMenu)
	$TrayEnableCon2 = TrayCreateItem(GUICtrlRead($Config2Name), $TraySubMenu)
	$TrayEnableCon3 = TrayCreateItem(GUICtrlRead($Config3Name), $TraySubMenu)
	$TrayEnableCon4 = TrayCreateItem(GUICtrlRead($Config4Name), $TraySubMenu)

EndFunc

Func show_hide_win()
	If $hide = 0 Then
		TrayItemSetText($trayshowhidewin, "Show")
		$hide = 1
		GUISetState(@SW_MINIMIZE, $WinTitle)
		GUISetState(@SW_HIDE, $WinTitle)
	ElseIf $hide = 1 Then
		TrayItemSetText($trayshowhidewin, "Hide")
		$hide = 0
		GUISetState(@SW_SHOW, $WinTitle)
		GUISetState(@SW_RESTORE, $WinTitle)
	Else
		Return
	EndIf
EndFunc

Func _CheckInput($hWnd, $ID)
    If $Mark = 0 And _IsFocused($hWnd, $ID) Then
        $Mark = 1

    ElseIf $Mark = 1 And Not _IsFocused($hWnd, $ID) Then
        $Mark = 0
		ChangeConfigName()
    EndIf
	
EndFunc

Func _IsFocused($hWnd, $nCID)
    Return ControlGetHandle($hWnd, '', $nCID) = ControlGetHandle($hWnd, '', ControlGetFocus($hWnd))
EndFunc

Func _Set_Gui($a)
	If $a = "Enable" Then
		GUICtrlSetState($DHCPCheck1, $GUI_ENABLE)
		GUICtrlSetState($DHCPCheck2, $GUI_ENABLE)
		GUICtrlSetState($DHCPCheck3, $GUI_ENABLE)
		GUICtrlSetState($DHCPCheck4, $GUI_ENABLE)
		If GUICtrlRead($DHCPCheck1) = $GUI_CHECKED Then
			GUICtrlSetState($IPaddressSet1, $GUI_DISABLE)
			GUICtrlSetState($SubnetSet1, $GUI_DISABLE)
			GUICtrlSetState($DefaultGWSet1, $GUI_DISABLE)
			GUICtrlSetState($DNSSet1, $GUI_DISABLE)
			GUICtrlSetState($WINSSet1, $GUI_DISABLE)
			$check1 = 0
		Else
			GUICtrlSetState($IPaddressSet1, $GUI_ENABLE)
			GUICtrlSetState($SubnetSet1, $GUI_ENABLE)
			GUICtrlSetState($DefaultGWSet1, $GUI_ENABLE)
			GUICtrlSetState($DNSSet1, $GUI_ENABLE)
			GUICtrlSetState($WINSSet1, $GUI_ENABLE)
		EndIf
		
		If GUICtrlRead($DHCPCheck2) = $GUI_CHECKED Then
			GUICtrlSetState($IPaddressSet2, $GUI_DISABLE)
			GUICtrlSetState($SubnetSet2, $GUI_DISABLE)
			GUICtrlSetState($DefaultGWSet2, $GUI_DISABLE)
			GUICtrlSetState($DNSSet2, $GUI_DISABLE)
			GUICtrlSetState($WINSSet2, $GUI_DISABLE)
			$check2 = 0
		Else
			GUICtrlSetState($IPaddressSet2, $GUI_ENABLE)
			GUICtrlSetState($SubnetSet2, $GUI_ENABLE)
			GUICtrlSetState($DefaultGWSet2, $GUI_ENABLE)
			GUICtrlSetState($DNSSet2, $GUI_ENABLE)
			GUICtrlSetState($WINSSet2, $GUI_ENABLE)
		EndIf
		
		If GUICtrlRead($DHCPCheck3) = $GUI_CHECKED Then
			GUICtrlSetState($IPaddressSet3, $GUI_DISABLE)
			GUICtrlSetState($SubnetSet3, $GUI_DISABLE)
			GUICtrlSetState($DefaultGWSet3, $GUI_DISABLE)
			GUICtrlSetState($DNSSet3, $GUI_DISABLE)
			GUICtrlSetState($WINSSet3, $GUI_DISABLE)
			$check3 = 0
		Else
			GUICtrlSetState($IPaddressSet3, $GUI_ENABLE)
			GUICtrlSetState($SubnetSet3, $GUI_ENABLE)
			GUICtrlSetState($DefaultGWSet3, $GUI_ENABLE)
			GUICtrlSetState($DNSSet3, $GUI_ENABLE)
			GUICtrlSetState($WINSSet3, $GUI_ENABLE)
		EndIf
		
		If GUICtrlRead($DHCPCheck4) = $GUI_CHECKED Then
			GUICtrlSetState($IPaddressSet4, $GUI_DISABLE)
			GUICtrlSetState($SubnetSet4, $GUI_DISABLE)
			GUICtrlSetState($DefaultGWSet4, $GUI_DISABLE)
			GUICtrlSetState($DNSSet4, $GUI_DISABLE)
			GUICtrlSetState($WINSSet4, $GUI_DISABLE)
			$check4 = 0
		Else
			GUICtrlSetState($IPaddressSet4, $GUI_ENABLE)
			GUICtrlSetState($SubnetSet4, $GUI_ENABLE)
			GUICtrlSetState($DefaultGWSet4, $GUI_ENABLE)
			GUICtrlSetState($DHCPCheck4, $GUI_ENABLE)
			GUICtrlSetState($DNSSet4, $GUI_ENABLE)
			GUICtrlSetState($WINSSet4, $GUI_ENABLE)
		EndIf
		GUICtrlSetState($Config1Name, $GUI_ENABLE)
		GUICtrlSetState($ButtonSave1, $GUI_ENABLE)
		GUICtrlSetState($SetConfig1, $GUI_ENABLE)
		GUICtrlSetState($Config2Name, $GUI_ENABLE)
		GUICtrlSetState($ButtonSave2, $GUI_ENABLE)
		GUICtrlSetState($SetConfig2, $GUI_ENABLE)
		GUICtrlSetState($Config3Name, $GUI_ENABLE)
		GUICtrlSetState($ButtonSave3, $GUI_ENABLE)
		GUICtrlSetState($SetConfig3, $GUI_ENABLE)
		GUICtrlSetState($Config4Name, $GUI_ENABLE)
		GUICtrlSetState($ButtonSave4, $GUI_ENABLE)
		GUICtrlSetState($SetConfig4, $GUI_ENABLE)
		GUICtrlSetState($Adapter, $GUI_ENABLE)
		GUICtrlSetState($GetAdp, $GUI_ENABLE)
		GUICtrlSetState($ButtonRefresh, $GUI_ENABLE)
		
	ElseIf $a = "Disable" Then
		GUICtrlSetState($Config1Name, $GUI_DISABLE)
		GUICtrlSetState($IPaddressSet1, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet1, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet1, $GUI_DISABLE)
		GUICtrlSetState($DHCPCheck1, $GUI_DISABLE)
		GUICtrlSetState($DNSSet1, $GUI_DISABLE)
		GUICtrlSetState($WINSSet1, $GUI_DISABLE)
		GUICtrlSetState($ButtonSave1, $GUI_DISABLE)
		GUICtrlSetState($SetConfig1, $GUI_DISABLE)
		GUICtrlSetState($Config2Name, $GUI_DISABLE)
		GUICtrlSetState($IPaddressSet2, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet2, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet2, $GUI_DISABLE)
		GUICtrlSetState($DHCPCheck2, $GUI_DISABLE)
		GUICtrlSetState($DNSSet2, $GUI_DISABLE)
		GUICtrlSetState($WINSSet2, $GUI_DISABLE)
		GUICtrlSetState($ButtonSave2, $GUI_DISABLE)
		GUICtrlSetState($SetConfig2, $GUI_DISABLE)
		GUICtrlSetState($Config3Name, $GUI_DISABLE)
		GUICtrlSetState($IPaddressSet3, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet3, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet3, $GUI_DISABLE)
		GUICtrlSetState($DHCPCheck3, $GUI_DISABLE)
		GUICtrlSetState($DNSSet3, $GUI_DISABLE)
		GUICtrlSetState($WINSSet3, $GUI_DISABLE)
		GUICtrlSetState($ButtonSave3, $GUI_DISABLE)
		GUICtrlSetState($SetConfig3, $GUI_DISABLE)
		GUICtrlSetState($Config4Name, $GUI_DISABLE)
		GUICtrlSetState($IPaddressSet4, $GUI_DISABLE)
		GUICtrlSetState($SubnetSet4, $GUI_DISABLE)
		GUICtrlSetState($DefaultGWSet4, $GUI_DISABLE)
		GUICtrlSetState($DHCPCheck4, $GUI_DISABLE)
		GUICtrlSetState($DNSSet4, $GUI_DISABLE)
		GUICtrlSetState($WINSSet4, $GUI_DISABLE)
		GUICtrlSetState($ButtonSave4, $GUI_DISABLE)
		GUICtrlSetState($SetConfig4, $GUI_DISABLE)
		GUICtrlSetState($GetAdp, $GUI_DISABLE)
		GUICtrlSetState($ButtonRefresh, $GUI_DISABLE)
	EndIf
EndFunc

Func _Apply_Button($a)
	Select 
		Case $a = "Button1"
			If GUICtrlRead($IPaddressSet1) = "" Then GUICtrlSetData($IPaddressSet1, "0.0.0.0")
			If GUICtrlRead($SubnetSet1) = "" Then GUICtrlSetData($SubnetSet1, "0.0.0.0")
			If GUICtrlRead($DefaultGWSet1) = "" Then GUICtrlSetData($DefaultGWSet1, "0.0.0.0")
			If GUICtrlRead($DNSSet1) = "" Then GUICtrlSetData($DNSSet1, "0.0.0.0")
			If GUICtrlRead($WINSSet1) = "" Then GUICtrlSetData($WINSSet1, "0.0.0.0")
			If GUICtrlRead($DHCPCheck1) = $GUI_CHECKED Then
				_Set_DHCP(GUICtrlRead($Adapter))
			Else
				_Set_ip(GUICtrlRead($Adapter), GUICtrlRead($IPaddressSet1), GUICtrlRead($SubnetSet1), GUICtrlRead($DefaultGWSet1), GUICtrlRead($DNSSet1), GUICtrlRead($WINSSet1))
			EndIf
			
		Case $a = "Button2"
			If GUICtrlRead($IPaddressSet2) = "" Then GUICtrlSetData($IPaddressSet2, "0.0.0.0")
			If GUICtrlRead($SubnetSet2) = "" Then GUICtrlSetData($SubnetSet2, "0.0.0.0")
			If GUICtrlRead($DefaultGWSet2) = "" Then GUICtrlSetData($DefaultGWSet2, "0.0.0.0")
			If GUICtrlRead($DNSSet2) = "" Then GUICtrlSetData($DNSSet2, "0.0.0.0")
			If GUICtrlRead($WINSSet2) = "" Then GUICtrlSetData($WINSSet2, "0.0.0.0")
			If GUICtrlRead($DHCPCheck2) = $GUI_CHECKED Then
				_Set_DHCP(GUICtrlRead($Adapter))
			Else
				_Set_ip(GUICtrlRead($Adapter), GUICtrlRead($IPaddressSet2), GUICtrlRead($SubnetSet2), GUICtrlRead($DefaultGWSet2), GUICtrlRead($DNSSet2), GUICtrlRead($WINSSet2))
			EndIf
			
		Case $a = "Button3"
			If GUICtrlRead($IPaddressSet3) = "" Then GUICtrlSetData($IPaddressSet3, "0.0.0.0")
			If GUICtrlRead($SubnetSet3) = "" Then GUICtrlSetData($SubnetSet3, "0.0.0.0")
			If GUICtrlRead($DefaultGWSet3) = "" Then GUICtrlSetData($DefaultGWSet3, "0.0.0.0")
			If GUICtrlRead($DNSSet3) = "" Then GUICtrlSetData($DNSSet3, "0.0.0.0")
			If GUICtrlRead($WINSSet3) = "" Then GUICtrlSetData($WINSSet3, "0.0.0.0")
			If GUICtrlRead($DHCPCheck3) = $GUI_CHECKED Then
				_Set_DHCP(GUICtrlRead($Adapter))
			Else
				_Set_ip(GUICtrlRead($Adapter), GUICtrlRead($IPaddressSet3), GUICtrlRead($SubnetSet3), GUICtrlRead($DefaultGWSet3), GUICtrlRead($DNSSet3), GUICtrlRead($WINSSet3))
			EndIf
			
		Case $a = "Button4"
			If GUICtrlRead($IPaddressSet4) = "" Then GUICtrlSetData($IPaddressSet4, "0.0.0.0")
			If GUICtrlRead($SubnetSet4) = "" Then GUICtrlSetData($SubnetSet4, "0.0.0.0")
			If GUICtrlRead($DefaultGWSet4) = "" Then GUICtrlSetData($DefaultGWSet4, "0.0.0.0")
			If GUICtrlRead($DNSSet4) = "" Then GUICtrlSetData($DNSSet4, "0.0.0.0")
			If GUICtrlRead($WINSSet4) = "" Then GUICtrlSetData($WINSSet4, "0.0.0.0")
			If GUICtrlRead($DHCPCheck4) = $GUI_CHECKED Then
				_Set_DHCP(GUICtrlRead($Adapter))
			Else
				_Set_ip(GUICtrlRead($Adapter), GUICtrlRead($IPaddressSet4), GUICtrlRead($SubnetSet4), GUICtrlRead($DefaultGWSet4), GUICtrlRead($DNSSet4), GUICtrlRead($WINSSet4))
			EndIf
			
	EndSelect
EndFunc

Func _Set_DHCP($name)
	$DHCP = _Test_DHCP($Adapter)
	ProgressOn($WinTitle, "Changing IP Address")
	GUICtrlSetData($statuslabel, "")
	$run = RunWait(@ComSpec & " /c " & 'netsh interface ip set address name="' & $name & '" source=dhcp', "", $show)
	ProgressSet(25)
	$run3 = RunWait(@ComSpec & " /c " & 'netsh int ip set dns "' & $name & '" dhcp', "", $show)
	RunWait(@ComSpec & " /c " & 'netsh interface ip set wins "' & $name & '" dhcp', "", $show)
	ProgressSet(50)
	RunWait(@ComSpec & " /c " & 'ipconfig /release "' & $name & '"', "", $show)
	ProgressSet(75)
	$run2 = RunWait(@ComSpec & " /c " & 'ipconfig /renew "' & $name & '"', "", $show)
	ProgressSet(100)
	Sleep(500)
	ProgressOff()
	If $run = 0 Then
		MsgBox(0, $WinTitle, $NoteString4)
		GUICtrlSetData($statuslabel, $NoteString4)
	ElseIf $run2 <> 0 Then
		MsgBox(0, $WinTitle, $NoteString6)
		GUICtrlSetData($statuslabel, $NoteString6)
	ElseIf $run3 <> 0 Then
		MsgBox(0, $WinTitle, $NoteString7)
	Else
		MsgBox(0, $WinTitle, $NoteString1)
		GUICtrlSetData($statuslabel, $NoteString1)
	EndIf
	_Refresh_Config()
EndFunc

Func _Set_ip($name, $IP, $Subnet, $DefaultGW, $DNS, $WINS)
	ProgressOn($WinTitle, "Changing IP Address")
	GUICtrlSetData($statuslabel, "")
	$run = RunWait(@ComSpec & " /c " & 'netsh interface ip set address name="' & $name & '" static ' & $IP & " " & $Subnet & " " & $DefaultGW & " 1", "", $show)
	ProgressSet(50)
	$rundns = RunWait(@ComSpec & " /c " & 'netsh interface ip set dns name="' & $name & '" static ' & $DNS, "", $show)
	$runwins = RunWait(@ComSpec & " /c " & 'netsh interface ip set wins name="' & $name & '" static ' & $WINS, "", $show)
	ProgressSet(100)
	Sleep(500)
	ProgressOff()
	If $run = 0 Then 
		MsgBox(0,$WinTitle, $NoteString4)
		GUICtrlSetData($statuslabel, $NoteString4)
		Sleep(1000)
	Else
		MsgBox(0,$WinTitle, $NoteString1)
		GUICtrlSetData($statuslabel, $NoteString1)
	EndIf
	_Refresh_Config()
EndFunc

Func _Refresh_Config()
	SplashTextOn($WinTitle, "Please Wait...", 170, 40)
	GUICtrlDelete($label1)
	GUICtrlDelete($label2)
	GUICtrlDelete($label3)
	GUICtrlDelete($label4)
	GUICtrlDelete($label5)
	GUICtrlDelete($label6)
	GUICtrlDelete($label7)
	GUICtrlSetData($statuslabel, "")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	for $objItem in $colItems
		If $objItem.NetConnectionID = GUICtrlRead($Adapter) Then
			$Adapter2 = $objItem.caption
			ExitLoop
		EndIf
	Next
	$IPaddress = _Refresh_IP($Adapter2)
	$Subnet = _Refresh_Subnet($Adapter2)
	$DefaultGW = _Refresh_DefaultGW($Adapter2)
	$DNSServer = _Refresh_DNSServer($Adapter2)
	$WINSServer = _Refresh_WINS($Adapter2)
	$DHCP = _Test_DHCP($Adapter2)
	$DHCPServer = _Refresh_DHCPServer($Adapter2)
	
	SplashOff()
	$Lab1Pos2 = $Lab1Pos2 - 120
	$label1 = GUICtrlCreateLabel($IPaddress, $Lab1Pos1, $Lab1Pos2)
	$Lab1Pos2 = $Lab1Pos2 + 20
	$label2 = GUICtrlCreateLabel($Subnet, $Lab1Pos1, $Lab1Pos2)
	$Lab1Pos2 = $Lab1Pos2 + 20
	$label3 = GUICtrlCreateLabel($DefaultGW, $Lab1Pos1, $Lab1Pos2)
	$Lab1Pos2 = $Lab1Pos2 + 20
	$label4 = GUICtrlCreateLabel($DNSServer, $Lab1Pos1, $Lab1Pos2)
	$Lab1Pos2 = $Lab1Pos2 + 20
	$label5 = GUICtrlCreateLabel($WINSServer, $Lab1Pos1, $Lab1Pos2)
	$Lab1Pos2 = $Lab1Pos2 + 20
	$label6 = GUICtrlCreateLabel($DHCP, $Lab1Pos1, $Lab1Pos2)
	$Lab1Pos2 = $Lab1Pos2 + 20
	$label7 = GUICtrlCreateLabel($DHCPServer, $Lab1Pos1, $Lab1Pos2)
	If $IPaddress = "0.0.0.0" Then GUICtrlSetData($statuslabel, $NoteString2)
EndFunc

Func _Refresh_IP($a)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	For $objItem In $colItems
		If $objItem.caption = $a Then 
			$IPaddress = $objItem.IPAddress(0)
			Return $IPaddress
		EndIf
	Next
EndFunc

Func _Refresh_Subnet($a)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	For $objItem In $colItems
		$Subnet = $objItem.IPSubnet(0)
		If $objItem.caption = $a Then
			Return $Subnet
		EndIf
	Next
EndFunc

Func _Refresh_DefaultGW($a)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	For $objItem In $colItems
		$DefaultGW = $objItem.DefaultIPGateway(0)
		If $objItem.caption = $a Then 
			Return $DefaultGW
		EndIf
	Next
EndFunc

Func _Refresh_DNSServer($a)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
		For $objItem In $colItems
		$DNSServer = $objItem.DNSServerSearchOrder(0)
		If $objItem.caption = $a Then 
			If $DNSServer = "0" Then $DNSServer = "Not Set"
			Return $DNSServer
		EndIf
	Next
EndFunc


Func _Refresh_WINS($a)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
		For $objItem In $colItems
		$WINSServer = $objItem.WINSPrimaryServer
		If $objItem.caption = $a Then 
			If $WINSServer = "0" or $WINSServer = "" Then $WINSServer = "Not Set"
			Return $WINSServer
		EndIf
	Next
EndFunc

Func _Test_DHCP($a)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	For $objItem In $colItems
		$DHCPEnabled = $objItem.DHCPEnabled
		If $objItem.caption = $a Then
			If $DHCPEnabled <> 0 Then
				$DHCPEnabled2 = "Yes"
			Else
				$DHCPEnabled2 = "No"
			EndIf
			Return $DHCPEnabled2
		EndIf
	Next
EndFunc

Func _Refresh_DHCPServer($a)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
		For $objItem In $colItems
		$DHCPServer = $objItem.DHCPServer
		If $objItem.caption = $a Then
			If $DHCPServer = "0" or $DHCPServer = "" Then $DHCPServer = "Not Set"
			Return $DHCPServer
		EndIf
	Next
EndFunc

Func _Get_Adapters()
	GUICtrlSetData($statuslabel, "")
	$Adapters = ""
	SplashTextOn($WinTitle, "Please Wait...", 170, 40)
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	For	$objItem in $colItems
		If $objItem.NetConnectionID <> "" Then $Adapters = $Adapters & "|" & $objItem.NetConnectionID
	Next
	SplashOff()
	If $Adapters = "" Then GUICtrlSetData($statuslabel, $NoteString3)
	Return $Adapters
EndFunc

Func _Exit()
	$ExitMsgBox = MsgBox(3, $WinTitle, "Save Settings?")
	If $ExitMsgBox = 6 Then
		_SaveIniFile()
		Exit
	ElseIf $ExitMsgBox = 7 Then 
		Exit
	Else
		_Set_Gui("Enable")
		Return
	EndIf
EndFunc

Func _SaveIniFile()
	FileDelete($SettingsFile)
	IniWrite($SettingsFile, "Program Options", "Welcome MSG", $WelcomeEnable)
	If $show = @SW_SHOW Then
		IniWrite($SettingsFile, "Program Options", "Debug", "1")
	Else
		IniWrite($SettingsFile, "Program Options", "Debug", "0")
	EndIf
	FileWriteLine($SettingsFile, @CRLF & @CRLF)
	IniWrite($SettingsFile, "Last Used Adapter", "Name", GUICtrlRead($Adapter))
	FileWriteLine($SettingsFile, @CRLF & @CRLF)
	IniWrite($SettingsFile, "IPConfig1", "Name", GUICtrlRead($Config1Name))
	IniWrite($SettingsFile, "IPConfig1", "IPAddress", GUICtrlRead($IPaddressSet1))
	IniWrite($SettingsFile, "IPConfig1", "SubnetMask", GUICtrlRead($SubnetSet1))
	IniWrite($SettingsFile, "IPConfig1", "DefaultGW", GUICtrlRead($DefaultGWSet1))
	IniWrite($SettingsFile, "IPConfig1","DNS", GUICtrlRead($DNSSet1))
	IniWrite($SettingsFile, "IPConfig1","WINS", GUICtrlRead($WINSSet1))
	If GUICtrlRead($DHCPCheck1) = $GUI_CHECKED Then
		IniWrite($SettingsFile, "IPConfig1", "Use DHCP", "1")
	Else
		IniWrite($SettingsFile, "IPConfig1", "Use DHCP", "0")
	EndIf
	FileWriteLine($SettingsFile, @CRLF & @CRLF)
	IniWrite($SettingsFile, "IPConfig2", "Name", GUICtrlRead($Config2Name))
	IniWrite($SettingsFile, "IPConfig2", "IPAddress", GUICtrlRead($IPaddressSet2))
	IniWrite($SettingsFile, "IPConfig2", "SubnetMask", GUICtrlRead($SubnetSet2))
	IniWrite($SettingsFile, "IPConfig2", "DefaultGW", GUICtrlRead($DefaultGWSet2))
	IniWrite($SettingsFile, "IPConfig2","DNS", GUICtrlRead($DNSSet2))
	IniWrite($SettingsFile, "IPConfig2","WINS", GUICtrlRead($WINSSet2))
	If GUICtrlRead($DHCPCheck2) = $GUI_CHECKED Then
		IniWrite($SettingsFile, "IPConfig2", "Use DHCP", "1")
	Else
		IniWrite($SettingsFile, "IPConfig2", "Use DHCP", "0")
	EndIf
	FileWriteLine($SettingsFile, @CRLF & @CRLF)
	IniWrite($SettingsFile, "IPConfig3", "Name", GUICtrlRead($Config3Name))
	IniWrite($SettingsFile, "IPConfig3", "IPAddress", GUICtrlRead($IPaddressSet3))
	IniWrite($SettingsFile, "IPConfig3", "SubnetMask", GUICtrlRead($SubnetSet3))
	IniWrite($SettingsFile, "IPConfig3", "DefaultGW", GUICtrlRead($DefaultGWSet3))
	IniWrite($SettingsFile, "IPConfig3","DNS", GUICtrlRead($DNSSet3))
	IniWrite($SettingsFile, "IPConfig3","WINS", GUICtrlRead($WINSSet3))
	If GUICtrlRead($DHCPCheck3) = $GUI_CHECKED Then
		IniWrite($SettingsFile, "IPConfig3", "Use DHCP", "1")
	Else
		IniWrite($SettingsFile, "IPConfig3", "Use DHCP", "0")
	EndIf
	FileWriteLine($SettingsFile, @CRLF & @CRLF)
	IniWrite($SettingsFile, "IPConfig4", "Name", GUICtrlRead($Config4Name))
	IniWrite($SettingsFile, "IPConfig4", "IPAddress", GUICtrlRead($IPaddressSet4))
	IniWrite($SettingsFile, "IPConfig4", "SubnetMask", GUICtrlRead($SubnetSet4))
	IniWrite($SettingsFile, "IPConfig4", "DefaultGW", GUICtrlRead($DefaultGWSet4))
	IniWrite($SettingsFile, "IPConfig4","DNS", GUICtrlRead($DNSSet4))
	IniWrite($SettingsFile, "IPConfig4","WINS", GUICtrlRead($WINSSet4))
	If GUICtrlRead($DHCPCheck4) = $GUI_CHECKED Then
		IniWrite($SettingsFile, "IPConfig4", "Use DHCP", "1")
	Else
		IniWrite($SettingsFile, "IPConfig4", "Use DHCP", "0")
	EndIf
EndFunc
