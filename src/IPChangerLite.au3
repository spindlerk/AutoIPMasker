#include <GuiConstants.au3>
#comments-start
Global configurations
#comments-end
Opt("TrayMenuMode",1)
Global $version = "Version 1.7"
Global $WinTitle = "Fast IP Changer (" & $version & ") - Timmo - written in AutoIT"
Global $Adapters, $Adapter3
Global $DefaultGW = ""
Global $IPaddress = ""
Global $Subnet = ""
Global $DHCP = ""
Global $Mark = 0
Global $statuslabel = ""
Global $localhost = "localhost"

$check1 = 0
$check2 = 0
$check3 = 0
$check4 = 0

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems1 = ""
$colItems2 = ""

$NoteXPos = 8
$NoteYPos = 257
$NoteString1 = "Note: IP Change Failed, Please Check settings (and cable?)"
$NoteString2 = "Note: No IP detected please check cable/connection"
$NoteString3 = "Note: No adapters detected, please check system config"
$NoteString4 = "IP Changed successfully"
$NoteString5 = "Please make sure you have selected your adapter"
$NoteString6 = "Failed to get IP via DHCP, make sure you are connected to a DHCP server."
$NoteString7 = "Failed to set DNS server (settings might still work)"

$objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
#comments-start
Call getActiveAdapter to get the NIC currently in use. Tries setting the IP to DHCP, if DHCP is not available, tests three static IP settings from ini file.
#comments-end
		$Adapter = getActiveAdapter($localhost)
		$DHCP=_Set_DHCP_Auto($Adapter)
		If $DHCP == 1 Then
		ElseIf $DHCP == 2 Then
			$IP_set=_Set_ip_Auto($Adapter, 192.168.0.100, 255.255.255.0, 192.168.0.1, 192.168.0.1)
				If $IP_set == 1 Then
				ElseIf $IP_set == 2 Then
				$IP_set=_Set_ip_Auto($Adapter, 172.168.0.12, 255.255.0.0, 172.168.0.1, 172.168.0.1)
					If $IP_set == 1 Then
					ElseIf $IP_set == 2 Then
					$IP_set=_Set_ip_Auto($Adapter, 10.1.1.235, 255.0.0.0, 10.1.1.1, 10.10.1.1)
						If $IP_set == 1 Then
						ElseIf $IP_set == 2 Then
						GUICtrlSetData($statuslabel, $NoteString1)
						EndIf
					EndIf
				EndIf
			EndIf
#comments-start
search and find the active network adapter
#comments-end	
Func getActiveAdapter($srv)
	
	Local $Description, $colItems, $colItem, $ping
	If $CmdLine[0] > 0 Then $ip = $CmdLine[1]
	$objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $localhost & "\root\cimv2")
    $ping = Ping($srv)
    If $ping Then

        $colItems = $objWMIService.ExecQuery("SELECT NetConnectionID, NetConnectionStatus FROM Win32_NetworkAdapter", "WQL", 0x30)
        If IsObj($colItems) Then
            For $objItem In $colItems
				If $objItem.NetConnectionStatus == 2 Then
				return $objItem.NetConnectionID
				ExitLoop
				EndIF
            Next
            SetError(0)
    Else
        SetError(1)
        Return "Host not reachable"
    EndIf
EndIF
EndFunc
#comments-start
configure the auto dhcp and set progress bar
#comments-end
Func _Set_DHCP_Auto($name)
	$DHCP = _Test_DHCP($Adapter)
	ProgressOn($WinTitle, "Changing IP Address")
	GUICtrlSetData($statuslabel, "")
	$run = RunWait(@ComSpec & " /c " & 'netsh interface ip set address name="' & $name & '" source=dhcp', "")
	ProgressSet(25)
	$run3 = RunWait(@ComSpec & " /c " & 'netsh int ip set dns "' & $name & '" dhcp', "")
	RunWait(@ComSpec & " /c " & 'netsh interface ip set wins "' & $name & '" dhcp', "")
	ProgressSet(50)
	RunWait(@ComSpec & " /c " & 'ipconfig /release "' & $name & '"', "")
	ProgressSet(75)
	$run2 = RunWait(@ComSpec & " /c " & 'ipconfig /renew "' & $name & '"', "")
	ProgressSet(100)
	Sleep(500)
	ProgressOff()
	If $run = 0 Then
		Return 1
	ElseIf $run2 <> 0 Then
		Return 2
	ElseIf $run3 <> 0 Then
		MsgBox(0, $WinTitle, $NoteString4)
	Else
		MsgBox(0, $WinTitle, $NoteString1)
		GUICtrlSetData($statuslabel, $NoteString1)
	EndIf
EndFunc
#comments-start
Takes in static IP settings and tries to ping default gateway to see if settings will work
#comments-end
Func _Set_ip_Auto($name, $IP, $Subnet, $DefaultGW, $DNS, $WINS)
	ProgressOn($WinTitle, "Changing IP Address")
	GUICtrlSetData($statuslabel, "")
	$run = RunWait(@ComSpec & " /c " & 'netsh interface ip set address name="' & $name & '" static ' & $IP & " " & $Subnet & " " & $DefaultGW & " 1", "")
	ProgressSet(50)
	$rundns = RunWait(@ComSpec & " /c " & 'netsh interface ip set dns name="' & $name & '" static ' & $DNS, "")
	$runwins = RunWait(@ComSpec & " /c " & 'netsh interface ip set wins name="' & $name & '" static ' & $WINS, "")
	ProgressSet(100)
	Sleep(500)
	ProgressOff()
	$ping = Ping($DefaultGW)
	If $ping == 0 Then 
		Return 0 ;Returns 0 if ping of default gateway fails
	Else
		Return 1 ;Returns 1 if ping of default gateway succeeds 
	EndIf
EndFunc
#comments-start
test if DHCP if enabled on the network
#comments-end
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