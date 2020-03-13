'* ***********************************************
'* Имя:: ARM_SpecInfo.vbs                        *
'* Язык: VBScript                                *
'* Описание: Вывод информации о спецификации АРМ *
'* Версия: 1.1.6                                 *
'* Автор: Кононов А.Ю.      	      24.06.2015 *
'*************************************************

'On Error Resume Next
strComputer = inputbox ("Введите имя компьютера или (IP-адрес):", "Спецификация компьютера",".")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
' ------------------ Сетевая информация (IP-адрес версии IPv4) -----------------------------------------------------------------
Set IPConfigSet = objWMIService.ExecQuery ("Select IPAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = true",,48)
For Each IPConfig in IPconfigSet
	if InStr(IPConfig.IPAddress(0),PPTPNework) > 0 then 
	IP_Address=IPConfig.IPAddress(0)
	end if
Next

sTime = Now
'Set objExplorer = CreateObject("InternetExplorer.Application")
Set objExplorer = WScript.CreateObject("InternetExplorer.Application", "IE_")
With objExplorer
    .Navigate "about:Blank"
    .Toolbar = 1
    .StatusBar = 0
    .Width = 800
    .Height = 800
    .Left = 75
    .Top = 0
    .Visible = 1

End With
Set fileOutput = objExplorer.Document

'This is the code for the web page to be displayed.
Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem",,48)
For Each objItem in colItems

fileOutput.Open
fileOutput.WriteLn "<html>"
fileOutput.WriteLn "    <head>"
fileOutput.WriteLn "        <title>" & objItem.Name & " </title>"
fileOutput.WriteLn "    </head>"
fileOutput.WriteLn "<style type='text/css'>"
fileOutput.WriteLn "   body    { font-size:80%; font-family:MS Shell Dlg; }"
fileOutput.WriteLn "   table   { font-size:90% }"
fileOutput.WriteLn "</style>"
fileOutput.WriteLn "        <center>"                    	
fileOutput.WriteLn "                 <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"
fileOutput.WriteLn "                       <TR style='font-size: 10pt'><TD style='color: #FFFFFF' width='100%' align=center bgcolor='#A0BACB' colspan='2'><b>ARM Special Information.</b></TD></TR>"
'----------------- This (start) current DateTime of the PC being scanned. ------------------------
fileOutput.WriteLn "                       <TR><TD bgcolor='#FEF7D6' colspan='2'>Starting Time: " & sTime & "</TD></TR>"
'----------------- PCName, UserName ---------------------------------------------------------
fileOutput.WriteLn "                           <tr><td colspan='2' bgcolor='#f0f0f0'>"
fileOutput.WriteLn "                               <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"

'Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem",,48)
'For Each objItem in colItems
'fileOutput.WriteLn " <TR><TD width='30%' align=center bgcolor='#ffffff'>IP Address: <b>" & objItem.Name & "</b></td><TD width='40%' align=center bgcolor='#ffffff'>Account: <b>" & objItem.UserName & "</b><TD width='40%' align=center bgcolor='#ffffff'>PC Name: <b>" & IP_Address & "</b></td><tr>"
fileOutput.WriteLn " <TR><TD width='30%' align=center bgcolor='#ffffff'><b>" & objItem.Name & "." & objItem.Domain & "</b></td><TD width='40%' align=center bgcolor='#ffffff'>Login User: <b>  " & objItem.UserName & "</b><TD width='40%' align=center bgcolor='#ffffff'><b>" & IP_Address & "</b></td><tr>"
fileOutput.WriteLn "                                            </table>"
Next

'----------------- Operating System Information -----------------------------------------------------
Set colItems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
fileOutput.WriteLn "                       <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Operating System: </TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Caption & "   (" & objItem.CSDVersion & ")</td></tr>"
Next

'----------------- License key -----------------------------------------------------
Set WshShell = CreateObject("WScript.Shell")
regKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
DigitalProductId = WshShell.RegRead(regKey & "DigitalProductId")
 
Win8ProductName = "Windows Product Name: " & WshShell.RegRead(regKey & "ProductName") & vbNewLine
Win8ProductID = "Windows Product ID: " & WshShell.RegRead(regKey & "ProductID") & vbNewLine
Win8ProductKey = ConvertToKey(DigitalProductId)
strProductKey ="Windows Key: " & Win8ProductKey 
Win8ProductID = Win8ProductName & Win8ProductID & strProductKey 

fileOutput.WriteLn "                       <TR><TD width='30%' align=left bgcolor='#e0e0e0'>License key: </TD><td width='70%' bgcolor=#f0f0f0 align=left>" & Win8ProductKey & "</td></tr>"

 
'MsgBox(Win8ProductKey)
'MsgBox(Win8ProductID)
 
Function ConvertToKey(regKey)
    Const KeyOffset = 52
    isWin8 = (regKey(66) \ 6) And 1
    regKey(66) = (regKey(66) And &HF7) Or ((isWin8 And 2) * 4)
    j = 24
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    Do
        Cur = 0
        y = 14
        Do
            Cur = Cur * 256
            Cur = regKey(y + KeyOffset) + Cur
            regKey(y + KeyOffset) = (Cur \ 24)
            Cur = Cur Mod 24
            y = y -1
        Loop While y >= 0
        j = j -1
        winKeyOutput = Mid(Chars, Cur + 1, 1) & winKeyOutput
        Last = Cur
    Loop While j >= 0
    If (isWin8 = 1) Then
        keypart1 = Mid(winKeyOutput, 2, Last)
        insert = "N"
        winKeyOutput = Replace(winKeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
        If Last = 0 Then winKeyOutput = insert & winKeyOutput
    End If
    a = Mid(winKeyOutput, 1, 5)
    b = Mid(winKeyOutput, 6, 5)
    c = Mid(winKeyOutput, 11, 5)
    d = Mid(winKeyOutput, 16, 5)
    e = Mid(winKeyOutput, 21, 5)
    ConvertToKey = a & "-" & b & "-" & c & "-" & d & "-" & e
End Function

'----------------- Local User account------------------------------
Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_UserAccount WHERE LocalAccount = True and Disabled = false",,48) 
For Each objItem in colItems 
'    LocalUser = objItem.Name
fileOutput.WriteLn "                       <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Local User: </TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Name & "</td></tr>"
Next

'------------------ OEM Information -----------------------------------------
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>General OEM Information</b></TD></TR>"
Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystemProduct",,48)
For Each objItem in colItems
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Manufacturer:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Vendor & "</TD></TR>"
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Model:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Name & "</TD></TR>"
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Serial Number:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.IdentifyingNumber & "</TD></TR>"
Next
'------------------ BIOS version -----------------------------------------
Set colItems = objWMIService.ExecQuery ("Select * from Win32_BIOS",,48)
For Each objItem in colItems
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>BIOS Version:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.SMBIOSBIOSVersion & "</TD></TR>"
Next


'------------------ Hardware Information -----------------------------------------
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>General Hardware Information</b></TD></TR>"
Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_Processor",,48)
For Each objItem in colItems
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'><b>Central Processor:</b></TD><td width='70%' bgcolor=#ffffff align=left>" & objItem.Name & "</td></tr>"
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Current Frequency:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.CurrentClockSpeed & " (MHz)</td></tr>"
Next

Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_CacheMemory",,48)
For Each objItem in colItems
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>" & objItem.Purpose & "</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.InstalledSize & " (Kb)</td></tr>"
Next


Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem",,48)
For Each objItem in colItems
fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'><b>Physical Memory:</b></TD><td width='70%' bgcolor=#ffffff align=left>" & Round(((objItem.TotalPhysicalMemory/1024)/1024)/1024,2) & " GB</td></tr>"
Next

Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_PhysicalMemory",,48) 
For Each objItem in colItems 
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Slot: " & objItem.DeviceLocator & "</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & int((objItem.Capacity/1024)/1024) & " MB ( " & objItem.Speed & " MHz )</td></tr>"
Next


'--------- Current HDD Space Information -------------------- 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive",,48) 
For Each objItem in colItems
	If objItem.MediaType = "Removable Media" Then
		fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Removable Dsik:</TD><td width='70%' bgcolor=#f0d0d0 align=left>" & objItem.Model & "</td></tr>"
	Else
		fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'>HDD Model Drive:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Model & "</td></tr>"

	End If
Next

Set colItems = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk Where Name='C:'",,48)
For Each objItem in colItems
fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Total HDD Space:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & Fix(((objItem.Size/1024)/1024)/1024) & " GB ( Free: " & Fix(((objItem.FreeSpace/1024)/1024)/1024) & " GB )</td></tr>"
Next

'--------- Current Video Information -------------------- 
Set colItems = objWMIService.ExecQuery ("Select * from Win32_VideoController WHERE DeviceID = 'VideoController1'",,48)
For Each objItem in colItems
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Video Card:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Caption & "</td></tr>"
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Video Processor:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.VideoProcessor & "</td></tr>"
fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Current Resolution:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.CurrentHorizontalResolution & " x " & objItem.CurrentVerticalResolution & " ( " & objItem.CurrentBitsPerPixel  & " bits  )</td></tr>"
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Video Adapter RAM:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & Fix ((objItem.AdapterRAM/1024)/1024) & " MB</td></tr>"
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Video Driver Version:</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.DriverVersion & "</td></tr>"
Next

'------------- USB Device Information --------------
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>USB Device Information</b></TD></TR>"
Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_PnPEntity WHERE Description like 'EPSON%' or Description like 'Xerox%' or  Description like 'Moto%' or Description = 'VPN Key' or Description = 'FB1200' or Description like 'Lexmark%' or Description like 'Xerox%' or Description like '%PIN%' or Description = 'iLook 300' or Description like 'Eye' or Description = 'HD Webcam C270' or Description = '1200dpi Scanner' or Description like 'iButton%'",,48)
	For Each objItem in colItems 
	fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'>" & objItem.Manufacturer & "</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Name & "</td></tr>"
        Next
'------------- BATTERY Information --------------
Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_PnPSignedDriver WHERE DeviceClass = 'BATTERY' and Description <> 'Составная батарея (Майкрософт)'",,48)
	For Each objItem in colItems 
	BatUPS = objItem.DeviceClass
        Next

Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_Battery",,48) 
	For Each objItem in colItems 
If objItem.Name <> "0" Then
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>" & BatUPS & "</b> ( " & objItem.DeviceID & " )</TD></TR>"
	fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'> Estimated Charge: </TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.EstimatedChargeRemaining & "%</td></tr>"		
	fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'> Estimated RunTime: </TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.EstimatedRunTime & "</td></tr>"		
End If
	Next
'------------- Input Device Information --------------
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>Input Device Information</b></TD></TR>"

Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_PnPSignedDriver WHERE DeviceClass = 'KEYBOARD' or DeviceClass = 'MOUSE' ",,48)
	For Each objItem in colItems 
	fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'>" & objItem.DeviceClass & "</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & objItem.Description & "</td></tr>"	
	Next
'------------------ Local Printers Information---------------
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>Local Printers Information </b></TD></TR>"
'Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_Printer where Local=True",,48)
Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_Printer",,48)
For Each objItem in colItems
		If objItem.Name <> "Отправить в OneNote 2010" then
			If objItem.Name <> "Microsoft XPS Document Writer" then
				If objItem.Name <> "Fax" then
fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>" & objItem.Name & " <TD width='70%' align=left bgcolor='#f0f0f0'>" & objItem.PortName & "</td></tr>"
				End If
			End If
		End If
Next

' ------------------ Serial Ports ---------------------------------
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>Serial Ports</b></TD></TR>"
Set colItems = objWMIService.ExecQuery ("SELECT * FROM Win32_SerialPort",,48) 
For Each objItem in colItems 
         If objItem.Name <> "0" Then
		If objItem.StatusInfo = "3" Then
		fileOutput.WriteLn " <TR><TD width='35%' align=left bgcolor='#e0e0e0'>" & objItem.Name & " <TD width='60%' align=left bgcolor='#f0f0f0'>used</td></tr>"
		Else
		fileOutput.WriteLn " <TR><TD width='35%' align=left bgcolor='#e0e0e0'>" & objItem.Name & " <TD width='60%' align=left bgcolor='#f0f0f0'>" & objItem.StatusInfo & "</td></tr>"
		End If
	End If
Next


' ------------------ Mapped Network Disk  (Permission: Administrators) ---------------------------------
fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>Mapped Network Disk</b><i>   ____ (for Local Administrators)</i></TD></TR>"
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_MappedLogicalDisk",,48) 
For Each objItem in colItems 
fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#e0e0e0'>Mapped Disk (" & objItem.Name & ") <TD width='70%' align=left bgcolor='#f0f0f0'>" & objItem.ProviderName & "</td></tr>"
Next

'------------------ HWD Process.-------------------------------------------------------------
'fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>HWD Process</b></TD></TR>"
'fileOutput.WriteLn "                                        <tr><td colspan='2' bgcolor='#f0f0f0'>"
'fileOutput.WriteLn "                                            <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"

'Set colItem = objWMIService.ExecQuery ("SELECT * FROM Win32_Process Where Name = 'hwd.exe' ") 
' Проверяем запущен ли процесс HWD.exe
'	If colItem.Count = 0 Then
'		fileOutput.WriteLn " <TR><TD width='40%' align=center bgcolor='#f0f0f0'>Daemon HWD</TD><td bgcolor=#f0d0d0 align=center>Not Running !!!</td></tr>"'
'		Else
'		For Each objItem in colItem
'		objItem.GetOwner User
'fileOutput.WriteLn " <TR><TD width='30%' align=center bgcolor='#c0c0c0'><b>Executable Path</b></td><TD width='40%' align=center bgcolor='#c0c0c0'><b>Process User</b><TD width='40%' align=center bgcolor='#c0c0c0'><b>Process State</b></td><tr>"
'fileOutput.WriteLn " <TR><TD width='30%' align=left bgcolor='#ffffff'>" & objItem.ExecutablePath & "</td><TD width='40%' align=center bgcolor='d0f0d0'>" & User & "<TD width='40%' align=center bgcolor='d0f0d0'>Running</td><tr>"
'			Next
'	End If
'fileOutput.WriteLn "                                            </table>"

'------------------ Amicon FPSU --------------------------------------------------------
'fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>Amicon Information</b></TD></TR>"
'fileOutput.WriteLn "                                        <tr><td colspan='2' bgcolor='#f0f0f0'>"
'fileOutput.WriteLn "                                            <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"

' --------------------------------------------
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const aKeyPath = "SOFTWARE\Amicon\Client FPSU-IP"

'Const ParamVhi = "VersionHi"
'Const ParamVlo = "VersionLo"
'Const ParamVbn = "VersionBn"
'Const ParamVs = "VpnKey_Serial"
'Const ParamVin = "VpnKey_Inserted"

'Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
' Чтение HEX - параметра реестра -------------------------
'objReg.GetDWORDValue HKEY_LOCAL_MACHINE, aKeyPath, ParamVhi, HiVal
'objReg.GetDWORDValue HKEY_LOCAL_MACHINE, aKeyPath, ParamVlo, LoVal
'objReg.GetDWORDValue HKEY_LOCAL_MACHINE, aKeyPath, ParamVbn, BnVal
'objReg.GetStringValue HKEY_LOCAL_MACHINE, aKeyPath, ParamVs, psVal
'objReg.GetDWORDValue HKEY_LOCAL_MACHINE, aKeyPath, ParamVin, InVal
'If HiVal <> "0" then
'fileOutput.WriteLn " <TR><TD width='30%' align=center bgcolor='#c0c0c0'><b>Component</b></td><TD width='40%' align=center bgcolor='#c0c0c0'><b>Serial Number:</b><TD width='40%' align=center bgcolor='#c0c0c0'><b>Product Version</b></td><tr>"
'	If psVal <> "0" Then
'		fileOutput.WriteLn " <TR><TD width='30%' align=center bgcolor='#f0f0f0'>Amicon FPSU-IP/Client</td><TD width='40%' align=center bgcolor='#d0f0d0'>" & psVal & "<TD width='40%' align=center bgcolor='d0f0d0'>" & HiVal & "." & LoVal & "." & BnVal & "</td><tr>"
'	Else
'		fileOutput.WriteLn " <TR><TD width='30%' align=center bgcolor='#f0f0f0'>Amicon FPSU-IP/Client</td><TD width='40%' align=center bgcolor='#f0d0c0'>" & psVal & "<TD width='40%' align=center bgcolor='f0d0c0'>" & HiVal & "." & LoVal & "." & BnVal & "</td><tr>"
'
'        End If
'Else
'		fileOutput.WriteLn " <TR><TD width='40%' align=center bgcolor='#f0f0f0'>Amicon FPSU-IP/Client:</TD><td bgcolor=#f0f0f0 align=center>Not Used</td></tr>"'


'End If
'fileOutput.WriteLn "                                            </table>"

'------------------ InfoCrypt Information --------------------------------------------------------
'fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>InfoCrypt Information</b></TD></TR>"
'fileOutput.WriteLn "                                        <tr><td colspan='2' bgcolor='#f0f0f0'>"
'fileOutput.WriteLn "                                            <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"
'fileOutput.WriteLn " <TR><TD width='20%' align=center bgcolor='#c0c0c0'><b>Component</b></td><TD width='40%' align=center bgcolor='#c0c0c0'><b>Executable Path</b><TD width='40%' align=center bgcolor='#c0c0c0'><b>Version</b></td><tr>"
' ----------------------WinNet----------------------

'Const tKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WinNet 3.0"
'Const sKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Sbersign50"
'Const iadKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\InfoCryptAdmin"

'Const Param1 = "UninstallString"
'Const Param2 = "DisplayName"

' Чтение Строкового параметра реестра -------------------------

'objReg.GetStringValue HKEY_LOCAL_MACHINE, tKeyPath, Param1, lVal
'objReg.GetStringValue HKEY_LOCAL_MACHINE, tKeyPath, Param2, pVal
'If pVal <> "0" then
'fileOutput.WriteLn " <TR><TD width='20%' align=center bgcolor='#f0f0f0'>WinNet</td><TD width='50%' align=left bgcolor='#ffffff'>" & lVal & "<TD width='40%' align=center bgcolor='#ffffff'>" & pVal & "</td><tr>"
'End If
'objReg.GetStringValue HKEY_LOCAL_MACHINE, sKeyPath, Param1, lVal
'objReg.GetStringValue HKEY_LOCAL_MACHINE, sKeyPath, Param2, pVal
'If pVal <> "0" then
'fileOutput.WriteLn " <TR><TD width='20%' align=center bgcolor='#f0f0f0'>SberSign</td><TD width='50%' align=left bgcolor='#ffffff'>" & lVal & "<TD width='40%' align=center bgcolor='#ffffff'>" & pVal & "</td><tr>"
'End If
'objReg.GetStringValue HKEY_LOCAL_MACHINE, iadKeyPath, Param1, lVal
'objReg.GetStringValue HKEY_LOCAL_MACHINE, iadKeyPath, Param2, pVal
'If pVal <> "0" then
'fileOutput.WriteLn " <TR><TD width='20%' align=center bgcolor='#f0f0f0'>InfoCryptAdmin</td><TD width='50%' align=left bgcolor='#d0f0d0'>" & lVal & "<TD width='40%' align=center bgcolor='#d0f0d0'>" & pVal & "</td><tr>"
'Else
'fileOutput.WriteLn " <TR><TD width='20%' align=center bgcolor='#f0f0f0'>InfoCryptAdmin</td><TD width='50%' align=left bgcolor='#f0d0d0'>" & lVal & "<TD width='40%' align=center bgcolor='#f0d0d0'>" & pVal & "</td><tr>"

'End If

'fileOutput.WriteLn "                                            </table>"

' ------------------ WinNet Users ---------------------------------
'Const uKeyPath = "SOFTWARE\Infocrypt\WinNet 3.0\Users"


' -------- Чтение параметров ветки в массив arrValues
'intRes = objReg.EnumValues (HKEY_LOCAL_MACHINE, uKeyPath, arrValues)
'If intRes = 0 Then
'fileOutput.WriteLn "                                        <TR><TD bgcolor='#A0BACB' colspan='2'><b>WinNet Users Logon</b></TD></TR>"
' -------- Чтение значений параметров в массиве PassKey
'	For i = Lbound(arrValues) To Ubound(arrValues)
'	intRes = objReg.GetStringValue (HKEY_LOCAL_MACHINE, uKeyPath, arrValues(i), PassKey)
' -------- проверка пустого значения
'		If PassKey <> "" Then
'fileOutput.WriteLn "                                        <TR><TD width='30%' align=left bgcolor='#e0e0e0'>" & arrValues(i) & "</TD><td width='70%' bgcolor=#f0f0f0 align=left>" & PassKey & "</td></tr>"
'		End If
'	Next
'End If

'------------------ Current Services Lists and status -------------------------------------------------------------
'fileOutput.WriteLn "                                        <TR><TD align=center bgcolor='#A0BACB' colspan='2'><b>Current Services Lists</b></TD></TR>"
'fileOutput.WriteLn "                                        <tr><td colspan='2' bgcolor='#f0f0f0'>"
'fileOutput.WriteLn "                                            <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"
'fileOutput.WriteLn "                                                <TR><TD width='70%' align=center bgcolor='#c0c0c0'><b>Service Name</b></td><TD width='30%' align=center bgcolor='#c0c0c0'><b>Service State</b></td><tr>"


'Set colRunningServices = objWMIService.ExecQuery ("Select * from Win32_Service",,48)
'For Each objService in colRunningServices
'	If objService.State = "Running" then'
'	fileOutput.WriteLn "                                <TR><TD align=left bgcolor=#ffffff>" & objService.DisplayName & "</TD><td bgcolor=#ffffff align=center>" & objService.State & "</td></tr>"
'	else
'	fileOutput.WriteLn "                                <TR><TD align=left bgcolor=#e0e0e0>" & objService.DisplayName & "</TD><td bgcolor=#e0e0e0 align=center>" & objService.State & "</td></tr>"
'	End If
'Next
'fileOutput.WriteLn "                                            </table>"

'------------------ Install Programs and Features -------------------------------------------------------------
'fileOutput.WriteLn "                                        <TR><TD align=center bgcolor='#A0BACB' colspan='2'><b>Install Programs and Features</b></TD></TR>"
'fileOutput.WriteLn "                                        <tr><td colspan='2' bgcolor='#f0f0f0'>"
'fileOutput.WriteLn "                                            <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"
'fileOutput.WriteLn " <tr><td width=35% align=center bgcolor='#c0c0c0'><b>Install Location:</b></td><td width=30% align=center bgcolor='#c0c0c0'><b>Name</b></td><td width=15% align=center bgcolor='#c0c0c0'><b>Vendor</b></td><td width=10% align=center bgcolor='#c0c0c0'><b>Version</b></td></tr>"
'fileOutput.WriteLn "                                            </table>"
'
'Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product",,48)

'For Each objSoftware in colSoftware
'fileOutput.WriteLn "                                            <TABLE width='100%' cellspacing='1' cellpadding='2' border='1' bordercolor='#c0c0c0' bordercolordark='#ffffff' bordercolorlight='#c0c0c0'>"
'	If objSoftware.Vendor = "Microsoft Corporation" then
'fileOutput.WriteLn " <tr><td width=35% align=left bgcolor=#ffffff>" & objSoftware.Name & "</td><td width=30% align=left bgcolor=#ffffff>" & objSoftware.InstallLocation & "</td><td width=15% align=left bgcolor=#ffffff>" & objSoftware.Vendor & "</td><td width=10% align=left bgcolor=#ffffff>" & objSoftware.Version & "</td></tr>"
'	Else
'fileOutput.WriteLn " <tr><td width=35% align=left bgcolor=#e0e0e0>" & objSoftware.Name & "</td><td width=30% align=left bgcolor=#e0e0e0>" & objSoftware.InstallLocation & "</td><td width=15% align=left bgcolor=#e0e0e0>" & objSoftware.Vendor & "</td><td width=10% align=left bgcolor=#e0e0e0>" & objSoftware.Version & "</td></tr>"
'	End If
'Next
'fileOutput.WriteLn "                                            </table>"
'------------------------- This (ended) current DateTime of the PC being scanned. ------------------------
eTime = Now
fileOutput.WriteLn "                       <TR><TD bgcolor='#FEF7D6' colspan='2'>Ended Time: " & eTime & "</TD></TR>"



fileOutput.WriteLn "            </table>"
fileOutput.WriteLn "            <p><small>2014, Alexandr Kononov (R)</small></p>"
fileOutput.WriteLn "        </center>"
fileOutput.WriteLn "    </body>"
fileOutput.WriteLn "<html>"
fileOutput.close
WScript.Quit
