#RequireAdmin

#Region
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion

#include <AutoItConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <GuiEdit.au3>

; Begin of Executable File Information
#pragma compile(Out, Automated_Windows_Maintenance_1.4.0.exe)
#pragma compile(Icon, D:\Programming\AutoIT\AWM Project\icon.ico)
#pragma compile(UPX, True)
#pragma compile(FileDescription, 'Automated Windows Maintenance 1.4.0 (Public)')
#pragma compile(ProductName, Automated Windows Maintenance)
#pragma compile(ProductVersion, 1.4.0)
#pragma compile(FileVersion, 1.4.0.0)
#pragma compile(LegalCopyright, � 2020-2025 Jeff Steveanus)
#pragma compile(LegalTrademarks, 'Tiakan oku Sinalau Bakas')
#pragma compile(CompanyName, 'Jeff Steveanus')
; End of Executable File Information

#Region
$MainForm = GUICreate("Automated Windows Maintenance 1.4.0 (Public)", 496, 494, 232, 124)
$GrpSettings = GUICtrlCreateGroup("Settings:", 8, 8, 481, 145)
$RemoveDrives = GUICtrlCreateCheckbox("Remove Non-Present Drives", 16, 72, 153, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$FlushDNS = GUICtrlCreateCheckbox("Flush DNS Resolver Cache", 16, 56, 153, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$CleanSSL = GUICtrlCreateCheckbox("Clean CryptNet SSL Certificate Cache (Vista and up)", 177, 56, 273, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$CleanWindowsLogs = GUICtrlCreateCheckbox("Clean Windows Logs", 16, 104, 121, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$DefragHomeDrive = GUICtrlCreateCheckbox("Defrag Home Drive", 16, 88, 113, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$LblCleaner = GUICtrlCreateLabel("Cleaner Program:", 16, 32, 85, 17)
$LblAction = GUICtrlCreateLabel("Action:", 244, 32, 37, 17)
$ComboCleaner = GUICtrlCreateCombo("", 104, 27, 129, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "Windows Disk Cleanup", "Windows Disk Cleanup")
$ComboAction = GUICtrlCreateCombo("", 284, 27, 193, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "Do Nothing|Close Application|Restart|Shutdown", "Do Nothing")
$DeleteAWMData = GUICtrlCreateCheckbox("Delete AWM (this app) Program data before defragmentation", 177, 120, 305, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$SFC_UTIL = GUICtrlCreateCheckbox("Run System File Checker and fix corrupted Windows files", 177, 73, 305, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$DISM_CHECKRESTOREHEALTH = GUICtrlCreateCheckbox("Run DISM Check and Restore Windows Health", 177, 88, 305, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$CheckDiskUtil = GUICtrlCreateCheckbox("Run Disk Check (chkdsk)", 177, 104, 305, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$GrpLogs = GUICtrlCreateGroup("Logs:", 8, 160, 481, 209)
$TxtCurrentProcess = GUICtrlCreateEdit("", 16, 176, 465, 185, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY, $ES_WANTRETURN, $WS_VSCROLL))
GUICtrlSetData(-1, StringFormat("What this program does:\r\n1. Clean all junk files, browser cookies, cache, and history.\r\n2. Clean CryptNet SSL Certificate Cache.\r\n3. Clean DNS Cache.\r\n4. Clean Logs.\r\n5. Remove Non-Present Drives.\r\n6. Clean Windows Update Cache.\r\n7. Defrag Home Drive.\r\n8. Check and restore Windows health and file integrity.\r\n\r\nWhat's New in Version 1.4.0:\r\n1. [Added] System File Checker\r\n2. [Added] Check and Restore Windows Health\r\n3. [Added] Automatic Check Drive Type (SSD or HDD)\r\n4. [Change] Disable Disk Defragmentation on SSD\r\n5. [Change] Disable Disk Check on SSD\r\n6. [Misc] Update all third-party programs used by this program\r\n7. [Misc] Major Program Update\r\n"))
GUICtrlCreateGroup("", -99, -99, 1, 1)
$BtnJobs = GUICtrlCreateButton("START NOW", 160, 376, 171, 33)
GUICtrlCreateLabel("This is an open-source program", 145, 424, 208, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$SourceCode = GUICtrlCreateLabel("Click here to get source code", 8, 464, 143, 17)
GUICtrlSetFont(-1, 8, 400, 4, "MS Sans Serif")
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlSetCursor (-1, 0)
$LblCopyright = GUICtrlCreateLabel("© 2020 - 2025 Jeff Steveanus", 339, 464, 147, 17)
GUISetState(@SW_SHOW)
#EndRegion

AutoItSetOption("TrayMenuMode", 0)

Const $DriveLetter = @HomeDrive & "\"

; Declare Variables
Global $AppTemp = (@HomeDrive & "\AWM\")
Global $GitHub = ("https://github.com/jeffsteveanus/Automated-Windows-Maintenance/")

; Detect drive type and disable defrag/disk check for SSD
$type = DriveGetType($DriveLetter)
If @error Then
    MsgBox(4096, "Error", "Invalid drive requested")
    Exit
EndIf

$ssd = DriveGetType($DriveLetter, $DT_SSDSTATUS)

Func CheckDriveType()
	If $ssd = "SSD" Then
	GUICtrlSetState($DefragHomeDrive, $GUI_DISABLE)
	GUICtrlSetState($DefragHomeDrive, $GUI_UNCHECKED)
	GUICtrlSetState($CheckDiskUtil, $GUI_DISABLE)
	GUICtrlSetState($CheckDiskUtil, $GUI_UNCHECKED)
	EndIf
EndFunc

CheckDriveType()

Func Disable_UI()
	  GUICtrlSetState($BtnJobs, $GUI_DISABLE)
	  GUICtrlSetState($FlushDNS, $GUI_DISABLE)
	  GUICtrlSetState($RemoveDrives, $GUI_DISABLE)
	  GUICtrlSetState($CleanSSL, $GUI_DISABLE)
	  GUICtrlSetState($DefragHomeDrive, $GUI_DISABLE)
	  GUICtrlSetState($CleanWindowsLogs, $GUI_DISABLE)
	  GUICtrlSetState($DeleteAWMData, $GUI_DISABLE)
	  GUICtrlSetState($ComboCleaner, $GUI_DISABLE)
	  GUICtrlSetState($SFC_UTIL, $GUI_DISABLE)
	  GUICtrlSetState($DISM_CHECKRESTOREHEALTH, $GUI_DISABLE)
EndFunc

Func Enable_UI()
	  GUICtrlSetState($BtnJobs, $GUI_ENABLE)
	  GUICtrlSetState($FlushDNS, $GUI_ENABLE)
	  GUICtrlSetState($RemoveDrives, $GUI_ENABLE)
	  GUICtrlSetState($CleanSSL, $GUI_ENABLE)
	  GUICtrlSetState($DefragHomeDrive, $GUI_ENABLE)
	  GUICtrlSetState($CleanWindowsLogs, $GUI_ENABLE)
	  GUICtrlSetState($DeleteAWMData, $GUI_ENABLE)
	  GUICtrlSetState($ComboCleaner, $GUI_ENABLE)
	  GUICtrlSetState($SFC_UTIL, $GUI_ENABLE)
	  GUICtrlSetState($DISM_CHECKRESTOREHEALTH, $GUI_ENABLE)
	  CheckDriveType()
EndFunc


; Main loop to handle events
While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
		Case $GUI_EVENT_CLOSE
		DirRemove($AppTemp, 1)
		Exit

        Case $BtnJobs
            Disable_UI()
            GUICtrlSetData($TxtCurrentProcess, "")
            Extract()
            If Not FileExists($AppTemp) Then
                MsgBox($MB_ICONWARNING, "Information", "Required files not found. Please restart this application and try again.", 120)
                Enable_UI()
                GUICtrlSetData($GrpLogs, "Logs:")
            ElseIf FileExists($AppTemp) Then
                ; Stop Windows Update Service
                _GUICtrlEdit_AppendText($TxtCurrentProcess, "Stopping Windows Update Service...." & @CRLF)
                GUICtrlSetData($GrpLogs, "Current Process: Stopping Windows Update Service...")
                RunWait(@ComSpec & " /c " & "net stop WUAUSERV & exit", @SW_HIDE)

                If GUICtrlRead($FlushDNS) == $GUI_CHECKED Then
                    FlushDNSCache()
                EndIf

                If GUICtrlRead($CleanSSL) == $GUI_CHECKED Then
                    CleanSSLCache()
                EndIf

                If GUICtrlRead($RemoveDrives) == $GUI_CHECKED Then
                    DriveCleanups()
                EndIf

                CleanUp()
                Delete_LeftOver()

                If GUICtrlRead($CleanWindowsLogs) == $GUI_CHECKED Then
                    ClearEventLogs()
                EndIf

                If GUICtrlRead($DeleteAWMData) == $GUI_CHECKED Then
                    DirRemove($AppTemp, 1)
                EndIf

                If GUICtrlRead($CheckDiskUtil) == $GUI_CHECKED Then
                    Disk_Check()
                EndIf

                If GUICtrlRead($DISM_CHECKRESTOREHEALTH) == $GUI_CHECKED Then
                    DISM_CHECK_RESTORE_HEALTH()
                EndIf

                If GUICtrlRead($SFC_UTIL) == $GUI_CHECKED Then
                    SFC_Run()
                EndIf

                If GUICtrlRead($DefragHomeDrive) == $GUI_CHECKED Then
                    Defrag()
                EndIf

                ; Start Windows Update Service
                Local $TakeAction = GUICtrlRead($ComboAction)
                If $TakeAction = "Do Nothing" Then
                    _GUICtrlEdit_AppendText($TxtCurrentProcess, "Starting Windows Update Service...." & @CRLF)
                    GUICtrlSetData($GrpLogs, "Current Process: Starting Windows Update Service...")
                    RunWait(@ComSpec & " /c " & "net start WUAUSERV & exit", @SW_HIDE)
                    _GUICtrlEdit_AppendText($TxtCurrentProcess, "Done" & @CRLF)
                    _GUICtrlEdit_AppendText($TxtCurrentProcess, "===========================================================" & @CRLF)
                EndIf

                DirRemove($AppTemp, 1)

                MsgBox($MB_SYSTEMMODAL, "Information", "Maintenance done. Do not forget to share this software with your family and friends if you find it useful.", 120)

                Enable_UI()
                GUICtrlSetData($GrpLogs, "Logs:")

                Local $Action = GUICtrlRead($ComboAction)
                If $Action = "Close Application" Then
                    Exit
                ElseIf $Action = "Shutdown" Then
                    Shutdown(1)
                    Exit
                ElseIf $Action = "Restart" Then
                    Shutdown(2)
                    Exit
                EndIf
            EndIf

        Case $SourceCode
            ShellExecute($GitHub)
    EndSwitch
WEnd

;Extract Files to Working Directory
Func Extract()
   DirCreate($AppTemp)
   GUICtrlSetData($GrpLogs, "Current Process: Extracting Program Files...")
   FileInstall("D:\Programming\AutoIT\AWM Project\AWM_ProgramData\7zxa.dll", $AppTemp & "7zxa.dll", 1)
   FileInstall("D:\Programming\AutoIT\AWM Project\AWM_ProgramData\7za.dll", $AppTemp & "7za.dll", 1)
   FileInstall("D:\Programming\AutoIT\AWM Project\AWM_ProgramData\7za.exe", $AppTemp & "7za.exe", 1)
   FileInstall("D:\Programming\AutoIT\AWM Project\AWM_ProgramData\DriveCleanup.7z", $AppTemp & "DriveCleanup.7z", 1)

   GUICtrlSetData($GrpLogs, "Current Process: Extracting DriveCleanup Utility...")
   RunWait(@ComSpec & " /c " & $AppTemp & "7za.exe" & " x -y " & $AppTemp & "DriveCleanup.7z -o" & $AppTemp, "", @SW_HIDE)
EndFunc

Func CleanUp()
   Local $CleanerProg = GUICtrlRead($ComboCleaner)

   _GUICtrlEdit_AppendText($TxtCurrentProcess, "Cleaning Windows Update...." & @CRLF)
   GUICtrlSetData($GrpLogs, "Current Process: Cleaning Windows Update...")
   RunWait(@ComSpec & " /c " & "rmdir /s /q %windir%\softwaredistribution & exit","", @SW_HIDE)

   _GUICtrlEdit_AppendText($TxtCurrentProcess, "Cleaning Internet Explorer...." & @CRLF)
   GUICtrlSetData($GrpLogs, "Current Process: Cleaning Internet Explorer...")
   RunWait(@ComSpec & " /c " & "rundll32.exe inetcpl.cpl,ClearMyTracksByProcess 4351 & exit","", @SW_HIDE)

	  _GUICtrlEdit_AppendText($TxtCurrentProcess,"Running Windows Disk CleanUp...." & @CRLF)
	  GUICtrlSetData($GrpLogs, "Current Process: Cleaning using Windows CleanUp...")

	;Fix Windows Disk Cleanup from cleaning Downloads Folder
	;Save Windows Disk Cleanup Preset
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Downloaded Program Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Internet Cache Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Memory Dump Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Old ChkDsk Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Recycle Bin", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Service Pack Cleanup", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Setup Log Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error memory dump files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error minidump files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Setup Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Thumbnail Cache", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Upgrade Discarded Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Archive Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Queue Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Archive Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting System Queue Files", "StateFlags1997", "REG_DWORD", "0x2")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files", "StateFlags1997", "REG_DWORD", "0x2")


	RunWait(@ComSpec & " /c " & 'cleanmgr.exe /sagerun:1997 & exit', "", @SW_HIDE)

	_GUICtrlEdit_AppendText($TxtCurrentProcess,"Windows Disk Cleanup successfully cleaned your computer.." & @CRLF)
EndFunc

Func SystemFileChecker()
	   RunWait(@ComSpec & " /c " & 'sfc /scannow', "", @SW_HIDE)
EndFunc

Func Delete_LeftOver()
   GUICtrlSetData($GrpLogs, "Current Process: Deleting Leftover files...")
   _GUICtrlEdit_AppendText($TxtCurrentProcess,"Deleting Leftover files..." & @CRLF)

   RunWait(@ComSpec & " /c " & 'DEL /F /Q /S %temp%', "", @SW_HIDE)
   RunWait(@ComSpec & " /c " & 'FOR /D  %i IN (%temp%) DO RD /S /Q "%i"', "", @SW_HIDE)
   RunWait(@ComSpec & " /c " & 'DEL /F /Q /S %systemroot%\Prefetch\', "", @SW_HIDE)
   RunWait(@ComSpec & " /c " & 'FOR /D  %i IN (%systemroot%\Prefetch\) DO RD /S /Q "%i"', "", @SW_HIDE)
   RunWait(@ComSpec & " /c " & 'DEL /F /Q /S %systemroot%\TEMP\', "", @SW_HIDE)
   RunWait(@ComSpec & " /c " & 'FOR /D  %i IN (%systemroot%\TEMP\) DO RD /S /Q "%i"',"", @SW_HIDE)
EndFunc

Func Defrag()
   _GUICtrlEdit_AppendText($TxtCurrentProcess, "Removing File(s) and Folder(s) created by this software...." & @CRLF)
   GUICtrlSetData($GrpLogs, "Current Process: Removing Program Files...")
   DirRemove($AppTemp, 1)

   _GUICtrlEdit_AppendText($TxtCurrentProcess,"===========================================================" & @CRLF)
   _GUICtrlEdit_AppendText($TxtCurrentProcess, "Please wait until the defragmentation process finishes..."  & @CRLF  & "WARNING: DO NOT CLOSE THE DEFRAGMENTATION WINDOW!" & @CRLF)
   GUICtrlSetData($GrpLogs, "Current Process: Defragmentation...")

		 If StringLen($DriveLetter) <> 1 Then
        MsgBox(16, "Error", "Invalid drive letter")
        Return
    EndIf

    Local $command = "defrag " & $DriveLetter & ": /H /U"

    ShellExecute("cmd.exe", "/c " & $command, "", "runas")
EndFunc


Func FlushDNSCache()
   _GUICtrlEdit_AppendText($TxtCurrentProcess, "Flushing DNS Cache...." & @CRLF)
   GUICtrlSetData($GrpLogs, "Current Process: Flushing DNS Cache...")
   RunWait(@ComSpec & " /c " & "ipconfig /flushdns & exit", "", @SW_HIDE)
EndFunc

Func DriveCleanups()
   _GUICtrlEdit_AppendText($TxtCurrentProcess, "Cleaning Non-Present Drives...." & @CRLF)
   GUICtrlSetData($GrpLogs, "Current Process: Cleaning Non-Present Drives...")
   if @CPUArch = "X86" Then
	  RunWait(@ComSpec & " /c " & $AppTemp & "DriveCleanup\x86\DriveCleanup.exe -n & exit", "", @SW_HIDE)
   ElseIf @CPUArch = "X64" Then
	  RunWait(@ComSpec & " /c " & $AppTemp & "DriveCleanup\x64\DriveCleanup.exe -n & exit", "", @SW_HIDE)
   EndIf
EndFunc

Func CleanSSLCache()
   _GUICtrlEdit_AppendText($TxtCurrentProcess, "Cleaning CryptNet SSL certificate cache...." & @CRLF)
   GUICtrlSetData($GrpLogs, "Current Process: Cleaning CryptNet SSL certificate cache...")
   RunWait(@ComSpec & " /c " & 'certutil -URLcache * delete' & ' & exit', "", @SW_HIDE)
EndFunc

Func Disk_Check()
    _GUICtrlEdit_AppendText($TxtCurrentProcess, "Checking Disk..." & @CRLF)
    GUICtrlSetData($GrpLogs, "Current Process: Checking Disk...")
    RunWait(@ComSpec & " /c " & "CHKDSK " & @HomeDrive & " /F & exit")
EndFunc

Func SFC_RUN()
    _GUICtrlEdit_AppendText($TxtCurrentProcess, "Checking filesystem files..." & @CRLF)
    GUICtrlSetData($GrpLogs, "Current Process: Checking filesystem files...")
    RunWait(@ComSpec & " /c " & "sfc /scannow & exit")
EndFunc

Func DISM_CHECK_RESTORE_HEALTH()
    _GUICtrlEdit_AppendText($TxtCurrentProcess, "Running DISM CheckHealth utility..." & @CRLF)
    GUICtrlSetData($GrpLogs, "Current Process: Running DISM CheckHealth utility...")
    RunWait(@ComSpec & " /c " & "DISM /Online /Cleanup-Image /CheckHealth & exit")

    _GUICtrlEdit_AppendText($TxtCurrentProcess, "Running DISM ScanHealth utility..." & @CRLF)
    GUICtrlSetData($GrpLogs, "Current Process: Running DISM ScanHealth utility...")
    RunWait(@ComSpec & " /c " & "DISM /Online /Cleanup-Image /ScanHealth & exit")

    _GUICtrlEdit_AppendText($TxtCurrentProcess, "Running DISM RestoreHealth utility..." & @CRLF)
    GUICtrlSetData($GrpLogs, "Current Process: Running DISM RestoreHealth utility...")
    RunWait(@ComSpec & " /c " & "DISM /Online /Cleanup-Image /RestoreHealth & exit")
EndFunc


Func ClearEventLogs()
    Local $result = Run(@ComSpec & " /c wevtutil.exe el", @SystemDir, @SW_HIDE, 2) ; Use 2 to redirect stdout
    Local $logList = ""

    While True
        $logList &= StdoutRead($result)
        If @error Then ExitLoop
    WEnd

    If $logList = "" Then
        MsgBox($MB_ICONERROR, "Error", "Failed to list event logs.")
        Exit
    EndIf

    Local $logArray = StringSplit($logList, @CRLF)

	_GUICtrlEdit_AppendText($TxtCurrentProcess, "Cleaning log..."  & @CRLF)

	For $i = 1 To $logArray[0]
        If $logArray[$i] <> "" Then
			GUICtrlSetData($GrpLogs, "Current Process: Cleaning " & $logArray[$i] & " log")

            RunWait(@ComSpec & " /c wevtutil.exe cl " & $logArray[$i], @SystemDir, @SW_HIDE)
        EndIf
    Next
EndFunc

