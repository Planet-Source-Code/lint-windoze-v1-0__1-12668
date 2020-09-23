Attribute VB_Name = "Windoze"
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'•                     Windoze  v1.0
'•                    Written by Lint
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'• Released: November 9th 2000
'• Coded in: Visual Basic 6.0
'• For use with: Windows 95 and 98 (not sure for ME or 2k)
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'• Site: www.pornstarz.org
'• AIM: virii
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'• Final Comments: I am having a horrible fucking day and
'•                 I hope you sons of bitches get some god
'•                 damn pleasure out of this fucking thing
'•                 and you better learn something.
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'• Greets: Blackout, Dee, Imperial, Yz, Bboy, John Dew,
'•         Diamond, Trend, Polar, Mosh, Kemikal, Happy,
'•         Zirc, Oblivic, Justin, Level, Reflex, Friction,
'•         Fury, Iron, Vrml, Aun, Grind, Sacrifice, Sic,
'•         Gore, Fuzzy and that's probably all.
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••


'Public Declare Function
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
(ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" _
(ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
(ByVal szHost As String) As Long
Public Declare Function GetCurrentProcessId _
Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess _
Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess _
Lib "kernel32" (ByVal dwProcessID As Long, _
ByVal dwType As Long) As Long

'Declare Function
Declare Function StretchBlt% Lib "gdi32" (ByVal hdc%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function setsyscolors Lib "user32" Alias "SetSysColors" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function SetComputerName Lib "kernel32" _
Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal _
nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal _
nSize As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Global nid As NOTIFYICONDATA

'Private Declare Function
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib _
"user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function GetVolumeInformation Lib _
"kernel32.dll" Alias "GetVolumeInformationA" (ByVal _
lpRootPathName As String, ByVal lpVolumeNameBuffer As _
String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength _
As Long, lpFileSystemFlags As Long, ByVal _
lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long
Private Declare Function DllGetVersion _
Lib "Shlwapi.dll" _
(dwVersion As DllVersionInfo) As Long

'Private Declare Sub
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

'Declare Sub
Declare Sub keybd_event Lib "user32" _
(ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal FLAGS As Long, ByVal ExtraInfo As Long)

'Const
Const REG_SZ = 1&

'Public Const
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
Public Const MAX_PATH = 260
Public Const HWND_ONTOP = -1
Public Const HWND_NOTOP = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public Const MAX_WSASYSStatus = 128
Public Const MAX_WSADescription = 256
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1
Public Const SRCCOPY = &HCC0020

'Private Const
Private Const SPI_SCREENSAVERRUNNING = 97
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_LEVEL_80386 As Long = 3
Private Const PROCESSOR_LEVEL_80486 As Long = 4
Private Const PROCESSOR_LEVEL_PENTIUM As Long = 5
Private Const PROCESSOR_LEVEL_PENTIUMII As Long = 6

'Global
Global LeftX
Global TopY

'Global Const
Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const WM_MOUSEMOVE = &H200
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202

'Dim
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long

'Type
Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

'Public Type
Public Type WSADATA
wVersion      As Integer
wHighVersion  As Integer
szDescription(0 To MAX_WSADescription)   As Byte
szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
wMaxSockets   As Integer
wMaxUDPDG     As Integer
dwVendorInfo  As Long
End Type

'Private Type
Private Type DllVersionInfo
cbSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
End Type
Private Type OSVERSIONINFOEX
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Public Type udtCPU
lClockSpeed As Variant
lProcType As Integer
strProcLevel As String
strProcRevision As String
lNumberOfProcessors As Long
End Type
Private Type SYSTEM_INFO
dwOemID As Long
dwPageSize As Long
lpMinimumApplicationAddress As Long
lpMaximumApplicationAddress As Long
dwActiveProcessorMask As Long
dwNumberOfProcessors As Long
dwProcessorType As Long
dwAllocationGranularity As Long
wProcessorLevel As Integer
wProcessorRevision As Integer
End Type
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type

Function FlipScreen(frm As Form, pixx As PictureBox)
Dim fliphorizontal As Boolean, flipvertical As Boolean, thechange
If frm.Caption = frm.Caption Then
pixx.AutoRedraw = True
thechange = SRCCOPY
flipvertical = True
DumpToWindow pixx, thechange, fliphorizontal, flipvertical, frm
End If
End Function

Sub FlipPictureHorizontal(pic1 As PictureBox, pic2 As PictureBox)
'don't fuck with this or i will kill your bitch ass
pic1.ScaleMode = 3
pic2.ScaleMode = 3
Dim px%
Dim py%
Dim retval%
px% = pic1.ScaleWidth
py% = pic1.ScaleHeight
retval% = StretchBlt(pic2.hdc, px%, 0, -px%, py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub

Sub FlipPictureVertical(pic1 As PictureBox, pic2 As PictureBox)
'don't.
pic1.ScaleMode = 3
pic2.ScaleMode = 3
Dim px%
Dim py%
Dim retval%
px% = pic1.ScaleWidth
py% = pic1.ScaleHeight
retval% = StretchBlt(pic2.hdc, 0, py%, px%, -py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub

Function Get_CPULevel()
Dim tCPU As udtCPU
Call GetCPUInfo(tCPU)
Get_CPULevel = tCPU.strProcLevel
End Function

Function Get_CPUType()
Dim tCPU As udtCPU
Call GetCPUInfo(tCPU)
Get_CPUType = tCPU.lProcType
End Function

Function Get_ManyCPU()
Dim tCPU As udtCPU
Call GetCPUInfo(tCPU)
Get_ManyCPU = tCPU.lNumberOfProcessors
End Function

Public Function Get_CPUInfo(ptCPUInfo As udtCPU)
Dim tSYS As SYSTEM_INFO
Dim intProcType As Integer
Dim strProcLevel As String
Dim strProcRevision As String
Call GetSystemInfo(tSYS)
Select Case tSYS.dwProcessorType
Case PROCESSOR_INTEL_386: intProcType = 386
Case PROCESSOR_INTEL_486: intProcType = 486
Case PROCESSOR_INTEL_PENTIUM: intProcType = 586
End Select
Select Case tSYS.wProcessorLevel
Case PROCESSOR_LEVEL_80386: strProcLevel = "Intel 80386"
Case PROCESSOR_LEVEL_80486: strProcLevel = "Intel 80486"
Case PROCESSOR_LEVEL_PENTIUM: strProcLevel = "Intel Pentium"
Case PROCESSOR_LEVEL_PENTIUMII: strProcLevel = "Intel Pentium Pro or Pentium II"
End Select
strProcRevision = "Model " & HiByte(tSYS.wProcessorRevision) & ", Stepping " & LoByte(tSYS.wProcessorRevision)
With ptCPUInfo
.lClockSpeed = GetCPUSpeed
.lNumberOfProcessors = tSYS.dwNumberOfProcessors
.lProcType = intProcType
.strProcLevel = IIf(strProcLevel = "", "None", strProcLevel)
.strProcRevision = IIf(strProcRevision = "", "None", strProcRevision)
End With
End Function

Function Get_ComputerOwner()
Get_ComputerOwner = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
End Function

Function Get_KeyboardInfo()
Get_Keyboard = GetStringValue("HKEY_LOCAL_MACHINE\enum\bios\*PNP0303\05", "DeviceDesc")
End Function

Function INI_GetFrom(Section As String, Key As String, Directory As String) As String
Dim Buff As String
Buff = String(750, Chr(0))
Key$ = (Key$)
INI_GetFrom$ = Left(Buff, GetPrivateProfileString(Section$, ByVal Key$, "", Buff, Len(Buff), Directory$))
End Function

Public Function Copy_File(currentFilename As String, newFilename As String)
Dim a%, buffer%, temp$, fRead&, fSize&, b%
On Error GoTo ErrHan:
a = FreeFile
buffer = 4048
Open currentFilename For Binary Access Read As a
b = FreeFile
Open newFilename For Binary Access Write As b
fSize = FileLen(currentFilename)
While fRead < fSize
DoEvents
If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
temp = Space(buffer)
Get a, , temp
Put b, , temp
fRead = fRead + buffer
Wend
Close b
Close a
Copy_File = 1
Exit Function
ErrHan:
Copy_File = 0
End Function

Sub Change_Colors(NUM As Long, COLOR As Long)
'this changes your windows colors
'ex: Call Change_Colors(1, vbBlack)
'desktop - 1
'enabled titlebar [left side] - 2
'disabled titlebar [left side] - 3
'file area(file, edit, search, help) - 4
'TextBox BackColor - 5
'Text COLOR - 8
'Enabled boardercolor - 10
'disabled boardercolor - 11
'highlighted Text - 13
'start button\taskbar\exit buttons\main! - 15
'shadow backcolor.. - 16
'disabled file area (file, edit, search, helt) - 17
'start button\taskbar\exit buttons\main forecolor! - 18
'disabled window caption text - 19
'boarder colors around window functions - 20
'second boarder around functions - 21
'help labels forecolor - 23
'help labels backcolor - 24
'enabled titlebar [right side] - 27
'disabled titlebar [right side] - 28
Dim SetSysColorz
SetSysColorz = a = setsyscolors(1, NUM, COLOR)
End Sub
Function Change_ComputerName(CPN As String)
Call SetComputerName(CPN)
End Function

Function Get_HoursRunning()
Dim lngCount As Long
Dim lngHours As Long
lngCount = GetTickCount
lngHours = ((lngCount / 1000) / 60) / 60
Get_Hours = lngHours
End Function

Function Get_MinutesRunning()
Dim lngCount As Long
Dim lngMinutes As Long
lngCount = GetTickCount
lngMinutes = ((lngCount / 1000) / 60) Mod 60
Get_Minutes = lngMinutes
End Function

Function Get_MonitorInfo()
Get_Monitor = GetStringValue("HKEY_LOCAL_MACHINE\enum\monitor\PHLB15C\PCI_VEN_1002&DEV_4742&SUBSYS_47421002&REV_5C_000800", "devicedesc")
End Function

Function Get_RegCpuIdentification()
Get_RegCPUIDENT = GetStringValue("HKEY_LOCAL_MACHINE\hardware\description\system\centralprocessor\0", "identifier")
End Function

Function Get_RegCpuVendorInfo()
Get_RegCPUVENDOR = GetStringValue("HKEY_LOCAL_MACHINE\hardware\description\system\centralprocessor\0", "VendorIdentifier")
End Function

Public Function Get_SysFolderPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
Get_SysPath = Left(strFolder, InStr(strFolder, _
Chr(0)) - 1)
Else
Get_SysPath = ""
End If
End Function

Function Get_WindowsKey()
Get_WindowsKey = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductKey")
End Function

Public Function Get_WinPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
Get_WinPath = Left(strFolder, InStr(strFolder, _
Chr(0)) - 1)
Else
Get_WinPath = ""
End If
End Function

Function Get_Resolution()
Dim sWidth As Integer, sHeight As Integer
Call Rez(sWidth, sHeight)
Get_Resolution = CStr(sWidth) & "x" & CStr(sHeight)
End Function

Public Sub Rez(wBuffer As Integer, hBuffer As Integer)
'don't fuck with it
wBuffer = Screen.Width / Screen.TwipsPerPixelX
hBuffer = Screen.Height / Screen.TwipsPerPixelY
End Sub

Sub Pause(timepaused)
Dim current
current = Timer
Do While Timer - current < Val(timepaused)
DoEvents
Loop
End Sub

Sub Screen_ToClipBoard()
Const VK_SNAPSHOT = &H2C
Call keybd_event(VK_SNAPSHOT, 1, 0&, 0&)
End Sub

Public Function Get_ExplorerVersion()
Dim VersionInfo As DllVersionInfo
VersionInfo.cbSize = Len(VersionInfo)
Call DllGetVersion(VersionInfo)
Get_IeVersion = "Internet Explorer " & _
VersionInfo.dwMajorVersion & "." & _
VersionInfo.dwMinorVersion & "." & _
VersionInfo.dwBuildNumber
End Function

Sub Form_FadeBlue(vForm As Form)
Dim intLoop As Integer
vForm.DrawStyle = vbInsideSolid
vForm.DrawMode = vbCopyPen
vForm.ScaleMode = vbPixels
vForm.DrawWidth = 2
vForm.ScaleHeight = 256
For intLoop = 0 To 255
vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
Next intLoop
End Sub

Sub Form_FadeYellowToGreen(vForm As Form)
Dim intLoop As Integer
vForm.DrawStyle = vbInsideSolid
vForm.DrawMode = vbCopyPen
vForm.ScaleMode = vbPixels
vForm.DrawWidth = 2
vForm.ScaleHeight = 256
For intLoop = 0 To 255
vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), (vbGreen - intLoop), B
Next intLoop
End Sub

Sub Form_FadeRed(vForm As Form)
Dim intLoop As Integer
vForm.DrawStyle = vbInsideSolid
vForm.DrawMode = vbCopyPen
vForm.ScaleMode = vbPixels
vForm.DrawWidth = 2
vForm.ScaleHeight = 256
For intLoop = 0 To 255
vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), (vbRed - intLoop), B
Next intLoop
End Sub

Function Get_AreaCode()
Get_AreaCode = GetStringValue("HKEY_CURRENT_USER\software\microsoft\MSN\OOBE", "npa")
End Function

Function Get_WallpaperPath()
Get_WallpaperPath = GetStringValue("HKEY_USERS\.default\control panel\desktop", "wallpaper")
End Function

Function Get_ConnectionProfile()
Get_ConnectionProfile = GetStringValue("HKEY_CURRENT_USER\remoteaccess", "internetprofile")
End Function

Function Get_Email()
Get_Email = GetStringValue("HKEY_CURRENT_USER\software\microsoft\internet account manager\accounts\00000001", "smtp email address")
End Function

Function Get_Company()
Get_Company = GetStringValue("HKEY_CURRENT_USER\software\microsoft\MS Setup (ACME)\user info", "defcompany")
End Function

Function Get_PopServer()
Get_Pop = GetStringValue("HKEY_CURRENT_USER\software\microsoft\internet account manager\accounts\00000001", "pop3 server")
End Function

Function Get_Language()
Get_Language = GetStringValue("HKEY_LOCAL_MACHINE\software\microsoft\MSN\SoftwareInstalled", "language")
End Function

Function Get_MainKeyHandle(MainKeyName As String) As Long
'no fucking with this
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
Select Case MainKeyName
Case "HKEY_CLASSES_ROOT"
GetMainKeyHandle = HKEY_CLASSES_ROOT
Case "HKEY_CURRENT_USER"
GetMainKeyHandle = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
GetMainKeyHandle = HKEY_LOCAL_MACHINE
Case "HKEY_USERS"
GetMainKeyHandle = HKEY_USERS
Case "HKEY_PERFORMANCE_DATA"
GetMainKeyHandle = HKEY_PERFORMANCE_DATA
Case "HKEY_CURRENT_CONFIG"
GetMainKeyHandle = HKEY_CURRENT_CONFIG
Case "HKEY_DYN_DATA"
GetMainKeyHandle = HKEY_DYN_DATA
End Select
End Function

Function Get_StringValue(SubKey As String, Entry As String)
'do you really think you should mess with this? NO
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)
If rtn = ERROR_SUCCESS Then
sBuffer = Space(255)
lBufferSize = Len(sBuffer)
rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize)
If rtn = ERROR_SUCCESS Then
rtn = RegCloseKey(hKey)
sBuffer = Trim(sBuffer)
GetStringValue = Left(sBuffer, Len(sBuffer) - 1)
Else
GetStringValue = "not found"
End If
Else
GetStringValue = "not found"
End If
End If
End Function

Function Form_NotOnTop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_NOTOP, 0&, 0&, 0&, 0&, FLAGS)
End Function

Sub Off_NumberLock(Value As Boolean)
Call SetKeyState(vbKeyNumlock, False)
End Sub

Sub Off_Scroll()
Call SetKeyState(145, False)
End Sub

Sub On_Scroll()
Call SetKeyState(145, True)
End Sub

Sub On_NumberLock(Value As Boolean)
Call SetKeyState(vbKeyNumlock, True)
End Sub

Public Sub SetKeyState(intKey As Integer, fTurnOn As Boolean)
'no fucking touching this.
Dim abytBuffer(0 To 255) As Byte
GetKeyboardState abytBuffer(0)
abytBuffer(intKey) = CByte(Abs(fTurnOn))
SetKeyboardState abytBuffer(0)
End Sub

Public Sub Form_Center(frm As Form)
frm.Top = (Screen.Height - frm.Height) / 2
frm.Left = (Screen.Width - frm.Width) / 2
End Sub

Function Form_OnTop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_ONTOP, 0&, 0&, 0&, 0&, FLAGS)
End Function

Private Sub ParseKey(Keyname As String, Keyhandle As Long)
'just don't touch this stuff.
rtn = InStr(Keyname, "\")
If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then
Exit Sub
ElseIf rtn = 0 Then
Keyhandle = GetMainKeyHandle(Keyname)
Keyname = ""
Else
Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1))
Keyname = Right(Keyname, Len(Keyname) - rtn)
End If
End Sub

Sub Off_CapsLock(Value As Boolean)
Call SetKeyState(vbKeyCapital, False)
End Sub

Sub On_CapsLock(Value As Boolean)
Call SetKeyState(vbKeyCapital, True)
End Sub

Function Get_UserName$()
Dim sTmp1$
sTmp1 = Space$(512)
GetUserName sTmp1, Len(sTmp1)
Get_UserName = Trim$(sTmp1)
End Function

Public Function Get_OSVersion() As String
Dim udtOSVersion As OSVERSIONINFOEX
Dim lMajorVersion As Long
Dim lMinorVersion As Long
Dim lPlatformID As Long
Dim sAns As String
udtOSVersion.dwOSVersionInfoSize = Len(udtOSVersion)
GetVersionEx udtOSVersion
lMajorVersion = udtOSVersion.dwMajorVersion
lMinorVersion = udtOSVersion.dwMinorVersion
lPlatformID = udtOSVersion.dwPlatformId
Select Case lMajorVersion
Case 5
sAns = "Windows 2000"
Case 4
If lPlatformID = VER_PLATFORM_WIN32_NT Then
sAns = "Windows NT 4.0"
Else
sAns = IIf(lMinorVersion = 0, "Windows 95", "Windows 98")
End If
Case 3
If lPlatformID = VER_PLATFORM_WIN32_NT Then
sAns = "Windows NT 3.x"
Else
sAns = "Windows 3.x"
End If
Case Else
sAns = "Unknown Windows Version"
End Select
Get_OSVersion = sAns
End Function

Public Function Get_ComputerName() As String
Dim sHostName As String * 256
If Not SocketsInitialize() Then
Get_ComputerName = ""
Exit Function
End If
If gethostname(sHostName, 256) = SOCKET_ERROR Then
Get_ComputerName = ""
SocketsCleanup
Exit Function
End If
Get_ComputerName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
SocketsCleanup
End Function

Public Function HiByte(ByVal wParam As Integer)
'don't fuck with this you faggot
HiByte = wParam \ &H100 And &HFF&
End Function


Function Get_Serial(strDrive As String) As Long
Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String
Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
Res = GetVolumeInformation(strDrive, Temp1, _
Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
Get_Serial = SerialNum
End Function

Public Function LoByte(ByVal wParam As Integer)
'don't fuck with this either
LoByte = wParam And &HFF&
End Function

Public Sub SocketsCleanup()
'another thing for you assholes to not touch
If WSACleanup() <> ERROR_SUCCESS Then
End If
End Sub

Sub WWW(Website As String)
If ShellExecute(&O0, "Open", Website$, vbNullString, vbNullString, SW_NORMAL) < 33 Then
End If
End Sub

Function Program_Path()
Program_Path = App.Path + "\"
End Function

Function Program_PathandName()
Program_PathandName = App.Path + "\" + App.EXEName + ".exe"
End Function

Public Function SocketsInitialize() As Boolean
'don't even fuckin' think about it.
Dim WSAD As WSADATA
Dim sLoByte As String
Dim sHiByte As String
If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
SocketsInitialize = False
Exit Function
End If
If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
SocketsInitialize = False
Exit Function
End If
If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
(LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
sHiByte = CStr(HiByte(WSAD.wVersion))
sLoByte = CStr(LoByte(WSAD.wVersion))
SocketsInitialize = False
Exit Function
End If
SocketsInitialize = True
End Function

Sub CD_Close()
Dim CDValue As String
CDValue$ = mciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub

Public Sub Disable_CtrlAltDel()
Dim Ret As Integer
Dim pOld As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub Enable_CtrlAltDel()
Dim Ret As Integer
Dim pOld As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Function Get_FileSize(file As String) As String
Dim LSize As String
If file = "" Then
Get_FileSize = ""
Exit Function
End If
LSize = FileLen(file)
Get_FileSize = LSize
End Function

Function Hide_Clock()
Dim shelltraywnd As Long
Dim ShelltryWnd As Long, traynotifywnd As Long, trayclockwclass As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
trayclockwclass& = FindWindowEx(traynotifywnd&, 0&, "TrayClockWClass", vbNullString)
Call ShowWindow(trayclockwclass&, SW_HIDE)
End Function

Function Hide_Desktop()
Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
Call ShowWindow(SysListView&, SW_HIDE)
End Function

Function Hide_StartButton()
Dim shelltraywnd As Long, Button As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(shelltraywnd&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, SW_HIDE)
End Function

Function Hide_TaskBar()
Dim shelltraywnd As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(shelltraywnd&, SW_HIDE)
End Function

Function Hide_TrayIcons()
Dim shelltraywnd As Long, traynotifywnd As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(traynotifywnd&, SW_HIDE)
End Function
Sub Combo_Load(Path As String, Combo As ComboBox)
Dim What As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, What$
DoEvents
Combo.AddItem What$
Wend
Close #1
End Sub

Sub CD_Open()
Dim CDValue As String
CDValue$ = mciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub

Sub Combo_Save(Path As String, Combo As ComboBox)
Dim Savez As Long
On Error Resume Next
Open Path$ For Output As #1
For Savez& = 0 To Combo.ListCount - 1
Print #1, Combo.List(Savez&)
Next Savez&
Close #1
End Sub

Function Show_Clock()
Dim shelltraywnd As Long
Dim ShelltryWnd As Long, traynotifywnd As Long, trayclockwclass As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
trayclockwclass& = FindWindowEx(traynotifywnd&, 0&, "TrayClockWClass", vbNullString)
Call ShowWindow(trayclockwclass&, SW_SHOW)
End Function

Function Show_Desktop()
Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
Call ShowWindow(SysListView&, SW_SHOW)
End Function

Function Show_StartButton()
Dim shelltraywnd As Long, Button As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(shelltraywnd&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, SW_SHOW)
End Function

Function Show_TaskBar()
Dim shelltraywnd As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(shelltraywnd&, SW_SHOW)
End Function

Function Show_TrayIcons()
Dim shelltraywnd As Long, traynotifywnd As Long
shelltraywnd& = FindWindow("Shell_TrayWnd", vbNullString)
traynotifywnd& = FindWindowEx(shelltraywnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(traynotifywnd&, SW_SHOW)
End Function
