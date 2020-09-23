VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmWindowsOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Options"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   3338
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   2258
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   278
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Screen Saver"
      TabPicture(0)   =   "frmWindowsOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "cmdEnable"
      Tab(0).Control(2)=   "cmdDisable"
      Tab(0).Control(3)=   "frmActivateSS"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Settings"
      TabPicture(1)   =   "frmWindowsOptions.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblComputerName"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblUserName"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblOSPlatform"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblOSVersion"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblUpdate"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblCPUMake"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblModel"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblSpeed"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblRegisteredUser"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblOrganization"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblProductID"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdMinimize"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdShowIcons"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdHideIcons"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdShowTaskbar"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdHideTaskbar"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Screen Resolution"
      TabPicture(2)   =   "frmWindowsOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(1)=   "List1"
      Tab(2).Control(2)=   "cmdApplyRes"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdHideTaskbar 
         Caption         =   "Hide Taskbar"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowTaskbar 
         Caption         =   "Show Taskbar"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdHideIcons 
         Caption         =   "Hide Desktop Icons"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowIcons 
         Caption         =   "Show Desktop Icons"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdMinimize 
         Caption         =   "Minimize ALL Windows"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdApplyRes 
         Caption         =   "Apply"
         Height          =   255
         Left            =   -70080
         TabIndex        =   7
         Top             =   2880
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   -74280
         TabIndex        =   6
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton frmActivateSS 
         Caption         =   "Activate Screen Saver"
         Height          =   255
         Left            =   -72742
         TabIndex        =   5
         Top             =   1620
         Width           =   2175
      End
      Begin VB.CommandButton cmdDisable 
         Caption         =   "Disable ScreenSaver"
         Height          =   375
         Left            =   -71542
         TabIndex        =   4
         Top             =   540
         Width           =   1935
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "Enable ScreenSaver"
         Height          =   375
         Left            =   -73702
         TabIndex        =   3
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label lblProductID 
         Caption         =   "Product ID"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label lblOrganization 
         Caption         =   "Organization"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label lblRegisteredUser 
         Caption         =   "Registered User"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label lblSpeed 
         Caption         =   "Speed"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lblModel 
         Caption         =   "Model"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label lblCPUMake 
         Caption         =   "CPU Make"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label lblUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label lblOSVersion 
         Caption         =   "OS Version"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label lblOSPlatform 
         Caption         =   "OS Platform"
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblUserName 
         Caption         =   "User Name"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblComputerName 
         Caption         =   "Computer Name"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ScreenSaver (Enabled-Disabled)"
         Height          =   255
         Left            =   -73462
         TabIndex        =   13
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmWindowsOptions.frx":0054
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmWindowsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Const VER_PLATFORM_WIN32_NT = 2
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32s = 0

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Private tmpVersionInfo As OSVERSIONINFO

Dim tmpRegKey As String
Dim tmpBuffer As String * 255

'Api's for starting screensaver
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_SCREENSAVE = &HF140&

'Api's for getting computer and user's name
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'Api's for hiding and showing taskbar
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

'Api's for showing and hiding desktop icons
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'Api's for minimizing all windows
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B

'Api for enabling and disabling screensaver
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
Private Const SPI_SETSCREENSAVEACTIVE = 17


Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DM_DISPLAYFLAGS = &H200000
Const DM_DISPLAYFREQUENCY = &H400000

Private Declare Function ChangeDisplaySettings Lib "User32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "User32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ExitWindowsEx Lib "User32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const BITSPIXEL = 12

' /* Flags for ChangeDisplaySettings */
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H2
Const CDS_FULLSCREEN = &H4
Const CDS_GLOBAL = &H8
Const CDS_SET_PRIMARY = &H10
Const CDS_RESET = &H40000000
Const CDS_SETRECT = &H20000000
Const CDS_NORESET = &H10000000

' /* Return values for ChangeDisplaySettings */
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const DISP_CHANGE_FAILED = -1
Const DISP_CHANGE_BADMODE = -2
Const DISP_CHANGE_NOTUPDATED = -3
Const DISP_CHANGE_BADFLAGS = -4
Const DISP_CHANGE_BADPARAM = -5

Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Dim D() As DEVMODE, lNumModes As Long

Private Sub cmdApplyRes_Click()
    Dim l As Long, Flags As Long, x As Long
    x = List1.ListIndex
    D(x).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    Flags = CDS_UPDATEREGISTRY
    l = ChangeDisplaySettings(D(x), Flags)
    Select Case l
        Case DISP_CHANGE_RESTART
            l = MsgBox("This Change Will Not Take Effect Until You Reboot. Reboot Now?", vbYesNo)
            If l = vbYes Then
                Flags = 0
                l = ExitWindowsEx(EWX_REBOOT, Flags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        Case Else
            MsgBox "Error Changing Resolution"
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDisable_Click()
    ToggleScreenSaverActive (False)
    Let Label1.Caption = "ScreenSaver Is Disabled"
End Sub

Private Sub cmdEnable_Click()
    ToggleScreenSaverActive (True)
    Let Label1.Caption = "ScreenSaver Is Enabled"
End Sub

Private Sub cmdHideIcons_Click()
    Dim hWnd As Long
    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 0
End Sub

Private Sub cmdHideTaskbar_Click()
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Private Sub cmdMinimize_Click()
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(77, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdShowIcons_Click()
    Dim hWnd As Long
    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 5
End Sub

Private Sub cmdShowTaskbar_Click()
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub

Private Sub Form_Load()
    tmpVersionInfo.dwOSVersionInfoSize = 148
    GetVersionEx tmpVersionInfo
    If tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        If tmpVersionInfo.dwMinorVersion = 0 Then
            lblOSPlatform.Caption = "OS System: Microsoft Windows '95"
        Else
            lblOSPlatform.Caption = "OS System: Microsoft Windows '98"
        End If
    ElseIf tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        If tmpVersionInfo.dwMajorVersion = 4 Then
            lblOSPlatform.Caption = "OS System: Microsoft Windows NT"
        Else
            lblOSPlatform.Caption = "OS System: Microsoft Windows 2000"
        End If
    End If
    lblOSVersion.Caption = "OS Version: " & tmpVersionInfo.dwMajorVersion & "." & Format(tmpVersionInfo.dwMinorVersion, "00") & "." & tmpVersionInfo.dwBuildNumber
    lblUpdate.Caption = "Update: " & Left(tmpVersionInfo.szCSDVersion, InStr(1, tmpVersionInfo.szCSDVersion, Chr(0)))
    If tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        tmpRegKey = RK_WIN32_OS
    ElseIf tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        tmpRegKey = RK_WIN32_OS_NT
    End If
    lblOrganization.Caption = "Organization: " & GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "RegisteredOrganization")
    lblRegisteredUser.Caption = "Registered User: " & GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "RegisteredOwner")
    lblProductID.Caption = "Product ID#: " & GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "ProductID")
    lblCPUMake.Caption = "CPU Make: " & GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "VendorIdentifier")
    lblModel.Caption = "Model: " & GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "Identifier")
    tmpBuffer = GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "~MHZ")
    lblSpeed.Caption = "Speed: " & Trim(tmpBuffer) & " MHz"
    Dim PCName As String
    Dim PCName2 As String
    Dim CompName As Long
    Dim CompUserName As Long
    CompUserName = NameOfPC(PCName2)
    CompName = NameOfPC(PCName)
    lblUserName.Caption = "Computer User Name: " & UserName
    lblComputerName.Caption = "Computer Name: " & PCName2
    Let Label1.Caption = "ScreenSaver Is Enabled"
    Dim l As Long, lMaxModes As Long
    Dim lBits As Long, lWidth As Long, lHeight As Long
    lBits = GetDeviceCaps(hdc, BITSPIXEL)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    lMaxModes = 8
    ReDim D(0 To lMaxModes) As DEVMODE
    lNumModes = 0
    l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Do While l
        List1.AddItem D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel
        If lBits = D(lNumModes).dmBitsPerPel And _
           lWidth = D(lNumModes).dmPelsWidth And _
           lHeight = D(lNumModes).dmPelsHeight Then
            List1.ListIndex = List1.NewIndex
        End If
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve D(0 To lMaxModes) As DEVMODE
        End If
        l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Loop
    lNumModes = lNumModes - 1
End Sub

Public Function ToggleScreenSaverActive(Active As Boolean) As Boolean
    Dim lActiveFlag As Long
    Dim retvaL As Long
    lActiveFlag = IIf(Active, 1, 0)
    retvaL = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, lActiveFlag, 0, 0)
    ToggleScreenSaverActive = retvaL > 0
End Function

Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
End Function

Public Property Get UserName() As Variant
     Dim sBuffer As String
     Dim lSize As Long
     sBuffer = Space$(255)
     lSize = Len(sBuffer)
     Call GetUserName(sBuffer, lSize)
     UserName = Left$(sBuffer, lSize)
End Property

Private Sub frmActivateSS_Click()
    'Activates Current ScreenSaver
    Dim startSS As Long
    startSS = SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
