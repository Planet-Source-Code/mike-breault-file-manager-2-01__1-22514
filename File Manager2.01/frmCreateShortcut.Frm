VERSION 5.00
Begin VB.Form frmCreateShortcut 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Shortcut"
   ClientHeight    =   3810
   ClientLeft      =   3855
   ClientTop       =   1710
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3810
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   3405
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtShortcutName 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   5130
      End
      Begin VB.TextBox txtExename 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   5130
      End
      Begin VB.TextBox txtShortcutDir 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   4770
      End
      Begin VB.CommandButton cmdShortcutPathDialog 
         Caption         =   "..."
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblShortcutName 
         Alignment       =   2  'Center
         Caption         =   "Enter Name For Shortcut:"
         Height          =   285
         Left            =   165
         TabIndex        =   17
         Top             =   720
         Width           =   4770
      End
      Begin VB.Label lblExename 
         Alignment       =   2  'Center
         Caption         =   "Path and File Name For Shortcut:"
         Height          =   285
         Left            =   165
         TabIndex        =   16
         Top             =   120
         Width           =   4770
      End
      Begin VB.Label lblShortcutDir 
         Alignment       =   2  'Center
         Caption         =   "Enter Directory For Shortcut To Be Created Or Browse:"
         Height          =   270
         Left            =   165
         TabIndex        =   15
         Top             =   1305
         Width           =   4890
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quick Access"
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   5415
      Begin VB.CheckBox chkQuickLaunchBar 
         Caption         =   "Quick Launch Bar"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkFavorites 
         Alignment       =   1  'Right Justify
         Caption         =   "&Favorites"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkAccessories 
         Alignment       =   1  'Right Justify
         Caption         =   "&Accessories"
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkStartup 
         Alignment       =   1  'Right Justify
         Caption         =   "S&tartup"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkPrograms 
         Caption         =   "&Programs"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkStartmenu 
         Caption         =   "&Start Menu"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkDesktop 
         Caption         =   "&Desktop"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Text            =   "5"
      Top             =   5400
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   5280
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3880
      TabIndex        =   2
      Top             =   3405
      Width           =   1650
   End
   Begin VB.CommandButton cmdCreateShortcut 
      Caption         =   "&Create Shortcut"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3405
      Width           =   1650
   End
End
Attribute VB_Name = "frmCreateShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function fCreateShellGroup Lib "VB5STKIT.DLL" (ByVal lpstrDirName As String) As Long
Private Declare Function fCreateShellLink Lib "VB5STKIT.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Private Declare Function fRemoveShellLink Lib "VB5STKIT.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long

'Function to get Windows directory
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Structure used by SHFileOperation
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
  
'Function and constants for performing operations on file system objects
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Const FO_DELETE = &H3
Const FO_COPY = &H2
Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const FOF_CONFIRMMOUSE = &H2
Const FOF_FILESONLY = &H80
Const FOF_MULTIDESTFILES = &H1
Const FOF_NOCONFIRMATION = &H10
Const FOF_NOCONFIRMMKDIR = &H200
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_SILENT = &H4
Const FOF_SIMPLEPROGRESS = &H100
Const FOF_WANTMAPPINGHANDLE = &H20
Const FOF_ALLOWUNDO = &H40

'Variable to store the Windows directory
Dim WinDir As String

'Buffer and constant used for API functions
Dim msBuffer As String * 255
Const BUFFERSIZE = 255

Private Sub CreateShortcut(ByVal sDir As String, ByVal sShortcut As String, ByVal sExecutable As String, ByVal sArguments As String)
    Dim lRet As Long
    Dim udtFileOp As SHFILEOPSTRUCT
    If LCase$(Left$(sShortcut, 4)) = ".lnk" Or LCase$(Left$(sShortcut, 4)) = ".pif" Then
        sShortcut = Left$(sShortcut, Len(sShortcut) - 4)
    End If
    lRet = fCreateShellLink("", sShortcut, sExecutable, sArguments)
    If lRet Then
        udtFileOp.hwnd = Me.hwnd
        udtFileOp.wFunc = FO_MOVE
        If LCase$(Right$(Dir$(WinDir & "\start menu\programs\" & sShortcut & ".*"), 3)) = "lnk" Then
            udtFileOp.pFrom = WinDir & "\start menu\programs\" & sShortcut & ".lnk"
            udtFileOp.pTo = sDir & "\" & sShortcut & ".lnk"
        Else
            udtFileOp.pFrom = WinDir & "\start menu\programs\" & sShortcut & ".pif"
            udtFileOp.pTo = sDir & "\" & sShortcut & ".pif"
        End If
        udtFileOp.fFlags = FOF_RENAMEONCOLLISION
        lRet = SHFileOperation(udtFileOp)
        If lRet <> 0 And Not udtFileOp.fAnyOperationsAborted Then
            MsgBox "Unable to create shortcut for " & sExecutable & ".", vbInformation
        End If
    Else
        MsgBox "Unable to create shortcut for " & sExecutable & ".", vbInformation
    End If
End Sub

Private Sub chkDesktop_Click()
    If chkDesktop.Value = 1 Then
        chkQuickLaunchBar.Value = 0
        chkStartmenu.Value = 0
        chkPrograms.Value = 0
        chkStartup.Value = 0
        chkAccessories.Value = 0
        chkFavorites.Value = 0
        txtShortcutDir = WinDir & "\Desktop"
    End If
End Sub

Private Sub chkQuickLaunchBar_Click()
    If chkQuickLaunchBar.Value = 1 Then
        chkDesktop.Value = 0
        chkStartmenu.Value = 0
        chkPrograms.Value = 0
        chkStartup.Value = 0
        chkAccessories.Value = 0
        chkFavorites.Value = 0
        txtShortcutDir = WinDir & "\Application Data\Microsoft\Internet Explorer\Quick Launch"
    End If
End Sub

Private Sub chkStartmenu_Click()
    If chkStartmenu.Value = 1 Then
        chkQuickLaunchBar.Value = 0
        chkDesktop.Value = 0
        chkPrograms.Value = 0
        chkStartup.Value = 0
        chkAccessories.Value = 0
        chkFavorites.Value = 0
        txtShortcutDir = WinDir & "\Start Menu"
    End If
End Sub

Private Sub chkPrograms_Click()
    If chkPrograms.Value = 1 Then
        chkQuickLaunchBar.Value = 0
        chkStartmenu.Value = 0
        chkDesktop.Value = 0
        chkStartup.Value = 0
        chkAccessories.Value = 0
        chkFavorites.Value = 0
        txtShortcutDir = WinDir & "\Start Menu\Programs"
    End If
End Sub

Private Sub chkStartup_Click()
    If chkStartup.Value = 1 Then
        chkQuickLaunchBar.Value = 0
        chkStartmenu.Value = 0
        chkDesktop.Value = 0
        chkPrograms.Value = 0
        chkAccessories.Value = 0
        chkFavorites.Value = 0
        txtShortcutDir = WinDir & "\Start Menu\Programs\Startup"
    End If
End Sub

Private Sub chkAccessories_Click()
    If chkAccessories.Value = 1 Then
        chkQuickLaunchBar.Value = 0
        chkStartup.Value = 0
        chkStartmenu.Value = 0
        chkDesktop.Value = 0
        chkPrograms.Value = 0
        chkFavorites.Value = 0
        txtShortcutDir = WinDir & "\Start Menu\Programs\Accessories"
    End If
End Sub

Private Sub chkFavorites_Click()
    If chkFavorites.Value = 1 Then
        chkQuickLaunchBar.Value = 0
        chkAccessories.Value = 0
        chkStartup.Value = 0
        chkStartmenu.Value = 0
        chkDesktop.Value = 0
        chkPrograms.Value = 0
        txtShortcutDir = WinDir & "\Favorites"
    End If
End Sub

Private Sub cmdCreateShortcut_Click()
    If Len(txtShortcutName) = 0 Then
        MsgBox "Please enter a Shortcut Name. Eg... Notepad", vbQuestion
        txtShortcutName.SetFocus
        Exit Sub
    End If
    If Len(txtShortcutDir) = 0 Then
        MsgBox "Please enter a Directory to create you new shortcut in." & vbCrLf & vbCrLf & "Eg... C:\Windows\Desktop", vbQuestion
        txtShortcutDir.SetFocus
        Exit Sub
    End If
    CreateShortcut txtShortcutDir, txtShortcutName, txtExename, ""
    Timer1.Enabled = True
    Unload Me
End Sub

Function FileOrDirExists(Optional ByVal sFileName As Variant, Optional ByVal sPath As Variant) As Boolean
    On Error GoTo Oops
    If IsMissing(sPath) Then
        If Len(Dir$(sFileName)) Then FileOrDirExists = True
    Else
        If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
        If IsMissing(sFileName) Then
            If Len(Dir$(sPath, vbDirectory)) Then FileOrDirExists = True
        Else
            If Len(Dir$(sPath & sFileName)) Then FileOrDirExists = True
        End If
    End If
    Exit Function
Oops:
    Exit Function
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
    Let txtShortcutName.Text = ""
    Let txtShortcutDir.Text = ""
    Let chkAccessories.Value = 0
    Let chkQuickLaunchBar.Value = 0
    Let chkStartmenu.Value = 0
    Let chkPrograms.Value = 0
    Let chkStartup.Value = 0
    Let chkFavorites.Value = 0
End Sub

Private Sub cmdShortcutPathDialog_Click()
    Dim udtBrowseInfo As BROWSEINFO
    Dim lRet As Long
    Dim lPathID As Long
    Dim sPath As String
    Dim nNullPos As Integer
    txtShortcutDir.SetFocus
    udtBrowseInfo.hOwner = Me.hwnd
    udtBrowseInfo.pidlRoot = 0&
    udtBrowseInfo.lpszTitle = "Select a folder for Shortcut"
    udtBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS
    lPathID = SHBrowseForFolder(udtBrowseInfo)
    sPath = Space$(512)
    lRet = SHGetPathFromIDList(lPathID, sPath)
    If lRet Then
        nNullPos = InStr(sPath, vbNullChar)
        txtShortcutDir = Left(sPath, nNullPos - 1)
    End If
End Sub

Private Sub Form_Load()
    Text1.Text = 5
    Dim lBytes As Long
    lBytes = GetWindowsDirectory(msBuffer, BUFFERSIZE)
    WinDir = Left$(msBuffer, lBytes)
End Sub

Private Sub Timer1_Timer()
    Text1.Text = Text1.Text - 1
    If Text1.Text = 0 Then
        txtExename.Text = ""
        txtShortcutName.Text = ""
    End If
    If txtExename.Text = "" Then
        Text1.Text = 5
        Timer1.Enabled = False
    End If
End Sub

Private Sub txtExeName_LostFocus()
    If Not FileOrDirExists(txtExename) And ActiveControl.Name <> "cmdExeDialog" Then
        MsgBox "File Not Found" & vbCrLf & vbCrLf & txtExename, vbInformation
        With txtExename
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
End Sub

Private Sub txtShortcutDir_LostFocus()
    If Not FileOrDirExists(, txtShortcutDir) And ActiveControl.Name <> "cmdShortcutPathDialog" Then
        MsgBox "Path Not Found" & vbCrLf & vbCrLf & txtShortcutDir, vbInformation
        With txtShortcutDir
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
End Sub

Private Sub txtShortcutName_KeyPress(KeyAscii As Integer)
    'Disallow characters which are not valid in file names
    Select Case KeyAscii
        Case 92, 47, 58, 42, 63, 34, 60, 62, 124
            MsgBox "A filename cannot contain any of the following characters:" & vbCrLf & vbCrLf & Space$(15) & Space$(15) & "/ \ : * ? "" < > |", vbCritical, App.Title
            KeyAscii = 0
    End Select
End Sub

