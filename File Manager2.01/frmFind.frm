VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find File"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtExt 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Text            =   "*.*"
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Search"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtText 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Case Sensitive"
      Height          =   195
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   38
      TabIndex        =   3
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   1305
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&New Search"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Named"
      Height          =   225
      Left            =   720
      TabIndex        =   0
      Top             =   390
      Width           =   615
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFind_Click()
    ListView1.ListItems.Clear
    If txtExt = "" Then txtExt = "*.*"
    Me.Caption = "Finding Files Named " & txtExt
    GetJobFiles txtPath, txtExt
    Let Me.Caption = ListView1.ListItems.Count & " Files Named " & txtExt.Text & " DONE!!"
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In Me
        ctrl.Text = ""
    Next ctrl
    ListView1.ListItems.Clear
    txtExt.SetFocus
End Sub
Private Sub Form_Load()
    ListView1.ColumnHeaders.Add 1, "name", "Name"
    ListView1.ColumnHeaders.Add 2, "infolder", "In Folder"
    ListView1.ColumnHeaders.Add 3, "size", "Size"
    ListView1.ColumnHeaders.Add 4, "type", "Type"
    ListView1.ColumnHeaders.Add 5, "date", "Date Modified"
    With txtExt
        .AddItem "*.*"
        .AddItem "*.ico"
        .AddItem "*.bmp"
        .AddItem "*.jpg"
        .AddItem "*.exe"
        .AddItem "*.vbp"
        .AddItem "*.frm"
        .AddItem "*.vbw"
        .AddItem "*.vbg"
        .AddItem "*.frx"
        .AddItem "*.dll"
        .AddItem "*.bas"
        .AddItem "*.cls"
        .AddItem "*.ocx"
        .AddItem "*.oca"
        .AddItem "*.exp"
        .AddItem "*.lib"
        .AddItem "*.ctl"
        .AddItem "*.ctx"
        .AddItem "*.txt"
        .AddItem "*.ani"
        .AddItem "*.doc"
        .AddItem "*.tm2"
        .AddItem "*.dat"
        .AddItem "*.mpg"
        .AddItem "*.avi"
        .AddItem "*.mov"
        .AddItem "*.wav"
        .AddItem "*.mp3"
        .AddItem "*.mid"
        .AddItem "*.log"
        .AddItem "*.inf"
        .AddItem "*.tmp"
        .AddItem "*.cab"
        .AddItem "*.zip"
        .AddItem "*.rar"
        .AddItem "*.ini"
        .AddItem "*.sys"
        .AddItem "*.com"
        .AddItem "*.pif"
'        .AddItem "*."
    End With
End Sub
Private Function GetJobFiles(drivepath As String, ext As String) As Boolean
    Dim XDir() As String
    Dim tmpdir As String
    Dim DirCount As Integer
    Dim x As Integer
    On Error Resume Next
    DirCount = 0
    ReDim XDir(0) As String
    XDir(DirCount) = ""
    Getfindfiles drivepath
    If Right(drivepath, 1) <> "\" Then
        drivepath = drivepath & "\"
    End If
    DoEvents
    tmpdir = Dir(drivepath, vbDirectory)
    Do While tmpdir <> ""
       If tmpdir <> "." And tmpdir <> ".." Then
           If (GetAttr(drivepath & tmpdir) And vbDirectory) = vbDirectory Then
                Getfindfiles drivepath & tmpdir
                XDir(DirCount) = drivepath & tmpdir & "\"
                DirCount = DirCount + 1
                ReDim Preserve XDir(DirCount) As String
                
            End If
        End If
        tmpdir = Dir
    Loop
    For x = 0 To (UBound(XDir) - 1)
        GetJobFiles1 XDir(x)
    Next x
End Function

Private Function GetJobFiles1(drivepath As String) As Boolean
    Dim XDir() As String
    Dim tmpdir As String
    Dim DirCount As Integer
    Dim x As Integer
    DirCount = 0
    ReDim XDir(0) As String
    XDir(DirCount) = ""
    If Right(drivepath, 1) <> "\" Then
        drivepath = drivepath & "\"
    End If
    DoEvents
    tmpdir = Dir(drivepath, vbDirectory)
    Do While tmpdir <> ""
       If tmpdir <> "." And tmpdir <> ".." Then
           If (GetAttr(drivepath & tmpdir) And vbDirectory) = vbDirectory Then
                Getfindfiles drivepath & tmpdir
                XDir(DirCount) = drivepath & tmpdir & "\"
                DirCount = DirCount + 1
                ReDim Preserve XDir(DirCount) As String
            End If
        End If
        tmpdir = Dir
    Loop
    For x = 0 To (UBound(XDir) - 1)
        GetJobFiles1 XDir(x)
    Next x
End Function

Public Function Getfindfiles(driv As String)
    Dim fs, fsFolder, fsFileList, fsFile, ls As ListItem
    Static fsFilecounter As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fsFolder = fs.GetFolder(driv & "\")
    Set fsFileList = fsFolder.Files
    For Each fsFile In fsFileList
    If txtExt = "*.*" Then
            Set ls = ListView1.ListItems.Add(1, driv & fsFile.Name, fsFile.Name)
            ls.SubItems(1) = driv
            ls.SubItems(2) = Round(fsFile.Size / 1024) & "KB"
            If Round(fsFile.Size / 1024) = 0 Then
                ls.SubItems(2) = fsFile.Size & "Bytes"
            End If
            ls.SubItems(3) = fsFile.Type
            ls.SubItems(4) = fsFile.DateLastModified
    ElseIf LCase(Right(fsFile.Name, 3)) = LCase(Right(txtExt, 3)) Then
        If txtText <> "" Then
            checkfile fsFile.Name, driv, txtText
        End If
            Set ls = ListView1.ListItems.Add(1, driv & fsFile.Name, fsFile.Name)
            ls.SubItems(1) = driv
            ls.SubItems(2) = Round(fsFile.Size / 1024) & "KB"
            If Round(fsFile.Size / 1024) = 0 Then
                ls.SubItems(2) = fsFile.Size & "Bytes"
            End If
            ls.SubItems(3) = fsFile.Type
            ls.SubItems(4) = fsFile.DateLastModified
    End If
    Next
End Function

  
Public Function GetSelectedFile(strPath As String, listnum As String) As String
    If Right(strPath, 1) <> "\" Then
        GetSelectedFile = strPath & "\" & listnum
    Else
        GetSelectedFile = strPath & listnum
    End If
End Function

Private Sub checkfile(listnum As String, drivpath As String, searchfor As String)
    Dim strSearchFor$, intFileCount%, intListCount%, strGsearch$
    Dim gFindPosition&, strFound$, ls As ListItem
    Dim fs, fsFolder, fsFileList, fsFile
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fsFolder = fs.GetFolder(drivpath & "\")
    Set fsFileList = fsFolder.Files
    For Each fsFile In fsFileList
        If fsFile.Name = listnum Then
            strSearchFor$ = searchfor
            Open GetSelectedFile(drivpath, listnum) For Binary As 1
            strGsearch$ = String(LOF(1), Chr$(0))
            Get 1, , strGsearch$
            Close 1
            If chkMatchCase.Value = vbChecked Then
                strGsearch$ = strGsearch$
            Else
                strGsearch$ = LCase(strGsearch$)
            End If
            gFindPosition& = InStr(strGsearch$, strSearchFor$)
            If gFindPosition& <> 0 Then
                strFound$ = Mid(strGsearch$, gFindPosition&, Len(strSearchFor$))
            End If
            If strFound$ = strSearchFor$ Then
               Set ls = ListView1.ListItems.Add(1, drivpath & fsFile.Name, fsFile.Name)
                ls.SubItems(1) = drivpath
                ls.SubItems(2) = Round(fsFile.Size / 1024) & "KB"
                If Round(fsFile.Size / 1024) = 0 Then
                    ls.SubItems(2) = fsFile.Size & "Bytes"
                End If
                ls.SubItems(3) = fsFile.Type
                ls.SubItems(4) = fsFile.DateCreated
            End If
        End If
    Next
    strGsearch$ = ""
    strFound$ = ""
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = vbKeyBack) Then 'TO GET ONLY NUMERIC INPUT
        KeyAscii = 0
    End If
End Sub
