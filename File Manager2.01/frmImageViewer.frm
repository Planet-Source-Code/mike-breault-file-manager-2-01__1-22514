VERSION 5.00
Begin VB.Form frmImageViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Viewer"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "frmImageViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Text            =   "*.jpg"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Loop SlideShow"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdStopSlideShow 
      Caption         =   "&End SlideShow"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   7680
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSlideShow 
      Caption         =   "&Slide Show"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   7200
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   3960
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   2280
      ScaleHeight     =   6915
      ScaleWidth      =   8115
      TabIndex        =   3
      Top             =   240
      Width           =   8175
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   1020
         Left            =   0
         Picture         =   "frmImageViewer.frx":0442
         ScaleHeight     =   960
         ScaleWidth      =   7815
         TabIndex        =   4
         Top             =   0
         Width           =   7875
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      LargeChange     =   20
      Left            =   2280
      Min             =   1
      SmallChange     =   20
      TabIndex        =   2
      Top             =   7200
      Value           =   1
      Width           =   8175
   End
   Begin VB.CommandButton cmdNothing 
      Enabled         =   0   'False
      Height          =   255
      Left            =   10440
      TabIndex        =   1
      Top             =   7200
      Width           =   255
   End
   Begin VB.VScrollBar VS 
      Height          =   6975
      LargeChange     =   20
      Left            =   10440
      SmallChange     =   20
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblCurrentFile 
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10695
   End
   Begin VB.Menu mnuProp 
      Caption         =   "File Properties"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeBackground 
         Caption         =   "Display As Background"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmImageViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER = 20

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSlideShow_Click()
    cmdSlideShow.Enabled = False
    cmdStopSlideShow.Enabled = True
    List2.Clear
    Dim i As Integer
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
            List2.AddItem File1.List(i)
        End If
    Next i
    
    Timer1.Interval = Val(3) * 500
    Timer1.Enabled = True
End Sub

Private Sub cmdStopSlideShow_Click()
    cmdSlideShow.Enabled = True
    cmdStopSlideShow.Enabled = False
    Timer1.Enabled = False
End Sub

Private Sub Combo1_Click()
    If Combo1 = "*.bmp" Then
        File1.Pattern = "*.bmp"
    ElseIf Combo1 = "*.jpg" Then
        File1.Pattern = "*.jpg"
    ElseIf Combo1 = "*.pcx" Then
        File1.Pattern = "*.pcx"
    ElseIf Combo1 = "*.gif" Then
        File1.Pattern = "*.gif"
    ElseIf Combo1 = "" Then
        File1.Pattern = "*."
    End If
End Sub

Private Sub Dir1_Change()
    lblCurrentFile.Caption = Dir1.path & "\"
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

Private Sub File1_Click()
    Dim FileName As String
    FileName = File1.path + "\" + File1.FileName
    MousePointer = vbHourglass
    DoEvents
    Picture2.Picture = LoadPicture(FileName)
    If Picture2.Height < Picture1.Height Then
        'make scroll bars invisible
        Let VS.Visible = False
        If Picture2.Width < Picture1.Width Then
            Let HS.Visible = False
            Let cmdNothing.Visible = False
        End If
    ElseIf Picture2.Height > Picture1.Height Then
        'make scroll bars visible
        Let VS.Visible = True
        Let cmdNothing.Visible = True
        If Picture2.Width > Picture1.Width Then
            Let HS.Visible = True
            Let cmdNothing.Visible = True
        End If
    End If
    MousePointer = vbDefault
    Let lblCurrentFile.Caption = FileName
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Right(File1.FileName, 4) = ".bmp" Then
        If Button <> vbRightButton Then
             Exit Sub
        End If
        mnuChangeBackground.Visible = True
        PopupMenu Me.mnuProp
    Else
        If Button <> vbRightButton Then
             Exit Sub
        End If
        mnuChangeBackground.Visible = False
        PopupMenu Me.mnuProp
    End If
End Sub

Private Sub Form_Load()
    cmdStopSlideShow.Enabled = False
    With Combo1
        .AddItem "*.jpg"
        .AddItem "*.bmp"
        .AddItem "*.gif"
        .AddItem "*.pcx"
    End With
    File1.Pattern = "*.jpg"
    lblCurrentFile.Caption = Dir1.path & "\"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MousePointer = vbDefault
End Sub

Private Sub HS_Change()
    Call Picture2.Move(-HS.Value, Picture2.Top)
End Sub

Private Sub HS_Scroll()
    Call Picture2.Move(-HS.Value, Picture2.Top)
End Sub

Private Sub mnuChangeBackground_Click()
    Dim WallPaper As Long
    WallPaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, Dir1.path & "\" & File1.FileName, 0)
End Sub

Private Sub mnuDelete_Click()
    Dim lResult As Long
    lResult = ShellDelete(Dir1.path & "\" & File1.FileName)
    Dir1.Refresh
    File1.Refresh
End Sub

Private Sub mnuProperties_Click()
    'Displays a files properties when popupmenu is clicked
    'and the properties item is choosen
    Dim r As Long
    r = ShowFileProperties(Dir1.path & "\" & File1.FileName, Me.hwnd) 'To show the properties dialog pass the filename and the owner of the dialog
    If r <= 32 Then
        MsgBox "Error In Properties Window", vbOKOnly, "Error"
    End If
    MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
    Static i As Integer
    Dim fn As String
    fn = Dir1.path + "\"
    If i < List2.ListCount Then
        lblCurrentFile.Caption = List2.List(i) + " (" + Format(i + 1) + "/" + Format(List2.ListCount) + ")"
        Display_Picture (fn + List2.List(i))
        i = i + 1
    Else
        i = 0
        If Check1.Value = 1 Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
            Let cmdSlideShow.Enabled = True
            Let cmdStopSlideShow.Enabled = False
        End If
    End If
End Sub

Private Sub VS_Change()
    Call Picture2.Move(Picture2.Left, -VS.Value)
End Sub

Private Sub VS_Scroll()
    Call Picture2.Move(Picture2.Left, -VS.Value)
End Sub

Private Sub Display_Picture(fn As String)
    Picture2.Picture = LoadPicture(fn)
End Sub

