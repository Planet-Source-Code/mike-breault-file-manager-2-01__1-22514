VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExtractIcon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extract Icon"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1785
   Icon            =   "frmExtractIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSaveIcon 
      Caption         =   "&Save Icon"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdExtractIcon 
      Caption         =   "&Extract Icon"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox pctIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   577
      ScaleHeight     =   555
      ScaleWidth      =   570
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmExtractIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExtractIcon_Click()
  On Error Resume Next
    CommonDialog1.FileName = strProgram
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Locate File To Extract Icon"
    CommonDialog1.Filter = "Icon Resources (*.ico;*.exe;*.dll)|*.ico;*.exe;*.dll|All files|*.*"
    CommonDialog1.Action = 1
    strProgram = CommonDialog1.FileName
    DestroyIcon lngIcon
    lngIcon = ExtractIcon(App.hInstance, strProgram, 0)
    If lngIcon = 0 Then
      MsgBox "Sorry, No Icon Available."
    Else
        pctIcon.Cls
        pctIcon.AutoSize = True
        pctIcon.AutoRedraw = True
        DrawIcon pctIcon.hdc, 0, 0, lngIcon
        pctIcon.Refresh
    End If
End Sub

Private Sub cmdSaveIcon_Click()
  On Error Resume Next
    CommonDialog1.DialogTitle = "Save Icon As..."
    CommonDialog1.FileName = strSaveIconFile
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Bitmap Image (*.bmp)|*.bmp"
    CommonDialog1.ShowSave
    strSaveIconFile = CommonDialog1.FileName
    SavePicture pctIcon.Image, strSaveIconFile
End Sub
