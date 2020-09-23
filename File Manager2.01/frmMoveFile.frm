VERSION 5.00
Begin VB.Form frmMoveFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move File"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3195
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3195
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblCaption 
      Caption         =   "Label1"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmMoveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim yesno As Integer
    yesno = MsgBox("Are You Sure You Want To Move File: " & SourcePath & frmFileManager.ListView1.SelectedItem.Text & "  TO  " & Dir1.path, vbYesNo, "Move File")
    If yesno = vbYes Then
        Call MoveFile(SourcePath & frmFileManager.ListView1.SelectedItem.Text, Dir1.path & "\" & frmFileManager.ListView1.SelectedItem.Text)
    End If
    Unload Me
End Sub

Private Sub Dir1_Change()
    Let lblCaption.Caption = "Move File: " & SourcePath & frmFileManager.ListView1.SelectedItem.Text & " To Directory: " & Dir1.path & "\" & frmFileManager.ListView1.SelectedItem.Text
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub
