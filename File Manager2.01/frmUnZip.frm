VERSION 5.00
Begin VB.Form frmUnZip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UnZip File"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label lblFolder 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblCaption 
      Caption         =   "Unzip File"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmUnZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo vbErrorHandler
    Dim oUnZip As CGUnzipFiles
    Set oUnZip = New CGUnzipFiles
    With oUnZip
        .ZipFileName = SourcePath & frmFileManager.ListView1.SelectedItem.Text
        .ExtractDir = Dir1.path
        .HonorDirectories = False
        If .Unzip <> 0 Then
            MsgBox .GetLastMessage
        End If
    End With
    Set oUnZip = Nothing
    MsgBox SourcePath & frmFileManager.ListView1.SelectedItem.Text & " Extracted Successfully to " & Dir1.path & "\"
    Exit Sub
vbErrorHandler:
    MsgBox Err.Description, vbOKOnly, "Error"
    Unload Me
End Sub

Private Sub Dir1_Change()
    Let lblFolder.Caption = Dir1.path & "\"
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Let lblCaption.Caption = "Unzip File: " & SourcePath & frmFileManager.ListView1.SelectedItem.Text & " To Folder:"
    Let lblFolder.Caption = Dir1.path & "\"
End Sub
