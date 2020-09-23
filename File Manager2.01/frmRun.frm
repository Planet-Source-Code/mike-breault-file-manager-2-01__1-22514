VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The Location Of The  File You Wish To Run, Or Click On Browse To Locate File:"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    'Displays The Browse Dialog To Locate File
    CommonDialog1.DialogTitle = "Browse For File"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "All Files (*.*)|*.*|"
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.FileName) = 0 Then
        Exit Sub
    End If
    txtLocation.Text = CommonDialog1.FileName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Call ShellExecute(hWnd, "Open", txtLocation.Text, "", App.path, 1)
    Unload Me
End Sub

Private Sub Form_Load()
    Let txtLocation.Text = App.path & "\"
End Sub

Private Sub txtLocation_Change()
    'If Nothing Is Into The TextBox Location, It Will
    'Disable The Ok Button
    If Len(txtLocation.Text) <= 0 Then
        Let cmdOk.Enabled = False
    Else
        Let cmdOk.Enabled = True
    End If
End Sub

