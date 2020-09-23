VERSION 5.00
Begin VB.Form frmFilters 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display In File List"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
      Begin VB.OptionButton optFoldersOnly 
         Caption         =   "F&olders Only"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton optFilesOnly 
         Caption         =   "&Files Only"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton optFilesandFolders 
         Caption         =   "Fil&es And Folders"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show ONLY The Following Files"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "*.*"
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Let optFilesOnly.Value = True
    With Combo1
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
