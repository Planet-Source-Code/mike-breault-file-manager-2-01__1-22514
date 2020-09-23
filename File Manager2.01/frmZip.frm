VERSION 5.00
Begin VB.Form frmZip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zip Files"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Click On The Files You Want To Remove From The Combo List"
      Top             =   3360
      Width           =   6015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3615
   End
   Begin VB.FileListBox File1 
      Height          =   5160
      Left            =   6000
      TabIndex        =   2
      ToolTipText     =   "Double Click On The Files You Want To Compile"
      Top             =   0
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   0
      Picture         =   "frmZip.frx":0000
      Top             =   0
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Click On The Files You Want To Remove From The Combo List"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Zip Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click On The Files You Want To Compile:"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim zipname As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If Text1.Text = "" Then
        MsgBox "Enter The Name Of The ZipFile", vbOKOnly, "Enter Zipfile Name"
    Else
        If List1.ListCount < 1 Then
            MsgBox "Input Some Files First, Then Continue", vbOKOnly, "Input Files"
        Else
            Dim oZip As CGZipFiles
            Set oZip = New CGZipFiles
            oZip.ZipFileName = Text1.Text
            Dim X As Integer
            For X = 1 To List1.ListCount
                oZip.AddFile List1.List(X)
            Next X
            If oZip.MakeZipFile <> 0 Then
               MsgBox oZip.GetLastMessage
            End If
            Set oZip = Nothing
            Unload Me
        End If
    End If
End Sub

Private Sub list1_Click()
    If List1.List(List1.ListIndex) <> "" Then
        Dim remove As Integer
        remove = MsgBox("Are You Sure You Want To Remove The File: " & List1.List(List1.ListIndex), vbYesNo, "Remove File")
        If remove = vbYes Then
            With List1
                .RemoveItem (List1.ListIndex)
            End With
        End If
    End If
End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    Dim yes As Integer
    yes = MsgBox("Are You Sure You Want To Add " & File1.FileName & " To The Zipfile: " & zipname, vbYesNo, "Add File")
    If yes = vbYes Then
        Dim X As Integer
        X = X + 1
        With List1
            .AddItem Dir1.path & "\" & File1.FileName
        End With
    End If
End Sub
