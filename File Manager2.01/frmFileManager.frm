VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileManager 
   Caption         =   "File Manager"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   11880
   Icon            =   "frmFileManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9551
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5415
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   9551
      _Version        =   393217
      Indentation     =   220
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10440
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9960
      Top             =   840
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0442
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0554
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0666
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0778
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":088A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":099C
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0AAE
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0BC0
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0CD2
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0DE4
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":0EF6
            Key             =   "Sort Ascending"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":1008
            Key             =   "Sort Descending"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":111A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":122E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":1686
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":1ADA
            Key             =   "back"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":1F2E
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":2382
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":27D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":2C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":2D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":2E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":2F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":307A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":35BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":36D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":47EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":5902
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":6A1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":7B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":7F86
            Key             =   "image"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6255
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4302
            MinWidth        =   4302
            Text            =   "Free Bytes: "
            TextSave        =   "Free Bytes: "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Number Of Files:"
            TextSave        =   "Number Of Files:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Selected   Bytes"
            TextSave        =   "Selected   Bytes"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4938
            MinWidth        =   4938
            Text            =   "Date and Time:"
            TextSave        =   "Date and Time:"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      DisabledImageList=   "imlToolbarIcons"
      HotImageList    =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Move File"
            Object.ToolTipText     =   "Move File"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Recycle"
            Object.ToolTipText     =   "Empty Recycle Bin"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            Object.ToolTipText     =   "Back Folder"
            ImageKey        =   "back"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            Object.ToolTipText     =   "Forward Folder"
            ImageKey        =   "forward"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sort Ascending"
            Object.ToolTipText     =   "Sort Ascending"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sort Descending"
            Object.ToolTipText     =   "Sort Descending"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8160
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":83DA
            Key             =   "dt"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":872E
            Key             =   "hd"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":8A82
            Key             =   "op"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":8DD6
            Key             =   "cl"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":912A
            Key             =   "rd"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":94FE
            Key             =   "fl"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":9852
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":9BE6
            Key             =   "rm"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":9F3A
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":A2CE
            Key             =   "dl"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":A622
            Key             =   "open"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":AA7A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":AED2
            Key             =   "cd1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":B32A
            Key             =   "adrive"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":B782
            Key             =   "harddrive"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":BBDA
            Key             =   "recycleempty"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":C032
            Key             =   "recyclefull"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":C48A
            Key             =   "transfer"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":C8E2
            Key             =   "mycomputer"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":CD3A
            Key             =   "desktop"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":D192
            Key             =   "controlpanel"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":D5E6
            Key             =   "netscape"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":E6FE
            Key             =   "printers"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "extracticon"
            Object.ToolTipText     =   "Extract Icon"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imageviewer"
            Object.ToolTipText     =   "Image Viewer"
            ImageKey        =   "image"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Internet Explorer"
            Object.ToolTipText     =   "Internet Explorer"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Notepad"
            Object.ToolTipText     =   "Notepad"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MS Paint"
            Object.ToolTipText     =   "MS Paint"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Wordpad"
            Object.ToolTipText     =   "Wordpad"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Begin VB.Menu mnuFolder 
            Caption         =   "&Folder"
         End
         Begin VB.Menu mnuShortcut 
            Caption         =   "&Shortcut"
         End
      End
      Begin VB.Menu mnuMoveTo 
         Caption         =   "&Move To"
      End
      Begin VB.Menu mnuCopyTo 
         Caption         =   "&Copy To"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Begin VB.Menu mnuDeleteFile 
            Caption         =   "&File"
         End
         Begin VB.Menu mnuDeleteFolder 
            Caption         =   "F&older"
         End
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "R&un"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Pri&nt"
         Begin VB.Menu mnuPrintFileList 
            Caption         =   "File List"
         End
         Begin VB.Menu mnuPrintFolderList 
            Caption         =   "Folder List"
         End
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "R&estart Computer"
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "&Shut Down Computer"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuline4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuShowAllFiles 
         Caption         =   "Show All Files"
      End
      Begin VB.Menu mnuLine68 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange"
         Begin VB.Menu mnuArrangeName 
            Caption         =   "Name"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuArrangeSize 
            Caption         =   "Size"
         End
         Begin VB.Menu mnuArrangeType 
            Caption         =   "Type"
         End
         Begin VB.Menu mnuArrangeLastModified 
            Caption         =   "Last Modified"
         End
      End
      Begin VB.Menu mnuLine38 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilters 
         Caption         =   "&Filters"
      End
   End
   Begin VB.Menu mnuGo 
      Caption         =   "&Go"
      Begin VB.Menu mnuRecentFilesTitle 
         Caption         =   "Recent Files"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuLine78 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSaveNow 
         Caption         =   "&Save Settings Now"
      End
      Begin VB.Menu mnuSaveExit 
         Caption         =   "S&ave Settings On Exit"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolBar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLaunchBar 
         Caption         =   "&Launch Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLine97 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowsOptions 
         Caption         =   "View Windows Options"
      End
      Begin VB.Menu mnuCheckBoxes 
         Caption         =   "Checkboxes"
         Begin VB.Menu mnuCHKOn 
            Caption         =   "On"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCHKOff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "&Grid"
         Begin VB.Menu mnuGridOn 
            Caption         =   "On"
         End
         Begin VB.Menu mnuGridOff 
            Caption         =   "Off"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuLine6997 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAscending 
         Caption         =   "&Ascending"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDescending 
         Caption         =   "&Descending"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuRegisterOCX 
         Caption         =   "&Register OCX"
      End
      Begin VB.Menu mnuUnregisterOCX 
         Caption         =   "&UnRegister OCX"
      End
      Begin VB.Menu mnuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtractIcon 
         Caption         =   "&Extract Icon From File"
      End
      Begin VB.Menu mnuLine8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatDisk 
         Caption         =   "F&ormat Disk"
      End
      Begin VB.Menu mnuLine9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecycleBin 
         Caption         =   "R&ecycle Bin"
         Begin VB.Menu mnuEmptyRecycleBin 
            Caption         =   "&Empty Recycle Bin"
         End
         Begin VB.Menu mnuRecycleBinProperties 
            Caption         =   "&Properties"
         End
      End
      Begin VB.Menu mnuLine10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageViewer 
         Caption         =   "&Image Viewer"
      End
      Begin VB.Menu mnuLine61 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenDOSWindow 
         Caption         =   "Open MS Dos Window"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh1 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "Properties"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "File Properties"
      End
      Begin VB.Menu mnuRenameFile 
         Caption         =   "&Rename File"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu mnuMoveFileTo 
         Caption         =   "&Move File To..."
      End
      Begin VB.Menu mnuLine69 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZip 
         Caption         =   "Zip"
      End
      Begin VB.Menu mnuUnzip 
         Caption         =   "UnZip"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintFList 
         Caption         =   "&Print File List"
      End
      Begin VB.Menu mnuLine73 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInvertAll 
         Caption         =   "&Invert All"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuChkBoxes 
         Caption         =   "CheckBoxes"
         Begin VB.Menu mnuTurnOn 
            Caption         =   "On"
         End
         Begin VB.Menu mnuTurnOff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu mnuRGrid 
         Caption         =   "&Grid"
         Begin VB.Menu mnuRGridOn 
            Caption         =   "On"
         End
         Begin VB.Menu mnuRGridOff 
            Caption         =   "Off"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuline49 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuCreateShortcut 
         Caption         =   "Create Shortcut"
      End
      Begin VB.Menu mnuLine13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRAscending 
         Caption         =   "Ascending"
      End
      Begin VB.Menu mnuRDescending 
         Caption         =   "Descending"
      End
   End
   Begin VB.Menu mnuRightClickFolder 
      Caption         =   "Properties"
      Visible         =   0   'False
      Begin VB.Menu mnuCreateDir 
         Caption         =   "Create Directory"
      End
      Begin VB.Menu mnuRRenameFolder 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuRDeleteFolder 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRFProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuline47 
         Caption         =   "-"
      End
      Begin VB.Menu mnuREmptyBin 
         Caption         =   "Empty Recycle Bin"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine87 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuLine48 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRZip 
         Caption         =   "Zip"
      End
      Begin VB.Menu mnuLine58 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuRB 
      Caption         =   "Recycle Bin"
      Visible         =   0   'False
      Begin VB.Menu mnuRBOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuPermDel 
         Caption         =   "Permanetly Delete"
      End
   End
End
Attribute VB_Name = "frmFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyrighted by: Mike Breault
'Date: February 25 2001
'Program Name: File Manager
'Program Purpose: Copy of File Manager/Explorer
'Note: Please report any bugs to me at - mike_breault@hotmail.com
'      Thank You for your co-operation!

Option Compare Text
Option Explicit

Private Const rDayZeroBias As Double = 109205#
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

'Const for max files in Recent Files List
Const MaxRecentFiles = 8

Private FullPath        As String
'Class for drive info
Private cDrv            As New CDriveInfo

'Api for Emptying Recycle Bin
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOPROGRESSUI = &H2

Private Const MAX_PATH = 260
Private Const ILD_TRANSPARENT = &H1

Private Type SHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Private ShInfo As SHFILEINFO

'Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Dim lngResult As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Function CurrToVbLocal(ByVal MyCurr As Currency) As Date
    CurrToVbLocal = (MyCurr / rMillisecondPerDay) - rDayZeroBias
End Function

Private Function LargeIntegerToDouble(low_part As Long, high_part As Long) As Double
    Dim result As Double
    result = high_part
    If high_part < 0 Then
        result = result + 2 ^ 32
    End If
    result = result * 2 ^ 32
    result = result + low_part
    If low_part < 0 Then
        result = result + 2 ^ 32
    End If
    LargeIntegerToDouble = result
End Function


Private Function SizeString(ByVal num_bytes As Double) As String
    Const SIZE_KB As Double = 1024
    Const SIZE_MB As Double = 1024 * SIZE_KB
    Const SIZE_GB As Double = 1024 * SIZE_MB
    Const SIZE_TB As Double = 1024 * SIZE_GB
    If num_bytes < SIZE_KB Then
        SizeString = Format$(num_bytes) & " bytes"
    ElseIf num_bytes < SIZE_MB Then
        SizeString = Format$(num_bytes / SIZE_KB, "0.00") & " KB"
    ElseIf num_bytes < SIZE_GB Then
        SizeString = Format$(num_bytes / SIZE_MB, "0.00") & " MB"
    Else
        SizeString = Format$(num_bytes / SIZE_GB, "0.00") & " GB"
    End If
End Function

Private Sub Form_Load()
    Let tbToolBar.Buttons(12).Enabled = False
    Let mnuTurnOn.Checked = True
    Let mnuCHKOn.Checked = True
    ListView1.View = lvwReport
    Let StatusBar1.Panels(2).Text = "Total Of: " & ListView1.ListItems.count & " File(s)"
    Dim path As String
    LoadTree6
    Initialise
    path = CurDir
    FillListView1WithFiles path
    'Displays Correct Path
    'Checks Free Space On Computer
    GetDiskInfo
    'Sizes The Controls On The Form
    Let mnuAscending.Checked = True
    Let mnuDescending.Checked = False
    Let mnuRAscending.Checked = True
    Let mnuRDescending.Checked = False
    If mnuAscending.Checked = True Then
        ListView1.SortOrder = 0
        ListView1.Sorted = True
    ElseIf mnuDescending.Checked = True Then
        ListView1.SortOrder = 1
        ListView1.Sorted = False
    End If
   ' Let ListView1.Width = (Me.Width - 2800)
   ' Let ListView1.Height = (Me.Height - 1856)
    Let Me.Caption = "File Manager   --->  " & SourcePath
    Dim maxim As Integer
    '0 = normal
    '1 = maximized
    maxim = GetSetting(App.Title, "Settings", "Maximized", 0)
    SourcePath = GetSetting(App.Title, "Settings", "Directory", CurDir)
    If maxim = 0 Then
        Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
        Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
        Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
        Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    ElseIf maxim = 1 Then
        Me.WindowState = vbMaximized
    End If
    Let ListView1.FullRowSelect = True
    Let StatusBar1.Panels(3).Text = "0 Selected" & " 0 Bytes"
End Sub

Private Sub Form_Resize()
    'Resizes Controls On The Form
    'If Form Gets To Small, Then Resize To Minimum
    If Me.WindowState <> 1 Then
        If Me.ScaleWidth < 7405 Then
            Me.ScaleWidth = 7405
        End If
        If Me.ScaleHeight < 6000 Then
            Me.ScaleHeight = 6000
        End If
        Dim ToolH
         
        ToolH = tbToolBar.Height + Toolbar1.Height
        TreeView1.Move 0, ToolH, 0.25 * Me.ScaleWidth, Me.ScaleHeight - ToolH - StatusBar1.Height
        ListView1.Move TreeView1.Width, ToolH, Me.ScaleWidth - TreeView1.Width, Me.ScaleHeight - ToolH - StatusBar1.Height

    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuSaveExit.Checked = True Then
        If Me.WindowState <> vbMinimized Then
            SaveSetting App.Title, "Settings", "MainLeft", Me.Left
            SaveSetting App.Title, "Settings", "MainTop", Me.Top
            SaveSetting App.Title, "Settings", "MainWidth", Me.Width
            SaveSetting App.Title, "Settings", "MainHeight", Me.Height
            SaveSetting App.Title, "Settings", "Maximized", 0
        ElseIf Me.WindowState = vbMaximized Then
            SaveSetting App.Title, "Settings", "Maximized", 1
        End If
        SaveSetting App.Title, "Settings", "Directory", SourcePath
        SaveSetting App.Title, "Settings", "ViewMode", ListView1.View
    End If
End Sub

Private Sub ListView1_Click()
    Dim count As Integer
    Dim yay As Integer
    Dim filesize As Single
    For count = 1 To ListView1.ListItems.count
        If ListView1.ListItems.Item(count).Checked = True Then
            yay = yay + 1
            filesize = Format(FileLen(ListView1.ListItems.Item(ListView1.ListItems.count)), "#,##0") + filesize
        End If
    Next count
       
    Let StatusBar1.Panels(3).Text = yay & " Selected  " & filesize & " Bytes"
End Sub

Private Sub ListView1_DblClick()
    'Loads Correct File When Double Clicked ON
    Me.MousePointer = vbHourglass
    If ListView1.SelectedItem.Text <> " " Then
        Dim FileName As String
        Dim A As Long
        FileName = SourcePath & ListView1.SelectedItem.Text
        If FileName = "" Then
            Exit Sub
        End If
        If mnuRecentFiles.count > MaxRecentFiles Then
            For A = 2 To mnuRecentFiles.UBound
                mnuRecentFiles(A - 1).Caption = mnuRecentFiles(A).Caption
            Next A
            Unload mnuRecentFiles(mnuRecentFiles.UBound)
        End If
        A = mnuRecentFiles.UBound + 1
        Load mnuRecentFiles(A)
        mnuRecentFiles(A).Caption = FileName
        If Right(FileName, 4) = ".zip" Then
            frmUnZip.Show vbModal, Me
        Else
            Call ShellExecute(hwnd, "Open", SourcePath & ListView1.SelectedItem.Text, "", App.path, 1)
        End If
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuCHKOff_Click()
    'Checks/Unchecks correct items in menus
    Let ListView1.Checkboxes = False
    Let mnuCHKOn.Checked = False
    Let mnuCHKOff.Checked = True
    Let mnuTurnOn.Checked = False
    Let mnuTurnOff.Checked = True
End Sub

Private Sub mnuCHKOn_Click()
    'Checks/Unchecks correct items in menus
    Let ListView1.Checkboxes = True
    Let mnuCHKOn.Checked = True
    Let mnuCHKOff.Checked = False
    Let mnuTurnOn.Checked = True
    Let mnuTurnOff.Checked = False
End Sub

Private Sub mnuCopyTo_Click()
    MsgBox "Copy File Code Goes Here", vbOKOnly, "Code"
End Sub

Private Sub mnuCreateDir_Click()
    'Calls create directory procedure
    Call mnuFolder_Click
End Sub

Private Sub mnuFilters_Click()
    'Displays the filters form
    frmFilters.Show vbModal, Me
End Sub

Private Sub mnuGridOff_Click()
    'Turns the grid off
    Let ListView1.GridLines = False
    Let mnuGridOn.Checked = False
    Let mnuGridOff.Checked = True
    Let mnuRGridOff.Checked = True
    Let mnuRGridOn.Checked = False
End Sub

Private Sub mnuGridOn_Click()
    'Turns the gridd on
    Let ListView1.GridLines = True
    Let mnuGridOn.Checked = True
    Let mnuGridOff.Checked = False
    Let mnuRGridOff.Checked = False
    Let mnuRGridOn.Checked = True
End Sub

Private Sub mnuInvertAll_Click()
    'Unchecks all checkboxes in listview
    CheckAllItems False
End Sub

Private Sub mnuMoveFileTo_Click()
    'Calls the move to procedure to move a file
    Call mnuMoveTo_Click
End Sub

Private Sub mnuPermDel_Click()
    'Destroys a file permanetly
    Dim yesno As Integer
    yesno = MsgBox("Are You Sure You Want To Delete File: " & SourcePath & ListView1.SelectedItem.Text & "?", vbYesNo, "Delete File")
    If yesno = vbYes Then
        Kill (SourcePath & ListView1.SelectedItem.Text)
    End If
    FillListView1WithFiles SourcePath
    ListView1.Refresh
End Sub

Private Sub mnuPrintFList_Click()
    'Calls procedure to print file list
    Call mnuPrintFileList_Click
End Sub

Private Sub mnuProperties_Click()
    'Calls Display Properties Sub
    Call mnuRFProperties_Click
End Sub

Private Sub mnuRBOpen_Click()
    'Opens correct file when user right clicks then chooses
    'open item
    If ListView1.SelectedItem.Text <> " " Then
        Call ShellExecute(hwnd, "Open", SourcePath & "\" & ListView1.SelectedItem.Text, "", App.path, 1)
    End If
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
    'Opens recent file user clicked on
    Call ShellExecute(hwnd, "Open", mnuRecentFiles(Index).Caption, "", App.path, 1)
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'displays the path and filename in me.caption and
    'tooltiptext for the name of the file if not all visible
    If Button = vbLeftButton Then ListView1.OLEDrag
    Dim l As ListItem
    Set l = ListView1.HitTest(x, y)
    If l Is Nothing Then Exit Sub
    Me.Caption = "File Manager   --->  " & SourcePath & l.Text
    ListView1.ToolTipText = SourcePath & l.Text
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Determines which popupmenu to use and what to display
    If Button <> vbRightButton Then
         Exit Sub
    End If
    
    If TreeView1.SelectedItem.Text = UCase("recycled") Then
        PopupMenu Me.mnuRB
    Else
        If LCase(Right(ListView1.SelectedItem.Text, 4)) = ".zip" Then
            mnuUnzip.Enabled = True
        Else
            mnuUnzip.Enabled = False
        End If
        PopupMenu Me.mnuRightClick
        ListView1.Refresh
    End If
End Sub

Private Function BuildFullPath(ByVal MyPath As String) As String
   On Error GoTo PROC_ERR
   Dim iPos As Integer

   iPos = InStr(MyPath, ":")
   If iPos < 2 Then
      Exit Function
   End If
   MyPath = Mid$(MyPath, iPos - 1)

   iPos = InStr(MyPath, "\")
   If iPos > 1 Then
      BuildFullPath = Left$(MyPath, 2) & Mid$(MyPath, iPos)
   Else
      BuildFullPath = Left$(MyPath, 2)
   End If

   BuildFullPath = BuildFullPath & "\"

PROC_EXIT:
  Exit Function
PROC_ERR:
  ErrMsgBox Me.Name & ".BuildFullPath"
  Resume Next

End Function

Private Sub mnuREmptyBin_Click()
    'Calls the empty recycle bin procedure
    Call mnuEmptyRecycleBin_Click
End Sub

Private Sub mnuRGridOff_Click()
    'Turns the grid off and checks and unchecks proper
    'items in menus
    Let ListView1.GridLines = False
    Let mnuGridOn.Checked = False
    Let mnuGridOff.Checked = True
    Let mnuRGridOff.Checked = True
    Let mnuRGridOn.Checked = False
End Sub

Private Sub mnuRGridOn_Click()
    'Turns the grid on and checks and unchecks proper
    'items in menus
    Let ListView1.GridLines = True
    Let mnuGridOn.Checked = True
    Let mnuGridOff.Checked = False
    Let mnuRGridOff.Checked = False
    Let mnuRGridOn.Checked = True
End Sub

Private Sub mnuSelectAll_Click()
    'Checks all checkboxes in listview
    CheckAllItems True
End Sub

Private Sub mnuShortcut_Click()
    'Displays create shortcut form
    Let frmCreateShortcut.txtExename.Text = SourcePath & ListView1.SelectedItem.Text
    Let frmCreateShortcut.txtShortcutName.Text = ListView1.SelectedItem.Text
    Let frmCreateShortcut.txtShortcutDir.Text = SourcePath
    frmCreateShortcut.Show vbModal, Me
    frmCreateShortcut.Refresh
    Call mnuRefresh1_Click
End Sub

Private Sub mnuShowAllFiles_Click()
    MsgBox "Show All Files Goes Here", vbOKOnly, "SHOW ALL FILES"
End Sub

Private Sub mnuTurnOff_Click()
    'Calls procedure to turn off grid
    Call mnuCHKOff_Click
End Sub

Private Sub mnuTurnOn_Click()
    'Calls procedure to turn on grid
    Call mnuCHKOn_Click
End Sub

Private Sub TreeView1_Click()
    'Determines which caption to use in heading
    If TreeView1.SelectedItem.Text = UCase("recycle bin") Then
        Let Me.Caption = "File Manager   --->  " & "Recycle Bin"
    ElseIf TreeView1.SelectedItem.Text = UCase("control panel") Then
        Let Me.Caption = "File Manager   --->  " & "Control Panel"
    ElseIf TreeView1.SelectedItem.Text = UCase("printers") Then
        Let Me.Caption = "File Manager   --->  " & "Printers"
    ElseIf TreeView1.SelectedItem.Text = UCase("desktop") Then
        Let Me.Caption = "File Manager   --->  " & "Desktop"
    Else
        Let Me.Caption = "File Manager   --->  " & SourcePath
    End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
   On Error GoTo PROC_ERR
   
   Screen.MousePointer = vbHourglass
   If Node.Children = 1 And Node.Child.Children <= 0 Then
       ' Remove the "dummy" item
       TreeView1.Nodes.remove Node.Child.Index
       ' Enumerate file system items under this node
       Node.Sorted = False
       EnumFilesUnder Node
       Node.Sorted = True
   End If

   Screen.MousePointer = vbDefault
   Let StatusBar1.Panels(2).Text = "Total Of: " & ListView1.ListItems.count & " File(s)"

PROC_EXIT:
  Exit Sub
PROC_ERR:
  ErrMsgBox Me.Name & ".TreeView1_Expand"
  Resume Next

End Sub
Private Sub EnumFilesUnder(ByVal n As Node)

   On Error GoTo PROC_ERR
    Dim sPath As String
    Dim hFind As Long
    Dim oldPath As String
    Dim wf As WIN32_FIND_DATA
    Dim n2 As Node
    Dim MyZip As Boolean
    Dim FolderPic As String

    TreeView1.Visible = False
    oldPath = ""
    sPath = BuildFullPath(n.FullPath) & "*.*"
    'old sPath = ucase$(n.FullPath & "\*.*")
    hFind = FindFirstFile(sPath, wf)
    Do
        ' Get the filename, if any.
        sPath = StripNull(wf.cFileName)
        If Len(sPath) = 0 Or StrComp(sPath, oldPath) = 0 Then
            ' Nothing found?
            Exit Do
        ElseIf Asc(sPath) <> 46 Then
            ' Add file with folder image

           If Right$(LCase$(sPath), 4) = ".zip" Then
              MyZip = True
           Else
              MyZip = False
           End If
           If UCase(sPath) = "recycled" Then
              Dim recycle As Boolean
              recycle = True
           Else
              recycle = False
           End If
           If (wf.dwFileAttributes And vbDirectory) Or MyZip Or recycle Then
              If MyZip Then
                 FolderPic = "zip"
              ElseIf recycle Then
                 FolderPic = "recycleempty"
              Else
                 FolderPic = "close"
              End If
              Set n2 = TreeView1.Nodes.Add(n, tvwChild, , sPath, FolderPic)
              If MyZip Then
                 TreeView1.Nodes.Item(TreeView1.Nodes.count).Bold = True
                 TreeView1.Nodes.Item(TreeView1.Nodes.count).ForeColor = vbBlue
                 TreeView1.Nodes.Item(TreeView1.Nodes.count).BackColor = vbWhite
              End If
              If Not (MyZip) Then
                 n2.ExpandedImage = "open"
                 ' Add a dummy item so the + sign is
                 ' displayed
                 If hasSubDirectory(BuildFullPath(n.FullPath) & sPath & "\") Then
                    TreeView1.Nodes.Add n2, tvwChild
                 End If

               End If
           End If
        End If
        FindNextFile hFind, wf
        oldPath = sPath
    Loop
    FindClose hFind
    TreeView1.Visible = True
    Exit Sub

PROC_EXIT:
  Exit Sub
PROC_ERR:
  ErrMsgBox Me.Name & ".EnumFilesUnder"
  Resume Next

End Sub

Private Function hasSubDirectory(ByVal sPath As String) As Boolean
    On Error GoTo PROC_ERR
    'function added by oigres P
    Dim hFind As Long
    Dim oldPath As String
    Dim wf As WIN32_FIND_DATA

    oldPath = ""
    hasSubDirectory = False 'assume false
    hFind = FindFirstFile(sPath & "*.*", wf)
    Do
        ' Get the filename, if any.
        sPath = StripNull(wf.cFileName)
        If Len(sPath) = 0 Or StrComp(sPath, oldPath) = 0 Then
            ' Nothing found?
            hasSubDirectory = False
            GoTo ExitFunction
        ElseIf Asc(sPath) <> 46 Then
            ' return true if we have found a directory under this path
            If (wf.dwFileAttributes And vbDirectory) Or Right$(LCase$(sPath), 4) = ".zip" Then
                hasSubDirectory = True
                GoTo ExitFunction
            End If
        End If
        FindNextFile hFind, wf
        oldPath = sPath
    Loop
ExitFunction:
    FindClose hFind

PROC_EXIT:
  Exit Function
PROC_ERR:
  ErrMsgBox Me.Name & ".hasSubDirectory"
  Resume Next

End Function

Private Sub LoadTree6()

   On Error GoTo PROC_ERR
   'For Vb6 control MsComCtl.OCX
   'We could use scripting here but let's go for speed
'------------------------------
   Dim FirstFixed  As Integer
   Dim NodeNum     As Integer
   Dim MaxPwr      As Integer
   Dim Pwr         As Integer
'------------------------------
   Dim DrvBitMask  As Long
'------------------------------
   Dim MyPic       As String
   Dim MyVol       As String
'------------------------------
   Dim nod1        As Node
'   Dim RC          As RECT
'------------------------------
  ' Set drv = New CDriveInfo
   TreeView1.ImageList = ImageList2    ' Initialize ImageList.
   
   '-- Not Active = 1
   '-- Removable  = 2
   '-- Fixed      = 3
   '-- Remote     = 4
   '-- CdRom      = 5
   '-- RamDisk    = 6

   Set nod1 = TreeView1.Nodes.Add(, , "mycmptr", "Desktop", "desktop")
   Set nod1 = TreeView1.Nodes.Add(, , "RDesktop", "My Computer", "mycomputer")
   Set nod1 = TreeView1.Nodes.Add(, , "CTRLPNL", "Control Panel", "controlpanel")
   Set nod1 = TreeView1.Nodes.Add(, , "PRNTR", "Printers", "printers")
   Set nod1 = TreeView1.Nodes.Add(, , "RCYCLBN", "Recycle Bin", "recycleempty")
   'Set nod1 = TreeView1.Nodes.Add(, , "WEBBRWSR", "Web Browser", "netscape")

   NodeNum = 3
   FirstFixed = 1

    DrvBitMask = GetLogicalDrives()
    ' DrvBitMask is a bitmask representing
    ' available disk drives. Bit position 0
    ' is drive A, bit position 2 is drive C, etc.
    ' If function fails, return value is zero.
    If DrvBitMask Then
     ' Get & search each available drive
       MaxPwr = Int(Log(DrvBitMask) / Log(2))   ' a little math...
       For Pwr = 0 To MaxPwr
          If 2 ^ Pwr And DrvBitMask Then
             cDrv.Drive = Chr$(65 + Pwr) & ":\"
             MyPic = Choose(cDrv.DriveType, "dl", "adrive", "harddrive", "transfer", "cd1", "rd")
             Select Case cDrv.DriveType
                Case 2, 3, 5, 6
                   If cDrv.Label = "" Then
                      MyVol = ""
                   Else
                      MyVol = "[" & cDrv.Label & "] "
                   End If
                   'MyVol = MyVol & cDrv.FormatSize(cDrv.AvailableSpace) & " " & GetResourceString(c1000)
                   If (FirstFixed = 1) And (cDrv.DriveType = 3) Then
                      FirstFixed = NodeNum + 1
                   End If
                Case Else
                   MyVol = ""
             End Select
             Set nod1 = TreeView1.Nodes.Add("RDesktop", tvwChild, "R" & Left(cDrv.Drive, 2) & TreeView1.Nodes.count, Left("(" & cDrv.Drive, 2) & MyVol & ":)", MyPic)
             TreeView1.Nodes.Add nod1, tvwChild
             NodeNum = NodeNum + 1
          End If
       Next
    End If

  ' Set nod1 = TreeView1.Nodes.Add("RDesktop", tvwChild, "CPL", "Control Panel", "cp")
  ' Set nod1 = TreeView1.Nodes.Add("RDesktop", tvwChild, "FTP", "FTP Natallink.com.br(dgs)", "rm")
   
   'ensure visible at drive level
   Set nod1 = TreeView1.Nodes(2) 'RDeskTop
   nod1.Expanded = True
   'expand first fixed drive
   Set nod1 = TreeView1.Nodes(FirstFixed)
   nod1.Expanded = True
   'ensure first entry (My Computer) is visible
   Set nod1 = TreeView1.Nodes(1)
   nod1.EnsureVisible
TreeView1.Refresh

PROC_EXIT:
  Exit Sub
PROC_ERR:
  ErrMsgBox Me.Name & ".LoadTree6"
  Resume Next

End Sub

Private Sub mnuAbout_Click()
    'Displays The About Form Onto Screen
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAscending_Click()
    'Sorts The ListView Alphabetically
    ListView1.SortOrder = 0
    ListView1.Sorted = True
    Let mnuAscending.Checked = True
    Let mnuDescending.Checked = False
    Let mnuRAscending.Checked = True
    Let mnuRDescending.Checked = False
End Sub

Private Sub mnuCreateShortcut_Click()
    'Calls mnushortcut_click to create a shortcut
    Call mnuShortcut_Click
End Sub

Private Sub mnuDeleteFile_Click()
    'Delete a single file
    Dim lResult As Long
    lResult = ShellDelete(SourcePath & ListView1.SelectedItem.Text)
    'Delete several files
    'lResult = ShellDelete("DELETE.ME", "LOVE_LTR.DOC", "COVERUP.TXT")
    Call mnuRefresh1_Click
End Sub

Private Sub mnuDeleteFolder_Click()
    MsgBox "Code Goes Here", vbOKOnly, "Code"
End Sub

Private Sub mnuDescending_Click()
    'Sorts The ListView Non-Alphabetically
    ListView1.SortOrder = 1
    ListView1.Sorted = True
    Let mnuAscending.Checked = False
    Let mnuDescending.Checked = True
    Let mnuRAscending.Checked = False
    Let mnuRDescending.Checked = True
End Sub

Private Sub mnuEmptyRecycleBin_Click()
    'Empties recycle bin then refreshs listview
    Dim retvaL
    retvaL = SHEmptyRecycleBin(frmFileManager.hwnd, "", SHERB_NOPROGRESSUI)
    Call mnuRefresh1_Click
End Sub

Private Sub mnuExit_Click()
    'Ends Program
    End
End Sub

Private Sub mnuExtractIcon_Click()
    'Displays the extract icon form
    frmExtractIcon.Show vbModal, Me
End Sub

Private Sub mnuFind_Click()
    'Displays The Find form
    frmFind.Show vbModal, Me
End Sub

Private Sub mnuFolder_Click()
    'Displays a InputBox and allows creation of Folder
    Dim secAtts As SECURITY_ATTRIBUTES
    Dim newfolder As String
    Dim tmpdir As String
    newfolder = InputBox("Type the name of the new folder", "New Folder")
    If Right$(SourcePath, 1) = "\" Then
        tmpdir = SourcePath & newfolder
    Else
        tmpdir = SourcePath & "\" & newfolder
    End If
    'MsgBox tmpDir
    CreateDirectory tmpdir, secAtts
End Sub

Private Sub mnuFormatDisk_Click()
    'Displays form for formatting disk
    Dim DriveLetter$, DriveNumber&, DriveType&
    Dim retvaL&, RetFromMsg%
    DriveLetter = UCase("A:\")
    DriveNumber = (Asc(DriveLetter) - 65)
    DriveType = GetDriveType(DriveLetter)
    If DriveType = 2 Then
        retvaL = SHFormatDrive(Me.hwnd, DriveNumber, 0&, 0&)
    Else
        MsgBox "Will Not Format Any Drive Except A:\", vbOKOnly, "Warning"
    End If
End Sub

Private Sub mnuImageViewer_Click()
    'Displays form for image viewing
    frmImageViewer.Show vbModal, Me
End Sub

Private Sub mnuLaunchBar_Click()
    'Hides or unhides launchbar and checks item in menus
    'appropiately
    If mnuLaunchBar.Checked = True Then
        Let mnuLaunchBar.Checked = False
        Let Toolbar1.Visible = False
    ElseIf mnuLaunchBar.Checked = False Then
        Let mnuLaunchBar.Checked = True
        Let Toolbar1.Visible = True
    End If
End Sub

Private Sub mnuMoveTo_Click()
    'Displays the moveto form sending it the file name and
    'then refreshes the listview
    Let frmMoveFile.lblCaption.Caption = "Move File: " & SourcePath & ListView1.SelectedItem.Text & " To Directory:"
    frmMoveFile.Refresh
    frmMoveFile.Show vbModal, Me
    Call mnuRefresh1_Click
End Sub

Private Sub mnuNewWindow_Click()
    'Opens Another File Manager Window
    Call ShellExecute(hwnd, "Open", App.path & "\" & "FileManager.exe", "", App.path, 1)
End Sub

Private Sub mnuOpenFile_Click()
    'opens appropriate file if a file is clicked on
    If ListView1.SelectedItem.Text <> " " Then
        Call ShellExecute(hwnd, "Open", SourcePath & "\" & ListView1.SelectedItem.Text, "", App.path, 1)
    End If
End Sub

Private Sub mnuPrintFileList_Click()
    'Prints the listview list/report
    PrintLW ListView1
End Sub

Private Sub mnuRecycleBinProperties_Click()
    'Displays the recycle bin properties
    Dim r As Long
    r = ShowFileProperties("C:\Recycled\", Me.hwnd) 'To show the properties dialog pass the filename and the owner of the dialog
    If r <= 32 Then
        MsgBox "Error In Properties Window", vbOKOnly, "Error"
    End If
    MousePointer = vbDefault
End Sub

Private Sub mnuRefresh_Click()
    'Calls procedure to refresh listview
    Call mnuRefresh1_Click
End Sub

Private Sub mnuRefresh1_Click()
    'Fills listview with new files with updates included
    FillListView1WithFiles SourcePath
    ListView1.Refresh
    'Still wondering on how to "update" treeview, any help?
End Sub

Private Sub mnuRename_Click()
    'renames a file then refreshes listview
    Dim renamedFile As String
    renamedFile = InputBox("Enter The New Name For " & ListView1.SelectedItem.Text & ":", "Rename File")
    If renamedFile <> "" Then
        Name SourcePath & "\" & ListView1.SelectedItem.Text As SourcePath & "\" & renamedFile
    Else
        Exit Sub
    End If
    Call mnuRefresh1_Click
End Sub

Private Sub mnuRenameFile_Click()
    'Calls rename procedure to rename a file
    Call mnuRename_Click
End Sub

Private Sub mnuRestart_Click()
    'Restarts The Computer
    Dim leave As Integer
    leave = MsgBox("Are You Sure You Want To Restart The Computer?", vbYesNo, "Restart Computer")
    If leave = vbYes Then
        lngResult = ExitWindowsEx(EWX_REBOOT, 0&)
    End If
End Sub

Private Sub mnuRFProperties_Click()
    'Displays a files properties when popupmenu is clicked
    'and the properties item is choosen
    Dim r As Long
    r = ShowFileProperties(SourcePath, Me.hwnd) 'To show the properties dialog pass the filename and the owner of the dialog
    If r <= 32 Then
        MsgBox "Error In Properties Window", vbOKOnly, "Error"
    End If
    MousePointer = vbDefault
End Sub

Private Sub mnuRRenameFolder_Click()
    'Renames a folder accordingly
    Dim renamedFolder As String
    renamedFolder = InputBox("Enter The New Name For " & SourcePath, "Rename Folder")
    If renamedFolder <> "" Then
        Name SourcePath As SourcePath & renamedFolder
    Else
        Exit Sub
    End If
End Sub

Private Sub mnuRun_Click()
    'Displays The Form Run For Locating Applications To Run
    frmRun.Show vbModal, Me
End Sub

Private Sub mnuSaveExit_Click()
    'Checks/unchecks item to save on exit
    Let mnuSaveExit.Checked = True
    Let mnuSaveNow.Checked = False
End Sub

Private Sub mnuSaveNow_Click()
    'Checks/unchecks item to save now
    Let mnuSaveNow.Checked = True
    Let mnuSaveExit.Checked = False
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    ElseIf Me.WindowState = vbMaximized Then
        SaveSetting App.Title, "Settings", "Maximized", 1
    End If
    SaveSetting App.Title, "Settings", "Directory", SourcePath
    SaveSetting App.Title, "Settings", "ViewMode", ListView1.View
End Sub

Private Sub mnuShutdown_Click()
    'Shuts Down The Computer
    Dim leave As Integer
    leave = MsgBox("Are You Sure You Want To Shut Down The Computer?", vbYesNo, "Shut Down Computer")
    If leave = vbYes Then
        lngResult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
    End If
End Sub

Public Sub GetDiskInfo()
    'Reads from computer the free space
    Dim root As String
    Dim volume_name As String
    Dim serial_number As Long
    Dim max_component_length As Long
    Dim file_system_flags As Long
    Dim file_system_name As String
    Dim pos As Integer
    Dim bytes_avail As LARGE_INTEGER
    Dim bytes_total As LARGE_INTEGER
    Dim bytes_free As LARGE_INTEGER
    Dim dbl_total As Double
    Dim dbl_free As Double
    On Error Resume Next
    GetDiskFreeSpaceEx Left(SourcePath, 3), bytes_avail, bytes_total, bytes_free
    dbl_total = LargeIntegerToDouble(bytes_total.lowpart, bytes_total.highpart)
    dbl_free = LargeIntegerToDouble(bytes_free.lowpart, bytes_free.highpart)
    StatusBar1.Panels.Item(4).Text = Date & "  " & Time
    StatusBar1.Panels.Item(1).Text = Left(SourcePath, 2) & " " & SizeString(dbl_total) & " (" & SizeString(dbl_free) & " Free)"
End Sub

Private Sub mnuStatusBar_Click()
    'Hides/unhides the statusbar
    If mnuStatusBar.Checked = True Then
        Let mnuStatusBar.Checked = False
        Let StatusBar1.Visible = False
    ElseIf mnuStatusBar.Checked = False Then
        Let mnuStatusBar.Checked = True
        Let StatusBar1.Visible = True
    End If
End Sub

Private Sub mnuToolbar_Click()
    'Hides/unhides the toolbar
    If mnuToolBar.Checked = True Then
        Let mnuToolBar.Checked = False
        Let tbToolBar.Visible = False
    ElseIf mnuToolBar.Checked = False Then
        Let mnuToolBar.Checked = True
        Let tbToolBar.Visible = True
    End If
End Sub

Private Sub munExtractIcon_Click()
    'Displays the extract icon form
    frmExtractIcon.Show vbModal, Me
End Sub

Private Sub mnuUnzip_Click()
    'Displays the unzip form
    frmUnZip.Show vbModal, Me
End Sub

Private Sub mnuWindowsOptions_Click()
    'Displays the windows options form
    frmWindowsOptions.Show vbModal, Me
End Sub

Private Sub mnuZip_Click()
    'Displays the zip form
    frmZip.Show vbModal, Me
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Sort Ascending"
            'Sorts listview ABC
            Let mnuAscending.Checked = True
            Let mnuDescending.Checked = False
            Let ListView1.SortOrder = 0
        Case "Sort Descending"
            'Sorts listview CBA
            Let mnuAscending.Checked = False
            Let mnuDescending.Checked = True
            Let ListView1.SortOrder = 1
        Case "Find"
            'Displays find form
            frmFind.Show vbModal, Me
        Case "Properties"
            'displays file properties
            Call mnuFileProperties_click
        Case "Delete"
            'Calls delete procedure for a file
            Call mnuDeleteFile_Click
        Case "Recycle"
            'Calls empty recycle bin procedure
            Call mnuEmptyRecycleBin_Click
        Case "Cut"
            MsgBox "Code Goes Here", vbOKOnly, "Code"
        Case "Copy"
            MsgBox "Code Goes Here", vbOKOnly, "Code"
        Case "Paste"
            MsgBox "Code Goes Here", vbOKOnly, "Code"
        Case ""
        Case ""
    End Select
End Sub

Private Sub mnuFileProperties_click()
    'Display File Properties
    Dim showProperties As Long
    showProperties = ShowFileProperties(SourcePath & "\" & ListView1.SelectedItem.Text, Me.hwnd)
    If showProperties <= 32 Then
        MsgBox "Error"
        MousePointer = vbDefault
        Exit Sub
    End If
End Sub

Public Sub mnuFileDelete_click()
    'Deletes The File Highlighted
    Call mnuDeleteFile_Click
End Sub

Public Sub mnuRAscending_click()
    'Calls ascending procedure for listview
    Call mnuAscending_Click
End Sub

Public Sub mnuRDescending_click()
    'Calls descending procedure for listview
    Call mnuDescending_Click
End Sub

Private Sub Timer2_Timer()
    'Updates the time on the statusbar every 1 second
    StatusBar1.Panels.Item(4).Text = Format(Date, "Long Date") & "  " & Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Internet Explorer"
            'Opens program internet explorer
            Call ShellExecute(hwnd, "Open", "C:\Program Files\Iexplore.exe", "", App.path, 1)
        Case "Notepad"
            'Opens program Notepad
            Call ShellExecute(hwnd, "Open", "C:\Windows\Notepad.exe", "", App.path, 1)
        Case "MS Paint"
            'Opens program MS Paint
            Call ShellExecute(hwnd, "Open", "C:\Windows\Pbrush.exe", "", App.path, 1)
        Case "Wordpad"
            'Opens program Wordpad
            Call ShellExecute(hwnd, "Open", "C:\Program Files\Accessories\Wordpad.exe", "", App.path, 1)
        Case "extracticon"
            'Displays extract icon form
            Call mnuExtractIcon_Click
        Case "imageviewer"
            'Displays Image Viewer Form
            Call mnuImageViewer_Click
        Case ""
        Case ""
        Case ""
    End Select
End Sub

Private Sub LoadZip2()
   On Error GoTo PROC_ERR
    Dim Ratio As Single
    Dim Sig As Long, MyCount As Long
    Dim ZipStream As Integer
    Dim Res As Long, Location  As Long
    Dim zFile As ZipFile
    Dim Name As String, FileName As String, ext As String
    Dim FTime As Currency, UTC As Currency
    Dim itmx As ListItem
    Dim DateTime As Date
    
    '--------------
    '--Load Listview Headers (Zip)
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.Visible = False
    ListView1.ColumnHeaders.Add , "nam", "Name", 1800
    ListView1.ColumnHeaders.Add , "ext", "Ext", 500
    ListView1.ColumnHeaders.Add , "siz", "Size", 1200, 1
    ListView1.ColumnHeaders.Add , "typ", "Type", 1000
    ListView1.ColumnHeaders.Add , "cmp", "Comp", 1200, 1
    ListView1.ColumnHeaders.Add , "rat", "Ratio", 800
    ListView1.ColumnHeaders.Add , "dat", "Date", 1000
    ListView1.ColumnHeaders.Add , "tim", "Time", 1200
    ListView1.ColumnHeaders.Add , "mtd", "Method", 1000
    ListView1.ColumnHeaders.Add , "enc", "Encoded", 1000
    ListView1.ColumnHeaders.Add , "pth", "Path", 2000

    '--------------
    
   ' ProgressInit 83
   ' InZip = True
   ' InFtp = False
   ' GridColumnHeaders
   ' LoadSpecialFiles
    frmFileManager.Caption = SourcePath & "\" 'was sFile
    
    ZipStream = FreeFile
    Open SourcePath For Binary As ZipStream
    Do
        Get ZipStream, , Sig
        'See if the file header has been found
            If Sig = LocalFileHeaderSig Then
                Get ZipStream, , zFile.Version
                Get ZipStream, , zFile.Flag
                Get ZipStream, , zFile.CompressionMethod
                Get ZipStream, , zFile.Time
                Get ZipStream, , zFile.Date
                Get ZipStream, , zFile.CRC32
                Get ZipStream, , zFile.CompressedSize
                Get ZipStream, , zFile.UncompressedSize
                Get ZipStream, , zFile.FileNameLength
                Get ZipStream, , zFile.ExtraFieldLength
                Name = String$(zFile.FileNameLength, " ")
                Get ZipStream, , Name
                zFile.FileName = Mid$(Name, 1, zFile.FileNameLength)
                Seek ZipStream, (Seek(ZipStream) + zFile.ExtraFieldLength)
                Seek ZipStream, (Seek(ZipStream) + zFile.CompressedSize)
                '-----------------
                MyCount = MyCount + 1
                Location = InStrRev(Name, "/", -1)
                FileName = Mid$(Name, Location + 1)
                ext = GetExt(FileName)
                Set itmx = ListView1.ListItems.Add(, , FileName)
                ext = GetExt(FileName)
                itmx.SubItems(1) = ext
                itmx.SubItems(2) = Format(zFile.UncompressedSize, "#,###")
                itmx.SubItems(3) = GetFileType(ext)
                itmx.SubItems(4) = Format(zFile.CompressedSize, "#,###")
                'Trap division by zero
                If zFile.UncompressedSize <> 0 Then
                    Ratio = 1 - zFile.CompressedSize / zFile.UncompressedSize
                Else
                    Ratio = 0
                End If
                'Ratio is single. Format as desired
                itmx.SubItems(5) = Format(Ratio, "00.0%")
                DateTime = GetDateTime(zFile.Date, zFile.Time)
                itmx.SubItems(6) = Format(DateTime, "Short Date")
                itmx.SubItems(7) = Format(DateTime, "Long Time")
                'Flag bits 1,2 hold additional Method info
                itmx.SubItems(8) = MethodVerbose(zFile.CompressionMethod, zFile.Flag)
                'Flag bit 0 is Encryption True/False
                If zFile.Flag And 1 Then
                   itmx.SubItems(9) = "True"
                Else
                   itmx.SubItems(9) = "False"
                End If
                itmx.SubItems(10) = Left$(Name, Location)
                'Save the item number for other operations
                itmx.Tag = MyCount
                
                Dim TotalSize As Single
                TotalSize = TotalSize + zFile.UncompressedSize
                
            Else
                Select Case Sig
                   Case 0, CentralFileHeaderSig, EndCentralDirSig
                      Exit Do
                End Select
            End If
        Loop
        Close ZipStream
        ListView1.Visible = True
                

PROC_EXIT:
  Exit Sub
PROC_ERR:
  ErrMsgBox Me.Name & ".LoadZip2"
  Resume Next
End Sub

Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
    'Converts the file date/time dos stamp from the archive
    'in to a normal date/time string
    Dim r As Long
    Dim FTime As Currency 'Makes it much easier to convert
    'Convert the dos stamp into a file time
    r = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    'Filetime to VbDate
    GetDateTime = CurrToVbLocal(FTime)
End Function

Private Function MethodVerbose(ByVal Method As Integer, ByVal BitFlag As Integer) As String
   On Error Resume Next
   'Conforms to PkZip 2.04g Specifications
'Methods are
'0    Stored (None)
'1    Shrunk
'2-5  Reduced:1,2,3,4
'(For Method 6 - Imploding)
' general purpose bit flag: (2 bytes)
'Bit 1: If the compression method used was type 6,
'       Imploding, then this bit, if set, indicates
'       an 8K sliding dictionary was used.  If clear,
'       then a 4K sliding dictionary was used.
'Bit 2: If the compression method used was type 6,
'       Imploding, then this bit, if set, indicates
'       3 Shannon-Fano trees were used to encode the
'       sliding dictionary output.  If clear, then 2
'       Shannon-Fano trees were used.
'6    Imploded:8kDict/4kDict:3Tree/2Tree
'7    Tokenized
'(For Method 8 - Deflating)
' general purpose bit flag: (2 bytes)
'Bit 2  Bit 1
'  0      0    Normal (-en) compression option was used.
'  0      1    Maximum (-ex) compression option was used.
'  1      0    Fast (-ef) compression option was used.
'  1      1    Super Fast (-es) compression option was used.
'8    Deflated:N,X,F,S
'9    EnhDefl
'10   ImplDCL
'Else Unknown

   Select Case Method
      Case 0
         MethodVerbose = "Stored (None)"
      Case 1
         MethodVerbose = "Shrunk"
      Case 2 To 5
         MethodVerbose = "Reduced:" & Method - 1
      Case 6
         MethodVerbose = "Imploded:" & Choose((BitFlag \ 2) + 1, "8KDict:2Tree", "4KDict:2Tree", "8KDict:3Tree", "4KDict:3Tree")
      Case 7
         MethodVerbose = "Tokenized"
      Case 8
         MethodVerbose = "Deflated:" & Choose((BitFlag \ 2) + 1, "N", "X", "F", "S")
      Case 9
         MethodVerbose = "EnhDef"
      Case 10
         MethodVerbose = "ImplDCL"
      Case Else
         MethodVerbose = "Unknown"
   End Select

End Function

Sub FillListView1WithFiles(ByVal path As String)
    Const strAttr As String = "rhsvda"
    Dim Item As ListItem
    Dim s As String, sAttr As String
    Dim Attr As Integer, L4 As Integer
    path = CheckPath(path)    'Add '\' to end if not present
    SourcePath = path 'Global with '\'
    '--------------
    '--Load Listview Headers (will change for Zip)
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.Visible = False
    ListView1.ColumnHeaders.Add , "nam", "Name:", 2100
    ListView1.ColumnHeaders.Add , "xtn", "Ext:", 500
    ListView1.ColumnHeaders.Add , "siz", "Size:", 1160, 1
    ListView1.ColumnHeaders.Add , "typ", "Type:", 2100
    ListView1.ColumnHeaders.Add , "dat", "Date:", 1100
    ListView1.ColumnHeaders.Add , "tim", "Time:", 1100
    ListView1.ColumnHeaders.Add , "atr", "Attr:", 640
    
    '--------------
    s = Dir(path, vbNormal)
    Do While s <> ""
      If Asc(s) <> 46 Then
         Set Item = ListView1.ListItems.Add()
         Item.Key = path & s
         'Item.SmallIcon = "Folder"
         Dim ext As String
         ext = GetExt(path & s)
         Item.Text = s
         Item.SubItems(1) = ext
         Item.SubItems(2) = Format(FileLen(Item.Key), "#,##0")
         Item.SubItems(4) = Format(FileDateTime(Item.Key), "Short Date")
         Item.SubItems(5) = Format(FileDateTime(Item.Key), "Long Time")
         Attr = GetAttr(Item.Key)
         sAttr = "......"
         For L4 = 0 To 5
            If Attr And 2 ^ L4 Then
               Mid(sAttr, L4 + 1, 1) = Mid(strAttr, L4 + 1, 1)
            End If
         Next
         Item.SubItems(6) = sAttr
         If Attr And vbDirectory Then
           Item.SubItems(3) = "<dir>"
         Else
           Item.SubItems(3) = GetFileType(GetExt(s))
         End If
      End If
      s = Dir
    Loop
    
    ListView1.Visible = True
    
End Sub

Private Function GetFileType(ByVal sExt As String) As String
   On Error GoTo PROC_ERR
   Dim sName2 As String
   Dim lRegKey As Long  'Registry Key

   If sExt <> "" Then
      If RegOpenKey(HKEY_CLASSES_ROOT, ByVal "." & sExt, lRegKey) = 0 Then
         RegQueryValueEx lRegKey, ByVal "", 0&, 1, ByVal Buffer, OFS_MAXPATHNAME
         sName2 = StripNull(Buffer)
         RegCloseKey lRegKey
         If Len(sName2) Then
            'get type
            If RegOpenKey(HKEY_CLASSES_ROOT, sName2, lRegKey) = 0 Then
               RegQueryValueEx lRegKey, ByVal "", 0&, 1, ByVal f_Type, 80
               GetFileType = StripNull(f_Type)
               RegCloseKey lRegKey
            End If
         End If
      End If
   End If
    
   If GetFileType = "" Then
      GetFileType = UCase$(sExt) & " File"
   End If
    
PROC_EXIT:
  Exit Function

PROC_ERR:
  ErrMsgBox Me.Name & ".GetFileType"
  Resume Next

End Function

Private Function GetExt(ByVal Name As String) As String
   On Error GoTo PROC_ERR
   Dim j As Integer
   j = InStrRev(Name, ".")
   If j > 0 And j < (Len(Name) - 1) Then
      GetExt = LCase$(Mid$(Name, j + 1))
   End If

PROC_EXIT:
  Exit Function
PROC_ERR:
  ErrMsgBox Me.Name & ".GetExt"
  Resume Next

End Function

Private Function CheckPath(ByVal path As String) As String
    If Right(path, 1) <> "\" Then
      CheckPath = path & "\"
    Else
      CheckPath = path
    End If
End Function

Private Sub Initialise()
    On Local Error Resume Next
    'Break the link to iml lists
    ListView1.ListItems.Clear
    ListView1.Icons = Nothing
    ListView1.SmallIcons = Nothing
    'Clear the image lists
End Sub

Public Function PrintLW(Lw As ListView)
    'Prints the listview to printer
    Dim sData As String
    Dim x As Integer
    Dim sIdx As Integer
    sIdx = Lw.ColumnHeaders.count - 1
    On Error Resume Next
    Printer.Print Tab(4), "Path Of:  " & SourcePath & "\"
    Printer.Print
    Dim i As Integer
    For i = 1 To Lw.ListItems.count
        sData = Lw.ListItems.Item(i).Text
        For x = 1 To sIdx
            sData = sData
        Next
        Printer.Print sData
        sData = ""
    Next
    Printer.EndDoc
    Lw.ListItems.Clear
    MsgBox "Done!!"
End Function

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Determines which popupmenu to display on treeview
    If Button <> vbRightButton Then
         Exit Sub
    End If
    If TreeView1.SelectedItem.Text = UCase("recycled") Then
        Let mnuREmptyBin.Enabled = True
        Let mnuRZip.Enabled = False
        Let mnuCreateDir.Enabled = False
        Let mnuRRenameFolder.Enabled = False
        Let mnuRDeleteFolder.Enabled = False
        PopupMenu Me.mnuRightClickFolder
    ElseIf TreeView1.SelectedItem.Text = UCase("recycle bin") Then
        Let mnuREmptyBin.Enabled = True
        Let mnuRZip.Enabled = False
        Let mnuCreateDir.Enabled = False
        Let mnuRRenameFolder.Enabled = False
        Let mnuRDeleteFolder.Enabled = False
        PopupMenu Me.mnuRightClickFolder
    ElseIf TreeView1.SelectedItem.Text = UCase("control panel") Then
    ElseIf TreeView1.SelectedItem.Text = UCase("printers") Then
    Else
        Let mnuREmptyBin.Enabled = False
        Let mnuRZip.Enabled = True
        Let mnuCreateDir.Enabled = True
        Let mnuRRenameFolder.Enabled = True
        Let mnuRDeleteFolder.Enabled = True
        PopupMenu Me.mnuRightClickFolder
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo PROC_ERR
    Dim Nodx  As Node, temp As String
    '-- Set the variable to the SelectedItem.
    Set Nodx = TreeView1.SelectedItem
    FullPath = BuildFullPath(Nodx.FullPath)
    temp = UCase$(Right$(FullPath, 5))
    
    If temp = ".ZIP\" Then
        SourcePath = Left$(FullPath, Len(FullPath) - 1)
        LoadZip2
    ElseIf Nodx.Key = "RDesktop" Then
        'Nothing
    ElseIf Nodx.Key = "CTRLPNL" Then
        ListView1.ListItems.Clear
        ListView1.ColumnHeaders.Clear
        ListView1.Visible = False
        ListView1.ColumnHeaders.Add , "nam", "Name", 3500
        'SourcePath = SysDir
        'Filter = "*.cpl"
        'FillListView1WithFiles SourcePath
        ListView1.Visible = True
    ElseIf Nodx.Key = "PRNTR" Then
        ListView1.ListItems.Clear
        ListView1.ColumnHeaders.Clear
        ListView1.Visible = False
        ListView1.ColumnHeaders.Add , "nam", "Name", 3500
        ListView1.Visible = True
    ElseIf Nodx.Key = "RCYCLBN" Then
        FillListView1WithFiles UCase("c:\recycled")
        ListView1.Refresh
    Else
        SourcePath = FullPath
        FillListView1WithFiles SourcePath
    End If
    Let StatusBar1.Panels(2).Text = "Total Of: " & ListView1.ListItems.count & " File(s)"
   
PROC_EXIT:
    Exit Sub
PROC_ERR:
    If SourcePath = UCase("a:\") Then
        MsgBox "Insert Disk In A:\ First then Continue", vbOKOnly, "Error:"
    End If
    'ErrMsgBox Me.Name & ".TreeView1_NodeClick"
    Resume Next
End Sub

Private Function fixed(ByVal path As String) As String
    fixed = path & IIf(Right(path, 1) = "\", "", "\")
End Function

Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    'Allows files to be dragged to other programs like:
    'explorer, file manager, this file manager, start menu
    'etc.
    AllowedEffects = vbDropEffectCopy
    Data.Clear
    Data.Files.Add ListView1.SelectedItem.Key
    Data.SetData , vbCFFiles
End Sub

Private Sub CheckAllItems(bState As Boolean)
    'Checks all file checkboxes
    Dim LV As LVITEM
    Dim lvCount As Long
    Dim lvIndex As Long
    Dim lvState As Long
    Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = ListView1.ListItems.count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .state = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(ListView1.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    'Sorts the files according to what user selected
    Dim currSortKey As Integer
    ListView1.SortKey = ColumnHeader.Index - 1
    currSortKey = ListView1.SortKey
    ListView1.SortOrder = Abs(Not ListView1.SortOrder = 1)
    ListView1.Sorted = True
    mnuAscending.Checked = ListView1.SortOrder = 0
    mnuDescending.Checked = mnuAscending.Checked = False
End Sub
