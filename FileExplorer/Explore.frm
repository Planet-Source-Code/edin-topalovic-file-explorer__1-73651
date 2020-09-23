VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplore 
   Caption         =   "Explorer  - Topka ® Software"
   ClientHeight    =   8685
   ClientLeft      =   1770
   ClientTop       =   1950
   ClientWidth     =   12300
   ClipControls    =   0   'False
   Icon            =   "Explore.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8685
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList tb1 
      Left            =   2160
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":2854
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":33AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":3F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":4A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":55BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":6116
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":6C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":77CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":8324
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":8E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":99D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":A532
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":B08C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer enb_ 
      Interval        =   1
      Left            =   0
      Top             =   6120
   End
   Begin VB.ListBox naprijed 
      Height          =   1035
      Left            =   5400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox nazad 
      Height          =   1035
      Left            =   3600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.Toolbar tbr 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   1482
      ButtonWidth     =   1588
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "tb1"
      DisabledImageList=   "tb1dis"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "topka"
                  Text            =   "Edin"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "utio"
                  Text            =   "utio"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Up"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Move to"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy to"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cut"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Paste"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            ImageIndex      =   14
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.PictureBox pictemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   555
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSComctlLib.ImageList ilSmall 
      Left            =   600
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":BBE6
            Key             =   "fldrClosed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMain 
      Left            =   0
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":C180
            Key             =   "fldrClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":C71A
            Key             =   "fldrOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":CCB4
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":D24E
            Key             =   "explorer"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilMain"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   8400
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16978
            MinWidth        =   2558
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "14:50"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "25.12.2010"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   5175
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      View            =   3
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "il1"
      SmallIcons      =   "il2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ime"
         Text            =   "Ime"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "velièina"
         Text            =   "Zauzeta memorija"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "tip"
         Text            =   "Tip"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "kreirano"
         Text            =   "Kreirano"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "mijenjano"
         Text            =   "Mijenjano"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "otvarano"
         Text            =   "Otvarano"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "osobine"
         Text            =   "Atributi"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList il2 
      Left            =   480
      Top             =   7320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList il1 
      Left            =   0
      Top             =   7320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList il5 
      Left            =   9000
      Top             =   6960
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   26
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":D7E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il4 
      Left            =   8520
      Top             =   6960
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":E05C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":E3F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":E790
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":EB2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":EEC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":F25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":F5F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":F992
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":FD2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il3 
      Left            =   8040
      Top             =   6960
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":100C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":10460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":107FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":10B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":10F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":112C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":11662
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":119FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":11D96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il7 
      Left            =   1440
      Top             =   7320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList il6 
      Left            =   960
      Top             =   7320
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList tb1dis 
      Left            =   2760
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":12130
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":12C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":137E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":1433E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":14E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":159F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":1654C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":170A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":17C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":1875A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":192B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":19E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":1A968
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explore.frx":1B4C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu osvjezi 
      Caption         =   "Refresh"
   End
   Begin VB.Menu optabout 
      Caption         =   "About"
   End
   Begin VB.Menu mnu_FileExit 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "frmExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public fls As Files, fl As File
Public fldrs As Folders, fldr As Folder
Public fldrs1 As Folders, fldr1 As Folder
Public drv As Drive, drvs As Drives
Public fso As FileSystemObject
Public duse As Boolean
Private mem_ln, mem_nn As String

Private mIsMoving As Boolean
Const mSplitLimit = 2000

Private Sub enb__Timer()
On Error Resume Next
If nazad.ListCount > 1 Then
    tbr.Buttons(1).Enabled = True
Else
    tbr.Buttons(1).Enabled = False
End If
If naprijed.ListCount > 0 Then
    tbr.Buttons(2).Enabled = True
Else
    tbr.Buttons(2).Enabled = False
End If
If tv.Nodes(1).Selected = True Then
    tbr.Buttons(3).Enabled = False
Else
    tbr.Buttons(3).Enabled = True
End If
a = lv.SelectedItem
If Err > 0 Then
    tbr.Buttons(7).Enabled = False
    tbr.Buttons(8).Enabled = False
    Err.Clear
Else
    tbr.Buttons(7).Enabled = True
    tbr.Buttons(8).Enabled = True
End If
k = 0
For i = 1 To lv.ListItems.Count
   If lv.ListItems(i).Selected = True Then
        k = k + 1
   End If
Next i
If k > 1 Or k < 1 Then
    us.purb.Enabled = False
Else
    us.purb.Enabled = True
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
fdrv = 0
Set fso = New FileSystemObject
For Each drv In fso.Drives
    If drv.IsReady Then
        If fdrv = 0 Then
            fdrv = 1
            drv_ = drv.path
        End If
        tv.Nodes.Add , , drv.path, drv.path, "drive"
        
        Set fldrs = fso.GetFolder(drv.path & "\").SubFolders
        For Each fldr In fldrs
            tv.Nodes.Add drv.path, tvwChild, fldr.path, fldr.Name, 1, 2
        Next
    Else
        tv.Nodes.Add , , drv.path, drv.path, "drive"
        'tv.Nodes(drv.Path).ForeColor = vbRed
    End If
Next
tv.Nodes(drv_).Selected = True
tv.SelectedItem.EnsureVisible
get_i
Me.caption = tv.SelectedItem
nazad.AddItem tv.SelectedItem
End Sub

Sub lvdblclck()
lv_DblClick
End Sub

Sub get_i()
If duse <> True Then
    On Error Resume Next
back:
    Set fso = New FileSystemObject
    Dim imgX As ListImage
    Screen.MousePointer = 11
    lv.ListItems.Clear
    Set lv.Icons = Nothing
    Set lv.SmallIcons = Nothing
    il1.ListImages.Clear
    il2.ListImages.Clear
    il1.ImageHeight = 32
    il1.ImageWidth = 32
    il2.ImageHeight = 16
    il2.ImageWidth = 16
    fpath = tv.SelectedItem.FullPath
    old_lst = fpath
    'DirBox.Refresh
    Set fldrs = fso.GetFolder(fpath & "\").SubFolders
    broj_f = fso.GetFolder(fpath & "\").SubFolders.Count
    t_br = 0
    If Err > 0 Then

        If MsgBox("Error: (" & Err.Description & ")" & vbCrLf & "Retry?", vbExclamation + vbYesNo) = vbYes Then
            GoTo back:
            Exit Sub
        Else
            nazad.RemoveItem 0
            tv.Nodes(act_dir).Selected = True
            tv.SelectedItem.EnsureVisible
            get_i
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    For Each fldr In fldrs
        rGetIcon = ExtractAssociatedIcon(0, fldr.path, 1)
        Set pictemp.Picture = Nothing
        DrawIcon pictemp.hdc, 0, 0, rGetIcon
        pictemp.Picture = pictemp.Image
        Set imgX = il1.ListImages.Add(, , pictemp.Picture)
        Set imgX = il2.ListImages.Add(, , pictemp.Picture)
        t_br = t_br + 1
        sbStatusBar.Panels(1).Text = "Folder icons: " & t_br & "/" & broj_f
        'DoEvents
        If (t_br Mod 50) = 0 Then
            DoEvents
        End If
    Next
    lv.Icons = il1: lv.SmallIcons = il2
    t_br = 0
    For Each fldr In fldrs
        k = k + 1
        lv.ListItems.Add k, fldr.path, fldr.Name, k, k
        tp = fldr.Type & "  "
        kr = fldr.DateCreated
        mn = fldr.DateLastModified
        ot = fldr.DateLastAccessed
        att = fldr.Attributes
        lv.ListItems(k).SubItems(1) = sz
        lv.ListItems(k).SubItems(2) = tp
        lv.ListItems(k).SubItems(3) = kr
        lv.ListItems(k).SubItems(4) = mn
        lv.ListItems(k).SubItems(5) = ot
        lv.ListItems(k).SubItems(6) = att
        t_br = t_br + 1
        sbStatusBar.Panels(1).Text = "Folders: " & t_br & "/" & broj_f
        If (t_br Mod 50) = 0 Then
            DoEvents
        End If
    Next
    
   
    
    Set fls = fso.GetFolder(fpath & "\").Files
    broj_f = fso.GetFolder(fpath & "\").Files.Count
    t_br = 0
    For Each fl In fls

        rGetIcon = ExtractAssociatedIcon(0, fl, 1)
        Set pictemp.Picture = Nothing
        DrawIcon pictemp.hdc, 0, 0, rGetIcon
        pictemp.Picture = pictemp.Image
        Set imgX = il1.ListImages.Add(, , pictemp.Picture)
        Set imgX = il2.ListImages.Add(, , pictemp.Picture)
        t_br = t_br + 1
        sbStatusBar.Panels(1).Text = "File icons: " & t_br & "/" & broj_f
        If (t_br Mod 50) = 0 Then
            DoEvents
        End If
    Next
    lv.Icons = il1: lv.SmallIcons = il2
    broj_f = fso.GetFolder(fpath & "\").Files.Count
    t_br = 0
    For Each fl In fls
        k = k + 1
        lv.ListItems.Add k, fl.path, fl.Name, k, k
        sz = fl.Size
        jd = ""
        If sz > 1024 Then
            sz = sz / 1024
            If sz > 1024 Then
                sz = sz / 1024
                If sz > 1024 Then
                    sz = sz / 1024
                    If sz > 1024 Then
                        sz = sz / 1024
                        jd = "TB"
                    Else
                        jd = "GB"
                    End If
                Else
                    jd = "MB"
                End If
            Else
                jd = "KB"
            End If
        Else
            sz = sz / 1024
            jd = "KB"
        End If
        sz = Round(sz, 2) & " " & jd
        tp = fl.Type
        kr = fl.DateCreated
        mn = fl.DateLastModified
        ot = fl.DateLastAccessed
        att = fl.Attributes
        lv.ListItems(k).SubItems(1) = sz
        lv.ListItems(k).SubItems(2) = tp
        lv.ListItems(k).SubItems(3) = kr
        lv.ListItems(k).SubItems(4) = mn
        lv.ListItems(k).SubItems(5) = ot
        lv.ListItems(k).SubItems(6) = att
        t_br = t_br + 1
        sbStatusBar.Panels(1).Text = "Files: " & t_br & "/" & broj_f
    Next
    sbStatusBar.Panels(1).Text = ""
    act_dir = tv.SelectedItem.FullPath
end_:
    Screen.MousePointer = 0
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload us
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
If Me.Height < 6105 Then
    Me.Height = 6105
End If
If Me.Width < 12420 Then
    Me.Width = 12420
End If
x = Me.ScaleWidth / 2.913876
tv.Width = x
lv.Left = tv.Left + tv.Width
lv.Width = Me.ScaleWidth - tv.Width
tv.Top = tbr.Height
lv.Top = tv.Top
tv.Height = Me.ScaleHeight - tbr.Height - sbStatusBar.Height
lv.Height = tv.Height
End If
End Sub

Private Sub lv_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error Resume Next
Dim fso As New FileSystemObject
If Right(lv.SelectedItem.ListSubItems(2).Text, 2) = "  " Then
    fldr_ = 1
Else
    fldr_ = 0
End If

fd = Left(mem_ln, InStrRev(mem_ln, "\"))
mem_nn = fd & NewString
If InStr(1, NewString, "\") > 0 Or InStr(1, NewString, "/") > 0 Or InStr(1, NewString, ":") > 0 Or InStr(1, NewString, "*") > 0 Or InStr(1, NewString, "?") > 0 Or InStr(1, NewString, """") > 0 Or InStr(1, NewString, "<") > 0 Or InStr(1, NewString, ">") > 0 Or InStr(1, NewString, "|") > 0 Then
    GoTo err_:
Else
    If fldr_ = 1 Then
        fso.GetFolder(mem_ln).Name = NewString
    Else
        fso.GetFile(mem_ln).Name = NewString
    End If
End If
If Err > 0 Then
err_:
    Cancel = True
    MsgBox "Error !!!" & vbCrLf & "Name contains char(s): \ / : * ? "" < > |", vbCritical, "Rename"
Else
    lv.SelectedItem.Key = mem_nn
End If
End Sub

Private Sub lv_BeforeLabelEdit(Cancel As Integer)
mem_ln = lv.SelectedItem.Key
End Sub

Private Sub lv_DblClick()
On Error GoTo end_:
If Right(lv.SelectedItem.SubItems(2), 2) = "  " Then
    tv.Nodes(tv.SelectedItem.FullPath & "\" & lv.SelectedItem).Selected = True
    nazad.AddItem tv.SelectedItem.FullPath, 0
    tv.SelectedItem.EnsureVisible
    tv.SelectedItem.Expanded = True
    Me.caption = tv.SelectedItem
    get_i
Else
    OpenFile1 frmExplore.lv.SelectedItem.Key, Me.hWnd
End If
end_:
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu us.izbor
End If
End Sub

Private Sub mnu_FileExit_Click()
End
End Sub

Private Sub optabout_Click()
MsgBox "Created by Edin Topalovic" & vbCrLf & "Email: topalovic.e@gmail.com", vbInformation
End Sub

Private Sub osvjezi_Click()
osvjezi_1
End Sub

Sub osvjezi_1()
If tv.SelectedItem.Image = "drive" Then
    osvjezi_11
Else
dio_k = tv.SelectedItem.Key
If tv.SelectedItem.Expanded = True Then
    exp_ = 1
End If
root_ = Left(dio_k, InStr(1, dio_k, "\") - 1)
tv.Nodes(root_).Selected = True
osvjezi_11
p = 0
For i = 1 To Len(dio_k)
    znak = Mid(dio_k, i, 1)
    If znak = "\" Then
        If p = 0 Then
            p = 1
        Else
            dio = Left(dio_k, i - 1)
            tv.Nodes(dio).Expanded = True
        End If
    End If
Next i
tv.Nodes(dio_k).Selected = True
If exp_ = 1 Then
    tv.Nodes(dio_k).Expanded = True
End If
End If
End Sub

Sub osvjezi_11()
On Error Resume Next
duse = True
If tv.Nodes(1).Selected = True Then
    tbr.Buttons(3).Enabled = False
End If
If tv.SelectedItem.Expanded = True Then
        exp_ = 1
End If
If tv.SelectedItem.Image = "drive" Then
    dio_k = tv.SelectedItem.Key
    dio_t = tv.SelectedItem.Text
    lind = tv.SelectedItem.Index
    tv.SelectedItem.Selected = False
    tv.Nodes.Remove lind
    tv.SelectedItem.Selected = False
    tv.Nodes.Add , , dio_k, dio_t, "drive", "drive"
    
    tv.Nodes(dio_k).Selected = True
    
    Set fldrs1 = fso.GetFolder(tv.SelectedItem.FullPath & "\").SubFolders
        For Each fldr1 In fldrs1
            tv.Nodes.Add dio_k, tvwChild, fldr1.path, fldr1.Name, 1, 2
            Set fldrs = fso.GetFolder(fldr1.path & "\").SubFolders
            For Each fldr In fldrs
                tv.Nodes.Add fldr1.path, tvwChild, fldr.path, fldr.Name, 1, 2
                
            Next
        Next
Else
    
    dio_k = tv.SelectedItem.Key
    dio_t = tv.SelectedItem.Text
    root_ = Left(dio_k, InStrRev(dio_k, "\") - 1)
    tv.Nodes.Remove tv.SelectedItem.Index
    tv.Nodes.Item(root_).Selected = True
    tv.Nodes.Add root_, tvwChild, dio_k, dio_t, 1, 2
    tv.Nodes.Item(root_).Sorted = True
    tv.Nodes(dio_k).Selected = True
    
    Set fldrs = fso.GetFolder(tv.SelectedItem.Key & "\").SubFolders
            For Each fldr In fldrs
                tv.Nodes.Add tv.SelectedItem.Key, tvwChild, fldr.path, fldr.Name, 1, 2
            Next
    
End If
If exp_ = 1 Then
    tv.SelectedItem.Expanded = True
End If
duse = False
'If tv.SelectedItem.FullPath <> old_lst Then
    naprijed.Clear
    nazad.AddItem tv.SelectedItem.FullPath, 0
    get_i
    Me.caption = tv.SelectedItem
    tbr.Buttons(3).Enabled = True
'End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
If Button.Index = 1 Then
    nazad_d
ElseIf Button.Index = 2 Then
    naprijed_d
ElseIf Button.Index = 3 Then
    gore_d
ElseIf Button.Index = 5 Then
    traži_d
ElseIf Button.Index = 7 Then
    premjesti_u_d
ElseIf Button.Index = 8 Then
    kopiraj_u_d
ElseIf Button.Index = 9 Then
    PopupMenu us.brisanje
ElseIf Button.Index = 11 Then
    PopupMenu us.pregled
ElseIf Button.Index = 12 Then
    cut_d
ElseIf Button.Index = 13 Then
    copy_d
ElseIf Button.Index = 14 Then
    paste_d
ElseIf Button.Index = 15 Then
    svojstva_
End If
End Sub
Sub svojstva_()
On Error Resume Next
Dim pth As String
    If ActiveControl.Name = "lv" Then
            pth = tv.SelectedItem.FullPath & "\" & lv.SelectedItem
    ElseIf ActiveControl.Name = "tv" Then
        pth = tv.SelectedItem.FullPath
    End If
    Call ShowPropWindow(pth, Me.hWnd)
End Sub

Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error Resume Next
Dim fso As New FileSystemObject
fd = Left(mem_ln, InStrRev(mem_ln, "\"))
mem_nn = fd & NewString
If InStr(1, NewString, "\") > 0 Or InStr(1, NewString, "/") > 0 Or InStr(1, NewString, ":") > 0 Or InStr(1, NewString, "*") > 0 Or InStr(1, NewString, "?") > 0 Or InStr(1, NewString, """") > 0 Or InStr(1, NewString, "<") > 0 Or InStr(1, NewString, ">") > 0 Or InStr(1, NewString, "|") > 0 Then
    GoTo err_:
Else
    fso.GetFolder(mem_ln).Name = NewString
End If
If Err > 0 Then
err_:
    Cancel = True
    MsgBox "Error !!!" & vbCrLf & "Name contains char(s): \ / : * ? "" < > |", vbCritical, "Rename"
Else
    tv.SelectedItem.Key = mem_nn
End If
End Sub

Private Sub tv_BeforeLabelEdit(Cancel As Integer)
mem_ln = tv.SelectedItem.Key
End Sub

Private Sub tv_Expand(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Screen.MousePointer = 11
Set fldrs1 = fso.GetFolder(Node.FullPath & "\").SubFolders
a = fso.GetFolder(Node.FullPath & "\").SubFolders.Count
s = 0
        For Each fldr1 In fldrs1
            Set fldrs = fso.GetFolder(fldr1.path & "\").SubFolders
            
            For Each fldr In fldrs
                tv.Nodes.Add fldr1.path, tvwChild, fldr.path, fldr.Name, 1, 2
            Next
            s = s + 1
            sbStatusBar.Panels(1).Text = "Folders: " & s & "/" & a
            If (s Mod 100) = 0 Then
                DoEvents
            End If
        Next
sbStatusBar.Panels(1).Text = ""
Screen.MousePointer = vbDefault
End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu us.izbor
End If
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
If tv.Nodes(1).Selected = True Then
    tbr.Buttons(3).Enabled = False
End If
If tv.SelectedItem.FullPath <> old_lst Then
    naprijed.Clear
    nazad.AddItem tv.SelectedItem.FullPath, 0
    get_i
    Me.caption = tv.SelectedItem
    tbr.Buttons(3).Enabled = True
End If
End Sub

Sub nazad_d()
    naprijed.AddItem nazad.List(0), 0
    tv.Nodes(nazad.List(1)).Selected = True
    If nazad.ListCount = 1 Then
    Else
    nazad.RemoveItem 0
    End If
    tv.SelectedItem.EnsureVisible
    get_i
End Sub

Sub naprijed_d()
On Error Resume Next
    nazad.AddItem naprijed.List(0), 0
    If nazad.ListCount = 1 Then
        naprijed.RemoveItem 0
        Exit Sub
    Else
        tv.Nodes(naprijed.List(1)).Selected = True
    End If
    naprijed.RemoveItem 0
    tv.SelectedItem.EnsureVisible
    get_i
End Sub

Sub gore_d()
    If tv.SelectedItem.Image = "drive" Then
        tv.Nodes(1).Selected = True
        nazad.AddItem tv.SelectedItem.FullPath, 0
        tbr.Buttons(3).Enabled = False
        tv.SelectedItem.EnsureVisible
        get_i
        Exit Sub
    End If
    fgh = Left(tv.SelectedItem.FullPath, Len(tv.SelectedItem.FullPath) - Len(tv.SelectedItem) - 1)
    tv.Nodes(fgh).Selected = True
    tv.SelectedItem.EnsureVisible
    get_i
End Sub

Sub traži_d()
    Dim strPathToSearch As String
    If ActiveControl.Name = "lv" Then
        If Right(lv.SelectedItem.SubItems(2), 2) = "  " Then
            strPathToSearch = tv.SelectedItem.FullPath & "\" & lv.SelectedItem
        Else
            strPathToSearch = tv.SelectedItem
        End If
    ElseIf ActiveControl.Name = "tv" Then
        strPathToSearch = tv.SelectedItem.FullPath
    End If
    Call ShellExecute(Me.hWnd, "find", strPathToSearch, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Sub premjesti_u_d()
    mv.caption = "Moving..."
    mv.lokacija = get_loc
    If mv.lokacija = "" Then
        Exit Sub
    End If
    If Me.ActiveControl.Name = "tv" Then
        mv.folderi = tv.SelectedItem.FullPath
    Else
        mv.folderi = folderi
        mv.datoteke = datoteke
    End If
    mv.operacija = 0
    mv.animacija = "Kopiranje"
    izvrši_1
End Sub

Sub kopiraj_u_d()
    mv.caption = "Copying..."
    mv.lokacija = get_loc
    If mv.lokacija = "" Then
        Exit Sub
    End If
    If Me.ActiveControl.Name = "tv" Then
        mv.folderi = tv.SelectedItem.FullPath
    Else
        mv.folderi = folderi
        mv.datoteke = datoteke
    End If
    mv.operacija = 1
    mv.animacija = "Kopiranje"
    izvrši_1
End Sub

Sub cut_d()
Dim fso As New FileSystemObject
    clp.opr.move = True
    clp.opr.copy = False
    get_pth
    clp.fldrs2 = fldrs11
    clp.fls2 = fls11
End Sub

Sub copy_d()
Dim fso As New FileSystemObject
    clp.opr.move = False
    clp.opr.copy = True
    get_pth
    clp.fldrs2 = fldrs11
    clp.fls2 = fls11
End Sub

Sub paste_d()
On Error Resume Next
If clp.fldrs2 = "" And clp.fls2 = "" Then Exit Sub
    mv.folderi = clp.fldrs2
    mv.datoteke = clp.fls2
    If clp.opr.copy = True Then
        mv.caption = "Copying..."
        mv.animacija = "Kopiranje"
        mv.operacija = 1
    Else
        mv.caption = "Moving..."
        mv.animacija = "Premještanje"
        mv.operacija = 0
    End If
    mv.lokacija = lfc
    If mv.lokacija = "" Then
        Exit Sub
    End If
    izvrši_1
End Sub
