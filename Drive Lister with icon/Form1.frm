VERSION 5.00
Object = "{3A94456F-33AF-4D40-A77B-936366DCE6FB}#1.0#0"; "AeroSuite.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3540
      ItemData        =   "Form1.frx":0000
      Left            =   6600
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ImageList il 
      Left            =   480
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":51FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AeroSuite.AeroListBox al 
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemHeight      =   48
      ItemHeightAuto  =   0   'False
      SelectMode      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim i As Integer
al.SetImageList il

'.........................Collect Fixed Drives.....
Call GetDrives(List1, Fixed)
i = 0
For i = 0 To List1.ListCount
    If List1.List(i) <> "" Then al.AddItem List1.List(i), 1, 1
Next i
'..................................................

'........................Collect CD-ROM Drives.....
Call GetDrives(List1, CDROM)
i = 0
For i = 0 To List1.ListCount
    If List1.List(i) <> "" Then al.AddItem List1.List(i), 2, 2
Next i
'..................................................

'......................Collect Removble Drives.....
Call GetDrives(List1, REMOVABLE)
i = 0
For i = 0 To List1.ListCount
    If List1.List(i) <> "" Then al.AddItem List1.List(i), 3, 3
Next i
'..................................................

'...........................Collect RAM Drives.....
Call GetDrives(List1, RAMDISK)
i = 0
For i = 0 To List1.ListCount
    If List1.List(i) <> "" Then al.AddItem List1.List(i), 4, 4
Next i
'..................................................

'........................Collect Remote Drives.....
Call GetDrives(List1, REMOTE)
i = 0
For i = 0 To List1.ListCount
    If List1.List(i) <> "" Then al.AddItem List1.List(i), 5, 5
Next i
'..................................................

End Sub
