VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect Manager"
   ClientHeight    =   6945
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniMenu UniMenu1 
      Left            =   0
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   767
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2880
      Top             =   4080
   End
   Begin UniControls.UniFrame UniFrame1 
      Height          =   2655
      Left            =   240
      Top             =   4200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Tho6ng tin"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniLabel UniLabel12 
         Height          =   255
         Left            =   5520
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "Bo65 nho71 RAM"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin VB.PictureBox PicIcon 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin UniControls.UniLabel UniLabel10 
         Height          =   255
         Left            =   1080
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "Thuo65c ti1nh:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniTextBox txtFileName 
         Height          =   270
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniLabel UniLabel9 
         Height          =   255
         Left            =   1080
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Caption         =   "Chie61m bo65 nho71:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel8 
         Height          =   255
         Left            =   1080
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "Mu71c u7u tie6n:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel7 
         Height          =   255
         Left            =   1080
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Caption         =   "Dung lu7o75ng:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel6 
         Height          =   255
         Left            =   1080
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Caption         =   "Phie6n ba3n:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel5 
         Height          =   255
         Left            =   1080
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Mie6u ta3 u71ng du5ng:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel4 
         Height          =   255
         Left            =   1080
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Caption         =   "Nha2 Pha1t Ha2nh:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel3 
         Height          =   255
         Left            =   1080
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Caption         =   "D9i5a chi3:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel2 
         Height          =   255
         Left            =   1080
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "Te6n File:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniTextBox txtPath 
         Height          =   270
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniTextBox txtCompany 
         Height          =   270
         Left            =   2640
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniTextBox txtDecription 
         Height          =   270
         Left            =   2640
         TabIndex        =   6
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniTextBox txtVersion 
         Height          =   270
         Left            =   2640
         TabIndex        =   7
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniTextBox txtFileSize 
         Height          =   270
         Left            =   2640
         TabIndex        =   8
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniTextBox txtUuTien 
         Height          =   270
         Left            =   2640
         TabIndex        =   9
         Top             =   2040
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniTextBox txtMemory 
         Height          =   270
         Left            =   2640
         TabIndex        =   10
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniTextBox txtThuocTinh 
         Height          =   270
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   255
         Left            =   5520
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Info 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   2
         Left            =   6720
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Info 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   1
         Left            =   6720
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Info 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   0
         Left            =   6720
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "System Cache:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5520
         TabIndex        =   14
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Available :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5520
         TabIndex        =   13
         Top             =   960
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00B99D7F&
         X1              =   5400
         X2              =   5400
         Y1              =   240
         Y2              =   2520
      End
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   240
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Qua3n ly1 tie61n tri2nh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImA"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Image Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PID"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5520
      Top             =   2040
   End
   Begin VB.ListBox lstPro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImA 
      Left            =   5880
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu thaotac 
      Caption         =   "Thao ta1c"
      Begin VB.Menu tatungdung 
         Caption         =   "Ta81t u71ng du5ng"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu tatvaxoa 
         Caption         =   "Ta81t va2 xo1a"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu lamtuoi 
         Caption         =   "La2m tu7o7i"
         Shortcut        =   {F5}
      End
      Begin VB.Menu denthumucgoc 
         Caption         =   "D9e61n thu7 mu5c go61c"
         Shortcut        =   ^G
      End
      Begin VB.Menu dongbang 
         Caption         =   "D9o1ng ba8ng"
      End
      Begin VB.Menu khoiphuc 
         Caption         =   "Mo73 d9o1ng ba8ng"
      End
      Begin VB.Menu datmucuutien 
         Caption         =   "D9a85t mu71c u7u tie6n"
         Begin VB.Menu mutCAONHAT 
            Caption         =   "Cao Nha61t"
         End
         Begin VB.Menu mutCAO 
            Caption         =   "Cao"
         End
         Begin VB.Menu mutTB 
            Caption         =   "Trung Bi2nh"
         End
         Begin VB.Menu mutTHAP 
            Caption         =   "Tha61p"
         End
      End
      Begin VB.Menu xemthuoctinh 
         Caption         =   "Xem thuo65c ti1nh"
      End
   End
   Begin VB.Menu about 
      Caption         =   "Copyright © Perfect Antivirus 2009"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SoLuong




Private Sub denthumucgoc_Click()
Shell "explorer.exe /select," & LV1.ListItems(LV1.SelectedItem.Index).SubItems(1), vbNormalFocus
End Sub

Private Sub dongbang_Click()
SuspendResumeProcess LV1.ListItems(LV1.SelectedItem.Index).SubItems(2), True
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End

UniMenu1.InitUnicodeMenu
End Sub

Private Sub khoiphuc_Click()
SuspendResumeProcess LV1.ListItems(LV1.SelectedItem.Index).SubItems(2), False
End Sub

Private Sub lamtuoi_Click()
    GetProcess LV1, ImA, Pic
End Sub


Private Sub LV1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)



'MsgBox LV1.SelectedItem.Text & "=" & LV1.ListItems(LV1.SelectedItem.Index).SubItems(1)
On Error Resume Next
Dim hVer As VERHEADER
With Me
    .txtFileName.Text = LV1.SelectedItem.Text
    .txtPath.Text = LV1.ListItems(LV1.SelectedItem.Index).SubItems(1)
    GetVerHeader txtPath.Text, hVer
    .txtCompany.Text = hVer.CompanyName
    .txtThuocTinh.Text = GetAttribute(.txtPath.Text)
    .txtDecription.Text = hVer.FileDescription
    .txtVersion.Text = hVer.FileVersion
    .txtFileSize.Text = Format(FileLen(.txtPath.Text) \ 1024 & " KB", "###,###")
    .txtUuTien.Text = GetBasePriority(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2))
    .txtMemory.Text = Format(GetMemory(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2)) \ 1024, "###,###") & " KB"
    GetLargeIcon .txtPath.Text, PicIcon
End With















If Button = 2 Then
Dim YY
'MsgBox GetPriority(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2))
YY = GetPriority(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2))
mutTB.Checked = False
mutTHAP.Checked = False
mutCAO.Checked = False
mutCAONHAT.Checked = False


If YY = 32 Then
    mutTB.Checked = True
ElseIf YY = 64 Then
    mutTHAP.Checked = True
ElseIf YY = 128 Then
    mutCAO.Checked = True
ElseIf YY = 256 Then
    mutCAONHAT.Checked = True
End If

PopupMenu thaotac
End If
End Sub

Private Sub mutCAO_Click()
Dim Pri As Long
Pri = OpenProcess(PROCESS_SET_INFORMATION, False, CLng(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2)))
SetPriorityClass Pri, HIGH_PRIORITY_CLASS
    GetProcess LV1, ImA, Pic
End Sub

Private Sub mutCAONHAT_Click()
Dim Pri As Long
Pri = OpenProcess(PROCESS_SET_INFORMATION, False, CLng(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2)))
SetPriorityClass Pri, REALTIME_PRIORITY_CLASS
    GetProcess LV1, ImA, Pic
End Sub

Private Sub mutTB_Click()
Dim Pri As Long
Pri = OpenProcess(PROCESS_SET_INFORMATION, False, CLng(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2)))
SetPriorityClass Pri, NORMAL_PRIORITY_CLASS
    GetProcess LV1, ImA, Pic
End Sub

Private Sub mutTHAP_Click()
Dim Pri As Long
Pri = OpenProcess(PROCESS_SET_INFORMATION, False, CLng(LV1.ListItems(LV1.SelectedItem.Index).SubItems(2)))
SetPriorityClass Pri, IDLE_PRIORITY_CLASS
    GetProcess LV1, ImA, Pic
End Sub

Private Sub tatungdung_Click()
If UniMsgBox("Ba5n co1 muo61n ta81t u71ng du5ng: " & LV1.SelectedItem.Text & " kho6ng?", vbYesNo) = vbYes Then basProcess.KillProcessById (LV1.ListItems(LV1.SelectedItem.Index).SubItems(2))

End Sub

Private Sub tatvaxoa_Click()

If UniMsgBox("Ba5n co1 muo61n ta81t va2 xo1a u71ng du5ng: " & LV1.SelectedItem.Text & " kho6ng?", vbYesNo) = vbYes Then

KillProcessById (LV1.ListItems(LV1.SelectedItem.Index).SubItems(2))
If basFile.zDeletefile(LV1.ListItems(LV1.SelectedItem.Index).SubItems(1)) = True Then
    UniMsgBox "D9a4 ta81t va2 xo1a " & LV1.ListItems(LV1.SelectedItem.Index).SubItems(1)
Else
    UniMsgBox "Kho6ng xo1a d9u7o75c! " & LV1.ListItems(LV1.SelectedItem.Index).SubItems(1)
End If

End If
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
    i = 0
   snap = CreateToolhelpSnapshot(TH32CS_SNAPALL, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0

      theloop = ProcessNext(snap, proc)
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
            'MsgBox ProcessPathByPID(proc.th32ProcessID)
              i = i + 1
              If i > LV1.ListItems.Count Then GoTo ThEmVaO
        Dim xProPa
        xProPa = ProcessPathByPID(proc.th32ProcessID)
        If Left(xProPa, 4) = "\??\" Then xProPa = Right(xProPa, Len(xProPa) - 4)
        If Left(xProPa, Len("\SystemRoot\")) = "\SystemRoot\" Then xProPa = "C:\WINDOWS\" & Right(xProPa, Len(xProPa) - Len("\SystemRoot\"))

            If LV1.ListItems(i).SubItems(1) <> xProPa Then GoTo ThEmVaO
      End If
   Wend

   CloseHandle snap
   EnumWindows AddressOf EnumWindowsProc, ByVal 0&
   If i < LV1.ListItems.Count - SoLuong Then GoTo ThEmVaO
Exit Sub
ThEmVaO:
    GetProcess LV1, ImA, Pic
End Sub

Private Sub Timer2_Timer()
MonitoringPerformance Info(0), Info(1), Info(2)
End Sub

Private Sub xemthuoctinh_Click()
ShowProperties LV1.ListItems(LV1.SelectedItem.Index).SubItems(1), Me.hWnd
End Sub
