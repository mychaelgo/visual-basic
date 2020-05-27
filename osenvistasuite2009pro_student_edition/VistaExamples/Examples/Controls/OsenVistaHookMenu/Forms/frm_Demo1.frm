VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_demo1 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Sample OsenXPHookMenu"
   ClientHeight    =   6315
   ClientLeft      =   1440
   ClientTop       =   930
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frm_Demo1.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaCheckBox OsenXPCheckBox1 
      Height          =   255
      Left            =   3090
      TabIndex        =   6
      Top             =   1410
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   450
      BackColor       =   16767935
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
      Caption         =   "Right to Left"
      Style           =   1
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_Demo1.frx":058A
      Left            =   1320
      List            =   "frm_Demo1.frx":0597
      TabIndex        =   1
      Top             =   1380
      Width           =   1545
   End
   Begin VistaSuitePro.OsenVistaToolBar OsenXPToolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   810
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XPBlend         =   0   'False
      TotalButton     =   20
      Bpic1           =   "frm_Demo1.frx":05C8
      Bname1          =   "New"
      Btype1          =   0
      Bwidth1         =   0
      Bchecked1       =   0   'False
      Bvalue1         =   0   'False
      Bpic2           =   "frm_Demo1.frx":091A
      BEnabled2       =   0   'False
      Bname2          =   "Open"
      Btype2          =   0
      Bwidth2         =   0
      Bchecked2       =   0   'False
      Bvalue2         =   0   'False
      BNI2            =   0
      BSI2            =   0
      Bpic3           =   "frm_Demo1.frx":0C6C
      BEnabled3       =   0   'False
      Bname3          =   "Save"
      Btype3          =   0
      Bwidth3         =   0
      Bchecked3       =   0   'False
      Bvalue3         =   0   'False
      BNI3            =   0
      BSI3            =   0
      Bname4          =   "Button4"
      Btype4          =   2
      Bwidth4         =   0
      Bchecked4       =   0   'False
      Bvalue4         =   0   'False
      Bpic5           =   "frm_Demo1.frx":0FBE
      Bname5          =   "Print"
      Btype5          =   0
      Bwidth5         =   0
      Bchecked5       =   0   'False
      Bvalue5         =   0   'False
      Bpic6           =   "frm_Demo1.frx":1310
      Bname6          =   "PreView"
      Btype6          =   0
      Bwidth6         =   0
      Bchecked6       =   0   'False
      Bvalue6         =   0   'False
      Bpic7           =   "frm_Demo1.frx":1662
      Bname7          =   "Check Spelling"
      Btype7          =   0
      Bwidth7         =   0
      Bchecked7       =   0   'False
      Bvalue7         =   0   'False
      Bname8          =   "Button8"
      Btype8          =   2
      Bwidth8         =   0
      Bchecked8       =   0   'False
      Bvalue8         =   0   'False
      Bpic9           =   "frm_Demo1.frx":19B4
      Bname9          =   "Cut"
      Btype9          =   0
      Bwidth9         =   0
      Bchecked9       =   0   'False
      Bvalue9         =   0   'False
      Bpic10          =   "frm_Demo1.frx":1D06
      Bname10         =   "Copy"
      Btype10         =   0
      Bwidth10        =   0
      Bchecked10      =   0   'False
      Bvalue10        =   0   'False
      Bpic11          =   "frm_Demo1.frx":2058
      Bname11         =   "Paste"
      Btype11         =   0
      Bwidth11        =   0
      Bchecked11      =   0   'False
      Bvalue11        =   0   'False
      Bname12         =   "Button12"
      Btype12         =   2
      Bwidth12        =   0
      Bchecked12      =   0   'False
      Bvalue12        =   0   'False
      Bpic13          =   "frm_Demo1.frx":23AA
      Bname13         =   "Undo"
      Btype13         =   1
      Bwidth13        =   0
      Bchecked13      =   0   'False
      Bvalue13        =   0   'False
      Bpic14          =   "frm_Demo1.frx":26FC
      Bname14         =   "Redo"
      Btype14         =   1
      Bwidth14        =   0
      Bchecked14      =   0   'False
      Bvalue14        =   0   'False
      Bname15         =   "Button15"
      Btype15         =   2
      Bwidth15        =   0
      Bchecked15      =   0   'False
      Bvalue15        =   0   'False
      Bpic16          =   "frm_Demo1.frx":2A4E
      Bname16         =   "History"
      Btype16         =   0
      Bwidth16        =   0
      Bchecked16      =   0   'False
      Bvalue16        =   0   'False
      Bpic17          =   "frm_Demo1.frx":2DA0
      Bname17         =   "Summary"
      Btype17         =   0
      Bwidth17        =   0
      Bchecked17      =   0   'False
      Bvalue17        =   0   'False
      Bpic18          =   "frm_Demo1.frx":30F2
      Bname18         =   "Sort Ascending"
      Btype18         =   0
      Bwidth18        =   0
      Bchecked18      =   0   'False
      Bvalue18        =   0   'False
      Bpic19          =   "frm_Demo1.frx":3444
      Bname19         =   "Sort Descending"
      Btype19         =   0
      Bwidth19        =   0
      Bchecked19      =   0   'False
      Bvalue19        =   0   'False
      Bname20         =   "Button20"
      Btype20         =   2
      Bwidth20        =   0
      Bchecked20      =   0   'False
      Bvalue20        =   0   'False
      BackColor       =   16244941
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaHookMenu OsenXPHookMenu1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   688
      BmpCount        =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GripperLeft     =   12
      MCountMenu      =   5
      Bmp:1           =   "frm_Demo1.frx":3796
      Key:1           =   "#mnuFile:0"
      Bmp:2           =   "frm_Demo1.frx":3BBE
      Key:2           =   "#mnuFile:1"
      Bmp:3           =   "frm_Demo1.frx":3FE6
      Key:3           =   "#mnuFile:2"
      Bmp:4           =   "frm_Demo1.frx":440E
      Key:4           =   "#mnuFile:5"
      Bmp:5           =   "frm_Demo1.frx":4836
      Key:5           =   "#mnuFile:6"
      Bmp:6           =   "frm_Demo1.frx":4C5E
      Key:6           =   "#mnuEdit:0"
      Bmp:7           =   "frm_Demo1.frx":5086
      Key:7           =   "#mnuEdit:2"
      Bmp:8           =   "frm_Demo1.frx":54AE
      Key:8           =   "#mnuEdit:3"
      Bmp:9           =   "frm_Demo1.frx":58D6
      Key:9           =   "#mnuEdit:4"
      Bmp:10          =   "frm_Demo1.frx":5CFE
      Key:10          =   "#mnuEdit:5"
      Bmp:11          =   "frm_Demo1.frx":6126
      Key:11          =   "#mnuEdit:7"
      Bmp:12          =   "frm_Demo1.frx":654E
      Key:12          =   "#mnuEdit:8"
      Bmp:13          =   "frm_Demo1.frx":6976
      Key:13          =   "#mnuEdit:10"
      Bmp:14          =   "frm_Demo1.frx":6D9E
      Key:14          =   "#mnuEdit:13"
      Bmp:15          =   "frm_Demo1.frx":71C6
      Key:15          =   "#mnuFormat:1"
      Bmp:16          =   "frm_Demo1.frx":75EE
      Key:16          =   "#mnuFormat:0"
      Bmp:17          =   "frm_Demo1.frx":7A16
      Key:17          =   "#mnuHelp:0"
      Bmp:18          =   "frm_Demo1.frx":7E3E
      Key:18          =   "#mnuHelp:1"
      XMenuA1         =   "&File "
      XMenuACS1       =   "f"
      XMenuC1         =   "mnuFileTop"
      XMenuE1         =   -1  'True
      XMenuH1         =   -1  'True
      XMenuA2         =   "&Edit "
      XMenuACS2       =   "e"
      XMenuC2         =   "mnuEditTop"
      XMenuE2         =   -1  'True
      XMenuH2         =   -1  'True
      XMenuA3         =   "F&ormat "
      XMenuACS3       =   "o"
      XMenuC3         =   "mnuFormatTop"
      XMenuE3         =   -1  'True
      XMenuH3         =   -1  'True
      XMenuA4         =   "&View "
      XMenuACS4       =   "v"
      XMenuC4         =   "mnuViewTop"
      XMenuE4         =   -1  'True
      XMenuH4         =   -1  'True
      XMenuA5         =   "&Help "
      XMenuACS5       =   "h"
      XMenuC5         =   "mnuHelpTop"
      XMenuE5         =   -1  'True
      XMenuH5         =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   5850
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   820
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   2
      HaveXPForm      =   -1  'True
      WindowColor     =   3
      PWidth1         =   100
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Osen Kusnadi"
      pTextAlignment1 =   0
      PanelPicture1   =   "frm_Demo1.frx":8266
      PanelPicAlignment1=   0
      PWidth2         =   200
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Email: support@osenxpsuite.net"
      pTextAlignment2 =   0
      PanelPicture2   =   "frm_Demo1.frx":85B8
      PanelPicAlignment2=   0
      GradientColor1  =   15382160
      GradientColor2  =   12752244
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sample OsenXPHookMenu"
      TitleTop        =   7
      icon            =   "frm_Demo1.frx":85D4
      BorderStyle     =   1
      UseDefaultTheme =   0   'False
      WindowColor     =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Scheme:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1410
      Width           =   1035
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &As..."
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Page Set&up"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print..."
         Index           =   6
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin VB.Menu mnuEditTop 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cu&t"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "De&lete"
         Index           =   5
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Find..."
         Index           =   7
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find &Next"
         Index           =   8
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Replace..."
         Index           =   9
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Go To..."
         Enabled         =   0   'False
         Index           =   10
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select &All"
         Index           =   12
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Time/&Date..."
         Index           =   13
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFormatTop 
      Caption         =   "F&ormat"
      Visible         =   0   'False
      Begin VB.Menu mnuFormat 
         Caption         =   "&Word Wrap"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "&Font..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuViewTop 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu mnuView 
         Caption         =   "&Status Bar"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelp 
         Caption         =   "Help T&opics..."
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About Notepad..."
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm_demo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************'
'*  OsenXPSuite 2006 - OsenXPHookMenu sample                             *'
'*  Copyright (c) 2006 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  This file is part of the OsenXPSuite 2006 sample applications.       *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenXPSuite 2006 documentation, and is provided "as is", without     *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
End Sub


Private Sub OsenXPCheckBox1_Click()
    Me.OsenXPHookMenu1.RightToLeft = Me.OsenXPCheckBox1.Value
    Me.OsenXPForm1.RightToLeft = Me.OsenXPCheckBox1.Value
    OsenXPCheckBox1.Alignment = OsenXPCheckBox1.Value
    Me.OsenXPToolBar1.RightToLeft = OsenXPCheckBox1.Value
    Me.OsenXPStatusBar1.RightToLeft = OsenXPCheckBox1.Value
End Sub

Private Sub OsenXPHookMenu1_Highlight(strMenuCaption As String)
    OsenXPStatusBar1.PanelCaption(2) = strMenuCaption
End Sub

Private Sub OsenXPToolBar1_Highlight(Index As Integer, sText As String)
    OsenXPStatusBar1.PanelCaption(2) = sText
End Sub

Private Sub Combo1_Click()
'    Me.OsenXPForm1.ColorScheme = Combo1.ListIndex
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' purpose: to allow shortcut key for hookmenu when user press alt+key
    Me.OsenXPHookMenu1.GetShortcut KeyCode, Shift
    
End Sub






















