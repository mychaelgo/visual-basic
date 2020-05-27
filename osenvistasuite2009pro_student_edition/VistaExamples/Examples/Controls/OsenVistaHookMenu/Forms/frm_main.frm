VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form Frm_Main 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Sample Menu,Toolbar and Statusbar usage"
   ClientHeight    =   7125
   ClientLeft      =   4515
   ClientTop       =   2205
   ClientWidth     =   12315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VistaSuitePro.MyImageList MyImg 
      Left            =   5820
      Top             =   5460
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   101024
      Images          =   "frm_main.frx":058A
      Version         =   131072
      KeyCount        =   88
      Keys            =   $"frm_main.frx":1904A
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   1605
      Left            =   1350
      TabIndex        =   14
      Top             =   4020
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2831
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      Caption         =   $"frm_main.frx":190FC
      AutoSize        =   0   'False
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   435
      Left            =   5190
      TabIndex        =   7
      Top             =   4230
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   767
      Caption         =   "&Copy Items to other listbox >>"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_main.frx":1918D
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frm_main.frx":191A9
      BinaryImageOver =   "frm_main.frx":191C1
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox2 
      Height          =   5145
      Left            =   7980
      TabIndex        =   6
      Top             =   1380
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   9075
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   16777215
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      HeaderFormatString=   "ico;27;1|Attribute1;150;0|Attribute2;75;1"
      Columns         =   3
      ShowGridLines   =   -1  'True
      AlternateRowColors=   -1  'True
      MaxAllColumnWidth=   252
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   "myimg"
      ForeColorSelected=   16576
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "frm_main.frx":191D9
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaLabel LblMail 
      Height          =   255
      Left            =   210
      TabIndex        =   11
      ToolTipText     =   "support@osenxpsuite2005.com"
      Top             =   7260
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   450
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frm_main.frx":191F1
      MousePointer    =   99
      Picture         =   "frm_main.frx":19353
      ForeColorOver   =   16748098
      ForeColorDown   =   5412350
      Caption         =   "Mail: support@osenxpsuite.net"
      HiperLink       =   "mailto:support@osenxpsuite.net"
      AutoSize        =   0   'False
      UnderlineOnOver =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton3 
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   3270
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      Caption         =   "Clear"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_main.frx":196ED
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frm_main.frx":19709
      BinaryImageOver =   "frm_main.frx":19721
   End
   Begin VistaSuitePro.OsenVistaButton CmdAdd2 
      Height          =   375
      Left            =   5130
      TabIndex        =   9
      Top             =   2550
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   661
      Caption         =   "Additem without LockUpdate"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_main.frx":19739
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frm_main.frx":19755
      BinaryImageOver =   "frm_main.frx":1976D
   End
   Begin VistaSuitePro.OsenVistaButton CmdAdd1 
      Height          =   375
      Left            =   5130
      TabIndex        =   10
      Top             =   1740
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "AddItem with LockUpdate"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_main.frx":19785
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frm_main.frx":197A1
      BinaryImageOver =   "frm_main.frx":197B9
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   5145
      Left            =   540
      TabIndex        =   2
      Top             =   1350
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   9075
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   16777215
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      HeaderFormatString=   "ico;27;1|Attribute1;150;0|Attribute2;75;1"
      Columns         =   3
      ShowGridLines   =   -1  'True
      AlternateRowColors=   -1  'True
      MaxAllColumnWidth=   252
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   "myimg"
      ForeColorSelected=   16576
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "frm_main.frx":197D1
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaStatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   714
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   8421504
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
      NumberOfPanels  =   7
      HaveXPForm      =   -1  'True
      WindowColor     =   3
      PWidth1         =   87
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   3
      pText1          =   "19.4.2004"
      pTextAlignment1 =   0
      PanelPicture1   =   "frm_main.frx":197E9
      PanelPicAlignment1=   0
      PWidth2         =   78
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   2
      pText2          =   "07:19:50"
      pTextAlignment2 =   0
      PanelPicture2   =   "frm_main.frx":19B3B
      PanelPicAlignment2=   0
      PWidth3         =   200
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   ""
      pTextAlignment3 =   0
      PanelPicture3   =   "frm_main.frx":19E8D
      PanelPicAlignment3=   0
      PWidth4         =   50
      PMinWidth4      =   0
      pTTText4        =   ""
      pType4          =   5
      pText4          =   "CAPS"
      pTextAlignment4 =   0
      PanelPicture4   =   "frm_main.frx":19EA9
      PanelPicAlignment4=   0
      PWidth5         =   48
      PMinWidth5      =   0
      pTTText5        =   ""
      pType5          =   6
      pText5          =   "NUM"
      pTextAlignment5 =   0
      PanelPicture5   =   "frm_main.frx":19EC5
      PanelPicAlignment5=   0
      PWidth6         =   60
      PMinWidth6      =   0
      pTTText6        =   ""
      pType6          =   7
      pText6          =   "SCROLL"
      pTextAlignment6 =   0
      PanelPicture6   =   "frm_main.frx":19EE1
      PanelPicAlignment6=   0
      PWidth7         =   100
      PMinWidth7      =   0
      pTTText7        =   ""
      pType7          =   0
      pText7          =   ""
      pTextAlignment7 =   0
      PanelPicAlignment7=   0
      GradientColor1  =   15382160
      GradientColor2  =   12752244
   End
   Begin VistaSuitePro.OsenVistaToolBar TBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   810
      Width           =   12315
      _ExtentX        =   21722
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
      TotalButton     =   23
      Bpic1           =   "frm_main.frx":19EFD
      Bname1          =   "New"
      Btype1          =   0
      Bwidth1         =   0
      Bchecked1       =   0   'False
      Bvalue1         =   0   'False
      Bpic2           =   "frm_main.frx":1A24F
      Bname2          =   "Open"
      Btype2          =   0
      Bwidth2         =   0
      Bchecked2       =   0   'False
      Bvalue2         =   0   'False
      Bpic3           =   "frm_main.frx":1A5A1
      Bname3          =   "Save"
      Btype3          =   0
      Bwidth3         =   0
      Bchecked3       =   0   'False
      Bvalue3         =   0   'False
      Bpic4           =   "frm_main.frx":1A8F3
      Bname4          =   "Search"
      Btype4          =   0
      Bwidth4         =   0
      Bchecked4       =   0   'False
      Bvalue4         =   0   'False
      Bname5          =   "Button5"
      Btype5          =   2
      Bwidth5         =   0
      Bchecked5       =   0   'False
      Bvalue5         =   0   'False
      Bpic6           =   "frm_main.frx":1AC45
      Bname6          =   "Print"
      Btype6          =   0
      Bwidth6         =   0
      Bchecked6       =   0   'False
      Bvalue6         =   0   'False
      Bpic7           =   "frm_main.frx":1AF97
      Bname7          =   "PreView"
      Btype7          =   0
      Bwidth7         =   0
      Bchecked7       =   0   'False
      Bvalue7         =   0   'False
      Bpic8           =   "frm_main.frx":1B2E9
      Bname8          =   "Spelling"
      Btype8          =   0
      Bwidth8         =   0
      Bchecked8       =   0   'False
      Bvalue8         =   0   'False
      Bname9          =   "Button9"
      Btype9          =   2
      Bwidth9         =   0
      Bchecked9       =   0   'False
      Bvalue9         =   0   'False
      Bpic10          =   "frm_main.frx":1B63B
      Bname10         =   "Cut"
      Btype10         =   0
      Bwidth10        =   0
      Bchecked10      =   0   'False
      Bvalue10        =   0   'False
      Bpic11          =   "frm_main.frx":1B98D
      Bname11         =   "Copy"
      Btype11         =   0
      Bwidth11        =   0
      Bchecked11      =   0   'False
      Bvalue11        =   0   'False
      Bpic12          =   "frm_main.frx":1BCDF
      Bname12         =   "Paste"
      Btype12         =   1
      Bwidth12        =   0
      Bchecked12      =   0   'False
      Bvalue12        =   0   'False
      Bpic13          =   "frm_main.frx":1C031
      Bname13         =   "Clean"
      Btype13         =   0
      Bwidth13        =   0
      Bchecked13      =   0   'False
      Bvalue13        =   0   'False
      Bname14         =   "Button15"
      Btype14         =   2
      Bwidth14        =   0
      Bchecked14      =   0   'False
      Bvalue14        =   0   'False
      Bpic15          =   "frm_main.frx":1C383
      Bname15         =   "Undo"
      Btype15         =   1
      Bwidth15        =   0
      Bchecked15      =   0   'False
      Bvalue15        =   0   'False
      Bpic16          =   "frm_main.frx":1C6D5
      Bname16         =   "Redo"
      Btype16         =   1
      Bwidth16        =   0
      Bchecked16      =   0   'False
      Bvalue16        =   0   'False
      Bname17         =   "Button17"
      Btype17         =   2
      Bwidth17        =   0
      Bchecked17      =   0   'False
      Bvalue17        =   0   'False
      Bpic18          =   "frm_main.frx":1CA27
      Bname18         =   "Sort Ascending"
      Btype18         =   0
      Bwidth18        =   0
      Bchecked18      =   0   'False
      Bvalue18        =   0   'False
      Bpic19          =   "frm_main.frx":1CD79
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
      Bpic21          =   "frm_main.frx":1D0CB
      Bname21         =   "Chart"
      Btype21         =   0
      Bwidth21        =   0
      Bchecked21      =   0   'False
      Bvalue21        =   0   'False
      Bhighlight22    =   0   'False
      Bname22         =   "Button22"
      Btype22         =   0
      Bwidth22        =   75
      Bchecked22      =   0   'False
      Bvalue22        =   0   'False
      Bpic23          =   "frm_main.frx":1D41D
      Bname23         =   "Help"
      Btype23         =   1
      Bwidth23        =   0
      Bchecked23      =   0   'False
      Bvalue23        =   0   'False
      BackColor       =   16244941
      WindowColor     =   3
      Begin VistaSuitePro.OsenVistaComboBox OsenXPComboBox1 
         Height          =   285
         Left            =   7350
         TabIndex        =   3
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ComboStyle      =   1
         MAXROWS         =   5
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LLID            =   0
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         HoverSelection  =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFontColor =   0
         ASURC           =   0   'False
         TextColumn      =   0
         Required        =   0   'False
         Unicode         =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
   End
   Begin VistaSuitePro.OsenVistaHookMenu Hmenu 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   13
      Top             =   420
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   688
      BmpCount        =   26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCountMenu      =   9
      Bmp:1           =   "frm_main.frx":1D76F
      Key:1           =   "#mnu_File:1"
      Bmp:2           =   "frm_main.frx":1DB97
      Key:2           =   "#mnu_File:2"
      Bmp:3           =   "frm_main.frx":1DFBF
      Key:3           =   "#mnu_File:5"
      Bmp:4           =   "frm_main.frx":1E3E7
      Key:4           =   "#mnu_File:7"
      Bmp:5           =   "frm_main.frx":1E80F
      Key:5           =   "#mnu_File:11"
      Bmp:6           =   "frm_main.frx":1EC37
      Key:6           =   "#mnu_File:12"
      Bmp:7           =   "frm_main.frx":1F05F
      Key:7           =   "#mnu_send:1"
      Bmp:8           =   "frm_main.frx":1F487
      Key:8           =   "#mnu_send:2"
      Bmp:9           =   "frm_main.frx":1F8AF
      Key:9           =   "#mnu_Edit:1"
      Bmp:10          =   "frm_main.frx":1FCD7
      Key:10          =   "#mnu_Edit:2"
      Bmp:11          =   "frm_main.frx":200FF
      Key:11          =   "#mnu_Edit:4"
      Bmp:12          =   "frm_main.frx":20527
      Key:12          =   "#mnu_Edit:5"
      Bmp:13          =   "frm_main.frx":2094F
      Key:13          =   "#mnu_Edit:6"
      Bmp:14          =   "frm_main.frx":20D77
      Key:14          =   "#mnu_Fill:1"
      Bmp:15          =   "frm_main.frx":2119F
      Key:15          =   "#mnu_Fill:2"
      Bmp:16          =   "frm_main.frx":215C7
      Key:16          =   "#mnu_Edit:12"
      Bmp:17          =   "frm_main.frx":219EF
      Key:17          =   "#mnu_Edit:13"
      Bmp:18          =   "frm_main.frx":21E17
      Key:18          =   "#mnu_view:1"
      Bmp:19          =   "frm_main.frx":2223F
      Key:19          =   "#mnu_view:14"
      Bmp:20          =   "frm_main.frx":22667
      Key:20          =   "#mnu_view:9"
      Bmp:21          =   "frm_main.frx":22A8F
      Key:21          =   "#mnu_view:2"
      Bmp:22          =   "frm_main.frx":22EB7
      Key:22          =   "#mnu_File:9"
      Bmp:23          =   "frm_main.frx":232DF
      Key:23          =   "#mnu_Paste:2"
      Bmp:24          =   "frm_main.frx":23707
      Key:24          =   "#mnu_File:6"
      Bmp:25          =   "frm_main.frx":23B2F
      Key:25          =   "#mnu_File:15"
      Bmp:26          =   "frm_main.frx":23F57
      Key:26          =   "#mnu_File:17"
      XMenuA1         =   "&File "
      XMenuACS1       =   "f"
      XMenuC1         =   "mnuFile"
      XMenuE1         =   -1  'True
      XMenuH1         =   -1  'True
      XMenuA2         =   "Edit "
      XMenuACS2       =   ""
      XMenuC2         =   "mnuEdit"
      XMenuE2         =   -1  'True
      XMenuH2         =   0   'False
      XMenuA3         =   "View "
      XMenuACS3       =   ""
      XMenuC3         =   "mnuView"
      XMenuE3         =   -1  'True
      XMenuH3         =   0   'False
      XMenuA4         =   "Insert "
      XMenuACS4       =   ""
      XMenuC4         =   "mnuInsert"
      XMenuE4         =   -1  'True
      XMenuH4         =   0   'False
      XMenuA5         =   "Format "
      XMenuACS5       =   ""
      XMenuC5         =   "mnuFormat"
      XMenuE5         =   -1  'True
      XMenuH5         =   0   'False
      XMenuA6         =   "Tools "
      XMenuACS6       =   ""
      XMenuC6         =   "mnuTools"
      XMenuE6         =   -1  'True
      XMenuH6         =   0   'False
      XMenuA7         =   "Data "
      XMenuACS7       =   ""
      XMenuC7         =   "mnuData"
      XMenuE7         =   -1  'True
      XMenuH7         =   0   'False
      XMenuA8         =   "Window "
      XMenuACS8       =   ""
      XMenuC8         =   "mnuWindow"
      XMenuE8         =   -1  'True
      XMenuH8         =   0   'False
      XMenuA9         =   "Help "
      XMenuACS9       =   ""
      XMenuC9         =   "mnuHelp"
      XMenuE9         =   -1  'True
      XMenuH9         =   0   'False
   End
   Begin VistaSuitePro.OsenVistaForm OXP 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
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
      Caption         =   "Sample Menu,Toolbar and Statusbar usage"
      icon            =   "frm_main.frx":2437F
      BorderStyle     =   1
      UseDefaultTheme =   0   'False
      WindowColor     =   3
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CONCLUSION --> Slower"
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
      Index           =   1
      Left            =   5130
      TabIndex        =   4
      Top             =   2280
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONCLUSION --> Faster"
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
      Index           =   0
      Left            =   5130
      TabIndex        =   5
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnu_File 
         Caption         =   "&New"
         Index           =   1
      End
      Begin VB.Menu mnu_File 
         Caption         =   "&Open"
         Index           =   2
      End
      Begin VB.Menu mnu_File 
         Caption         =   "&Close"
         Index           =   3
      End
      Begin VB.Menu mnu_File 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnu_File 
         Caption         =   "&Save"
         Index           =   5
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Save &As"
         Index           =   6
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Search"
         Index           =   7
      End
      Begin VB.Menu mnu_File 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Page Setup ..."
         Index           =   9
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Print Area"
         Index           =   10
         Begin VB.Menu mnu_Print_Area 
            Caption         =   "Set Print Area"
            Index           =   1
         End
         Begin VB.Menu mnu_Print_Area 
            Caption         =   "Clear Print Area"
            Index           =   2
         End
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Preview"
         Index           =   11
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Print"
         Index           =   12
      End
      Begin VB.Menu mnu_File 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Send To"
         Index           =   14
         Begin VB.Menu mnu_send 
            Caption         =   "Mail Recipient (For review)"
            Index           =   1
         End
         Begin VB.Menu mnu_send 
            Caption         =   "Mail Recipient (As attachment)"
            Index           =   2
         End
         Begin VB.Menu mnu_send 
            Caption         =   "Routing Recipient"
            Index           =   3
         End
      End
      Begin VB.Menu mnu_File 
         Caption         =   "Properties"
         Index           =   15
      End
      Begin VB.Menu mnu_File 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnu_File 
         Caption         =   "E&xit"
         Index           =   17
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnu_Edit 
         Caption         =   "Undo"
         Index           =   1
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Redo"
         Index           =   2
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Cut"
         Index           =   4
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Copy"
         Index           =   5
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Paste"
         Index           =   6
         Begin VB.Menu mnu_Paste 
            Caption         =   "Formulas"
            Index           =   1
         End
         Begin VB.Menu mnu_Paste 
            Caption         =   "Values"
            Index           =   2
         End
         Begin VB.Menu mnu_Paste 
            Caption         =   "No Borders"
            Index           =   3
         End
         Begin VB.Menu mnu_Paste 
            Caption         =   "Transpose"
            Index           =   4
         End
         Begin VB.Menu mnu_Paste 
            Caption         =   "Paste Link"
            Index           =   5
         End
         Begin VB.Menu mnu_Paste 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnu_Paste 
            Caption         =   "Paste Special"
            Index           =   7
         End
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Fill"
         Index           =   8
         Begin VB.Menu mnu_Fill 
            Caption         =   "Down"
            Index           =   1
         End
         Begin VB.Menu mnu_Fill 
            Caption         =   "Right"
            Index           =   2
         End
         Begin VB.Menu mnu_Fill 
            Caption         =   "Up"
            Index           =   3
         End
         Begin VB.Menu mnu_Fill 
            Caption         =   "Left"
            Index           =   4
         End
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Clear"
         Index           =   9
         Begin VB.Menu mnu_Clear 
            Caption         =   "All"
            Index           =   1
         End
         Begin VB.Menu mnu_Clear 
            Caption         =   "Formats"
            Index           =   2
         End
         Begin VB.Menu mnu_Clear 
            Caption         =   "Contants"
            Index           =   3
         End
         Begin VB.Menu mnu_Clear 
            Caption         =   "Comments"
            Index           =   4
         End
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Delete"
         Index           =   10
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Find"
         Index           =   12
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Replace"
         Index           =   13
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Go To"
         Index           =   14
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "Link Object"
         Index           =   16
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Visible         =   0   'False
      Begin VB.Menu mnu_view 
         Caption         =   "Normal"
         Index           =   1
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Page Break Preview"
         Index           =   2
      End
      Begin VB.Menu mnu_view 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Taskpane"
         Index           =   4
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Toolbars"
         Index           =   5
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Standard"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Formatting"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Borders"
            Index           =   3
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Charts"
            Index           =   4
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Control Toolbox"
            Index           =   5
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Drawing"
            Index           =   6
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Forms"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Picture"
            Index           =   8
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Text to Speech"
            Checked         =   -1  'True
            Index           =   9
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Visual Basic"
            Checked         =   -1  'True
            Index           =   10
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnu_Toolbar 
            Caption         =   "Customize"
            Index           =   12
         End
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Formula bar"
         Checked         =   -1  'True
         Index           =   6
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Status bar"
         Checked         =   -1  'True
         Index           =   7
      End
      Begin VB.Menu mnu_view 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Header and Footer"
         Index           =   9
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Comments"
         Index           =   11
      End
      Begin VB.Menu mnu_view 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Custom views"
         Index           =   13
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Full Screen"
         Index           =   14
      End
      Begin VB.Menu mnu_view 
         Caption         =   "Zoom..."
         Index           =   15
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "Insert"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenXPSuite 2006 - OsenXPHookMenu sample                             *'
'*  Copyright (c) 2006 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  This file is part of the OsenXPSuite 2006 sample applications.       *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenXPSuite 2006 documentation, and is provided "as is", without     *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'


Private Sub CmdAdd1_Click()
' with lockupdate --> very fast
Dim I As Long
Dim timeX As Long
    ' get Start
    timeX = GTick
    
    With Me.OsenXPListBox1
        .LockUpdate = True
        For I = 1 To 10000
            .AddItem vbTab & "OsenXpListbox" & I & vbTab & .ListCount + 1, , , , RndImage, RndImage
        Next
        .LockUpdate = False
    End With
    
    MsgBoxGT "10.000 items on " & GTick - timeX & " ms", vbExclamation, "Conclusion", 2
    
    
End Sub

Private Sub CmdAdd2_Click()

' without lockupdate --> slower
Dim I As Long
Dim timeX As Long

    ' get Start
    timeX = GTick
    
    With Me.OsenXPListBox1
        
        For I = 1 To 1000
            .AddItem vbTab & "OsenXpListbox" & I & vbTab & .ListCount + 1, , , , RndImage, RndImage
        Next
    
    End With
    
    MsgBoxXP "1000 items on " & GTick - timeX & " ms", vbExclamation, "Conclusion"
    

End Sub

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OXP.Init Me
    
    With Me.OsenXPComboBox1
        .AddItem "400%"
        .AddItem "200%"
        .AddItem "150%"
        .AddItem "125%"
        .AddItem "100%"
        .AddItem "75%"
        .AddItem "50%"
        .AddItem "25%"
    End With
    

    
End Sub


Private Sub mnuAbout_Click()
    ' show about dialog
    OXP.About
End Sub

Private Sub OsenXPButton1_Click()
' copy items from list1 to list2
    OsenXPListBox2.SetListItems OsenXPListBox1.GetListItems
End Sub

Private Sub OsenXPButton3_Click()

    'clear listbox
    Me.OsenXPListBox1.Clear
    
End Sub

Private Sub r2l_Click()
'TBar.RightToLeft = r2l.Value
End Sub

Private Sub TBar_ButtonClick(Index As Integer, sText As String)
    Debug.Print "ButtonCLicked: "; ; Index; ; Timer
End Sub

Private Sub TBar_Highlight(Index As Integer, sText As String)

    ' highlight toolbarbutton
    Sbar.PanelCaption(3) = "Index: " & Index & "    sText: " & sText
    
End Sub

Private Sub TBar_PopUpMainMenu(Index As Integer, sText As String, X As Long, Y As Long)
        Debug.Print "popupmenu: "; ; Index; ; Timer

    ' popup menu
    If sText = "Paste" Then
        Me.PopupMenu mnu_Edit(6), , X, Y
    End If
    
End Sub

Private Function RndImage() As Integer
    RndImage = CInt((Rnd(1) * 44)) + 1
End Function





















































































































































































































