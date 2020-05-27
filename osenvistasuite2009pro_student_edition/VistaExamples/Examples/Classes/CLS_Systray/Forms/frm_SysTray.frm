VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_SysTray 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Systray Icon Demo"
   ClientHeight    =   7020
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_SysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaHookMenu OsenXPHookMenu1 
      Height          =   375
      Left            =   1770
      TabIndex        =   18
      Top             =   4590
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   688
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GripperLeft     =   8
      MCountMenu      =   1
      XMenuA1         =   "System "
      XMenuACS1       =   ""
      XMenuC1         =   "mnuSystem"
      XMenuE1         =   -1  'True
      XMenuH1         =   0   'False
   End
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   1650
      Top             =   3300
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   24108
      Images          =   "frm_SysTray.frx":058A
      Version         =   131072
      KeyCount        =   21
      Keys            =   "????????????????????????????????????????ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
      Height          =   5175
      Left            =   4110
      TabIndex        =   5
      Top             =   1170
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   9128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BorderColor     =   14396553
      Appearance      =   1
      BinaryImage     =   "frm_SysTray.frx":63D6
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaButton cmdClear 
         Height          =   405
         Left            =   210
         TabIndex        =   17
         Top             =   4590
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   714
         Caption         =   "&Clear event history"
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "frm_SysTray.frx":63EE
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_SysTray.frx":640A
         BinaryImageOver =   "frm_SysTray.frx":6422
      End
      Begin VistaSuitePro.OsenVistaButton cmdHide 
         Height          =   405
         Left            =   2220
         TabIndex        =   16
         Top             =   4080
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
         Caption         =   "Hide Popup Balloon"
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "frm_SysTray.frx":643A
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_SysTray.frx":6456
         BinaryImageOver =   "frm_SysTray.frx":646E
      End
      Begin VistaSuitePro.OsenVistaButton cmdShow 
         Height          =   375
         Left            =   210
         TabIndex        =   15
         Top             =   4080
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         Caption         =   "&Display Popup Balloon"
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "frm_SysTray.frx":6486
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_SysTray.frx":64A2
         BinaryImageOver =   "frm_SysTray.frx":64BA
      End
      Begin VistaSuitePro.OsenVistaComboBox cboMsgIcon 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   3540
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "xpTrayIcon"
         ComboStyle      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSGL            =   -1  'True
         LLID            =   4
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         HoverSelection  =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Required        =   -1  'True
         Unicode         =   0   'False
         DataList        =   "xpNone|xpInformation|xpWarning|xpError|xpTrayIcon"
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
      Begin VistaSuitePro.OsenVistaTextBox txtBody 
         Height          =   1215
         Left            =   210
         TabIndex        =   12
         Top             =   2160
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2143
         Text            =   $"frm_SysTray.frx":64D2
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MultiLine       =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Required        =   -1  'True
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaTextBox txtTitle 
         Height          =   345
         Left            =   210
         TabIndex        =   10
         Top             =   1710
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   609
         Text            =   "OsenVistaSuite 2008 Professional Edition"
         BackColor       =   16777215
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Required        =   -1  'True
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaTextBox txtToolTip 
         Height          =   345
         Left            =   1440
         TabIndex        =   8
         Top             =   930
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         Text            =   "Test Tooltip"
         BackColor       =   16777215
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
         ButtonCaption   =   ""
         ButtonPicture   =   "frm_SysTray.frx":6591
         ButtonVisible   =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaComboBox cboIcon 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ComboStyle      =   1
         DisplayPicture  =   -1  'True
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSGL            =   -1  'True
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         HoverSelection  =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   540
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Change Icon:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   300
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   960
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Change Tooltip:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   300
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   1410
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Message Title:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   300
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   3570
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Display Icon:"
         ForeColor       =   0
         BackStyle       =   0
      End
   End
   Begin VistaSuitePro.OsenVistaButton cmdRemove 
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   600
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      Caption         =   "&Remove Systray"
      Enabled         =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frm_SysTray.frx":6B2B
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "frm_SysTray.frx":6B47
      BinaryImageOver =   "frm_SysTray.frx":6B5F
   End
   Begin VistaSuitePro.OsenVistaButton cmdCreate 
      Height          =   375
      Left            =   4650
      TabIndex        =   3
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "&Create Systray"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frm_SysTray.frx":6B77
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "frm_SysTray.frx":6B93
      BinaryImageOver =   "frm_SysTray.frx":6BAB
   End
   Begin VistaSuitePro.OsenVistaStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   6465
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   979
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   3
      HaveXPForm      =   -1  'True
      WindowColor     =   3
      PWidth1         =   132
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "OsenVistaSuite 2008"
      pTextAlignment1 =   0
      pTextBold1      =   -1  'True
      PanelPicture1   =   "frm_SysTray.frx":6BC3
      PanelPicAlignment1=   0
      PWidth2         =   120
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Professional Edition"
      pTextAlignment2 =   0
      PanelPicture2   =   "frm_SysTray.frx":6BDF
      PanelPicAlignment2=   0
      PWidth3         =   120
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Version 2.0.0.19"
      pTextAlignment3 =   0
      pTextBold3      =   -1  'True
      PanelPicture3   =   "frm_SysTray.frx":6BFB
      PanelPicAlignment3=   0
      GradientColor1  =   15382160
      GradientColor2  =   12752244
   End
   Begin VistaSuitePro.OsenVistaListBox lstEvents 
      Height          =   5775
      Left            =   150
      TabIndex        =   1
      Top             =   570
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   10186
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   0
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      AllowEdit       =   0   'False
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BinaryImage     =   "frm_SysTray.frx":6C17
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Systray Icon Demo"
      TitleTop        =   7
      icon            =   "frm_SysTray.frx":6C2F
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      WindowColor     =   3
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "System"
      Visible         =   0   'False
      Begin VB.Menu mnutest 
         Caption         =   "Menu1"
         Index           =   1
      End
      Begin VB.Menu mnutest 
         Caption         =   "Menu2"
         Index           =   2
      End
      Begin VB.Menu mnutest 
         Caption         =   "Menu3"
         Index           =   3
      End
      Begin VB.Menu mnutest 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnutest 
         Caption         =   "Exit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frm_SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenVistaSuite 2008 - CLS_SysTray sample                             *'
'*  Copyright (c) 2008 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  [frm_systray.frm]                                                    *'
'*                                                                       *'
'*  This file is part of the OsenVistaSuite 2008 sample applications.    *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenVistaSuite 2008 documentation, and is provided "as is", without  *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

' Declare variable ...
Private WithEvents SysTray As CLS_SysTray
Attribute SysTray.VB_VarHelpID = -1

Private Sub Form_Load()

    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize event)
    Me.OsenXPForm1.Init Me
    
    ' Create new instance for CLS_SysTray
    Set SysTray = New CLS_SysTray
    
    
    cboIcon.ImageListName = "Myimagelist1"
    
    'Enter picture/icon from imagelist to cboicon
    cboIcon.DisplayIconFormImageList
    
    cboMsgIcon.ListIndex = 1
    
End Sub

Private Sub cboIcon_Click()
    
    ' Change systray icon ...
    Set SysTray.TrayIcon = MyImageList1.ItemPicture(cboIcon.ListIndex + 1)
    
    ' Display message (using ballon message)
    SysTray.BalloonShow "The trayicon has been changed", "Trayicon was changed successfull.", xpInformation
    
    
    lstEvents.AddItem "Trayicon was successfull changed"

End Sub

Private Sub cmdClear_Click()

    ' CLear all item from lstevents
    lstEvents.Clear
    
End Sub

Private Sub cmdCreate_Click()
    
    ' Systray icon initializing ...
    ' Please make sure don't use form.hwnd (me.hwnd) for entering pHwnd references
    ' for that reason, I used the cmdcreate.hwnd
    SysTray.Create "OsenXPSuite 2006 Enterprise Edition", cmdCreate.hWnd, OsenXPForm1.Icon
    
    ' Display ballon message
    SysTray.BalloonShow "Thank you for using OsenXPSuite 2006.", "Welcome to ....", xpTrayIcon
    
    ' change button status
    cmdCreate.Enabled = False
    cmdRemove.Enabled = True
    OsenXPFrame1.Enabled = True
    
End Sub

Private Sub cmdHide_Click()

    ' Close the ballon message
    SysTray.BalloonClose
    
End Sub

Private Sub cmdRemove_Click()
    
    ' Remove the icon from the tray icon
    SysTray.Remove
    
    ' change button status/position
    cmdCreate.Enabled = 1
    cmdRemove.Enabled = 0
    OsenXPFrame1.Enabled = 0

End Sub

Private Sub cmdShow_Click()

    ' Display ballon message
    SysTray.BalloonShow txtBody.Text, txtTitle.Text, cboMsgIcon.ListIndex
    
End Sub

Private Sub mnutest_Click(Index As Integer)
    
    ' Testing only ...
    If Index = 5 Then
        Unload Me
    Else
        ' Give a respon...
        MsgBoxGT mnutest(Index).Caption & " clicked", vbInformation, "Information"
    End If
    
End Sub

Private Sub SysTray_BalloonClick()
    
    ' This event occurs when user pressed ballon message
    lstEvents.AddItem "BaloonClick"
    
End Sub

Private Sub SysTray_BalloonClose()

    ' This event occurs when the ballon message closed
    lstEvents.AddItem "BaloonClose"
    
End Sub

Private Sub SysTray_BalloonHide()

    ' This event occurs when the ballon message hidden
    lstEvents.AddItem "BaloonHide"

End Sub

Private Sub SysTray_BalloonShow()

    ' This event occurs when the ballon message appear
    lstEvents.AddItem "BaloonShow"

End Sub

Private Sub SysTray_LeftButtonClick()

    ' This event occurs when the ballon message has been clicked with left mouse button
    lstEvents.AddItem "LeftButtonClick"

End Sub

Private Sub SysTray_LeftButtonDblClick()

    ' This event occurs when the ballon message has been dblclicked with left mouse button
    lstEvents.AddItem "LeftButtonDblClick"

End Sub

Private Sub SysTray_MouseOver()

    ' This event occurs when user movement of mouse at balloon message
    lstEvents.AddItem "MouseOver"
    
End Sub

Private Sub SysTray_RightButtonClick()

    ' This event occurs when the ballon message has been clicked with right mouse button
    lstEvents.AddItem "RightButtonClick"
    
    ' Now, open a sample popup menu
    PopupMenu mnuSystem
    
End Sub

Private Sub SysTray_RightButtonDblClick()

    ' This event occurs when the ballon message has been dblclicked with right mouse button
    lstEvents.AddItem "RightButtonDblClick"

End Sub

Private Sub txtToolTip_ButtonClick()
    
    ' Change the systray tooltiptext
    SysTray.ToolTip = txtToolTip.Text
    
    lstEvents.AddItem "Tooltip was successfull changed"
    
    ' Display message
    SysTray.BalloonShow "Tooltip has been changed", "Tooltip was changed successfull.", xpInformation

End Sub























