VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "ImySQL5_Connection Sample"
   ClientHeight    =   7950
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   8010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
      Height          =   1395
      Left            =   210
      TabIndex        =   13
      Top             =   1410
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   2461
      Caption         =   "Connection properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      DropDownButton  =   -1  'True
      BinaryImage     =   "Form1.frx":038A
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaTextBox txtPort 
         Height          =   315
         Left            =   3060
         TabIndex        =   2
         Top             =   480
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Text            =   "3306"
         Alignment       =   2
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
         Value           =   3306
         BorderColor     =   8370596
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaTextBox txtHost 
         Height          =   330
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         Text            =   "localhost"
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
         BorderColor     =   8370596
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         LabelBackColor  =   16767935
         LabelCaption    =   "Server Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtDBName 
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   582
         Text            =   "nwind2008"
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
         BorderColor     =   8370596
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         LabelBackColor  =   16767935
         LabelCaption    =   "Database Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtUID 
         Height          =   330
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   582
         Text            =   "root"
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
         BorderColor     =   8370596
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         LabelBackColor  =   16767935
         LabelCaption    =   "User Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtPwd 
         Height          =   330
         Left            =   4440
         TabIndex        =   4
         Top             =   900
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   582
         Text            =   "vb"
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
         BorderColor     =   8370596
         PasswordChar    =   "*"
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         LabelBackColor  =   16767935
         LabelCaption    =   "Password:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
   End
   Begin VistaSuitePro.OsenVistaOptionButton optRestore 
      Height          =   285
      Left            =   3210
      TabIndex        =   20
      Top             =   2370
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BackColor       =   16767935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      Caption         =   "Restore"
      Style           =   1
   End
   Begin VistaSuitePro.OsenVistaOptionButton OptBackUp 
      Height          =   255
      Left            =   2100
      TabIndex        =   19
      Top             =   2370
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   450
      BackColor       =   16767935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      Value           =   -1  'True
      Caption         =   "BackUp"
      Style           =   1
   End
   Begin VistaSuitePro.OsenVistaButton cmdBackUp 
      Height          =   315
      Left            =   6150
      TabIndex        =   18
      Top             =   2340
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Caption         =   "&BackUp Database"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "Form1.frx":03A2
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":03BE
      BinaryImageOver =   "Form1.frx":03D6
   End
   Begin VistaSuitePro.OsenVistaTextBox txtFile 
      Height          =   315
      Left            =   2130
      TabIndex        =   16
      Top             =   1950
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   556
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
      BorderColor     =   8370596
      ButtonCaption   =   ""
      ButtonPicture   =   "Form1.frx":03EE
      ButtonVisible   =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaStatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   15
      Top             =   7425
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   926
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   12632256
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   1
      HaveXPForm      =   -1  'True
      WindowColor     =   3
      PWidth1         =   1080
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   ""
      pTextAlignment1 =   0
      PanelPicture1   =   "Form1.frx":0788
      PanelPicAlignment1=   0
      GradientColor1  =   15382160
      GradientColor2  =   12752244
   End
   Begin VistaSuitePro.OsenVistaListBox LstData 
      Height          =   2745
      Left            =   210
      TabIndex        =   14
      Top             =   4560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4842
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   0
      BackSelected    =   7381139
      BackSelectedG1  =   16777215
      BackSelectedG2  =   8632490
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
      HeaderCaption   =   "OsenXPListBox1"
      TransparencyLevel=   22
      ReadOnDemand    =   -1  'True
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientTop=   8569007
      HeaderGradientBottom=   4487779
      BinaryImage     =   "Form1.frx":07A4
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaTextBox txtSQL 
      Height          =   1155
      Left            =   210
      TabIndex        =   11
      Top             =   3360
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   2037
      Text            =   "call vw1_42400();"
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
      BorderColor     =   8421504
      MultiLine       =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaButton CmdExec 
      Height          =   345
      Left            =   6600
      TabIndex        =   10
      Top             =   2910
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Caption         =   "&Execute"
      Enabled         =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "Form1.frx":07BC
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":07D8
      BinaryImageOver =   "Form1.frx":07F0
   End
   Begin VistaSuitePro.OsenVistaButton CmdProc 
      Height          =   345
      Left            =   5100
      TabIndex        =   9
      Top             =   2910
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      Caption         =   "Show Processlist"
      Enabled         =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "Form1.frx":0808
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":0824
      BinaryImageOver =   "Form1.frx":083C
   End
   Begin VistaSuitePro.OsenVistaButton cmdVar 
      Height          =   345
      Left            =   3600
      TabIndex        =   8
      Top             =   2910
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      Caption         =   "Show Variables"
      Enabled         =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "Form1.frx":0854
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":0870
      BinaryImageOver =   "Form1.frx":0888
   End
   Begin VistaSuitePro.OsenVistaButton cmdDisconnect 
      Height          =   345
      Left            =   1320
      TabIndex        =   7
      Top             =   2910
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "Form1.frx":08A0
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":08BC
      BinaryImageOver =   "Form1.frx":08D4
   End
   Begin VistaSuitePro.OsenVistaButton cmdConnect 
      Height          =   345
      Left            =   210
      TabIndex        =   6
      Top             =   2910
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      Caption         =   "&Connect"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "Form1.frx":08EC
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":0908
      BinaryImageOver =   "Form1.frx":0920
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   12
      Top             =   420
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   1561
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
      Picture         =   "Form1.frx":0938
      BorderColor     =   8632490
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   16310477
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The Following is example of usage ImySQL5_Connection"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Class Name: IMySQL5_Connection"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "Form1.frx":158A
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
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
      Caption         =   "ImySQL5_Connection Sample"
      TitleTop        =   7
      icon            =   "Form1.frx":15A2
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      WindowColor     =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Script Filename (Target):"
      Height          =   195
      Left            =   270
      TabIndex        =   17
      Top             =   1980
      Width           =   1725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************'
'*  OsenVistaSuite 2008 - IMySQL5_Connection sample                      *'
'*  Copyright (c) 2008 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  [Form1.frm]                                                          *'
'*                                                                       *'
'*  This file is part of the OsenVistaSuite 2008 sample applications.    *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenVistaSuite 2008 documentation, and is provided "as is", without  *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

Option Explicit

' Declare variable
Private myCN            As IMySQL5_Connection
Private lxTime          As Long

Private Sub cmdBackUp_Click()
    
    ' check connection status
    If myCN.State Then
        
        ' Check the sql script filename ...
        If Len(txtFile.Text) Then
            
            If OptBackUp.Value Then
            
                    
                    ' Now it's the moment to backup database process
                    myCN.BackUpDatabase txtFile, lxTime
                    
                    MsgBoxGT "Backup database finished" & vbCrLf & _
                             lxTime & " ms taken", vbInformation, "Backup database"
                    
                
            Else
            
                ' Now it's the moment to restore database process
                myCN.Restore txtFile, , lxTime
                
                MsgBoxGT "Restore database finished" & vbCrLf & _
                         lxTime & " ms taken", vbInformation, "Restore database"
            
            End If
            
            
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize event)
    Me.OsenXPForm1.Init Me
    
    ' Now, you can make a new instance of CLS_Encrypt
    Set myCN = New IMySQL5_Connection

End Sub

Private Sub cmdConnect_Click()
    
    ' Try to open connection to specified server address
    myCN.OpenConnection txtHost, txtUID, txtPwd, txtPort.Value, txtDBName
    
    If myCN.State Then
    
        MsgBoxGT "Connection to MySQL server successfull." & vbCrLf & _
                 "Connection Id: " & myCN.ConnectionID & vbCrLf & _
                "Client Version: " & myCN.ClientVersion & vbCrLf & _
                "Server version: " & myCN.ServerVersion, vbInformation, "IMySQL5_Connection Sample"
        
        cmdDisconnect.Enabled = True
        cmdVar.Enabled = True
        CmdProc.Enabled = True
        CmdExec.Enabled = True
        cmdConnect.Enabled = False
                
    Else
         MsgBoxGT "Connection failed.", vbCritical, "CLS_MySQL Sample", 3
    End If
    
    
End Sub


Private Sub cmdDisconnect_Click()

    ' Close opened connection ...
    myCN.CloseConnection
    
    cmdDisconnect.Enabled = False
    cmdVar.Enabled = False
    CmdProc.Enabled = False
    CmdExec.Enabled = False
    cmdConnect.Enabled = True

End Sub

Private Sub CmdExec_Click()
    On Error GoTo Err_Msg
    
    ' Clear the previous data
    LstData.Clear True, True
    lxTime = GTick
    
    ' Execute the sql script in txtSQL, and display the resultset at the lstdata
    LstData.InsertItemByRecordset myCN.Recordset(txtSQL), , , True, , AutoColumnWidthEx:=True
    
    lxTime = GTick - lxTime
    
    ' Display message in StatusBar
    sBar.PanelCaption(1) = LstData.ListCount & " row(s) returned [" & lxTime & " ms taken]"
    
    Debug.Print myCN.ConnectionID
    
    
    Exit Sub
Err_Msg:
    MsgBoxGT Err.Description, vbExclamation, "Error"
    
End Sub

Private Sub CmdProc_Click()

    ' Show processlist
    LstData.InsertItemByRecordset myCN.Recordset("show processlist"), , , True, lxTime, AutoColumnWidthEx:=True
        
    ' Display message in StatusBar
    sBar.PanelCaption(1) = LstData.ListCount & " row(s) returned [" & lxTime & " ms taken]"
    
End Sub

Private Sub cmdVar_Click()

    '  Show variables
    LstData.InsertItemByRecordset myCN.Recordset("show variables"), , , True, lxTime, AutoColumnWidthEx:=True
    
    ' Display message in StatusBar
    sBar.PanelCaption(1) = LstData.ListCount & " row(s) returned [" & lxTime & " ms taken]"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' CLean Up
    Set myCN = Nothing
    
End Sub

Private Sub OptBackUp_Click()

    ' Change the cmdBackup caption
    cmdBackUp.Caption = "&BackUp database"

End Sub

Private Sub optRestore_Click()
    
    ' Change the cmdBackup caption
    cmdBackUp.Caption = "&Restore database"
    
End Sub

Private Sub txtFile_ButtonClick()
    
    If OptBackUp.Value Then
    
        ' Show save dialog and get the filename
        txtFile.ShowSaveDialog "Backup database as SQL statement", "SQL|*.sql", "sql"
    
    Else
    
        ' Show save dialog and get the filename
        txtFile.ShowOpenDialog "Restore database from File", "SQL|*.sql", "sql"
        
    End If
        
End Sub













