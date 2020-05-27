VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_user_mgmt 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "User Management"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_user_mgmt2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyADODC MyADODC1 
      Height          =   435
      Left            =   1950
      TabIndex        =   21
      Top             =   7890
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   767
      GradientColor1  =   16777215
      GradientColor2  =   12752244
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
      BeginProperty FontButton {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   6
      Gradient        =   -1  'True
      AutoConfirmBeforeDelete=   -1  'True
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaTextBox txtdata 
      Height          =   435
      Index           =   3
      Left            =   1980
      TabIndex        =   20
      Top             =   7380
      Visible         =   0   'False
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   767
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "frm_user_mgmt2.frx":058A
      Top             =   6810
      Width           =   9015
   End
   Begin VistaSuitePro.OsenVistaHookMenu OsenXPHookMenu1 
      Height          =   375
      Left            =   6510
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      XMenuA1         =   "Check All "
      XMenuACS1       =   ""
      XMenuC1         =   "mnuCheckAll"
      XMenuE1         =   -1  'True
      XMenuH1         =   0   'False
   End
   Begin VistaSuitePro.OsenVistaTreeView tvwMain 
      Height          =   5355
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   9446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectedBackColor=   13603685
      SelectedColor   =   16777215
      CheckBoxes      =   -1  'True
      BorderStyle     =   0
      HeaderCaption   =   "Northwind Traders"
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowHeader      =   -1  'True
      GradientColor2  =   12632256
      HeaderGradientAllow=   0   'False
   End
   Begin VistaSuitePro.OsenVistaTab MyTab 
      Height          =   5385
      Left            =   3030
      TabIndex        =   2
      Top             =   1380
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9499
      TabHeight       =   22
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FrameColor      =   12164479
      MaskColor       =   16711935
      SelectedTab     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   1
      BackColorParent =   16767935
      TabWidth1       =   61
      TabText1        =   "User Info"
      TabEnabled1     =   -1  'True
      TabVisible1     =   0   'False
      TabBack21       =   16777215
      TabCountCtls1   =   11
      TabNo1CtlID1    =   "LblGetNodes"
      TabNo1CtlIX1    =   1
      TabNo1CtlIT1    =   -1  'True
      TabNo1CtlID2    =   "cmdOK"
      TabNo1CtlIX2    =   1
      TabNo1CtlIT2    =   -1  'True
      TabNo1CtlID3    =   "cmdCancel"
      TabNo1CtlIX3    =   1
      TabNo1CtlIT3    =   -1  'True
      TabNo1CtlID4    =   "txtID(0)"
      TabNo1CtlIX4    =   1
      TabNo1CtlIT4    =   -1  'True
      TabNo1CtlID5    =   "OsenXPLabel1(0)"
      TabNo1CtlIX5    =   1
      TabNo1CtlIT5    =   -1  'True
      TabNo1CtlID6    =   "Picture2"
      TabNo1CtlIX6    =   1
      TabNo1CtlIT6    =   -1  'True
      TabNo1CtlID7    =   "OsenXPLabel1(1)"
      TabNo1CtlIX7    =   1
      TabNo1CtlIT7    =   -1  'True
      TabNo1CtlID8    =   "TxtName(1)"
      TabNo1CtlIX8    =   1
      TabNo1CtlIT8    =   -1  'True
      TabNo1CtlID9    =   "OsenXPLabel1(2)"
      TabNo1CtlIX9    =   1
      TabNo1CtlIT9    =   -1  'True
      TabNo1CtlID10   =   "TxtPwd(2)"
      TabNo1CtlIX10   =   1
      TabNo1CtlIT10   =   -1  'True
      TabNo1CtlID11   =   "OsenXPLabel5"
      TabNo1CtlIX11   =   1
      TabNo1CtlIT11   =   -1  'True
      Begin VistaSuitePro.OsenVistaLabel LblGetNodes 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   2130
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ForeColorOver   =   12582912
         ForeColorDown   =   33023
         Caption         =   "Get All Selected Node(s)"
         FontUnderline   =   -1  'True
         GradientOnOver  =   -1  'True
         GradientColor1  =   16777215
         GradientColor2  =   16777215
         ForeColor       =   0
         GradientColor3  =   16777215
         BorderColor     =   16777215
         AutoSize        =   0   'False
      End
      Begin VistaSuitePro.OsenVistaButton cmdOK 
         Height          =   285
         Left            =   1740
         TabIndex        =   6
         Top             =   1680
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Caption         =   "&OK"
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
         MPTR            =   99
         MICON           =   "frm_user_mgmt2.frx":07A6
         UMCOL           =   -1  'True
         XPBlendPicture  =   -1  'True
         Style           =   1
         BinaryImageNormal=   "frm_user_mgmt2.frx":07C2
         BinaryImageOver =   "frm_user_mgmt2.frx":07DA
      End
      Begin VistaSuitePro.OsenVistaButton cmdCancel 
         Height          =   285
         Left            =   2910
         TabIndex        =   1
         Top             =   1680
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Caption         =   "&Cancel"
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
         MPTR            =   99
         MICON           =   "frm_user_mgmt2.frx":07F2
         UMCOL           =   -1  'True
         XPBlendPicture  =   -1  'True
         Style           =   1
         BinaryImageNormal=   "frm_user_mgmt2.frx":080E
         BinaryImageOver =   "frm_user_mgmt2.frx":0826
      End
      Begin VistaSuitePro.OsenVistaTextBox txtdata 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   570
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   570
         Width           =   720
         _ExtentX        =   1270
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
         Caption         =   "User ID:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   930
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "User Name:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox txtdata 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   930
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   1290
         Width           =   870
         _ExtentX        =   1535
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
         Caption         =   "Password:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox txtdata 
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   1290
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PasswordChar    =   "•"
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VistaSuitePro.OsenVistaLabel LblDetail 
         Height          =   315
         Left            =   2280
         TabIndex        =   16
         Top             =   2130
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ForeColorOver   =   12582912
         ForeColorDown   =   33023
         Caption         =   "&Detail Privileges >>"
         FontUnderline   =   -1  'True
         GradientOnOver  =   -1  'True
         GradientColor1  =   16777215
         GradientColor2  =   16777215
         ForeColor       =   0
         GradientColor3  =   16777215
         BorderColor     =   16777215
         AutoSize        =   0   'False
      End
      Begin VistaSuitePro.OsenVistaTreeView tvwTest 
         Height          =   2745
         Left            =   150
         TabIndex        =   17
         Top             =   2520
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   4842
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SelectedBackColor=   13603685
         SelectedColor   =   16777215
         BorderStyle     =   0
         HeaderCaption   =   "Northwind Traders  >> User Privileges"
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowHeader      =   -1  'True
         GradientColor2  =   15779735
         HeaderGradientAllow=   0   'False
      End
      Begin VistaSuitePro.OsenVistaPicture picX 
         Height          =   1995
         Left            =   4050
         TabIndex        =   22
         Top             =   420
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   3519
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
         FieldName       =   "photo"
         BorderColor     =   14854529
         GradientColor2  =   16310477
         BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BinaryImage     =   "frm_user_mgmt2.frx":083E
      End
   End
   Begin VistaSuitePro.OsenVistaListBox lvwPrivileges 
      Height          =   2745
      Left            =   120
      TabIndex        =   15
      Top             =   3990
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
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
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      ItemTextLeft    =   20
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      HeaderFormatString=   "Key;70;0;0;|Caption;250;0;0;|Add;50;1;1;|Edit;50;1;1;|Delete;50;1;1;|Preview;57;1;1;|Export To Excel;60;1;1;"
      Columns         =   7
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      AlternateRowColors=   -1  'True
      MaxAllColumnWidth=   587
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ASURC           =   -1  'True
      IMGLIST         =   ""
      ForeColorSelected=   16576
      AllowSortItem   =   0   'False
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BinaryImage     =   "frm_user_mgmt2.frx":0856
      WindowColor     =   3
   End
   Begin VistaSuitePro.MyImageList SmallIcons 
      Left            =   2790
      Top             =   5940
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   59696
      Images          =   "frm_user_mgmt2.frx":086E
      Version         =   131072
      KeyCount        =   52
      Keys            =   "???????????????????????????????????????????????????ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   614
      TabIndex        =   8
      Top             =   420
      Width           =   9210
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel4 
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         Caption         =   "Add/Edit User ..."
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel3 
         Height          =   285
         Left            =   630
         TabIndex        =   13
         Top             =   420
         Width           =   4980
         _ExtentX        =   8784
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
         Caption         =   "This dialog allow you to create a new user and set global privilege(s)."
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   8160
         Picture         =   "frm_user_mgmt2.frx":F1BE
         Top             =   30
         Width           =   720
      End
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
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
      Caption         =   "User Management"
      TitleTop        =   7
      icon            =   "frm_user_mgmt2.frx":10088
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      WindowColor     =   3
   End
   Begin VB.Menu mnuCheckAll 
      Caption         =   "Check All"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Child_All 
         Caption         =   "&Select All"
         Index           =   1
      End
      Begin VB.Menu Mnu_Child_All 
         Caption         =   "&Deselect All"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm_user_mgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrValidation   As String
Private Flags   As Integer

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    ' On Error Resume Next
    StrValidation = lvwPrivileges.GetPrivileges
    txtData(3).Text = StrValidation
    MyADODC1.SendAction 8
    
    Unload Me

End Sub

Private Sub Form_Load()
    On Error Resume Next
    ' Make sure OsenXPForm Work Fine with this method
    Me.OsenXPForm1.Init Me

    ' Draw gradient Color for Picture1
'    DrawGradient4Pic Picture1

    ' Prepared Query
    If IsNew Then
        MyADODC1.OpenMySQLTable "users", "userid", MyCN, "where userid='-1'", txtData, picX
        MyADODC1.SendAction 5
    Else
        MyADODC1.OpenMySQLTable "users", "userid", MyCN, "where userid='" & KeyValue & "'", txtData, picX
        MyADODC1.SendAction 6
    End If

    FillData
    
    ' Prepared and set Imagelist handle
    tvwMain.ImageList = SmallIcons.hIml
    tvwTest.ImageList = SmallIcons.hIml
    lvwPrivileges.SmallIcons = SmallIcons.hIml

    ' Insert Nodes into tvwMain
    CreateNode

End Sub

' Purpose : Insert Node item by recordset >> From table "Nodes" <<
Public Sub CreateNode()
    On Error Resume Next

    With tvwMain
        ' Clean Up
        .Clear

        .LockUpdate = True

        ' Prepared recordset ...
        Set .TableNodes = GetRST("select * from nodes")

        ' Create node by recordset ...
        .CreateNodeByCurrentRecordset "parent", "NWIND", 0, 1, 3, 4, , , False

        ' Check Count of nodes to expand
        If .Nodes.Count Then
            .Nodes(1).Expanded = True
        End If

        ' Unlock , and draw all nodes
        .LockUpdate = False

    End With

End Sub

' Purpose : Insert Node item by recordset >> From table "Nodes" <<
Public Sub CreateNodeDemo()
    On Error Resume Next

    With tvwTest
        ' Clean Up
        .Clear

        .LockUpdate = True

        ' Prepared recordset ...
        Set .TableNodes = GetRST("select * from nodes")
        
        ' Create node by recordset ...
        .CreateNodeByCurrentRecordset "parent", "NWIND", 0, 1, 3, 4, , , False

        ' Check Count of nodes to expand
        If .Nodes.Count Then
            .Nodes(1).Expanded = True
        End If

        ' Unlock , and draw all nodes
        .LockUpdate = False

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    frmMain.RefreshView
    frmMain.Show
End Sub

' Purpose: Show or Hide lvwPrivileges
Private Sub LblDetail_Click()
    On Error Resume Next

    If tvwMain.Height <> 170 Then
        lvwPrivileges.Visible = True
        
        tvwMain.Height = 170
        MyTab.Height = 170
        LblDetail.Caption = "&Detail Privileges <<"
    Else
        lvwPrivileges.Visible = 0
        tvwMain.Height = 359
        MyTab.Height = 359
        LblDetail.Caption = "&Detail Privileges >>"
    End If

End Sub

Private Sub LblGetNodes_Click()
    On Error Resume Next

    lvwPrivileges.InsertItemFromNodes tvwMain.Nodes
    StrValidation = lvwPrivileges.GetPrivileges
    txtData(3).Text = StrValidation
    
    CreateNodeDemo

End Sub

Private Sub lvwPrivileges_HeaderRightClick(Index As Integer, _
                                           X As Single, _
                                           Y As Single)
    On Error Resume Next

    If Index < 2 Then Exit Sub

    Flags = 1
    PopupMenu mnuCheckAll, , X, Y

    DoEvents
    
    If Flags <> 1 Then
        lvwPrivileges.CheckAll Index, Abs(Flags)
    End If

End Sub

Private Sub Mnu_Child_All_Click(Index As Integer)
    Flags = Index - 2
End Sub


Private Sub tvwMain_AllowAddNode(NewNodeKey As String, _
                                 Allow As Boolean, _
                                 DATA As String, _
                                 iStart As Long)

    Allow = True
    DATA = "00000"

End Sub

Private Sub tvwtest_AllowAddNode(NewNodeKey As String, _
                                 Allow As Boolean, _
                                 DATA As String, _
                                 iStart As Long)

    Dim X As Long
    X = InStr(1, StrValidation, VALID_NODE_KEY(NewNodeKey))

    DATA = Mid(StrValidation, X + Len(NewNodeKey) + 5, 5)

    Allow = X

End Sub


' Purpose: Show data from recordset
Private Sub FillData()
    On Error Resume Next

    If MyADODC1.Rs.RecordCount Then
    
        StrValidation = txtData(3).Text
        
        CreateNodeDemo
        
        lvwPrivileges.InsertItemFromNodes tvwTest.Nodes, False

    End If

End Sub





