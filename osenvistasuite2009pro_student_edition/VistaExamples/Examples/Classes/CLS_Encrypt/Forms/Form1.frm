VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "CLS_Encrypt Sample"
   ClientHeight    =   4365
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   7800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton cmdTest 
      Height          =   345
      Left            =   5970
      TabIndex        =   8
      Top             =   3780
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Caption         =   "&Test"
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
      MICON           =   "Form1.frx":038A
      PICN            =   "Form1.frx":03A6
      UMCOL           =   -1  'True
      BinaryImageNormal=   "Form1.frx":0740
      BinaryImageOver =   "Form1.frx":0758
   End
   Begin VistaSuitePro.OsenVistaTextBox txtKey 
      Height          =   330
      Left            =   1830
      TabIndex        =   7
      Top             =   3750
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   582
      Text            =   "masterkey"
      Alignment       =   2
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
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
      LabelBackColor  =   15790320
      LabelCaption    =   "Encrypt Key:"
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
   Begin VistaSuitePro.OsenVistaCheckBox chkHex 
      Height          =   315
      Left            =   270
      TabIndex        =   6
      Top             =   3750
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BackColor       =   15790320
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
      Caption         =   "Hex Format"
   End
   Begin VistaSuitePro.OsenVistaTab TabX 
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   1380
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4048
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
      NumberOfTabs    =   4
      BackColorParent =   14215660
      Style           =   0
      TabsPerRow      =   4
      TabWidth1       =   112
      TabText1        =   "Plain Text"
      TabEnabled1     =   -1  'True
      TabVisible1     =   0   'False
      TabPicture1     =   "Form1.frx":0770
      TabCountCtls1   =   1
      TabNo1CtlID1    =   "txtData(0)"
      TabNo1CtlIX1    =   1
      TabNo1CtlIT1    =   -1  'True
      TabWidth2       =   146
      TabText2        =   "Encrypted Text"
      TabEnabled2     =   -1  'True
      TabVisible2     =   0   'False
      TabPicture2     =   "Form1.frx":0AC2
      TabCountCtls2   =   1
      TabNo2CtlID1    =   "txtData(1)"
      TabNo2CtlIX1    =   2
      TabNo2CtlIT1    =   -1  'True
      TabWidth3       =   155
      TabText3        =   "Descrypted Text"
      TabEnabled3     =   -1  'True
      TabVisible3     =   0   'False
      TabPicture3     =   "Form1.frx":0E14
      TabCountCtls3   =   1
      TabNo3CtlID1    =   "txtData(2)"
      TabNo3CtlIX1    =   3
      TabNo3CtlIT1    =   -1  'True
      TabWidth4       =   72
      TabText4        =   "File"
      TabEnabled4     =   -1  'True
      TabVisible4     =   0   'False
      TabPicture4     =   "Form1.frx":1166
      TabCountCtls4   =   9
      TabNo4CtlID1    =   "cmdEncFile"
      TabNo4CtlIX1    =   4
      TabNo4CtlIT1    =   -1  'True
      TabNo4CtlID2    =   "txtSource"
      TabNo4CtlIX2    =   4
      TabNo4CtlIT2    =   -1  'True
      TabNo4CtlID3    =   "OsenXPLabel2"
      TabNo4CtlIX3    =   4
      TabNo4CtlIT3    =   -1  'True
      TabNo4CtlID4    =   "OsenXPLabel3"
      TabNo4CtlIX4    =   4
      TabNo4CtlIT4    =   -1  'True
      TabNo4CtlID5    =   "txtEnc"
      TabNo4CtlIX5    =   4
      TabNo4CtlIT5    =   -1  'True
      TabNo4CtlID6    =   "txtFileKey"
      TabNo4CtlIX6    =   4
      TabNo4CtlIT6    =   -1  'True
      TabNo4CtlID7    =   "OsenXPLabel4"
      TabNo4CtlIX7    =   4
      TabNo4CtlIT7    =   -1  'True
      TabNo4CtlID8    =   "OsenXPLabel1"
      TabNo4CtlIX8    =   4
      TabNo4CtlIT8    =   -1  'True
      TabNo4CtlID9    =   "txtDes"
      TabNo4CtlIX9    =   4
      TabNo4CtlIT9    =   -1  'True
      Begin VistaSuitePro.OsenVistaButton cmdEncFile 
         Height          =   345
         Left            =   80910
         TabIndex        =   15
         Top             =   1800
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         Caption         =   "&Encrypt Now!"
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
         MICON           =   "Form1.frx":14B8
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         BinaryImageNormal=   "Form1.frx":14D4
         BinaryImageOver =   "Form1.frx":14EC
      End
      Begin VistaSuitePro.OsenVistaTextBox txtSource 
         Height          =   345
         Left            =   76380
         TabIndex        =   10
         Top             =   480
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   609
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
         Locked          =   -1  'True
         ButtonCaption   =   ""
         ButtonPicture   =   "Form1.frx":1504
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
         BackColorOver   =   12648447
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel2 
         Height          =   225
         Left            =   75120
         TabIndex        =   9
         Top             =   510
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Source file:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   1875
         Index           =   0
         Left            =   30
         TabIndex        =   3
         ToolTipText     =   "Writes down any word which would in encrypt"
         Top             =   360
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3307
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
         BackColorOver   =   12648447
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
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   1875
         Index           =   1
         Left            =   75030
         TabIndex        =   4
         Top             =   360
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3307
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
         Locked          =   -1  'True
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
         BackColorOver   =   14737632
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
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   1875
         Index           =   2
         Left            =   75030
         TabIndex        =   5
         Top             =   360
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3307
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
         Locked          =   -1  'True
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
         BackColorOver   =   14737632
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel3 
         Height          =   285
         Left            =   75120
         TabIndex        =   11
         Top             =   960
         Width           =   1140
         _ExtentX        =   2011
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
         Caption         =   "Encrypted file:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox txtEnc 
         Height          =   345
         Left            =   76380
         TabIndex        =   12
         Top             =   930
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   609
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
         ForeColor       =   16711680
         Locked          =   -1  'True
         ButtonCaption   =   ""
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   14737632
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
      Begin VistaSuitePro.OsenVistaTextBox txtFileKey 
         Height          =   315
         Left            =   76380
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Text            =   "masterkey"
         Alignment       =   2
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel4 
         Height          =   285
         Left            =   75120
         TabIndex        =   14
         Top             =   1830
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
         Caption         =   "Encrypt Key:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Left            =   75120
         TabIndex        =   16
         Top             =   1380
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "Descrypted file:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox txtDes 
         Height          =   345
         Left            =   76380
         TabIndex        =   17
         Top             =   1350
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   609
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
         ForeColor       =   4210752
         Locked          =   -1  'True
         ButtonCaption   =   ""
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   14737632
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
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   7800
      _ExtentX        =   13758
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
      Picture         =   "Form1.frx":189E
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The Following is example of usage CLS_Encrypt"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Class Name: CLS_Encrypt"
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
      BinaryImage     =   "Form1.frx":24F0
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
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
      Caption         =   "CLS_Encrypt Sample"
      TitleTop        =   7
      icon            =   "Form1.frx":2508
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************'
'*  OsenVistaSuite 2008 - CLS_Encrypt sample                             *'
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
Private c_Enc        As CLS_Encrypt

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize event)
    Me.OsenXPForm1.Init Me
    
    ' Now, you can make a new instance of CLS_Encrypt
    Set c_Enc = New CLS_Encrypt
    
End Sub

Private Sub cmdTest_Click()
    
    ' Check the Plain Text situation
    ' Ignore when the plain text is empty
    
    If Len(txtData(0).Text) > 0 Then
    
        Dim Lx1             As Long
        Dim Lx2             As Long
        Dim str_result      As String
        
        ' Call API Gettickcount
        Lx1 = GTick() ' Gtick() function can be find in XP class members
        
        ' Now it's the moment to encryption process
        str_result = c_Enc.EncryptString(txtData(0).Text, txtKey.Text, chkHex.Value)
        
        ' Calculated the duration encryption process time
        Lx1 = GTick() - Lx1 ' ticktime in ms [mili second]
        
        ' saves result of encryption at txtdata(1)
        txtData(1).Text = str_result
        
        ' Call API Gettickcount
        Lx2 = GTick()
        
        ' Now it's the moment to descryption process
        str_result = c_Enc.DecryptString(txtData(1).Text, txtKey.Text, chkHex.Value)
        
        ' Calculated the duration descryption process time
        Lx2 = GTick() - Lx2 ' ticktime in ms [mili second]
        
        ' saves result of descryption at txtdata(2)
        txtData(2).Text = str_result
    
        ' Display report ...
        MsgBoxGT "The number of byte which in processing: " & Len(txtData(0).Text) & " Byte" & vbCrLf & _
                 Lx1 & " ms taken at encryption process" & vbCrLf & _
                 Lx2 & " ms taken at descryption process", vbInformation
    
    End If
    
    
End Sub

Private Sub cmdEncFile_Click()

    Dim Lx1 As Long
    Dim Lx2 As Long
    
    ' As same as gettickcount
    Lx1 = GTick()
    
    ' Encrypt ...
    c_Enc.EncryptFile txtSource.Text, txtEnc.Text, txtFileKey.Text, False

    'Calculated the duration encryption process time
    Lx1 = GTick() - Lx1
    
    Lx2 = GTick()
    
    ' Descrypt
    c_Enc.DecryptFile txtEnc.Text, txtDes.Text, txtFileKey.Text, False
    
    'Calculated the duration descryption process time
    Lx2 = GTick() - Lx2
    
    ' Display report ...
    MsgBoxGT "The file size which in processing: " & Format$(FileLen(txtSource.Text) / 1024, "#,###,##0") & " Kb" & vbCrLf & _
             Lx1 & " ms taken at encryption process" & vbCrLf & _
             Lx2 & " ms taken at descryption process", vbInformation
             

End Sub

Private Sub txtSource_ButtonClick()

    ' Show the open dialog and get the filename
    txtSource.ShowOpenDialog
    
    ' check the result, ignore if empty
    If Len(txtSource.Text) Then
        
        ' set the filename for result of encryption
        txtEnc.Text = txtSource.Text & ".enc"
        
        ' set the filename for result of descryption
        txtDes.Text = txtSource.Text & ".des"
    
    End If
    
End Sub


Private Sub TabX_TabSelected(ByVal iTabIndex As Integer)

    chkHex.Visible = iTabIndex Mod 4
    txtKey.Visible = iTabIndex Mod 4
    cmdTest.Visible = iTabIndex Mod 4
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' CLean Up
    Set c_Enc = Nothing
    
End Sub




