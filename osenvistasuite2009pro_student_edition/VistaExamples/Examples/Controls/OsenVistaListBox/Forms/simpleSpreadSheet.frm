VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form frm_SS 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Simple Spreadsheet Demo"
   ClientHeight    =   6900
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "simpleSpreadSheet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaComboBox OsenXPComboBox1 
      Height          =   315
      Left            =   7590
      TabIndex        =   3
      Top             =   6390
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
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
      Text            =   "Silver"
      ComboStyle      =   1
      LBN             =   16777215
      LBS             =   7381139
      LBG1            =   16777215
      LBG2            =   8632490
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LLID            =   2
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
      DataList        =   "Blue|Olive Green|Silver"
      BorderColor     =   12164479
      BorderColorOver =   12164479
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   5745
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10134
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
      BackSelected    =   7381139
      BackSelectedG1  =   16777215
      BackSelectedG2  =   8632490
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   5
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
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientTop=   8569007
      HeaderGradientBottom=   4487779
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "simpleSpreadSheet.frx":038A
      Begin VistaSuitePro.OsenVistaTextBox OsenXPTextBox1 
         Height          =   315
         Left            =   3540
         TabIndex        =   2
         Top             =   2730
         Visible         =   0   'False
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Text            =   "TextBox1"
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
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
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
      Caption         =   "Simple Spreadsheet Demo"
      TitleTop        =   7
      icon            =   "simpleSpreadSheet.frx":03A2
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Scheme:"
      Height          =   195
      Left            =   6330
      TabIndex        =   4
      Top             =   6450
      Width           =   1035
   End
End
Attribute VB_Name = "frm_SS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenXPSuite 2006 - OsenXPListBox sample                              *'
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

  ' Set The Default Color scheme for All forms in this projects
  ' and set osenxpform1.usedefaulttheme=true
    OsenXPComboBox1.ListIndex = 1

    ' Set Simple SpreadSheet on OsenXPListbox
    ' lngRows = Number of row
    ' lngCols = Number of column
    ' TxtData = OsenXPTextBox (Editable cell)
    Me.OsenXPListBox1.SetupSimpleSpreadSheet 100, 26, Me.OsenXPTextBox1
    
    
End Sub

