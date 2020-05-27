VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_Main 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "MyImageList Sample"
   ClientHeight    =   3360
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   4905
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
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   327
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   2190
      Top             =   2490
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   19516
      Images          =   "frm_main.frx":038A
      Version         =   65536
      KeyCount        =   17
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaComboBox OsenXPComboBox1 
      Height          =   375
      Left            =   1740
      TabIndex        =   2
      Top             =   1860
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
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
      ComboStyle      =   1
      DisplayPicture  =   -1  'True
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LIO             =   2
      LITL            =   22
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   16777215
      ASURC           =   0   'False
      TextColumn      =   0
      Required        =   0   'False
      Unicode         =   0   'False
      BorderColor     =   12164479
      BorderColorOver =   12164479
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   450
      Width           =   4905
      _ExtentX        =   8652
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
      Picture         =   "frm_main.frx":4FE6
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The following is example of MyImageList Usage"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Control Name: MyImageList "
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
      BinaryImage     =   "frm_main.frx":5C38
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "MyImageList Sample"
      TitleTop        =   7
      icon            =   "frm_main.frx":5C50
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenVistaSuite 2008 -  MyImageList sample                                *'
'*  Copyright (c) 2008 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  This file is part of the OsenVistaSuite 2008 sample applications.    *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenVistaSuite 2008 documentation, and is provided "as is", without  *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
    ' Display all image/picture from MyImageList object at the OsenXPComboBox
    Me.OsenXPComboBox1.DisplayIconFormImageList "Myimagelist1"
    
   
End Sub

Private Sub OsenXPComboBox1_Click()
    
    ' Test
    MsgBoxXP "Display selected icon at the message box", 4096, "Test", , , MyImageList1.ItemPicture(Me.OsenXPComboBox1.ListIndex + 1), True
    
    
End Sub






















