VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_1 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "OsenXPHookMenu Sample"
   ClientHeight    =   4275
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   405
      Left            =   1470
      TabIndex        =   3
      Top             =   2910
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   714
      Caption         =   "Demo2"
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
      MICON           =   "frm_1.frx":058A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_1.frx":05A6
      BinaryImageOver =   "frm_1.frx":05BE
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   435
      Left            =   1470
      TabIndex        =   2
      Top             =   1980
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   767
      Caption         =   "Demo1"
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
      MICON           =   "frm_1.frx":05D6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_1.frx":05F2
      BinaryImageOver =   "frm_1.frx":060A
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   5355
      _ExtentX        =   9446
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
      Picture         =   "frm_1.frx":0622
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   14854529
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The following is example of OsenXPHookMenu Usage"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Control Name: OsenXPHookMenu"
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
      BinaryImage     =   "frm_1.frx":1274
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      _ExtentX        =   9446
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
      Caption         =   "OsenXPHookMenu Sample"
      TitleTop        =   7
      icon            =   "frm_1.frx":128C
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
   End
End
Attribute VB_Name = "frm_1"
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

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
       
End Sub

Private Sub OsenXPButton1_Click()
    frm_demo1.Show
End Sub

Private Sub OsenXPButton2_Click()
    Frm_Main.Show
End Sub






















