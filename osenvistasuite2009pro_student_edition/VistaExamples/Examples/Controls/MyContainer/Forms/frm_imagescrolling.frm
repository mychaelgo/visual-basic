VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_image 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Scrolling picture with MyContainerCTL"
   ClientHeight    =   6915
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_imagescrolling.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyContainerCtl MyContainerCtl1 
      Height          =   6225
      Left            =   210
      TabIndex        =   1
      Top             =   570
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10980
      BackColor       =   8421504
      ScaleWidth      =   449
      ScaleHeight     =   415
      ClientWidth     =   12315
      ClientHeight    =   9315
      Picture         =   "frm_imagescrolling.frx":038A
      ScrollingPicture=   -1  'True
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
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
      Caption         =   "Scrolling picture with MyContainerCTL"
      TitleTop        =   7
      icon            =   "frm_imagescrolling.frx":7824
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      CaptionAlignment=   1
      AllowFadeIn     =   -1  'True
   End
End
Attribute VB_Name = "frm_image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************'
'*  OsenVistaSuite 2008 -  MyContainerCTL sample                         *'
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
    
End Sub




























