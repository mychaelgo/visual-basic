VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Advanced Demo"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   753
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyContainerCtl MyContainerCtl1 
      Height          =   8805
      Left            =   90
      TabIndex        =   1
      Top             =   480
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   15531
      BackColor       =   8421504
      ScaleWidth      =   741
      ScaleHeight     =   587
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   25000
         Left            =   0
         ScaleHeight     =   25005
         ScaleWidth      =   25005
         TabIndex        =   2
         Top             =   0
         Width           =   25000
      End
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Advanced Demo"
      TitleTop        =   7
      icon            =   "Form1.frx":058A
      BorderStyle     =   1
      UseDefaultTheme =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenXPSuite 2006 - MyContainerCTL sample                             *'
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
 
 
 
 
 
 
 
 
 
 
 
 
 
 








