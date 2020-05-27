VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm_basic1 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Basic Editable Cell #1 (Simple)"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   Icon            =   "frm_basic1.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   3  'Windows Default
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   5715
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   10081
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
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
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
      HeaderCaption   =   "OsenXPListBox1"
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      Begin VistaSuitePro.OsenVistaTextBox txtCell 
         Height          =   345
         Left            =   7470
         TabIndex        =   2
         Top             =   5190
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
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
      Caption         =   "Basic Editable Cell #1 (Simple)"
      TitleTop        =   7
      icon            =   "frm_basic1.frx":038A
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
   End
End
Attribute VB_Name = "frm_basic1"
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
    
    ' Open Recordset
    Me.OsenXPListBox1.InsertItemBySQL AdoCN, "Select * from customers", , , True
    
    ' Bind OsenXPTextBox with OsenXPListBox
    Me.OsenXPListBox1.BindObject txtCell, 1, 11 '1= First Column Index, Count of columns/fields
    
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








