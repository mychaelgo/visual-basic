VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm6_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Sample Drag Selected Demo"
   ClientHeight    =   5040
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm6_basic.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   3405
      Left            =   240
      TabIndex        =   1
      Top             =   630
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   6006
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
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      XPAlphaBlend    =   0   'False
      MaxAllColumnWidth=   600
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   "MyImageList1"
      HeaderCaption   =   "OsenXPListBox1"
      DataList        =   "Sample1|Sample2|Sample3|Sample4|Sample5|Sample6|Sample7|Sample8|Sample9"
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      OLEDropMode     =   1
      AutoDragAndDrop =   -1  'True
      DragSelected    =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4185
      _ExtentX        =   7382
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
      Caption         =   "Sample Drag Selected Demo"
      TitleTop        =   7
      icon            =   "frm6_basic.frx":038A
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      CtlAutoChangeScheme=   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Moves selected rows in Mouse Enter  and finish the dragging in Mouse Up."
      Height          =   765
      Left            =   390
      TabIndex        =   2
      Top             =   4200
      Width           =   3255
   End
End
Attribute VB_Name = "frm6_basic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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








