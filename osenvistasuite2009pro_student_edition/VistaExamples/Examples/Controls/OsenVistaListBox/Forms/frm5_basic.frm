VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm5_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Sample Drag and Drop"
   ClientHeight    =   7470
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm5_basic.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   345
      Left            =   4170
      TabIndex        =   8
      Top             =   6750
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      Caption         =   "&Clear"
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
      MICON           =   "frm5_basic.frx":038A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.OsenVistaOptionButton chkSelectedMode 
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   7
      Top             =   3390
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   450
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      Alignment       =   1
      Caption         =   "Multiple"
   End
   Begin VistaSuitePro.OsenVistaOptionButton chkSelectedMode 
      Height          =   255
      Index           =   0
      Left            =   7800
      TabIndex        =   6
      Top             =   3390
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   450
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      Value           =   -1  'True
      Caption         =   "Single"
   End
   Begin VistaSuitePro.OsenVistaListBox lstTarget 
      Height          =   2325
      Left            =   330
      TabIndex        =   5
      Top             =   4320
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4101
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
      HeaderCaption   =   "OsenXPListBox2"
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
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3450
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Add Items"
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
      MICON           =   "frm5_basic.frx":03A6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   4680
      Top             =   3480
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   44772
      Images          =   "frm5_basic.frx":03C2
      Version         =   720920
      KeyCount        =   39
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   2475
      Left            =   330
      TabIndex        =   1
      Top             =   870
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   4366
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
      BackSelectedG2  =   33023
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
      OLEDragMode     =   1
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
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
      Caption         =   "Sample Drag and Drop"
      TitleTop        =   7
      icon            =   "frm5_basic.frx":B2C6
      BorderStyle     =   1
      CtlAutoChangeScheme=   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target:"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   4020
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   630
      Width           =   555
   End
End
Attribute VB_Name = "frm5_basic"
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

End Sub

'********************* DRAG AND DROP **********************
' OsenXPListBox1.AutoDragAndDrop=True
' OsenXPListBox1.OLEDragMode=1
' OsenXPListBox1.OLEDropMode=1
'
' OsenXPListBox2.AutoDragAndDrop=True
' OsenXPListBox2.OLEDragMode=0
' OsenXPListBox2.OLEDropMode=1
'**********************************************************

' Populate list
Private Sub OsenXPButton1_Click()
    
    With Me.OsenXPListBox1
        .AddItem MD5(Timer), , , , 1, , vbRed, True, , , vbYellow
        .AddItem MD5(Timer), , , , 2, , vbBlue, True, True
        .AddItem MD5(Timer), , , , 4, , , , , True, vbYellow
        .AddItem MD5(Timer), , , , 3, , vbGreen, True, , , vbBlack
        .AddItem MD5(Timer), , , , 7, , vbBlue, True, , , vbYellow
        .AddItem MD5(Timer), , , , 12, , vbRed, , , True, vbBlack
        .AddItem MD5(Timer), , , , 14
        .AddItem MD5(Timer), , , , 17, , vbBlue
    End With
    
End Sub

Private Sub OsenXPButton2_Click()
    lstTarget.Clear
End Sub

Private Sub chkSelectedMode_Click(Index As Integer)
    OsenXPListBox1.SelectMode = Index
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








