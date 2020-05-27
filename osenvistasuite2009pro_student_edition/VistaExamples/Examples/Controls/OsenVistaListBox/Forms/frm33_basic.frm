VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm3_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm33_basic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   666
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   405
      Left            =   5370
      TabIndex        =   3
      Top             =   6180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      Caption         =   "Clear"
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
      MICON           =   "frm33_basic.frx":000C
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   405
      Left            =   3240
      TabIndex        =   2
      Top             =   6180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
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
      MICON           =   "frm33_basic.frx":0028
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   5385
      Left            =   210
      TabIndex        =   1
      Top             =   630
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   9499
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
      HeaderFormatString=   "No.;100;0;0;;-1|Message 1;300;0;0;;-1|Message 2;200;0;0;;-1"
      Columns         =   3
      ShowGridLines   =   -1  'True
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
      IMGLIST         =   ""
      AutoSetRowHeight=   -1  'True
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
      BinaryImage     =   "frm33_basic.frx":0044
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
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
      Caption         =   "Form1"
      TitleTop        =   7
      BorderStyle     =   1
      AllowFadeIn     =   -1  'True
   End
End
Attribute VB_Name = "frm3_basic"
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

Private Sub OsenXPButton1_Click()
    OsenXPListBox1.AddItem OsenXPListBox1.ListCount & vbTab & MD5(Timer) & vbTab & MD5(Timer) & vbCrLf & CRC32(Timer) & vbCrLf & MD5(Timer)
    OsenXPListBox1.AddItem OsenXPListBox1.ListCount & vbTab & MD5(Timer)   ' & vbTab & MD5(Timer) & vbCrLf & CRC32FromString(Timer) & vbCrLf & MD5(Timer)
    OsenXPListBox1.AddItem OsenXPListBox1.ListCount & vbTab & MD5(Timer) & vbTab & MD5(Timer) & vbCrLf & CRC32(Timer) & vbCrLf & MD5(Timer) & vbCrLf & MD5(Timer)
    OsenXPListBox1.AddItem OsenXPListBox1.ListCount & vbTab & MD5(Timer)   '& vbTab & MD5(Timer) & vbCrLf & CRC32FromString(Timer) & vbCrLf & MD5(Timer)
    OsenXPListBox1.AddItem OsenXPListBox1.ListCount & vbTab & MD5(Timer) & vbTab & MD5(Timer) & vbCrLf & CRC32(Timer) & vbCrLf & MD5(Timer) & vbCrLf & MD5(Timer) & vbCrLf & MD5(Timer)
    OsenXPListBox1.LockUpdate = False

End Sub

Private Sub OsenXPButton2_Click()
     OsenXPListBox1.Clear
End Sub































