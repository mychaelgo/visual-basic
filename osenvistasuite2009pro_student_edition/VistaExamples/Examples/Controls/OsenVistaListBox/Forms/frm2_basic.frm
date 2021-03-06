VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm2_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "ListBox With Image"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   Icon            =   "frm2_basic.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   228
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   720
      Top             =   3210
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   19516
      Images          =   "frm2_basic.frx":038A
      Version         =   720920
      KeyCount        =   17
      Keys            =   "����������������"
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   4470
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   767
      Caption         =   "Populate 100,000 Items"
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
      MICON           =   "frm2_basic.frx":4FE6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   15779735
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   3855
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   6800
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
      ItemHeight      =   21
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      ItemTextLeft    =   20
      SelectModeStyle =   2
      Lstyle          =   2
      XPAlphaBlend    =   0   'False
      AlternateRowColors=   -1  'True
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
      Picture         =   "frm2_basic.frx":5002
      PicturePosition =   2
      TransparencyLevel=   22
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
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3420
      _ExtentX        =   6033
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
      Caption         =   "ListBox With Image"
      TitleTop        =   7
      icon            =   "frm2_basic.frx":5BCE
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
   End
End
Attribute VB_Name = "frm2_basic"
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
    Dim l As Long
    Dim z   As Long
    
    ' Clear all items at osenxplistbox
    OsenXPListBox1.Clear
    
    ' set the imagelist handle
    OsenXPListBox1.SmallIcons = Me.MyImageList1.hIml
    
    ' lock the listbox for increase speed performance
    OsenXPListBox1.LockUpdate = True
    
    ' GettickCount()
    z = gTick
    
    For l = 1 To 10000
        Me.OsenXPListBox1.AddItem "Osen Kusnadi: " & l, , , , CInt(Rnd() * 17), CInt(Rnd() * 17)
    Next
    
    z = gTick - z
    
    OsenXPListBox1.LockUpdate = False
    DoEvents
    
    MsgBoxGT "10,000 Items was successful inserted" & vbCrLf & z & " ms taken", vbInformation
    
    
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








