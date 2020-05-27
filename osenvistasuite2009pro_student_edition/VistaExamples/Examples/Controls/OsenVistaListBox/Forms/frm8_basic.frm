VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm8_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Large Icon"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   Icon            =   "frm8_basic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   4260
      Top             =   4530
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      Iconsize        =   2
      Size            =   75004
      Images          =   "frm8_basic.frx":038A
      Version         =   720920
      KeyCount        =   17
      Keys            =   "����������������"
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   4395
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   7752
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
      ViewMode        =   1
      HeaderCaption   =   "Northwind Traders 2006"
      HeaderAlignment =   1
      Picture         =   "frm8_basic.frx":128A6
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
      Width           =   7500
      _ExtentX        =   13229
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
      Caption         =   "Large Icon"
      TitleTop        =   7
      icon            =   "frm8_basic.frx":428F8
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
   End
End
Attribute VB_Name = "frm8_basic"
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

    
    ' set the imagelist handle
    OsenXPListBox1.LargeIcons = Me.MyImageList1.hIml
    
    Dim L As Long
    
    For L = 1 To 15
        Me.OsenXPListBox1.AddItem "Osen " & 2000 + L, , , , CInt(Rnd() * 16) + 1, CInt(Rnd * 16) + 1
    Next

End Sub


Private Sub OsenXPListBox1_IconClick(lRowIndex As Long)
    Debug.Print "IconIndex: "; ; lRowIndex; ; " :: Caption: "; ; OsenXPListBox1.List(lRowIndex - 1)
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








