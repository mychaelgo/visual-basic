VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm7_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Customize an icon for each cells"
   ClientHeight    =   5850
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm7_basic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   375
      Left            =   4410
      TabIndex        =   3
      Top             =   5280
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      Caption         =   "Test 2"
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
      MICON           =   "frm7_basic.frx":058A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   375
      Left            =   2670
      TabIndex        =   2
      Top             =   5280
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      Caption         =   "Test 1"
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
      MICON           =   "frm7_basic.frx":05A6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   5310
      Top             =   4560
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   44772
      Images          =   "frm7_basic.frx":05C2
      Version         =   720920
      KeyCount        =   39
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   4635
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8176
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
      ShowHeader      =   -1  'True
      HeaderFormatString=   "No.;40;2;0;;0|test2;100;0;0;;-1|test3;100;0;0;;-1|test4;100;0;0;;-1|test5;100;0;0;;-1|test6;100;0;1;;-1"
      Columns         =   6
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      MaxAllColumnWidth=   540
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
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
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
      Caption         =   "Customize an icon for each cells"
      TitleTop        =   7
      icon            =   "frm7_basic.frx":B4C6
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
   End
End
Attribute VB_Name = "frm7_basic"
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

Dim sData As String
Dim sIcon As String

Private Sub Form_Load()

    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me


End Sub

Private Sub OsenXPButton1_Click()

    ' Add new item
    Me.OsenXPListBox1.AddItem OsenXPListBox1.ListCount + 1 & vbTab & "sample2" & vbTab & "sample3" & vbTab & MD5(Timer)
    
    ' Now, specify or formatting each cell with randomize icon
    OsenXPListBox1.SetReportIcon OsenXPListBox1.ListIndex, 1, CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1)
    OsenXPListBox1.SetReportIcon OsenXPListBox1.ListIndex, 2, CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1)
    OsenXPListBox1.SetReportIcon OsenXPListBox1.ListIndex, 3, CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1), enAlignRight
    OsenXPListBox1.SetReportIcon OsenXPListBox1.ListIndex, 4, CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1), enAlignCenter
    
End Sub


Private Sub OsenXPButton2_Click()

    sData = OsenXPListBox1.ListCount + 1 & vbTab & MD5(Timer) & vbTab & gTick & vbTab & "Test4"
    sIcon = "-1," & CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1) & "," & CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1) & "," & CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1) & ";2," & CLng(Rnd(1) * Me.MyImageList1.ImageCount - 1) & ";1"
    
    ' Now, Add new item with icon for each cells that was specified
    OsenXPListBox1.AddItem sData, iconcollections:=sIcon

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








