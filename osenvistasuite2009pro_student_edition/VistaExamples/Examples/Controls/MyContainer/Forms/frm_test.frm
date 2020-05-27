VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_test 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Sample #3"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   Icon            =   "frm_test.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyContainerCtl MyContainerCtl1 
      Height          =   6885
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12144
      BackColor       =   8421504
      ScaleWidth      =   793
      ScaleHeight     =   459
      ClientWidth     =   13140
      ClientHeight    =   10440
      Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
         Height          =   10245
         Left            =   -30
         TabIndex        =   2
         Top             =   -30
         Width           =   12945
         _ExtentX        =   22834
         _ExtentY        =   18071
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
         BorderColor     =   12563634
         GradientBackGround=   -1  'True
         GradientColor2  =   12563634
         GradientOrientation=   1
         BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DescriptionLeft =   42
         BorderStyle     =   1
         BinaryImage     =   "frm_test.frx":058A
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame6 
            Height          =   3135
            Left            =   6300
            TabIndex        =   11
            Top             =   6630
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   5530
            Caption         =   "Statistic #3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14215660
            ForeColor       =   16777215
            BorderColor     =   5800032
            Appearance      =   1
            DropDownButton  =   -1  'True
            Picture         =   "frm_test.frx":05A2
            PicturePosition =   2
            BinaryImage     =   "frm_test.frx":3294
            Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   17
               Top             =   2700
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "frm_test.frx":32AC
               MousePointer    =   99
               ForeColorDown   =   16711935
               Caption         =   "http://osenxpsuite.net"
               ForeColor       =   16748098
               HiperLink       =   "http://osenxpsuite.net"
               AutoSize        =   0   'False
               UnderlineOnOver =   -1  'True
               BackStyle       =   0
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   $"frm_test.frx":340E
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Index           =   2
               Left            =   150
               TabIndex        =   16
               Top             =   510
               Width           =   4875
            End
         End
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame5 
            Height          =   3255
            Left            =   6300
            TabIndex        =   10
            Top             =   3300
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   5741
            Caption         =   "Statistic #2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14215660
            ForeColor       =   16777215
            BorderColor     =   12017457
            Appearance      =   1
            DropDownButton  =   -1  'True
            Picture         =   "frm_test.frx":34FB
            PicturePosition =   2
            BinaryImage     =   "frm_test.frx":59BB
            Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   15
               Top             =   2880
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "frm_test.frx":59D3
               MousePointer    =   99
               ForeColorDown   =   16711935
               Caption         =   "http://osenxpsuite.net"
               ForeColor       =   16748098
               HiperLink       =   "http://osenxpsuite.net"
               AutoSize        =   0   'False
               UnderlineOnOver =   -1  'True
               BackStyle       =   0
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   $"frm_test.frx":5B35
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Index           =   1
               Left            =   120
               TabIndex        =   14
               Top             =   390
               Width           =   4875
            End
         End
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame4 
            Height          =   3165
            Left            =   6300
            TabIndex        =   9
            Top             =   90
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   5583
            Caption         =   "Statistic #1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14215660
            ForeColor       =   16777215
            BorderColor     =   9731196
            Appearance      =   1
            DropDownButton  =   -1  'True
            Picture         =   "frm_test.frx":5C22
            PicturePosition =   2
            BinaryImage     =   "frm_test.frx":7E97
            Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   13
               Top             =   2700
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "frm_test.frx":7EAF
               MousePointer    =   99
               ForeColorDown   =   16711935
               Caption         =   "http://osenxpsuite.net"
               ForeColor       =   16748098
               HiperLink       =   "http://osenxpsuite.net"
               AutoSize        =   0   'False
               UnderlineOnOver =   -1  'True
               BackStyle       =   0
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   $"frm_test.frx":8011
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   510
               Width           =   4875
            End
         End
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame3 
            Height          =   3165
            Left            =   180
            TabIndex        =   5
            Top             =   6630
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   5583
            Caption         =   "Products"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14215660
            ForeColor       =   16777215
            BorderColor     =   9731196
            Appearance      =   1
            DropDownButton  =   -1  'True
            image           =   "frm_test.frx":80FE
            BinaryImage     =   "frm_test.frx":8498
            Begin VistaSuitePro.OsenVistaListBox OsenXPListBox3 
               Height          =   2775
               Left            =   30
               TabIndex        =   8
               Top             =   360
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   4895
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
               BackSelected    =   12563634
               BackSelectedG1  =   16777215
               BackSelectedG2  =   14140358
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
               HeaderCaption   =   "OsenXPListBox3"
               TransparencyLevel=   22
               ReadOnDemand    =   -1  'True
               BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               HeaderGradientTop=   12560296
               HeaderGradientBottom=   9531248
               HeaderGradientAllow=   -1  'True
               HeaderForeColor =   16777215
               BinaryImage     =   "frm_test.frx":84B0
            End
         End
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame2 
            Height          =   3225
            Left            =   180
            TabIndex        =   4
            Top             =   3330
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   5689
            Caption         =   "Suppliers"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14215660
            ForeColor       =   16777215
            BorderColor     =   5800032
            Appearance      =   1
            DropDownButton  =   -1  'True
            image           =   "frm_test.frx":84C8
            BinaryImage     =   "frm_test.frx":8A62
            Begin VistaSuitePro.OsenVistaListBox OsenXPListBox2 
               Height          =   2835
               Left            =   30
               TabIndex        =   7
               Top             =   360
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   5001
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
               BackSelected    =   7381139
               BackSelectedG1  =   16777215
               BackSelectedG2  =   8632490
               WordWrap        =   0   'False
               ItemHeightAuto  =   0   'False
               ItemOffset      =   2
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
               HeaderCaption   =   "OsenXPListBox2"
               TransparencyLevel=   22
               ReadOnDemand    =   -1  'True
               BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               HeaderGradientTop=   8569007
               HeaderGradientBottom=   4487779
               BinaryImage     =   "frm_test.frx":8A7A
            End
         End
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
            Height          =   3195
            Left            =   180
            TabIndex        =   3
            Top             =   90
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   5636
            Caption         =   "Customers"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14215660
            ForeColor       =   16777215
            BorderColor     =   12017457
            Appearance      =   1
            DropDownButton  =   -1  'True
            image           =   "frm_test.frx":8A92
            BinaryImage     =   "frm_test.frx":902C
            Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
               Height          =   2805
               Left            =   30
               TabIndex        =   6
               Top             =   360
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   4948
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
               TransparencyLevel=   22
               ReadOnDemand    =   -1  'True
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
               BinaryImage     =   "frm_test.frx":9044
            End
         End
      End
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sample #3"
      TitleTop        =   7
      icon            =   "frm_test.frx":905C
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      CtlAutoChangeScheme=   0   'False
      UseDefaultTheme =   0   'False
   End
End
Attribute VB_Name = "frm_test"
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
    
    ' Initialize data
    Me.OsenXPListBox1.InsertItemBySQL ADOCN, "select * from customers", , , True
    Me.OsenXPListBox2.InsertItemBySQL ADOCN, "select * from suppliers", , , True
    Me.OsenXPListBox3.InsertItemBySQL ADOCN, "select * from products", , , True
    
    
End Sub


Private Sub Repos1()

    Me.OsenXPFrame2.Top = Me.OsenXPFrame1.Top + Me.OsenXPFrame1.Height + 60
    Me.OsenXPFrame3.Top = Me.OsenXPFrame2.Top + Me.OsenXPFrame2.Height + 60
    
End Sub

Private Sub Repos2()

    Me.OsenXPFrame5.Top = Me.OsenXPFrame4.Top + Me.OsenXPFrame4.Height + 60
    Me.OsenXPFrame6.Top = Me.OsenXPFrame5.Top + Me.OsenXPFrame5.Height + 60
    
End Sub


Private Sub OsenXPFrame1_DropDownClick()
    Repos1
End Sub
Private Sub OsenXPFrame2_DropDownClick()
    Repos1
End Sub
Private Sub OsenXPFrame3_DropDownClick()
    Repos1
End Sub
Private Sub OsenXPFrame4_DropDownClick()
    Repos2
End Sub
Private Sub OsenXPFrame5_DropDownClick()
    Repos2
End Sub
Private Sub OsenXPFrame6_DropDownClick()
    Repos2
End Sub























