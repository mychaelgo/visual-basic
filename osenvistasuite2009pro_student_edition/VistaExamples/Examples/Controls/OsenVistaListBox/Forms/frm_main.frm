VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_1Main 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "OsenXPListBox Sample"
   ClientHeight    =   6390
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaFrame fra3 
      Height          =   1515
      Left            =   150
      TabIndex        =   4
      Top             =   6330
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   2672
      Caption         =   "Spreadsheet Demo"
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
      BorderColor     =   14396553
      MouseIcon       =   "frm_main.frx":038A
      MousePointer    =   99
      Appearance      =   1
      DropDownButton  =   -1  'True
      Picture         =   "frm_main.frx":04EC
      PicturePosition =   2
      image           =   "frm_main.frx":138B
      BinaryImage     =   "frm_main.frx":1725
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaButton cmdAS 
         Height          =   405
         Left            =   270
         TabIndex        =   6
         Top             =   960
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   714
         Caption         =   "Advance Spreadsheet Demo"
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
         MICON           =   "frm_main.frx":173D
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":1759
         BinaryImageOver =   "frm_main.frx":1771
      End
      Begin VistaSuitePro.OsenVistaButton cmdSS 
         Height          =   405
         Left            =   270
         TabIndex        =   5
         Top             =   480
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   714
         Caption         =   "Simple Spreadsheet Demo"
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
         MICON           =   "frm_main.frx":1789
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":17A5
         BinaryImageOver =   "frm_main.frx":17BD
      End
   End
   Begin VistaSuitePro.OsenVistaFrame fra2 
      Height          =   2085
      Left            =   150
      TabIndex        =   3
      Top             =   4170
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   3678
      Caption         =   "Advance OsenXPListbox Demo"
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
      BorderColor     =   14396553
      MouseIcon       =   "frm_main.frx":17D5
      MousePointer    =   99
      Appearance      =   1
      DropDownButton  =   -1  'True
      image           =   "frm_main.frx":1937
      BinaryImage     =   "frm_main.frx":1ED1
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaButton cmd_exf 
         Height          =   405
         Left            =   1980
         TabIndex        =   11
         Top             =   1020
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   714
         Caption         =   "Export Data From OsenXPListbox"
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
         MICON           =   "frm_main.frx":1EE9
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":1F05
         BinaryImageOver =   "frm_main.frx":1F1D
      End
      Begin VistaSuitePro.OsenVistaButton cmd_csf 
         Height          =   405
         Left            =   1980
         TabIndex        =   10
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   714
         Caption         =   "Customize Search and Filter Dialog"
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
         MICON           =   "frm_main.frx":1F35
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":1F51
         BinaryImageOver =   "frm_main.frx":1F69
      End
      Begin VistaSuitePro.OsenVistaButton cmd_rs1 
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   767
         Caption         =   "Basic Editable Cell #1"
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
         MICON           =   "frm_main.frx":1F81
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":1F9D
         BinaryImageOver =   "frm_main.frx":1FB5
      End
      Begin VistaSuitePro.OsenVistaButton cmd_Rs2 
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         Caption         =   "Basic Editable Cell #2"
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
         MICON           =   "frm_main.frx":1FCD
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":1FE9
         BinaryImageOver =   "frm_main.frx":2001
      End
      Begin VistaSuitePro.OsenVistaButton cmd_lsql 
         Height          =   405
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   714
         Caption         =   "Customize ListBox Item by SQL (Recordset)"
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
         MICON           =   "frm_main.frx":2019
         PICN            =   "frm_main.frx":2035
         UMCOL           =   -1  'True
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":25CF
         BinaryImageOver =   "frm_main.frx":25E7
      End
   End
   Begin VistaSuitePro.OsenVistaFrame fra1 
      Height          =   2685
      Left            =   150
      TabIndex        =   2
      Top             =   1410
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4736
      Caption         =   ""
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
      BorderColor     =   14396553
      MouseIcon       =   "frm_main.frx":25FF
      MousePointer    =   99
      Appearance      =   1
      DropDownButton  =   -1  'True
      image           =   "frm_main.frx":2761
      BinaryImage     =   "frm_main.frx":2CFB
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaButton cmd1 
         Height          =   405
         Left            =   210
         TabIndex        =   12
         Top             =   450
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Basic Demo #1"
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
         MICON           =   "frm_main.frx":2D13
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2D2F
         BinaryImageOver =   "frm_main.frx":2D47
      End
      Begin VistaSuitePro.OsenVistaButton cmd2 
         Height          =   405
         Left            =   2460
         TabIndex        =   13
         Top             =   450
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Basic Demo with Image"
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
         MICON           =   "frm_main.frx":2D5F
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2D7B
         BinaryImageOver =   "frm_main.frx":2D93
      End
      Begin VistaSuitePro.OsenVistaButton cmd4 
         Height          =   405
         Left            =   2460
         TabIndex        =   14
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Multiple Checkbox"
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
         MICON           =   "frm_main.frx":2DAB
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2DC7
         BinaryImageOver =   "frm_main.frx":2DDF
      End
      Begin VistaSuitePro.OsenVistaButton cmd3 
         Height          =   405
         Left            =   210
         TabIndex        =   15
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Multiline at single row"
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
         MICON           =   "frm_main.frx":2DF7
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2E13
         BinaryImageOver =   "frm_main.frx":2E2B
      End
      Begin VistaSuitePro.OsenVistaButton cmd5 
         Height          =   405
         Left            =   210
         TabIndex        =   16
         Top             =   1500
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Drag and Drop"
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
         MICON           =   "frm_main.frx":2E43
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2E5F
         BinaryImageOver =   "frm_main.frx":2E77
      End
      Begin VistaSuitePro.OsenVistaButton cmd6 
         Height          =   405
         Left            =   2460
         TabIndex        =   17
         Top             =   1500
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Drag selected"
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
         MICON           =   "frm_main.frx":2E8F
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2EAB
         BinaryImageOver =   "frm_main.frx":2EC3
      End
      Begin VistaSuitePro.OsenVistaButton cmd7 
         Height          =   405
         Left            =   210
         TabIndex        =   18
         Top             =   2070
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Icon in each cell"
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
         MICON           =   "frm_main.frx":2EDB
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2EF7
         BinaryImageOver =   "frm_main.frx":2F0F
      End
      Begin VistaSuitePro.OsenVistaButton cmd8 
         Height          =   405
         Left            =   2460
         TabIndex        =   19
         Top             =   2070
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         Caption         =   "Large Icon"
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
         MICON           =   "frm_main.frx":2F27
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         Style           =   1
         BinaryImageNormal=   "frm_main.frx":2F43
         BinaryImageOver =   "frm_main.frx":2F5B
      End
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   1561
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
      Picture         =   "frm_main.frx":2F73
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   16310477
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The following is example of OsenXPListBox Usage"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Control Name: OsenXPListBox "
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "frm_main.frx":3BC5
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5130
      _ExtentX        =   9049
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
      Caption         =   "OsenXPListBox Sample"
      TitleTop        =   7
      icon            =   "frm_main.frx":3BDD
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
End
Attribute VB_Name = "frm_1Main"
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

'mStrSQL is global variable from XP Library (OsenXPSuite2006.XP class object) [String]
'MsAccessConnString is global function from XP Library (OsenXPSuite2006.XP class object) [Return: ConnectionString]
'Ado_Open is global function form XP Library (OsenXPSuite2006.XP Class Object) [Boolean; True if connection opened and FALSE if failed]
'AdoCN is global variable from XP Library (OsenXPSuite2006.XP class object) [ADODB.Connection]

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
    
    ' Get Application Path
    mStrSQL = App.Path
    
    ' Retrieve location of database file (nwind.mdb)
    mStrSQL = App.Path & "\..\nwind.mdb"
    
    ' FYI (For Your Information)
    Debug.Print mStrSQL
    
    ' prepare connectionstring for opening database
    mStrSQL = MsAccessConnString(mStrSQL)
    
    ' FYI
    Debug.Print mStrSQL
    
    ' Open the database connection right now ...
    If ADO_OPEN(mStrSQL) Then
        Debug.Print "Connection successful"
        ' Now, the AdoCN has been set
    Else
        MsgBoxGT "The database can't be connected!", vbExclamation, "Connection failed"
    End If
    
End Sub

Private Sub Repos()
    fra2.Top = fra1.Height + fra1.Top + 4
    fra3.Top = fra2.Height + fra2.Top + 4
End Sub

Private Sub fra1_DropDownClick()
    Repos
End Sub

Private Sub fra2_DropDownClick()
    Repos
End Sub

Private Sub fra3_DropDownClick()
    Repos
End Sub

Private Sub cmdSS_Click()
    frm_SS.Show 1
End Sub

Private Sub cmdAS_Click()
    frm_AS.Show 1
End Sub

Private Sub cmd_rs1_Click()
    frm_basic1.Show 1
End Sub

Private Sub cmd_Rs2_Click()
    frm_basic2.Show 1
End Sub

Private Sub cmd_csf_Click()
    frm_csf.Show
End Sub

Private Sub cmd_exf_Click()
    frm_exd.Show 1
End Sub

Private Sub cmd_lsql_Click()
    frm_SQL.Show 1
End Sub

Private Sub cmd1_Click()
    frm1_basic.Show 1
End Sub

Private Sub cmd2_Click()
    frm2_basic.Show 1
End Sub

Private Sub cmd3_Click()
    frm3_basic.Show
End Sub

Private Sub cmd4_Click()
    frm4_basic.Show 1
End Sub

Private Sub cmd5_Click()
    frm5_basic.Show 1
End Sub

Private Sub cmd6_Click()
    frm6_basic.Show 1
End Sub

Private Sub cmd7_Click()
    frm7_basic.Show 1
End Sub

Private Sub cmd8_Click()
    frm8_basic.Show 1
End Sub























