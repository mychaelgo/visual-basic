VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_Main 
   BackColor       =   &H00EAF7F7&
   BorderStyle     =   0  'None
   Caption         =   "MyADODC Sample"
   ClientHeight    =   3915
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton3 
      Height          =   465
      Left            =   900
      TabIndex        =   5
      Top             =   3030
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   820
      Caption         =   "Complete MyADODC Demo"
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
      MICON           =   "Form1.frx":038A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "Form1.frx":03A6
      BinaryImageOver =   "Form1.frx":03BE
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   465
      Left            =   900
      TabIndex        =   4
      Top             =   2340
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   820
      Caption         =   "&MyADODC Customize"
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
      MICON           =   "Form1.frx":03D6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "Form1.frx":03F2
      BinaryImageOver =   "Form1.frx":040A
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   465
      Left            =   930
      TabIndex        =   3
      Top             =   1620
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   820
      Caption         =   "&Simple MyADODC Demo"
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
      MICON           =   "Form1.frx":0422
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "Form1.frx":043E
      BinaryImageOver =   "Form1.frx":0456
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   4665
      _ExtentX        =   8229
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
      Picture         =   "Form1.frx":046E
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   13089392
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The following is example of MyADODC Usage"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Control Name: MyADODC"
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
      BinaryImage     =   "Form1.frx":10C0
      WindowColor     =   2
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4665
      _ExtentX        =   8229
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
      Caption         =   "MyADODC Sample"
      TitleTop        =   7
      icon            =   "Form1.frx":10D8
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      WindowColor     =   2
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   2
      Top             =   5100
      Width           =   75
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenVistaSuite 2008 -  MyADODC sample                                *'
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

Private Sub OsenXPButton1_Click()

    ' Check connection status
    ' If the database opened, show frm_simple
    If ADOCN.State Then
        frm_simple.Show 1
    End If
    
End Sub

Private Sub OsenXPButton2_Click()

    ' Check connection status
    ' If the database opened, show frm_customize
    If ADOCN.State Then
        frm_customize.Show 1
    End If
    
End Sub

Private Sub OsenXPButton3_Click()

    ' Check connection status
    ' If the database opened, show frm_employees
    If ADOCN.State Then
        frm_employees.Show 1
    End If
    
End Sub























