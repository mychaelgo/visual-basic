VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_Main 
   BackColor       =   &H00EAF7F7&
   BorderStyle     =   0  'None
   Caption         =   "MyContainerCtl Sample"
   ClientHeight    =   3795
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   5220
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
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   435
      Left            =   1290
      TabIndex        =   3
      Top             =   2370
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   767
      Caption         =   "Picture Scrolling"
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
      MPTR            =   99
      MICON           =   "frm_main.frx":038A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_main.frx":04EC
      BinaryImageOver =   "frm_main.frx":0504
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   435
      Left            =   1290
      TabIndex        =   2
      Top             =   1710
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   767
      Caption         =   "Sales Order"
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
      MPTR            =   99
      MICON           =   "frm_main.frx":051C
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_main.frx":067E
      BinaryImageOver =   "frm_main.frx":0696
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   5220
      _ExtentX        =   9208
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
      Picture         =   "frm_main.frx":06AE
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   15177840
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The following is example of MyContainerCtl Usage"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Control Name: MyContainerCtl "
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
      BinaryImage     =   "frm_main.frx":1300
      WindowColor     =   1
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5220
      _ExtentX        =   9208
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
      Caption         =   "MyContainerCtl Sample"
      TitleTop        =   7
      icon            =   "frm_main.frx":1318
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      WindowColor     =   1
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************'
'*  OsenVistaSuite 2008 -  MyContainerCTL sample                           *'
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
  
    ' prepare connectionstring for opening database
    mStrSQL = MsAccessConnString(mStrSQL)
       
    ' Open the database connection right now ...
    If ADO_OPEN(mStrSQL) Then
        Debug.Print "Connection successful"
        ' Now, the AdoCN has been set
    Else
        MsgBoxGT "The database can't be connected!", vbExclamation, "Connection failed"
    End If
    
End Sub

Private Sub OsenXPButton1_Click()
    frm_orders.Show
End Sub

Private Sub OsenXPButton2_Click()
    frm_image.Show 1
End Sub



