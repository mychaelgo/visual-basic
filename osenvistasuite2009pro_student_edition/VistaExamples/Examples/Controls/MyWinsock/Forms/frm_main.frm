VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_Main 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "MyWinsock Sample"
   ClientHeight    =   3885
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   4890
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
   ScaleHeight     =   259
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton4 
      Height          =   405
      Left            =   1350
      TabIndex        =   6
      Top             =   3180
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   714
      Caption         =   "Send Email"
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
      MICON           =   "frm_main.frx":058A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton3 
      Height          =   405
      Left            =   1350
      TabIndex        =   5
      Top             =   2670
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   714
      Caption         =   "FTP Upload"
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
      MICON           =   "frm_main.frx":06EC
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   405
      Left            =   1350
      TabIndex        =   4
      Top             =   2130
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   714
      Caption         =   "Ping && Trace route"
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
      MICON           =   "frm_main.frx":084E
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   405
      Left            =   1350
      TabIndex        =   3
      Top             =   1590
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   714
      Caption         =   "Basic Demo"
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
      MICON           =   "frm_main.frx":09B0
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   4890
      _ExtentX        =   8625
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
      Picture         =   "frm_main.frx":0B12
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The following is example of MyWinsock Usage"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Control Name: MyWinsock "
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
      BinaryImage     =   "frm_main.frx":1764
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4890
      _ExtentX        =   8625
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
      Caption         =   "MyWinsock Sample"
      TitleTop        =   7
      icon            =   "frm_main.frx":177C
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
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
Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
   
End Sub

Private Sub OsenXPButton1_Click()
    frmServer.Show
End Sub

Private Sub OsenXPButton2_Click()
    frm_ping.Show
End Sub

Private Sub OsenXPButton3_Click()
    frm_upload.Show
End Sub

Private Sub OsenXPButton4_Click()
    frm_contactUs.Show 1
End Sub







