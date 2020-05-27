VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_change_password 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Change Password ..."
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   Icon            =   "frm_change_password.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaPicture OsenVistaPicture1 
      Align           =   1  'Align Top
      Height          =   825
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   1455
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBorderColor  =   0   'False
      Description     =   "Type the password you want to use."
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Change your password ..."
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
      BinaryImage     =   "frm_change_password.frx":038A
      WindowColor     =   0
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Change Password ..."
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin VistaSuitePro.OsenVistaButton CmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   1
      Top             =   2820
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      Caption         =   "&OK"
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
      MICON           =   "frm_change_password.frx":03A2
      PICN            =   "frm_change_password.frx":03BE
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      BinaryImageNormal=   "frm_change_password.frx":0958
      BinaryImageOver =   "frm_change_password.frx":0970
   End
   Begin VistaSuitePro.OsenVistaButton cmdCancel 
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   2820
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      MICON           =   "frm_change_password.frx":0988
      PICN            =   "frm_change_password.frx":09A4
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      BinaryImageNormal=   "frm_change_password.frx":0F3E
      BinaryImageOver =   "frm_change_password.frx":0F56
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtUser 
      Height          =   360
      Left            =   180
      TabIndex        =   3
      Top             =   1410
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "•"
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelBackColor  =   15790320
      LabelCaption    =   "Old Password:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtPwd 
      Height          =   360
      Left            =   180
      TabIndex        =   4
      Top             =   1830
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "•"
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelBackColor  =   15790320
      LabelCaption    =   "New Password:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox OsenXPTextBox1 
      Height          =   360
      Left            =   180
      TabIndex        =   5
      Top             =   2250
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "•"
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelBackColor  =   15790320
      LabelCaption    =   "Confrim Password:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
End
Attribute VB_Name = "frm_change_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Do Nothing
Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    On Error Resume Next

    ' Xp Form initialize
    Me.OsenXPForm1.Init Me
    
    ' Draw gradient color for Pic1
'    DrawGradient4Pic Picture1
    
End Sub


















