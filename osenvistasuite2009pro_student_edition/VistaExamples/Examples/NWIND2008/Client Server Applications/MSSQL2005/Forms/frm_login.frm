VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_login 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   5
      Top             =   420
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   1561
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
      Picture         =   "frm_login.frx":058A
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "Please enter user name and password to connect to the server ..."
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "User Information"
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
      BinaryImage     =   "frm_login.frx":20DC
      WindowColor     =   0
   End
   Begin VistaSuitePro.OsenVistaButton CmdLogin 
      Default         =   -1  'True
      Height          =   345
      Left            =   3150
      TabIndex        =   2
      Top             =   1500
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Caption         =   "&Log In"
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
      MICON           =   "frm_login.frx":20F4
      PICN            =   "frm_login.frx":2110
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      BinaryImageNormal=   "frm_login.frx":26AA
      BinaryImageOver =   "frm_login.frx":26C2
   End
   Begin VistaSuitePro.OsenVistaButton cmdCancel 
      Height          =   345
      Left            =   3150
      TabIndex        =   3
      Top             =   1950
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
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
      MICON           =   "frm_login.frx":26DA
      PICN            =   "frm_login.frx":26F6
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      BinaryImageNormal=   "frm_login.frx":2C90
      BinaryImageOver =   "frm_login.frx":2CA8
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtUser 
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   1500
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   582
      Text            =   "Admin"
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      AutoTab         =   -1  'True
      LabelBackColor  =   15790320
      LabelCaption    =   "User Name:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   75
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4605
      _ExtentX        =   8123
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
      Caption         =   "Login"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      AllowFadeOut    =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtPwd 
      Height          =   360
      Left            =   210
      TabIndex        =   1
      Top             =   1920
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   635
      Text            =   "vb"
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
      BackColorOver   =   12648447
      AutoTab         =   -1  'True
      LabelBackColor  =   15790320
      LabelCaption    =   "Password:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   75
      LabelStyle      =   2
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Ncount As Integer

Private valid As Boolean

Private Sub cmdCancel_Click()
    On Error Resume Next
    
    Unload Me
    
End Sub

Private Sub CmdLogin_Click()
    
    StrUserID = TxtUser
    Ncount = Ncount + 1
    ' Prepared Query
    mStrSQL = "select * from users where userid='" & StrUserID & "' and password='" & TxtPwd & "'"
    
    ' Execute current query
    If GetRST(mStrSQL).RecordCount Then  ' user validation
        ' valid user
        StrUserName = ADO_SQL_RESULT(mStrSQL, , 1)
        valid = 1
        
        DoEvents
        Unload Me
        
    Else
    
        valid = 0
        MsgBoxGT "Access denied for user " & TxtUser, vbCritical, "Login Failed", 5

        If Ncount = 3 Then
            Unload Me
            CloseProgram
        End If
    End If
    
End Sub

Private Sub Form_Activate()
    TxtUser.SetFocusEx
End Sub

Private Sub Form_Load()
    On Error Resume Next

    ' Xp Form initialize
    Me.OsenXPForm1.Init Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If valid Then
        If AlreadyExist Then
            frmMain.CreateNode
        Else
            Load frmMain
            frmMain.Show
        End If
    End If

End Sub


















