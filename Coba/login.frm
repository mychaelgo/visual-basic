VERSION 5.00
Object = "{EE17B266-A61D-48F0-BB3E-5C4EC9EE2D1D}#1.1#0"; "osenxpsuite2009.ocx"
Begin VB.Form login 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   124
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   247
   StartUpPosition =   2  'CenterScreen
   Begin OSENXPSUITE2009OCX.OsenXPButton cmd_login 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Login"
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
      MICON           =   "login.frx":0000
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
      Style           =   1
      BinaryImageNormal=   "login.frx":001C
      BinaryImageOver =   "login.frx":0034
   End
   Begin OSENXPSUITE2009OCX.OsenXPTextBox txt 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
      ForeColor       =   0
      DecimalSeparator=   ","
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      ThousandSeparator=   "."
   End
   Begin OSENXPSUITE2009OCX.OsenXPLabel lbl 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Username"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin OSENXPSUITE2009OCX.OsenXPLabel lbl 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Password"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin OSENXPSUITE2009OCX.OsenXPTextBox txt 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
      ForeColor       =   0
      PasswordChar    =   "*"
      DecimalSeparator=   ","
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      ThousandSeparator=   "."
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_login_Click()
    sambung
    sql = ("select * from login where user='" & txt(0).Text & "' and password='" & txt(1) & "' ")
    Set rs = con.Execute(sql)
    If Not rs.EOF Then
        mdi.Show
        login.Hide
    Else
        xp.MsgBoxXP "Periksa Username dan Password Anda...", vbInformation, "Konfirmasi"
        txt(0).SetFocus
    End If
End Sub

Private Sub Form_activate()
    txt(0).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    msg = xp.MsgBoxXP("Anda Yakin Tidak Ingin Login ???", vbYesNo, "Konfirmasi")
    If msg = vbYes Then
        Unload Me
    Else
        Cancel = 1
        txt(0).SetFocus
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


