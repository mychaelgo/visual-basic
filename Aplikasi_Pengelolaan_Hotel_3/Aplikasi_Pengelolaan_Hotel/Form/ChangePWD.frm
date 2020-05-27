VERSION 5.00
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Begin VB.Form ChangePwd 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3060
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "ChangePWD.frx":0000
   ScaleHeight     =   1807.949
   ScaleMode       =   0  'User
   ScaleWidth      =   4971.718
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPassword2 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1320
      Top             =   2400
   End
   Begin LvButtons.lvButtons_H btnCLOSE 
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "&STOP"
      Top             =   1320
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Caption         =   "&STOP"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "ChangePWD.frx":7D7E
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnOK 
      Height          =   1035
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "OK!!!"
      Top             =   1200
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "ChangePWD.frx":824C
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   1845
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "INSERT PASSWORD"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "ChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCLOSE_Click()
Unload Me
End Sub

Private Sub btnOK_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from tuser where name='" & xUser & "' and pwd='" & txtPassword & "'"
        If Not Rs.EOF Then
            lblLabels.Caption = "Insert New Password"
            txtPassword2.Visible = True
          txtPassword2.SetFocus
          
    If txtPassword2 <> "" Then
            pesan = MsgBox("Are You Sure Change Password ???", vbQuestion + vbYesNo, "mYHoTEL")
            If pesan = vbYes Then
                KOneKsi.Execute " update tuser set pwd='" & txtPassword2 & "' where name='" & xUser & "'"
                MsgBox "New Password Is Sucessfully Changed ", vbInformation, "mYHoTEL"
                Unload Me
        txtPassword2 = ""
          End If
           Else
            txtPassword.Visible = False
            MsgBox "Please Enter New Password", vbExclamation, "mYHoTEL"
            End If
        Else
            bersih Me
                txtPassword.SetFocus
        End If

End Sub

Private Sub btnOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Timer1_Timer()
If lblLabels.ForeColor = &HFF00& Then
    lblLabels.ForeColor = vbBlack
Else
    lblLabels.ForeColor = &HFF00&
End If
End Sub
