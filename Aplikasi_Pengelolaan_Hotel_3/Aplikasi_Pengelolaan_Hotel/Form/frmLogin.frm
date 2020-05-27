VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LOGIN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LOGIN"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tUSER 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2040
      Width           =   2325
   End
   Begin VB.ComboBox cmbLevel 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "frmLogin.frx":0000
      Left            =   3720
      List            =   "frmLogin.frx":000A
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   3600
      Left            =   0
      ScaleHeight     =   3540
      ScaleWidth      =   7995
      TabIndex        =   4
      Top             =   0
      Width           =   8055
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   120
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   120
         Top             =   600
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   390
         Left            =   6360
         TabIndex        =   3
         Top             =   1920
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   390
         Left            =   6360
         TabIndex        =   5
         Top             =   2280
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3720
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3000
         Width           =   2325
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   2040
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   0
         Top             =   1560
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   120
         Top             =   1080
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         MousePointer    =   2
         Max             =   255
         Scrolling       =   1
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password  :"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   2
         Left            =   1080
         TabIndex        =   15
         Top             =   2880
         Width           =   1980
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2040
         Left            =   5400
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2040
         Left            =   4800
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2040
         Left            =   2160
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2040
         Left            =   720
         TabIndex        =   11
         Top             =   -120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2040
         Left            =   3480
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name  :"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   1920
         Width           =   2040
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&LEVEl            :"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   7560
         TabIndex        =   7
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image1 
         Height          =   5385
         Left            =   -720
         Picture         =   "frmLogin.frx":0023
         Stretch         =   -1  'True
         Top             =   -1080
         Width           =   9720
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmLogin.frx":1DA18
      Top             =   480
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Const Invert = 1

Option Explicit

Dim x As Byte
Dim D As Byte
Dim F As Byte
Dim z As Byte
Dim g As Byte
Dim pesan As String
Dim m1, m2, m3, m4, m5, m6 As Byte
Dim B As Byte
Public LoginSucceeded As Boolean
Public Function TranslucentForm(frm As Form, TranslucenceLevel As Byte) As Boolean
SetWindowLong frm.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes frm.hWnd, 0, TranslucenceLevel, LWA_ALPHA
TranslucentForm = Err.LastDllError = 0
End Function

Private Sub cmdCancel_Click()
   
    LoginSucceeded = False
   
    Timer3.Enabled = True
If Timer3.Enabled = False Then
    Unload Me
    Unload MAIN
End If

End Sub

Private Sub cmdOK_Click()
Dim OK As String

If Rs.State = 1 Then Rs.Close
   OK = "select * from tuser where NAMe='" & Replace(tUSER, "'", "''") & "' and pwd='" & Replace(txtPassword, "'", "''") & "' and levels='" & cmbLevel & "'"
   
   Rs.Open OK, KOneKsi, 3, 3
   Debug.Print OK
   If Not Rs.EOF Then
        If cmbLevel = "Administrator" Then
    
   
                Timer4.Enabled = True
            Else
            Timer4.Enabled = True
            LeveL = "User"
'                    If MAIN.Visible = True Then
'                           With MAIN
'                            .ADDLAUNDRY.Enabled = False
'                            .ADDREST.Enabled = False
'                            .USR.Enabled = False
'
'                          End With
'                    End If
                    End If

    xUser = tUSER
    Else
        bersih Me
        tUSER.SetFocus
       pesan = MsgBox("TANYA_KENAPA?", vbCritical, "COBA_LAGI")
        z = z + 1
        If z >= 3 Then
            pesan = MsgBox("BYE...BYE...BYE...", vbExclamation, "JANGAN PUTUS ASA")
            Timer3.Enabled = True
        End If
   End If

'If Rs.State = 1 Then Rs.Close
'   OK = "select * from tuser where NAMe='" & Replace(tUSER, "'", "''") & "' and pwd='" & Replace(txtPassword, "'", "''") & "' and level='User'"
'
'   Rs.Open OK, KOneKsi, 3, 3
'   Debug.Print OK
'   If Not Rs.EOF Then
'    MAIN.master.Enabled = False
'
'
'
''        MAIN.Timer2.Enabled = False
''
'            Timer4.Enabled = True '
'
'    Else
'        bersih Me
'        tUSER.SetFocus
'       PESAN = MsgBox("TANYA_KENAPA?", vbCritical, "COBA_LAGI")
'        Z = Z + 1
'        If Z >= 3 Then
'            PESAN = MsgBox("BYE...BYE...BYE...", vbExclamation, "JANGAN PUTUS ASA")
'            Timer3.Enabled = True
'        End If
'   End If
'
'



End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys ("{TAB}")
End If
End Sub

Private Sub Form_Load()
'On Error Resume Next
Skin1.LoadSkin App.Path & "\Document\paper.skn"
    Skin1.ApplySkin hWnd
    
OPENDATA
Buka Me
D = 10
x = 0
B = 56
F = 255
g = 255
m1 = Asc(Trim(Left(Label1.Caption, 1)))

m2 = Asc(Trim(Left(Label2.Caption, 1)))
m3 = Asc(Trim(Left(Label3.Caption, 1)))
m4 = Asc(Trim(Left(Label4.Caption, 1)))
m5 = Asc(Trim(Left(Label5.Caption, 1)))
m6 = Asc(Trim(Left(Label6.Caption, 1)))

m1 = m1 - B
m2 = m2 + B
m3 = m3 - B
m4 = m4 + B
m5 = m5 - B
m6 = m6 + B

TranslucentForm Me, 0
'MakeitSkin Me, Picture1

Timer1.Enabled = True
Timer2.Enabled = True
End Sub



Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label6.ForeColor = vbRed
End Sub


Private Sub Label6_Click()

Timer3.Enabled = True
If Timer3.Enabled = False Then
    Unload Me
    Unload MAIN
End If

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label6.ForeColor = vbBlack
End Sub






Private Sub Timer1_Timer()


Label1.Caption = Chr(m1 + x)
Label2.Caption = Chr(m2 - x)
Label3.Caption = Chr(m3 + x)
Label4.Caption = Chr(m4 - x)
Label5.Caption = Chr(m5 + x)
Label6.Caption = Chr(m6 - x)
'Label7.Caption = Right(Label7.Caption, Len(Label7.Caption) - 1) & Left(Label7.Caption, 1)
If x < B Then
x = x + 1
Else

    If Label1.Caption = "L" And Label2.Caption = "O" And Label3.Caption = "G" And Label4.Caption = "I" And Label4.Caption = "N" Then
    Timer1.Enabled = False
    Buka Me
    End If
End If


End Sub


Private Sub Timer2_Timer()


If D < 255 Then
TranslucentForm Me, D
ProgressBar1.Value = D
Else
If D = 255 Then
TranslucentForm Me, 255
ProgressBar1.Value = D
D = 5

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Timer2.Enabled = False
Timer5.Enabled = True

ProgressBar1.Visible = False
End If
End If
D = D + 1
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If F > 0 Then
TranslucentForm Me, F
Else
If F >= 100 Then
Timer3.Enabled = False
TranslucentForm Me, 0
End
End If
End If
F = F - 1
End Sub

Private Sub Timer4_Timer()
On Error Resume Next

If g >= 101 Then
TranslucentForm Me, g
Else
If g <= 0 Then

Timer4.Enabled = False
Unload Me
MAIN.Show
TranslucentForm Me, 5
TranslucentForm Me, 3
TranslucentForm Me, 1
TranslucentForm Me, 0

End If
End If
g = g - 1
End Sub

Private Sub Timer5_Timer()
 FlashWindow Me.hWnd, Invert
End Sub
