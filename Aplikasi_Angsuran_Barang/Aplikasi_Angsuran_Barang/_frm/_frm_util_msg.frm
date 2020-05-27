VERSION 5.00
Begin VB.Form frm_util_msg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informasi"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_util_msg.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Status 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1365
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   525
      Width           =   5220
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6420
      Top             =   1305
   End
   Begin SysInfo_Nardhika.vbButton cmd 
      Height          =   360
      Index           =   2
      Left            =   5505
      TabIndex        =   0
      Top             =   2265
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   635
      BTYPE           =   8
      TX              =   "c"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "_frm_util_msg.frx":239E
      PICN            =   "_frm_util_msg.frx":26B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfo_Nardhika.vbButton cmd 
      Height          =   360
      Index           =   1
      Left            =   4170
      TabIndex        =   1
      Top             =   2265
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   635
      BTYPE           =   8
      TX              =   "b"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "_frm_util_msg.frx":2A52
      PICN            =   "_frm_util_msg.frx":2D6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfo_Nardhika.vbButton cmd 
      Height          =   360
      Index           =   0
      Left            =   2835
      TabIndex        =   2
      Top             =   2265
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   635
      BTYPE           =   8
      TX              =   "a"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "_frm_util_msg.frx":3106
      PICN            =   "_frm_util_msg.frx":3420
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      TabIndex        =   3
      Top             =   2445
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   480
      Picture         =   "_frm_util_msg.frx":37BA
      Top             =   600
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   6345
      Picture         =   "_frm_util_msg.frx":4684
      Stretch         =   -1  'True
      Top             =   225
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label cgmHyperLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Informasi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5190
      TabIndex        =   4
      Top             =   255
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   1920
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   6705
   End
End
Attribute VB_Name = "frm_util_msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cgmGradientLabel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then MoveIt hWnd

End Sub

Private Sub cgmHyperLabel1_Click()
On Error Resume Next
Static X As Boolean
If X = False Then
   Status.Text = Image1.Tag
   X = True
   cgmHyperLabel1.ForeColor = &HFF0000
   Image3.Tag = Status.ForeColor
   Status.ForeColor = vbRed
Else
   Status.Text = Status.Tag
   cgmHyperLabel1.ForeColor = 0
   X = False
   Status.ForeColor = Image3.Tag
End If
End Sub

Private Sub Check1_Click()
On Error Resume Next
If cgmHyperLabel1.Tag = "" Then
   If Trim(Me.Tag) <> "" Then SaveSetting Me.Tag, "Message", "showme", Check1.Value
Else
   SaveSetting Me.Tag, "Message", cgmHyperLabel1.Tag, Check1.Value
End If
End Sub

Private Sub cmd_Click(Index As Integer)
On Error Resume Next
Select Case UCase(cmd(Index).Caption)
       Case "&OK": SelectMsg = vbOK
       Case "&YA": SelectMsg = vbYes
       Case "&TIDAK": SelectMsg = vbNo
       Case "&BATAL": SelectMsg = vbCancel
       Case "&COBA LAGI": SelectMsg = vbRetry
       Case "&BATALKAN": SelectMsg = vbAbort
       Case "&LANJUT": SelectMsg = vbIgnore
End Select
Unload Me
End Sub


Private Sub Form_Load()
On Error Resume Next
If GetSetting(vbReg, "Setting\View", "AllTransparent", "") <> "All" Then
   'MakeTransparent GetSetting(vbReg, "Setting\View\" & Me.Name, "Transparent", 255), Me.hWnd
Else
   'MakeTransparent GetSetting(vbReg, "Setting\View", "Transparent", 255), Me.hWnd
End If

Set cgmHyperLabel1.MouseIcon = LoadResPicture(101, 2)
cgmHyperLabel1.MousePointer = 99
Check1.Caption = "Tampilkan Dialog ini kembali"
HideMenu Me.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then MoveIt Me.hWnd
End Sub

Private Sub Status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then MoveIt Me.hWnd
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If cgmHyperLabel1.Visible Then Image3.Visible = Not Image3.Visible = True
End Sub



