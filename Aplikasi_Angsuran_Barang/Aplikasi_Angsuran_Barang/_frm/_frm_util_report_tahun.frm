VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tahun Pendapatan"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_util_report_tahun.frx":0000
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SysInfo_Nardhika.vbButton vbButton1 
      Height          =   375
      Left            =   3090
      TabIndex        =   4
      Top             =   810
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "_frm_util_report_tahun.frx":038A
      PICN            =   "_frm_util_report_tahun.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   1215
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   33023
      AutoTab         =   -1  'True
      FocusBackColor  =   12640511
      FocusForeColor  =   8388736
      FocusBackMainColor=   8438015
      FocusBorderColor=   33023
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      ForeColor       =   0
      MaxLength       =   4
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   825
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   33023
      AutoTab         =   -1  'True
      FocusBackColor  =   12640511
      FocusForeColor  =   8388736
      FocusBackMainColor=   8438015
      FocusBorderColor=   33023
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      ForeColor       =   0
      MaxLength       =   4
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbButton vbButton2 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3090
      TabIndex        =   5
      Top             =   1230
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "_frm_util_report_tahun.frx":0740
      PICN            =   "_frm_util_report_tahun.frx":075C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   75
      X2              =   4260
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   60
      X2              =   4260
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "_frm_util_report_tahun.frx":0AF6
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Masukan Tahun Pendapatan yang akan ditampilkan"
      Height          =   435
      Left            =   645
      TabIndex        =   6
      Top             =   105
      Width           =   2805
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mulai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   885
      Width           =   435
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sampai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   1275
      Width           =   600
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next
If CekUser("23", "P") = False Then
   ShowDlgMsg Me, "Anda tidak diperkenankan untuk mencetak laporan ini", vbOK, , True, False
Else
   HideMenu Me.hwnd
End If
End Sub

Private Sub vbButton1_Click()
On Error Resume Next
SelectMsg = vbOK
varTgl1 = txtFields(0).Text
varTgl2 = txtFields(1).Text
Unload Me
End Sub

Private Sub vbButton2_Click()
On Error Resume Next
SelectMsg = vbCancel
Unload Me
End Sub
