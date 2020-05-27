VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Tip of the Day"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTip.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4320
      Top             =   1320
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   2880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3075
      Left            =   120
      Picture         =   "frmTip.frx":2B01
      ScaleHeight     =   3015
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Byte
'The in-memory database of tips.
Dim Tips As New Collection

' Nama File Tipsnya
Const TEKS_FILE = "Q!W#T.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()
    'Kalo Tips pengen RANDOM
    'CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    'Kalo Pengen tipsnya MUTER
    
    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    'Tampil Euy
    frmTip.DisplayCurrentTip
    
End Sub

Function TampilTips(sFile As String) As Boolean
    Dim NextTip As String   ' Tiap tips dibaca dari file
    Dim InFile As Integer   ' Descriptor Bwat file
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    'Pastiin File-nya Udah ditentuin
    If sFile = "" Then
        TampilTips = False
        Exit Function
    End If
    
    ' Pastiin File Ada sblm nyoba dibuka
    If Dir(sFile) = "" Then
        TampilTips = False
        Exit Function
    End If
    
    'Ngebaca text yang ada di dalem file
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    'tampil di random
    DoNextTip
    
    TampilTips = True
    
End Function


Private Sub cmdNextTip_Click()
    DoNextTip
End Sub
Sub TENGAH()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub
Private Sub cmdOK_Click()
  Me.Hide
  Pembukaan.Timer2.Enabled = False
LOGIN.Show
End Sub

Private Sub Form_Load()
'    Dim Tampil_Start_Up As Long
'        'Kalo pengen Ditampilin pas startup
'      Tampil_Start_Up = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
'    If Tampil_Start_Up = 0 Then
'       Unload Me
'        Exit Sub
'    End If
'        TENGAH
'    'Set CHECKBOX, ini akan mengeksekusi nilai untuk dituliskan kedalam registry
'    Me.chkLoadTipsAtStartup.Value = vbChecked
'
'    ' Seed Rnd
'    Randomize
    
    ' Read in the tips file and display a tip at random.
    If TampilTips(App.Path & "\Document\" & TEKS_FILE) = False Then
        lblTipText.Caption = "That the " & TEKS_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TEKS_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If
    Pembukaan.Timer2.Enabled = False
   Timer1.Enabled = True
    Timer1.Interval = 20
    MakeTaskbarTransparent frmTip.hWnd, 0
    'MakeFlat cmdOK.hWnd
    'MakeFlat cmdNextTip.hWnd
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub






Private Sub Timer1_Timer()
  For i = 0 To 255
        MakeTaskbarTransparent frmTip.hWnd, i
        If i >= 254 Then
            MakeTaskbarTransparent frmTip.hWnd, 255
            Timer1.Enabled = False
            
            Exit Sub
        End If
    Next

End Sub
