VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PasWoRdx dOoNg!"
   ClientHeight    =   2040
   ClientLeft      =   3585
   ClientTop       =   3345
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Juice ITC"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menu.frx":0000
   ScaleHeight     =   2040
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   0
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   0
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "open"
      Height          =   495
      Left            =   2520
      MaskColor       =   &H0080FFFF&
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   480
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "masukkan paswordx :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
MDIForm1.Visible = True
Form2.Show
Unload Form1
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Text1_Change()
If Text1.Text = "pasword" Then
Command1.Enabled = True
Timer1.Enabled = True
Timer2.Enabled = True
Label2.Caption = "iyyow itu paswordx!"
Else
Command1.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
Label2.ForeColor = vbRed
Label2.Caption = "bukan itu paswordx!"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MousePointer = Default Then
Text1.BackColor = vbWhite
Text1.ForeColor = vbBlack
End If
End Sub

Private Sub text1_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
End Sub

Private Sub Timer1_Timer()
j = Timer1.Interval
For j = 1 To 250
If j Mod 2 = 0 Then
Label2.ForeColor = vbBlue
Else
If j Mod 2 = 1 Then
Label2.ForeColor = vbYellow
Else
End If
End If
Next
End Sub

Private Sub Timer2_Timer()
j = Timer1.Interval
For j = 250 To 499
If j Mod 2 = 0 Then
Label2.ForeColor = vbBlue
Else
If j Mod 2 = 1 Then
Label2.ForeColor = vbYellow
Else
End If
End If
Next
End Sub


