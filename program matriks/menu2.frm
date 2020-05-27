VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "PrOgRaM MatRiks [MN]"
   ClientHeight    =   5970
   ClientLeft      =   1650
   ClientTop       =   1650
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Juice ITC"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "menu2.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image10 
      Height          =   615
      Left            =   5040
      Picture         =   "menu2.frx":BF21
      ToolTipText     =   "kolom jadi baris"
      Top             =   2520
      Width           =   2250
   End
   Begin VB.Image Image9 
      Height          =   615
      Left            =   5040
      Picture         =   "menu2.frx":CC86
      ToolTipText     =   "exit program"
      Top             =   4200
      Width           =   2250
   End
   Begin VB.Image Image8 
      Height          =   615
      Left            =   5040
      Picture         =   "menu2.frx":D6D4
      ToolTipText     =   "skalar matriks"
      Top             =   3360
      Width           =   2250
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   5040
      Picture         =   "menu2.frx":E09F
      ToolTipText     =   "perhitungan matriks (+,/,-,*)"
      Top             =   1680
      Width           =   2250
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "alamat email and facebookx emend.ohyeah@yahoo.com. tammand diia he"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   7335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "namax emend, dia i2 skuul di esteem jurusand teerpeel"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   5160
      Width           =   6495
   End
   Begin VB.Image Image5 
      Height          =   3000
      Left            =   600
      Picture         =   "menu2.frx":EA64
      ToolTipText     =   "apa tunjuk2!"
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "ExiIt"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "sKaLaR"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "tRaNsPoS"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "oRdo"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   2640
      Left            =   1080
      Picture         =   "menu2.frx":12F48
      Top             =   2040
      Width           =   2550
   End
   Begin VB.Image Image3 
      Height          =   2640
      Left            =   720
      Picture         =   "menu2.frx":149F0
      Top             =   2040
      Width           =   2700
   End
   Begin VB.Image Image2 
      Height          =   2640
      Left            =   960
      Picture         =   "menu2.frx":174DB
      Top             =   2160
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   960
      Picture         =   "menu2.frx":1A0B3
      Top             =   2040
      Width           =   2700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "matriksx emend"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   -120
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = True
Label6.Visible = False
Label7.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MousePointer = Default Then
Label1.Caption = "matriksx emend"
Image5.Visible = True
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Label6.Visible = False
Label7.Visible = False
Image6.Visible = False
Image10.Visible = False
Image8.Visible = False
Image9.Visible = False
End If
End Sub

Private Sub image5_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.Visible = True
Label7.Visible = True
Label1.Caption = "emend ohyeah!"
End Sub

Private Sub Image8_Click()
MDIForm1.Visible = False
Form5.Show
End Sub

Private Sub Image9_Click()
Unload MDIForm1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Label1.Caption = "MATriksx emend"
Image6.Visible = True
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
End Sub

Private Sub label3_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Label1.Caption = "matRIKSx emend"
Image10.Visible = True
Image2.Visible = True
Image1.Visible = False
Image3.Visible = False
Image4.Visible = False
End Sub

Private Sub label4_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Label1.Caption = "matriksX emend"
Image8.Visible = True
Image3.Visible = True
Image1.Visible = False
Image2.Visible = False
Image4.Visible = False
End Sub

Private Sub label5_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Label1.Caption = "matriksx EMEND"
Image9.Visible = True
Image4.Visible = True
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
End Sub

Private Sub image6_Click()
MDIForm1.Visible = False
Form4.Show
End Sub

Private Sub image10_Click()
MDIForm1.Visible = False
Form3.Show
End Sub
