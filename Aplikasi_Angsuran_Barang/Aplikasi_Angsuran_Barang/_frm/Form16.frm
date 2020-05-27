VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Product"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00F4FEFF&
      Height          =   285
      Left            =   210
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3795
      Width           =   5010
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   210
      TabIndex        =   3
      Top             =   4440
      Width           =   5010
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Register"
      Height          =   375
      Left            =   2805
      TabIndex        =   4
      Top             =   4890
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4050
      TabIndex        =   5
      Top             =   4890
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vbbego.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   3405
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   3540
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "License Key"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   3525
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activation Key"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   4170
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form16.frx":038A
      ForeColor       =   &H00404040&
      Height          =   2700
      Left            =   150
      TabIndex        =   7
      Top             =   540
      Width           =   5235
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Registration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   6
      Top             =   225
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   15
      X2              =   5490
      Y1              =   3225
      Y2              =   3225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   15
      X2              =   5490
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3210
      Left            =   -195
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim h As String
Static CobaAh As Byte
h = Replace(MySerial, "-", "")
h = Crypto(h, Chr(255) & Chr(254) & Chr(253), "vbbego.com")
Dim i As Integer, X As String
For i = 1 To Len(h)
   If i Mod 4 = 0 Then
      X = X & Hex(Asc(Mid(h, i, 1))) & "-"
   Else
      X = X & Hex(Asc(Mid(h, i, 1)))
   End If
Next i
h = IIf(Right(X, 1) = "-", Left(X, Len(X) - 1), X)
Clipboard.Clear
Clipboard.SetText h
If h = Text2 Then
   MsgBox "Aktifasi kode berhasil." & vbCrLf & vbCrLf & _
          "License Key:" & Text1 & vbCrLf & _
          "ActivationKey:" & Text2 & vbCrLf & vbCrLf & _
          "Terima kasih atas pembelian produk kami, simpan" & vbCrLf & _
          "baik-baik aktifasi kode program tersebut." & vbCrLf & vbCrLf & _
          "Apabila terjadi kesulitan dalam penggunaan program" & vbCrLf & _
          "hubungi kami di support@vbbego.com", 64
    
    If LockUnlock(StripPath(App.Path) & "_support\_syslog.sys", False) Then
       If srvUSER.State = 1 Then srvUSER.Close
       srvUSER.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_support\_syslog.sys;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=False"
       LockUnlock StripPath(App.Path) & "_support\_syslog.sys", True
       srvUSER.Execute "INSERT INTO SETTINGS(NAMA,ISI) VALUES('license','" & Replace(Text2, "'", "") & "')"
    End If
  End
Else
  MsgBox "Aktifasi kode yang anda masukan salah, silahkan" & vbCrLf & _
         "masukan kode aktifasi yang valid." & vbCrLf & vbCrLf & _
         "Untuk mendapatkan kode aktifasi yang valid, silahkan" & vbCrLf & _
         "hubungi kami:" & vbCrLf & vbCrLf & _
         "   email: market@vbbego.com" & vbCrLf & _
         "   home site: www.vb-bego.com", 48, "Product Activation Error"
  If CobaAh = 2 Then
     End
  End If
  CobaAh = CobaAh + 1
  Text2.SelStart = 0
  Text2.SelLength = 100
  Text2.SetFocus
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If pubMenu Then
   Unload Me
Else
   End
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Text1 = MySerial
    HideMenu Me.hwnd
End Sub

Private Sub Label1_Click(index As Integer)
On Error Resume Next
Shell "explorer " & Label1(index), 1
End Sub
