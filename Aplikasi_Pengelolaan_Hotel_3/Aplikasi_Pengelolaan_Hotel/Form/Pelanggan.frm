VERSION 5.00
Begin VB.Form GUest 
   Caption         =   "PeLanggan"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Height          =   6375
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4680
         Top             =   1560
      End
      Begin VB.ComboBox cmbAgama 
         Height          =   315
         ItemData        =   "Pelanggan.frx":0000
         Left            =   1200
         List            =   "Pelanggan.frx":0013
         TabIndex        =   30
         Top             =   4920
         Width           =   2055
      End
      Begin VB.ComboBox CmbSex 
         Height          =   315
         ItemData        =   "Pelanggan.frx":0041
         Left            =   1200
         List            =   "Pelanggan.frx":004B
         TabIndex        =   27
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox tPINCODE 
         Height          =   285
         Left            =   4440
         TabIndex        =   25
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox tNational 
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox tID 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&NEW"
         Height          =   975
         Left            =   3480
         TabIndex        =   13
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton cmdLAst 
         Caption         =   "last"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdPREV 
         Caption         =   "prev"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "next"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "first"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox tdate 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox tname 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox tAge 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox tAddR 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox tCity 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox tPhone 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   975
         Left            =   5040
         TabIndex        =   2
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox tIDCard 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Religion"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   5040
         Width           =   570
      End
      Begin VB.Label Label4 
         Caption         =   "Sex"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Pincode"
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "NationaLity"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IDGUEST"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Date of arrival"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Phone"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "City"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "IDCARD"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
   End
End
Attribute VB_Name = "GUest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub auto()
Dim X As String
With Rs
If .State = 1 Then .Close
.Open "select * from Guest ", KOneKsi, 3, 3

    If .RecordCount = 0 Then
        MsgBox "awal"
        tID = "FJ" + Format(Date, "yymm") + "001"
    Else
        .MoveLast
       
'        If Left(Rs(1), 9) = "FJ" + Format(Date, "yymm") Then
'        tID = "FJ" + Format(Date, "yymm") + "001"
'        MsgBox "else1"
'        Else
         
        tID = Right(Rs(1), 1) + 1
        MsgBox tID
        X = Len(tID)
        MsgBox " ini adalah Left(000, 3 - Len(tID)) " + " " + X
        tID = "FJ" + Format(Date, "yymm") + Left("000", 3 - Len(tID)) + tID
        
'        End If
    End If
End With
End Sub

Private Sub cmdADD_Click()
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from Guest where idguest='" & tID & "'", KOneKsi, 3, 3
    If Rs.EOF Then
    KOneKsi.Execute "insert into guest(arrivaldate,idguest,idcard,name,address,city,nationality,phone,sex,age,religion,pincode)values('" & tDate & "' ,'" & tID & " ',  '" & tIDCard & "', '" & tname & "', '" & tAddr & "','" & tcity & "',  '" & tNational & "','" & tPhone & "', '" & CmbSex & "', '" & tAge & "','" & cmbAgama & "', '" & tPINCODE & "')"
    MsgBox ("Data added. Room alloted for visitor"), vbInformation, "mYHoTEL"
    End If
End Sub

Private Sub cmdClear_Click()
bersih Me
tID.SetFocus

auto
'tDate = Format(Date, "dd/mm/yyyy")
Buka Me
End Sub

Private Sub Form_Load()
OPENDATA
End Sub

Private Sub Timer1_Timer()
tDate = Date
End Sub
