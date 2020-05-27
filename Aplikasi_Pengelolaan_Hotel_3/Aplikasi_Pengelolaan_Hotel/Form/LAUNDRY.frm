VERSION 5.00
Begin VB.Form LAUNDRY2 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3015
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Width           =   8415
      Begin Project1.vbbutton btnNew 
         Height          =   495
         Left            =   5640
         TabIndex        =   12
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&NEW"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "LAUNDRY.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox lstQTY 
         Height          =   2205
         ItemData        =   "LAUNDRY.frx":001C
         Left            =   1680
         List            =   "LAUNDRY.frx":001E
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.ListBox LstItem 
         Height          =   2205
         ItemData        =   "LAUNDRY.frx":0020
         Left            =   360
         List            =   "LAUNDRY.frx":0022
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin Project1.vbbutton btnSAVE 
         Height          =   495
         Left            =   5640
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&SAVE"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "LAUNDRY.frx":0024
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "quantity"
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
         Height          =   240
         Left            =   1680
         TabIndex        =   11
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "item "
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
         Height          =   240
         Left            =   480
         TabIndex        =   9
         Top             =   120
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.ComboBox cmbIdLaundry 
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbIDGuest 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "idlaundry"
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
         Height          =   240
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "idGuest"
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
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   900
         Left            =   5880
         TabIndex        =   2
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   660
         Left            =   1800
         TabIndex        =   1
         Top             =   120
         Width           =   135
      End
   End
End
Attribute VB_Name = "LAUNDRY2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmbIdLaundry_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from itemlaundry where idlaundry=" & cmbIdLaundry & "", KOneKsi, 3, 3
        If Not Rs.EOF Then
            LstItem.AddItem Rs!Name
        End If
End Sub

Private Sub Form_Load()
OPENDATA
End Sub

Private Sub LstItem_Click()
Dim Qty As String
Qty = InputBox("Please Enter Quantity" + " " + LstItem, "mYHoTEL")

If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from itemlaundry where name='" & LstItem & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            lstQTY.AddItem Qty
        End If

End Sub
