VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   3840
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   5160
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.ListBox list1 
      Height          =   2595
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblkembalian 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label lblDiskon 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label lblJumlah 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblHarga 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblBarang 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "jumlah :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "pilih barang :"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
list1.AddItem "Disket"
list1.AddItem "Buku"
list1.AddItem "Kertas"
list1.AddItem "Pulpen"
End Sub
Private Sub Command1_Click()
Dim harga As Currency, total As Currency
Dim jumlah As Integer
Dim diskon As Single
Dim satuan As String
If list1.Text = "" Then
MsgBox "Anda belum memilih barang !!"
list1.ListIndex = 0
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "Anda belum mengisi jumlah barang !!"
Text1.SetFocus
Exit Sub
End If
Select Case list1.Text
Case "Disket"
harga = 35000
satuan = "Box"
Case "Buku"
harga = 20000
satuan = "Lusin"
Case "Kertas"
harga = 25000
satuan = "Rim"
Case "Pulpen"
harga = 10000
satuan = "Pak"
End Select
lblBarang.Caption = "Barang : " & list1.Text
lblHarga.Caption = "Harga : " & Format(harga, "Currency") & "/" & satuan
lblJumlah.Caption = "Jumlah : " & Text1.Text & " " & satuan
jumlah = Text1.Text
Select Case jumlah
Case Is < 10
diskon = 0
Case 10 To 20
diskon = 0.15
Case Else
diskon = 0.2
End Select
total = jumlah * (harga * (1 - diskon))
lblDiskon.Caption = "Diskon : " & Format(diskon, "0 %")
lblTotal.Caption = "Total Bayar : " & Format(total, "Currency")
Text2.Text = "total "
lblkembalian.Caption = "kembalian :  Rp " & Format(, "currency")
End Sub
