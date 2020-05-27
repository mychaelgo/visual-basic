VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form frmTest 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaButton OsenVistaButton2 
      Height          =   375
      Left            =   9390
      TabIndex        =   3
      Top             =   7470
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "Test"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frmtest.frx":0000
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frmtest.frx":001C
      BinaryImageOver =   "frmtest.frx":0034
   End
   Begin VistaSuitePro.OsenVistaButton OsenVistaButton1 
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   7440
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      Caption         =   "Expand / Collape"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frmtest.frx":004C
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frmtest.frx":0068
      BinaryImageOver =   "frmtest.frx":0080
   End
   Begin VistaSuitePro.OsenVistaListBox OsenVistaListBox1 
      Height          =   6885
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   12144
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   0
      BackSelected    =   12563634
      BackSelectedG1  =   16777215
      BackSelectedG2  =   12632256
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      HeaderFormatString=   "Column1;100;0;0;;-1|Column2;100;0;0;;-1|Column3;100;0;0;;-1|Column4;100;0;0;;-1"
      Columns         =   4
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      MaxAllColumnWidth=   400
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
      HeaderCaption   =   "OsenVistaListBox1"
      TransparencyLevel=   22
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorOver =   12958375
      BinaryImage     =   "frmtest.frx":0098
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaForm OsenVistaForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Form1"
      TitleTop        =   7
      BorderStyle     =   1
      UseDefaultTheme =   0   'False
      WindowColor     =   3
   End
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   34440
      Images          =   "frmtest.frx":00B0
      Version         =   196608
      KeyCount        =   30
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As SQLite3_Connection
Dim Rs As CLS_ADODB_Recordset

Private Sub Form_Load()
    Set cn = New SQLite3_Connection
    cn.OpenDB ".\account.db3"
    Set Rs = New CLS_ADODB_Recordset
    Rs.RsOpen cn, "select * from groupaccount"
End Sub

Private Sub OsenVistaButton1_Click()
    Static b As Boolean
    Me.OsenVistaListBox1.ExpandAllNodes b
    b = Not b
End Sub

Private Sub OsenVistaButton2_Click()
    Me.OsenVistaListBox1.AddItemsByRecordset cn.Recordset("select * from account"), "3", basecolumnforsummary:=4
End Sub

Private Sub OsenVistaListBox1_DrawGroupItems(Node As VistaSuitePro.CLS_TreeList, Caption() As String, Alignment() As Integer, Forecolor() As Long, ReDraw As Boolean)
    Dim data(4) As String
    Dim adata(4) As Integer
    Dim fdata(4) As Long
    
    data(0) = Node.Caption
    Rs.Filter = "parent=" & Node.Caption
    data(1) = Rs.sField(1)
    data(2) = Me.OsenVistaListBox1.Cell(Node.Lists(1), 2)
    
    data(4) = Format$(Node.Summary, "0,000.00")
    
    adata(3) = 1
    adata(4) = 2
    fdata(4) = vbRed
    
    Caption = data
    Alignment = adata
    Forecolor = fdata
    
    ReDraw = True
    
End Sub

