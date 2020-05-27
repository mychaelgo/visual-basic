VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaButton cmdTree 
      Height          =   405
      Left            =   7140
      TabIndex        =   5
      Top             =   7470
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   714
      Caption         =   "Treeview with custom event"
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
      MICON           =   "Form1.frx":0000
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":001C
      BinaryImageOver =   "Form1.frx":0034
   End
   Begin VistaSuitePro.OsenVistaButton cmdTvw 
      Height          =   405
      Left            =   4650
      TabIndex        =   4
      Top             =   7470
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
      Caption         =   "TreeView Version 2"
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
      MICON           =   "Form1.frx":004C
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":0068
      BinaryImageOver =   "Form1.frx":0080
   End
   Begin VistaSuitePro.OsenVistaButton OsenVistaButton2 
      Height          =   405
      Left            =   2850
      TabIndex        =   3
      Top             =   7470
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "Create TreeListView"
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
      MICON           =   "Form1.frx":0098
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":00B4
      BinaryImageOver =   "Form1.frx":00CC
   End
   Begin VistaSuitePro.OsenVistaButton OsenVistaButton1 
      Height          =   405
      Left            =   330
      TabIndex        =   2
      Top             =   7470
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   714
      Caption         =   "create treelist from recordset"
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
      MICON           =   "Form1.frx":00E4
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":0100
      BinaryImageOver =   "Form1.frx":0118
   End
   Begin VistaSuitePro.OsenVistaListBox vList 
      Height          =   6765
      Left            =   330
      TabIndex        =   1
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11933
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   0
      BackSelected    =   12648447
      BackSelectedG1  =   16777215
      BackSelectedG2  =   33023
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   1
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
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
      AllowSortItem   =   0   'False
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
      BinaryImage     =   "Form1.frx":0130
      TreeViewMode    =   -1  'True
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaForm OsenVistaForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
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
      Images          =   "Form1.frx":0148
      Version         =   196608
      KeyCount        =   30
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As SQLite3_Connection

Private Sub cmdTree_Click()
    frmTest.Show 1
End Sub

Private Sub cmdTvw_Click()
    With vList
        .Clear True
        .TreeListViewMode = -1
        .InsertColumn , "Name", 170
        .InsertColumn , "Parent", 80
        .InsertColumn , "Status", 80
        .LockUpdate = True
        .SmallIcons = Me.MyImageList1.hIml
        
        .AddItem "Node 1" & vbTab & "root" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node1"
        .AddItem "Node 2" & vbTab & "Node 1" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node2", strparentkey:="node1"
        .AddItem "Node 3" & vbTab & "Node 1" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node3", strparentkey:="node1"
        .AddItem "Node 4" & vbTab & "root" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node4"
        .AddItem "Node 5" & vbTab & "root" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node5"
        .AddItem "Node 6" & vbTab & "Node 5" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node6", strparentkey:="node5"
        .AddItem "Node 7" & vbTab & "Node 6" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node7", strparentkey:="node6"
        .AddItem "Node 8" & vbTab & "Node 7" & vbTab & "InProgress", iconcollections:=CInt(Rnd() * 29), strkey:="node8", strparentkey:="node7"
        .AddItem "Node 9" & vbTab & "Node 7" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node9", strparentkey:="node7"
        
        .AddItem "Node 6.1" & vbTab & "Node 5" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node61", strparentkey:="node5"
        .AddItem "Node 7.1" & vbTab & "Node 6.1" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node71", strparentkey:="node61"
        .AddItem "Node 8.1" & vbTab & "Node 7.1" & vbTab & "InProgress", iconcollections:=CInt(Rnd() * 29), strkey:="node81", strparentkey:="node71"
        .AddItem "Node 9.1" & vbTab & "Node 7.1" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node91", strparentkey:="node71"
        
        .AddItem "Node 7.2" & vbTab & "Node 6.1" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node72", strparentkey:="node61"
        .AddItem "Node 8.2" & vbTab & "Node 7.2" & vbTab & "InProgress", iconcollections:=CInt(Rnd() * 29), strkey:="node82", strparentkey:="node72"
        .AddItem "Node 9.2" & vbTab & "Node 7.2" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node92", strparentkey:="node72"
        
        .AddItem "Node End" & vbTab & "Node 5" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="nodee", strparentkey:="node5"
        .AddItem "Node 10" & vbTab & "root" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node10"
        .AddItem "Node 11" & vbTab & "Node 10" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node11", strparentkey:="node10"
        .AddItem "Node 12" & vbTab & "Node 10" & vbTab & "Active", iconcollections:=CInt(Rnd() * 29), strkey:="node12", strparentkey:="node10"
        
        
        .LockUpdate = False
    End With

End Sub

Private Sub Form_Load()
    Me.OsenVistaForm1.Init Me
    Set cn = New SQLite3_Connection
    cn.OpenDB ".\nwind2008.db3"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cn.CloseDB
    Set cn = Nothing
End Sub

Private Sub OsenVistaButton1_Click()
    vList.AddItemsByRecordset cn.Recordset("select * from vw7 limit 0,500"), "1,2,0", autocolumnwidthex:=True
End Sub

Private Sub OsenVistaButton2_Click()
Dim L As Long
With vList
    .Clear True
    
    ' Add Columns
    .InsertColumn , "Account No." ' 0
    .InsertColumn , "Account Name" ' 1
    .InsertColumn , "Group" 'Here will be set as parent '2
    .InsertColumn , "Balance", , enAlignRight, 6, "#,##0" ' 3
     
     .LockUpdate = True 'Lock updating process
     
    .AddItem "1101.001" & vbTab & "Cash Rupiah" & vbTab & "Cash" & vbTab & "2460000", BaseTreeColumn:="2"
    .AddItem "1101.002" & vbTab & "Cash USDollar" & vbTab & "Cash" & vbTab & "17000000", BaseTreeColumn:="2"
    .AddItem "1101.003" & vbTab & "Cash SinDollar" & vbTab & "Cash" & vbTab & "13032225", BaseTreeColumn:="2"
    .AddItem "1101.201" & vbTab & "Petty Cash Rupiah" & vbTab & "Cash" & vbTab & "470000", BaseTreeColumn:="2"
    .AddItem "1101.202" & vbTab & "Petty Cash USD" & vbTab & "Cash" & vbTab & "1700000", BaseTreeColumn:="2"
    
    .AddItem "1102.001" & vbTab & "BALI USDollar" & vbTab & "Bank" & vbTab & "20691419", BaseTreeColumn:="2"
    .AddItem "1102.002" & vbTab & "Mandiri Rupiah" & vbTab & "Bank" & vbTab & "69250000", BaseTreeColumn:="2"
    .AddItem "1102.003" & vbTab & "BCA Rupiah" & vbTab & "Bank" & vbTab & "39821500", BaseTreeColumn:="2"
     
     .LockUpdate = False ' refresh/update list
    
End With

End Sub

