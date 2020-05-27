VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "SQLite3_Connection Sample"
   ClientHeight    =   7500
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   7695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton cmdMyADODC 
      Height          =   375
      Left            =   2460
      TabIndex        =   6
      Top             =   6870
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   661
      Caption         =   "SQL&ite3 in MyADODC"
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
      MICON           =   "Form1.frx":038A
      PICN            =   "Form1.frx":03A6
      UMCOL           =   -1  'True
      BinaryImageNormal=   "Form1.frx":0740
      BinaryImageOver =   "Form1.frx":0758
   End
   Begin VistaSuitePro.OsenVistaButton cmdQuery 
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   6870
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "&Simple Query Analyzer..."
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
      MICON           =   "Form1.frx":0770
      PICN            =   "Form1.frx":078C
      UMCOL           =   -1  'True
      BinaryImageNormal=   "Form1.frx":0D26
      BinaryImageOver =   "Form1.frx":0D3E
   End
   Begin VistaSuitePro.OsenVistaButton cmdTest 
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   6900
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Running Test..."
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
      MICON           =   "Form1.frx":0D56
      PICN            =   "Form1.frx":0D72
      UMCOL           =   -1  'True
      BinaryImageNormal=   "Form1.frx":110C
      BinaryImageOver =   "Form1.frx":1124
   End
   Begin VistaSuitePro.OsenVistaListBox lstDemo 
      Height          =   3105
      Left            =   210
      TabIndex        =   3
      Top             =   3660
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   5477
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
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
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
      HeaderCaption   =   "OsenXPListBox1"
      TransparencyLevel=   22
      ReadOnDemand    =   -1  'True
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BinaryImage     =   "Form1.frx":113C
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1561
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":1154
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "The Following is example of usage SQLite3_Connection"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Class Name: SQLite3_Connection"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "Form1.frx":3066
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SQLite3_Connection Sample"
      TitleTop        =   7
      icon            =   "Form1.frx":307E
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin VB.Label LbInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2265
      Left            =   210
      TabIndex        =   2
      Top             =   1380
      Width           =   7275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************'
'*  OsenVistaSuite 2008 - SQLite3_Connection sample                      *'
'*  Copyright (c) 2008 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  [Form1.frm]                                                          *'
'*                                                                       *'
'*  This file is part of the OsenVistaSuite 2008 sample applications.    *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenVistaSuite 2008 documentation, and is provided "as is", without  *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

Option Explicit

' Declare variable
Private SQLite3        As SQLite3_Connection


Private Sub cmdMyADODC_Click()
    frm_main.Show 1
End Sub

Private Sub cmdQuery_Click()
    Form2.Show 1
End Sub

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize event)
    Me.OsenXPForm1.Init Me
    
    ' Now, you can make a new instance of SQLite3_Connection
    Set SQLite3 = New SQLite3_Connection
    
    ' Open Memory Database
    SQLite3.OpenMemorDB

End Sub


Private Sub cmdTest_Click()
On Error Resume Next

Dim L As Long
Dim n As Long
 
    ' Return the SQLite3 version
    LbInfo.Caption = "SQLite version: " & SQLite3.ExecScalar("select sqlite_version();")
      
    ' As same as  Gettickcount function
    L = GTick()
    
    ' Create a demo table
    SQLite3.Execute "create table if not exists demo2(nomor integer primary key,name text,address text,email text,website text)"
    
    ' Calculated the duration table creation process time
    L = GTick - L
    
    ' Display message
    LbInfo.Caption = LbInfo.Caption & vbLf & vbLf & "Create table: " & L & " ms" & vbCrLf
    
    DoEvents
    
    ' GetTickCount() ...
    L = GTick
    
    ' Start Transaction
    SQLite3.BeginTrans
    
    For n = 1 To 50000
        ' Execute a query
        SQLite3.Execute "insert into demo2 values(null,'osen','bekasi','support@osenxpsuite.net','http://osenxpsuite.net')"
    Next
    
    ' commit trasaction
    SQLite3.CommitTrans
    
    ' Calculated how many time which in applies to process data inclusion
    L = GTick - L
    
    ' Display a report
    LbInfo.Caption = LbInfo.Caption & vbLf & "Insert 50,000 records on trasaction: " & L & " ms" & vbCrLf
    DoEvents
    
    ' Insert data using prepare method
    
    ' GetTickCount()
    L = GTick
    
    SQLite3.BeginTrans  ' Begin transaction
    
    ' Prepare a query for inserting record ...
    SQLite3.Prepare "insert into demo2 values(?,?,?,?,?)"
    
    
    For n = 1 To 50000
    
        ' Binding new value into current statement
        SQLite3.BindValue 2, "osen"
        SQLite3.BindValue 3, "bekasi"
        SQLite3.BindValue 4, "support@osenxpsuite.net"
        SQLite3.BindValue 5, "http://osenxpsuite.net"
        
        ' Execute
        SQLite3.ExecuteNonQuery
        
    Next
    
    SQLite3.CommitTrans  ' commit trasaction
    
    ' Calculated how many time which in applies to process data inclusion (Using prepare statement)
    L = GTick - L
    
    ' Display a report...
    LbInfo.Caption = LbInfo.Caption & vbLf & "Insert 50,000 records on trasaction (Use prepare method): " & L & " ms" & vbCrLf
    DoEvents
    
    ' GetTick() ...
    L = GTick
    
    ' Execute select statement and convert the resultset into adodb.recordset (using Get_ADORS function)
    lstDemo.InsertItemByRecordset SQLite3.Recordset("select * from demo2 limit 0,20000 "), AutoColumnWIdthEx:=True
    
    ' Calculated how many time which in applies to process selection and convertion data (Select,convert)
    L = GTick - L
    
    ' Display a report
    LbInfo.Caption = LbInfo.Caption & vbLf & "Select 20,000 records and convert it into ADODB.recordset: " & L & " ms" & vbCrLf
    DoEvents
    
    L = GTick
    
    n = SQLite3.ExecScalar("select count(*) from demo2")
    L = GTick - L
    
    LbInfo.Caption = LbInfo.Caption & vbLf & "SELECT count(*) FROM demo2' = " & Format(n, "#,##0") & " [" & L & " ms taken]"
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' CLean Up
    Set SQLite3 = Nothing
    
End Sub




