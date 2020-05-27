VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Query Analyzer [Press F5 to Execute]"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaStatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   5805
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   979
      BackColor       =   14936810
      ForeColor       =   16777215
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   2
      HaveXPForm      =   -1  'True
      PWidth1         =   320
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "0 row(s) [0 ms taken]"
      pTextAlignment1 =   0
      PanelPicture1   =   "Form2.frx":058A
      PanelPicAlignment1=   0
      PWidth2         =   1000
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Click Here to Execute"
      pTextAlignment2 =   0
      pTextBold2      =   -1  'True
      PanelPicture2   =   "Form2.frx":05A6
      PanelPicAlignment2=   0
      GradientColor1  =   10000535
      GradientColor2  =   5460819
   End
   Begin VistaSuitePro.OsenVistaListBox lstData 
      Height          =   2805
      Left            =   270
      TabIndex        =   2
      Top             =   2880
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4948
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
      BackSelected    =   7381139
      BackSelectedG1  =   16777215
      BackSelectedG2  =   8632490
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
      HeaderGradientTop=   8569007
      HeaderGradientBottom=   4487779
      BinaryImage     =   "Form2.frx":05C2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtSQL 
      Height          =   945
      Left            =   270
      TabIndex        =   1
      ToolTipText     =   "Press F5 to Execute"
      Top             =   1830
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   1667
      Text            =   "SELECT sqlite_version(),md5('osen kusnadi'),now();"
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
      ForeColor       =   0
      BorderColor     =   8370596
      MultiLine       =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox txtDBFile 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   1410
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   582
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
      ForeColor       =   0
      BorderColor     =   8370596
      ButtonCaption   =   ""
      ButtonPicture   =   "Form2.frx":05DA
      ButtonVisible   =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      LabelBackColor  =   15790320
      LabelCaption    =   "DB Filename:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Query Analyzer [Press F5 to Execute]"
      TitleTop        =   7
      icon            =   "Form2.frx":0974
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   4
      Top             =   420
      Width           =   7560
      _ExtentX        =   13335
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
      Picture         =   "Form2.frx":0F0E
      BorderColor     =   8632490
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
      BinaryImage     =   "Form2.frx":2E20
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenVistaSuite 2008 - SQLite3_Connection sample                      *'
'*  Copyright (c) 2008 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  [Form2.frm]                                                          *'
'*                                                                       *'
'*  This file is part of the OsenVistaSuite 2008 sample applications.    *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenVistaSuite 2008 documentation, and is provided "as is", without  *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

' Declare variable
Private SQLite3        As SQLite3_Connection

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrX
    ' Execute SQL [F5]
    If KeyCode = 116 Then
        If SQLite3.State Then
            Dim L As Long
            lstData.InsertItemByRecordset SQLite3.Recordset(txtSQL), LngTickTime:=L, AutoColumnWIdthEx:=True
            sBar.PanelCaption(1) = lstData.ListCount & " row(s)"
            sBar.ExtendedCaption 1, "[ " & L & " ms taken]", enAlignLeft, vbRed, True
        Else
            lstData.Message "There are no active database connection"
        End If
    End If
    Exit Sub
ErrX:
    MsgBoxGT SQLite3.ErrDescription, vbExclamation, "Query Analyzer"
On Error GoTo 0
End Sub

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize event)
    Me.OsenXPForm1.Init Me
    
    ' Now, you can make a new instance of SQLite3_Connection
    Set SQLite3 = New SQLite3_Connection

    ' Try to open sample database
    SQLite3.OpenDB App.Path & "\adbook.db3", "osenxpsuite"
    
    txtDBFile = App.Path & "\adbook.db3"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' CLean Up
    Set SQLite3 = Nothing
    
End Sub

Private Sub sBar_MouseDownInPanel(iPanel As Long)
On Error GoTo ErrX

    If iPanel = 2 Then
         If SQLite3.State Then
             Dim L As Long
             lstData.InsertItemByRecordset SQLite3.Recordset(txtSQL), LngTickTime:=L, AutoColumnWIdthEx:=True
             sBar.PanelCaption(1) = lstData.ListCount & " row(s)"
             sBar.ExtendedCaption 1, "[ " & L & " ms taken]", enAlignLeft, vbRed, True
         Else
             lstData.Message "There are no active database connection"
         End If
    End If
    
    Exit Sub
ErrX:
    MsgBoxGT SQLite3.ErrDescription, vbExclamation, "Query Analyzer"
On Error GoTo 0
End Sub

' Open Database connection
Private Sub txtDBFile_ButtonClick()

    ' Display a SHow open dialog and return the selected filename
    txtDBFile.ShowOpenDialog "Open SQLite3 Database", "SQLite3 Database|*.DB3;*.SDB;*.DB", "DB3"
    
    ' check the filename exists or not ...
    If Len(txtDBFile.Text) Then
        SQLite3.OpenDB txtDBFile
    Else
        SQLite3.CloseDB
    End If
    
End Sub























