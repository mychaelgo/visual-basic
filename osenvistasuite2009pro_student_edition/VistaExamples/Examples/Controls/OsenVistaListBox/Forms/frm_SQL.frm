VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm_SQL 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Customize ListBox by Query"
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_SQL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   510
      Top             =   5580
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   37884
      Images          =   "frm_SQL.frx":038A
      Version         =   720920
      KeyCount        =   33
      Keys            =   "????????????????????????????????ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaButton CmdRefresh 
      Height          =   345
      Left            =   8340
      TabIndex        =   4
      Top             =   3930
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BCOL            =   16777215
      BCOLO           =   16777215
      Caption         =   "&Resfresh"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_SQL.frx":97A6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.OsenVistaButton cmdFilter 
      Height          =   345
      Left            =   8340
      TabIndex        =   3
      Top             =   3510
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BCOL            =   16777215
      BCOLO           =   16777215
      Caption         =   "&Filter"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_SQL.frx":97C2
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.OsenVistaButton cmdSearch 
      Height          =   345
      Left            =   8340
      TabIndex        =   2
      Top             =   3090
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BCOL            =   16777215
      BCOLO           =   16777215
      Caption         =   "&Search"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_SQL.frx":97DE
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.OsenVistaListBox LField 
      Height          =   1545
      Left            =   1650
      TabIndex        =   22
      Top             =   4980
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   2725
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      ShowHeader      =   -1  'True
      HeaderFormatString=   "No;40;2|FieldName;120;0"
      Columns         =   2
      ShowGridLines   =   -1  'True
      MaxAllColumnWidth=   160
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtFields 
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   8
      Top             =   4920
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      Text            =   "8"
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
      Value           =   8
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaButton CmdExecute 
      Height          =   405
      Left            =   180
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      BCOL            =   16777215
      BCOLO           =   16777215
      Caption         =   "&Execute"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm_SQL.frx":97FA
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtSQL 
      Height          =   1665
      Left            =   180
      TabIndex        =   1
      Top             =   3060
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   2937
      Text            =   $"frm_SQL.frx":9816
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaListBox LView 
      Height          =   2445
      Left            =   180
      TabIndex        =   6
      Top             =   570
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4313
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      ShowHeader      =   -1  'True
      ShowGridLines   =   -1  'True
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   "MyImageList1"
      ReadOnDemand    =   -1  'True
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Customize ListBox by Query"
      icon            =   "frm_SQL.frx":991E
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      MaximizeEnabled =   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtFields 
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   11
      Top             =   5370
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      Text            =   "-1"
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
      Value           =   -1
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtFields 
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   14
      Top             =   5820
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      Text            =   "7"
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
      Value           =   7
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtFields 
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   17
      Top             =   6270
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   661
      Text            =   "6"
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
      Value           =   6
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can set the field for set the color or bold for List Item, and it's easy of use. please learn it from sample"
      Height          =   195
      Left            =   270
      TabIndex        =   21
      Top             =   7080
      Width           =   7650
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "List item color or bold can be set by value of field in recordset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   20
      Top             =   6750
      Width           =   6255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REMARK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   19
      Top             =   6390
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "fieldname = bold, index no.6 frm SQL above"
      Height          =   615
      Index           =   3
      Left            =   7200
      TabIndex        =   18
      Top             =   6270
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field for Custom Bold:"
      Height          =   195
      Index           =   3
      Left            =   4560
      TabIndex        =   16
      Top             =   6330
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fieldname = color (see SQL above)"
      Height          =   195
      Index           =   2
      Left            =   7170
      TabIndex        =   15
      Top             =   5880
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field for Custom Color:"
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   13
      Top             =   5880
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Field Index from recordset)"
      Height          =   195
      Index           =   1
      Left            =   7170
      TabIndex        =   12
      Top             =   5430
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field for Selected Icon:"
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   10
      Top             =   5430
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Field Index from recordset)"
      Height          =   195
      Index           =   0
      Left            =   7170
      TabIndex        =   9
      Top             =   4980
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field for Normal Icon:"
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   7
      Top             =   4980
      Width           =   1545
   End
End
Attribute VB_Name = "frm_SQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenXPSuite 2006 - OsenXPListBox sample                              *'
'*  Copyright (c) 2006 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  This file is part of the OsenXPSuite 2006 sample applications.       *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenXPSuite 2006 documentation, and is provided "as is", without     *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

Private Sub Form_Load()

    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
End Sub

Private Sub CmdExecute_Click()

    
    ' Just writes down a single line of code , you can present data format as according to condition of at query which in executing
    LView.InsertItemBySQL AdoCN, TxtSQL, , , True, , TxtFields(0).Value, TxtFields(1).Value, TxtFields(2).Value, TxtFields(3).Value
    
    ' Now, customize a specified column alignment
    LView.ColumnStyle(6) = 1 ' As Checkbox
    LView.ColumnAlignment(5) = enAlignRight
    LView.ColumnAlignment(4) = enAlignRight
    LView.ColumnFormat(4) = "[$]#,##0.00" ' Format String for Currency
    
    ' Retrieve FieldName from SQL
    Dim I As Integer
    With LField
        .Clear
        If LView.Columns Then
            For I = 1 To LView.Columns
                .AddItem I - 1 & vbTab & LView.ColumnText(I)
            Next
        End If
    End With
    
End Sub

Private Sub cmdFilter_Click()
    LView.List_Filter
End Sub

Private Sub CmdRefresh_Click()
    LView.List_Refresh
End Sub

Private Sub cmdSearch_Click()
    LView.List_Search
End Sub

 
 
 
 
 
 
 
 
 
 
 
 
 
 








