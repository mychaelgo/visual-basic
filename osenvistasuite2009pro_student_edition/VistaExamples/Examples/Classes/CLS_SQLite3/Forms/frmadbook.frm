VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form frm_main 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   0  'None
   Caption         =   "AddressBook XP with SQLite3 engine"
   ClientHeight    =   5115
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmadbook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyADODC MySQLite 
      Height          =   555
      Left            =   300
      TabIndex        =   19
      Top             =   4380
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   979
      GradientColor1  =   10000535
      GradientColor2  =   5460819
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BeginProperty FontButton {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "MyADODC1"
      Style           =   1
      BorderStyle     =   6
      Gradient        =   -1  'True
      AutoConfirmBeforeDelete=   -1  'True
   End
   Begin VistaSuitePro.OsenVistaPicture pic 
      Height          =   2985
      Left            =   5910
      TabIndex        =   18
      Top             =   1290
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   5265
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
      FieldName       =   "photo"
      BorderColor     =   12563634
      GradientColor2  =   12563634
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BorderStyle     =   1
      BinaryImage     =   "frmadbook.frx":058A
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   17
      Top             =   450
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   1402
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
      Picture         =   "frmadbook.frx":05A2
      BorderColor     =   12563634
      PictureAlignment=   6
      GradientBackGround=   -1  'True
      GradientColor2  =   12563634
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Space           =   15
      Description     =   "<Description ... xxxxxxxxxx xxxxxxxx xxxxxxxxxxxx>"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "AddressBook with SQLite3 database"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextLeft        =   25
      DescriptionLeft =   45
      BinaryImage     =   "frmadbook.frx":20F4
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   315
      Index           =   0
      Left            =   1230
      TabIndex        =   9
      Top             =   1380
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Locked          =   -1  'True
      BackColor       =   12648447
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOver   =   255
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1410
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Number:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
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
      Caption         =   "AddressBook XP with SQLite3 engine"
      TitleTop        =   7
      icon            =   "frmadbook.frx":210C
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Name:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2220
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Address:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3060
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "City:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Country:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   5
      Left            =   3180
      TabIndex        =   6
      Top             =   3120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Postal Code:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   6
      Left            =   3180
      TabIndex        =   7
      Top             =   3510
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Phone:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   3990
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "E-mail:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   315
      Index           =   1
      Left            =   1230
      TabIndex        =   10
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   705
      Index           =   2
      Left            =   1230
      TabIndex        =   11
      Top             =   2220
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1244
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
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
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   315
      Index           =   3
      Left            =   1230
      TabIndex        =   12
      Top             =   3060
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   315
      Index           =   4
      Left            =   4320
      TabIndex        =   13
      Top             =   3090
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   315
      Index           =   5
      Left            =   1230
      TabIndex        =   14
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   315
      Index           =   6
      Left            =   4320
      TabIndex        =   15
      Top             =   3480
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   315
      Index           =   7
      Left            =   1230
      TabIndex        =   16
      Top             =   3930
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoTab         =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3120
      Picture         =   "frmadbook.frx":26A6
      Top             =   1380
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

' Very simple and easy in building an embedded database application with osenxpsuite
' [BLOB/Image has supported]
Private Sub Form_Load()

    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize event)
    Me.OsenXPForm1.Init Me

    'Open SQLite3 database
    MySQLite.OpenSQLite3Database App.Path & "\adbook.db3", "osenxpsuite"

    'Open SQLite3 table ..., and then binding that returned recordset with txtdata (osenxptextbox) and pic (osenxppicture)
    MySQLite.OpenSQLite3Table "adbook", , , txtData, pic
    
    ' Add new separator onto MyADODC
    MySQLite.AddSeparator
    
    ' Add new button ...
    MySQLite.AddButton 14, Image1.Picture, "Backup database"
    
End Sub


Private Sub MySQLite_ButtonClick(ByVal ButtonName As VistaSuitePro.EnumButtonName, Cancel As Boolean, Is_MySQL_RS As Boolean)
    
    
    If ButtonName = 14 Then ' Backup database
    
        ' Backup database ???
        If MsgBoxXP("Do you really want to backup your database?", vbQuestion + vbYesNo, "Backup Database?", , xpSilver, , , True) = vbYes Then
        
            Dim o As New CLS_CommonDialog
            Dim L As Long
            Dim sInfo As String
            
                ' set ...
                o.hWnd = Me.hWnd
                o.DialogTitle = "Backup database"
                o.Filter = "SQLite3 database format|*.db3;*.db;*.sdb"
                o.DefaultExt = "db3"
                
                ' display save dialog
                o.ShowSave
                
                If LenB(o.FileName) Then
                    
                    ' return Gettickcount Value
                    L = GTick
                    ' backup database now ...
'                    MySQLite.SQLiteConn.CopyDB o.FileName
                    L = GTick - L
                    
                    ' Prepare report ...
                    sInfo = "Backup database successfull" & vbCrLf & _
                    L & " ms taken" & vbCrLf & "Filename: " & o.FileName & vbCrLf & _
                    "Filesize: " & Format$(FileLen(o.FileName) \ 1024, "#,##0 Kb")
                    
                    ' Display message
                    MsgBoxXP sInfo, vbInformation, "Backup Finished", , xpSilver, , , True
                    
                End If
            
            Set o = Nothing
            
        End If
        
    End If
    
End Sub






















