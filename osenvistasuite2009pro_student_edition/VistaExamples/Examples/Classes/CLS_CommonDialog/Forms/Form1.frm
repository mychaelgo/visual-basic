VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "CLS_CommonDialog Sample"
   ClientHeight    =   3360
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton cmdColor 
      Height          =   405
      Left            =   2820
      TabIndex        =   5
      Top             =   2100
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   714
      Caption         =   "Show &Color Dialog"
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
   End
   Begin VistaSuitePro.OsenVistaButton cmdSave 
      Height          =   405
      Left            =   300
      TabIndex        =   4
      Top             =   2100
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   714
      Caption         =   "Show &Save Dialog"
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
      MICON           =   "Form1.frx":0940
      PICN            =   "Form1.frx":095C
      UMCOL           =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaButton cmdOpen 
      Height          =   405
      Left            =   300
      TabIndex        =   3
      Top             =   1500
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   714
      Caption         =   "Show &Open Dialog"
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
      MICON           =   "Form1.frx":0CF6
      PICN            =   "Form1.frx":0D12
      UMCOL           =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaTextBox txtFile 
      Height          =   375
      Left            =   300
      TabIndex        =   2
      Top             =   2730
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   661
      Text            =   "OsenVistaSuite 2008 Express Edition"
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
      ButtonGradient  =   -1  'True
      ForeColorOver   =   33023
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
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   5280
      _ExtentX        =   9313
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
      Picture         =   "Form1.frx":12AC
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "How to use the CLS_CommonDialog class object ?"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Class Name: CLS_CommonDialog"
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
      BinaryImage     =   "Form1.frx":1EFE
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5280
      _ExtentX        =   9313
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
      Caption         =   "CLS_CommonDialog Sample"
      TitleTop        =   7
      icon            =   "Form1.frx":1F16
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin VistaSuitePro.OsenVistaButton cmdFont 
      Height          =   405
      Left            =   2850
      TabIndex        =   6
      Top             =   1500
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   714
      Caption         =   "Show &Font Dialog"
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
      MICON           =   "Form1.frx":22B0
      PICN            =   "Form1.frx":22CC
      UMCOL           =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************'
'*  OsenVistaSuite 2008 - CLS_CommonDialog sample                           *'
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


Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    
    Me.OsenXPForm1.Init Me
     
End Sub

Private Sub cmdFont_Click()

    Dim cDLG As New CLS_CommonDialog
    
    ' set the parent object of cls_commondialog by Hwnd property
    cDLG.hWnd = Me.hWnd ' Make form1 (this form) as the parent for cls_commondialog
    
    ' Initialize font style
    cDLG.Font.Name = txtFile.Font.Name
    cDLG.Font.Size = txtFile.Font.Size
    cDLG.Font.Italic = txtFile.Font.Italic
    cDLG.Font.Charset = txtFile.Font.Charset
    cDLG.Font.Bold = txtFile.Font.Bold
    cDLG.Font.Strikethrough = txtFile.Font.Strikethrough
    cDLG.Font.Underline = txtFile.Font.Underline
    
    ' Show the Font dialog
    cDLG.ShowFont
    
    ' Get the result
    Set txtFile.Font = cDLG.Font
    
    ' clean up
    Set cDLG = Nothing
    
End Sub

Private Sub cmdOpen_Click()

    Dim cDLG As New CLS_CommonDialog
    
    ' set the parent object of cls_commondialog by Hwnd property
    cDLG.hWnd = Me.hWnd ' Make form1 (this form) as the parent for cls_commondialog
    
    ' clear current filename
    cDLG.FileName = ""
    
    ' Make a filter
    cDLG.Filter = "Microsoft Access|*mdb|Microsoft Excel|*.xls|All Files|*.*"
    
    ' Show the open dialog
    cDLG.ShowOpen
    
    ' Get a result
    txtFile.Text = cDLG.FileName
    
    ' clean up
    Set cDLG = Nothing
    
End Sub

Private Sub cmdColor_Click()

    Dim cDLG As New CLS_CommonDialog
    
    ' set the parent object of cls_commondialog by Hwnd property
    cDLG.hWnd = Me.hWnd ' Make form1 (this form) as the parent for cls_commondialog
    
    ' Show the printer dialog
    cDLG.ShowColor
    
    ' Get a result
    txtFile.Text = "Get a color from cls_commondialog"
    txtFile.ForeColor = cDLG.Color
    
    Set cDLG = Nothing
    
End Sub

Private Sub cmdSave_Click()

    Dim cDLG As New CLS_CommonDialog
    
    ' set the parent object of cls_commondialog by Hwnd property
    cDLG.hWnd = Me.hWnd ' Make form1 (this form) as the parent for cls_commondialog
    
    ' clear current filename
    cDLG.FileName = ""
    
    ' Make a filter
    cDLG.Filter = "Microsoft Access|*mdb|Microsoft Excel|*.xls|All Files|*.*"
    
    ' Make a default file type (which is specified in default extention)
    cDLG.DefaultExt = "MDB"
    
    ' Show the save dialog
    cDLG.ShowSave
    
    ' Get a result
    txtFile.Text = cDLG.FileName
    
    ' clean up
    Set cDLG = Nothing
    
End Sub






















