VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form frm_exd 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Export Data"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaButton OsenXPButton4 
      Height          =   345
      Left            =   4290
      TabIndex        =   7
      Top             =   6030
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "Export To XML"
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
      MPTR            =   99
      MICON           =   "Form1.frx":058A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "Form1.frx":06EC
      BinaryImageOver =   "Form1.frx":0704
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton3 
      Height          =   345
      Left            =   2910
      TabIndex        =   6
      Top             =   6030
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      Caption         =   "Export to HTML"
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
      MPTR            =   99
      MICON           =   "Form1.frx":071C
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "Form1.frx":087E
      BinaryImageOver =   "Form1.frx":0896
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   345
      Left            =   1500
      TabIndex        =   5
      Top             =   6030
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      Caption         =   "Export To Csv"
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
      MPTR            =   99
      MICON           =   "Form1.frx":08AE
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "Form1.frx":0A10
      BinaryImageOver =   "Form1.frx":0A28
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   6030
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      Caption         =   "Export to Excel"
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
      MPTR            =   99
      MICON           =   "Form1.frx":0A40
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "Form1.frx":0BA2
      BinaryImageOver =   "Form1.frx":0BBA
   End
   Begin VistaSuitePro.OsenVistaComboBox CboScheme 
      Height          =   315
      Left            =   7170
      TabIndex        =   2
      Top             =   6030
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
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
      Text            =   "ComboBox1"
      ComboStyle      =   1
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSGL            =   -1  'True
      LLID            =   2
      LIO             =   2
      LITL            =   2
      IMGLIST         =   ""
      HoverSelection  =   -1  'True
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   0
      ASURC           =   0   'False
      TextColumn      =   0
      Required        =   0   'False
      Unicode         =   0   'False
      DataList        =   "Blue|Olive Green|Silver"
      BorderColor     =   12164479
      BorderColorOver =   12164479
   End
   Begin VistaSuitePro.OsenVistaListBox LstX 
      Height          =   5445
      Left            =   90
      TabIndex        =   1
      Top             =   510
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   9604
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
      SelectModeStyle =   2
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
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "Form1.frx":0BD2
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
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
      Caption         =   "Export Data"
      TitleTop        =   7
      icon            =   "Form1.frx":0BEA
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colorscheme:"
      Height          =   195
      Left            =   6030
      TabIndex        =   3
      Top             =   6090
      Width           =   960
   End
End
Attribute VB_Name = "frm_exd"
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

    ' Get Customers Information, and display it at LstX
    LstX.InsertItemByRecordset GetADORecordset("select * from customers"), , , True
    
End Sub

Private Sub CboScheme_Click()

    ' Change the colorscheme of OsenXPForm
'    Me.OsenXPForm1.ColorScheme = CboScheme.ListIndex
    
End Sub

Private Sub OsenXPButton1_Click()

    LstX.ExportToExcel
    
End Sub

Private Sub OsenXPButton2_Click()

    ' prepare filename for output data (destination)
    mStrSQL = "c:\test_customers.csv"
    
    ' Here is the method for export data from osenxplistbox into csv file format
    LstX.ExportToCSV mStrSQL
    
    ' Display message
    MsgBoxGT "Export data finished." & vbCrLf & "Filename: " & mStrSQL, vbInformation
    
End Sub

Private Sub OsenXPButton3_Click()

    ' prepare filename for output data (destination)
    mStrSQL = "c:\test_customers.html"
    
    ' Here is the method for export data from osenxplistbox into html file format
    LstX.ExportToHTML mStrSQL, "Customers Information"
    
    ' Display message
    MsgBoxGT "Export data finished." & vbCrLf & "Filename: " & mStrSQL, vbInformation
    
    ' Now, Display that HTML file
    OpenBrowser 0, mStrSQL
    ' OpenBrowser is global function from XP Library (OsenXPSuite2006.XP class object)

End Sub

Private Sub OsenXPButton4_Click()

    ' prepare filename for output data (destination)
    mStrSQL = "c:\test_customers.xml"
    
    ' Here is the method for export data from osenxplistbox into xml file format
    LstX.ExportToXML mStrSQL
    
    ' Display message
    MsgBoxGT "Export data finished." & vbCrLf & "Filename: " & mStrSQL, vbInformation
    
End Sub






















