VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form frm_csf 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Customize a Search/Filter dialog at OsenXPListBox"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   Icon            =   "frm_customize_search_filter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaButton OsenXPButton4 
      Height          =   345
      Left            =   3840
      TabIndex        =   7
      Top             =   6030
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      Caption         =   "&Show Report"
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
      MICON           =   "frm_customize_search_filter.frx":058A
      PICN            =   "frm_customize_search_filter.frx":06EC
      UMCOL           =   -1  'True
      BinaryImageNormal=   "frm_customize_search_filter.frx":0A86
      BinaryImageOver =   "frm_customize_search_filter.frx":0A9E
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton3 
      Height          =   345
      Left            =   2550
      TabIndex        =   6
      Top             =   6030
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Caption         =   "&Refresh"
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
      MICON           =   "frm_customize_search_filter.frx":0AB6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_customize_search_filter.frx":0C18
      BinaryImageOver =   "frm_customize_search_filter.frx":0C30
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   345
      Left            =   1320
      TabIndex        =   5
      Top             =   6030
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      MCOL            =   16711935
      MPTR            =   99
      MICON           =   "frm_customize_search_filter.frx":0C48
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_customize_search_filter.frx":0DAA
      BinaryImageOver =   "frm_customize_search_filter.frx":0DC2
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   6030
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      MCOL            =   16711935
      MPTR            =   99
      MICON           =   "frm_customize_search_filter.frx":0DDA
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_customize_search_filter.frx":0F3C
      BinaryImageOver =   "frm_customize_search_filter.frx":0F54
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
      BinaryImage     =   "frm_customize_search_filter.frx":0F6C
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
      Caption         =   "Customize a Search/Filter dialog at OsenXPListBox"
      TitleTop        =   7
      icon            =   "frm_customize_search_filter.frx":0F84
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
      Top             =   6060
      Width           =   960
   End
End
Attribute VB_Name = "frm_csf"
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
    
    ' Customize language of search/Filter dialog
    Dim SF As New CLS_SFDialog
    With SF
        
        ' Searching dialog
        .SearchButtonCaption = "&Cari"
        .SearchConditionCaption = "Kondisi pencarian"
        .SearchDescription = "Keterangan untuk pencarian data boleh ditulis disini"
        .SearchLookByCaption = "Cari data berdasarkan"
        .SearchLookForCaption = "Cari data untuk"
        .SearchTitleBar = "Pencarian data"
        
        ' Filter dialog
        .FilterButtonCaption = "&Saring"
        .FilterConditionCaption = "Kondisi penyaringan"
        .FilterDescription = "Keterangan untuk penyaringan data boleh ditulis disini juga :)"
        .FilterLookByCaption = "saring data dengan"
        .FilterLookForCaption = "saring data untuk"
        .FilterTitleBar = "Penyaringan data"
        
        ' Miscellaneous
        .BeginWithCaption = "Di mulai dengan"
        .ContainWithCaption = "Terisi oleh"
        .CancelButtonCaption = "&Batal"
        
        ' Font
        .TitlebarFont.Name = "Tahoma"
        .TitlebarFont.Size = 11
        .TitlebarFont.Bold = True
        
        .DefaultFont.Name = "Comic Sans MS"
        .DefaultFont.Size = 9
        .DefaultFont.Bold = False
    End With
    
End Sub

Private Sub CboScheme_Click()

    ' Change the colorscheme of OsenXPForm
'    Me.OsenXPForm1.ColorScheme = CboScheme.ListIndex
    
End Sub

Private Sub OsenXPButton1_Click()
    LstX.List_Search
End Sub

Private Sub OsenXPButton2_Click()
    LstX.List_Filter
End Sub

Private Sub OsenXPButton3_Click()
    LstX.List_Refresh
End Sub

Private Sub OsenXPButton4_Click()
On Error Resume Next

    
    LstX.LockRecordset = False
    
    ' Using this method to display/preview activereport
    ' Please make sure that you have activereport from datadynamics was installed in your PC
'    DisplayARV LstX.ActiveRst, "Customer Information", App.Path & "\arv\customers.rpx"
        
    LstX.LockRecordset = True
    
End Sub
































