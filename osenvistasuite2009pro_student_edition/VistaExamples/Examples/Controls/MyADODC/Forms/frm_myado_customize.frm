VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_customize 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Customers"
   ClientHeight    =   5025
   ClientLeft      =   3750
   ClientTop       =   1530
   ClientWidth     =   7275
   Icon            =   "frm_myado_customize.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   6420
      Top             =   1290
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   14924
      Images          =   "frm_myado_customize.frx":058A
      Version         =   131072
      KeyCount        =   13
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   1005
      Left            =   0
      TabIndex        =   24
      Top             =   420
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1773
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
      Picture         =   "frm_myado_customize.frx":3FF6
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "This form contains all information about the customer"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Customer records"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "frm_myado_customize.frx":5B48
      WindowColor     =   0
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   10
      Left            =   4380
      TabIndex        =   10
      Top             =   3990
      Width           =   2625
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   9
      Left            =   1530
      TabIndex        =   9
      Top             =   3990
      Width           =   1995
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   8
      Left            =   4380
      TabIndex        =   8
      Top             =   3600
      Width           =   2625
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   7
      Left            =   1530
      TabIndex        =   7
      Top             =   3600
      Width           =   1995
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   6
      Left            =   4380
      TabIndex        =   6
      Top             =   3210
      Width           =   2625
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   1530
      TabIndex        =   5
      Top             =   3210
      Width           =   1995
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   1530
      TabIndex        =   4
      Top             =   2820
      Width           =   5475
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   4950
      TabIndex        =   3
      Top             =   2430
      Width           =   2055
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   1530
      TabIndex        =   2
      Top             =   2430
      Width           =   2055
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   1530
      TabIndex        =   1
      Top             =   2040
      Width           =   5475
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   1650
      Width           =   1515
   End
   Begin VistaSuitePro.MyADODC MyADODC1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   12
      Top             =   4470
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   979
      GradientColor1  =   10000535
      GradientColor2  =   5460819
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
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
      Alignment       =   2
      Style           =   1
      BorderStyle     =   3
      Gradient        =   -1  'True
      ListCaption     =   "Index Konsumen|Nama Perusahaan|Kontak|-|-|Kota|-|-|Negara"
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
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
      Caption         =   "Customers"
      TitleTop        =   7
      icon            =   "frm_myado_customize.frx":5B60
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      BorderStyle     =   1
      AllowFadeIn     =   -1  'True
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      Height          =   195
      Index           =   10
      Left            =   3660
      TabIndex        =   23
      Top             =   4050
      Width           =   300
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      Height          =   195
      Index           =   9
      Left            =   300
      TabIndex        =   22
      Top             =   4050
      Width           =   510
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      Height          =   195
      Index           =   8
      Left            =   3660
      TabIndex        =   21
      Top             =   3660
      Width           =   585
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Code:"
      Height          =   195
      Index           =   7
      Left            =   300
      TabIndex        =   20
      Top             =   3660
      Width           =   900
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region:"
      Height          =   195
      Index           =   6
      Left            =   3660
      TabIndex        =   19
      Top             =   3270
      Width           =   555
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   195
      Index           =   5
      Left            =   300
      TabIndex        =   18
      Top             =   3300
      Width           =   300
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   17
      Top             =   2910
      Width           =   615
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Title:"
      Height          =   195
      Index           =   3
      Left            =   3870
      TabIndex        =   16
      Top             =   2490
      Width           =   945
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name:"
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   15
      Top             =   2490
      Width           =   1065
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   14
      Top             =   2100
      Width           =   1170
   End
   Begin VB.Label LbInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID:"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   13
      Top             =   1710
      Width           =   870
   End
End
Attribute VB_Name = "frm_customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenVistaSuite 2008 - MyADODC Control sample                         *'
'*  Copyright (c) 2008 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  [Form1.frm]                                                          *'
'*                                                                       *'
'*  This file is part of the OsenVistaSuite 2008 sample applications.    *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenVistaSuite 2008 documentation, and is provided "as is", without  *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

Private Sub Form_Load()
On Error GoTo Err_On_Load

    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
    ' Retrieve all customers information (Open recordset --> customers table)
    ' Proc: OpenRecordset (SQL,Conn,TxtField,...)
    MyADODC1.OpenRecordset "select * from customers", ADOCN, TxtData
    
    '/// Change the Image Button Of MyADODC
    Dim I As Integer
    For I = 1 To 13
        MyADODC1.ChangeImageButton MyImageList1.ItemPicture(I), I
    Next
    
    '// Change Tooltip of MyADODC
    With MyADODC1
        .ChangeToolTipButton "Awal", 1
        .ChangeToolTipButton "Mundur", 2
        .ChangeToolTipButton "Maju", 3
        .ChangeToolTipButton "Akhir", 4
        .ChangeToolTipButton "Tambah", 5
        .ChangeToolTipButton "Ubah", 6
        .ChangeToolTipButton "Hapus", 7
        .ChangeToolTipButton "Simpan", 8
        .ChangeToolTipButton "Batal", 9
        .ChangeToolTipButton "Cari Data", 10
        .ChangeToolTipButton "Saring Data", 11
        .ChangeToolTipButton "Refresh Data", 12
        .ChangeToolTipButton "Cetak Laporan", 13
    End With
    
    
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
    
    
    Exit Sub
Err_On_Load:
    MsgBoxXP Err.Description, vbCritical, "Error", , 5, , 1
    Err.Clear
    Unload Me
    
End Sub

Private Sub MyADODC1_ButtonClick(ByVal ButtonName As VistaSuitePro.EnumButtonName, Cancel As Boolean, Is_MySQL_RS As Boolean)
On Error GoTo Err_Main
    Select Case ButtonName
    
        Case 7 'buttonname=[Delete Record]
        
            Dim data
            
            ' set button caption
            
            data = Array("&Ya", "&Tidak")
            ' vbDefaultButton2 -> "&Tidak"
            
            Dim tFont As New StdFont
            
            With tFont
                .Name = "Arial"
                .Size = 10
                .Bold = True
            End With
            
            Dim bFont As New StdFont
            
            With bFont
                .Name = "Comic Sans MS"
                .Size = 10
                .Bold = False
            End With
            
            If MsgBoxGT("Anda yakin akan menghapus data ini?", vbQuestion + 1024 + vbDefaultButton2, "Hapus data", , , data) = 1 Then     ' if user select &Tidak=1
                Cancel = True
            End If
        
        Case 13 'buttonname=[Print Preview]
        
            ' handle report viewer here
            MsgBoxXP "The customers report does not exist.", vbExclamation, "Print"
            
    End Select
Exit Sub
Err_Main:
    MsgBoxGT Err.Description, vbCritical, "Error"
    Err.Clear
End Sub


Private Sub TxtData_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub























