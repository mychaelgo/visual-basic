VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_products 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Products"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   Icon            =   "frm_products.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyADODC MyADODC1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   14
      Top             =   4995
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   979
      GradientColor1  =   16777215
      GradientColor2  =   12752244
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
      Style           =   1
      BorderStyle     =   6
      Gradient        =   -1  'True
      ShowFindButton  =   0   'False
      ShowFilterButton=   0   'False
      ShowPrinterButton=   0   'False
      ShowGriper      =   0   'False
      AutoConfirmBeforeDelete=   -1  'True
      CaptionWidth    =   70
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaTab OsenXPTab1 
      Height          =   4455
      Left            =   150
      TabIndex        =   11
      Top             =   480
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   7858
      TabHeight       =   22
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FrameColor      =   12164479
      MaskColor       =   16711935
      SelectedTab     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   1
      BackColorParent =   16767935
      TabWidth1       =   76
      TabText1        =   "Product Info"
      TabEnabled1     =   -1  'True
      TabVisible1     =   0   'False
      Begin VistaSuitePro.OsenVistaCheckBox chkDiscount 
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   4020
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
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
         Caption         =   "Discontinued"
         Style           =   1
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel2 
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Top             =   1680
         Width           =   810
         _ExtentX        =   1429
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
         Caption         =   "Category:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaComboBox CboCtg 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   1680
         Width           =   3705
         _ExtentX        =   6535
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
         ComboStyle      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSH             =   -1  'True
         LSGL            =   -1  'True
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         HoverSelection  =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Required        =   -1  'True
         Unicode         =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   510
         Width           =   4905
         _ExtentX        =   8652
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
         Enabled         =   0   'False
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
         Required        =   -1  'True
         LabelCaption    =   "Product ID:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   900
         Width           =   4905
         _ExtentX        =   8652
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
         Required        =   -1  'True
         LabelCaption    =   "Product Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   1320
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
         Caption         =   "Supplier:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   4
         Left            =   180
         TabIndex        =   4
         Top             =   2070
         Width           =   4905
         _ExtentX        =   8652
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
         LabelCaption    =   "Qty per unit:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   5
         Left            =   180
         TabIndex        =   5
         Top             =   2460
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         Alignment       =   1
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
         NumberOnly      =   -1  'True
         UseFormat       =   -1  'True
         FormatNumber    =   "#,##0.00"
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
         CurrencySymbol  =   "US $"
         LabelCaption    =   "Unit price:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   6
         Left            =   180
         TabIndex        =   6
         Top             =   2850
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         Alignment       =   1
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
         LabelCaption    =   "Units in stock:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   3240
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         Alignment       =   1
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
         LabelCaption    =   "Units On Order:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   8
         Left            =   180
         TabIndex        =   8
         Top             =   3630
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         Alignment       =   1
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
         LabelCaption    =   "Reorder Level:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaComboBox cboSP 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   1290
         Width           =   3705
         _ExtentX        =   6535
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
         ComboStyle      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSH             =   -1  'True
         LSGL            =   -1  'True
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         HoverSelection  =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Required        =   -1  'True
         Unicode         =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
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
      Caption         =   "Products"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
End
Attribute VB_Name = "frm_products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.OsenXPForm1.Init Me
    
    cboSP.InsertItemByRecordset GetRST("select * from vlistSp"), True, True
    cboSP.TextColumn = 1 ' Display >> CompanyName
    cboSP.ColumnWidth(1) = 0 ' Hidden SupplierID
    cboSP.FitColumnWidth 2
    
    CboCtg.InsertItemByRecordset GetRST("select * from vlistctg"), True, 1
    CboCtg.TextColumn = 1 ' Display Category NAme
    CboCtg.ColumnWidth(1) = 0 ' Hidden CategoryID
    CboCtg.FitColumnWidth 2

    If IsNew Then
        MyADODC1.OpenMySQLTable "products", "productid", MyCN, "where productid=-1", txtdata ' open empty record
    Else
        MyADODC1.OpenMySQLTable "products", "productid", MyCN, "where productid=" & KeyValue, txtdata
    End If
    
    ' Bind ComboBox with MyADODC
    MyADODC1.Bind cboSP, 2 ' Suppliers
    MyADODC1.Bind CboCtg, 3 ' Categories
    
    ' Bind Checkbox
    MyADODC1.Bind chkDiscount, 9
    
    ' Prepare MyADODC action (Addnew or Update)
    If IsNew Then
        MyADODC1.SendAction 5 ' Addnew
    Else
        MyADODC1.SendAction 6 ' Update
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    frmMain.RefreshView
    frmMain.Show
End Sub




