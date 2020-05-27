VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_AddItem 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Insert item"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   Icon            =   "frm_AddItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4350
      _ExtentX        =   7673
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
      Caption         =   "Insert item"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
   End
   Begin VistaSuitePro.OsenVistaButton cmdOK 
      Height          =   345
      Left            =   3000
      TabIndex        =   5
      Top             =   3210
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Caption         =   "&OK"
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
      MICON           =   "frm_AddItem.frx":058A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_AddItem.frx":05A6
      BinaryImageOver =   "frm_AddItem.frx":05BE
   End
   Begin VistaSuitePro.OsenVistaButton cmdCancel 
      Height          =   345
      Left            =   1710
      TabIndex        =   6
      Top             =   3210
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Caption         =   "&Cancel"
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
      MICON           =   "frm_AddItem.frx":05D6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm_AddItem.frx":05F2
      BinaryImageOver =   "frm_AddItem.frx":060A
   End
   Begin VistaSuitePro.OsenVistaTab OsenXPTab1 
      Height          =   2685
      Left            =   150
      TabIndex        =   7
      Top             =   450
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   4736
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
      BackColorParent =   15790320
      TabWidth1       =   55
      TabText1        =   "General"
      TabEnabled1     =   -1  'True
      TabVisible1     =   0   'False
      Begin VistaSuitePro.OsenVistaComboBox cboProduct 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   540
         Width           =   2505
         _ExtentX        =   4419
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
         TextColumn      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSGL            =   -1  'True
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
         TextColumn      =   1
         Required        =   -1  'True
         Unicode         =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   33023
         AutoChangeBorderColor=   0   'False
      End
      Begin VistaSuitePro.OsenVistaTextBox txtInfo 
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
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
         CurrencyBackColor=   8421504
         CurrencyForeColor=   16777215
         LabelCaption    =   "Unit Price:"
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
      Begin VistaSuitePro.OsenVistaTextBox txtInfo 
         Height          =   330
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   1350
         Width           =   3735
         _ExtentX        =   6588
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
         LabelCaption    =   "Quantity:"
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
      Begin VistaSuitePro.OsenVistaTextBox txtInfo 
         Height          =   330
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Top             =   1740
         Width           =   3735
         _ExtentX        =   6588
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
         LabelCaption    =   "Discount (%):"
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
      Begin VistaSuitePro.OsenVistaTextBox txtInfo 
         Height          =   330
         Index           =   4
         Left            =   180
         TabIndex        =   4
         Top             =   2130
         Width           =   3735
         _ExtentX        =   6588
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
         Locked          =   -1  'True
         NumberOnly      =   -1  'True
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
         LabelCaption    =   "Extended Price:"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   570
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frm_AddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboProduct_Click()

    If cboProduct.ListIndex > -1 Then
        txtInfo(1) = cboProduct.ColumnText(2) ' UnitPrice
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    ReDim strItem(6) As String
    strItemX = cboProduct.GetKeyValue  ' Get productID
    strItemX = strItemX & vbTab & cboProduct.Text   ' Get productName
    strItemX = strItemX & vbTab & txtInfo(1) ' unitprice
    strItemX = strItemX & vbTab & txtInfo(2) ' quantity
    strItemX = strItemX & vbTab & (Val(txtInfo(3)) / 100) 'discount
    strItemX = strItemX & vbTab & Format(txtInfo(4), "0.00") ' Extended Price
    bItemChanged = True
    Unload Me
End Sub

Private Sub Form_Load()

    Me.OsenXPForm1.Init Me
    
    cboProduct.InsertItemByRecordset GetRST("select productid,productname,unitprice from products"), True, True
    cboProduct.ColumnWidth(1) = 0 ' Hidden the productID column
    cboProduct.ColumnWidth(2) = 200
    cboProduct.ColumnWidth(3) = 0 ' Hidden the UnitPrice column
    cboProduct.TextColumn = 1 ' Display Productname as Text
    
    bItemChanged = False
    
    If UBound(strItem) Then
    
        cboProduct.KeyValue = strItem(1)
        txtInfo(1) = strItem(2)
        txtInfo(2) = strItem(3)
        txtInfo(3) = strItem(4) * 100
        
    End If
    
End Sub

Private Sub txtInfo_Change(Index As Integer)

    On Error Resume Next
    
    If Index <> 4 Then
        ' calculate Extended Price
        txtInfo(4) = Val(txtInfo(1)) * Val(txtInfo(2))
        
        If txtInfo(3).Value > 0 Then
            txtInfo(4) = Val(txtInfo(4)) - (txtInfo(4) * (txtInfo(3) / 100))
        End If
        
    End If
    
End Sub


















