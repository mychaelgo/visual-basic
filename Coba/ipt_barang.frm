VERSION 5.00
Object = "{EE17B266-A61D-48F0-BB3E-5C4EC9EE2D1D}#1.1#0"; "osenxpsuite2009.ocx"
Begin VB.Form ipt_barang 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Input Barang Baru"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   Begin OSENXPSUITE2009OCX.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Input Barang Baru"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      BorderStyle     =   2
      UseDefaultTheme =   0   'False
   End
   Begin OSENXPSUITE2009OCX.OsenXPComboBox cbo_satuan 
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
      DropDownList    =   -1  'True
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LSMS            =   0
      LIO             =   2
      IMGLIST         =   ""
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   16777215
      ASURC           =   0   'False
      TextColumn      =   0
      Required        =   0   'False
      Unicode         =   0   'False
      CurrencySymbol  =   "Rp."
      BorderColor     =   12164479
      BorderColorOver =   12164479
      UseSystemGradientColor=   0   'False
      HeaderGradientAllow=   -1  'True
   End
   Begin OSENXPSUITE2009OCX.MyADODC ado 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   9
      Top             =   3165
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   979
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
         Name            =   "Arial"
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
   Begin OSENXPSUITE2009OCX.OsenXPTextBox txt_data 
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
      NumberOnly      =   -1  'True
      DecimalSeparator=   ","
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
      ThousandSeparator=   "."
   End
   Begin OSENXPSUITE2009OCX.OsenXPTextBox txt_data 
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
      DecimalSeparator=   ","
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
      ThousandSeparator=   "."
   End
   Begin OSENXPSUITE2009OCX.OsenXPTextBox txt_data 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
      DecimalSeparator=   ","
      ButtonEnabled   =   0   'False
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
      ThousandSeparator=   "."
   End
   Begin OSENXPSUITE2009OCX.OsenXPTextBox txt_data 
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
      NumberOnly      =   -1  'True
      DecimalSeparator=   ","
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
      ThousandSeparator=   "."
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Satuan"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Harga Barang"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   930
   End
End
Attribute VB_Name = "ipt_barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
