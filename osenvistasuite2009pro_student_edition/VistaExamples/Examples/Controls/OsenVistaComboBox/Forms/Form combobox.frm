VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Sample OsenXPComboBox"
   ClientHeight    =   7215
   ClientLeft      =   4980
   ClientTop       =   795
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form combobox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame2 
      Height          =   1485
      Left            =   270
      TabIndex        =   22
      Top             =   5160
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   2619
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
      BorderColor     =   14396553
      Appearance      =   1
      BinaryImage     =   "Form combobox.frx":014A
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaComboBox OsenXPComboBox1 
         Height          =   345
         Left            =   1260
         TabIndex        =   23
         Top             =   690
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   609
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
         LSH             =   -1  'True
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
         TextColumn      =   0
         Required        =   0   'False
         Unicode         =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
   End
   Begin VistaSuitePro.OsenVistaStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   20
      Top             =   6780
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   767
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   2
      HaveXPForm      =   -1  'True
      WindowColor     =   3
      PWidth1         =   100
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Osen Kusnadi"
      pTextAlignment1 =   0
      PanelPicAlignment1=   0
      PWidth2         =   350
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   ""
      pTextAlignment2 =   0
      PanelPicture2   =   "Form combobox.frx":0162
      PanelPicAlignment2=   0
      GradientColor1  =   15382160
      GradientColor2  =   12752244
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   4260
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   609
      Caption         =   "&Get Cell value by listIndex and Column"
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
      MICON           =   "Form combobox.frx":017E
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "Form combobox.frx":019A
      BinaryImageOver =   "Form combobox.frx":01B2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtCellValue 
      Height          =   315
      Left            =   2520
      TabIndex        =   19
      Top             =   4680
      Width           =   3435
      _ExtentX        =   6059
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VistaSuitePro.OsenVistaTextBox TxtTextColumn 
      Height          =   285
      Left            =   1260
      TabIndex        =   18
      Top             =   4680
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
      Text            =   "0"
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VistaSuitePro.OsenVistaTextBox TxtListIndex 
      Height          =   285
      Left            =   1260
      TabIndex        =   17
      Top             =   4320
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
      Text            =   "0"
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   375
      Left            =   2550
      TabIndex        =   13
      Top             =   3360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "&List of selected cell of OsenXPComboBox"
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
      MICON           =   "Form combobox.frx":01CA
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "Form combobox.frx":01E6
      BinaryImageOver =   "Form combobox.frx":01FE
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   1845
      Left            =   2550
      TabIndex        =   11
      Top             =   1440
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3254
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
      FontNormal      =   16777215
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      HeaderFormatString=   "ID;55;1|Company Name;150;0"
      Columns         =   2
      ShowGridLines   =   -1  'True
      AlternateRowColors=   -1  'True
      RowColor1       =   15527922
      MaxAllColumnWidth=   205
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
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
      BinaryImage     =   "Form combobox.frx":0216
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtColumn 
      Height          =   345
      Left            =   1320
      TabIndex        =   9
      Top             =   3480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      Text            =   "1"
      Alignment       =   2
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
      Value           =   1
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VistaSuitePro.OsenVistaCheckBox ChkAltRow 
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   2520
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   450
      BackColor       =   16767935
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
      Value           =   1
      Caption         =   "Alternate Row"
      Style           =   1
   End
   Begin VistaSuitePro.OsenVistaCheckBox ChkShowGrid 
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   2130
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   450
      BackColor       =   16767935
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
      Value           =   1
      Caption         =   "Show Grid"
      Style           =   1
   End
   Begin VistaSuitePro.OsenVistaCheckBox ChkShowHeader 
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   1740
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   450
      BackColor       =   16767935
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
      Value           =   1
      Caption         =   "Show Header Column"
      Style           =   1
   End
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
      Height          =   1065
      Left            =   300
      TabIndex        =   21
      Top             =   600
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1879
      Caption         =   "Combobox Style:"
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
      BorderColor     =   14396553
      Appearance      =   1
      BinaryImage     =   "Form combobox.frx":022E
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaOptionButton ChkCboStyle 
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         Caption         =   "Office XP 2003"
         BackStyle       =   0
         Style           =   1
      End
      Begin VistaSuitePro.OsenVistaOptionButton ChkCboStyle 
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         Value           =   -1  'True
         Caption         =   "Windows XP"
         BackStyle       =   0
         Style           =   1
      End
   End
   Begin VistaSuitePro.OsenVistaComboBox MyCbo 
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   720
      Width           =   3285
      _ExtentX        =   5794
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
      TextColumn      =   1
      MAXROWS         =   5
      LBN             =   16777215
      LBS             =   10841658
      LBG1            =   16777215
      LBG2            =   14854529
      LAR             =   -1  'True
      LS              =   1
      LSH             =   -1  'True
      LSGL            =   -1  'True
      LARC1           =   16380137
      LLID            =   0
      LIO             =   2
      LITL            =   17
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
      Required        =   0   'False
      Unicode         =   0   'False
      BorderColor     =   12164479
      BorderColorOver =   12164479
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
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
      Caption         =   "Sample OsenXPComboBox"
      icon            =   "Form combobox.frx":0246
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Column:"
      Height          =   195
      Left            =   300
      TabIndex        =   16
      Top             =   4710
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ListIndex:"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   4350
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Value:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   4020
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get Selected Value of MyCBO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2550
      TabIndex        =   12
      Top             =   1200
      Width           =   2475
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0..2 (or max = no of columns - 1)"
      Height          =   465
      Left            =   300
      TabIndex        =   10
      Top             =   3030
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TextColumn"
      Height          =   195
      Left            =   300
      TabIndex        =   8
      Top             =   3540
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenXPSuite 2006 - OsenXPFrame sample                                *'
'*  Copyright (c) 2006 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  This file is part of the OsenXPSuite 2006 sample applications.       *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenXPSuite 2006 documentation, and is provided "as is", without     *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

'mStrSQL is global variable from XP Library (OsenXPSuite2006.XP class object) [String]
'MsAccessConnString is global function from XP Library (OsenXPSuite2006.XP class object) [Return: ConnectionString]
'Ado_Open is global function form XP Library (OsenXPSuite2006.XP Class Object) [Boolean; True if connection opened and FALSE if failed]
'AdoCN is global variable from XP Library (OsenXPSuite2006.XP class object) [ADODB.Connection]

Private MyRow As Long
Private MyCol   As Integer

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
    ' Retrieve location of database file (nwind.mdb)
    mStrSQL = "..\nwind.mdb"
    
    ' prepare connectionstring for opening database
    mStrSQL = MsAccessConnString(mStrSQL)
    
    ' FYI
    Debug.Print mStrSQL
    
    ' Open the database connection right now ...
    If ADO_OPEN(mStrSQL) Then
    
        Debug.Print "Connection successful"
        ' Now, the AdoCN has been set
        
        ' sample combobox multi column
        MyCbo.InsertItemBySQL ADOCN, "Select customerid,companyname,contactname from customers", True, True
        
        OsenXPComboBox1.InsertItemBySQL ADOCN, "select ProductID,ProductName,UnitPrice from products", True, True
        OsenXPComboBox1.ColumnFormat(2) = "$ #,###.00"
        OsenXPComboBox1.TextColumn = 1
    
    Else
        MsgBoxGT "The database can't be connected!", vbExclamation, "Connection failed"
    End If
    
End Sub

Private Sub MyCbo_Click()
    TxtListIndex.Value = MyCbo.ListIndex
End Sub

Private Sub OsenXPButton1_Click()
On Error Resume Next
    Dim I As Long
    If MyCbo.ListCount > 0 Then
        Me.OsenXPListBox1.LockUpdate = True
        Me.OsenXPListBox1.Clear
        For I = 0 To MyCbo.ListCount - 1
            If MyCbo.Selected(I) Then
                Me.OsenXPListBox1.AddItem MyCbo.List(I)
            End If
        Next I
        Me.OsenXPListBox1.LockUpdate = False
    End If
End Sub

Private Sub OsenXPButton2_Click()
On Error Resume Next
    If MyCbo.ListCount > 0 Then
        If TxtListIndex.Value > -1 And TxtListIndex.Value < MyCbo.ListCount Then
            txtCellValue.Text = MyCbo.TextMatrix(TxtListIndex.Value, TxtTextColumn.Value)
        End If
    End If
End Sub

Private Sub TxtColumn_Change()
    If TxtColumn.Value >= 0 And TxtColumn.Value <= 2 Then
        MyCbo.TextColumn = TxtColumn.Value
    Else
        TxtColumn.Value = MyCbo.TextColumn
    End If
End Sub

Private Sub ChkAltRow_Click()
    MyCbo.AlternateRowColors = ChkAltRow.Value
    Me.OsenXPListBox1.AlternateRowColors = ChkAltRow.Value
End Sub

Private Sub ChkCboStyle_Click(Index As Integer)
    MyCbo.ComboStyle = Index
End Sub

Private Sub ChkShowGrid_Click()
    MyCbo.ShowGridLines = ChkShowGrid.Value
    OsenXPListBox1.ShowGridLines = ChkShowGrid.Value
End Sub

Private Sub ChkShowHeader_Click()
    MyCbo.ShowHeader = ChkShowHeader.Value
    Me.OsenXPListBox1.ShowHeader = ChkShowHeader.Value
End Sub























