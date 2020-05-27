VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form frm_AS 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Advance Spreadsheet Demo"
   ClientHeight    =   6480
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_advance_spreadsheet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   7560
      Top             =   4980
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   4592
      Images          =   "frm_advance_spreadsheet.frx":038A
      Version         =   196608
      KeyCount        =   4
      Keys            =   "ÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   5745
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10134
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
      BackSelected    =   7381139
      BackSelectedG1  =   16777215
      BackSelectedG2  =   8632490
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   5
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      HeaderFormatString=   $"frm_advance_spreadsheet.frx":159A
      Columns         =   9
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      MaxAllColumnWidth=   805
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   "MyImageList1"
      ForeColorSelected=   16576
      HeaderCaption   =   "OsenXPListBox1"
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientTop=   8569007
      HeaderGradientBottom=   4487779
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      CenteredAllColumnsAlignment=   -1  'True
      BinaryImage     =   "frm_advance_spreadsheet.frx":164D
      WindowColor     =   3
      Begin VistaSuitePro.OsenVistaSpin OsenXPSpin1 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   3150
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         Text            =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaComboBox OsenXPComboBox3 
         Height          =   345
         Left            =   6270
         TabIndex        =   5
         Top             =   2100
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
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
         Text            =   "ComboBox3"
         ComboStyle      =   1
         DisplayPicture  =   -1  'True
         DropDownList    =   -1  'True
         LBN             =   16777215
         LBS             =   7381139
         LBG1            =   16777215
         LBG2            =   8632490
         LAR             =   -1  'True
         LS              =   2
         LSGL            =   -1  'True
         LIO             =   2
         LITL            =   22
         IMGLIST         =   "MyImageList1"
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
      Begin VistaSuitePro.OsenVistaComboBox OsenXPComboBox2 
         Height          =   345
         Left            =   5370
         TabIndex        =   4
         Top             =   3420
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
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
         Text            =   "Firebird"
         ComboStyle      =   1
         DropDownList    =   -1  'True
         LBN             =   16777215
         LBS             =   7381139
         LBG1            =   16777215
         LBG2            =   8632490
         LAR             =   -1  'True
         LSGL            =   -1  'True
         LLID            =   11
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
         DataList        =   "Visual Basic 6|Visual C++|Delphi|Java|PHP|.NET|SQL server|Oracle|MySQL|DB2|SQLite|Firebird"
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
      Begin VistaSuitePro.OsenVistaDTPicker OsenXPDTPicker1 
         Height          =   345
         Left            =   3270
         TabIndex        =   3
         Top             =   4080
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
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
         FormatDate      =   "mmmm dd, yyyy"
         YEAR            =   0
         MONTH           =   0
         MYDATE          =   0
         thisdate        =   38844
         Text            =   "2006-05-07"
         BorderColor     =   7381139
         BorderColorOver =   12624503
         Mask            =   3
         BinaryImage     =   "frm_advance_spreadsheet.frx":1665
      End
      Begin VistaSuitePro.OsenVistaTextBox OsenXPTextBox1 
         Height          =   315
         Left            =   3540
         TabIndex        =   2
         Top             =   2730
         Visible         =   0   'False
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Text            =   "TextBox1"
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
         BorderColor     =   8370596
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
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
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
      Caption         =   "Advance Spreadsheet Demo"
      TitleTop        =   7
      icon            =   "frm_advance_spreadsheet.frx":167D
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
End
Attribute VB_Name = "frm_AS"
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
    
  ' Set The Default Color scheme for All forms in this projects
  ' and set osenxpform1.usedefaulttheme=true
'    OsenXPComboBox1.ListIndex = 0

    ' Init rowcount
    Me.OsenXPListBox1.Rows = 1000
    
    
    ' Populate list for osenxpcombobox3
    With OsenXPComboBox3
        .Clear
        .SmallIcon = Me.MyImageList1.hIml ' set imagelist handle
        .AddItem "Database", NormalIcon:=0
        .AddItem "Home", NormalIcon:=1
        .AddItem "Personal", NormalIcon:=2
        .AddItem "Corporate", NormalIcon:=3
    End With
    
    ' Bind object(osenxpsuite controls) into osenxplistbox
    OsenXPListBox1.BindObject Me.OsenXPTextBox1, 2
    OsenXPListBox1.BindObject Me.OsenXPDTPicker1, 3
    OsenXPListBox1.BindObject Me.OsenXPSpin1, 4
    
    ' set custom cell style (as Button)
    OsenXPListBox1.ColumnStyle(5) = 11 ' Button
    
    OsenXPListBox1.BindObject Me.OsenXPComboBox2, 6
    OsenXPListBox1.BindObject Me.OsenXPComboBox3, 7
    
    OsenXPListBox1.ColumnStyle(8) = 1 ' CheckBox
    
    
End Sub

' Change colorscheme
Private Sub OsenXPComboBox1_ListIndexChanged()
'    Me.OsenXPForm1.ColorScheme = OsenXPComboBox1.ListIndex
End Sub


Private Sub OsenXPListBox1_CellButtonClick(lrow As Long, iCol As Integer, Value As String)
    
    Dim s As String
    
    s = InputBoxGT("Current Value: " & Value & vbCrLf & vbCrLf & "Please enter new value.", "Input")
    
    If s <> "" Then
         Value = s
    End If
    
End Sub






















