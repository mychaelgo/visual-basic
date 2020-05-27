VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_employees 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Employees"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   Icon            =   "frm_employees.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyADODC MyADODC1 
      Height          =   555
      Left            =   210
      TabIndex        =   23
      Top             =   5310
      Width           =   6315
      _ExtentX        =   11139
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
      ShowRefreshButton=   0   'False
      ShowPrinterButton=   0   'False
      ShowGriper      =   0   'False
      AutoConfirmBeforeDelete=   -1  'True
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture2 
      Align           =   1  'Align Top
      Height          =   885
      Left            =   0
      TabIndex        =   22
      Top             =   420
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   1561
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
      Picture         =   "frm_employees.frx":038A
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   16310477
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "detail of an employee is as follow ..."
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Employees records"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextLeft        =   17
      DescriptionLeft =   37
      BinaryImage     =   "frm_employees.frx":1EDC
   End
   Begin VistaSuitePro.OsenVistaTab OsenXPTab1 
      Height          =   3855
      Left            =   210
      TabIndex        =   16
      Top             =   1410
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6800
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
      NumberOfTabs    =   2
      BackColorParent =   16767935
      TabWidth1       =   83
      TabText1        =   "Company Info"
      TabEnabled1     =   -1  'True
      TabVisible1     =   0   'False
      TabCountCtls1   =   10
      TabNo1CtlID1    =   "picX"
      TabNo1CtlIX1    =   1
      TabNo1CtlIT1    =   -1  'True
      TabNo1CtlID2    =   "cboReportTo"
      TabNo1CtlIX2    =   1
      TabNo1CtlIT2    =   -1  'True
      TabNo1CtlID3    =   "dtHire"
      TabNo1CtlIX3    =   1
      TabNo1CtlIT3    =   -1  'True
      TabNo1CtlID4    =   "TxtData(0)"
      TabNo1CtlIX4    =   1
      TabNo1CtlIT4    =   -1  'True
      TabNo1CtlID5    =   "TxtData(1)"
      TabNo1CtlIX5    =   1
      TabNo1CtlIT5    =   -1  'True
      TabNo1CtlID6    =   "TxtData(2)"
      TabNo1CtlIX6    =   1
      TabNo1CtlIT6    =   -1  'True
      TabNo1CtlID7    =   "TxtData(3)"
      TabNo1CtlIX7    =   1
      TabNo1CtlIT7    =   -1  'True
      TabNo1CtlID8    =   "OsenXPLabel1(4)"
      TabNo1CtlIX8    =   1
      TabNo1CtlIT8    =   -1  'True
      TabNo1CtlID9    =   "TxtData(13)"
      TabNo1CtlIX9    =   1
      TabNo1CtlIT9    =   -1  'True
      TabNo1CtlID10   =   "OsenXPLabel1(5)"
      TabNo1CtlIX10   =   1
      TabNo1CtlIT10   =   -1  'True
      TabWidth2       =   80
      TabText2        =   "Personal Info"
      TabEnabled2     =   -1  'True
      TabVisible2     =   0   'False
      TabCountCtls2   =   10
      TabNo2CtlID1    =   "dtBirth"
      TabNo2CtlIX1    =   2
      TabNo2CtlIT1    =   -1  'True
      TabNo2CtlID2    =   "TxtData(7)"
      TabNo2CtlIX2    =   2
      TabNo2CtlIT2    =   -1  'True
      TabNo2CtlID3    =   "TxtData(9)"
      TabNo2CtlIX3    =   2
      TabNo2CtlIT3    =   -1  'True
      TabNo2CtlID4    =   "TxtData(8)"
      TabNo2CtlIX4    =   2
      TabNo2CtlIT4    =   -1  'True
      TabNo2CtlID5    =   "TxtData(10)"
      TabNo2CtlIX5    =   2
      TabNo2CtlIT5    =   -1  'True
      TabNo2CtlID6    =   "TxtData(11)"
      TabNo2CtlIX6    =   2
      TabNo2CtlIT6    =   -1  'True
      TabNo2CtlID7    =   "TxtData(15)"
      TabNo2CtlIX7    =   2
      TabNo2CtlIT7    =   -1  'True
      TabNo2CtlID8    =   "TxtData(12)"
      TabNo2CtlIX8    =   2
      TabNo2CtlIT8    =   -1  'True
      TabNo2CtlID9    =   "TxtData(4)"
      TabNo2CtlIX9    =   2
      TabNo2CtlIT9    =   -1  'True
      TabNo2CtlID10   =   "OsenXPLabel1(15)"
      TabNo2CtlIX10   =   2
      TabNo2CtlIT10   =   -1  'True
      Begin VistaSuitePro.OsenVistaPicture picX 
         Height          =   3285
         Left            =   3360
         TabIndex        =   21
         Top             =   450
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   5794
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
         BorderColor     =   14854529
         GradientColor2  =   16310477
         BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BinaryImage     =   "frm_employees.frx":1EF4
      End
      Begin VistaSuitePro.OsenVistaDTPicker dtBirth 
         Height          =   315
         Left            =   76320
         TabIndex        =   13
         Top             =   3120
         Width           =   1635
         _ExtentX        =   2884
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
         FormatDate      =   "mmmm dd, yyyy"
         YEAR            =   0
         MONTH           =   0
         MYDATE          =   0
         thisdate        =   38533
         Text            =   "2005-06-30"
         BorderColor     =   14456432
         BorderColorOver =   12624503
         Mask            =   5
         Picture         =   "frm_employees.frx":1F0C
         FadeInEffect    =   -1  'True
         BinaryImage     =   "frm_employees.frx":5AFF
      End
      Begin VistaSuitePro.OsenVistaComboBox cboReportTo 
         Height          =   315
         Left            =   1305
         TabIndex        =   4
         Top             =   2310
         Width           =   1875
         _ExtentX        =   3307
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
         TextColumn      =   1
         Required        =   0   'False
         Unicode         =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
      Begin VistaSuitePro.OsenVistaDTPicker dtHire 
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Top             =   2790
         Width           =   1875
         _ExtentX        =   3307
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
         FormatDate      =   "mmmm dd, yyyy"
         YEAR            =   0
         MONTH           =   0
         MYDATE          =   0
         thisdate        =   38533
         Text            =   "2005-06-30"
         BorderColor     =   14456432
         BorderColorOver =   12624503
         Mask            =   5
         Picture         =   "frm_employees.frx":5B17
         FadeInEffect    =   -1  'True
         BinaryImage     =   "frm_employees.frx":970A
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   510
         Width           =   3000
         _ExtentX        =   5292
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
         Locked          =   -1  'True
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
         LabelCaption    =   "EmployeeID:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   3000
         _ExtentX        =   5292
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
         LabelCaption    =   "First Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   1410
         Width           =   3000
         _ExtentX        =   5292
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
         LabelCaption    =   "Last Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Top             =   1860
         Width           =   3000
         _ExtentX        =   5292
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
         LabelCaption    =   "Title:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   2340
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "Report To:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   13
         Left            =   180
         TabIndex        =   6
         Top             =   3240
         Width           =   3000
         _ExtentX        =   5292
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
         LabelCaption    =   "Extention:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   18
         Top             =   2790
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "Hire Date:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   645
         Index           =   7
         Left            =   75150
         TabIndex        =   7
         Top             =   450
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   1138
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
         MultiLine       =   -1  'True
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
         LabelCaption    =   "Address:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   9
         Left            =   78150
         TabIndex        =   9
         Top             =   1200
         Width           =   2955
         _ExtentX        =   5212
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
         LabelCaption    =   "Region:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   8
         Left            =   75150
         TabIndex        =   8
         Top             =   1230
         Width           =   2805
         _ExtentX        =   4948
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
         LabelCaption    =   "City:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   10
         Left            =   75150
         TabIndex        =   10
         Top             =   1680
         Width           =   2805
         _ExtentX        =   4948
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
         LabelCaption    =   "Postal Code:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   11
         Left            =   78150
         TabIndex        =   11
         Top             =   1680
         Width           =   2985
         _ExtentX        =   5265
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
         LabelCaption    =   "Country:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   1485
         Index           =   15
         Left            =   78180
         TabIndex        =   14
         Top             =   2160
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   2619
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
         MultiLine       =   -1  'True
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
         LabelCaption    =   "Notes:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelStyle      =   1
         LabelGradient   =   -1  'True
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   12
         Left            =   75150
         TabIndex        =   12
         Top             =   2130
         Width           =   2775
         _ExtentX        =   4895
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
         LabelCaption    =   "Home Phone:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   75
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox TxtData 
         Height          =   330
         Index           =   4
         Left            =   75180
         TabIndex        =   19
         Top             =   2610
         Width           =   2775
         _ExtentX        =   4895
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
         LabelCaption    =   "Title of Courtesy:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Index           =   15
         Left            =   75150
         TabIndex        =   20
         Top             =   3120
         Width           =   885
         _ExtentX        =   1561
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
         Caption         =   "Birth Date:"
         ForeColor       =   0
         BackStyle       =   0
      End
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
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
      Caption         =   "Employees"
      TitleTop        =   7
      icon            =   "frm_employees.frx":9722
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
End
Attribute VB_Name = "frm_employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    ' Make sure this form work fine ...
    Me.OsenXPForm1.Init Me
    
    ' Fill Employee Name into Cbo Reportto by Query (queryName:=vwEmployeeList)
    cboReportTo.InsertItemByRecordset GetRST("select * from vListEmp"), True, True
    cboReportTo.ColumnWidth(1) = 0

    If IsNew Then
        MyADODC1.OpenMySQLTable "employees", "employeeid", MyCN, " where employeeid=-1", txtdata, picX
    Else
        MyADODC1.OpenMySQLTable "employees", "employeeid", MyCN, " where employeeid=" & KeyValue, txtdata, picX
    End If

    MyADODC1.Bind cboReportTo, 16 ' Bind the combobox object (osenxpcombobox) into MyADODC1
    MyADODC1.Bind Me.dtBirth, 5
    MyADODC1.Bind Me.dtHire, 6
    
    If IsNew Then
        MyADODC1.SendAction 5 ' Addnew
    Else
        MyADODC1.SendAction 6 ' Edit
    End If

End Sub

Private Sub Form_Terminate()

    SetMyParent 0, hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    frmMain.RefreshView
    frmMain.Show
End Sub




