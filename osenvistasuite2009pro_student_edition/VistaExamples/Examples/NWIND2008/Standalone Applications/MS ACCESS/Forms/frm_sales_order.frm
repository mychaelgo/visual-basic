VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_orders 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Sales Order"
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_sales_order.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   665
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaStatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   9570
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   714
      BackColor       =   14936810
      ForeColor       =   16777215
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   4
      HaveXPForm      =   -1  'True
      PWidth1         =   230
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "http://www.osenvistasuite.com"
      pTextAlignment1 =   0
      PanelPicture1   =   "frm_sales_order.frx":058A
      PanelPicAlignment1=   0
      PWidth2         =   150
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Sub Total:"
      pTextAlignment2 =   0
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frm_sales_order.frx":08DC
      PanelPicAlignment2=   0
      PWidth3         =   150
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Freight:"
      pTextAlignment3 =   0
      pTextBold3      =   -1  'True
      PanelPicture3   =   "frm_sales_order.frx":08F8
      PanelPicAlignment3=   0
      PWidth4         =   150
      PMinWidth4      =   0
      pTTText4        =   ""
      pType4          =   0
      pText4          =   "Total:"
      pTextAlignment4 =   0
      pTextBold4      =   -1  'True
      PanelPicture4   =   "frm_sales_order.frx":0914
      PanelPicAlignment4=   0
      GradientColor1  =   10000535
      GradientColor2  =   5460819
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
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
      Caption         =   "Sales Order"
      TitleTop        =   7
      icon            =   "frm_sales_order.frx":0930
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AutoBackColor   =   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaToolBar tBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XPBlend         =   0   'False
      TotalButton     =   9
      Bpic1           =   "frm_sales_order.frx":0ECA
      Bname1          =   "New"
      Btype1          =   0
      Bwidth1         =   0
      Bchecked1       =   0   'False
      Bvalue1         =   0   'False
      Bpic2           =   "frm_sales_order.frx":121C
      Bname2          =   "Save"
      Btype2          =   0
      Bwidth2         =   0
      Bchecked2       =   0   'False
      Bvalue2         =   0   'False
      Bpic3           =   "frm_sales_order.frx":156E
      Bname3          =   "Delete"
      Btype3          =   0
      Bwidth3         =   0
      Bchecked3       =   0   'False
      Bvalue3         =   0   'False
      Bname4          =   ""
      Btype4          =   2
      Bwidth4         =   0
      Bchecked4       =   0   'False
      Bvalue4         =   0   'False
      BNI4            =   0
      BSI4            =   0
      Bpic5           =   "frm_sales_order.frx":18C0
      Bname5          =   "Preview Sales Order"
      Btype5          =   0
      Bwidth5         =   0
      Bchecked5       =   0   'False
      Bvalue5         =   0   'False
      BNI5            =   0
      BSI5            =   0
      Bname6          =   "Button6"
      Btype6          =   2
      Bwidth6         =   0
      Bchecked6       =   0   'False
      Bvalue6         =   0   'False
      Bpic7           =   "frm_sales_order.frx":1C12
      Bname7          =   "Create Invoice"
      Btype7          =   0
      Bwidth7         =   0
      Bchecked7       =   0   'False
      Bvalue7         =   0   'False
      Bname8          =   "Button9"
      Btype8          =   2
      Bwidth8         =   0
      Bchecked8       =   0   'False
      Bvalue8         =   0   'False
      Bpic9           =   "frm_sales_order.frx":1F64
      Bname9          =   "Close"
      Btype9          =   0
      Bwidth9         =   0
      Bchecked9       =   0   'False
      Bvalue9         =   0   'False
   End
   Begin VistaSuitePro.OsenVistaHookMenu OsenXPHookMenu1 
      Height          =   375
      Left            =   3330
      TabIndex        =   3
      Top             =   8370
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   688
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GripperLeft     =   8
   End
   Begin VistaSuitePro.MyContainerCtl Ctl 
      Align           =   1  'Align Top
      Height          =   7470
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   13176
      BackColor       =   8421504
      ScaleWidth      =   799
      ScaleHeight     =   498
      ClientWidth     =   10710
      ClientHeight    =   9900
      OffsetLR        =   8
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   9225
         Left            =   210
         ScaleHeight     =   9225
         ScaleWidth      =   10035
         TabIndex        =   5
         Top             =   210
         Width           =   10035
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame2 
            Height          =   1635
            Left            =   6210
            TabIndex        =   6
            Top             =   7320
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2884
            Caption         =   "Grand Total"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777215
            ForeColor       =   16777215
            BorderColor     =   7368816
            Appearance      =   1
            CaptionPosition =   1
            BinaryImage     =   "frm_sales_order.frx":22B6
            WindowColor     =   0
            GradientColor1  =   10000535
            GradientColor2  =   5460819
            Begin VistaSuitePro.OsenVistaTextBox TxtSubTotal 
               Height          =   315
               Left            =   1170
               TabIndex        =   7
               Top             =   420
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
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
               CurrencySymbol  =   "US $"
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
            Begin VistaSuitePro.OsenVistaTextBox txtFreight 
               Height          =   315
               Left            =   1170
               TabIndex        =   8
               Top             =   810
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
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
               CurrencySymbol  =   "US $"
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
            Begin VistaSuitePro.OsenVistaTextBox txtTotal 
               Height          =   315
               Left            =   1170
               TabIndex        =   9
               Top             =   1200
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
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
               CurrencySymbol  =   "US $"
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Total:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   180
               TabIndex        =   12
               Top             =   480
               Width           =   900
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Freight:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   180
               TabIndex        =   11
               Top             =   870
               Width           =   660
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   180
               TabIndex        =   10
               Top             =   1230
               Width           =   510
            End
         End
         Begin VistaSuitePro.OsenVistaDTPicker dtInfo 
            Height          =   315
            Index           =   0
            Left            =   1590
            TabIndex        =   13
            Top             =   4770
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
            FormatDate      =   "mmm dd,yyyy"
            YEAR            =   0
            MONTH           =   0
            MYDATE          =   0
            thisdate        =   38577
            Text            =   "2005-08-13"
            BorderColor     =   14456432
            BorderColorOver =   12624503
            Mask            =   5
            Picture         =   "frm_sales_order.frx":22CE
            Required        =   -1  'True
            FadeInEffect    =   -1  'True
            BinaryImage     =   "frm_sales_order.frx":5EC1
         End
         Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
            Height          =   735
            Left            =   6150
            TabIndex        =   14
            Top             =   3930
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   1296
            Caption         =   "&Ship Via:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777215
            ForeColor       =   16777215
            BorderColor     =   7368816
            Appearance      =   1
            image           =   "frm_sales_order.frx":5ED9
            BinaryImage     =   "frm_sales_order.frx":6473
            WindowColor     =   0
            GradientColor1  =   10000535
            GradientColor2  =   5460819
            Begin VistaSuitePro.OsenVistaOptionButton optVia 
               Height          =   285
               Index           =   1
               Left            =   210
               TabIndex        =   15
               Tag             =   "Speedy Express"
               Top             =   390
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               Value           =   -1  'True
               Caption         =   "&Speedy"
            End
            Begin VistaSuitePro.OsenVistaOptionButton optVia 
               Height          =   285
               Index           =   2
               Left            =   1350
               TabIndex        =   16
               Tag             =   "United Package"
               Top             =   390
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               Caption         =   "&United"
            End
            Begin VistaSuitePro.OsenVistaOptionButton optVia 
               Height          =   285
               Index           =   3
               Left            =   2430
               TabIndex        =   17
               Tag             =   "Federal Shipping"
               Top             =   390
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
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
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               Caption         =   "&Federal"
            End
         End
         Begin VistaSuitePro.OsenVistaListBox lstItemDetails 
            Height          =   1845
            Left            =   390
            TabIndex        =   18
            Top             =   5250
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   3254
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
            FontNormal      =   0
            BackSelected    =   10841658
            BackSelectedG1  =   16777215
            BackSelectedG2  =   33023
            AllowEdit       =   0   'False
            WordWrap        =   0   'False
            ItemHeightAuto  =   0   'False
            ItemOffset      =   2
            ItemTextLeft    =   17
            BorderColor     =   12958375
            Lstyle          =   2
            ShowHeader      =   -1  'True
            HeaderFormatString=   $"frm_sales_order.frx":648B
            Columns         =   7
            ShowGridLines   =   -1  'True
            XPAlphaBlend    =   0   'False
            AlternateRowColors=   -1  'True
            MaxAllColumnWidth=   590
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFontColor =   255
            ASURC           =   -1  'True
            IMGLIST         =   ""
            ForeColorSelected=   16576
            Picture         =   "frm_sales_order.frx":6529
            PicturePosition =   2
            TransparencyLevel=   33
            BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseSystemGradientColor=   0   'False
            BinaryImage     =   "frm_sales_order.frx":70C8
         End
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   615
            Index           =   0
            Left            =   1080
            TabIndex        =   19
            Top             =   2550
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   1085
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
            Locked          =   -1  'True
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
         Begin VistaSuitePro.OsenVistaComboBox CboInfo 
            Height          =   315
            Left            =   1080
            TabIndex        =   20
            Top             =   2160
            Width           =   3525
            _ExtentX        =   6218
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
            MAXROWS         =   5
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
               Name            =   "Arial"
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
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   21
            Top             =   3240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   22
            Top             =   3240
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   3
            Left            =   3630
            TabIndex        =   23
            Top             =   3240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   24
            Top             =   3600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   615
            Index           =   5
            Left            =   6150
            TabIndex        =   25
            Top             =   2520
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   1085
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
            Locked          =   -1  'True
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   6
            Left            =   6150
            TabIndex        =   26
            Top             =   3210
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   7
            Left            =   7950
            TabIndex        =   27
            Top             =   3210
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   8
            Left            =   8700
            TabIndex        =   28
            Top             =   3210
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   9
            Left            =   6150
            TabIndex        =   29
            Top             =   3570
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
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
         Begin VistaSuitePro.OsenVistaComboBox CboSales 
            Height          =   315
            Left            =   1590
            TabIndex        =   30
            Top             =   4290
            Width           =   3015
            _ExtentX        =   5318
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
            MAXROWS         =   3
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
               Name            =   "Arial"
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
         Begin VistaSuitePro.OsenVistaTextBox txtOrder 
            Height          =   285
            Left            =   7920
            TabIndex        =   31
            Top             =   1110
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
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
            Required        =   -1  'True
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
         Begin VistaSuitePro.OsenVistaTextBox txtInfo 
            Height          =   285
            Index           =   10
            Left            =   6150
            TabIndex        =   32
            Top             =   2160
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   503
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
            Required        =   -1  'True
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
         Begin VistaSuitePro.OsenVistaDTPicker dtInfo 
            Height          =   315
            Index           =   1
            Left            =   5040
            TabIndex        =   33
            Top             =   4770
            Width           =   1485
            _ExtentX        =   2619
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
            FormatDate      =   "mmm dd,yyyy"
            YEAR            =   0
            MONTH           =   0
            MYDATE          =   0
            thisdate        =   38577
            Text            =   "2005-08-13"
            BorderColor     =   14456432
            BorderColorOver =   12624503
            Mask            =   5
            Picture         =   "frm_sales_order.frx":70E0
            FadeInEffect    =   -1  'True
            BinaryImage     =   "frm_sales_order.frx":ACD3
         End
         Begin VistaSuitePro.OsenVistaDTPicker dtInfo 
            Height          =   315
            Index           =   2
            Left            =   8250
            TabIndex        =   34
            Top             =   4770
            Width           =   1455
            _ExtentX        =   2566
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
            FormatDate      =   "mmm dd,yyyy"
            YEAR            =   0
            MONTH           =   0
            MYDATE          =   0
            thisdate        =   38577
            Text            =   "2005-08-13"
            BorderColor     =   14456432
            BorderColorOver =   12624503
            Mask            =   5
            Picture         =   "frm_sales_order.frx":ACEB
            FadeInEffect    =   -1  'True
            BinaryImage     =   "frm_sales_order.frx":E8DE
         End
         Begin VB.Image Image1 
            Height          =   1335
            Left            =   450
            Picture         =   "frm_sales_order.frx":E8F6
            Top             =   150
            Width           =   3720
         End
         Begin VB.Label LBlINE 
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   45
            Left            =   420
            TabIndex        =   45
            Top             =   1500
            Width           =   9255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "One Portals Way, Twin Points WA  98156 Phone: 1-206-555-1417   Fax: 1-206-555-5938"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   390
            TabIndex        =   44
            Top             =   1590
            Width           =   3465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   8010
            TabIndex        =   43
            Top             =   1620
            Width           =   495
         End
         Begin VB.Label lbCurDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12-Aug-2005"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   42
            Top             =   1650
            Width           =   915
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Bill To:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   0
            Left            =   390
            TabIndex        =   41
            Top             =   2190
            Width           =   555
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Ship To:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   5310
            TabIndex        =   40
            Top             =   2160
            Width           =   675
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Salesperson:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   2
            Left            =   390
            TabIndex        =   39
            Top             =   4320
            Width           =   1125
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Order ID:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   3
            Left            =   7050
            TabIndex        =   38
            Top             =   1140
            Width           =   750
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Order Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   4
            Left            =   390
            TabIndex        =   37
            Top             =   4800
            Width           =   975
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Required Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   5
            Left            =   3630
            TabIndex        =   36
            Top             =   4800
            Width           =   1245
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Shipped Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   6
            Left            =   7020
            TabIndex        =   35
            Top             =   4800
            Width           =   1170
         End
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   "System Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAction 
         Caption         =   "&Insert"
         Index           =   1
      End
      Begin VB.Menu mnuAction 
         Caption         =   "&Edit"
         Index           =   2
      End
      Begin VB.Menu mnuAction 
         Caption         =   "&Remove"
         Index           =   3
      End
      Begin VB.Menu mnuAction 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuAction 
         Caption         =   "&Clear"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frm_orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs As CLS_ADODB_Recordset

Private strVia As String

Private IntVia As Integer

Private Sub CboInfo_Click()

    On Error Resume Next

    If CboInfo.ListIndex > -1 Then
        Dim i As Long

        For i = 0 To 4
            txtInfo(i) = CboInfo.ColumnText(i + 2)
        Next i

        For i = 5 To 9
            txtInfo(i) = CboInfo.ColumnText(i - 3)
        Next i

        txtInfo(10) = CboInfo.ColumnText(1) ' Company name (ship to)
        Me.OsenXPForm1.Caption = "Sales Order [" & txtOrder & "] - " & CboInfo.Text

    End If

End Sub

Private Sub Form_Load()

    ' Initialize OsenXPForm

    OsenXPForm1.Init Me

    ' Fill cboInfo with CustomerID and Company Name Address, City, Region, PostalCode, Country from customers table, but customerid hiden(invisible)
    CboInfo.InsertItemByRecordset GetRST("SELECT CustomerID, CompanyName, Address, City, Region, PostalCode, Country FROM Customers;"), True
    CboInfo.ColumnWidth(1) = 0 ' Hiden the customerid column
    CboInfo.ColumnWidth(2) = 215
    CboInfo.ColumnWidth(3) = 0 ' Hiden the customer addresses
    CboInfo.ColumnWidth(4) = 0 ' Hiden the customer city
    CboInfo.ColumnWidth(5) = 0 ' Hiden the customer region
    CboInfo.ColumnWidth(6) = 0 ' Hiden the customer postal code
    CboInfo.ColumnWidth(7) = 0 ' Hiden the customer country
    CboInfo.TextColumn = 1 ' Display Custome Name as Text
    
    ' Fill CboSales with employeeID and Employee name from employees table, and hiden the employeeID column
    CboSales.InsertItemByRecordset GetRST("spGetEmpFL"), True
    CboSales.ColumnWidth(1) = 0 ' Hiden/Invisible the employeeID column
    CboSales.ColumnWidth(2) = 215
    
    lstItemDetails.ColumnFormat(6) = "0.00 %"

    ' Set date picker with current date
    dtInfo(0).Value = Date
    dtInfo(1).Value = Date
    dtInfo(2).Value = Date
    
    lbCurDate.Caption = Format(Now(), "dd-mmm-yyyy")
    
    ' set default shipper
    optVia_Click 1
    IntVia = 1
    
    If IsNew Then
        CreateNewOrder
    Else
        DisplayOrder KeyValue
    End If

End Sub

Private Sub DisplayOrder(OrderID As String)

    If Len(OrderID) Then
        On Error Resume Next
        
        Set Rs = New CLS_ADODB_Recordset
        
        Set Rs.DBRecordset = GetRST("select * from orders where orderid=" & OrderID)

        If Rs.State And Rs.Have_Records Then

            txtOrder = OrderID

            ' set customer info
            CboInfo.KeyValue = Rs.sField("customerid")
                
            ' set sales info
            CboSales.KeyValue = Rs.sField("employeeid")

            ' set shipper info
            optVia(CInt(Rs.sField("shipvia"))).Value = True

            ' Display Order Date
            dtInfo(0).Value = IIf(Rs.sField("OrderDate") <> "", Rs.sField("OrderDate"), 0)

            ' Display Requested date
            dtInfo(1).Value = IIf(Rs.sField("RequiredDate") <> "", Rs.sField("RequiredDate"), 0)

            ' Display shipped date
            If Not IsNull(Rs.sField("ShippedDate")) Then
                dtInfo(2).Value = Rs.sField("ShippedDate")
            Else
                dtInfo(2).Text = ""
            End If
                
            txtFreight = Rs.sField("freight")

            ' Display Order Details
            DIsplayOrderDetails OrderID

        End If
    End If

End Sub

Private Sub DIsplayOrderDetails(OrderID As String)

    On Error Resume Next

    If Len(OrderID) Then
            
        mStrSQL = "vw_orderdetails " & OrderID
            
        ' Fill order detail info lstItemDetail
        lstItemDetails.InsertItemByRecordset GetRST(mStrSQL), , False
        lstItemDetails.ColumnFormat(6) = "0.00 %"
        lstItemDetails.ColumnWidth(1) = 0
        lstItemDetails.ColumnWidth(2) = 0
        lstItemDetails.Refresh
        
        TxtSubTotal = lstItemDetails.Sum(6)       ' aggregate function in listbox
        txtTotal = Val(TxtSubTotal) + Val(txtFreight)
            
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' On error resume next
    If IsNew Then
        If MsgBoxGT("Are you sure to discard new order.", vbYesNo + vbQuestion, "Confirm") = vbYes Then
            DeleteOrder txtOrder
        Else
            Cancel = 1
        End If
    End If
    
    frmMain.RefreshView
    frmMain.Show
    
End Sub

Private Sub lstItemDetails_DblClick()

    If lstItemDetails.ListCount > 0 Then
        MnuAction_Click 2
    End If

End Sub

Private Sub lstItemDetails_MouseDown(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)

    If Button = 2 Then
        ' set active child menu
        mnuAction(2).Enabled = lstItemDetails.ListCount
        mnuAction(3).Enabled = lstItemDetails.ListCount
        mnuAction(5).Enabled = lstItemDetails.ListCount
        
        ' Now, display context menu
        PopupMenu mnuItem, , , , mnuAction(1)
        
    End If

End Sub

Private Sub MnuAction_Click(Index As Integer)

    Dim ItemEdit As Long
    
    bItemChanged = False
    ItemEdit = -1
    
    Select Case Index

        Case 1
            ReDim strItem(0)
            frm_AddItem.Show 1

        Case 2
            ReDim strItem(5)

            If lstItemDetails.ListIndex > -1 Then
                ReDim strItem(4)
                strItem(1) = lstItemDetails.ColumnValue(1)
                strItem(2) = lstItemDetails.ColumnValue(3)
                strItem(3) = lstItemDetails.ColumnValue(4)
                strItem(4) = lstItemDetails.ColumnValue(5)
                ItemEdit = lstItemDetails.ListIndex
            Else
                ReDim strItem(0)
            End If

            frm_AddItem.Show 1

        Case 3

            If MsgBoxGT("Are you sure to delete these item?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
                lstItemDetails.RemoveItem lstItemDetails.ListIndex
            End If

        Case 5
            lstItemDetails.Clear
    End Select
    
    If bItemChanged Then
    
        If ItemEdit > -1 Then ' Edit current item
            lstItemDetails.List(ItemEdit) = txtOrder.Text & vbTab & strItemX
        Else ' Insert new item
            lstItemDetails.AddItem txtOrder.Text & vbTab & strItemX
        End If

    End If
    
    TxtSubTotal = lstItemDetails.Sum(6)   ' aggregate function in listbox
    txtTotal = TxtSubTotal.Value + txtFreight.Value
    
End Sub

Private Sub optVia_Click(Index As Integer)

    strVia = optVia(Index).Tag
    IntVia = Index

End Sub

Private Sub tBar_ButtonClick(Index As Integer, _
                             sText As String)

    Select Case Index

        Case 1
            CreateNewOrder
            IsNew = True

        Case 2
            SaveRecord
            IsNew = False

        Case 3
            DeleteOrder txtOrder
            IsNew = False

        Case 5
            DisplayReport

        Case 7
            DisplayReport "INVOICE"

        Case 9
            Unload Me

        Case Else
    End Select

End Sub

Private Sub txtFreight_Change()

    sBar.ExtendedCaption 3, txtFreight, enAlignRight, vbRed, True
    txtTotal.Value = TxtSubTotal.Value + txtFreight.Value
End Sub

Private Sub txtOrder_Change()

    Me.OsenXPForm1.Caption = "Sales Order [" & txtOrder & "] - " & CboInfo.Text

End Sub

Private Sub TxtSubTotal_Change()

    sBar.ExtendedCaption 2, TxtSubTotal, enAlignRight, vbBlue, True

End Sub

Private Sub txtTotal_Change()

    sBar.ExtendedCaption 4, txtTotal, enAlignRight, RGB(100, 100, 255), True

End Sub

Private Sub DisplayReport(Optional Title As String = "SALES ORDER")

    'On Error Resume Next

    If lstItemDetails.ListCount > 0 Then

        ' Create new recordset
        Dim RsReport As New Recordset
        Dim i As Long, J As Long

        ' Record field definition
        With RsReport

            .Fields.Append "productid", adInteger
            .Fields.Append "productname", adVarChar, 255 '15
            .Fields.Append "unitprice", adCurrency '16
            .Fields.Append "quantity", adInteger '17
            .Fields.Append "discount", adSingle '18
            .Fields.Append "extprice", adCurrency '19

            ' Open this recordset
            .Open

            J = lstItemDetails.ListCount

            For i = 1 To J

                .AddNew

                ' field definition (order details)
                .Fields(0).Value = CInt(lstItemDetails.TextMatrix(i - 1, 1)) ' ProductID
                .Fields(1).Value = lstItemDetails.TextMatrix(i - 1, 2) ' Product Name
                .Fields(2).Value = CCur(lstItemDetails.TextMatrix(i - 1, 3)) ' Unitprice
                .Fields(3).Value = CInt(lstItemDetails.TextMatrix(i - 1, 4)) ' Quantity
                .Fields(4).Value = CSng(lstItemDetails.TextMatrix(i - 1, 5)) ' Discount
                .Fields(5).Value = CCur(lstItemDetails.TextMatrix(i - 1, 6)) ' Extended Price

                .Update

            Next i

            rpt_SalesOrder.Sections(2).Controls("lb0").Caption = txtInfo(10) ' Company Name
            rpt_SalesOrder.Sections(2).Controls("lb1").Caption = txtInfo(0) ' Address
            rpt_SalesOrder.Sections(2).Controls("lb2").Caption = txtInfo(1) & "  " & txtInfo(3) ' City and Postal code
            rpt_SalesOrder.Sections(2).Controls("lb3").Caption = txtInfo(4) ' Country
            rpt_SalesOrder.Sections(2).Controls("lb4").Caption = txtInfo(10)
            rpt_SalesOrder.Sections(2).Controls("lb5").Caption = txtInfo(0)
            rpt_SalesOrder.Sections(2).Controls("lb6").Caption = txtInfo(1) & "  " & txtInfo(3)
            rpt_SalesOrder.Sections(2).Controls("lb7").Caption = txtInfo(4)

            rpt_SalesOrder.Sections(2).Controls("lbOrderID").Caption = txtOrder ' Order ID
            rpt_SalesOrder.Sections(2).Controls("lbCustomerID").Caption = CboInfo.GetKeyValue ' CustomerID
            rpt_SalesOrder.Sections(2).Controls("lbSalesPerson").Caption = CboSales.Text ' Salesperson
            rpt_SalesOrder.Sections(2).Controls("lbOrderDate").Caption = dtInfo(0).Text ' Order Date
            rpt_SalesOrder.Sections(2).Controls("lbRequiredDate").Caption = dtInfo(1).Text ' Required date
            rpt_SalesOrder.Sections(2).Controls("lbShippedDate").Caption = dtInfo(2).Text ' Shipped date
            rpt_SalesOrder.Sections(2).Controls("lbVia").Caption = strVia ' Shipper

            rpt_SalesOrder.Sections(2).Controls("lbDate").Caption = lbCurDate.Caption ' Current date
            rpt_SalesOrder.Sections(2).Controls("lbTitle").Caption = Title ' Title

            rpt_SalesOrder.Sections(4).Controls("lbsubtotal").Caption = Format$(TxtSubTotal, "$ 0.00")
            rpt_SalesOrder.Sections(4).Controls("lbfreight").Caption = Format$(txtFreight.Value, "$ 0.00")
            rpt_SalesOrder.Sections(4).Controls("lbtotal").Caption = Format$(txtTotal, "$ 0.00")

            rpt_SalesOrder.Caption = "Sales Order [" & txtOrder & "] - " & txtInfo(10)

            Set rpt_SalesOrder.DataSource = RsReport ' Fill order details

            rpt_SalesOrder.Show

        End With

    End If

End Sub

Private Sub CLearAll()

    Dim i As Integer

    For i = 0 To 10
        txtInfo(i).Text = ""
    Next i

    For i = 0 To 2
        dtInfo(i).Text = ""
    Next i

    CboInfo.ListIndex = -1
    CboSales.ListIndex = -1
    optVia(1).Value = True
    optVia_Click 1
    
    txtOrder = ""
    TxtSubTotal.Value = 0
    txtFreight.Value = 0
    txtTotal.Value = 0

    ' Cleanup list items
    lstItemDetails.Clear

End Sub

Private Sub DeleteOrder(OrderID As String)

    On Error Resume Next

    If Len(OrderID) Then

        If MsgBoxGT("Are you sure to delete these order?" & vbLf & "OrderID: " & OrderID & vbLf & "Customer: " & txtInfo(10).Text, vbQuestion + vbYesNo, "Delete Order: [" & OrderID & "] ?") = vbYes Then

            ' Delete Order details
            mStrSQL = "delete from `order details` where orderid=" & OrderID
            ADOCN.Execute mStrSQL

            ' Delete Order
            mStrSQL = "delete from orders where orderid=" & OrderID
            ADOCN.Execute mStrSQL

            ' Clean up TextBOx and other controls
            CLearAll

            ' Display report
            MsgBoxGT "Delete order successfully.", vbInformation, "Information"

        End If

    End If

End Sub

Private Sub CreateNewOrder()

    On Error Resume Next

    ' Clear all last data from controls
    CLearAll

    Set Rs = New CLS_ADODB_Recordset
    
    Rs.RsOpen ADOCN, "select * from orders where orderid=-1"
    
    Rs.AddNew
    
    Rs.sField("orderdate") = MySQLDate(Now(), True)
    
    Rs.Update
   
    ' get autonumber of OrderID from current recordset
    txtOrder = Rs.sField(0)
    
End Sub

Private Sub SaveRecord()

    On Error Resume Next
    Dim i As Long

    If lstItemDetails.ListCount Then
        
        mStrSQL = "Update Orders " & "Set CustomerID='" & CboInfo.GetKeyValue & "'" & _
                    ", EmployeeID=" & CboSales.GetKeyValue & ", OrderDate=" & MySQLDate(dtInfo(0).Text, True) & _
                    ", RequiredDate=" & MySQLDate(dtInfo(1).Text, True) & ", ShippedDate=" & MySQLDate(dtInfo(2).Text, True) & _
                    ", Shipvia=" & IntVia & ", Freight=" & txtFreight.Value & ", Shipname='" & CboInfo.Text & _
                    "', " & "ShipAddress='" & txtInfo(0) & "', " & "ShipCity='" & txtInfo(1) & _
                    "', " & "ShipRegion='" & txtInfo(2) & " ', " & "ShipPostalCode='" & txtInfo(3) & _
                    "', " & "ShipCountry='" & txtInfo(4) & "' " & "Where orderid=" & txtOrder
                    
        
        ADOCN.Execute mStrSQL
              
        ' Now update the [Order details] table
        mStrSQL = "delete from `order details` where orderid=" & txtOrder.Text
        ADOCN.Execute mStrSQL

        For i = 0 To lstItemDetails.ListCount - 1

            mStrSQL = "Insert into `order details` values(" & txtOrder & "," & lstItemDetails.TextMatrix(i, 1) & "," & lstItemDetails.TextMatrix(i, 3) & "," & lstItemDetails.TextMatrix(i, 4) & "," & lstItemDetails.TextMatrix(i, 5) & ")"

            ADOCN.Execute mStrSQL

        Next i

        ' display message
        MsgBoxGT "Save data finished.", vbInformation, "Data saved."

    End If

End Sub

