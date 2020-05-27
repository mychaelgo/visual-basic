VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Pelanggan"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   Icon            =   "_frm_mst_pelanggan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   10725
   Begin VB.Frame Frame2 
      Caption         =   "IDENTITAS USAHA / PEKERJAAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7695
      Index           =   0
      Left            =   5280
      TabIndex        =   55
      Top             =   630
      Width           =   5340
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   28
         Left            =   1770
         TabIndex        =   99
         Top             =   6645
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FontFormat      =   2
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   26
         Left            =   1770
         TabIndex        =   92
         Top             =   5790
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         FontFormat      =   2
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   25
         Left            =   1770
         TabIndex        =   81
         Top             =   4200
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Index           =   3
         Left            =   1770
         TabIndex        =   83
         Top             =   4470
         Width           =   3435
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Karyawan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   22
            Left            =   2265
            TabIndex        =   86
            Top             =   180
            Width           =   1110
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Milik Pribadi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   21
            Left            =   90
            TabIndex        =   84
            Top             =   180
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Kontrak"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   20
            Left            =   1305
            TabIndex        =   85
            Top             =   180
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Index           =   1
         Left            =   7290
         TabIndex        =   104
         Top             =   0
         Width           =   3615
         Begin VB.OptionButton Option1 
            Caption         =   "2 Minggu"
            Height          =   300
            Index           =   14
            Left            =   1005
            TabIndex        =   107
            Top             =   180
            Width           =   1080
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1 Hari"
            Height          =   300
            Index           =   13
            Left            =   90
            TabIndex        =   106
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "3 Bulan"
            Height          =   300
            Index           =   3
            Left            =   2160
            TabIndex        =   105
            Top             =   180
            Width           =   1080
         End
      End
      Begin VB.Frame Frame3 
         Height          =   525
         Left            =   7290
         TabIndex        =   100
         Top             =   0
         Width           =   3615
         Begin VB.OptionButton Option1 
            Caption         =   "2 Minggu"
            Height          =   300
            Index           =   5
            Left            =   1005
            TabIndex        =   103
            Top             =   180
            Width           =   1080
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1 Hari"
            Height          =   300
            Index           =   4
            Left            =   90
            TabIndex        =   102
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "3 Bulan"
            Height          =   300
            Index           =   2
            Left            =   2160
            TabIndex        =   101
            Top             =   180
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Index           =   2
         Left            =   1770
         TabIndex        =   94
         Top             =   6075
         Width           =   3435
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "2. Minggu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   12
            Left            =   1185
            TabIndex        =   96
            Top             =   180
            Width           =   1050
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "1. Hari"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   150
            TabIndex        =   95
            Top             =   180
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "3. Bulan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   2340
            TabIndex        =   97
            Top             =   165
            Width           =   915
         End
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   13
         Left            =   1755
         TabIndex        =   57
         Top             =   255
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   14
         Left            =   1770
         TabIndex        =   59
         Top             =   615
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   15
         Left            =   1770
         TabIndex        =   61
         Top             =   960
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   16
         Left            =   1770
         TabIndex        =   63
         Top             =   1320
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   17
         Left            =   2310
         TabIndex        =   65
         Top             =   1680
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   18
         Left            =   4065
         TabIndex        =   67
         Top             =   1680
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   19
         Left            =   1770
         TabIndex        =   69
         Top             =   2040
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   20
         Left            =   1770
         TabIndex        =   71
         Top             =   2400
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   21
         Left            =   1770
         TabIndex        =   73
         Top             =   2760
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   22
         Left            =   1770
         TabIndex        =   75
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   23
         Left            =   1770
         TabIndex        =   77
         Top             =   3480
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   24
         Left            =   1770
         TabIndex        =   79
         Top             =   3840
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   27
         Left            =   1770
         TabIndex        =   90
         Top             =   5415
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   31
         Left            =   1770
         TabIndex        =   88
         Top             =   5040
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   62
         Left            =   2910
         TabIndex        =   142
         Top             =   5085
         Width           =   510
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Usaha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   59
         Left            =   165
         TabIndex        =   82
         Top             =   4650
         Width           =   1065
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Usaha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   58
         Left            =   165
         TabIndex        =   87
         Top             =   5085
         Width           =   990
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan ++"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   75
         Left            =   165
         TabIndex        =   98
         Top             =   6705
         Width           =   1230
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Usaha 1 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   72
         Left            =   165
         TabIndex        =   76
         Top             =   3570
         Width           =   1080
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   71
         Left            =   165
         TabIndex        =   62
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bidang Usaha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   70
         Left            =   165
         TabIndex        =   60
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   69
         Left            =   165
         TabIndex        =   58
         Top             =   660
         Width           =   1470
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   68
         Left            =   1770
         TabIndex        =   64
         Top             =   1740
         Width           =   300
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelurahan "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   67
         Left            =   165
         TabIndex        =   68
         Top             =   2115
         Width           =   870
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kecamatan "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   66
         Left            =   165
         TabIndex        =   70
         Top             =   2460
         Width           =   945
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RW :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   64
         Left            =   3585
         TabIndex        =   66
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   63
         Left            =   165
         TabIndex        =   74
         Top             =   3210
         Width           =   825
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan Tambahan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   53
         Left            =   7830
         TabIndex        =   137
         Top             =   4065
         Width           =   1905
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   45
         Left            =   7065
         TabIndex        =   136
         Top             =   4305
         Width           =   2355
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   44
         Left            =   7335
         TabIndex        =   135
         Top             =   3870
         Width           =   1980
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   43
         Left            =   7290
         TabIndex        =   134
         Top             =   3165
         Width           =   1620
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Usaha 2 Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   42
         Left            =   7245
         TabIndex        =   133
         Top             =   4530
         Width           =   1845
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Usaha 1 Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   41
         Left            =   7260
         TabIndex        =   132
         Top             =   3660
         Width           =   1845
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   40
         Left            =   7170
         TabIndex        =   131
         Top             =   1290
         Width           =   1920
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bidang Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   39
         Left            =   7170
         TabIndex        =   130
         Top             =   1020
         Width           =   1905
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   38
         Left            =   7170
         TabIndex        =   129
         Top             =   570
         Width           =   2280
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   37
         Left            =   7095
         TabIndex        =   128
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelurahan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   36
         Left            =   7320
         TabIndex        =   127
         Top             =   2220
         Width           =   2175
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kecamatan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   7275
         TabIndex        =   126
         Top             =   2655
         Width           =   2250
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   7275
         TabIndex        =   125
         Top             =   2895
         Width           =   1710
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RW Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   7290
         TabIndex        =   124
         Top             =   1860
         Width           =   1605
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   7290
         TabIndex        =   123
         Top             =   3375
         Width           =   2130
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan Tambahan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   35
         Left            =   7830
         TabIndex        =   122
         Top             =   4065
         Width           =   1905
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   33
         Left            =   7065
         TabIndex        =   121
         Top             =   4305
         Width           =   2355
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   32
         Left            =   7335
         TabIndex        =   120
         Top             =   3870
         Width           =   1980
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   31
         Left            =   7290
         TabIndex        =   119
         Top             =   3165
         Width           =   1620
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Usaha 2 Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   30
         Left            =   7245
         TabIndex        =   118
         Top             =   4530
         Width           =   1845
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Usaha 1 Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   29
         Left            =   7260
         TabIndex        =   117
         Top             =   3660
         Width           =   1845
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   28
         Left            =   7170
         TabIndex        =   116
         Top             =   1290
         Width           =   1920
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bidang Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   27
         Left            =   7170
         TabIndex        =   115
         Top             =   1020
         Width           =   1905
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   26
         Left            =   7170
         TabIndex        =   114
         Top             =   570
         Width           =   2280
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   25
         Left            =   7095
         TabIndex        =   113
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelurahan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   7320
         TabIndex        =   112
         Top             =   2220
         Width           =   2175
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kecamatan Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   23
         Left            =   7275
         TabIndex        =   111
         Top             =   2655
         Width           =   2250
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   22
         Left            =   7275
         TabIndex        =   110
         Top             =   2895
         Width           =   1710
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RW Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   21
         Left            =   7290
         TabIndex        =   109
         Top             =   1860
         Width           =   1605
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos Usaha Penjamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   7290
         TabIndex        =   108
         Top             =   3375
         Width           =   2130
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Usaha "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   165
         TabIndex        =   56
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   47
         Left            =   150
         TabIndex        =   72
         Top             =   2820
         Width           =   405
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Usaha 2 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   48
         Left            =   165
         TabIndex        =   78
         Top             =   3900
         Width           =   1080
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   49
         Left            =   165
         TabIndex        =   89
         Top             =   5490
         Width           =   630
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   50
         Left            =   165
         TabIndex        =   80
         Top             =   4275
         Width           =   315
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   51
         Left            =   165
         TabIndex        =   91
         Top             =   5850
         Width           =   1005
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan PER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   52
         Left            =   150
         TabIndex        =   93
         Top             =   6270
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IDENTITAS PELANGGAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7695
      Left            =   90
      TabIndex        =   38
      Top             =   630
      Width           =   5220
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   32
         Left            =   1440
         TabIndex        =   48
         Top             =   6705
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin VB.Frame Frame7 
         Height          =   555
         Left            =   1440
         TabIndex        =   50
         Top             =   6990
         Width           =   3615
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "KTP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   26
            Left            =   120
            TabIndex        =   51
            Top             =   180
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "SIM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   25
            Left            =   960
            TabIndex        =   52
            Top             =   180
            Width           =   630
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "KK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   24
            Left            =   1710
            TabIndex        =   53
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Lain - lain"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   23
            Left            =   2475
            TabIndex        =   54
            Top             =   180
            Width           =   1005
         End
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   10
         Left            =   1440
         TabIndex        =   35
         Top             =   5205
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin VB.Frame Frame6 
         Height          =   795
         Left            =   1440
         TabIndex        =   37
         Top             =   5475
         Width           =   3600
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Kost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   19
            Left            =   2565
            TabIndex        =   42
            Top             =   435
            Width           =   660
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Milik Pribadi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   18
            Left            =   90
            TabIndex        =   39
            Top             =   150
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Rumah Dinas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   17
            Left            =   1425
            TabIndex        =   40
            Top             =   165
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Milik Ortu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   16
            Left            =   90
            TabIndex        =   138
            Top             =   435
            Width           =   1035
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Kontrak"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   15
            Left            =   1425
            TabIndex        =   41
            Top             =   435
            Width           =   930
         End
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   12
         Left            =   1440
         TabIndex        =   7
         Top             =   1350
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin VB.Frame Frame4 
         Height          =   525
         Left            =   1440
         TabIndex        =   9
         Top             =   1620
         Width           =   3615
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Perempuan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   6
            Left            =   1380
            TabIndex        =   11
            Top             =   180
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Laki - Laki"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   7
            Left            =   90
            TabIndex        =   10
            Top             =   180
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   630
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin VB.Frame Frame5 
         Height          =   555
         Left            =   1440
         TabIndex        =   13
         Top             =   2085
         Width           =   3615
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Duda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   8
            Left            =   2730
            TabIndex        =   17
            Top             =   195
            Width           =   750
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Janda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   9
            Left            =   1875
            TabIndex        =   16
            Top             =   210
            Width           =   900
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Belum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   10
            Left            =   1050
            TabIndex        =   15
            Top             =   210
            Width           =   900
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Menikah"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   11
            Left            =   90
            TabIndex        =   14
            Top             =   195
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   285
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         Icon            =   "_frm_mst_pelanggan.frx":058A
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   -1  'True
         BorderColor     =   33023
         Locked          =   -1  'True
         AutoTab         =   -1  'True
         FocusBackColor  =   14737632
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         MaxLength       =   17
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   19
         Top             =   2670
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   21
         Top             =   3015
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         Alignment       =   2
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   4
         Left            =   2895
         TabIndex        =   23
         Top             =   3030
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         Alignment       =   2
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   5
         Left            =   1440
         TabIndex        =   25
         Top             =   3390
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   6
         Left            =   1440
         TabIndex        =   27
         Top             =   3765
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   7
         Left            =   1440
         TabIndex        =   29
         Top             =   4125
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   8
         Left            =   1440
         TabIndex        =   31
         Top             =   4485
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   9
         Left            =   1440
         TabIndex        =   33
         Top             =   4845
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   11
         Left            =   1440
         TabIndex        =   5
         Top             =   990
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FontFormat      =   1
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   29
         Left            =   1440
         TabIndex        =   44
         Top             =   6345
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   30
         Left            =   4200
         TabIndex        =   46
         Top             =   6330
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
         AutoTab         =   -1  'True
         FocusBackColor  =   12640511
         FocusForeColor  =   8388736
         FocusBackMainColor=   8438015
         FocusBorderColor=   33023
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   0
         Text            =   ""
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. ID Jenis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   61
         Left            =   150
         TabIndex        =   49
         Top             =   7185
         Width           =   1005
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   60
         Left            =   150
         TabIndex        =   47
         Top             =   6750
         Width           =   510
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml. Tanggungan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   57
         Left            =   2640
         TabIndex        =   45
         Top             =   6375
         Width           =   1410
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Tinggal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   56
         Left            =   150
         TabIndex        =   43
         Top             =   6390
         Width           =   1095
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rumah"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   55
         Left            =   150
         TabIndex        =   139
         Top             =   5910
         Width           =   570
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   54
         Left            =   150
         TabIndex        =   36
         Top             =   5670
         Width           =   570
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tmp Lahir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   150
         TabIndex        =   6
         Top             =   1410
         Width           =   840
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Lahir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   150
         TabIndex        =   4
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   1815
         Width           =   300
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No HP "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   150
         TabIndex        =   34
         Top             =   5235
         Width           =   510
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RW :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   2475
         TabIndex        =   22
         Top             =   3075
         Width           =   345
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   1455
         TabIndex        =   20
         Top             =   3090
         Width           =   300
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   150
         TabIndex        =   12
         Top             =   2310
         Width           =   525
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   150
         TabIndex        =   18
         Top             =   2715
         Width           =   615
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   2
         Top             =   705
         Width           =   495
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelurahan "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   150
         TabIndex        =   24
         Top             =   3480
         Width           =   870
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kecamatan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   150
         TabIndex        =   26
         Top             =   3795
         Width           =   900
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   150
         TabIndex        =   28
         Top             =   4170
         Width           =   405
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   150
         TabIndex        =   30
         Top             =   4545
         Width           =   825
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Telp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   150
         TabIndex        =   32
         Top             =   4905
         Width           =   615
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   345
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   4710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":09D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":0D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":110C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":14A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":1840
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":1F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":230E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":26A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":2A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":2DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":3176
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":3510
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":38AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":3E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pelanggan.frx":43DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   140
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   1005
      ButtonWidth     =   1270
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baru"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Simpan"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hapus"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cari"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Batal"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "P.Jamin"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Option"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   141
      Top             =   8400
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   -120
      X2              =   19375
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   -120
      X2              =   19375
      Y1              =   540
      Y2              =   540
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurRec As New ADODB.Recordset
Dim hBtn As MSComctlLib.Button

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Shift = 0 Then
    Select Case KeyCode
           Case vbKeyEscape
                If txtFields(0).Text = "" Then
                   Unload Me
                Else
                    Set hBtn = Toolbar2.Buttons(6)
                        Toolbar1_ButtonClick hBtn
                        Set hBtn = Nothing
                End If
                KeyCode = 0
          Case vbKeyF2
                Set hBtn = Toolbar2.Buttons(1)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
          Case vbKeyF3
                Set hBtn = Toolbar2.Buttons(2)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
          Case vbKeyF4
                Set hBtn = Toolbar2.Buttons(4)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
          Case vbKeyF5
                Set hBtn = Toolbar2.Buttons(5)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
                    
    End Select
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
txtFields(0).Locked = CekAktifNo("005")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
PostFindForm "#" & txtFields(0).Hwnd1
CurRec.Close
Set CurRec = Nothing
End Sub

Private Sub Option1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    Select Case index
            Case 7, 6
                Option1(11).SetFocus
            Case 8 To 11
                txtFields(2).SetFocus
            Case 15 To 19
                txtFields(29).SetFocus
            Case 23 To 26
                txtFields(13).SetFocus
            Case 20 To 22
                txtFields(31).SetFocus
            Case 0, 1, 12
                txtFields(28).SetFocus
                
    End Select
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
 Select Case Button.index
       Case 1
           If CekUser("06", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            ClearControl Me
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            If CekAktifNo("005") Then
               txtFields(0).Text = getAutoNo("005")
               txtFields(1).SetFocus
            Else
               txtFields(0).SetFocus
            End If
           End If
       Case 2
           If CekUser("06", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
              SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("06", "D") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            HapusData (txtFields(0).Text)
           End If
       Case 5
            txtFields_DownButtonClick 0
       Case 6
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            ClearControl Me
       Case 7
           If CekUser("06", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            Dim StrSql As String, Form4 As New frm_util_report
            Load Form4
            StrSql = "SELECT mst_Pelanggan.[Kode Pelanggan], mst_Pelanggan.Nama, mst_Pelanggan.[Tgl Lahir], mst_Pelanggan.[Tmp Lahir], mst_Pelanggan.Sex, mst_Pelanggan.Status, mst_Pelanggan.Alamat, mst_Pelanggan.RT, mst_Pelanggan.RW, mst_Pelanggan.Kelurahan, mst_Pelanggan.Kecamatan, mst_Pelanggan.Kota, mst_Pelanggan.[Kode Pos], mst_Pelanggan.Telp, mst_Pelanggan.[No HP], mst_Pelanggan.[Status Rumah], mst_Pelanggan.[Lama Tinggal], mst_Pelanggan.[Jml Tanggungan], mst_Pelanggan.[Jenis Usaha], " & _
                     "mst_Pelanggan.[Nama Perusahaan], mst_Pelanggan.[Bidang Usaha], mst_Pelanggan.[Alamat Usaha], mst_Pelanggan.[RT Usaha], mst_Pelanggan.[RW Usaha], mst_Pelanggan.[Kelurahan Usaha], mst_Pelanggan.[Kecamatan Usaha], mst_Pelanggan.[Kota Usaha], mst_Pelanggan.[Kode Pos Usaha], mst_Pelanggan.[Telp Usaha1], mst_Pelanggan.[Telp Usaha2], mst_Pelanggan.[Fax Usaha], mst_Pelanggan.[Status Usaha], mst_Pelanggan.[Lama Usaha], mst_Pelanggan.[Jabatan Usaha], mst_Pelanggan.[Penghasilan Usaha],  " & _
                     "mst_Pelanggan.[Penghasilan Jenis Usaha], mst_Pelanggan.[Penghasilan Tambahan], mst_Pelanggan.RefID, mst_Pelanggan.[RefID Jenis] From mst_Pelanggan <!where> ORDER BY mst_Pelanggan.[Kode Pelanggan];"


            Form4.ARView.Tag = "lap_pelanggan|" & StrSql
            Form4.ShowField StrSql
            Form4.Show
            Form4.Left = 0
            Form4.Top = 0
            Form4.ZOrder 0
           End If
       Case 8
            Form7.Show
       Case 9
           
       Case 12
            Unload Me
End Select
End Sub

Sub SimpanData(nKey As String)
On Error Resume Next
Dim h As String, hErr As String, sStatus, sSex, sSRumah, sSUsaha, sPenghasilan, sRefID, i As Integer

h = FindRecord("SELECT * FROM mst_pelanggan WHERE (((mst_pelanggan.[kode pelanggan])='" & nKey & "'));")

'---Sex---'
If Option1(6).Value = True Then
sSex = 2
ElseIf Option1(7).Value = True Then
sSex = 1
End If

'---Status---'
If Option1(8).Value = True Then
sStatus = 4
ElseIf Option1(9).Value = True Then
sStatus = 3
ElseIf Option1(10).Value = True Then
sStatus = 2
ElseIf Option1(11).Value = True Then
sStatus = 1
End If

'---Status Rumah---'
If Option1(15).Value = True Then
sSRumah = 4
ElseIf Option1(16).Value = True Then
sSRumah = 3
ElseIf Option1(17).Value = True Then
sSRumah = 2
ElseIf Option1(18).Value = True Then
sSRumah = 1
ElseIf Option1(19).Value = True Then
sSRumah = 5
End If

'---Status Usaha---'
If Option1(20).Value = True Then
sSUsaha = 2
ElseIf Option1(21).Value = True Then
sSUsaha = 1
ElseIf Option1(22).Value = True Then
sSUsaha = 3
End If

'---Penghasilan---'
If Option1(0).Value = True Then
sPenghasilan = 3
ElseIf Option1(1).Value = True Then
sPenghasilan = 1
ElseIf Option1(12).Value = True Then
sPenghasilan = 2
End If

'---RefID---'
If Option1(23).Value = True Then
sRefID = 4
ElseIf Option1(24).Value = True Then
sRefID = 3
ElseIf Option1(25).Value = True Then
sRefID = 2
ElseIf Option1(26).Value = True Then
sRefID = 1
End If

If h = "0" Then
                                                 
   h = SaveRecord("mst_pelanggan", Array("kode pelanggan=" & txtFields(0).Text, "nama=" & txtFields(1).Text, "@tgl lahir=" & txtFields(11).Text, _
                                                  "tmp lahir=" & txtFields(12).Text, "sex=" & sSex, "status=" & sStatus, "alamat=" & txtFields(2).Text, "rt=" & txtFields(3).Text, "rw=" & txtFields(4).Text, "kelurahan=" & txtFields(5).Text, "kecamatan=" & txtFields(6).Text, "kota=" & txtFields(7).Text, "kode pos=" & txtFields(8).Text, "telp=" & txtFields(9).Text, "no hp=" & txtFields(10).Text, "status rumah=" & sSRumah, _
                                                  "lama tinggal=" & txtFields(28).Text, "jml tanggungan=" & txtFields(29).Text, "jenis usaha=" & txtFields(13).Text, "nama perusahaan=" & txtFields(14).Text, "bidang usaha=" & txtFields(15).Text, "alamat usaha=" & txtFields(16).Text, "rt usaha=" & txtFields(17).Text, "rw usaha=" & txtFields(18).Text, "kelurahan usaha = " & txtFields(19).Text, "kecamatan usaha=" & txtFields(20).Text, "kota usaha=" & txtFields(21).Text, _
                                                  "kode pos usaha=" & txtFields(22).Text, _
                                                  "telp usaha1=" & txtFields(23).Text, _
                                                  "telp usaha2=" & txtFields(24).Text, _
                                                  "fax usaha=" & txtFields(25).Text, _
                                                  "status usaha=" & sSUsaha, _
                                                  "lama usaha=" & txtFields(31).Text, _
                                                  "jabatan usaha=" & txtFields(27).Text, _
                                                  "$penghasilan usaha=" & txtFields(26).Text, _
                                                  "penghasilan jenis usaha=" & sPenghasilan, _
                                                  "$penghasilan tambahan=" & txtFields(28).Text, _
                                                  "refid=" & txtFields(32).Text, _
                                                  "refid jenis=" & sRefID))
                                                  
                                                 
  If h = "" Then
       If CekAktifNo("005") Then txtFields(0).Text = getAutoNo("005", True)
       txtFields(0).Tag = txtFields(0).Text
       Me.Caption = Replace(Me.Caption, "*", "")
       Me.Tag = ""
  Else
     ShowDlgMsg Me, "Proses penyimpanan data gagal!", vbOK, h, True, False
  End If
                                                 
ElseIf h = "1" Then
   If ShowDlgMsg(Me, "Data sudah terdaftar!, update dengan data baru?", vbYesNo, Error, False, True, , , , , Me.name & "_update") = False Then
      GoSub SimpanLabel
   Else
      If SelectMsg = vbYes Then
SimpanLabel:
         h = UpdateRecord("mst_pelanggan", Array("kode pelanggan=" & txtFields(0).Text, "nama=" & txtFields(1).Text, "@tgl lahir=" & txtFields(11).Text, _
                                                  "tmp lahir=" & txtFields(12).Text, "sex=" & sSex, "status=" & sStatus, "alamat=" & txtFields(2).Text, "rt=" & txtFields(3).Text, "rw=" & txtFields(4).Text, "kelurahan=" & txtFields(5).Text, "kecamatan=" & txtFields(6).Text, "kota=" & txtFields(7).Text, "kode pos=" & txtFields(8).Text, "telp=" & txtFields(9).Text, "no hp=" & txtFields(10).Text, "status rumah=" & sSRumah, _
                                                  "lama tinggal=" & txtFields(28).Text, "jml tanggungan=" & txtFields(29).Text, "jenis usaha=" & txtFields(13).Text, "nama perusahaan=" & txtFields(14).Text, "bidang usaha=" & txtFields(15).Text, "alamat usaha=" & txtFields(16).Text, "rt usaha=" & txtFields(17).Text, "rw usaha=" & txtFields(18).Text, "kelurahan usaha = " & txtFields(19).Text, "kecamatan usaha=" & txtFields(20).Text, "kota usaha=" & txtFields(21).Text, _
                                                  "kode pos usaha=" & txtFields(22).Text, _
                                                  "telp usaha1=" & txtFields(23).Text, _
                                                  "telp usaha2=" & txtFields(24).Text, _
                                                  "fax usaha=" & txtFields(25).Text, _
                                                  "status usaha=" & sSUsaha, _
                                                  "lama usaha=" & txtFields(31).Text, _
                                                  "jabatan usaha=" & txtFields(27).Text, _
                                                  "$penghasilan usaha=" & txtFields(26).Text, _
                                                  "penghasilan jenis usaha=" & sPenghasilan, _
                                                  "$penghasilan tambahan=" & txtFields(28).Text, _
                                                  "refid=" & txtFields(32).Text, _
                                                  "refid jenis=" & sRefID), " WHERE [kode pelanggan]='" & txtFields(0).Text & "'")
        If h = "" Then
             Me.Caption = Replace(Me.Caption, "*", "")
             txtFields(0).Tag = txtFields(0).Text
             Me.Tag = ""
        Else
           ShowDlgMsg Me, "Proses penyimpanan data gagal!", vbOK, h, True, False
        End If
                                                          
                                                          
      End If
   End If
End If
End Sub

Sub HapusData(hKey As String)
On Error Resume Next
Dim hErr, h As String
hErr = FindRecord("SELECT mst_pelanggan.[kode pelanggan] From mst_pelanggan WHERE (((mst_pelanggan.[kode pelanggan])='" & hKey & "'));")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Pelanggan ?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") = False Then
      GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE mst_pelanggan.[kode pelanggan] From mst_pelanggan WHERE (((mst_pelanggan.[kode pelanggan])='" & hKey & "'));")
         If hErr = "" Then
            ClearControl Me
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = ""
         Else
            ShowDlgMsg Me, "Proses penghapusan data gagal!", vbOK, h, True, False
         End If
      End If
   End If
ElseIf hErr = "0" Then
    ShowDlgMsg Me, "Tidak ada data yang akan dihapus", vbOK, , True, False
End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If CurRec.State = 0 Then
   GoSub subLoadDB
End If

Me.Caption = Replace(Me.Caption, "*", "")
Me.Tag = "*"
txtFields(0).Tag = ""
Select Case Button.index
       Case 1
            'If Not CurRec.BOF Then
               CurRec.MoveFirst
               ShowPelanggan NotNull(CurRec("Kode Pelanggan")) & "|"
            'End If
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               If CurRec.BOF Then
                  CurRec.MoveNext
               End If
               ShowPelanggan NotNull(CurRec("Kode Pelanggan")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               If CurRec.EOF Then
                  CurRec.MovePrevious
               End If
               ShowPelanggan NotNull(CurRec("Kode Pelanggan")) & "|"
            End If
       Case 4
            'If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowPelanggan NotNull(CurRec("Kode Pelanggan")) & "|"
            'End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowPelanggan NotNull(CurRec("Kode Pelanggan")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [Kode Pelanggan] From mst_Pelanggan ORDER BY [Kode Pelanggan]"
Return

End Sub

Private Sub txtFields_DownButtonClick(index As Integer)
On Error Resume Next
Select Case index
       Case 0
                 ShowFindForm "SELECT mst_pelanggan.[kode pelanggan],mst_pelanggan.[nama],mst_pelanggan.[tgl lahir],mst_pelanggan.[tmp lahir],mst_pelanggan.[sex],mst_pelanggan.[status],mst_pelanggan.[alamat],mst_pelanggan.[rt],mst_pelanggan.[rw],mst_pelanggan.[kelurahan],mst_pelanggan.[kecamatan], mst_pelanggan.kota,mst_pelanggan.[kode pos], mst_pelanggan.[telp], mst_pelanggan.[no hp],mst_pelanggan.[status rumah], " & _
                                                 "mst_pelanggan.[lama tinggal],mst_pelanggan.[jml tanggungan],mst_pelanggan.[jenis usaha],mst_pelanggan.[nama perusahaan],mst_pelanggan.[bidang usaha],mst_pelanggan.[alamat usaha], mst_pelanggan.[rt usaha], mst_pelanggan.[rw usaha],mst_pelanggan.[kelurahan usaha],mst_pelanggan.[kecamatan usaha],mst_pelanggan.[kota usaha], " & _
                                                  "mst_pelanggan.[kode pos usaha]," & _
                                                  "mst_pelanggan.[telp usaha1], " & _
                                                  "mst_pelanggan.[telp usaha2], " & _
                                                  "mst_pelanggan.[fax usaha], " & _
                                                  "mst_pelanggan.[status usaha], " & _
                                                  "mst_pelanggan.[lama usaha], " & _
                                                  "mst_pelanggan.[jabatan usaha], " & _
                                                  "mst_pelanggan.[penghasilan usaha], " & _
                                                  "mst_pelanggan.[penghasilan jenis usaha], " & _
                                                  "mst_pelanggan.[penghasilan tambahan], " & _
                                                  "mst_pelanggan.[refid],mst_pelanggan.[refid jenis], " & _
                                                  "FROM mst_pelanggan <!where> ORDER BY mst_pelanggan.[Kode Pelanggan]; ", "#" & txtFields(index).Hwnd1, Me, "ShowPelanggan"
End Select
End Sub

Private Sub txtFields_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case 0
            Select Case index
                   Case 0
                    ShowPelanggan txtFields(index).Text & "|"
            End Select
       Case 13
            Select Case index
                   Case 12
                     Option1(7).SetFocus
                   Case 10
                     Option1(18).SetFocus
                   Case 32
                     Option1(26).SetFocus
                   Case 25
                     Option1(21).SetFocus
                   Case 26
                     Option1(1).SetFocus
            End Select
       Case Else
            If Me.Tag = "" Then
               Me.Tag = "*"
               Me.Caption = Me.Caption & Me.Tag
            End If
End Select
End Sub

Sub ShowPelanggan(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
Dim sStatus, sSex, sSRumah, sSUsaha, sPenghasilan, sRefID, i As Integer

hKey = Split(nKey, "|")

hErr = SelectQuery(rc, "Select * from mst_Pelanggan WHERE [Kode Pelanggan]='" & hKey(0) & "' ORDER BY mst_pelanggan.[Kode Pelanggan]")

'MsgBox Herr

If hErr = "" Then
    If Not rc.EOF Then
        txtFields(0).Text = NotNull(rc("Kode Pelanggan"))
        txtFields(1).Text = NotNull(rc("Nama"))
        txtFields(12).Text = NotNull(rc("tmp lahir"))
        txtFields(11).Text = NotNull(rc("tgl lahir"))
        sSex = NotNull(rc("sex"))
        If sSex = 1 Then
        Option1(7).Value = True
        ElseIf sSex = 2 Then
        Option1(6).Value = True
        End If
        
        sStatus = NotNull(rc("status"))
        If sStatus = 1 Then
        Option1(11).Value = True
        ElseIf sStatus = 2 Then
        Option1(10).Value = True
        ElseIf sStatus = 3 Then
        Option1(9).Value = True
        ElseIf sStatus = 4 Then
        Option1(8).Value = True
        End If
              
        txtFields(2).Text = NotNull(rc("alamat"))
        txtFields(3).Text = NotNull(rc("rt"))
        txtFields(4).Text = NotNull(rc("rw"))
        txtFields(5).Text = NotNull(rc("Kelurahan"))
        txtFields(6).Text = NotNull(rc("Kecamatan"))
        txtFields(7).Text = NotNull(rc("Kota"))
        txtFields(8).Text = NotNull(rc("Kode Pos"))
        txtFields(9).Text = NotNull(rc("Telp"))
        txtFields(10).Text = NotNull(rc("no hp"))
        
        sSRumah = NotNull(rc("status rumah"))
        If sSRumah = 4 Then
        Option1(15).Value = True
        ElseIf sSRumah = 3 Then
        Option1(16).Value = True
        ElseIf sSRumah = 2 Then
        Option1(17).Value = True
        ElseIf sSRumah = 1 Then
        Option1(18).Value = True
        ElseIf sSRumah = 5 Then
        Option1(19).Value = True
        End If
        
        txtFields(29).Text = NotNull(rc("lama tinggal"))
        txtFields(30).Text = NotNull(rc("jml tanggungan"))
        txtFields(13).Text = NotNull(rc("jenis usaha"))
        txtFields(14).Text = NotNull(rc("nama perusahaan"))
        txtFields(15).Text = NotNull(rc("bidang usaha"))
        txtFields(16).Text = NotNull(rc("alamat usaha"))
        txtFields(17).Text = NotNull(rc("rt usaha"))
        txtFields(18).Text = NotNull(rc("rw usaha"))
        txtFields(19).Text = NotNull(rc("kelurahan usaha"))
        txtFields(20).Text = NotNull(rc("kecamatan usaha"))
        txtFields(21).Text = NotNull(rc("kota usaha"))
        txtFields(22).Text = NotNull(rc("kode pos usaha"))
        txtFields(23).Text = NotNull(rc("telp usaha1"))
        txtFields(24).Text = NotNull(rc("telp usaha2"))
        txtFields(25).Text = NotNull(rc("fax usaha"))
       
        sSUsaha = NotNull(rc("status usaha"))
        If sSUsaha = 1 Then
        Option1(21).Value = True
        ElseIf sSUsaha = 2 Then
        Option1(20).Value = True
        ElseIf sSUsaha = 3 Then
        Option1(22).Value = True
        End If
             
        txtFields(31).Text = NotNull(rc("lama usaha"))
        txtFields(27).Text = NotNull(rc("jabatan usaha"))
        txtFields(26).Text = NotNull(rc("penghasilan usaha"))
           
        sPenghasilan = NotNull(rc("penghasilan jenis usaha"))
        If sPenghasilan = 1 Then
        Option1(1).Value = True
        ElseIf sPenghasilan = 2 Then
        Option1(12).Value = True
        ElseIf sPenghasilan = 3 Then
        Option1(0).Value = True
        End If
        
        txtFields(28).Text = NotNull(rc("penghasilan tambahan"))
        txtFields(32).Text = NotNull(rc("refid"))
               
        sRefID = NotNull(rc("refid jenis"))
        If sRefID = 4 Then
        Option1(23).Value = True
        ElseIf sRefID = 3 Then
        Option1(24).Value = True
        ElseIf sRefID = 2 Then
        Option1(25).Value = True
        ElseIf sRefID = 1 Then
        Option1(26).Value = True
        End If
        
       Else
kembali:
      ClearControl Me
    End If
Else
   GoSub kembali
End If
rc.Close
End Sub
                 
