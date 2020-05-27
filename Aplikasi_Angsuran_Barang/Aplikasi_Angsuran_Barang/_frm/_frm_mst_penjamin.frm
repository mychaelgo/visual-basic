VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Penjamin Konsumen"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   Icon            =   "_frm_mst_penjamin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   10815
   Begin VB.Frame Frame2 
      Caption         =   "IDENTITAS USAHA PELANGGAN"
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
      Height          =   6180
      Index           =   0
      Left            =   5325
      TabIndex        =   34
      Top             =   735
      Width           =   5310
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   26
         Left            =   1770
         TabIndex        =   64
         Top             =   4935
         Width           =   2235
         _ExtentX        =   3942
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
      Begin VB.Frame Frame2 
         Height          =   525
         Index           =   2
         Left            =   1770
         TabIndex        =   66
         Top             =   5220
         Width           =   3435
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "3 Bulan"
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
            Left            =   2370
            TabIndex        =   69
            Top             =   150
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "1 Hari"
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
            Left            =   150
            TabIndex        =   67
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "2 Minggu"
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
            Left            =   1200
            TabIndex        =   68
            Top             =   150
            Width           =   1080
         End
      End
      Begin VB.Frame Frame5 
         Height          =   525
         Left            =   7290
         TabIndex        =   76
         Top             =   0
         Width           =   3615
         Begin VB.OptionButton Option1 
            Caption         =   "3 Bulan"
            Height          =   300
            Index           =   6
            Left            =   2160
            TabIndex        =   79
            Top             =   180
            Width           =   1080
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1 Hari"
            Height          =   300
            Index           =   4
            Left            =   90
            TabIndex        =   78
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2 Minggu"
            Height          =   300
            Index           =   5
            Left            =   1005
            TabIndex        =   77
            Top             =   180
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Index           =   1
         Left            =   7290
         TabIndex        =   72
         Top             =   0
         Width           =   3615
         Begin VB.OptionButton Option1 
            Caption         =   "3 Bulan"
            Height          =   300
            Index           =   7
            Left            =   2160
            TabIndex        =   75
            Top             =   180
            Width           =   1080
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1 Hari"
            Height          =   300
            Index           =   8
            Left            =   90
            TabIndex        =   74
            Top             =   150
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2 Minggu"
            Height          =   300
            Index           =   9
            Left            =   1005
            TabIndex        =   73
            Top             =   180
            Width           =   1080
         End
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   11
         Left            =   1770
         TabIndex        =   36
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
         Index           =   12
         Left            =   1770
         TabIndex        =   38
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
         Index           =   13
         Left            =   1770
         TabIndex        =   40
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
         Index           =   14
         Left            =   1770
         TabIndex        =   42
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
         Index           =   15
         Left            =   2145
         TabIndex        =   44
         Top             =   1680
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   3495
         TabIndex        =   46
         Top             =   1680
         Width           =   780
         _ExtentX        =   1376
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
         Left            =   1770
         TabIndex        =   48
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
         Index           =   19
         Left            =   1770
         TabIndex        =   50
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
         Index           =   20
         Left            =   1770
         TabIndex        =   52
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
         Index           =   21
         Left            =   1770
         TabIndex        =   54
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
         Index           =   22
         Left            =   1770
         TabIndex        =   56
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
         Index           =   23
         Left            =   1770
         TabIndex        =   58
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
         Index           =   24
         Left            =   1770
         TabIndex        =   60
         Top             =   4200
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
         Index           =   25
         Left            =   1770
         TabIndex        =   62
         Top             =   4575
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
         Index           =   27
         Left            =   1770
         TabIndex        =   71
         Top             =   5775
         Width           =   2235
         _ExtentX        =   3942
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
         Left            =   180
         TabIndex        =   65
         Top             =   5385
         Width           =   1350
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
         Left            =   180
         TabIndex        =   63
         Top             =   4995
         Width           =   1005
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
         Left            =   180
         TabIndex        =   59
         Top             =   4290
         Width           =   315
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
         TabIndex        =   61
         Top             =   4635
         Width           =   630
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
         Index           =   48
         Left            =   150
         TabIndex        =   57
         Top             =   3900
         Width           =   1080
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
         TabIndex        =   51
         Top             =   2820
         Width           =   405
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
         Index           =   6
         Left            =   150
         TabIndex        =   35
         Top             =   330
         Width           =   1035
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
         TabIndex        =   111
         Top             =   3375
         Width           =   2130
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
         TabIndex        =   110
         Top             =   1860
         Width           =   1605
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
         TabIndex        =   109
         Top             =   2895
         Width           =   1710
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
         TabIndex        =   108
         Top             =   2655
         Width           =   2250
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
         TabIndex        =   107
         Top             =   2220
         Width           =   2175
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
         TabIndex        =   106
         Top             =   1560
         Width           =   1560
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
         TabIndex        =   105
         Top             =   570
         Width           =   2280
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
         TabIndex        =   104
         Top             =   1020
         Width           =   1905
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
         TabIndex        =   103
         Top             =   1290
         Width           =   1920
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
         TabIndex        =   102
         Top             =   3660
         Width           =   1845
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
         TabIndex        =   101
         Top             =   4530
         Width           =   1845
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
         TabIndex        =   100
         Top             =   3165
         Width           =   1620
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
         TabIndex        =   99
         Top             =   3870
         Width           =   1980
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
         TabIndex        =   98
         Top             =   4305
         Width           =   2355
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan Jenis Usaha"
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
         Index           =   34
         Left            =   5325
         TabIndex        =   97
         Top             =   4485
         Width           =   2040
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
         TabIndex        =   96
         Top             =   4065
         Width           =   1905
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
         Index           =   13
         Left            =   7290
         TabIndex        =   95
         Top             =   3375
         Width           =   2130
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
         Index           =   14
         Left            =   7290
         TabIndex        =   94
         Top             =   1860
         Width           =   1605
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
         Index           =   16
         Left            =   7275
         TabIndex        =   93
         Top             =   2895
         Width           =   1710
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
         Index           =   17
         Left            =   7275
         TabIndex        =   92
         Top             =   2655
         Width           =   2250
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
         Index           =   20
         Left            =   7320
         TabIndex        =   91
         Top             =   2220
         Width           =   2175
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
         Index           =   36
         Left            =   7095
         TabIndex        =   90
         Top             =   1560
         Width           =   1560
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
         Index           =   37
         Left            =   7170
         TabIndex        =   89
         Top             =   570
         Width           =   2280
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
         Index           =   38
         Left            =   7170
         TabIndex        =   88
         Top             =   1020
         Width           =   1905
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
         Index           =   39
         Left            =   7170
         TabIndex        =   87
         Top             =   1290
         Width           =   1920
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
         Index           =   40
         Left            =   7260
         TabIndex        =   86
         Top             =   3660
         Width           =   1845
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
         Index           =   41
         Left            =   7245
         TabIndex        =   85
         Top             =   4530
         Width           =   1845
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
         Index           =   42
         Left            =   7290
         TabIndex        =   84
         Top             =   3165
         Width           =   1620
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
         Index           =   43
         Left            =   7335
         TabIndex        =   83
         Top             =   3870
         Width           =   1980
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
         Index           =   44
         Left            =   7065
         TabIndex        =   82
         Top             =   4305
         Width           =   2355
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penghasilan Jenis Usaha"
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
         Left            =   5325
         TabIndex        =   81
         Top             =   4485
         Width           =   2040
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
         Index           =   46
         Left            =   7830
         TabIndex        =   80
         Top             =   4065
         Width           =   1905
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
         Left            =   150
         TabIndex        =   53
         Top             =   3195
         Width           =   825
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
         Left            =   3060
         TabIndex        =   45
         Top             =   1740
         Width           =   345
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
         Left            =   150
         TabIndex        =   49
         Top             =   2445
         Width           =   945
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
         Left            =   150
         TabIndex        =   47
         Top             =   2100
         Width           =   870
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
         TabIndex        =   43
         Top             =   1740
         Width           =   300
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
         Left            =   150
         TabIndex        =   37
         Top             =   705
         Width           =   1470
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
         Left            =   150
         TabIndex        =   39
         Top             =   1050
         Width           =   1095
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
         Left            =   150
         TabIndex        =   41
         Top             =   1410
         Width           =   615
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
         Left            =   150
         TabIndex        =   55
         Top             =   3540
         Width           =   1080
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pengh. Tambahan"
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
         Left            =   180
         TabIndex        =   70
         Top             =   5850
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IDENTITAS PENJAMIN"
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
      Height          =   5070
      Left            =   165
      TabIndex        =   5
      Top             =   1830
      Width           =   5040
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   9
         Top             =   615
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
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   1305
         TabIndex        =   11
         Top             =   885
         Width           =   3615
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Suami"
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
            Left            =   60
            TabIndex        =   12
            Top             =   180
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Istri"
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
            Left            =   1050
            TabIndex        =   13
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Ortu"
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
            Index           =   2
            Left            =   1800
            TabIndex        =   14
            Top             =   180
            Width           =   780
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Lain-lain"
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
            Index           =   3
            Left            =   2580
            TabIndex        =   15
            Top             =   180
            Width           =   900
         End
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   7
         Top             =   270
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         Icon            =   "_frm_mst_penjamin.frx":058A
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
         Text            =   ""
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   2
         Left            =   1290
         TabIndex        =   17
         Top             =   1470
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
         Left            =   1650
         TabIndex        =   19
         Top             =   1830
         Width           =   735
         _ExtentX        =   1296
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
         Index           =   4
         Left            =   3150
         TabIndex        =   21
         Top             =   1830
         Width           =   690
         _ExtentX        =   1217
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
         Index           =   5
         Left            =   1290
         TabIndex        =   23
         Top             =   2205
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
         Left            =   1290
         TabIndex        =   25
         Top             =   2565
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
         Left            =   1290
         TabIndex        =   27
         Top             =   2925
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
         Left            =   1290
         TabIndex        =   29
         Top             =   3285
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
         Left            =   1290
         TabIndex        =   31
         Top             =   3645
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
         Index           =   10
         Left            =   1290
         TabIndex        =   33
         Top             =   4005
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
         TabIndex        =   6
         Top             =   300
         Width           =   420
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
         Left            =   135
         TabIndex        =   30
         Top             =   3720
         Width           =   615
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
         Left            =   135
         TabIndex        =   28
         Top             =   3315
         Width           =   825
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
         Left            =   135
         TabIndex        =   26
         Top             =   3000
         Width           =   405
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
         Left            =   135
         TabIndex        =   24
         Top             =   2655
         Width           =   900
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
         Left            =   135
         TabIndex        =   22
         Top             =   2295
         Width           =   870
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
         TabIndex        =   8
         Top             =   675
         Width           =   495
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
         TabIndex        =   16
         Top             =   1575
         Width           =   615
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis "
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
         Left            =   135
         TabIndex        =   10
         Top             =   1095
         Width           =   495
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
         Left            =   1305
         TabIndex        =   18
         Top             =   1890
         Width           =   300
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
         Left            =   2685
         TabIndex        =   20
         Top             =   1875
         Width           =   345
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
         Left            =   135
         TabIndex        =   32
         Top             =   4095
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "PELANGGAN"
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
      Height          =   1020
      Left            =   165
      TabIndex        =   0
      Top             =   750
      Width           =   5040
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   16
         Left            =   1290
         TabIndex        =   2
         Top             =   225
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         Icon            =   "_frm_mst_penjamin.frx":09D8
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   -1  'True
         BorderColor     =   33023
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
         Index           =   28
         Left            =   1290
         TabIndex        =   4
         Top             =   585
         Width           =   3645
         _ExtentX        =   6429
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
         Index           =   15
         Left            =   135
         TabIndex        =   3
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode "
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
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":0E26
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":11C0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":155A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":18F4
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":1C8E
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":2028
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":23C2
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":275C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":2AF6
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":2E90
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":322A
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":35C4
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":395E
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":3CF8
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_penjamin.frx":4292
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   112
      Top             =   7230
      Width           =   10815
      _ExtentX        =   19076
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   113
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1005
      ButtonWidth     =   1244
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baru"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Simpan"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hapus"
            ImageKey        =   "IMG3"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cari"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Batal"
            ImageKey        =   "IMG5"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Option"
            ImageKey        =   "IMG7"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageKey        =   "IMG8"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            ImageKey        =   "IMG9"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   0
      X2              =   19495
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   0
      X2              =   19495
      Y1              =   7140
      Y2              =   7140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   4
      X1              =   195
      X2              =   19690
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   19495
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   19495
      Y1              =   585
      Y2              =   585
   End
End
Attribute VB_Name = "Form7"
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
txtFields(0).Locked = CekAktifNo("006")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
PostFindForm "#" & txtFields(16).Hwnd1
PostFindForm "#" & txtFields(0).Hwnd1
CurRec.Close
Set CurRec = Nothing
End Sub

Private Sub Option1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case 13
            Select Case index
                   Case 0 To 3
                        txtFields(2).SetFocus
                   Case 10 To 12
                        txtFields(27).SetFocus
            End Select
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.index
       Case 1
           If CekUser("05", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            ClearControl Me
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            If CekAktifNo("006") Then txtFields(0).Text = getAutoNo("006")
            txtFields(16).SetFocus
           End If
       Case 2
           If CekUser("05", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("05", "D") = False Then
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
           If CekUser("05", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            Dim StrSql As String, Form4 As New frm_util_report
            Load Form4
            StrSql = "SELECT mst_Penjamin.[Kode Penjamin], mst_Penjamin.[Kode Pelanggan], mst_Penjamin.[Nama Penjamin], mst_Penjamin.[Alamat Penjamin], mst_Penjamin.[RT Penjamin], mst_Penjamin.[RW Penjamin], mst_Penjamin.[Kelurahan Penjamin], mst_Penjamin.[Kecamatan Penjamin], mst_Penjamin.[Kota Penjamin], mst_Penjamin.[Kode Pos Penjamin], " & _
                     "mst_Penjamin.[Telp Penjamin], mst_Penjamin.[No HP Penjamin] From mst_Penjamin <!where> ORDER BY mst_Penjamin.[Kode Penjamin];"

            Form4.ARView.Tag = "lap_penjamin|" & StrSql
            Form4.ShowField StrSql
            Form4.Show
            Form4.Left = 0
            Form4.Top = 0
            Form4.ZOrder 0
           End If
       Case 11
           Unload Me
End Select
End Sub

Sub SimpanData(nKey As String)
On Error Resume Next
Dim h As String, hErr As String, sJPenjamin, sPenghasilan, i As Integer

h = FindRecord("SELECT * FROM mst_penjamin WHERE (((mst_penjamin.[kode penjamin])='" & nKey & "'));")

'---Jenis Penjamin---'
If Option1(0).Value = True Then
sJPenjamin = 1
ElseIf Option1(1).Value = True Then
sJPenjamin = 2
ElseIf Option1(2).Value = True Then
sJPenjamin = 3
ElseIf Option1(3).Value = True Then
sJPenjamin = 4
End If

'---Penghasilan Usaha Penjamin---'
If Option1(10).Value = True Then
sPenghasilan = 3
ElseIf Option1(11).Value = True Then
sPenghasilan = 1
ElseIf Option1(12).Value = True Then
sPenghasilan = 2
End If

If h = "0" Then
                                                 
   h = SaveRecord("mst_penjamin", Array("kode penjamin=" & txtFields(0).Text, "kode pelanggan=" & txtFields(16).Text, "nama penjamin=" & txtFields(1).Text, "jenis penjamin=" & sJPenjamin, _
                                                  "alamat penjamin=" & txtFields(2).Text, "rt penjamin=" & txtFields(3).Text, "rw penjamin=" & txtFields(4).Text, "kelurahan penjamin=" & txtFields(5).Text, "kecamatan penjamin=" & txtFields(6).Text, "kota penjamin=" & txtFields(7).Text, "kode pos penjamin=" & txtFields(8).Text, "telp penjamin=" & txtFields(9).Text, "no hp penjamin=" & txtFields(10).Text, _
                                                  "jenis usaha penjamin=" & txtFields(11).Text, "nama perusahaan penjamin=" & txtFields(12).Text, "bidang usaha penjamin=" & txtFields(13).Text, "alamat usaha penjamin=" & txtFields(14).Text, "rt usaha penjamin=" & txtFields(15).Text, "rw usaha penjamin=" & txtFields(17).Text, "kelurahan usaha penjamin= " & txtFields(18).Text, "kecamatan usaha penjamin=" & txtFields(19).Text, "kota usaha penjamin=" & txtFields(20).Text, _
                                                  "kode pos usaha penjamin=" & txtFields(21).Text, _
                                                  "telp usaha1 penjamin=" & txtFields(22).Text, _
                                                  "telp usaha2 penjamin=" & txtFields(23).Text, _
                                                  "fax usaha penjamin=" & txtFields(24).Text, _
                                                  "jabatan usaha=" & txtFields(25).Text, _
                                                  "$penghasilan usaha=" & txtFields(26).Text, _
                                                  "penghasilan jenis usaha=" & sPenghasilan, _
                                                  "$penghasilan tambahan=" & txtFields(27).Text))
                                                  
  If h = "" Then
       If CekAktifNo("006") Then txtFields(0).Text = getAutoNo("006", True)
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
         h = UpdateRecord("mst_penjamin", Array("kode penjamin=" & txtFields(0).Text, "kode pelanggan=" & txtFields(16).Text, "nama penjamin=" & txtFields(1).Text, "jenis penjamin=" & sJPenjamin, _
                                                  "alamat penjamin=" & txtFields(2).Text, "rt penjamin=" & txtFields(3).Text, "rw penjamin=" & txtFields(4).Text, "kelurahan penjamin=" & txtFields(5).Text, "kecamatan penjamin=" & txtFields(6).Text, "kota penjamin=" & txtFields(7).Text, "kode pos penjamin=" & txtFields(8).Text, "telp penjamin=" & txtFields(9).Text, "no hp penjamin=" & txtFields(10).Text, _
                                                  "jenis usaha penjamin=" & txtFields(11).Text, "nama perusahaan penjamin=" & txtFields(12).Text, "bidang usaha penjamin=" & txtFields(13).Text, "alamat usaha penjamin=" & txtFields(14).Text, "rt usaha penjamin=" & txtFields(15).Text, "rw usaha penjamin=" & txtFields(17).Text, "kelurahan usaha penjamin= " & txtFields(18).Text, "kecamatan usaha penjamin=" & txtFields(19).Text, "kota usaha penjamin=" & txtFields(20).Text, _
                                                  "kode pos usaha penjamin=" & txtFields(21).Text, _
                                                  "telp usaha1 penjamin=" & txtFields(22).Text, _
                                                  "telp usaha2 penjamin=" & txtFields(23).Text, _
                                                  "fax usaha penjamin=" & txtFields(24).Text, _
                                                  "jabatan usaha=" & txtFields(25).Text, _
                                                  "$penghasilan usaha=" & txtFields(26).Text, _
                                                  "penghasilan jenis usaha=" & sPenghasilan, _
                                                  "$penghasilan tambahan=" & txtFields(27).Text), " WHERE [kode penjamin]='" & txtFields(0).Text & "'")
                                                          
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
hErr = FindRecord("SELECT mst_penjamin.[kode penjamin] From mst_penjamin WHERE (((mst_penjamin.[kode penjamin])='" & hKey & "'));")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Pelanggan ?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") = False Then
      GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE mst_penjamin.[kode penjamin] From mst_penjamin WHERE (((mst_penjamin.[kode penjamin])='" & hKey & "'));")
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
               ShowPenjamin NotNull(CurRec("Kode Penjamin")) & "|"
            'End If
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               If CurRec.BOF Then
                  CurRec.MoveNext
               End If
               ShowPenjamin NotNull(CurRec("Kode Penjamin")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               If CurRec.EOF Then
                  CurRec.MovePrevious
               End If
               ShowPenjamin NotNull(CurRec("Kode Penjamin")) & "|"
            End If
       Case 4
            'If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowPenjamin NotNull(CurRec("Kode Penjamin")) & "|"
            'End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowPenjamin NotNull(CurRec("Kode Penjamin")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [Kode Penjamin] From mst_Penjamin ORDER BY [Kode Penjamin]"
Return
End Sub

Private Sub txtFields_DownButtonClick(index As Integer)
On Error Resume Next
Select Case index
        Case 0
                  ShowFindForm "SELECT mst_penjamin.[kode penjamin], mst_penjamin.[kode pelanggan],mst_penjamin.[nama penjamin],mst_penjamin.[jenis penjamin]," & _
                                                  "mst_penjamin.[alamat penjamin],mst_penjamin.[rt penjamin],mst_penjamin.[rw penjamin],mst_penjamin.[kelurahan penjamin],mst_penjamin.[kecamatan penjamin],mst_penjamin.[kota penjamin],mst_penjamin.[Kode pos penjamin],mst_penjamin.[telp penjamin],mst_penjamin.[no hp penjamin], " & _
                                                  "mst_penjamin.[jenis usaha penjamin],mst_penjamin.[nama perusahaan penjamin],mst_penjamin.[bidang usaha penjamin],mst_penjamin.[alamat usaha penjamin],mst_penjamin.[rt usaha penjamin],mst_penjamin.[rw usaha penjamin], mst_penjamin.[kelurahan usaha penjamin], mst_penjamin.[kecamatan usaha penjamin],mst_penjamin.[kota usaha penjamin], " & _
                                                  "mst_penjamin.[kode pos usaha penjamin], " & _
                                                  "mst_penjamin.[telp usaha1 penjamin], " & _
                                                  "mst_penjamin.[telp usaha2 penjamin], " & _
                                                  "mst_penjamin.[fax usaha penjamin], " & _
                                                  "mst_penjamin.[jabatan usaha], " & _
                                                  "mst_penjamin.[penghasilan usaha], " & _
                                                  "mst_penjamin.[penghasilan jenis usaha], " & _
                                                  "mst_penjamin.[penghasilan tambahan], " & _
                                                  "FROM mst_penjamin <!where> ORDER BY mst_penjamin.[Kode Penjamin]; ", "#" & txtFields(index).Hwnd1, Me, "ShowPenjamin"
                                                  
       Case 16
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

Sub ShowPelanggan(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "Select * from mst_Pelanggan WHERE [Kode Pelanggan]='" & hKey(0) & "' ORDER BY mst_pelanggan.[Kode Pelanggan]")
If hErr = "" Then
    If Not rc.EOF Then
        txtFields(16).Text = NotNull(rc("Kode Pelanggan"))
        txtFields(28).Text = NotNull(rc("Nama"))
    Else
kembali:
       ClearControl Me
End If
Else
   GoSub kembali
End If
rc.Close
End Sub

Private Sub txtFields_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case 16
            Select Case index
               Case 0
                      ShowPenjamin txtFields(index).Text & "|"
                   Case 16
                      ShowPelanggan txtFields(index).Text & "|"
           End Select
      Case 13
            Select Case index
               Case 1
                   Option1(0).SetFocus
               Case 26
                   Option1(11).SetFocus
                   
           End Select
       
       Case Else
            If Me.Tag = "" Then
               Me.Tag = "*"
               Me.Caption = Me.Caption & Me.Tag
            End If
End Select
End Sub

Sub ShowPenjamin(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
Dim sJPenjamin, sPenghasilan, i As Integer

hKey = Split(nKey, "|")

hErr = SelectQuery(rc, "Select * from mst_Penjamin WHERE [Kode Penjamin]='" & hKey(0) & "' ORDER BY mst_penjamin.[Kode Penjamin]")

If hErr = "" Then
    If Not rc.EOF Then
        ShowPelanggan NotNull(rc("Kode Pelanggan").Value) & "|"
        txtFields(0).Text = NotNull(rc("Kode Penjamin"))
        txtFields(16).Text = NotNull(rc("Kode Pelanggan"))
        txtFields(1).Text = NotNull(rc("Nama Penjamin"))
        
        sJPenjamin = NotNull(rc("jenis penjamin"))
        If sJPenjamin = 1 Then
        Option1(0).Value = True
        ElseIf sJPenjamin = 2 Then
        Option1(1).Value = True
         ElseIf sJPenjamin = 3 Then
        Option1(2).Value = True
         ElseIf sJPenjamin = 4 Then
        Option1(3).Value = True
        End If
                  
        txtFields(2).Text = NotNull(rc("alamat penjamin"))
        txtFields(3).Text = NotNull(rc("rt penjamin"))
        txtFields(4).Text = NotNull(rc("rw penjamin"))
        txtFields(5).Text = NotNull(rc("Kelurahan penjamin"))
        txtFields(6).Text = NotNull(rc("Kecamatan penjamin"))
        txtFields(7).Text = NotNull(rc("Kota penjamin"))
        txtFields(8).Text = NotNull(rc("Kode Pos penjamin"))
        txtFields(9).Text = NotNull(rc("Telp penjamin"))
        txtFields(10).Text = NotNull(rc("no hp penjamin"))
        
        txtFields(11).Text = NotNull(rc("jenis usaha penjamin"))
        txtFields(12).Text = NotNull(rc("nama perusahaan penjamin"))
        txtFields(13).Text = NotNull(rc("bidang usaha penjamin"))
        txtFields(14).Text = NotNull(rc("alamat usaha penjamin"))
        txtFields(15).Text = NotNull(rc("rt usaha penjamin"))
        txtFields(17).Text = NotNull(rc("rw usaha penjamin"))
        txtFields(18).Text = NotNull(rc("kelurahan usaha penjamin"))
        txtFields(19).Text = NotNull(rc("kecamatan usaha penjamin"))
        txtFields(20).Text = NotNull(rc("kota usaha penjamin"))
        txtFields(21).Text = NotNull(rc("kode pos usaha penjamin"))
        txtFields(22).Text = NotNull(rc("telp usaha1 penjamin"))
        txtFields(23).Text = NotNull(rc("telp usaha2 penjamin"))
        txtFields(24).Text = NotNull(rc("fax usaha penjamin"))
        txtFields(25).Text = NotNull(rc("jabatan usaha"))
        txtFields(26).Text = NotNull(rc("penghasilan usaha"))
           
        sPenghasilan = NotNull(rc("penghasilan jenis usaha"))
        If sPenghasilan = 1 Then
        Option1(11).Value = True
        ElseIf sPenghasilan = 2 Then
        Option1(12).Value = True
        ElseIf sPenghasilan = 3 Then
        Option1(10).Value = True
        End If
        
        txtFields(27).Text = NotNull(rc("penghasilan tambahan"))
     Else
kembali:
      ClearControl Me
    End If
Else
   GoSub kembali
End If
rc.Close
End Sub
                 
