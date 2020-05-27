VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surat Perjanjian Sewa Beli - SPSB"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frmSPer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   7755
   Begin VB.Frame Frame1 
      Caption         =   "Bukti Pembayaran:"
      Height          =   3855
      Left            =   210
      TabIndex        =   38
      Top             =   3750
      Width           =   7320
      Begin VB.TextBox txtPrint 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5835
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   39
         Text            =   "_frmSPer.frx":038A
         Top             =   1395
         Visible         =   0   'False
         Width           =   1155
      End
      Begin SysInfo_Nardhika.vbButton vbCetakBukti 
         Height          =   450
         Left            =   5340
         TabIndex        =   35
         Top             =   255
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   794
         BTYPE           =   5
         TX              =   "Cetak No Bukti"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "_frmSPer.frx":0BC0
         PICN            =   "_frmSPer.frx":0BDC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   10
         Left            =   1950
         TabIndex        =   20
         Top             =   285
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Left            =   1950
         TabIndex        =   22
         Top             =   645
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Index           =   12
         Left            =   1950
         TabIndex        =   24
         Top             =   1005
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Left            =   1950
         TabIndex        =   26
         Top             =   1365
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Left            =   1950
         TabIndex        =   28
         Top             =   1725
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Height          =   615
         Index           =   15
         Left            =   1950
         TabIndex        =   30
         Top             =   2085
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   1085
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Height          =   615
         Index           =   16
         Left            =   1950
         TabIndex        =   32
         Top             =   2745
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   1085
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Left            =   1950
         TabIndex        =   34
         Top             =   3405
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Alignment       =   1
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   8421504
         Locked          =   -1  'True
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
         Caption         =   "Sisa Angsuran"
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
         Left            =   165
         TabIndex        =   33
         Top             =   3450
         Width           =   1200
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
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
         Left            =   165
         TabIndex        =   31
         Top             =   2790
         Width           =   435
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
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
         TabIndex        =   29
         Top             =   2175
         Width           =   1065
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angsuran Ke 10"
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
         Left            =   165
         TabIndex        =   27
         Top             =   1815
         Width           =   1290
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP/Uang Muka"
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
         Left            =   165
         TabIndex        =   25
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Biaya Administrasi"
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
         Left            =   165
         TabIndex        =   23
         Top             =   1065
         Width           =   1530
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   165
         TabIndex        =   21
         Top             =   735
         Width           =   645
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Bukti"
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
         Left            =   165
         TabIndex        =   19
         Top             =   360
         Width           =   675
      End
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   9
      Left            =   4965
      TabIndex        =   5
      Top             =   1155
      Width           =   1545
      _ExtentX        =   2725
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
      Index           =   7
      Left            =   1650
      TabIndex        =   17
      Top             =   3360
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Icon            =   "_frmSPer.frx":0F76
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   -1  'True
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   8
      Left            =   3000
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3360
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   556
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
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
      Left            =   1650
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2985
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   556
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
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
      Index           =   2
      Left            =   1650
      TabIndex        =   7
      Top             =   1515
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Icon            =   "_frmSPer.frx":13C4
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   -1  'True
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   690
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":1812
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":1BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":1F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":22E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":267A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":2A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":2DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":34E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":387C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":3C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":3FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":434A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":46E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":4C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":5218
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":55B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":594C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":5CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmSPer.frx":84C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   7755
      _ExtentX        =   13679
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
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Option"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   37
      Top             =   7845
      Width           =   7755
      _ExtentX        =   13679
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   0
      Left            =   1650
      TabIndex        =   1
      Top             =   795
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   556
      Icon            =   "_frmSPer.frx":8A5E
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   -1  'True
      BorderColor     =   33023
      Locked          =   -1  'True
      AutoTab         =   -1  'True
      FocusBackColor  =   12640511
      FocusForeColor  =   8388736
      FocusBackMainColor=   8438015
      FocusBorderColor=   33023
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      ForeColor       =   0
      MaxLength       =   20
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   1
      Left            =   1650
      TabIndex        =   3
      Top             =   1155
      Width           =   1545
      _ExtentX        =   2725
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
      Index           =   4
      Left            =   1650
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2265
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   556
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
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
      Index           =   5
      Left            =   1650
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2610
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   556
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
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
      Index           =   3
      Left            =   1650
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1890
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   556
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
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
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Angsuran"
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
      Left            =   3345
      TabIndex        =   4
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   5
      X1              =   -285
      X2              =   19210
      Y1              =   7755
      Y2              =   7755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   -285
      X2              =   19210
      Y1              =   7740
      Y2              =   7740
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disetujui Oleh"
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
      Left            =   210
      TabIndex        =   16
      Top             =   3420
      Width           =   1140
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
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
      Left            =   210
      TabIndex        =   14
      Top             =   3045
      Width           =   945
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelanggan"
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
      Left            =   210
      TabIndex        =   8
      Top             =   1965
      Width           =   855
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
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
      Left            =   210
      TabIndex        =   10
      Top             =   2310
      Width           =   570
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kota"
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
      Left            =   210
      TabIndex        =   12
      Top             =   2670
      Width           =   360
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Pemohon"
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
      Left            =   210
      TabIndex        =   6
      Top             =   1575
      Width           =   1050
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No SPSB"
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
      Left            =   210
      TabIndex        =   0
      Top             =   855
      Width           =   675
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   210
      TabIndex        =   2
      Top             =   1215
      Width           =   645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   -45
      X2              =   19450
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   -45
      X2              =   19450
      Y1              =   585
      Y2              =   585
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurRec As New ADODB.Recordset
Dim hBtn As MSComctlLib.Button
Sub ShowDataPSPB(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT trn_Permohonan_Head.*, trn_Permohonan_Detail.[No Barang], trn_Permohonan_Detail.[Kode Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type, mst_Barang.Satuan, trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Awal Angsuran], trn_Permohonan_Detail.[No Seri], trn_Permohonan_Detail.Keterangan As Ket " & _
                       ", [trn_Permohonan_Detail]![Harga Kredit]*[trn_Permohonan_Detail]![Qty] AS Total FROM mst_Barang RIGHT JOIN (trn_Permohonan_Head LEFT JOIN trn_Permohonan_Detail ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang] Where (((trn_Permohonan_Head.[No Permohonan]) = '" & hKey(0) & "')) ORDER BY trn_Permohonan_Detail.[No Barang];")

If hErr = "" Then
   If Not rc.EOF Then
      txtFields(13).Text = fNum(NotNull(rc("Uang Muka").Value), True)
      txtFields(12).Text = fNum(NotNull(rc("Biaya Adm").Value), True)
      
      Dim NamaBarang As String, Merk As String
      Dim Ang10 As Currency, Total As Currency
      While Not rc.EOF
            NamaBarang = NamaBarang & NotNull(rc("Nama Barang").Value) & " & "
            Merk = Merk & NotNull(rc("Merk").Value) & "/" & NotNull(rc("Type").Value) & " & "
            'GridMe.TextMatrix(pos, 5) = fNum(NotNull(rc("Harga Kredit").Value))
            'GridMe.TextMatrix(pos, 6) = NotNull(rc("Qty").Value)
            'GridMe.TextMatrix(pos, 7) = NotNull(rc("Lama Angsuran").Value)
            'GridMe.TextMatrix(pos, 8) = NotNull(rc("Jenis Angsuran").Value)
            Ang10 = Ang10 + (Val(NotNull(rc("Jumlah Angsuran").Value)) * (Val(NotNull(rc("Awal Angsuran").Value))))
            'GridMe.TextMatrix(pos, 11) = NotNull(rc("No Seri").Value)
            'GridMe.TextMatrix(pos, 12) = NotNull(rc("Ket").Value)
            Total = Total + Val(NotNull(rc("Total").Value))
            rc.MoveNext
      Wend
      txtFields(14) = fNum(Ang10, True)
      txtFields(15) = IIf(Right(NamaBarang, 2) = "& ", Mid(NamaBarang, 1, Len(NamaBarang) - 2), NamaBarang)
      txtFields(16) = IIf(Right(Merk, 2) = "& ", Mid(Merk, 1, Len(Merk) - 2), Merk)
      txtFields(17) = fNum(Total - (Val(rNum(txtFields(13).Text)) + Ang10), True)
      txtFields(17).Tag = fNum(Val(rNum(txtFields(13).Text)) + Ang10 + rNum(txtFields(12)), True)
   End If
   rc.Close
End If
End Sub

Sub ShowAllData(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT trn_Perjanjian.[No Bukti],trn_Perjanjian.[Tgl Bukti],trn_Perjanjian.[Tgl Mulai],trn_Perjanjian.[No Perjanjian], trn_Perjanjian.[No Permohonan], trn_Perjanjian.Keterangan, trn_Perjanjian.[Kode Pegawai], trn_Perjanjian.Status, trn_Perjanjian.[Tgl Perjanjian] From trn_Perjanjian WHERE trn_Perjanjian.[No Perjanjian]='" & hKey(0) & "';")

If hErr = "" Then
   If Not rc.EOF Then
      txtFields(0).Text = NotNull(rc("No Perjanjian").Value)
      txtFields(0).Tag = NotNull(rc("No Perjanjian").Value)
      
      txtFields(1).Text = NotNull(rc("Tgl Perjanjian").Value)
      txtFields(6).Text = NotNull(rc("Keterangan").Value)
      txtFields(9).Text = NotNull(rc("Tgl Mulai").Value)
      txtFields(10).Text = NotNull(rc("No Bukti").Value)
      txtFields(11).Text = NotNull(rc("Tgl Bukti").Value)
      ShowPelanggan NotNull(rc("No Permohonan").Value) & "|"
      ShowInspektur NotNull(rc("Kode Pegawai").Value) & "|"
                  
   Else
        ClearControl Me
        Me.Caption = Replace(Me.Caption, "*", "")
        txtFields(0).Tag = ""
        Me.Caption = Me.Caption & Me.Tag
   End If
   rc.Close
End If
End Sub
Sub HapusData(hKey As String)
Dim hErr As String
hErr = FindRecord("SELECT  [No Perjanjian] From trn_Perjanjian WHERE  [No Perjanjian] ='" & hKey & "';")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Perjanjian Sewa Beli Barang?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") = False Then
      GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE From trn_Perjanjian WHERE [No Perjanjian]='" & hKey & "';")
         If hErr = "" Then
            ClearControl Me
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = ""
         Else
            ShowDlgMsg Me, "Proses penghapusan data gagal!", vbOK, hErr, True, False
         End If
      End If
   End If
ElseIf hErr = "0" Then
    ShowDlgMsg Me, "Tidak ada data yang akan dihapus", vbOK, , True, False
End If
End Sub

Sub ShowInspektur(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "Select * from mst_Pegawai WHERE [Kode Divisi]='" & GetDivisi(1) & "' AND [Kode Pegawai]='" & hKey(0) & "'")
If hErr = "" Then
   If Not rc.EOF Then
    txtFields(7).Text = NotNull(rc("Kode Pegawai"))
    txtFields(8).Text = NotNull(rc("Nama Pegawai"))
   Else
    txtFields(7).Text = ""
    txtFields(8).Text = ""
   End If
Else
    txtFields(7).Text = ""
    txtFields(8).Text = ""
End If
rc.Close
End Sub

Sub ShowPelanggan(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")

hErr = SelectQuery(rc, "SELECT trn_Permohonan_Head.[No Permohonan], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota " & _
                       "FROM mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]  " & _
                       "WHERE (((trn_Permohonan_Head.[No Permohonan])='" & hKey(0) & "'));")
       
If hErr = "" Then
   If Not rc.EOF Then
    txtFields(2).Text = NotNull(rc("No Permohonan"))
    txtFields(3).Text = NotNull(rc("Nama"))
    txtFields(4).Text = NotNull(rc("Alamat"))
    txtFields(5).Text = NotNull(rc("Kota"))
    ShowDataPSPB NotNull(rc("No Permohonan"))
   Else
    txtFields(2).Text = ""
    txtFields(3).Text = ""
    txtFields(4).Text = ""
    txtFields(5).Text = ""
   End If
Else
    txtFields(2).Text = ""
    txtFields(3).Text = ""
    txtFields(4).Text = ""
    txtFields(5).Text = ""
End If
rc.Close
End Sub

Sub SimpanData(nKey As String)
Dim h As String, hErr As String, i As Integer
Dim X
h = FindRecord("SELECT [No Perjanjian] From trn_Perjanjian WHERE [No Perjanjian]='" & nKey & "';")
If h = "0" Then
   hErr = SaveRecord("trn_Perjanjian", Array("No Perjanjian=" & txtFields(0).Text, _
                                                  "No Permohonan=" & txtFields(2).Text, _
                                                  "Keterangan=" & txtFields(6).Text, _
                                                  "Kode Pegawai=" & txtFields(7).Text, _
                                                  "@Tgl Perjanjian=" & txtFields(1).Text, _
                                                  "@Tgl Mulai=" & txtFields(9).Text, _
                                                  "Status=OPEN"))
  If hErr = "" Then
      If CekAktifNo("001") Then txtFields(0).Text = getAutoNo("001", True)
      txtFields(0).Tag = txtFields(0).Text
      Me.Caption = Replace(Me.Caption, "*", "")
      Me.Tag = ""
  Else
     ShowDlgMsg MainMenu, "Error!!!<br><br>Tidak dapat menyimpan data!", vbOK, Error, True, False
  End If
ElseIf h = "1" Then
   If ShowDlgMsg(Me, "Data sudah terdaftar!, update dengan data baru?", vbYesNo, Error, False, True, , , , , Me.name & "_update") = False Then
      GoSub SimpanLabel
   Else
      If SelectMsg = vbYes Then
SimpanLabel:
      
        hErr = UpdateRecord("trn_Perjanjian", Array("No Perjanjian=" & txtFields(0).Text, _
                                                         "No Permohonan=" & txtFields(2).Text, _
                                                         "Keterangan=" & txtFields(6).Text, _
                                                         "Kode Pegawai=" & txtFields(7).Text, _
                                                         "@Tgl Mulai=" & txtFields(9).Text, _
                                                         "@Tgl Perjanjian=" & txtFields(1).Text), " WHERE [No Perjanjian]='" & txtFields(0).Tag & "' ")

        If hErr = "" Then
            Me.Caption = Replace(Me.Caption, "*", "")
            txtFields(0).Tag = txtFields(0).Text
            Me.Tag = ""
        Else
            ShowDlgMsg MainMenu, "Error!!!<br><br>Tidak dapat menyimpan data!", vbOK, Error, True, False
        End If
     End If
   End If
End If
End Sub

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
txtFields(0).Locked = CekAktifNo("001")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
PostFindForm "#" & txtFields(0).Hwnd1
CurRec.Close
Set CurRec = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.index
       Case 1
           If CekUser("07", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            ClearControl Me
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            If CekAktifNo("001") Then
            txtFields(0).Text = getAutoNo("001")
            txtFields(1).SetFocus
            Else
            txtFields(0).SetFocus
            End If
           End If
       Case 2
           If CekUser("07", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("07", "S") = False Then
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
           If CekUser("07", "P") = False Then
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
               ShowAllData NotNull(CurRec("No Perjanjian")) & "|"
            'End If
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               If CurRec.BOF Then
                  CurRec.MoveNext
               End If
               ShowAllData NotNull(CurRec("No Perjanjian")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               If CurRec.EOF Then
                  CurRec.MovePrevious
               End If
               ShowAllData NotNull(CurRec("No Perjanjian")) & "|"
            End If
       Case 4
            'If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowAllData NotNull(CurRec("No Perjanjian")) & "|"
            'End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowAllData NotNull(CurRec("No Perjanjian")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [No Perjanjian] From trn_Perjanjian ORDER BY [No Perjanjian]"
Return
End Sub

Private Sub txtFields_DownButtonClick(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            ShowFindForm "SELECT trn_Perjanjian.[No Perjanjian], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, mst_Pelanggan.[Kode Pos], trn_Perjanjian.[Tgl Perjanjian] " & _
                         "FROM mst_Pelanggan RIGHT JOIN (trn_Permohonan_Head RIGHT JOIN trn_Perjanjian ON trn_Permohonan_Head.[No Permohonan] = trn_Perjanjian.[No Permohonan]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan] " & _
                         " <!where> ", "#" & txtFields(index).Hwnd1, Me, "ShowAllData"
       
       Case 2
            ShowFindForm "SELECT trn_Permohonan_Head.[No Permohonan], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota " & _
                         "FROM mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan] " & _
                         " <!where> ", "#" & txtFields(index).Hwnd1, Me, "ShowPelanggan"
       Case 7
            ShowFindForm "SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai],mst_Divisi.Jabatan " & _
                         "FROM mst_Divisi INNER JOIN mst_Pegawai ON mst_Divisi.[Kode Divisi] = mst_Pegawai.[Kode Divisi] " & _
                         " <!where> ;", "#" & txtFields(index).Hwnd1, Me, "ShowInspektur", " mst_Pegawai.[Kode Divisi]='" & GetDivisi(1) & "' AND "

End Select

End Sub

Private Sub txtFields_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Select Case index
          Case 0: ShowAllData txtFields(index).Text & "|"
          Case 2: ShowPelanggan txtFields(index).Text & "|"
          Case 7: ShowInspektur txtFields(index).Text & "|"
   End Select
End If
End Sub

Private Sub vbCetakBukti_Click()
On Error Resume Next
If CekUser("07", "P") = False Then
   ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
Else
    If txtFields(10).Text = "" Then Exit Sub
    Dim hDlg As Boolean
    hDlg = ShowDlgMsg(Me, "Cetak No Pembayaran?", vbYesNo, , False, True, , , , , "confirm_" & Me.name)
    If hDlg Then If SelectMsg = vbNo Then Exit Sub
    
    Dim hErr As String, rc As New ADODB.Recordset
    hErr = SelectQuery(rc, "SELECT trn_Perjanjian.[No Bukti],trn_Perjanjian.[No Perjanjian], trn_Perjanjian.[No Permohonan] " & _
                    "From trn_Perjanjian WHERE (((trn_Perjanjian.[No Perjanjian])='" & txtFields(0) & "') AND ((trn_Perjanjian.[No Permohonan])='" & txtFields(2) & "'));")
    If hErr = "" Then
        If Not rc.EOF Then
           If NotNull(rc("No Bukti")) = "" Then
              txtFields(10).Text = getAutoNo("010", True)
              txtFields(11).Text = fDate(Date)
              hErr = UpdateRecord("trn_Perjanjian", Array("No Bukti=" & txtFields(10).Text, _
                                 "@Tgl Bukti=" & txtFields(11).Text), " WHERE [No Perjanjian]='" & txtFields(0).Tag & "' ")
           End If
               If txtFields(10).Text <> "" Then
                    Dim hText As String, hAlamat As String
                    'hText = txtPrint
                    hAlamat = txtFields(4).Text
'                    hText = Replace(hText, String(15, "a"), AddSpace(txtFields(10), 15))
'                    hText = Replace(hText, String(62, "b"), AddSpace(txtFields(3).Text, 62))
'                    hText = Replace(hText, String(62, "c"), AddSpace(hAlamat, 62))
'                    hText = Replace(hText, String(62, "d"), AddSpace(Mid(hAlamat, 63) & " " & txtFields(5), 62))
'                    hText = Replace(hText, String(23, "e"), AddSpace(txtFields(17).Tag & ";-", 23)) 'Uang sebesar
'                    hText = Replace(hText, String(27, "f"), AddSpace(txtFields(12) & ";-", 27)) 'biaya Adminstrasi
'                    hText = Replace(hText, String(27, "g"), AddSpace(txtFields(13) & ";-", 27)) 'Uang Muka
'                    hText = Replace(hText, String(27, "h"), AddSpace(txtFields(14) & ";-", 27)) 'Angsuran Ke 10
'                    hText = Replace(hText, String(27, "i"), AddSpace(txtFields(17).Tag & ";-", 27)) 'Total
'                    hText = Replace(hText, String(87, "j"), AddSpace(txtFields(15), 87)) 'Nama Barang
'                    hText = Replace(hText, String(87, "k"), AddSpace(txtFields(16), 87)) 'Merk
'                    hText = Replace(hText, String(61, "l"), AddSpace(txtFields(0), 61)) 'no SPSB
'                    hText = Replace(hText, String(30, "m"), AddSpace(txtFields(17).Text & ";-", 30)) 'Sisa Angsuran
'                    hText = Replace(hText, String(73, "n"), AddSpace(Terbilang(rNum(txtFields(17).Text)) & "Rupiah", 73))   'Terbilang
'                    hText = Replace(hText, String(16, "o"), AddSpace(txtFields(11), 16))
'                    PrintText hText

                    
                    Load rpt_buktibayar
                    With rpt_buktibayar
                         .Label28.Caption = Replace(.Label28.Caption, String(15, "a"), AddSpace(txtFields(10), 15))
                         .Label5.Caption = Replace(.Label5.Caption, String(62, "b"), AddSpace(txtFields(3).Text, 62))
                         .Label6.Caption = Replace(.Label6.Caption, String(62, "c"), AddSpace(hAlamat, 62))
                         If Mid(Trim(hAlamat), 63) = "" Then
                            .Label7.Caption = Replace(.Label7.Caption, String(62, "d"), AddSpace(Mid(Trim(hAlamat), 63) & Trim(txtFields(5)), 62))
                         Else
                            .Label7.Caption = Replace(.Label7.Caption, String(62, "d"), AddSpace(Mid(Trim(hAlamat), 63) & " " & Trim(txtFields(5)), 62))
                         End If
                         .Label8.Caption = Replace(.Label8.Caption, String(23, "e"), AddSpace(txtFields(17).Tag & "", 23))  'Uang sebesar
                         .Label9.Caption = Replace(.Label9.Caption, String(27, "f"), AddSpace(txtFields(12) & "", 27, True))  'biaya Adminstrasi
                         .Label10.Caption = Replace(.Label10.Caption, String(27, "g"), AddSpace(txtFields(13) & "", 27, True))  'Uang Muka
                         .Label11.Caption = Replace(.Label11.Caption, String(27, "h"), AddSpace(txtFields(14) & "", 27, True))  'Angsuran Ke 10
                         .Label13.Caption = Replace(.Label13.Caption, String(27, "i"), AddSpace(txtFields(17).Tag & "", 27, True))  'Total
                         .Label15.Caption = Replace(.Label15.Caption, String(87, "j"), AddSpace(txtFields(15), 87))  'Nama Barang
                         .Label16.Caption = Replace(.Label16.Caption, String(87, "k"), AddSpace(txtFields(16), 87))  'Merk
                         .Label18.Caption = Replace(.Label18.Caption, String(61, "l"), AddSpace(txtFields(0), 61))  'no SPSB
                         .Label19.Caption = Replace(.Label19.Caption, String(30, "m"), AddSpace(txtFields(17).Text & "", 30))  'Sisa Angsuran
                         .Label21.Caption = Replace(.Label21.Caption, String(73, "n"), AddSpace(Terbilang(rNum(txtFields(17).Text)) & "Rupiah", 73))    'Terbilang
                         .Label23.Caption = Replace(.Label23.Caption, String(16, "o"), AddSpace(txtFields(11), 16))
                          
                          .PrintReport False
                    End With
                    Unload rpt_buktibayar
                    
                    hText = ""
               Else
                    ShowDlgMsg Me, "Tidak ada data faktur yang akan dicetak!", vbOK, , True, False
               End If
           
        Else
           ShowDlgMsg Me, "Data perjanjian belum tersimpan, proses pencetakan Bukti Pembayaran Dibatalkan", vbOK, , True, False
        End If
    End If
End If
End Sub
