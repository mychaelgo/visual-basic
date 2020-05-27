VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Angsuran - Sewa Beli Barang"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_trn_Angsuran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11670
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   13
      Left            =   10680
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   720
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
      AutoTab         =   -1  'True
      FontFormat      =   2
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
   Begin SysInfo_Nardhika.vbButton vbButton4 
      Height          =   375
      Left            =   150
      TabIndex        =   61
      Top             =   6750
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "&Kalkulasi Tanggal"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "_frm_trn_Angsuran.frx":038A
      PICN            =   "_frm_trn_Angsuran.frx":06A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfo_Nardhika.vbTextBoxMulti txtFieldsLine 
      Height          =   750
      Left            =   3930
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1323
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
      MultiLine       =   -1  'True
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   19
      Left            =   1845
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   16
      Left            =   1845
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   14
      Left            =   1845
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
      AutoTab         =   -1  'True
      FontFormat      =   1
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
      Index           =   12
      Left            =   1845
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   11
      Left            =   9510
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   10
      Left            =   9510
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   9
      Left            =   9510
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   7
      Left            =   10320
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1050
      _ExtentX        =   1852
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   6
      Left            =   9510
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1080
      Width           =   750
      _ExtentX        =   1323
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   3
      Left            =   9510
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Alignment       =   1
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   8
      Left            =   5610
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
      AutoTab         =   -1  'True
      FontFormat      =   2
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
      Left            =   5610
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1860
      _ExtentX        =   3281
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   4
      Left            =   5610
      TabIndex        =   15
      Top             =   2160
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Icon            =   "_frm_trn_Angsuran.frx":0A3E
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
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   23
      Left            =   9630
      TabIndex        =   40
      Top             =   6885
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
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
      Left            =   9630
      TabIndex        =   38
      Top             =   6585
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
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
      Left            =   1845
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8610
      Top             =   7590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":0E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":1226
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":15C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":195A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":1CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":208E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":2428
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":27C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":2B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":2EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":3290
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":362A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":39C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":3D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":4892
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":4C2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":4FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Angsuran.frx":5360
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1005
      ButtonWidth     =   1244
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
            Caption         =   "Print"
            Object.ToolTipText     =   "Cetak Bukti Angsuran"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageIndex      =   6
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
      TabIndex        =   42
      Top             =   7410
      Width           =   11670
      _ExtentX        =   20585
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
      Left            =   1845
      TabIndex        =   1
      Top             =   720
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Icon            =   "_frm_trn_Angsuran.frx":56FA
      Alignment       =   1
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
      MaxLength       =   17
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   1
      Left            =   1845
      TabIndex        =   3
      Top             =   1080
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      Icon            =   "_frm_trn_Angsuran.frx":5B48
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3195
      Left            =   150
      TabIndex        =   43
      Top             =   3480
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   5636
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   $"_frm_trn_Angsuran.frx":5F96
      TabPicture(0)   =   "_frm_trn_Angsuran.frx":602F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "picBrg"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "GridMe"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Toolbar3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCell"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.TextBox txtCell 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4440
         TabIndex        =   56
         Top             =   2865
         Width           =   630
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hari Sabtu"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2070
         TabIndex        =   36
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hari Libur"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   975
         TabIndex        =   35
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Left            =   5205
         TabIndex        =   53
         Top             =   2820
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   16
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   17
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   18
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox GridMe 
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   2700
         Left            =   105
         ScaleHeight     =   2640
         ScaleWidth      =   11070
         TabIndex        =   34
         Top             =   90
         Width           =   11130
      End
      Begin VB.PictureBox picBrg 
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         Height          =   2400
         Left            =   135
         ScaleHeight     =   2400
         ScaleWidth      =   10740
         TabIndex        =   44
         Top             =   105
         Visible         =   0   'False
         Width           =   10740
         Begin VB.OptionButton Option2 
            BackColor       =   &H00F9F9F9&
            Caption         =   "Berisi Kata"
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
            Left            =   9360
            TabIndex        =   47
            Top             =   2100
            Width           =   1260
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F9F9F9&
            Caption         =   "Kata Awalan"
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
            Left            =   9360
            TabIndex        =   46
            Top             =   1860
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.TextBox txtCari 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1365
            TabIndex        =   45
            Top             =   60
            Width           =   7800
         End
         Begin SysInfo_Nardhika.vbButton vbButton1 
            Height          =   360
            Left            =   9360
            TabIndex        =   48
            Top             =   135
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "&Cari"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "_frm_trn_Angsuran.frx":604B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.PictureBox GridFind 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   2025
            Left            =   0
            ScaleHeight     =   2025
            ScaleWidth      =   9255
            TabIndex        =   49
            Top             =   405
            Width           =   9255
         End
         Begin SysInfo_Nardhika.vbButton vbButton2 
            Height          =   360
            Left            =   9360
            TabIndex        =   50
            Top             =   555
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "&Pilih"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "_frm_trn_Angsuran.frx":6067
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SysInfo_Nardhika.vbButton vbButton3 
            Height          =   360
            Left            =   9360
            TabIndex        =   51
            Top             =   975
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "&Batal"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "_frm_trn_Angsuran.frx":6083
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Line Line2 
            X1              =   9330
            X2              =   10620
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label Label1 
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
            Left            =   135
            TabIndex        =   52
            Top             =   105
            Width           =   1065
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000004&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   510
            Left            =   -30
            Top             =   -105
            Width           =   9270
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generating date, wait..."
         Height          =   195
         Left            =   270
         TabIndex        =   60
         Top             =   315
         Width           =   1755
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   240
         Left            =   4425
         Top             =   2850
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agsuran Ke :"
         Height          =   195
         Left            =   3465
         TabIndex        =   55
         Top             =   2865
         Width           =   930
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         X1              =   3345
         X2              =   3345
         Y1              =   2850
         Y2              =   3090
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Termasuk :"
         Height          =   195
         Left            =   105
         TabIndex        =   54
         Top             =   2865
         Width           =   795
      End
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   15
      Left            =   4890
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   945
      Width           =   2580
      _ExtentX        =   4551
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
   Begin SysInfo_Nardhika.vbTextBoxMulti txtFieldsLine2 
      Height          =   750
      Left            =   8955
      TabIndex        =   58
      Top             =   2535
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1323
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
      MultiLine       =   -1  'True
      Text            =   ""
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Awal"
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
      Left            =   10185
      TabIndex        =   63
      Top             =   780
      Width           =   405
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
      Index           =   19
      Left            =   7875
      TabIndex        =   59
      Top             =   2550
      Width           =   945
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inspektur"
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
      Left            =   195
      TabIndex        =   12
      Top             =   2925
      Width           =   810
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salesmen"
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
      Left            =   195
      TabIndex        =   10
      Top             =   2565
      Width           =   825
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Angsuran"
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
      Left            =   195
      TabIndex        =   8
      Top             =   2205
      Width           =   1110
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
      Index           =   11
      Left            =   195
      TabIndex        =   6
      Top             =   1845
      Width           =   675
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
      Index           =   13
      Left            =   4905
      TabIndex        =   31
      Top             =   705
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
      Index           =   12
      Left            =   3975
      TabIndex        =   33
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Seri Barang"
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
      Left            =   7860
      TabIndex        =   29
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jatuh Tempo"
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
      Left            =   7860
      TabIndex        =   27
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Angsuran"
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
      Left            =   7860
      TabIndex        =   25
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lama Angsuran"
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
      Left            =   7860
      TabIndex        =   22
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Left            =   7860
      TabIndex        =   20
      Top             =   765
      Width           =   675
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Kredit"
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
      Left            =   3975
      TabIndex        =   18
      Top             =   2955
      Width           =   1005
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
      Index           =   8
      Left            =   3975
      TabIndex        =   16
      Top             =   2595
      Width           =   1065
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Barang"
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
      Left            =   3975
      TabIndex        =   14
      Top             =   2220
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   0
      X2              =   19495
      Y1              =   7335
      Y2              =   7335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   0
      X2              =   19495
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   90
      X2              =   19585
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   90
      X2              =   19585
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Angsuran"
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
      Left            =   195
      TabIndex        =   0
      Top             =   765
      Width           =   1065
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Permohonan"
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
      Left            =   195
      TabIndex        =   2
      Top             =   1140
      Width           =   1320
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Permohonan"
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
      Left            =   195
      TabIndex        =   4
      Top             =   1470
      Width           =   1365
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Angsuran"
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
      Left            =   7995
      TabIndex        =   37
      Top             =   6675
      Width           =   1260
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sisa Tagihan"
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
      Left            =   7995
      TabIndex        =   39
      Top             =   6960
      Width           =   1035
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColDelete As New Collection
Dim CurRec As New ADODB.Recordset
Dim hBtn As MSComctlLib.Button


Sub ShowAllData(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT trn_Angsuran_Head.[No Angsuran], trn_Angsuran_Head.[No Permohonan], trn_Angsuran_Head.[No Barang], trn_Angsuran_Head.[Kode Pegawai], trn_Angsuran_Head.Keterangan, trn_Angsuran_Head.Status, trn_Angsuran_Head.[Hari Libur], trn_Angsuran_Head.[Hari Sabtu] From trn_Angsuran_Head WHERE (((trn_Angsuran_Head.[No Angsuran])='" & hKey(0) & "'));")
If hErr = "" Then
   If Not rc.EOF Then
      txtFields(0).Text = NotNull(rc("No Angsuran"))
      txtFields(0).Tag = NotNull(rc("No Angsuran"))
      ShowDataPermohonan NotNull(rc("No Permohonan")) & "|"
      ShowDataNoBarang NotNull(rc("No Barang")) & "|", , True
      txtFieldsLine2.Text = NotNull(rc("Keterangan"))
      Check1.Value = NotNull(rc("Hari Libur"))
      Check2.Value = NotNull(rc("Hari Sabtu"))
      
      rc.Close
      hErr = SelectQuery(rc, "SELECT trn_Angsuran_Detail.[No Angsuran], trn_Angsuran_Detail.[No Bayar], trn_Angsuran_Detail.[Tgl Bayar], trn_Angsuran_Detail.[Tgl Dibayar], trn_Angsuran_Detail.[Jumlah Bayar], trn_Angsuran_Detail.Keterangan, trn_Angsuran_Detail.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], trn_Angsuran_Detail.[Angsuran Ke] " & _
                          "FROM mst_Pegawai RIGHT JOIN trn_Angsuran_Detail ON mst_Pegawai.[Kode Pegawai] = trn_Angsuran_Detail.[Kode Pegawai] Where (((trn_Angsuran_Detail.[No Angsuran]) = '" & txtFields(0).Tag & "')) ORDER BY trn_Angsuran_Detail.[No Bayar];")
                                  
        If hErr = "" Then
           If Not rc.EOF Then
              InitGrid
              Dim pos As Integer
              GridMe.Visible = False
              While Not rc.EOF
                 pos = Val(NotNull(rc("No Bayar")))
                 GridMe.AddItem pos
                 GridMe.TextMatrix(pos + 1, 1) = Format(NotNull(rc("Tgl Bayar")), "DD-MM-YYYY")
                 GridMe.TextMatrix(pos + 1, 2) = Format(NotNull(rc("Tgl Dibayar")), "DD-MM-YYYY")
                 GridMe.TextMatrix(pos + 1, 3) = fNum(NotNull(rc("Jumlah Bayar")))
                 GridMe.TextMatrix(pos + 1, 7) = NotNull(rc("Kode Pegawai"))
                 GridMe.TextMatrix(pos + 1, 5) = NotNull(rc("Nama Pegawai"))
                 GridMe.TextMatrix(pos + 1, 6) = NotNull(rc("Keterangan"))
                 GridMe.TextMatrix(pos + 1, 8) = NotNull(rc("No Bayar"))
                 If NotNull(rc("Tgl Dibayar")) <> "" Then
                 If pos = 1 Then
                    GridMe.TextMatrix(pos + 1, 4) = (GridMe.TextMatrix(pos + 1, 3))
                 Else
                    GridMe.TextMatrix(pos + 1, 4) = fNum(rNum(NotNull(rc("Jumlah Bayar"))) + rNum(GridMe.TextMatrix(pos, 4)))
                 End If
                 End If
                 rc.MoveNext
                 DoEvents
              Wend
              GridMe.Visible = True
           End If
        End If
        Akumulasi
   End If
End If
End Sub

Sub HapusData(hKey As String)
On Error Resume Next
Dim hErr As String
hErr = FindRecord("SELECT trn_Angsuran_Head.[No Angsuran] From trn_Angsuran_Head WHERE trn_Angsuran_Head.[No Angsuran]='" & hKey & "';")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Angsuran Sewa Beli Barang?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") Then
      If SelectMsg = vbYes Then GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE From trn_Angsuran_Head WHERE trn_Angsuran_Head.[No Angsuran]='" & hKey & "';")
         If hErr = "" Then
            ClearControl Me
            InitGrid
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = ""
            Set ColDelete = New Collection
         Else
            ShowDlgMsg Me, "Proses penghapusan data gagal!", vbOK, hErr, True, False
         End If
      End If
   End If
ElseIf hErr = "0" Then
    ShowDlgMsg Me, "Tidak ada data yang akan dihapus", vbOK, , True, False
End If
End Sub

Sub SimpanData(nKey As String)
On Error Resume Next
Dim h As String, hErr As String, i As Integer
Dim X
h = FindRecord("SELECT trn_Angsuran_Head.[No Angsuran] From trn_Angsuran_Head WHERE trn_Angsuran_Head.[No Angsuran]='" & nKey & "';")
If h = "0" Then
   hErr = SaveRecord("trn_Angsuran_Head", Array("No Angsuran=" & txtFields(0).Text, _
                                                "No Permohonan=" & txtFields(1).Text, _
                                                "No Barang=" & txtFields(4).Text, _
                                                "Keterangan=" & txtFieldsLine2.Text, _
                                                "Status=open", _
                                                "Hari Libur=" & Check1.Value, _
                                                "Hari Sabtu=" & Check2.Value))
  If hErr = "" Then
      
      For i = 1 To ColDelete.Count
          X = Split(ColDelete(i), Chr(255))
          ExecQuery "DELETE FROM trn_Angsuran_Head WHERE ([No Angsuran]= '" & AllowChar(CStr(X(0))) & "') AND ([No Bayar]= '" & AllowChar(CStr(X(1))) & "')"
      Next i

     txtFields(0).Tag = txtFields(0).Text
     For i = 2 To GridMe.Rows - 1
        If GridMe.TextMatrix(i, 1) <> "" Then
           hErr = SaveRecord("trn_Angsuran_Detail", Array("No Angsuran=" & txtFields(0), _
                                                            "No Bayar=" & Format(i - 1, "000#"), _
                                                            "@Tgl Bayar=" & GridMe.TextMatrix(i, 1), _
                                                            "@Tgl Dibayar=" & GridMe.TextMatrix(i, 2), _
                                                            "$Jumlah Bayar=" & rNum(GridMe.TextMatrix(i, 3)), _
                                                            "Keterangan=" & GridMe.TextMatrix(i, 6), _
                                                            "Kode Pegawai=" & GridMe.TextMatrix(i, 7), _
                                                            "Angsuran Ke=" & i - 1))
           
           GridMe.TextMatrix(i, 8) = Format(i - 1, "000#")
           If hErr = "" Then
                'Error Catch here!
           End If
        End If
     Next i
     If CekAktifNo("003") Then txtFields(0).Text = getAutoNo("003", True)
     txtFields(0).Tag = txtFields(0).Text
     Me.Caption = Replace(Me.Caption, "*", "")
     Set ColDelete = New Collection
     Me.Tag = ""
  Else
     ShowDlgMsg MainMenu, "Error!!!<br><br>Tidak dapat menyimpan data!", vbOK, hErr, True, False
  End If
ElseIf h = "1" Then
   If ShowDlgMsg(Me, "Data sudah terdaftar!, update dengan data baru?", vbYesNo, Error, False, True, , , , , Me.name & "_update") Then
      GoSub SimpanLabel
   Else
      If SelectMsg = vbYes Then
SimpanLabel:
      
      For i = 1 To ColDelete.Count
          X = Split(ColDelete(i), Chr(255))
          ExecQuery "DELETE FROM trn_Angsuran_Detail WHERE ([No Angsuran]= '" & AllowChar(CStr(X(0))) & "') AND ([No Bayar]= '" & AllowChar(CStr(X(1))) & "')"
      Next i

        hErr = UpdateRecord("trn_Angsuran_Head", Array("No Angsuran=" & txtFields(0).Text, _
                                                     "No Permohonan=" & txtFields(1).Text, _
                                                     "No Barang=" & txtFields(4).Text, _
                                                     "Keterangan=" & txtFieldsLine2.Text, _
                                                     "Status=open", _
                                                     "Hari Libur=" & Check1.Value, _
                                                     "Hari Sabtu=" & Check2.Value), " WHERE [No Angsuran]='" & txtFields(0).Text & "' ")
        If hErr = "" Then
           For i = 2 To GridMe.Rows - 1
              If GridMe.TextMatrix(i, 1) <> "" Then
                 If GridMe.TextMatrix(i, 8) = "" Then
                    hErr = SaveRecord("trn_Angsuran_Detail", Array("No Angsuran=" & txtFields(0), _
                                                                   "No Bayar=" & Format(i - 1, "000#"), _
                                                                   "@Tgl Bayar=" & GridMe.TextMatrix(i, 1), _
                                                                   "@Tgl Dibayar=" & GridMe.TextMatrix(i, 2), _
                                                                   "$Jumlah Bayar=" & rNum(GridMe.TextMatrix(i, 3)), _
                                                                   "Keterangan=" & GridMe.TextMatrix(i, 6), _
                                                                   "Kode Pegawai=" & GridMe.TextMatrix(i, 7), _
                                                                   "Angsuran Ke=" & i - 1))
                      
                      GridMe.TextMatrix(i, 8) = Format(i - 1, "000#")
                      If hErr = "" Then
                           'Error Catch here!
                      End If
                  Else '!>
                    If GridMe.Cell(flexcpBackColor, i, 1, , 6) = &HE3E3FF Then
                    hErr = UpdateRecord("trn_Angsuran_Detail", Array("No Angsuran=" & txtFields(0), _
                                                                     "No Bayar=" & Format(i - 1, "000#"), _
                                                                     "@Tgl Bayar=" & GridMe.TextMatrix(i, 1), _
                                                                     "@Tgl Dibayar=" & GridMe.TextMatrix(i, 2), _
                                                                     "$Jumlah Bayar=" & rNum(GridMe.TextMatrix(i, 3)), _
                                                                     "Keterangan=" & GridMe.TextMatrix(i, 6), _
                                                                     "Kode Pegawai=" & GridMe.TextMatrix(i, 7), _
                                                                     "Angsuran Ke=" & i - 1), " WHERE [No Angsuran]='" & txtFields(0).Tag & "' AND [No Bayar]='" & GridMe.TextMatrix(i, 8) & "'")
                      
                      GridMe.TextMatrix(i, 8) = Format(i - 1, "000#")
                        If hErr = "" Then
                           'Error Catch here!
                        End If
                       GridMe.Cell(flexcpBackColor, i, 1, , 6) = vbWhite
                    End If
                 End If
              End If
           Next i
     
            txtFields(0).Tag = txtFields(0).Text
            Me.Caption = Replace(Me.Caption, "*", "")
            Set ColDelete = New Collection
            Me.Tag = ""
           
        End If
     End If
   End If
End If
End Sub


Sub ShowDataKorektor(nKey As String, Optional onRow As Integer = -1)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")

Dim lErr As String
lErr = SelectQuery(rc, "SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], mst_Divisi.Jabatan FROM mst_Divisi RIGHT JOIN mst_Pegawai ON mst_Divisi.[Kode Divisi] = mst_Pegawai.[Kode Divisi] " & _
                       "WHERE (((mst_Pegawai.[Kode Pegawai])='" & hKey(0) & "')); ")
If lErr = "" Then
    If Not rc.EOF Then
       If onRow = -1 Then
          GridMe.TextMatrix(GridMe.Row, 7) = NotNull(rc("Kode Pegawai"))
          GridMe.TextMatrix(GridMe.Row, 5) = NotNull(rc("Nama Pegawai"))
       Else
          GridMe.TextMatrix(onRow, 7) = NotNull(rc("Kode Pegawai"))
          GridMe.TextMatrix(onRow, 5) = NotNull(rc("Nama Pegawai"))
       End If
    Else
        If onRow = -1 Then
          GridMe.TextMatrix(GridMe.Row, 7) = ""
          GridMe.TextMatrix(GridMe.Row, 5) = ""
        Else
          GridMe.TextMatrix(onRow, 7) = ""
          GridMe.TextMatrix(onRow, 5) = ""
        End If
    End If
    rc.Close
End If
End Sub

Sub ShowDataNoBarang(nKey As String, Optional nKey2 As String, Optional AllData As Boolean = False)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
Dim hFind As String

hKey = Split(nKey, "|")
If nKey2 = "" Then nKey2 = txtFields(1).Text

If AllData = True Then
   GoSub GetData
Else
    hFind = SelectQuery(rc, "SELECT trn_Angsuran_Head.[No Angsuran], trn_Angsuran_Head.[No Permohonan], trn_Angsuran_Head.[No Barang] From trn_Angsuran_Head WHERE (((trn_Angsuran_Head.[No Permohonan])='" & nKey2 & "') AND ((trn_Angsuran_Head.[No Barang])='" & hKey(0) & "'));")
    If hFind = "" Then
       If Not rc.EOF Then
          ShowAllData NotNull(rc("No Angsuran"))
          rc.Close
       Else
GetData:
            If rc.State = adStateOpen Then rc.Close
            hErr = SelectQuery(rc, "SELECT trn_Permohonan_Detail.[Awal Angsuran],trn_Permohonan_Detail.[No Permohonan], trn_Permohonan_Detail.[No Barang], trn_Permohonan_Detail.[Kode Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type, mst_Barang.Satuan, trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Angsuran JT], trn_Permohonan_Detail.[No Seri], trn_Permohonan_Detail.Keterangan " & _
                                   "FROM mst_Barang RIGHT JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang] " & _
                                   "WHERE (((trn_Permohonan_Detail.[No Permohonan])='" & nKey2 & "') AND ((trn_Permohonan_Detail.[No Barang])='" & hKey(0) & "'));")
            If hErr = "" Then
               If Not rc.EOF Then
                  txtFields(4).Text = NotNull(rc("No Barang").Value)
                  txtFields(5).Text = NotNull(rc("Nama Barang").Value)
                  txtFields(8).Text = fNum(NotNull(rc("Harga Kredit").Value))
                  txtFields(3).Text = NotNull(rc("QTY").Value)
                  txtFields(6).Text = NotNull(rc("Lama Angsuran").Value)
                  txtFields(7).Text = NotNull(rc("Jenis Angsuran").Value)
                  txtFields(9).Text = fNum(NotNull(rc("Jumlah Angsuran").Value))
                  txtFields(10).Text = NotNull(rc("Angsuran JT").Value)
                  txtFields(11).Text = NotNull(rc("No Seri").Value)
                  txtFields(13).Text = NotNull(rc("Awal Angsuran").Value)
               End If
            End If
       End If
    End If
End If
End Sub

Sub ShowDataPermohonan(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")

Dim hFind As String
hFind = FindRecord("SELECT trn_Perjanjian.[No Permohonan] From trn_Perjanjian WHERE (((trn_Perjanjian.[No Permohonan])='" & hKey(0) & "'));")
If hFind = "1" Then
    hErr = SelectQuery(rc, "SELECT trn_Permohonan_Head.[Kode Inspektur], trn_Permohonan_Head.[No Permohonan], trn_Permohonan_Head.[Tgl Permohonan], trn_Perjanjian.[No Perjanjian], trn_Perjanjian.[Tgl Perjanjian],  trn_Perjanjian.[Tgl Mulai], trn_Permohonan_Head.[Kode Pegawai], trn_Permohonan_Head.[Kode Pelanggan], mst_Pelanggan.Nama, [mst_Pelanggan]![Alamat] & ' ' & [mst_Pelanggan]![RT] & '/' & [mst_Pelanggan]![RW] & '<br>Kel.' & [mst_Pelanggan]![Kelurahan] & ', Kec.' & [mst_Pelanggan]![Kecamatan] & ', ' & [mst_Pelanggan]![Kota] & ' ' & [mst_Pelanggan]![Kode Pos] AS [Alamat Pelanggan], trn_Permohonan_Head.[Uang Muka], trn_Permohonan_Head.[Biaya Adm], trn_Permohonan_Head.Disc " & _
                           "FROM mst_Pelanggan RIGHT JOIN (trn_Permohonan_Head LEFT JOIN trn_Perjanjian ON trn_Permohonan_Head.[No Permohonan] = trn_Perjanjian.[No Permohonan]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan] WHERE (((trn_Permohonan_Head.[No Permohonan])='" & hKey(0) & "'));")
    If hErr = "" Then
       If Not rc.EOF Then
          txtFields(1).Text = NotNull(rc("No Permohonan").Value)
          txtFields(2).Text = NotNull(rc("Tgl Permohonan").Value)
          txtFields(12).Text = NotNull(rc("No Perjanjian").Value)
          txtFields(14).Text = NotNull(rc("Tgl Mulai").Value)
          txtFields(15).Text = NotNull(rc("Nama").Value)
          txtFieldsLine.Text = Replace(NotNull(rc("Alamat Pelanggan").Value), "<br>", vbCrLf)
          ShowInspektur NotNull(rc("Kode Inspektur").Value) & "|"
          ShowSales NotNull(rc("Kode Pegawai").Value) & "|"
       Else
kembali:
          txtFields(1).Text = ""
          txtFields(2).Text = ""
          txtFields(12).Text = ""
          txtFields(14).Text = ""
          txtFields(15).Text = ""
          txtFields(13).Text = ""
          txtFields(16).Text = ""
          txtFields(19).Text = ""
       End If
    Else
    GoSub kembali
    End If
Else
   ShowDlgMsg Me, "No Permohonan tersebut belum terdaftar di surat perjanjian!, apakah anda akan mengisi Surat Perjanjian terlebih dahulu?", vbYesNo, , True, False
   If SelectMsg = vbYes Then
      Form7.Show
      Me.WindowState = vbMinimized
   End If
End If
End Sub

Sub ShowSales(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "Select * from mst_Pegawai WHERE [Kode Divisi]='" & GetDivisi(2) & "' AND [Kode Pegawai]='" & hKey(0) & "'")
If hErr = "" Then
   If Not rc.EOF Then
    txtFields(16).Text = NotNull(rc("Nama Pegawai"))
   Else
    txtFields(16).Text = ""
   End If
Else
    txtFields(16).Text = ""
End If
rc.Close
End Sub

Sub ShowInspektur(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "Select * from mst_Pegawai WHERE [Kode Divisi]='" & GetDivisi(1) & "' AND [Kode Pegawai]='" & hKey(0) & "'")
If hErr = "" Then
   If Not rc.EOF Then
     txtFields(19).Text = NotNull(rc("Nama Pegawai"))
   Else
    txtFields(19).Text = ""
   End If
Else
    txtFields(19).Text = ""
End If
rc.Close
End Sub

Sub InitGrid()
On Error Resume Next
GridMe.Rows = 2
GridMe.MergeCells = flexMergeFixedOnly
GridMe.MergeCol(0) = True
GridMe.MergeCol(1) = True

GridMe.MergeCol(2) = True
GridMe.MergeCol(3) = True

GridMe.MergeCol(4) = True
GridMe.MergeCol(5) = True

GridMe.MergeCol(6) = True
GridMe.MergeCol(7) = True

GridMe.MergeRow(0) = True
GridMe.MergeRow(1) = True


GridMe.Cell(flexcpFontBold, 0, 0, 1, 7) = True
End Sub

Sub PaintDate()
On Error Resume Next
Dim h As Integer, pos As Integer
Dim rc As New ADODB.Recordset
Dim JumHari As Integer
Dim curTgl As String
GridMe.Visible = False
If LCase(txtFields(7).Text) = "minggu" Then
   JumHari = 7 '+ pos
   If (9 - Weekday(CDate(txtFields(14).Text))) < 8 Then
      curTgl = CDate(txtFields(14).Text) + (9 - Weekday(CDate(txtFields(14).Text)))
   Else
      curTgl = (CDate(txtFields(14).Text) - 2)
      curTgl = CDate(curTgl) + (9 - Weekday(CDate(curTgl)))
   End If
Else
   JumHari = 1
   curTgl = txtFields(14).Text
   pos = 0
End If
If (GridMe.Rows - 2) <> (Val(txtFields(6).Text) + Val(txtFields(13).Text)) Then GridMe.Rows = 2

Dim JmlBaris As Integer
JmlBaris = Val(txtFields(6).Text) + Val(txtFields(13).Text)
While h < (JmlBaris)
  If h >= Val(txtFields(13).Text) Then
        If Check1.Value = 1 And Check2.Value = 1 Then
           If (GridMe.Rows - 2) <> JmlBaris Then GridMe.AddItem h + 1
           GridMe.TextMatrix(h + 2, 1) = Format(CDate(curTgl) + pos, "DD-MM-YYYY")
           GridMe.Cell(flexcpForeColor, h + 2, 1) = vbBlue
           pos = pos + JumHari
           h = h + 1
        ElseIf Check1.Value = 1 And Check2.Value = 0 Then
           rc.Open "SELECT mst_hari_libur.Tanggal, mst_hari_libur.Keterangan " & _
                   "From mst_hari_libur WHERE (((mst_hari_libur.Tanggal)=#" & CDate(curTgl) + pos & "#));", srvLogon, LockType1, LockType2
           
           If Not rc.EOF Then
              If (GridMe.Rows - 2) <> JmlBaris Then GridMe.AddItem h + 1
              GridMe.TextMatrix(h + 2, 1) = Format(CDate(curTgl) + pos, "DD-MM-YYYY")
              GridMe.Cell(flexcpForeColor, h + 2, 1) = vbBlue
              h = h + 1
           Else
              Select Case Weekday(CDate(curTgl) + pos)
                     Case 1, 2, 3, 4, 5, 6
                         If (GridMe.Rows - 2) <> JmlBaris Then GridMe.AddItem h + 1
                         GridMe.TextMatrix(h + 2, 1) = Format(CDate(curTgl) + pos, "DD-MM-YYYY")
                         GridMe.Cell(flexcpForeColor, h + 2, 1) = vbBlue
                         h = h + 1
                     Case 7
                        pos = pos + 1
              End Select
           End If
           pos = pos + JumHari
           rc.Close
        ElseIf Check1.Value = 0 And Check2.Value = 1 Then
           rc.Open "SELECT mst_hari_libur.Tanggal, mst_hari_libur.Keterangan " & _
                   "From mst_hari_libur WHERE (((mst_hari_libur.Tanggal)=#" & CDate(curTgl) + pos & "#));", srvLogon, LockType1, LockType2
           
           If rc.EOF Then
              Select Case Weekday(CDate(curTgl) + pos)
                     Case 2, 3, 4, 5, 6, 7
                         If (GridMe.Rows - 2) <> JmlBaris Then GridMe.AddItem h + 1
                         GridMe.TextMatrix(h + 2, 1) = Format(CDate(curTgl) + pos, "DD-MM-YYYY")
                         GridMe.Cell(flexcpForeColor, h + 2, 1) = vbBlue
                         h = h + 1
              End Select
           End If
           pos = pos + JumHari
           rc.Close
        ElseIf Check1.Value = 0 And Check2.Value = 0 Then
           rc.Open "SELECT mst_hari_libur.Tanggal, mst_hari_libur.Keterangan " & _
                   "From mst_hari_libur WHERE (((mst_hari_libur.Tanggal)=#" & CDate(curTgl) + pos & "#));", srvLogon, LockType1, LockType2
           
           If rc.EOF Then
              Select Case Weekday(CDate(curTgl) + pos)
                     Case 2, 3, 4, 5, 6
                         If (GridMe.Rows - 2) <> JmlBaris Then GridMe.AddItem h + 1
                         GridMe.TextMatrix(h + 2, 1) = Format(CDate(curTgl) + pos, "DD-MM-YYYY")
                         GridMe.Cell(flexcpForeColor, h + 2, 1) = vbBlue
                         h = h + 1
                     Case Else
                         'pos = pos + 1
              End Select
           End If
           pos = pos + JumHari
           rc.Close
        End If
   Else
        If (GridMe.Rows - 2) <> (Val(txtFields(6).Text) + Val(txtFields(13).Text)) Then
           GridMe.AddItem h + 1
        End If
        GridMe.TextMatrix(h + 2, 1) = Format(txtFields(14).Text, "DD-MM-YYYY")
        GridMe.TextMatrix(h + 2, 2) = Format(txtFields(14).Text, "DD-MM-YYYY")
        GridMe.TextMatrix(h + 2, 3) = fNum(txtFields(9).Text)
        GridMe.TextMatrix(h + 2, 6) = "PEMBAYARAN AWAL"
        If h = 0 Then
           GridMe.TextMatrix(h + 2, 4) = GridMe.TextMatrix(h + 2, 3)
        Else
           GridMe.TextMatrix(h + 2, 4) = fNum(rNum(txtFields(9).Text) + rNum(GridMe.TextMatrix(h + 1, 4)))
        End If
        GridMe.Cell(flexcpForeColor, h + 2, 1, , 6) = &HC000C0
        h = h + 1
   End If
  DoEvents
Wend
Akumulasi
GridMe.Visible = True
End Sub

Sub Akumulasi()
On Error Resume Next
txtFields(18) = fNum(GridMe.Aggregate(flexSTSum, 2, 3, GridMe.Rows - 1, 3))
txtFields(23) = fNum(rNum(txtFields(8).Text) - rNum(txtFields(18)))
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
txtFields(0).Locked = CekAktifNo("003")
InitGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
PostFindForm "#" & txtFields(0).Hwnd1
PostFindForm "#" & txtFields(1).Hwnd1
PostFindForm "#" & txtFields(4).Hwnd1
CurRec.Close
Set CurRec = Nothing

End Sub

Private Sub GridFind_DblClick()
On Error Resume Next
picBrg.Visible = False
GridMe.Enabled = True
GridMe.Col = 6
GridMe.SetFocus
End Sub

Private Sub GridFind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp And GridFind.Row = 1 Then txtCari.SetFocus
End Sub

Private Sub GridFind_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then GridFind_DblClick
End Sub

Private Sub GridMe_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
Select Case Col
       Case 1, 2
            If Trim(GridMe.TextMatrix(Row, Col)) <> "" Then
               If Not IsDate(GridMe.TextMatrix(Row, Col)) Then
                  GridMe.TextMatrix(Row, Col) = ""
               Else
                  GridMe.TextMatrix(Row, Col) = Format(GridMe.TextMatrix(Row, Col), "DD-MM-YYYY")
                  GridMe.Col = Col + 1
               End If
            End If
            If Col = 2 Then
               If Trim(GridMe.TextMatrix(Row, 3)) = "" Then
                  GridMe.TextMatrix(Row, 3) = fNum(txtFields(9))
                  GridMe_AfterEdit Row, 3
               End If
            End If
       Case 3
            If Not IsNumeric(GridMe.TextMatrix(Row, Col)) Then
               GridMe.TextMatrix(Row, Col) = ""
            Else
               GridMe.TextMatrix(Row, Col) = fNum(GridMe.TextMatrix(Row, Col))
               If Row = 2 Then
                  GridMe.TextMatrix(Row, 4) = GridMe.TextMatrix(Row, 3)
               Else
                  GridMe.TextMatrix(Row, 4) = fNum(rNum(GridMe.TextMatrix(Row, 3)) + rNum(GridMe.TextMatrix(Row - 1, 4)))
               End If
               GridMe.Col = 5
            End If

       Case 6
            If GridMe.Row + 1 < GridMe.Rows Then
               GridMe.Row = GridMe.Row + 1
               GridMe.Col = 2
            End If
 End Select
 GridMe.Cell(flexcpBackColor, Row, 1, , 6) = &HE3E3FF
 Akumulasi
End Sub

Private Sub GridMe_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
Select Case Col
       Case 2, 3, 5, 6
            If Trim(GridMe.TextMatrix(Row, 1)) = "" Then
               Cancel = True
            End If
       Case 4
            Cancel = True
       Case 6

          ' GridMe.TextMatrix(Row, Col) = rNum(GridMe.TextMatrix(Row, Col))
End Select
End Sub

Private Sub GridMe_EnterCell()
On Error Resume Next
Select Case GridMe.Col
       Case 3
         If GridMe.Row > 1 Then
          If GridMe.TextMatrix(GridMe.Row, GridMe.Col) <> "" Then
           GridMe.TextMatrix(GridMe.Row, GridMe.Col) = rNum(GridMe.TextMatrix(GridMe.Row, GridMe.Col))
          End If
         End If
End Select
End Sub

Private Sub GridMe_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case vbKeyInsert
            GridMe.AddItem GridMe.Rows
            SendKeys "{DOWN}", 1
       Case vbKeyDelete
            If GridMe.Rows > 3 Then
               ColDelete.Add txtFields(0).Text & Chr(255) & GridMe.TextMatrix(GridMe.Row, 8)
               'GridMe.RemoveItem GridMe.Row
               GridMe.Cell(flexcpText, GridMe.Row, 1, , 8) = ""
            Else
               ColDelete.Add txtFields(0).Text & Chr(255) & GridMe.TextMatrix(GridMe.Row, 8)
               GridMe.Cell(flexcpText, 2, 1, , 8) = ""
            End If
End Select
End Sub

Private Sub GridMe_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   Select Case Col
          Case 5
            Dim curIsi As String
            curIsi = GridMe.EditText
            If Trim(curIsi) <> "" Then
               Dim lErr As String
               lErr = FindRecord("SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], mst_Divisi.Jabatan " & _
                      "FROM mst_Divisi RIGHT JOIN mst_Pegawai ON mst_Divisi.[Kode Divisi] = mst_Pegawai.[Kode Divisi] WHERE (((mst_Pegawai.[Kode Pegawai])='" & curIsi & "') AND ((mst_Pegawai.[Kode Divisi])='" & GetDivisi(3) & "'));")
               If lErr = "1" Then
                  GridMe.Visible = False
                  ShowDataKorektor curIsi & "|"
                  GridMe.Visible = True
                  GridMe.Col = 6
               ElseIf lErr = "0" Then
                   GridMe.Visible = False
                   GridMe.EditText = ""
                   GridMe.TextMatrix(GridMe.Row, 7) = ""
                   GridMe.TextMatrix(GridMe.Row, 5) = ""
                   ShowFindForm "SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], mst_Divisi.Jabatan " & _
                      "FROM mst_Divisi RIGHT JOIN mst_Pegawai ON mst_Divisi.[Kode Divisi] = mst_Pegawai.[Kode Divisi] <!where> ", "#" & GridMe.hWnd, Me, "ShowDataKorektor", " mst_Pegawai.[Kode Divisi]='" & GetDivisi(3) & "' AND "
                   
                  GridMe.Visible = True
               End If
               'GridMe.Editable = flexEDKbdMouse
            End If
         Case 5
            GridMe.Col = 6
         Case 6
            
   End Select
End If
End Sub

Private Sub GridMe_LeaveCell()
On Error Resume Next
Select Case GridMe.Col
       Case 3
        If GridMe.Row > 1 Then
           If GridMe.TextMatrix(GridMe.Row, GridMe.Col) <> "" Then
            If Not IsNumeric(GridMe.TextMatrix(GridMe.Row, GridMe.Col)) Then
               GridMe.TextMatrix(GridMe.Row, GridMe.Col) = ""
            Else
               GridMe.TextMatrix(GridMe.Row, GridMe.Col) = fNum(GridMe.TextMatrix(GridMe.Row, GridMe.Col))
            End If
           End If
       End If
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.index
       Case 1
           If CekUser("09", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            ClearControl Me
            InitGrid
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            Set ColDelete = New Collection
            If CekAktifNo("003") Then
            txtFields(0).Text = getAutoNo("003")
            txtFields(1).SetFocus
            Else
            txtFields(0).SetFocus
            End If
           End If
       Case 2
           If CekUser("09", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
              SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("09", "D") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            If txtFields(0).Tag <> "" Then
                HapusData txtFields(0).Tag
            Else
                HapusData txtFields(0).Text
            End If
           End If
      Case 5
          txtFields_DownButtonClick 0
      Case 6
            ClearControl Me
            InitGrid
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = ""
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            Set ColDelete = New Collection
      Case 7
           If CekUser("09", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            If GridMe.Rows > 1 Then
            
                Dim hDlg As Boolean
                hDlg = ShowDlgMsg(Me, "Cetak Bukti Angsuran Pembayaran Pelanggan?", vbYesNo, , False, True, , , , , "confirm_" & Me.name)
                If hDlg Then If SelectMsg = vbNo Then Exit Sub
            
                Load frm_util_print_redirect
                frm_util_print_redirect.Tag = "angsuran"
                frm_util_print_redirect.Show 1
            Else
               ShowDlgMsg Me, "Silahkan pilih terlebih dahulu angsuran yang akan dicetak", vbOK, , True, False
            End If
           End If
      Case 12
            Unload Me
End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If CurRec.State = 0 Then
   GoSub subLoadDB
End If

Select Case Button.index
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               ShowAllData NotNull(CurRec("No Angsuran")) & "|"
            End If
            If CurRec.BOF Then
               CurRec.MoveNext
            End If
       Case 1
            If Not CurRec.BOF Then
               CurRec.MoveFirst
               ShowAllData NotNull(CurRec("No Angsuran")) & "|"
            End If
       Case 4
            If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowAllData NotNull(CurRec("No Angsuran")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               ShowAllData NotNull(CurRec("No Angsuran")) & "|"
            End If
            If CurRec.EOF Then
               CurRec.MovePrevious
            End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowAllData NotNull(CurRec("No Angsuran")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [No Angsuran] From trn_Angsuran_Head ORDER BY [No Angsuran]"
'Return

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.index
       Case 1
            txtCell_KeyPress 13
End Select
End Sub

Private Sub txtCari_GotFocus()
BlokX txtCari, 0
End Sub

Private Sub txtCari_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDown Then
   GridFind.SetFocus
   GridFind.Row = 1
End If
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   Dim rc As New ADODB.Recordset
   If Option1.Value Then
      rc.Open "SELECT mst_Barang.[Kode Barang] As Kode, mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type, mst_Barang.Satuan, mst_Barang.[Harga Jual] FROM mst_barang WHERE Left([Nama Barang]," & Len(txtCari.Text) & ")='" & txtCari & "'", srvLogon, LockType1, LockType2
   Else
      rc.Open "SELECT mst_Barang.[Kode Barang] As Kode, mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type, mst_Barang.Satuan, mst_Barang.[Harga Jual] FROM mst_barang WHERE [Nama Barang] LIKE '%" & txtCari & "%'", srvLogon, LockType1, LockType2
   End If
   Set GridFind.DataSource = rc
       GridFind.Refresh
  GridFind.ColWidth(0) = 1000
  GridFind.ColWidth(1) = 3000
  GridFind.ColWidth(2) = 1400
  GridFind.ColWidth(3) = 1400
  GridFind.Cell(flexcpFontBold, 0, 0, , 5) = True
  KeyAscii = 0
End If
End Sub

Private Sub txtCell_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   On Error Resume Next
    GridMe.ShowCell Val(txtCell.Text) + 1, 0
    GridMe.Select Val(txtCell.Text) + 1, 0
    txtCell.SelStart = 0
    txtCell.SelLength = Len(txtCell.Text)
    KeyAscii = 0
    If Err.Number <> 0 Then
       txtCell.BackColor = vbRed
    Else
       txtCell.BackColor = vbWhite
       GridMe.Col = 1
       GridMe.SetFocus
    End If
End If
End Sub

Private Sub txtFields_DownButtonClick(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            ShowFindForm "SELECT trn_Angsuran_Head.[No Angsuran], trn_Angsuran_Head.[No Permohonan], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, trn_Angsuran_Head.[No Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type " & _
                         "FROM mst_Barang RIGHT JOIN ((mst_Pelanggan RIGHT JOIN (trn_Permohonan_Head RIGHT JOIN trn_Angsuran_Head ON trn_Permohonan_Head.[No Permohonan] = trn_Angsuran_Head.[No Permohonan]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) LEFT JOIN trn_Permohonan_Detail ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang] " & _
                         " <!where> ORDER BY trn_Angsuran_Head.[No Angsuran];", "#" & txtFields(index).Hwnd1, Me, "ShowAllData"
       Case 1
            ShowFindForm "SELECT trn_Permohonan_Head.[No Permohonan], trn_Permohonan_Head.[Tgl Permohonan], trn_Permohonan_Head.[Tgl Wawancara], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kelurahan, mst_Pelanggan.Kecamatan, mst_Pelanggan.Kota " & _
                         "FROM mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan] <!where>  " & _
                         "ORDER BY trn_Permohonan_Head.[No Permohonan];", "#" & txtFields(index).Hwnd1, Me, "ShowDataPermohonan"
              
       Case 4
            ShowFindForm "SELECT trn_Permohonan_Detail.[No Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type, mst_Barang.Satuan " & _
                         "FROM mst_Barang RIGHT JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang] <!where> ", _
                         "#" & txtFields(index).Hwnd1, Me, "ShowDataNoBarang", " trn_Permohonan_Detail.[No Permohonan]='" & txtFields(1).Text & "' AND "
        
End Select
End Sub

Private Sub vbButton1_Click()
txtCari_KeyPress 13
End Sub

Private Sub vbButton2_Click()
GridFind_DblClick
End Sub


Private Sub vbButton4_Click()
On Error Resume Next
vbButton4.Enabled = False
PaintDate
vbButton4.Enabled = True
End Sub
