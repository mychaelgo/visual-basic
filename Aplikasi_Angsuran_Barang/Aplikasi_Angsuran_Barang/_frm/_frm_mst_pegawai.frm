VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Data Pegawai"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_mst_pegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   10695
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   1770
      TabIndex        =   15
      Top             =   2730
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
         Index           =   5
         Left            =   1380
         TabIndex        =   17
         Top             =   150
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
         Index           =   4
         Left            =   90
         TabIndex        =   16
         Top             =   150
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   1770
      TabIndex        =   9
      Top             =   2175
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
         Index           =   3
         Left            =   2775
         TabIndex        =   13
         Top             =   210
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
         Index           =   2
         Left            =   1980
         TabIndex        =   12
         Top             =   210
         Width           =   795
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
         Index           =   1
         Left            =   1125
         TabIndex        =   11
         Top             =   210
         Width           =   795
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
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   195
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   3
      Left            =   1770
      TabIndex        =   7
      Top             =   1860
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   33023
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
      Index           =   6
      Left            =   1770
      TabIndex        =   19
      Top             =   3315
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
      Index           =   0
      Left            =   1770
      TabIndex        =   1
      Top             =   780
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Icon            =   "_frm_mst_pegawai.frx":038A
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
      Index           =   1
      Left            =   1770
      TabIndex        =   3
      Top             =   1140
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
      Index           =   2
      Left            =   1770
      TabIndex        =   5
      Top             =   1500
      Width           =   2190
      _ExtentX        =   3863
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
      Left            =   1770
      TabIndex        =   27
      Top             =   4755
      Width           =   2190
      _ExtentX        =   3863
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
      Left            =   1770
      TabIndex        =   25
      Top             =   4395
      Width           =   2190
      _ExtentX        =   3863
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
      Left            =   1770
      TabIndex        =   21
      Top             =   3675
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   1770
      TabIndex        =   23
      Top             =   4035
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
      Index           =   11
      Left            =   1770
      TabIndex        =   29
      Top             =   5130
      Width           =   1545
      _ExtentX        =   2725
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
      Left            =   7245
      TabIndex        =   31
      Top             =   1140
      Width           =   1920
      _ExtentX        =   3387
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
      Left            =   7245
      TabIndex        =   33
      Top             =   1500
      Width           =   3210
      _ExtentX        =   5662
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
      Left            =   7245
      TabIndex        =   35
      Top             =   1860
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Index           =   15
      Left            =   7245
      TabIndex        =   37
      Top             =   2220
      Width           =   2190
      _ExtentX        =   3863
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
      Left            =   7245
      TabIndex        =   39
      Top             =   2595
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      Icon            =   "_frm_mst_pegawai.frx":07D8
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
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   330
      Index           =   5
      Left            =   7245
      TabIndex        =   41
      Top             =   2955
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   582
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8070
      Top             =   4440
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
            Picture         =   "_frm_mst_pegawai.frx":0C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":0FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":135A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":16F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":1E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":21C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":255C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":28F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":2C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":302A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":33C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":375E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":3AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_pegawai.frx":4092
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
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
      TabIndex        =   43
      Top             =   5835
      Width           =   10695
      _ExtentX        =   18865
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
      Index           =   5
      X1              =   -1125
      X2              =   18370
      Y1              =   5730
      Y2              =   5730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   -1125
      X2              =   18370
      Y1              =   5715
      Y2              =   5715
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pegawai"
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
      Left            =   180
      TabIndex        =   0
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Divisi"
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
      Left            =   5685
      TabIndex        =   38
      Top             =   2655
      Width           =   900
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No KTP/SIM/KK"
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
      Left            =   5685
      TabIndex        =   36
      Top             =   2325
      Width           =   1170
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Left            =   5685
      TabIndex        =   32
      Top             =   1605
      Width           =   435
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Divisi"
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
      Left            =   5685
      TabIndex        =   40
      Top             =   3000
      Width           =   435
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Masuk"
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
      Left            =   5685
      TabIndex        =   34
      Top             =   1980
      Width           =   855
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No HP"
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
      Left            =   5685
      TabIndex        =   30
      Top             =   1245
      Width           =   465
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
      Left            =   180
      TabIndex        =   28
      Top             =   5250
      Width           =   615
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pos"
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
      Left            =   180
      TabIndex        =   20
      Top             =   3735
      Width           =   780
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
      Index           =   9
      Left            =   180
      TabIndex        =   22
      Top             =   4110
      Width           =   360
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
      Left            =   180
      TabIndex        =   24
      Top             =   4485
      Width           =   900
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kelurahan"
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
      Left            =   180
      TabIndex        =   26
      Top             =   4830
      Width           =   825
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
      Left            =   180
      TabIndex        =   18
      Top             =   3375
      Width           =   570
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
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
      Left            =   180
      TabIndex        =   2
      Top             =   1215
      Width           =   450
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
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   1545
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
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   1920
      Width           =   720
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
      Left            =   180
      TabIndex        =   8
      Top             =   2370
      Width           =   525
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
      Index           =   5
      Left            =   180
      TabIndex        =   14
      Top             =   2955
      Width           =   300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   165
      X2              =   19660
      Y1              =   7515
      Y2              =   7515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   165
      X2              =   19660
      Y1              =   7500
      Y2              =   7500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   19495
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   19495
      Y1              =   615
      Y2              =   615
   End
End
Attribute VB_Name = "Form4"
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
           Case vbKeyEscape 'Keluar
                If txtFields(0).Text = "" Then
                   Unload Me
                Else
                    Set hBtn = Toolbar2.Buttons(6)
                        Toolbar1_ButtonClick hBtn
                        Set hBtn = Nothing
                End If
                KeyCode = 0
          Case vbKeyF2 'Baru
                Set hBtn = Toolbar2.Buttons(1)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
          Case vbKeyF3 'Simpan
                Set hBtn = Toolbar2.Buttons(2)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
          Case vbKeyF4 'Hapus
                Set hBtn = Toolbar2.Buttons(4)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
          Case vbKeyF5 'Cari
                Set hBtn = Toolbar2.Buttons(5)
                    Toolbar1_ButtonClick hBtn
                    Set hBtn = Nothing
                    
    End Select
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
txtFields(0).Locked = CekAktifNo("004")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
PostFindForm "#" & txtFields(0).Hwnd1
PostFindForm "#" & txtFields(4).Hwnd1
End Sub

Private Sub Option1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   Select Case index
         Case 0 To 3
            Option1(4).SetFocus
         Case 4, 5
            txtFields(6).SetFocus
  End Select
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
 Select Case Button.index
       Case 1
           If CekUser("03", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            ClearControl Me
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            If CekAktifNo("004") Then
               txtFields(0).Text = getAutoNo("004")
               txtFields(1).SetFocus
            Else
               txtFields(0).SetFocus
            End If
           End If
       Case 2
           If CekUser("03", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("03", "D") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            HapusData (txtFields(0).Text)
           End If
       Case 5

       Case 6
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            ClearControl Me
       Case 7
           If CekUser("03", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            Dim StrSql As String, Form4 As New frm_util_report
            Load Form4
            StrSql = "SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], mst_Pegawai.[Tmp Lahir], mst_Pegawai.[Tgl Lahir], mst_Pegawai.Status, mst_Pegawai.Sex, mst_Pegawai.Alamat, mst_Pegawai.[Kode Pos], mst_Pegawai.Kota, mst_Pegawai.Kecamatan, mst_Pegawai.Kelurahan, mst_Pegawai.Telp, mst_Pegawai.Hp, mst_Pegawai.Email, mst_Pegawai.[Tgl Masuk], mst_Pegawai.[Ref ID], mst_Pegawai.[Kode Divisi] " & _
                     "From mst_Pegawai  <!where> ORDER BY mst_Pegawai.[Kode Pegawai];"

            Form4.ARView.Tag = "lap_pegawai|" & StrSql
            Form4.ShowField StrSql
            Form4.Show
            Form4.Left = 0
            Form4.Top = 0
            Form4.ZOrder 0
           End If
       Case 8
            
       Case 9
           
       Case 11
            Unload Me
            End Select
End Sub

Sub SimpanData(nKey As String)
On Error Resume Next
Dim h As String, hErr As String, sStatus, sSex, i As Integer

h = FindRecord("SELECT mst_pegawai.[kode pegawai], mst_pegawai.[nama pegawai], mst_pegawai.[tmp lahir], mst_pegawai.[tgl lahir], mst_pegawai.[status], mst_pegawai.[sex], " & _
                      "mst_pegawai.[alamat],  mst_pegawai.[kode pos], mst_pegawai.[kota], mst_pegawai.[kecamatan], mst_pegawai.[kelurahan], " & _
                      "mst_pegawai.[telp],  mst_pegawai.[hp], mst_pegawai.[email], mst_pegawai.[tgl masuk], mst_pegawai.[ref id] " & _
                      "FROM mst_pegawai WHERE (((mst_pegawai.[kode pegawai])='" & nKey & "'));")
         'MsgBox H
If Option1(0).Value = True Then
sStatus = 2
ElseIf Option1(1).Value = True Then
sStatus = 1
ElseIf Option1(2).Value = True Then
sStatus = 3
ElseIf Option1(3).Value = True Then
sStatus = 4
End If

If Option1(4).Value = True Then
sSex = 1
ElseIf Option1(5).Value = True Then
sSex = 2
End If

If h = "0" Then
           
   h = SaveRecord("mst_pegawai", Array("kode pegawai=" & txtFields(0).Text, _
                                                  "nama pegawai=" & txtFields(1).Text, _
                                                  "tmp lahir=" & txtFields(2).Text, _
                                                  "@tgl lahir=" & txtFields(3).Text, _
                                                  "status=" & sStatus, _
                                                  "sex=" & sSex, _
                                                  "alamat=" & txtFields(6).Text, _
                                                  "kode pos=" & txtFields(9).Text, _
                                                  "kota=" & txtFields(10).Text, _
                                                  "kecamatan=" & txtFields(8).Text, _
                                                  "kelurahan=" & txtFields(7).Text, _
                                                  "telp=" & txtFields(11).Text, _
                                                  "hp=" & txtFields(12).Text, _
                                                  "email=" & txtFields(13).Text, _
                                                  "@tgl masuk=" & txtFields(14).Text, _
                                                  "ref id=" & txtFields(15).Text, _
                                                  "kode divisi=" & txtFields(4).Text))
                                                  
                                                  
  If h = "" Then
       If CekAktifNo("004") Then txtFields(0).Text = getAutoNo("004", True)
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
         h = UpdateRecord("mst_pegawai", Array("kode pegawai=" & txtFields(0).Text, _
                                                  "nama pegawai=" & txtFields(1).Text, _
                                                  "tmp lahir=" & txtFields(2).Text, _
                                                  "@tgl lahir=" & txtFields(3).Text, _
                                                  "status=" & sStatus, _
                                                  "sex=" & sSex, _
                                                  "alamat=" & txtFields(6).Text, _
                                                  "kode pos=" & txtFields(9).Text, _
                                                  "kota=" & txtFields(10).Text, _
                                                  "kecamatan=" & txtFields(8).Text, _
                                                  "kelurahan=" & txtFields(7).Text, _
                                                  "telp=" & txtFields(11).Text, _
                                                  "hp=" & txtFields(12).Text, _
                                                  "email=" & txtFields(13).Text, _
                                                  "@tgl masuk=" & txtFields(14).Text, _
                                                  "ref id=" & txtFields(15).Text, _
                                                  "kode divisi=" & txtFields(4).Text), " WHERE [Kode Pegawai]='" & txtFields(0).Text & "'")
                                                          
     
        If h = "" Then
             txtFields(0).Tag = txtFields(0).Text
             Me.Caption = Replace(Me.Caption, "*", "")
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
hErr = FindRecord("SELECT mst_pegawai.[kode pegawai] From mst_pegawai WHERE (((mst_pegawai.[kode pegawai])='" & hKey & "'));")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Pegawai?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") = False Then
      GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE mst_pegawai.[kode pegawai] From mst_pegawai WHERE (((mst_pegawai.[kode pegawai])='" & hKey & "'));")
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

Sub ShowDivisi(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String

hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT mst_divisi.[Kode Divisi],  mst_divisi.[Jabatan] from mst_divisi where [Kode Divisi]='" & hKey(0) & "' ORDER BY mst_divisi.[Kode Divisi]; ")
If hErr = "" Then
    If Not rc.EOF Then
        txtFields(4).Text = NotNull(rc("Kode Divisi"))
        txtFields(5).Text = NotNull(rc("Jabatan"))
    Else
kembali:
        txtFields(4).Text = ""
        txtFields(5).Text = ""
End If
Else
   GoSub kembali
End If
rc.Close
End Sub


Sub ShowPegawai(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
Dim sStatus, sSex As Integer

hKey = Split(nKey, "|")

hErr = SelectQuery(rc, "Select * from mst_Pegawai WHERE [Kode Pegawai]='" & hKey(0) & "' ORDER BY mst_pegawai.[Kode Pegawai]")

'MsgBox Herr

If hErr = "" Then
    If Not rc.EOF Then
        txtFields(0).Text = NotNull(rc("Kode Pegawai"))
        txtFields(1).Text = NotNull(rc("Nama Pegawai"))
        txtFields(2).Text = NotNull(rc("tmp lahir"))
        txtFields(3).Text = NotNull(rc("tgl lahir"))
        sStatus = NotNull(rc("status"))
        
        If sStatus = 1 Then
        Option1(1).Value = True
        ElseIf sStatus = 2 Then
        Option1(0).Value = True
        ElseIf sStatus = 3 Then
        Option1(2).Value = True
        ElseIf sStatus = 4 Then
        Option1(3).Value = True
        End If
        
        sSex = NotNull(rc("sex"))
        If sSex = 1 Then
        Option1(4).Value = True
        ElseIf sStatus = 2 Then
        Option1(5).Value = True
        End If
        txtFields(6).Text = NotNull(rc("alamat"))
        txtFields(9).Text = NotNull(rc("Kode Pos"))
        txtFields(10).Text = NotNull(rc("Kota"))
        txtFields(8).Text = NotNull(rc("Kecamatan"))
        txtFields(7).Text = NotNull(rc("Kelurahan"))
        txtFields(11).Text = NotNull(rc("Telp"))
        txtFields(12).Text = NotNull(rc("hp"))
        txtFields(13).Text = NotNull(rc("email"))
        txtFields(14).Text = NotNull(rc("tgl masuk"))
        txtFields(15).Text = NotNull(rc("ref id"))
        ShowDivisi NotNull(rc("Kode Divisi").Value) & "|"
        Else
kembali:
      ClearControl Me
    End If
Else
   GoSub kembali
End If
rc.Close
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
               ShowPegawai NotNull(CurRec("Kode Pegawai")) & "|"
            'End If
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               If CurRec.BOF Then
                  CurRec.MoveNext
               End If
               ShowPegawai NotNull(CurRec("Kode Pegawai")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               If CurRec.EOF Then
                  CurRec.MovePrevious
               End If
               ShowPegawai NotNull(CurRec("Kode Pegawai")) & "|"
            End If
       Case 4
            'If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowPegawai NotNull(CurRec("Kode Pegawai")) & "|"
            'End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowPegawai NotNull(CurRec("Kode Pegawai")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [Kode Pegawai] From mst_Pegawai ORDER BY [Kode Pegawai]"
Return
End Sub

Private Sub txtFields_DownButtonClick(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            ShowFindForm "SELECT mst_pegawai.[kode pegawai], mst_pegawai.[nama pegawai], mst_pegawai.[tmp lahir], mst_pegawai.[tgl lahir], mst_pegawai.[status], mst_pegawai.[sex], " & _
                      "mst_pegawai.[alamat],  mst_pegawai.[kode pos], mst_pegawai.[kota], mst_pegawai.[kecamatan], mst_pegawai.[kelurahan], " & _
                      "mst_pegawai.[telp],  mst_pegawai.[hp], mst_pegawai.[email], mst_pegawai.[tgl masuk], mst_pegawai.[ref id], " & _
                      "FROM mst_pegawai <!where> ORDER BY mst_pegawai.[Kode Pegawai]; ", "#" & txtFields(index).Hwnd1, Me, "ShowPegawai"
      
      Case 4
            ShowFindForm "SELECT mst_divisi.[kode divisi], mst_divisi.[jabatan] FROM mst_divisi <!where> ORDER BY mst_divisi.[kode divisi]; ", "#" & txtFields(index).Hwnd1, Me, "ShowDivisi"
       End Select
End Sub

Private Sub txtFields_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case 0
            Select Case index
                   Case 0
                      ShowPegawai txtFields(index).Text & "|"
                   Case 4
                      ShowDivisi txtFields(index).Text & "|"
            End Select
       Case 13
            Select Case index
                   Case 3
                      Option1(0).SetFocus
            End Select
       Case Else
            If Me.Tag = "" Then
               Me.Tag = "*"
               Me.Caption = Me.Caption & Me.Tag
            End If
End Select
End Sub
