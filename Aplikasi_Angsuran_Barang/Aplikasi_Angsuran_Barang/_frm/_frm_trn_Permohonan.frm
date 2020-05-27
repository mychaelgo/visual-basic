VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permohonan Sewa Beli Barang - PSBB"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_trn_Permohonan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   12555
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   28
      Left            =   6795
      TabIndex        =   60
      Top             =   7680
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
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
      Index           =   21
      Left            =   10965
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   885
      Width           =   1485
      _ExtentX        =   2619
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
      Index           =   20
      Left            =   7260
      TabIndex        =   40
      Top             =   7290
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
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
      Index           =   19
      Left            =   6795
      TabIndex        =   39
      Top             =   7290
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   556
      Alignment       =   2
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
      AutoTab         =   -1  'True
      FontFormat      =   3
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
      Left            =   10590
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   7650
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
      Locked          =   -1  'True
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
      Index           =   18
      Left            =   10590
      TabIndex        =   44
      Top             =   7290
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
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
      Index           =   17
      Left            =   10590
      TabIndex        =   42
      Top             =   6930
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      Alignment       =   1
      BackColor       =   16777215
      BackColorMain   =   14737632
      DownButton      =   0   'False
      BorderColor     =   8421504
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
      Index           =   16
      Left            =   6795
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6930
      Width           =   1905
      _ExtentX        =   3360
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
      Left            =   7335
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1245
      Width           =   2445
      _ExtentX        =   4313
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
      Left            =   7335
      TabIndex        =   15
      Top             =   885
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Icon            =   "_frm_trn_Permohonan.frx":038A
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
      Index           =   11
      Left            =   9915
      TabIndex        =   28
      Top             =   885
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Icon            =   "_frm_trn_Permohonan.frx":07D8
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
      Height          =   315
      Index           =   10
      Left            =   7335
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2625
      Width           =   2445
      _ExtentX        =   4313
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
      Index           =   9
      Left            =   7335
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2445
      _ExtentX        =   4313
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
      Index           =   8
      Left            =   7335
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1935
      Width           =   2445
      _ExtentX        =   4313
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
      Index           =   7
      Left            =   7335
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2445
      _ExtentX        =   4313
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
      Index           =   6
      Left            =   8385
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   885
      Width           =   1395
      _ExtentX        =   2461
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4335
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":0C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":0FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":135A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":16F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":1E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":21C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":255C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":28F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":2C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":302A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":33C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":375E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":3AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":4092
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":462C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":49C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":4D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":50FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":78D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_Permohonan.frx":7E72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
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
            Caption         =   "Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SPSB"
            ImageIndex      =   21
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
      TabIndex        =   48
      Top             =   8190
      Width           =   12555
      _ExtentX        =   22146
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3525
      Left            =   105
      TabIndex        =   49
      Top             =   3450
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   6218
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "                            "
      TabPicture(0)   =   "_frm_trn_Permohonan.frx":820C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picBrg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GridMe"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Toolbar3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Left            =   105
         TabIndex        =   59
         Top             =   3150
         Width           =   1305
         _ExtentX        =   2302
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
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   16
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   17
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   18
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox GridMe 
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   2970
         Left            =   105
         ScaleHeight     =   2910
         ScaleWidth      =   12060
         TabIndex        =   35
         Top             =   90
         Width           =   12120
      End
      Begin VB.PictureBox picBrg 
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   135
         ScaleHeight     =   2955
         ScaleWidth      =   12060
         TabIndex        =   50
         Top             =   105
         Visible         =   0   'False
         Width           =   12060
         Begin VB.TextBox txtCari 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1365
            TabIndex        =   54
            Top             =   60
            Width           =   9045
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
            Left            =   10665
            TabIndex        =   52
            Top             =   1815
            Value           =   -1  'True
            Width           =   1305
         End
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
            Left            =   10665
            TabIndex        =   51
            Top             =   2055
            Width           =   1260
         End
         Begin SysInfo_Nardhika.vbButton vbButton1 
            Height          =   360
            Left            =   10665
            TabIndex        =   53
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
            MICON           =   "_frm_trn_Permohonan.frx":8228
            PICN            =   "_frm_trn_Permohonan.frx":8244
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
            Height          =   2430
            Left            =   0
            ScaleHeight     =   2430
            ScaleWidth      =   10485
            TabIndex        =   55
            Top             =   405
            Width           =   10485
         End
         Begin SysInfo_Nardhika.vbButton vbButton2 
            Height          =   360
            Left            =   10665
            TabIndex        =   56
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
            MICON           =   "_frm_trn_Permohonan.frx":85DE
            PICN            =   "_frm_trn_Permohonan.frx":85FA
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
            Left            =   10665
            TabIndex        =   57
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
            MICON           =   "_frm_trn_Permohonan.frx":8994
            PICN            =   "_frm_trn_Permohonan.frx":89B0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000004&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   510
            Left            =   -30
            Top             =   -105
            Width           =   10515
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
            TabIndex        =   58
            Top             =   105
            Width           =   1065
         End
         Begin VB.Line Line2 
            X1              =   10635
            X2              =   11925
            Y1              =   1620
            Y2              =   1620
         End
      End
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   14
      Left            =   3150
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2625
      Width           =   2520
      _ExtentX        =   4445
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
      Index           =   15
      Left            =   1800
      TabIndex        =   12
      Top             =   2625
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Icon            =   "_frm_trn_Permohonan.frx":8D4A
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
      Index           =   12
      Left            =   1800
      TabIndex        =   9
      Top             =   2265
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Icon            =   "_frm_trn_Permohonan.frx":9198
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
      Index           =   4
      Left            =   3150
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2265
      Width           =   2520
      _ExtentX        =   4445
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1905
      Width           =   1545
      _ExtentX        =   2725
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
      MaxLength       =   5
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   1545
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
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   825
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      Icon            =   "_frm_trn_Permohonan.frx":95E6
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1185
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
      Index           =   22
      Left            =   9915
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1245
      Width           =   2535
      _ExtentX        =   4471
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
      Index           =   24
      Left            =   9915
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2535
      _ExtentX        =   4471
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
      Index           =   25
      Left            =   9915
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2295
      Width           =   2535
      _ExtentX        =   4471
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
      Index           =   26
      Left            =   9915
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2535
      _ExtentX        =   4471
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
      Index           =   27
      Left            =   9915
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Angsuran x kali"
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
      Left            =   5190
      TabIndex        =   61
      Top             =   7740
      Width           =   1275
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   6000
      TabIndex        =   38
      Top             =   7365
      Width           =   720
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
      Index           =   12
      Left            =   195
      TabIndex        =   11
      Top             =   2685
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
      Index           =   4
      Left            =   195
      TabIndex        =   8
      Top             =   2325
      Width           =   825
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jam Wawancara"
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
      Left            =   195
      TabIndex        =   6
      Top             =   1950
      Width           =   1320
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Wawancara"
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
      Top             =   1590
      Width           =   1230
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Pemohon"
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
      Top             =   1245
      Width           =   1095
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
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   885
      Width           =   1050
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
      Left            =   9435
      TabIndex        =   45
      Top             =   7695
      Width           =   1035
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DP/UANG MUKA"
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
      Left            =   9240
      TabIndex        =   43
      Top             =   7350
      Width           =   1230
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
      Index           =   14
      Left            =   8940
      TabIndex        =   41
      Top             =   6990
      Width           =   1530
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Kredit"
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
      Left            =   5775
      TabIndex        =   36
      Top             =   6990
      Width           =   945
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Penjamin"
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
      Left            =   9945
      TabIndex        =   27
      Top             =   660
      Width           =   765
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat Bekerja"
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
      Left            =   5895
      TabIndex        =   25
      Top             =   2670
      Width           =   1245
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempat Bekerja"
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
      Left            =   5895
      TabIndex        =   23
      Top             =   2325
      Width           =   1305
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Usaha"
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
      Left            =   5895
      TabIndex        =   21
      Top             =   2025
      Width           =   990
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
      Left            =   5895
      TabIndex        =   19
      Top             =   1650
      Width           =   360
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
      Left            =   5895
      TabIndex        =   17
      Top             =   1290
      Width           =   570
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
      Left            =   5895
      TabIndex        =   14
      Top             =   945
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   5
      X1              =   195
      X2              =   19690
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   4
      X1              =   195
      X2              =   19690
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   19495
      Y1              =   615
      Y2              =   615
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
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   -90
      X2              =   19405
      Y1              =   8115
      Y2              =   8115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   -90
      X2              =   19405
      Y1              =   8130
      Y2              =   8130
   End
End
Attribute VB_Name = "Form5"
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
hErr = SelectQuery(rc, "SELECT trn_Permohonan_Head.*, trn_Permohonan_Detail.[No Barang], trn_Permohonan_Detail.[Kode Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type, mst_Barang.Satuan, trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Awal Angsuran], trn_Permohonan_Detail.[No Seri], trn_Permohonan_Detail.Keterangan As Ket " & _
                       ", [trn_Permohonan_Detail]![Harga Kredit]*[trn_Permohonan_Detail]![Qty] AS Total FROM mst_Barang RIGHT JOIN (trn_Permohonan_Head LEFT JOIN trn_Permohonan_Detail ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang] Where (((trn_Permohonan_Head.[No Permohonan]) = '" & hKey(0) & "')) ORDER BY trn_Permohonan_Detail.[No Barang];")

If hErr = "" Then
   If Not rc.EOF Then
      txtFields(0).Text = NotNull(rc("No Permohonan").Value)
      txtFields(0).Tag = NotNull(rc("No Permohonan").Value)
      
      txtFields(1).Text = NotNull(rc("Tgl Permohonan").Value)
      txtFields(2).Text = NotNull(rc("Tgl Wawancara").Value)
      txtFields(3).Text = NotNull(rc("Jam Wawancara").Value)
      
      ShowSales NotNull(rc("Kode Pegawai").Value) & "|"
      ShowPelanggan NotNull(rc("Kode Pelanggan").Value) & "|"
      ShowInspektur NotNull(rc("Kode Inspektur").Value) & "|"
      ShowPenjamin NotNull(rc("Kode Penjamin").Value) & "|", NotNull(rc("Kode Pelanggan").Value)

      txtFields(18).Text = fNum(NotNull(rc("Uang Muka").Value))
      txtFields(17).Text = fNum(NotNull(rc("Biaya Adm").Value))
      txtFields(19).Text = (NotNull(rc("Disc").Value))
            
      Dim pos As Integer
      pos = 2
      GridMe.Rows = pos
      While Not rc.EOF
            GridMe.AddItem pos - 1
            GridMe.TextMatrix(pos, 14) = NotNull(rc("No Barang").Value)
            GridMe.TextMatrix(pos, 15) = NotNull(rc("Kode Barang").Value)
            GridMe.TextMatrix(pos, 1) = NotNull(rc("Kode Barang").Value)
            GridMe.TextMatrix(pos, 2) = NotNull(rc("Nama Barang").Value)
            GridMe.TextMatrix(pos, 3) = NotNull(rc("Merk").Value) & " " & NotNull(rc("Type").Value)
            GridMe.TextMatrix(pos, 4) = NotNull(rc("Satuan").Value)
            GridMe.TextMatrix(pos, 5) = fNum(NotNull(rc("Harga Kredit").Value))
            GridMe.TextMatrix(pos, 6) = NotNull(rc("Qty").Value)
            GridMe.TextMatrix(pos, 7) = NotNull(rc("Lama Angsuran").Value)
            GridMe.TextMatrix(pos, 8) = NotNull(rc("Jenis Angsuran").Value)
            GridMe.TextMatrix(pos, 9) = fNum(NotNull(rc("Jumlah Angsuran").Value))
            GridMe.TextMatrix(pos, 10) = NotNull(rc("Awal Angsuran").Value)
            GridMe.TextMatrix(pos, 11) = NotNull(rc("No Seri").Value)
            GridMe.TextMatrix(pos, 12) = NotNull(rc("Ket").Value)
            GridMe.TextMatrix(pos, 13) = fNum(NotNull(rc("Total").Value))
            rc.MoveNext
            pos = pos + 1
      Wend
      AkumulasiJumlah
      txtFields_LostFocus 19
   Else
        ClearControl Me
        InitGrid
        Me.Caption = Replace(Me.Caption, "*", "")
        txtFields(0).Tag = ""
        Me.Caption = Me.Caption & Me.Tag
   End If
   rc.Close
End If
End Sub

Sub HapusData(hKey As String)
On Error Resume Next
Dim hErr As String
hErr = FindRecord("SELECT trn_Permohonan_Head.[No Permohonan] From trn_Permohonan_Head WHERE trn_Permohonan_Head.[No Permohonan]='" & hKey & "';")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Permohonan Sewa Beli Barang?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") Then
      If SelectMsg = vbYes Then GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE trn_Permohonan_Head.[No Permohonan] From trn_Permohonan_Head WHERE (((trn_Permohonan_Head.[No Permohonan])='" & hKey & "'));")
         If hErr = "" Then
            ClearControl Me
            InitGrid
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = ""
            Set ColDelete = New Collection
         End If
      End If
   End If
ElseIf hErr = "0" Then
    ShowDlgMsg Me, "Tidak ada data yang akan dihapus", vbOK, , True, False
End If
End Sub
Sub ShowPelanggan(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT mst_Pelanggan.[Kode Pelanggan], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, mst_Pelanggan.[Jenis Usaha], mst_Pelanggan.[Nama Perusahaan], mst_Pelanggan.[Alamat Usaha]  FROM mst_Pelanggan where [Kode Pelanggan]='" & hKey(0) & "' ORDER BY mst_Pelanggan.[Kode Pelanggan]; ")
If hErr = "" Then
    If Not rc.EOF Then
        txtFields(5).Text = NotNull(rc("Kode Pelanggan"))
        txtFields(6).Text = NotNull(rc("Nama"))
        txtFields(13).Text = NotNull(rc("Alamat"))
        txtFields(7).Text = NotNull(rc("Kota"))
        txtFields(8).Text = NotNull(rc("Jenis Usaha"))
        txtFields(9).Text = NotNull(rc("Nama Perusahaan"))
        txtFields(10).Text = NotNull(rc("Alamat Usaha"))
    Else
kembali:
        txtFields(5).Text = ""
        txtFields(6).Text = ""
        txtFields(13).Text = ""
        txtFields(7).Text = ""
        txtFields(8).Text = ""
        txtFields(9).Text = ""
        txtFields(10).Text = ""
    End If
Else
   GoSub kembali
End If
rc.Close
End Sub

Sub ShowSales(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "Select * from mst_Pegawai WHERE [Kode Divisi]='" & GetDivisi(2) & "' AND [Kode Pegawai]='" & hKey(0) & "'")
If hErr = "" Then
   If Not rc.EOF Then
    txtFields(12).Text = NotNull(rc("Kode Pegawai"))
    txtFields(4).Text = NotNull(rc("Nama Pegawai"))
   Else
    txtFields(12).Text = ""
    txtFields(4).Text = ""
   End If
Else
    txtFields(12).Text = ""
    txtFields(4).Text = ""
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
    txtFields(15).Text = NotNull(rc("Kode Pegawai"))
    txtFields(14).Text = NotNull(rc("Nama Pegawai"))
   Else
    txtFields(15).Text = ""
    txtFields(14).Text = ""
   End If
Else
    txtFields(15).Text = ""
    txtFields(14).Text = ""
End If
rc.Close
End Sub

Sub ShowPenjamin(nKey As String, Optional hkey2 As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
If hkey2 = "" Then hkey2 = txtFields(5).Text
hErr = SelectQuery(rc, "SELECT mst_Penjamin.[Kode Penjamin], mst_Penjamin.[Nama Penjamin], mst_Penjamin.[Alamat Penjamin], mst_Penjamin.[Kota Penjamin], mst_Penjamin.[Jenis Usaha Penjamin], mst_Penjamin.[Nama Perusahaan Penjamin], mst_Penjamin.[Alamat Usaha Penjamin] " & _
       "FROM mst_Pelanggan INNER JOIN mst_Penjamin ON mst_Pelanggan.[Kode Pelanggan] = mst_Penjamin.[Kode Pelanggan] " & _
       "WHERE [Kode Penjamin]='" & hKey(0) & "' AND (mst_Pelanggan.[Kode Pelanggan]='" & hkey2 & "');")
       
If hErr = "" Then
   If Not rc.EOF Then
    txtFields(11).Text = NotNull(rc("Kode Penjamin"))
    txtFields(21).Text = NotNull(rc("Nama Penjamin"))
    txtFields(22).Text = NotNull(rc("Alamat Penjamin"))
    txtFields(27).Text = NotNull(rc("Kota Penjamin"))
    txtFields(26).Text = NotNull(rc("Jenis Usaha Penjamin"))
    txtFields(25).Text = NotNull(rc("Nama Perusahaan Penjamin"))
    txtFields(24).Text = NotNull(rc("Alamat Usaha Penjamin"))
   Else
    txtFields(11).Text = ""
    txtFields(21).Text = ""
    txtFields(22).Text = ""
    txtFields(27).Text = ""
    txtFields(26).Text = ""
    txtFields(25).Text = ""
    txtFields(24).Text = ""
   End If
Else
    txtFields(11).Text = ""
    txtFields(21).Text = ""
    txtFields(22).Text = ""
    txtFields(27).Text = ""
    txtFields(26).Text = ""
    txtFields(25).Text = ""
    txtFields(24).Text = ""
End If
rc.Close
End Sub

Sub AkumulasiJumlah()
On Error Resume Next
Dim i As Integer, Akum As Currency, Ang As Currency
For i = 2 To GridMe.Rows - 1
   If GridMe.TextMatrix(i, 13) <> "" Then
      Akum = Akum + rNum(GridMe.TextMatrix(i, 13))
   End If
   Ang = Ang + rNum(GridMe.TextMatrix(i, 9) * rNum(GridMe.TextMatrix(i, 10)))
Next i
'TK-(Disc+Adm+um)
txtFields(16).Text = fNum(Akum)
txtFields(28).Text = fNum(Ang)
txtFields(23).Text = fNum(rNum(txtFields(16)) - (rNum(txtFields(18)) + rNum(txtFields(20)) + Ang))
End Sub
Sub SimpanData(nKey As String)
On Error Resume Next
Dim h As String, hErr As String, i As Integer
Dim X
h = FindRecord("SELECT trn_Permohonan_Head.[No Permohonan] From trn_Permohonan_Head WHERE (((trn_Permohonan_Head.[No Permohonan])='" & nKey & "'));")
If h = "0" Then
   hErr = SaveRecord("trn_Permohonan_Head", Array("No Permohonan=" & txtFields(0).Text, _
                                                  "@Tgl Permohonan=" & txtFields(1).Text, _
                                                  "@Tgl Wawancara=" & txtFields(2).Text, _
                                                  "Jam Wawancara=" & txtFields(3).Text, _
                                                  "Kode Pegawai=" & txtFields(12).Text, _
                                                  "Kode Pelanggan=" & txtFields(5).Text, _
                                                  "Kode Inspektur=" & txtFields(15).Text, _
                                                  "$Uang Muka=" & txtFields(18).Text, _
                                                  "$Biaya Adm=" & txtFields(17).Text, _
                                                  "$Disc=" & txtFields(19).Text, _
                                                  "Kode Penjamin=" & txtFields(11).Text))
  If hErr = "" Then
      
      For i = 1 To ColDelete.Count
          X = Split(ColDelete(i), Chr(255))
          ExecQuery "DELETE FROM trn_Permohonan_Detail WHERE ([No Permohonan]= '" & AllowChar(CStr(X(0))) & "') AND ([No Barang]= '" & AllowChar(CStr(X(1))) & "')"
      Next i

     txtFields(0).Tag = txtFields(0).Text
     For i = 2 To GridMe.Rows - 1
        If GridMe.TextMatrix(i, 1) <> "" Then
           hErr = SaveRecord("trn_Permohonan_Detail", Array("No Permohonan=" & txtFields(0), _
                                                            "No Barang=" & GridMe.TextMatrix(i, 1) & Format(i, "00#"), _
                                                            "Kode Barang=" & GridMe.TextMatrix(i, 1), _
                                                            "$Harga Kredit=" & GridMe.TextMatrix(i, 5), _
                                                            "#Qty=" & GridMe.TextMatrix(i, 6), _
                                                            "Lama Angsuran=" & GridMe.TextMatrix(i, 7), _
                                                            "Jenis Angsuran=" & GridMe.TextMatrix(i, 8), _
                                                            "Jumlah Angsuran=" & GridMe.TextMatrix(i, 9), _
                                                            "#Awal Angsuran=" & GridMe.TextMatrix(i, 10), _
                                                            "No Seri=" & GridMe.TextMatrix(i, 11), _
                                                            "Keterangan=" & GridMe.TextMatrix(i, 12)))
           
           GridMe.TextMatrix(i, 14) = GridMe.TextMatrix(i, 1) & Format(i, "00#")
           GridMe.TextMatrix(i, 15) = GridMe.TextMatrix(i, 1)
           If hErr = "" Then
                'Error Catch here!
           End If
        End If
     Next i
     If CekAktifNo("002") Then txtFields(0).Text = getAutoNo("002", True)
     txtFields(0).Tag = txtFields(0).Text
     Me.Caption = Replace(Me.Caption, "*", "")
     Set ColDelete = New Collection
     Me.Tag = ""
  Else
     ShowDlgMsg MainMenu, "Error!!!<br><br>Tidak dapat menyimpan data!", vbOK, Error, True, False
  End If
ElseIf h = "1" Then
   If ShowDlgMsg(Me, "Data sudah terdaftar!, update dengan data baru?", vbYesNo, Error, False, True, , , , , Me.name & "_update") Then
      GoSub SimpanLabel
   Else
      If SelectMsg = vbYes Then
SimpanLabel:
      
      For i = 1 To ColDelete.Count
          X = Split(ColDelete(i), Chr(255))
          ExecQuery "DELETE FROM trn_Permohonan_Detail WHERE ([No Permohonan]= '" & AllowChar(CStr(X(0))) & "') AND ([No Barang]= '" & AllowChar(CStr(X(1))) & "')"
      Next i

         hErr = UpdateRecord("trn_Permohonan_Head", Array("No Permohonan=" & txtFields(0).Text, _
                                                          "@Tgl Permohonan=" & txtFields(1).Text, _
                                                          "@Tgl Wawancara=" & txtFields(2).Text, _
                                                          "Jam Wawancara=" & txtFields(3).Text, _
                                                          "Kode Pegawai=" & txtFields(12).Text, _
                                                          "Kode Pelanggan=" & txtFields(5).Text, _
                                                          "Kode Inspektur=" & txtFields(15).Text, _
                                                          "$Uang Muka=" & txtFields(18).Text, _
                                                          "$Biaya Adm=" & txtFields(17).Text, _
                                                          "$Disc=" & txtFields(19).Text, _
                                                          "Kode Penjamin=" & txtFields(11).Text), " WHERE [No Permohonan]='" & txtFields(0).Tag & "' ")
        If hErr = "" Then
           For i = 2 To GridMe.Rows - 1
              If GridMe.TextMatrix(i, 1) <> "" Then
                 If GridMe.TextMatrix(i, 14) = "" Then
                      hErr = SaveRecord("trn_Permohonan_Detail", Array("No Permohonan=" & txtFields(0), _
                                                                       "No Barang=" & GridMe.TextMatrix(i, 1) & Format(i, "00#"), _
                                                                       "Kode Barang=" & GridMe.TextMatrix(i, 1), _
                                                                       "$Harga Kredit=" & GridMe.TextMatrix(i, 5), _
                                                                       "#Qty=" & GridMe.TextMatrix(i, 6), _
                                                                       "Lama Angsuran=" & GridMe.TextMatrix(i, 7), _
                                                                       "Jenis Angsuran=" & GridMe.TextMatrix(i, 8), _
                                                                       "Jumlah Angsuran=" & GridMe.TextMatrix(i, 9), _
                                                                       "#Awal Angsuran=" & GridMe.TextMatrix(i, 10), _
                                                                       "No Seri=" & GridMe.TextMatrix(i, 11), _
                                                                       "Keterangan=" & GridMe.TextMatrix(i, 12)))
                      
                      GridMe.TextMatrix(i, 14) = GridMe.TextMatrix(i, 1) & Format(i, "00#")
                      GridMe.TextMatrix(i, 15) = GridMe.TextMatrix(i, 1)
                      If hErr = "" Then
                           'Error Catch here!
                      End If
                  Else '!>
                      hErr = UpdateRecord("trn_Permohonan_Detail", Array("No Permohonan=" & txtFields(0), _
                                                                         "No Barang=" & GridMe.TextMatrix(i, 1) & Format(i, "00#"), _
                                                                         "Kode Barang=" & GridMe.TextMatrix(i, 1), _
                                                                         "$Harga Kredit=" & GridMe.TextMatrix(i, 5), _
                                                                         "#Qty=" & GridMe.TextMatrix(i, 6), _
                                                                         "Lama Angsuran=" & GridMe.TextMatrix(i, 7), _
                                                                         "Jenis Angsuran=" & GridMe.TextMatrix(i, 8), _
                                                                         "Jumlah Angsuran=" & GridMe.TextMatrix(i, 9), _
                                                                         "#Awal Angsuran=" & GridMe.TextMatrix(i, 10), _
                                                                         "No Seri=" & GridMe.TextMatrix(i, 11), _
                                                                         "Keterangan=" & GridMe.TextMatrix(i, 12)), " WHERE (trn_Permohonan_Detail.[No Permohonan]='" & txtFields(0).Tag & "') AND (trn_Permohonan_Detail.[No Barang]='" & GridMe.TextMatrix(i, 14) & "')")
                      
                      GridMe.TextMatrix(i, 14) = GridMe.TextMatrix(i, 1) & Format(i, "00#")
                      GridMe.TextMatrix(i, 15) = GridMe.TextMatrix(i, 1)
                      If hErr = "" Then
                           'Error Catch here!
                      End If
                 End If
              End If
           Next i
            txtFields(0).Tag = txtFields(0).Text
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = ""
            Set ColDelete = New Collection
        Else
            ShowDlgMsg MainMenu, "Error!!!<br><br>Tidak dapat menyimpan data!", vbOK, Error, True, False
        End If
     End If
   End If
End If
End Sub

Sub ShowDataBrg(nKey As String, Optional onRow As Integer = -1)
On Error Resume Next
Dim rc As New ADODB.Recordset
Dim lErr As String
lErr = SelectQuery(rc, "Select * FROM mst_barang WHERE [Kode Barang]='" & nKey & "'")
If lErr = "" Then
    If Not rc.EOF Then
       If onRow = -1 Then
          GridMe.TextMatrix(GridMe.Row, 1) = NotNull(rc("Kode Barang"))
          GridMe.TextMatrix(GridMe.Row, 2) = NotNull(rc("Nama Barang"))
          GridMe.TextMatrix(GridMe.Row, 3) = NotNull(rc("Merk")) & " " & NotNull(rc("Type"))
          GridMe.TextMatrix(GridMe.Row, 4) = NotNull(rc("Satuan"))
          GridMe.TextMatrix(GridMe.Row, 5) = fNum(NotNull(rc("Harga Jual")))
       Else
          GridMe.TextMatrix(onRow, 1) = NotNull(rc("Kode Barang"))
          GridMe.TextMatrix(onRow, 2) = NotNull(rc("Nama Barang"))
          GridMe.TextMatrix(onRow, 3) = NotNull(rc("Merk")) & " " & NotNull(rc("Type"))
          GridMe.TextMatrix(onRow, 4) = NotNull(rc("Satuan"))
          GridMe.TextMatrix(onRow, 5) = fNum(NotNull(rc("Harga Jual")))
       End If
    Else
        If onRow = -1 Then
          GridMe.TextMatrix(GridMe.Row, 1) = ""
          GridMe.TextMatrix(GridMe.Row, 2) = ""
          GridMe.TextMatrix(GridMe.Row, 3) = ""
          GridMe.TextMatrix(GridMe.Row, 4) = ""
          GridMe.TextMatrix(GridMe.Row, 5) = ""
        Else
          GridMe.TextMatrix(onRow, 1) = ""
          GridMe.TextMatrix(onRow, 2) = ""
          GridMe.TextMatrix(onRow, 3) = ""
          GridMe.TextMatrix(onRow, 4) = ""
          GridMe.TextMatrix(onRow, 5) = ""
        End If
    End If
    rc.Close
End If
End Sub
Sub InitGrid()
On Error Resume Next
GridMe.MergeCells = flexMergeFixedOnly
GridMe.Rows = 2
GridMe.Rows = 11
Dim i As Integer
For i = 0 To 13
GridMe.MergeCol(i) = True
Next i
GridMe.MergeRow(0) = True
GridMe.MergeRow(1) = True

GridMe.MergeRow(2) = False

GridMe.Cell(flexcpFontBold, 0, 0, 1, 13) = True
GridMe.FrozenCols = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Shift = 0 Then
    Select Case KeyCode
           Case vbKeyEscape
              If picBrg.Visible = False Then
                If txtFields(0).Text = "" Then
                   Unload Me
                Else
                    Set hBtn = Toolbar2.Buttons(6)
                        Toolbar1_ButtonClick hBtn
                        Set hBtn = Nothing
                End If
              Else
                picBrg.Visible = False
                GridMe.Enabled = True
                GridMe.SetFocus
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
txtFields(0).Locked = CekAktifNo("002")
InitGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
PostFindForm "#" & txtFields(0).Hwnd1
PostFindForm "#" & txtFields(12).Hwnd1
PostFindForm "#" & txtFields(15).Hwnd1
PostFindForm "#" & txtFields(5).Hwnd1
PostFindForm "#" & txtFields(11).Hwnd1
End Sub

Private Sub GridFind_DblClick()
On Error Resume Next
ShowDataBrg GridFind.TextMatrix(GridFind.Row, 0)
picBrg.Visible = False
GridMe.Enabled = True
GridMe.Col = 6
GridMe.SetFocus
End Sub

Private Sub GridFind_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyUp And GridFind.Row = 1 Then txtCari.SetFocus
End Sub

Private Sub GridFind_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then GridFind_DblClick
End Sub

Private Sub GridMe_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
Select Case Col
       Case 5
            If Not IsNumeric(GridMe.TextMatrix(Row, Col)) Then
               GridMe.TextMatrix(Row, Col) = ""
            Else
               GridMe.TextMatrix(Row, Col) = fNum(GridMe.TextMatrix(Row, Col))
            End If
            GridMe.TextMatrix(GridMe.Row, 13) = fNum(rNum(GridMe.TextMatrix(GridMe.Row, 5)) * rNum(GridMe.TextMatrix(GridMe.Row, 6)))
       Case 6
            If Not IsNumeric(GridMe.TextMatrix(GridMe.Row, 6)) And GridMe.TextMatrix(GridMe.Row, 6) <> "" Then
               GridMe.TextMatrix(GridMe.Row, 6) = ""
            End If
            GridMe.TextMatrix(GridMe.Row, 13) = fNum(rNum(GridMe.TextMatrix(GridMe.Row, 5)) * rNum(GridMe.TextMatrix(GridMe.Row, 6)))
            GridMe.Col = 7
       Case 7
            If Not IsNumeric(GridMe.TextMatrix(Row, Col)) And GridMe.TextMatrix(Row, Col) <> "" Then
               GridMe.TextMatrix(Row, Col) = ""
            Else
               GridMe.TextMatrix(Row, Col) = GridMe.TextMatrix(Row, Col)
               GridMe.Col = 8
            End If
       Case 8
            GridMe.Col = 9
       Case 9
            If Not IsNumeric(GridMe.TextMatrix(Row, Col)) And GridMe.TextMatrix(Row, Col) <> "" Then
               GridMe.TextMatrix(Row, Col) = ""
            Else
               GridMe.TextMatrix(Row, Col) = fNum(GridMe.TextMatrix(Row, Col))
               GridMe.Col = 10
            End If
       Case 10
            If Not IsNumeric(GridMe.TextMatrix(Row, Col)) And GridMe.TextMatrix(Row, Col) <> "" Then
               GridMe.TextMatrix(Row, Col) = ""
            Else
               GridMe.TextMatrix(Row, Col) = rNum(GridMe.TextMatrix(Row, Col))
               GridMe.Col = 11
            End If
       
'            If Not IsDate(GridMe.TextMatrix(Row, Col)) And GridMe.TextMatrix(Row, Col) <> "" Then
'               GridMe.TextMatrix(Row, Col) = ""
'            Else
'               GridMe.TextMatrix(Row, Col) = Format(GridMe.TextMatrix(Row, Col), "dd/mm/yyyy")
'               GridMe.Col = 11
'            End If
       Case 11
            GridMe.Col = 12
       Case 12
            If GridMe.Row + 1 = GridMe.Rows Then
               GridMe.AddItem GridMe.Rows
            End If
            GridMe.Row = GridMe.Row + 1
            GridMe.Col = 1
 End Select
 AkumulasiJumlah
End Sub

Private Sub GridMe_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
Select Case Col
       Case 1
            If GridMe.TextMatrix(GridMe.Row - 1, 1) = "" Then Cancel = True

       Case 2, 3, 4, 13
            Cancel = True
       Case 5, 6
           If GridMe.TextMatrix(GridMe.Row, 1) = "" Then Cancel = True
          ' GridMe.TextMatrix(Row, Col) = rNum(GridMe.TextMatrix(Row, Col))
End Select
End Sub

Private Sub GridMe_EnterCell()
On Error Resume Next
Select Case GridMe.Col
       Case 2, 3, 4
            'Cancel = True
       Case 5, 9
          If GridMe.TextMatrix(GridMe.Row, GridMe.Col) <> "" Then
           
           GridMe.TextMatrix(GridMe.Row, GridMe.Col) = rNum(GridMe.TextMatrix(GridMe.Row, GridMe.Col))
          End If
End Select
If GridMe.Row > 1 Then GridMe.TextMatrix(GridMe.Row, 0) = GridMe.Row - 1
End Sub

Private Sub GridMe_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case vbKeyInsert
            GridMe.AddItem GridMe.Rows
            SendKeys "{DOWN}", 1
       Case vbKeyDelete
            If GridMe.Rows > 3 Then
               ColDelete.Add txtFields(0).Text & Chr(255) & GridMe.TextMatrix(GridMe.Row, 14)
               GridMe.RemoveItem GridMe.Row
               GridMe.TextMatrix(GridMe.Row, 0) = GridMe.Row
            Else
               ColDelete.Add txtFields(0).Text & Chr(255) & GridMe.TextMatrix(GridMe.Row, 14)
               GridMe.Cell(flexcpText, 2, 1, , 13) = ""
            End If
End Select
End Sub

Private Sub GridMe_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   Select Case Col
          Case 1
            Dim curIsi As String
            curIsi = GridMe.EditText
            If Trim(curIsi) <> "" Then
               Dim lErr As String
               lErr = FindRecord("SELECT [Kode Barang] From mst_barang WHERE [Kode Barang]='" & curIsi & "'")
               If lErr = "1" Then
                  ShowDataBrg curIsi
                  GridMe.Col = 6
               ElseIf lErr = "0" Then
                   GridMe.Enabled = False
                   GridMe.Editable = flexEDNone
                   GridMe.EditText = ""
                   GridMe.TextMatrix(GridMe.Row, 1) = ""
                   GridMe.TextMatrix(GridMe.Row, 2) = ""
                   GridMe.TextMatrix(GridMe.Row, 3) = ""
                   GridMe.TextMatrix(GridMe.Row, 4) = ""
                   GridMe.TextMatrix(GridMe.Row, 5) = ""
                   picBrg.Visible = True
                   picBrg.ZOrder 0
                   txtCari.SetFocus
                   GridMe.Editable = flexEDKbdMouse
               End If
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
       Case 5, 9
           If GridMe.TextMatrix(GridMe.Row, GridMe.Col) <> "" Then
            If Not IsNumeric(GridMe.TextMatrix(GridMe.Row, GridMe.Col)) Then
               GridMe.TextMatrix(GridMe.Row, GridMe.Col) = ""
            Else
               GridMe.TextMatrix(GridMe.Row, GridMe.Col) = fNum(GridMe.TextMatrix(GridMe.Row, GridMe.Col))
            End If
           End If
 End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.index
       Case 1
           If CekUser("08", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            ClearControl Me
            InitGrid
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            If CekAktifNo("002") Then
            txtFields(0).Text = getAutoNo("002")
            txtFields(2).SetFocus
            Else
            txtFields(0).SetFocus
            End If
            txtFields(1).Text = Format(Date, "DD-MMM-YYYY")
            
            Set ColDelete = New Collection
            End If
       Case 2
           If CekUser("08", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
              SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("08", "D") = False Then
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
      Case 8
           Form9.Show
           Form9.ZOrder 0
      Case 12
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
               ShowAllData NotNull(CurRec("No Permohonan")) & "|"
            'End If
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               If CurRec.BOF Then
                  CurRec.MoveNext
               End If
               ShowAllData NotNull(CurRec("No Permohonan")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               If CurRec.EOF Then
                  CurRec.MovePrevious
               End If
               ShowAllData NotNull(CurRec("No Permohonan")) & "|"
            End If
       Case 4
            'If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowAllData NotNull(CurRec("No Permohonan")) & "|"
            'End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowAllData NotNull(CurRec("No Permohonan")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [No Permohonan] From trn_Permohonan_Head ORDER BY [No Permohonan]"
Return
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.index
       Case 1
            GridMe_KeyDown vbKeyInsert, 0
       Case 2
           GridMe_KeyDown vbKeyDelete, 0
       Case 4
            InitGrid
End Select
End Sub

Private Sub txtCari_GotFocus()
On Error Resume Next
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

Private Sub txtFields_DownButtonClick(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            ShowFindForm "SELECT trn_Permohonan_Head.[No Permohonan], trn_Permohonan_Head.[Tgl Permohonan], trn_Permohonan_Head.[Tgl Wawancara], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kelurahan, mst_Pelanggan.Kecamatan, mst_Pelanggan.Kota " & _
                         "FROM mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan] <!where>  " & _
                         "ORDER BY trn_Permohonan_Head.[No Permohonan];", "#" & txtFields(index).Hwnd1, Me, "ShowAllData"
              
              
       Case 11
            ShowFindForm "SELECT mst_Penjamin.[Kode Penjamin], mst_Penjamin.[Nama Penjamin], mst_Penjamin.[Alamat Penjamin], mst_Penjamin.[Kota Penjamin], IIf([mst_Penjamin]![Jenis Penjamin]=1,'Suami',IIf([mst_Penjamin]![Jenis Penjamin]=2,'Istri',IIf([mst_Penjamin]![Jenis Penjamin]=3,'Orang Tua','Lain-Lain'))) AS Penjamin " & _
                         "From mst_Penjamin <!where> ", "#" & txtFields(index).Hwnd1, Me, "ShowPenjamin", " (mst_Penjamin.[Kode Pelanggan]='" & txtFields(5).Text & "') AND "
       
       Case 5
            ShowFindForm "SELECT mst_Pelanggan.[Kode Pelanggan], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, mst_Pelanggan.[Jenis Usaha], mst_Pelanggan.[Nama Perusahaan], mst_Pelanggan.[Alamat Usaha] " & _
                         " FROM mst_Pelanggan <!where> ORDER BY mst_Pelanggan.[Kode Pelanggan]; ", "#" & txtFields(index).Hwnd1, Me, "ShowPelanggan"

       Case 12
            ShowFindForm "SELECT mst_Pegawai.[Kode Pegawai],  mst_Pegawai.[Nama Pegawai],mst_Divisi.Jabatan " & _
                         "FROM mst_Divisi INNER JOIN mst_Pegawai ON mst_Divisi.[Kode Divisi] = mst_Pegawai.[Kode Divisi] " & _
                         " <!where> ;", "#" & txtFields(index).Hwnd1, Me, "ShowSales", " mst_Pegawai.[Kode Divisi]='" & GetDivisi(2) & "' AND "
       Case 15
            ShowFindForm "SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai],mst_Divisi.Jabatan " & _
                         "FROM mst_Divisi INNER JOIN mst_Pegawai ON mst_Divisi.[Kode Divisi] = mst_Pegawai.[Kode Divisi] " & _
                         " <!where> ;", "#" & txtFields(index).Hwnd1, Me, "ShowInspektur", " mst_Pegawai.[Kode Divisi]='" & GetDivisi(1) & "' AND "
End Select
End Sub

Private Sub txtFields_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case 13
            Select Case index
                   Case 5
                      ShowPelanggan txtFields(index).Text & "|"
                   Case 15
                      ShowInspektur txtFields(index).Text & "|"
                   Case 11
                      ShowPenjamin txtFields(index).Text & "|"
                   Case 12
                      ShowSales txtFields(index).Text & "|"
            End Select
       Case Else
            If Me.Tag = "" Then
               Me.Tag = "*"
               Me.Caption = Me.Caption & Me.Tag
            End If
End Select
End Sub

Private Sub txtFields_LostFocus(index As Integer)
On Error Resume Next
         Dim Jumlah As Currency
        Select Case index
                
               Case 19
                   If Val(txtFields(16).Text) <> 0 And txtFields(16).Text <> "" Then
                    If IsNumeric(txtFields(index)) Then
                       Jumlah = (rNum(txtFields(index)) / 100)
                       Jumlah = Jumlah * rNum(txtFields(16))
                       txtFields(20).Text = fNum(Jumlah)
                    End If
                   End If
                   AkumulasiJumlah
               Case 20
                  If Val(txtFields(16).Text) <> 0 And txtFields(16).Text <> "" Then
                   If IsNumeric(txtFields(index)) Then
                      Jumlah = rNum(txtFields(index)) / rNum(txtFields(16))
                      Jumlah = Jumlah * 100
                      txtFields(19).Text = Jumlah
                   End If
                  End If
                  AkumulasiJumlah
             Case 17, 18
                 AkumulasiJumlah
        End Select
End Sub

Private Sub vbButton1_Click()
On Error Resume Next
txtCari_KeyPress 13
End Sub

Private Sub vbButton2_Click()
On Error Resume Next
GridFind_DblClick
End Sub

Private Sub vbButton3_Click()
On Error Resume Next
                picBrg.Visible = False
                GridMe.Enabled = True
                GridMe.SetFocus

End Sub
