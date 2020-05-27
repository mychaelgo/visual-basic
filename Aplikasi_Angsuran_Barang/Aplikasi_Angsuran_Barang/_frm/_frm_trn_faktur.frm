VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faktur Sewa Beli Barang"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_trn_faktur.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form14"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   8895
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   16
      Left            =   6105
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1575
      Width           =   2505
      _ExtentX        =   4419
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
      Left            =   6105
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2505
      _ExtentX        =   4419
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
      Index           =   14
      Left            =   6105
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   825
      Width           =   2505
      _ExtentX        =   4419
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
      Left            =   6105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   30
      Text            =   "_frm_trn_faktur.frx":038A
      Top             =   4635
      Visible         =   0   'False
      Width           =   1155
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   13
      Left            =   1695
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5595
      Width           =   6945
      _ExtentX        =   12250
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
      Index           =   12
      Left            =   1695
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5235
      Width           =   6945
      _ExtentX        =   12250
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
      Index           =   11
      Left            =   1680
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4875
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Alignment       =   1
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
      Index           =   10
      Left            =   1680
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4515
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Alignment       =   1
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
      Left            =   1695
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4155
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Alignment       =   1
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
      Left            =   1695
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3795
      Width           =   6945
      _ExtentX        =   12250
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
      Left            =   1695
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3435
      Width           =   6945
      _ExtentX        =   12250
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
      Left            =   1695
      TabIndex        =   13
      Top             =   3060
      Width           =   6945
      _ExtentX        =   12250
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
      Left            =   1695
      TabIndex        =   5
      Top             =   1575
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   556
      Icon            =   "_frm_trn_faktur.frx":0AE2
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
      Text            =   ""
   End
   Begin SysInfo_Nardhika.vbTextBox txtFields 
      Height          =   315
      Index           =   0
      Left            =   1695
      TabIndex        =   1
      Top             =   825
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   556
      Icon            =   "_frm_trn_faktur.frx":0F30
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
      Left            =   1695
      TabIndex        =   3
      Top             =   1200
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
      Left            =   1695
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2325
      Width           =   6945
      _ExtentX        =   12250
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
      Left            =   1695
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2685
      Width           =   6945
      _ExtentX        =   12250
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
      Left            =   1695
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1965
      Width           =   6945
      _ExtentX        =   12250
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
      Left            =   7890
      Top             =   4410
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
            Picture         =   "_frm_trn_faktur.frx":137E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":1718
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":1AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":1E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":21E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":2580
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":291A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":2CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":304E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":33E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":3782
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":3B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":3EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":4250
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":47EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":4D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":511E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":54B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_trn_faktur.frx":5852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
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
      TabIndex        =   29
      Top             =   6225
      Width           =   8895
      _ExtentX        =   15690
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
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inspektur"
      Height          =   210
      Index           =   16
      Left            =   5160
      TabIndex        =   36
      Top             =   1620
      Width           =   810
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salesmen"
      Height          =   210
      Index           =   15
      Left            =   5160
      TabIndex        =   34
      Top             =   1260
      Width           =   825
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No PSPB"
      Height          =   210
      Index           =   14
      Left            =   5160
      TabIndex        =   31
      Top             =   885
      Width           =   675
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Seri"
      Height          =   210
      Index           =   13
      Left            =   255
      TabIndex        =   26
      Top             =   5670
      Width           =   585
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banyak Satuan"
      Height          =   210
      Index           =   12
      Left            =   255
      TabIndex        =   24
      Top             =   5295
      Width           =   1185
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sisa Tagihan"
      Height          =   210
      Index           =   11
      Left            =   255
      TabIndex        =   22
      Top             =   4920
      Width           =   1035
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uang Muka"
      Height          =   210
      Index           =   10
      Left            =   255
      TabIndex        =   20
      Top             =   4590
      Width           =   900
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Harga"
      Height          =   210
      Index           =   9
      Left            =   255
      TabIndex        =   18
      Top             =   4185
      Width           =   1110
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Merk/Type"
      Height          =   210
      Index           =   8
      Left            =   255
      TabIndex        =   16
      Top             =   3840
      Width           =   885
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      Height          =   210
      Index           =   4
      Left            =   255
      TabIndex        =   14
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      Height          =   210
      Index           =   3
      Left            =   255
      TabIndex        =   12
      Top             =   3135
      Width           =   945
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kota"
      Height          =   210
      Index           =   7
      Left            =   255
      TabIndex        =   10
      Top             =   2730
      Width           =   360
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      Height          =   210
      Index           =   6
      Left            =   255
      TabIndex        =   8
      Top             =   2400
      Width           =   570
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pelanggan"
      Height          =   210
      Index           =   5
      Left            =   255
      TabIndex        =   6
      Top             =   2025
      Width           =   855
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No SPSB"
      Height          =   210
      Index           =   2
      Left            =   255
      TabIndex        =   4
      Top             =   1635
      Width           =   675
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   210
      Index           =   1
      Left            =   255
      TabIndex        =   2
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Faktur"
      Height          =   210
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   900
      Width           =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   -120
      X2              =   19375
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   5
      X1              =   -120
      X2              =   19375
      Y1              =   6165
      Y2              =   6165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   19495
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   19495
      Y1              =   570
      Y2              =   570
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hBtn As MSComctlLib.Button
Dim CurRec As New ADODB.Recordset

Function ShowPegawai(nKey As String) As String
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "Select * from mst_Pegawai WHERE [Kode Pegawai]='" & hKey(0) & "'")
If hErr = "" Then
   If Not rc.EOF Then
    ShowPegawai = NotNull(rc("Nama Pegawai"))
   End If
End If
rc.Close
End Function
Sub ShowFaktur(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT trn_faktur.[No Faktur],trn_faktur.Keterangan, trn_faktur.[Tgl Faktur], trn_faktur.[No Perjanjian] From trn_faktur WHERE (((trn_faktur.[No Faktur])='" & hKey(0) & "'));")
If hErr = "" Then
    If Not rc.EOF Then
        txtFields(0).Text = NotNull(rc("No Faktur"))
        txtFields(0).Tag = NotNull(rc("No Faktur"))
        txtFields(1).Text = NotNull(rc("Tgl Faktur"))
        txtFields(2).Text = NotNull(rc("No Perjanjian"))
        txtFields(6).Text = NotNull(rc("Keterangan"))
        ShowAllData NotNull(rc("No Perjanjian")) & "|"
    Else
kembali:
      ClearControl Me
    End If
Else
   GoSub kembali
End If
rc.Close
End Sub

Sub HapusData(hKey As String)
On Error Resume Next
Dim hErr, h As String
hErr = FindRecord("SELECT trn_faktur.[No Faktur] From trn_faktur WHERE (((trn_faktur.[No Faktur])='" & hKey & "'));")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Faktur?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") = False Then
      GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE From trn_faktur WHERE ([No Faktur]='" & hKey & "');")
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

Sub SimpanData(nKey As String)
On Error Resume Next
Dim h As String, hErr As String, i As Integer
h = FindRecord("SELECT trn_faktur.[No Faktur] From trn_faktur WHERE (((trn_faktur.[No Faktur])='" & nKey & "'));")

If h = "0" Then
   h = SaveRecord("trn_faktur", Array("No Faktur=" & txtFields(0).Text, _
                                      "Tgl Faktur=" & txtFields(1).Text, _
                                      "Keterangan=" & txtFields(6).Text, _
                                      "No Perjanjian=" & txtFields(2).Text))
  If h = "" Then
       If CekAktifNo("009") Then txtFields(0).Text = getAutoNo("009", True)
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
         h = UpdateRecord("trn_faktur", Array("No Faktur=" & txtFields(0).Text, _
                                              "Tgl Faktur=" & txtFields(1).Text, _
                                              "No Perjanjian=" & txtFields(2).Text, _
                                              "Keterangan=" & txtFields(3).Text), " WHERE [No Faktur]='" & txtFields(0).Text & "' ")
    If h = "" Then
         Me.Caption = Replace(Me.Caption, "*", "")
         Me.Tag = ""
         txtFields(0).Tag = txtFields(0).Text
    Else
       ShowDlgMsg Me, "Proses penyimpanan data gagal!", vbOK, h, True, False
    End If
    End If
   End If
End If
End Sub

Sub ShowPelanggan(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")

hErr = SelectQuery(rc, "SELECT trn_Permohonan_Head.[No Permohonan],trn_Permohonan_Head.[Kode Inspektur],trn_Permohonan_Head.[Kode Pegawai], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota " & _
                       "FROM mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]  " & _
                       "WHERE (((trn_Permohonan_Head.[No Permohonan])='" & hKey(0) & "'));")
       
If hErr = "" Then
   If Not rc.EOF Then
    txtFields(3).Text = NotNull(rc("Nama"))
    txtFields(4).Text = NotNull(rc("Alamat"))
    txtFields(5).Text = NotNull(rc("Kota"))
    txtFields(15).Text = ShowPegawai(NotNull(rc("Kode Pegawai")))
    txtFields(16).Text = ShowPegawai(NotNull(rc("Kode Inspektur")))
   Else
    txtFields(3).Text = ""
    txtFields(4).Text = ""
    txtFields(5).Text = ""
   End If
Else
    txtFields(3).Text = ""
    txtFields(4).Text = ""
    txtFields(5).Text = ""
End If
rc.Close
End Sub

Sub ShowDataBarang(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT trn_Permohonan_Detail.[No Permohonan], trn_Permohonan_Detail.[No Barang], trn_Permohonan_Detail.[Kode Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Satuan, trn_Permohonan_Detail.[Harga Kredit], " & _
                       "trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Angsuran JT], trn_Permohonan_Detail.[No Seri], " & _
                       "trn_Permohonan_Detail.Keterangan, trn_Permohonan_Detail.[Awal Angsuran], [harga kredit]*[qty] AS Total, trn_Permohonan_Head.[Kode Pegawai], trn_Permohonan_Head.[Kode Inspektur], trn_Permohonan_Head.[Uang Muka], trn_Permohonan_Head.Disc, trn_Permohonan_Head.[Biaya Adm], mst_Barang.Type " & _
                       "FROM trn_Permohonan_Head INNER JOIN (mst_Barang RIGHT JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang]) ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan] WHERE (((trn_Permohonan_Detail.[No Permohonan])='" & hKey(0) & "'));")

If hErr = "" Then
   If Not rc.EOF Then
      Dim Harga As Currency, Qty As Long
      Dim Barang As String, Merk As String, NoSeri As String
      While Not rc.EOF
            Barang = Barang & NotNull(rc("Nama Barang").Value) & " & "
            Merk = Merk & NotNull(rc("Merk").Value) & " " & NotNull(rc("Type").Value) & " & "
            'txtFields(7).Text = NotNull(RC("Satuan").Value)
            Harga = Harga + Val(NotNull(rc("Total").Value))
            Qty = Qty + Val(NotNull(rc("Qty").Value))
            NoSeri = NoSeri & NotNull(rc("No Seri").Value) & " & "
            txtFields(12).Text = NotNull(rc("Satuan").Value)
            txtFields(10).Text = fNum(NotNull(rc("Uang Muka").Value), False)
            rc.MoveNext
      Wend
      txtFields(7).Text = IIf(Right(Barang, 2) = "& ", Mid(Barang, 1, Len(Barang) - 2), Barang)
      txtFields(8).Text = IIf(Right(Merk, 2) = "& ", Mid(Merk, 1, Len(Merk) - 2), Merk)
      txtFields(13).Text = IIf(Right(NoSeri, 2) = "& ", Mid(NoSeri, 1, Len(NoSeri) - 2), NoSeri)
      txtFields(9).Text = fNum(Harga, False)
      txtFields(11).Text = fNum(Harga - rNum(txtFields(10).Text), False)
      txtFields(12).Text = Qty & " " & txtFields(12).Text
   Else
        ClearControl Me
        Me.Caption = Replace(Me.Caption, "*", "")
        txtFields(0).Tag = ""
        Me.Caption = Me.Caption & Me.Tag
   End If
   rc.Close
End If
End Sub

Sub ShowAllData(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT trn_Perjanjian.[Tgl Mulai],trn_Perjanjian.[No Perjanjian], trn_Perjanjian.[No Permohonan], trn_Perjanjian.Keterangan, trn_Perjanjian.[Kode Pegawai], trn_Perjanjian.Status, trn_Perjanjian.[Tgl Perjanjian] From trn_Perjanjian WHERE trn_Perjanjian.[No Perjanjian]='" & hKey(0) & "';")

If hErr = "" Then
   If Not rc.EOF Then
      txtFields(2).Text = NotNull(rc("No Perjanjian").Value)
      txtFields(2).Tag = NotNull(rc("No Perjanjian").Value)
      txtFields(14).Text = NotNull(rc("No Permohonan").Value)
      ShowPelanggan NotNull(rc("No Permohonan").Value) & "|"
      ShowDataBarang NotNull(rc("No Permohonan").Value) & "|"
   Else
        ClearControl Me
        Me.Caption = Replace(Me.Caption, "*", "")
        txtFields(0).Tag = ""
        Me.Caption = Me.Caption & Me.Tag
   End If
   rc.Close
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
txtFields(0).Locked = CekAktifNo("009")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
 Select Case Button.index
       Case 1
           If CekUser("26", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            ClearControl Me
            Me.Caption = Replace(Me.Caption, "*", "")
            Me.Tag = "*"
            txtFields(0).Tag = ""
            Me.Caption = Me.Caption & Me.Tag
            If CekAktifNo("009") Then
            txtFields(0).Text = getAutoNo("009")
            txtFields(1).SetFocus
            Else
            txtFields(0).SetFocus
            End If
           End If
       Case 2
           If CekUser("26", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
              SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("26", "D") = False Then
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
        If CekUser("26", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
        Else
            Dim hDlg As Boolean
            hDlg = ShowDlgMsg(Me, "Cetak Faktur Sewa/Beli Barang?", vbYesNo, , False, True, , , , , "confirm_" & Me.name)
            If hDlg Then If SelectMsg = vbNo Then Exit Sub
    
           If txtFields(0) <> "" Then
                Dim hText As String, hAlamat As String, hket As String
                hText = txtPrint
                hAlamat = txtFields(4).Text
                hket = txtFields(6).Text
                    Load rpt_buktifaktur
                    With rpt_buktifaktur
                        .Label5.Caption = Replace(.Label5.Caption, String(15, "0"), AddSpace(txtFields(0), 15))
                        .Label6.Caption = Replace(.Label6.Caption, String(12, "1"), AddSpace(" ", 12))
                        .Label8.Caption = Replace(.Label8.Caption, String(52, "a"), AddSpace(txtFields(3), 52))
                        .Label9.Caption = Replace(.Label9.Caption, String(52, "b"), AddSpace(hAlamat, 52))
                        .Label10.Caption = Replace(.Label10.Caption, String(52, "c"), AddSpace(Mid(hAlamat, 53), 52))
                        .Label11.Caption = Replace(.Label11.Caption, String(52, "d"), AddSpace(txtFields(5).Text, 52))
                        .Label8.Caption = Replace(.Label8.Caption, String(19, "e"), AddSpace(txtFields(14), 19)) 'no permohonan
        
                        .Label9.Caption = Replace(.Label9.Caption, String(19, "f"), AddSpace(txtFields(15), 19)) 'salesmen
                        .Label10.Caption = Replace(.Label10.Caption, String(19, "g"), AddSpace(txtFields(16), 19)) 'insp
                        .Label11.Caption = Replace(.Label11.Caption, String(19, "h"), AddSpace(txtFields(2), 19))
                        .Label15.Caption = Replace(.Label15.Caption, String(72, "i"), AddSpace(txtFields(7), 72))
                        .Label16.Caption = Replace(.Label16.Caption, String(72, "j"), AddSpace(txtFields(8), 72))
                        .Label17.Caption = Replace(.Label17.Caption, String(27, "k"), AddSpace("Rp. " & fNum(txtFields(9), True), 27))
                        .Label18.Caption = Replace(.Label18.Caption, String(27, "l"), AddSpace("Rp. " & fNum(txtFields(10), True), 27))
                        .Label19.Caption = Replace(.Label19.Caption, String(27, "m"), AddSpace("Rp. " & fNum(txtFields(11), True), 27))
                        .Label18.Caption = Replace(.Label18.Caption, String(25, "n"), AddSpace(txtFields(12), 25))
                        .Label19.Caption = Replace(.Label19.Caption, String(25, "o"), AddSpace(txtFields(13), 25))
                        
                        .Label20.Caption = Replace(.Label20.Caption, String(72, "p"), AddSpace(hket, 72))
                        .Label21.Caption = Replace(.Label21.Caption, String(72, "s"), AddSpace(Trim(Mid(hket, 73)), 72))
                        
                        .Label24.Caption = Replace(.Label24.Caption, String(14, "q"), AddSpace(txtFields(1), 14))
                        .Label32.Caption = Replace(.Label32.Caption, String(30, "r"), AddSpace(txtFields(3), 30))
                          .PrintReport False
                    End With
                    Unload rpt_buktifaktur
                
'                hText = Replace(hText, String(15, "0"), AddSpace(txtFields(0), 15))
'                hText = Replace(hText, String(12, "1"), AddSpace(" ", 12))
'                hText = Replace(hText, String(52, "a"), AddSpace(txtFields(3), 52))
'                hText = Replace(hText, String(52, "b"), AddSpace(hAlamat, 52))
'                hText = Replace(hText, String(52, "c"), AddSpace(Mid(hAlamat, 53), 52))
'                hText = Replace(hText, String(52, "d"), AddSpace(txtFields(5).Text, 52))
'                hText = Replace(hText, String(19, "e"), AddSpace(txtFields(14), 19)) 'no permohonan
'
'                hText = Replace(hText, String(19, "f"), AddSpace(txtFields(15), 19))  'salesmen
'                hText = Replace(hText, String(19, "g"), AddSpace(txtFields(16), 19)) 'insp
'                hText = Replace(hText, String(19, "h"), AddSpace(txtFields(2), 19))
'                hText = Replace(hText, String(72, "i"), AddSpace(txtFields(7), 72))
'                hText = Replace(hText, String(72, "j"), AddSpace(txtFields(8), 72))
'                hText = Replace(hText, String(27, "k"), AddSpace("Rp. " & fNum(txtFields(9), False) & ",-", 27))
'                hText = Replace(hText, String(27, "l"), AddSpace("Rp. " & fNum(txtFields(10), False) & ",-", 27))
'                hText = Replace(hText, String(27, "m"), AddSpace("Rp. " & fNum(txtFields(11), False) & ",-", 27))
'                hText = Replace(hText, String(25, "n"), AddSpace(txtFields(12), 25))
'                hText = Replace(hText, String(25, "o"), AddSpace(txtFields(13), 25))
'                hText = Replace(hText, String(72, "p"), AddSpace(txtFields(6), 72))
'                hText = Replace(hText, String(14, "q"), AddSpace(txtFields(1), 14))
'                hText = Replace(hText, String(19, "r"), AddSpace(txtFields(3), 19))
'                PrintText hText
                hText = ""
           Else
                ShowDlgMsg Me, "Tidak ada data faktur yang akan dicetak!", vbOK, , True, False
           End If
        End If
       Case 8
'           If CekUser("03", "N") = False Then
'              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
'           Else
'            Dim StrSql As String, Form4 As New frm_util_report
'            Load Form4
'            StrSql = "SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], mst_Pegawai.[Tmp Lahir], mst_Pegawai.[Tgl Lahir], mst_Pegawai.Status, mst_Pegawai.Sex, mst_Pegawai.Alamat, mst_Pegawai.[Kode Pos], mst_Pegawai.Kota, mst_Pegawai.Kecamatan, mst_Pegawai.Kelurahan, mst_Pegawai.Telp, mst_Pegawai.Hp, mst_Pegawai.Email, mst_Pegawai.[Tgl Masuk], mst_Pegawai.[Ref ID], mst_Pegawai.[Kode Divisi] " & _
'                     "From mst_Pegawai  <!where> ORDER BY mst_Pegawai.[Kode Pegawai];"
'
'            Form4.ARView.Tag = "lap_pegawai|" & StrSql
'            Form4.ShowField StrSql
'            Form4.Show
'            Form4.Left = 0
'            Form4.Top = 0
'            Form4.ZOrder 0
'           End If
       Case 8
            
       Case 9
           
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
               ShowFaktur NotNull(CurRec("No Faktur")) & "|"
            'End If
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               If CurRec.BOF Then
                  CurRec.MoveNext
               End If
               ShowFaktur NotNull(CurRec("No Faktur")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               If CurRec.EOF Then
                  CurRec.MovePrevious
               End If
               ShowFaktur NotNull(CurRec("No Faktur")) & "|"
            End If
       Case 4
            'If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowFaktur NotNull(CurRec("No Faktur")) & "|"
            'End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowFaktur NotNull(CurRec("No Faktur")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [No Faktur] From trn_faktur ORDER BY [No Faktur]"
Return
End Sub

Private Sub txtFields_DownButtonClick(index As Integer)
Select Case index
       Case 0
    
            ShowFindForm "SELECT trn_faktur.[No Faktur], trn_faktur.[Tgl Faktur], trn_faktur.[No Perjanjian], trn_faktur.Keterangan, mst_Pelanggan.Nama, mst_Pelanggan.Alamat " & _
                         "FROM mst_Pelanggan INNER JOIN (trn_Permohonan_Head INNER JOIN (trn_Perjanjian INNER JOIN trn_faktur ON trn_Perjanjian.[No Perjanjian] = trn_faktur.[No Perjanjian]) ON trn_Permohonan_Head.[No Permohonan] = trn_Perjanjian.[No Permohonan]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan] " & _
                         " <!where> ORDER BY trn_faktur.[No Faktur];", "#" & txtFields(index).Hwnd1, Me, "ShowFaktur"
       
       Case 2
            ShowFindForm "SELECT trn_Perjanjian.[No Perjanjian], mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, mst_Pelanggan.[Kode Pos], trn_Perjanjian.[Tgl Perjanjian] " & _
                         "FROM mst_Pelanggan RIGHT JOIN (trn_Permohonan_Head RIGHT JOIN trn_Perjanjian ON trn_Permohonan_Head.[No Permohonan] = trn_Perjanjian.[No Permohonan]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan] " & _
                         " <!where> ", "#" & txtFields(index).Hwnd1, Me, "ShowAllData"
       
End Select

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

