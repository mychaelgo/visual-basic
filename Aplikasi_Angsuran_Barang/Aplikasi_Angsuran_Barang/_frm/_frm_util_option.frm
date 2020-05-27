VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_util_option.frx":0000
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6735
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_option.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_option.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_option.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_option.frx":1258
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_option.frx":15F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_option.frx":198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_option.frx":1D26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SysInfo_Nardhika.vbButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   105
      TabIndex        =   35
      Top             =   5355
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   635
      BTYPE           =   5
      TX              =   "&Close"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "_frm_util_option.frx":20C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TreeView Editor 
      Height          =   5115
      Left            =   105
      TabIndex        =   33
      Top             =   105
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   9022
      _Version        =   393217
      Indentation     =   2
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "Company Profile"
      Height          =   5715
      Left            =   2820
      TabIndex        =   42
      Top             =   15
      Visible         =   0   'False
      Width           =   6105
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   4845
         Width           =   3735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   4425
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4005
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3615
         Width           =   3735
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   4815
         Width           =   1005
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   4410
         Width           =   1005
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4005
         Width           =   1005
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3600
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1935
         TabIndex        =   23
         Top             =   2760
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   1935
         TabIndex        =   21
         Top             =   2400
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1935
         TabIndex        =   19
         Top             =   2040
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1935
         TabIndex        =   17
         Top             =   1695
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1935
         TabIndex        =   15
         Top             =   1350
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1935
         TabIndex        =   13
         Top             =   1005
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1935
         TabIndex        =   11
         Top             =   660
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1935
         TabIndex        =   9
         Top             =   330
         Width           =   3990
      End
      Begin SysInfo_Nardhika.vbButton vbButton2 
         Height          =   360
         Left            =   4560
         TabIndex        =   32
         Top             =   5235
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BTYPE           =   5
         TX              =   "&Update"
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
         MPTR            =   99
         MICON           =   "_frm_util_option.frx":23DA
         PICN            =   "_frm_util_option.frx":26F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Divisi"
         Height          =   195
         Left            =   2175
         TabIndex        =   55
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Left            =   1140
         TabIndex        =   54
         Top             =   3360
         Width           =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         Index           =   0
         X1              =   135
         X2              =   4875
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Staff"
         Height          =   195
         Left            =   5040
         TabIndex        =   53
         Top             =   3135
         Width           =   915
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kolektor"
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   4875
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesmen"
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   4470
         Width           =   675
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inspektur"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   4065
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   3660
         Width           =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   135
         X2              =   4875
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Site"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2790
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2460
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2115
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1845
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1035
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Setup Data"
      Height          =   5715
      Left            =   2820
      TabIndex        =   45
      Top             =   15
      Visible         =   0   'False
      Width           =   6105
      Begin SysInfo_Nardhika.vbButton cmdupdateno 
         Height          =   360
         Left            =   4875
         TabIndex        =   48
         Top             =   5220
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         BTYPE           =   5
         TX              =   "&Update"
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
         MICON           =   "_frm_util_option.frx":2A8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox grid 
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   4380
         Left            =   135
         ScaleHeight     =   4320
         ScaleWidth      =   5775
         TabIndex        =   46
         Top             =   765
         Width           =   5835
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"_frm_util_option.frx":2AAA
         Height          =   435
         Left            =   135
         TabIndex        =   47
         Top             =   255
         Width           =   5715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "General Setting"
      Height          =   5715
      Left            =   2820
      TabIndex        =   41
      Top             =   15
      Width           =   6105
      Begin VB.CheckBox Check5 
         Caption         =   "Gunakan Tanggal Komputer (Apabila tgl kosong)"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   1695
         Width           =   4425
      End
      Begin VB.ComboBox cboPrint 
         Height          =   315
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   4365
         Width           =   5805
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2865
         Width           =   5805
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Buka Report Di Window Baru"
         Height          =   195
         Left            =   225
         TabIndex        =   0
         Top             =   405
         Width           =   2670
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Tampilkan Nama Login Terakhir"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   705
         Width           =   2670
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Tampilkan Toolbar"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   1035
         Width           =   2670
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Tampilkan StatusBar"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   1380
         Width           =   2670
      End
      Begin SysInfo_Nardhika.vbButton vbButton3 
         Height          =   390
         Left            =   4740
         TabIndex        =   7
         Top             =   3270
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "&Browse"
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
         MPTR            =   99
         MICON           =   "_frm_util_option.frx":2B3F
         PICN            =   "_frm_util_option.frx":2E59
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer Redirect Yang Digunakan"
         Height          =   195
         Left            =   195
         TabIndex        =   43
         Top             =   4095
         Width           =   2325
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   120
         X2              =   5940
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         Index           =   2
         X1              =   135
         X2              =   5955
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Directory"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         Index           =   1
         X1              =   135
         X2              =   5955
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   120
         X2              =   5940
         Y1              =   2235
         Y2              =   2235
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Management"
      Height          =   5715
      Left            =   2820
      TabIndex        =   34
      Top             =   15
      Visible         =   0   'False
      Width           =   6105
      Begin SysInfo_Nardhika.vbButton cmdAdd 
         Height          =   375
         Left            =   2895
         TabIndex        =   38
         Top             =   270
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "&Tambah"
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
         MICON           =   "_frm_util_option.frx":31F3
         PICN            =   "_frm_util_option.frx":350D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvUser 
         Height          =   4845
         Left            =   120
         TabIndex        =   36
         Top             =   735
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   8546
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User Login"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
      End
      Begin SysInfo_Nardhika.vbButton cmdEdit 
         Height          =   375
         Left            =   3945
         TabIndex        =   39
         Top             =   270
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "&Edit"
         ENAB            =   0   'False
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
         MICON           =   "_frm_util_option.frx":38A7
         PICN            =   "_frm_util_option.frx":3BC1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton cmdHapus 
         Height          =   375
         Left            =   4950
         TabIndex        =   40
         Top             =   270
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "&Hapus"
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
         MICON           =   "_frm_util_option.frx":3F5B
         PICN            =   "_frm_util_option.frx":4275
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daftar User"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   510
         Width           =   840
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ShowDivisi(nKey As String, obj1 As Object, obj2 As Object)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT mst_divisi.[Kode Divisi],  mst_divisi.[Jabatan] from mst_divisi where [Kode Divisi]='" & hKey(0) & "' ORDER BY mst_divisi.[Kode Divisi]; ")
If hErr = "" Then
    If Not rc.EOF Then
        'obj1.Text = NotNull(rc("Kode Divisi"))
        obj2.Text = NotNull(rc("Jabatan"))
    Else
kembali:
       'obj1.Text = ""
       obj2.Text = ""
End If
Else
   GoSub kembali
End If
rc.Close
End Sub

Sub ShowOnCombo()
On Error Resume Next
Dim rc As New ADODB.Recordset, hErr As String
hErr = SelectQuery(rc, "Select * from mst_divisi order by [Kode Divisi]")
If hErr = "" Then
   If Not rc.EOF Then
      While Not rc.EOF
          Combo1.AddItem NotNull(rc("Kode Divisi"))
          Combo2.AddItem NotNull(rc("Kode Divisi"))
          Combo3.AddItem NotNull(rc("Kode Divisi"))
          Combo4.AddItem NotNull(rc("Kode Divisi"))
          rc.MoveNext
      Wend
   End If
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
Form13.Show 1, Me
ShowUser
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
SaveSetting "vbbego.com\SISRent", "Setting", "opt1", Check1.Value
SaveSetting "vbbego.com\SISRent", "Setting", "opt2", Check2.Value
SaveSetting "vbbego.com\SISRent", "Setting", "opt3", Check3.Value
SaveSetting "vbbego.com\SISRent", "Setting", "opt4", Check4.Value
SaveSetting "vbbego.com\SISRent", "Setting", "opt5", Check5.Value
SaveSetting "vbbego.com\SISRent", "Setting", "PrintRedirect", cboPrint.Text
If Text2 <> "" Then
   SaveSetting "vbbego.com\SISRent", "Setting", "backuppath", Text2
End If
Unload Me
MainMenu.StatusBar1.Panels(3).Text = "Size: " & Format(((FileLen(StripPath(App.Path) & "_dba\_defbasis.xdb") / 1024) / 1024), "##.##") & " MB"
MainMenu.Toolbar1.Visible = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt3", 1))
MainMenu.StatusBar1.Visible = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt4", 1))
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
   Form13.ShowData lvUser.SelectedItem.Text
   Form13.Show 1, Me
   ShowUser
End Sub

Private Sub cmdHapus_Click()
On Error Resume Next
If lvUser.ListItems.Count > 0 Then
    If LCase(lvUser.SelectedItem.Text) <> "admin" Then
       ShowDlgMsg Me, "Anda yakin mau menghapusnya?", vbYesNo, , True, False
       If SelectMsg = vbYes Then
          srvUSER.Execute "DELETE FROM users WHERE (login = '" & AllowChar(lvUser.SelectedItem.Text) & "')"
          ShowUser
       End If
    Else
       ShowDlgMsg Me, "Administrator tidak dapat dihapus", vbOK, , True, False
    End If
End If
End Sub

Private Sub cmdupdateno_Click()
On Error Resume Next
Dim h As String
Dim i As Integer
For i = 1 To Grid.Rows - 1
     h = FindRecord("SELECT mst_Nomor.[Kode No] From mst_Nomor WHERE (((mst_Nomor.[Kode No])='" & Grid.TextMatrix(i, 0) & "'));")
     If h = "1" Then
        UpdateRecord "mst_Nomor", Array("Keterangan=" & Grid.TextMatrix(i, 1), _
                                        "FormatNo=" & Grid.TextMatrix(i, 2), _
                                        "#LenNo=" & Grid.TextMatrix(i, 3), _
                                        "LastYear=" & Grid.TextMatrix(i, 4), _
                                        "LastMonth=" & Grid.TextMatrix(i, 5), _
                                        "#LastNo=" & Grid.TextMatrix(i, 6), _
                                        "^aktif=" & Grid.TextMatrix(i, 8), _
                                        "ChangeNo=" & Grid.TextMatrix(i, 7)), " WHERE [Kode No]='" & Grid.TextMatrix(i, 0) & "' "
    ElseIf h = "0" Then
        SaveRecord "mst_Nomor", Array("Kode No=" & Grid.TextMatrix(i, 0), _
                                        "Keterangan=" & Grid.TextMatrix(i, 1), _
                                        "FormatNo=" & Grid.TextMatrix(i, 2), _
                                        "#LenNo=" & Grid.TextMatrix(i, 3), _
                                        "LastYear=" & Grid.TextMatrix(i, 4), _
                                        "LastMonth=" & Grid.TextMatrix(i, 5), _
                                        "#LastNo=" & Grid.TextMatrix(i, 6), _
                                        "^aktif=" & Grid.TextMatrix(i, 8), _
                                        "ChangeNo=" & Grid.TextMatrix(i, 7)), " WHERE [Kode No]='" & Grid.TextMatrix(i, 0) & "' "
    End If
Next i
End Sub

Private Sub Combo1_Click()
ShowDivisi Combo1.Text, Combo1, Text3
End Sub

Private Sub Combo2_Click()
ShowDivisi Combo2.Text, Combo2, Text4
End Sub

Private Sub Combo3_Click()
ShowDivisi Combo3.Text, Combo3, Text5
End Sub

Private Sub Combo4_Click()
ShowDivisi Combo4.Text, Combo4, Text6
End Sub

Private Sub Editor_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
If Node.Key = "root_general" Then
   Frame1.Visible = False
   Frame2.Visible = True
   Frame3.Visible = False
   Frame4.Visible = False
ElseIf Node.Key = "root_profile" Then
   Frame2.Visible = False
   Frame1.Visible = False
   Frame3.Visible = True
   Frame4.Visible = False
ElseIf Node.Key = "root_user" Then
   Frame1.Visible = True
   Frame2.Visible = False
   Frame3.Visible = False
   Frame4.Visible = False
ElseIf Node.Key = "root_database" Then
   Frame1.Visible = False
   Frame2.Visible = False
   Frame3.Visible = False
   Frame4.Visible = True
   LoadData
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Node As Node
    Set Node = Editor.Nodes.Add(, , "root", "Options", 4)
        Set Node = Editor.Nodes.Add("root", tvwChild, "root_general", "General", 5)
            Node.EnsureVisible
        Set Node = Editor.Nodes.Add("root", tvwChild, "root_user", "User Setting", 1)
            Node.EnsureVisible
        Set Node = Editor.Nodes.Add("root", tvwChild, "root_profile", "Company Profile", 6)
            Node.EnsureVisible
        Set Node = Editor.Nodes.Add("root", tvwChild, "root_database", "Setup Data", 7)
            Node.EnsureVisible

'  Dim h(1 To 8) As String * 1
'    h(1) = Chr(222)
'    h(2) = Chr(222)
'    h(3) = Chr(221)
'    h(4) = Chr(221)
'    h(5) = "r"
'    h(6) = "o"
'    h(7) = "o"
'    h(8) = "t"
'    syslog = h(1) & h(2) & h(3) & h(4) & h(5) & h(6) & h(7) & h(8)
'LockUnlock StripPath(App.Path) & "_support\_syslog.sys", False
'srvUSER.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_support\_syslog.sys;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=False"
'
'Dim hFile As String
'hFile = StripPath(App.Path) & "_dba\_defbasis.xdb"
'LockUnlock hFile, False
'Call LoadDatabase(hFile, syslog)
'LockUnlock hFile, True


ShowUser

Dim i As Integer
For i = 0 To Text1.Count - 1
  Text1(i).Text = GetSetting("vbbego.com\SISRent", "Setting", "Profile" & i)
Next i
Text2 = GetSetting("vbbego.com\SISRent", "Setting", "backuppath", StripPath(App.Path) & "backup_xdb")
Check1.Value = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt1", 1))
Check2.Value = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt2", 1))
Check3.Value = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt3", 1))
Check4.Value = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt4", 1))
Check5.Value = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt5", 0))
HideMenu Me.hWnd
ShowOnCombo
ShowCurrrentCombo
LoadPrinterName
End Sub

Sub LoadPrinterName()
On Error Resume Next
Dim h As Integer
Set LocalPrinter = New Collection
GetPrinterName LocalPrinter
If LocalPrinter.Count > 0 Then
   For h = 1 To LocalPrinter.Count
      cboPrint.AddItem LocalPrinter(h)
   Next h
   cboPrint.ListIndex = 0
Else
   ShowDlgMsg Me, "Tidak ada printer yang terpasang di komputer anda", vbOK, , True, False
   Unload Me
End If
cboPrint.Text = GetSetting("vbbego.com\SISRent", "Setting", "PrintRedirect")
End Sub
Sub ShowUser()
 Dim lv As ListItem
 Dim rc As New ADODB.Recordset
 Dim hErr As String
 Dim hIcon As Integer
 hErr = SelectQuery(rc, "SELECT * From Users ORDER BY users.login;", True, srvUSER)
 If hErr = "" Then
    lvUser.ListItems.Clear
    If Not rc.EOF Then
       While Not rc.EOF
            hIcon = Val(NotNull(rc("admin").Value))
            hIcon = IIf(hIcon = 0, 1, 3)
            
            hIcon = IIf(Val(NotNull(rc("aktif").Value)) = 0, 2, hIcon)
            
            Set lv = lvUser.ListItems.Add(, , NotNull(rc("login").Value), , hIcon)
                lv.SubItems(1) = NotNull(rc("description").Value)
          rc.MoveNext
       Wend
    End If
 End If
End Sub

Sub ShowCurrrentCombo()
On Error Resume Next
Dim rc As New ADODB.Recordset
Dim hErr As String
hErr = SelectQuery(rc, "Select * from settings where nama='divisi'", True, srvUSER)
If hErr = "" Then
   If Not rc.EOF Then
      Dim h
      If Trim(NotNull(rc("isi"))) <> "" Then
         h = Split(NotNull(rc("isi")), ";")
         If UBound(h) > 0 Then
            Combo1.Text = h(0)
            Combo2.Text = h(1)
            Combo3.Text = h(2)
            Combo4.Text = h(3)
            ShowDivisi Combo1.Text, Combo1, Text3
            ShowDivisi Combo2.Text, Combo2, Text4
            ShowDivisi Combo3.Text, Combo3, Text5
            ShowDivisi Combo4.Text, Combo4, Text6
         End If
      End If
   End If
End If
End Sub

Private Sub grid_EnterCell()
'MsgBox grid.TextMatrix(grid.Row, grid.Col)
End Sub

Private Sub lvUser_DblClick()
On Error Resume Next
If lvUser.ListItems.Count > 0 Then
   Form13.ShowData lvUser.SelectedItem.Text
   Form13.Show 1
End If
End Sub

Private Sub lvUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
cmdEdit.Enabled = True
End Sub

Private Sub vbButton2_Click()
On Error Resume Next
Dim i As Integer
For i = 0 To Text1.Count - 1
  SaveSetting "vbbego.com\SISRent", "Setting", "Profile" & i, Text1(i).Text
Next i
srvUSER.Execute "DELETE FROM settings Where Nama='divisi'"
srvUSER.Execute "INSERT INTO settings (Nama,Isi) Values('divisi','" & Combo1.Text & ";" & Combo2.Text & ";" & Combo3.Text & ";" & Combo4.Text & "')"
End Sub

Private Sub vbButton3_Click()
On Error Resume Next
Dim h As String
h = GetFolderBrowse(hWnd)
If Trim(h) <> "" Then
    Text2 = h
End If
End Sub

Sub LoadData()
On Error Resume Next
Dim rc As New ADODB.Recordset, pos As Long
rc.Open "select * from mst_Nomor order by [Kode No]", srvLogon
Grid.Rows = 1
pos = 1
If Not rc.EOF Then
   While Not rc.EOF
         Grid.AddItem NotNull(rc("Kode No"))
         Grid.TextMatrix(pos, 1) = NotNull(rc("Keterangan"))
         Grid.TextMatrix(pos, 2) = NotNull(rc("FormatNo"))
         Grid.TextMatrix(pos, 3) = NotNull(rc("LenNo"))
         Grid.TextMatrix(pos, 4) = NotNull(rc("LastYear"))
         Grid.TextMatrix(pos, 5) = NotNull(rc("LastMonth"))
         Grid.TextMatrix(pos, 6) = NotNull(rc("LastNo"))
         Grid.TextMatrix(pos, 7) = NotNull(rc("ChangeNo"))
         Grid.TextMatrix(pos, 8) = NotNull(rc("Aktif"))
         pos = pos + 1
         rc.MoveNext
   Wend
End If
rc.Close
Set rc = Nothing
End Sub
