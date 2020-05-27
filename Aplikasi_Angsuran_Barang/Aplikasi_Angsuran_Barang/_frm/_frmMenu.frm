VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MainMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SysInfo Nardhika - [ Sistem Informasi Sewa Beli Barang ]"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11505
   Icon            =   "_frmMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "_frmMenu.frx":038A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   5505
      Top             =   3870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   7905
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7567
            Text            =   "RETAIL VERSION  -   vbbego ,© Copyright 2006"
            TextSave        =   "RETAIL VERSION  -   vbbego ,© Copyright 2006"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "_frmMenu.frx":A311
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10-05-2008"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "11:54 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7095
      Top             =   1485
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
            Picture         =   "_frmMenu.frx":A6AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":AA45
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":AFDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":B379
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":B713
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":BAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":C047
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":C3E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":C77B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":CB15
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":CEAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":D249
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":D5E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":D97D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frmMenu.frx":DD17
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   3069
      ButtonWidth     =   1349
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pegawai"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Kons"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "P.Jamin"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Divisi"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Barang"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Libur"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PSBB"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SPSB"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Faktur"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Angs"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageIndex      =   9
            Style           =   5
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Analisis"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Option"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Manual"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tentang"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnufiles 
         Caption         =   "&LogOff..."
         Index           =   0
      End
      Begin VB.Menu mnufiles 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnufiles 
         Caption         =   "&Backup Database"
         Index           =   2
      End
      Begin VB.Menu mnufiles 
         Caption         =   "&Restore Database"
         Index           =   3
      End
      Begin VB.Menu mnufiles 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnufiles 
         Caption         =   "Management User"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnufiles 
         Caption         =   "&Properties"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnufiles 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnufiles 
         Caption         =   "&Keluar"
         Index           =   8
      End
      Begin VB.Menu strComp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "&Compact && Repair Database"
      End
   End
   Begin VB.Menu mnuRef 
      Caption         =   "&Referensi"
      Begin VB.Menu mnuRefs 
         Caption         =   "&1. Pegawai"
         Index           =   0
      End
      Begin VB.Menu mnuRefs 
         Caption         =   "&2. Konsumen"
         Index           =   1
      End
      Begin VB.Menu mnuRefs 
         Caption         =   "&3. Penjamin"
         Index           =   2
      End
      Begin VB.Menu mnuRefs 
         Caption         =   "&4. Divisi"
         Index           =   3
      End
      Begin VB.Menu mnuRefs 
         Caption         =   "&5. Barang"
         Index           =   4
      End
      Begin VB.Menu mnuRefs 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuRefs 
         Caption         =   "&6. Daftar Hari Libur"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "&Transaksi"
      Begin VB.Menu Mnutran 
         Caption         =   "&1. Permohonan Sewa Beli"
         Index           =   0
      End
      Begin VB.Menu Mnutran 
         Caption         =   "&2. Surat Perjanjian"
         Index           =   1
      End
      Begin VB.Menu Mnutran 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Mnutran 
         Caption         =   "&3. Angsuran Sew Beli"
         Index           =   3
      End
      Begin VB.Menu Mnutran 
         Caption         =   "&4. Faktur/ Invoice"
         Index           =   4
      End
   End
   Begin VB.Menu mnuLaps 
      Caption         =   "&Laporan"
      Begin VB.Menu mnuLaporan 
         Caption         =   "&a. Daftar Barang"
         Index           =   0
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&b. Daftar Pegawai"
         Index           =   1
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&c. Daftar Pelanggan"
         Index           =   2
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&d. Daftar Penjamin"
         Index           =   3
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&e. Daftar Divisi"
         Index           =   4
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&f. Laporan Angsuran Konsumen"
         Index           =   6
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&g. Laporan Angsuran Jatuh Tempo"
         Index           =   7
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&h. Laporan Penagihan Angsuran"
         Index           =   8
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&i. Laporan Penerimaan Angsuran"
         Index           =   9
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&j. Laporan Sewa Beli Per Inspektur"
         Index           =   11
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&k. Laporan Sewa Beli Per Salesmen"
         Index           =   12
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&l. Laporan Sewa Beli Per Pelanggan"
         Index           =   13
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&m. Laporan Rekap Pendapatan Bulanan"
         Index           =   15
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "&n. Top 10 Sewa Beli  Barang"
         Index           =   16
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuLaporan 
         Caption         =   "Daftar Hari Libur"
         Index           =   18
      End
   End
   Begin VB.Menu mnuanal 
      Caption         =   "&Analisis"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tool Option"
      Begin VB.Menu mnuconfig 
         Caption         =   "&Sistem Konfigurasi"
      End
      Begin VB.Menu mntud 
         Caption         =   "-"
      End
      Begin VB.Menu mnuta 
         Caption         =   "&Backup Database"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnutf 
         Caption         =   "&Restore Database"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuwin 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelps 
         Caption         =   "&Manual Help"
         Index           =   0
      End
      Begin VB.Menu mnuHelps 
         Caption         =   "&Online Help"
         Index           =   1
      End
      Begin VB.Menu mnuHelps 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelps 
         Caption         =   "&About System"
         Index           =   3
      End
      Begin VB.Menu mnuHelps 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuHelps 
         Caption         =   "&Registered Version"
         Enabled         =   0   'False
         Index           =   5
      End
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TerminateAll As Boolean

Private Sub MDIForm_Initialize()
    On Error Resume Next
    MkDir StripPath(App.Path) & "backup_xdb"
    MkDir StripPath(App.Path) & "log"
    StatusBar1.Panels(3).Text = "Size: " & Format(((FileLen(StripPath(App.Path) & "_dba\_defbasis.xdb") / 1024) / 1024), "##.##") & " MB"
    Toolbar1.Visible = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt3", 1))
    StatusBar1.Visible = Val(GetSetting("vbbego.com\SISRent", "Setting", "opt4", 1))
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
    Dim h(1 To 8) As String * 1
    h(1) = Chr(222)
    h(2) = Chr(222)
    h(3) = Chr(221)
    h(4) = Chr(221)
    h(5) = "r"
    h(6) = "o"
    h(7) = "o"
    h(8) = "t"
    syslog = h(1) & h(2) & h(3) & h(4) & h(5) & h(6) & h(7) & h(8)
GlobalUser = "admin"
Dim hFile As String
hFile = StripPath(App.Path) & "_dba\_defbasis.xdb"
LockUnlock hFile, False
Call LoadDatabase(hFile, syslog)
LockUnlock hFile, True
StatusBar1.Panels(2).Text = GlobalUser
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If TerminateAll = False Then End
End Sub

Private Sub mnuCompact_Click()
On Error GoTo salah
    If srvLogon.State = 1 Then srvLogon.Close
    Dim JRO As New DAO.DBEngine
    Dim hFile As String
    hFile = StripPath(App.Path) & "_dba\_defbasis.xdb"
    LockUnlock hFile, False
    JRO.CompactDatabase hFile, StripPath(App.Path) & "log.mbf", , , ";pwd=" & syslog
    ShowDlgMsg Me, "Repair && Compact database selesai?", vbOK, "", True, False
    DeleteFile hFile
    MoveFile StripPath(App.Path) & "log.mbf", hFile
    On Error Resume Next
    StatusBar1.Panels(3).Text = "Size: " & Format(((FileLen(StripPath(App.Path) & "_dba\_defbasis.xdb") / 1024) / 1024), "##.##") & " MB"
    If srvLogon.State = 0 Then
       srvLogon.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_dba\_defbasis.xdb;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=False"
    End If
    LockUnlock hFile, True
    Exit Sub
salah:
  ShowDlgMsg Me, "Tidak dapat Repair & Compact database, kemungkinan database sedang digunakan.<br>Pastikan form yang aktif ditutup semuanya?", vbOK, Error, True, False
  LockUnlock hFile, True

End Sub

Private Sub mnuconfig_Click()
On Error Resume Next
If GlobalAdmin Then
   Form12.Show 1
Else
    ShowDlgMsg Me, "Hanya Administrator yang dapat menggunakan setting ini", vbOK, , True, False
End If
End Sub

Private Sub mnufiles_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            TerminateAll = True
            GlobalAdmin = False
            GlobalUser = ""
            Unload Me
            Form1.Show
      Case 2
            ShowDlgMsg Me, "Backup file database?", vbYesNo, "", True, False
            If SelectMsg = vbYes Then
               BackUpDB
            End If
       Case 3
        With CD
             .CancelError = True
             On Error GoTo Salah2
             .Filter = "Data base file|*.mbf"
             .ShowOpen
             If BackUpDB(False) Then
                On Error GoTo Salah3
                srvLogon.Close
                Kill StripPath(App.Path) & "\_dba\_defbasis.xdb"
                FileCopy .Filename, StripPath(App.Path) & "_dba\_defbasis.xdb"
                If srvLogon.State = 0 Then
                    LockUnlock StripPath(App.Path) & "_dba\_defbasis.xdb", False
                    srvLogon.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_dba\_defbasis.xdb;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=False"
                    LockUnlock StripPath(App.Path) & "_dba\_defbasis.xdb", True
                End If
                ShowDlgMsg Me, "Restore database selesai?", vbOK, "", True, False
             End If
        End With
       Case 5
            'Form14.Show 1, Me
       Case 6
            'Form17.Show 1, Me
       Case 8
            Unload Me
End Select
Exit Sub
Salah2:
ShowDlgMsg Me, "Tidak dapat merestore database?<br>Silahkan tutup semua form yang aktif<br>" & Error, vbOK, Error, True, False
Exit Sub
Salah3:
ShowDlgMsg Me, "Tidak dapat merestore database?<br>Silahkan tutup semua form yang aktif<br>" & Error, vbOK, Error, True, False
End Sub

Private Sub mnuHelps_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            ShowDlgMsg Me, "File bantuan tidak ada, kemungkinan hilang atau rusak<br><br>" & StripPath(App.Path), vbOK, "", True, False
       Case 1
            Shell "explorer http://www.vb-bego.com/product/mynard/sysrent", 1
       Case 3
            ShowDlgMsg Me, "Copyright by vbBeGo Team 2006<br>Sistem Informasi Sewa Beli Barang<br><br>Team:<br>- Puji Susanto, ST<br>- Ghany Darmawan<br><br>For More Information:<br>- support@vbbego.com<br>- http://www.vb-bego.com", vbOK, , True, False

End Select
End Sub

Private Sub mnuLaporan_Click(index As Integer)
On Error Resume Next
Dim StrSql As String, Form4 As New frm_util_report
Dim isPrint As Boolean, IsReportName As String

Select Case index
       Case 0
           If CekUser("01", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT mst_Barang.[Kode Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type, mst_Barang.[Harga Jual], mst_Barang.Satuan " & _
                     "From mst_Barang <!where> ORDER BY mst_Barang.[Kode Barang];"
            IsReportName = "lap_barang|" & StrSql
            isPrint = True
           End If
       Case 1
           If CekUser("03", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT mst_Pegawai.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], mst_Pegawai.[Tmp Lahir], mst_Pegawai.[Tgl Lahir], mst_Pegawai.Status, mst_Pegawai.Sex, mst_Pegawai.Alamat, mst_Pegawai.[Kode Pos], mst_Pegawai.Kota, mst_Pegawai.Kecamatan, mst_Pegawai.Kelurahan, mst_Pegawai.Telp, mst_Pegawai.Hp, mst_Pegawai.Email, mst_Pegawai.[Tgl Masuk], mst_Pegawai.[Ref ID], mst_Pegawai.[Kode Divisi] " & _
                     "From mst_Pegawai <!where> ORDER BY mst_Pegawai.[Kode Pegawai];"
            IsReportName = "lap_pegawai|" & StrSql
            isPrint = True
           End If
       Case 2
           If CekUser("06", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT mst_Pelanggan.[Kode Pelanggan], mst_Pelanggan.Nama, mst_Pelanggan.[Tgl Lahir], mst_Pelanggan.[Tmp Lahir], IIf([mst_Pelanggan]![Sex]='1','Laki-Laki','Perempuan') " & _
                     "AS [Jenis Kelamin], IIf([mst_Pelanggan]![Status]='1','Menikah','Belum') AS Status, mst_Pelanggan.Alamat, mst_Pelanggan.RT, mst_Pelanggan.RW, mst_Pelanggan.Kelurahan, mst_Pelanggan.Kecamatan, mst_Pelanggan.Kota, mst_Pelanggan.[Kode Pos], mst_Pelanggan.Telp, mst_Pelanggan.[No HP], mst_Pelanggan.[Status Rumah], mst_Pelanggan.[Lama Tinggal], mst_Pelanggan.[Jml Tanggungan], mst_Pelanggan.[Jenis Usaha], mst_Pelanggan.[Nama Perusahaan], mst_Pelanggan.[Bidang Usaha], mst_Pelanggan.[Alamat Usaha], mst_Pelanggan.[RT Usaha], mst_Pelanggan.[RW Usaha], mst_Pelanggan.[Kelurahan Usaha], mst_Pelanggan.[Kecamatan Usaha], mst_Pelanggan.[Kota Usaha], mst_Pelanggan.[Kode Pos Usaha], mst_Pelanggan.[Telp Usaha1], mst_Pelanggan.[Telp Usaha2], mst_Pelanggan.[Fax Usaha], mst_Pelanggan.[Status Usaha],  " & _
                     "mst_Pelanggan.[Lama Usaha], mst_Pelanggan.[Jabatan Usaha], mst_Pelanggan.[Penghasilan Usaha] , mst_Pelanggan.[Penghasilan Jenis Usaha], mst_Pelanggan.[Penghasilan Tambahan], mst_Pelanggan.RefID, mst_Pelanggan.[RefID Jenis] " & _
                     "From mst_Pelanggan <!where> ORDER BY mst_Pelanggan.[Kode Pelanggan];"
            IsReportName = "lap_pelanggan|" & StrSql
            isPrint = True
           End If
       Case 3
           If CekUser("05", "P") = False Then
             ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT mst_Penjamin.[Kode Penjamin], mst_Penjamin.[Kode Pelanggan], mst_Penjamin.[Nama Penjamin], mst_Penjamin.[Jenis Penjamin], mst_Penjamin.[Alamat Penjamin], mst_Penjamin.[RT Penjamin], mst_Penjamin.[RW Penjamin], mst_Penjamin.[Kelurahan Penjamin], mst_Penjamin.[Kecamatan Penjamin], mst_Penjamin.[Kota Penjamin], mst_Penjamin.[Kode Pos Penjamin], mst_Penjamin.[Telp Penjamin], mst_Penjamin.[No HP Penjamin], mst_Penjamin.[Jenis Usaha Penjamin], mst_Penjamin.[Nama Perusahaan Penjamin], mst_Penjamin.[Bidang Usaha Penjamin], mst_Penjamin.[Alamat Usaha Penjamin], mst_Penjamin.[RT Usaha Penjamin], mst_Penjamin.[RW Usaha Penjamin], mst_Penjamin.[Kelurahan Usaha Penjamin], mst_Penjamin.[Kecamatan Usaha Penjamin], mst_Penjamin.[Kota Usaha Penjamin], mst_Penjamin.[Kode Pos Usaha Penjamin], mst_Penjamin.[Telp Usaha1 Penjamin], " & _
                     "mst_Penjamin.[Telp Usaha2 Penjamin], mst_Penjamin.[Fax Usaha Penjamin], mst_Penjamin.[Jabatan Usaha], mst_Penjamin.[Penghasilan Usaha], mst_Penjamin.[Penghasilan Jenis Usaha], mst_Penjamin.[Penghasilan Tambahan] " & _
                     "From mst_Penjamin <!where> ORDER BY mst_Penjamin.[Kode Penjamin];"
            IsReportName = "lap_penjamin|" & StrSql
            isPrint = True
           End If
       Case 4
           If CekUser("02", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT mst_Divisi.[Kode Divisi], mst_Divisi.Jabatan From mst_Divisi <!where> ORDER BY mst_Divisi.[Kode Divisi];"
            IsReportName = "lap_divisi|" & StrSql
            isPrint = True
           End If
       Case 6
           If CekUser("11", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT trn_Angsuran_Head.[No Angsuran], trn_Angsuran_Head.[No Permohonan], trn_Angsuran_Head.[No Barang], trn_Angsuran_Head.[Kode Pegawai], trn_Angsuran_Head.Keterangan, trn_Angsuran_Head.Status, " & _
                     "trn_Angsuran_Head.[Hari Libur], trn_Angsuran_Head.[Hari Sabtu], trn_Angsuran_Detail.[No Bayar], trn_Angsuran_Detail.[Tgl Bayar], trn_Angsuran_Detail.[Tgl Dibayar], trn_Angsuran_Detail.[Jumlah Bayar],  " & _
                     "trn_Angsuran_Detail.Keterangan, trn_Angsuran_Detail.[Kode Pegawai], Val([trn_Angsuran_Detail]![Angsuran Ke]) AS Angsuran, mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, trn_Permohonan_Detail.[Kode Barang],  " & _
                     "[mst_Barang]![Nama Barang] + ' ' + [mst_Barang]![Merk] + ' ' + [mst_Barang]![Type] AS [Nama Brg], trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Awal Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], mst_Barang.Satuan, trn_Permohonan_Head.[Kode Pegawai] AS [Kode Salesmen], trn_Permohonan_Head.[Kode Inspektur] " & _
                     "FROM ((mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) INNER JOIN (trn_Angsuran_Head LEFT JOIN trn_Angsuran_Detail ON trn_Angsuran_Head.[No Angsuran] = trn_Angsuran_Detail.[No Angsuran]) ON trn_Permohonan_Head.[No Permohonan] = trn_Angsuran_Head.[No Permohonan]) INNER JOIN (mst_Barang RIGHT JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] =  " & _
                     "trn_Permohonan_Detail.[Kode Barang]) ON (trn_Angsuran_Head.[No Barang] = trn_Permohonan_Detail.[No Barang]) AND (trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) <!where> " & _
                     "ORDER BY trn_Angsuran_Head.[No Angsuran], Val([trn_Angsuran_Detail]![Angsuran Ke]);"


            IsReportName = "lap_angsuran_konsumen|" & StrSql
            isPrint = True
           End If
       Case 7
           If CekUser("10", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT trn_Angsuran_Head.[No Angsuran], trn_Angsuran_Head.[No Permohonan], trn_Angsuran_Head.[No Barang], trn_Angsuran_Head.[Kode Pegawai], trn_Angsuran_Head.Keterangan, trn_Angsuran_Head.Status, " & _
                     "trn_Angsuran_Head.[Hari Libur], trn_Angsuran_Head.[Hari Sabtu], trn_Angsuran_Detail.[No Bayar], trn_Angsuran_Detail.[Tgl Bayar], trn_Angsuran_Detail.[Tgl Dibayar], trn_Angsuran_Detail.[Jumlah Bayar],  " & _
                     "trn_Angsuran_Detail.Keterangan, trn_Angsuran_Detail.[Kode Pegawai], Val([trn_Angsuran_Detail]![Angsuran Ke]) AS Angsuran, mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, trn_Permohonan_Detail.[Kode Barang],  " & _
                     "[mst_Barang]![Nama Barang] + ' ' + [mst_Barang]![Merk] + ' ' + [mst_Barang]![Type] AS [Nama Brg], trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Awal Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], mst_Barang.Satuan, trn_Permohonan_Head.[Kode Pegawai] AS [Kode Salesmen], trn_Permohonan_Head.[Kode Inspektur] " & _
                     "FROM ((mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) INNER JOIN (trn_Angsuran_Head LEFT JOIN trn_Angsuran_Detail ON trn_Angsuran_Head.[No Angsuran] = trn_Angsuran_Detail.[No Angsuran]) ON trn_Permohonan_Head.[No Permohonan] = trn_Angsuran_Head.[No Permohonan]) INNER JOIN (mst_Barang RIGHT JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] =  " & _
                     "trn_Permohonan_Detail.[Kode Barang]) ON (trn_Angsuran_Head.[No Barang] = trn_Permohonan_Detail.[No Barang]) AND (trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) <!where> " & _
                     "ORDER BY trn_Angsuran_Head.[No Angsuran], Val([trn_Angsuran_Detail]![Angsuran Ke]);"


            IsReportName = "lap_angsuran_jt|" & StrSql & "|TGL BAYARþ<=þ<!date>þANDÿTGL DIBAYARþ IS NULL"
            isPrint = True
           End If
       Case 8
           If CekUser("12", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT trn_Angsuran_Head.[No Angsuran], trn_Angsuran_Head.[No Permohonan], trn_Angsuran_Head.[No Barang], trn_Angsuran_Head.[Kode Pegawai], trn_Angsuran_Head.Keterangan, trn_Angsuran_Head.Status, " & _
                     "trn_Angsuran_Head.[Hari Libur], trn_Angsuran_Head.[Hari Sabtu], trn_Angsuran_Detail.[No Bayar], trn_Angsuran_Detail.[Tgl Bayar], trn_Angsuran_Detail.[Tgl Dibayar], trn_Angsuran_Detail.[Jumlah Bayar],  " & _
                     "trn_Angsuran_Detail.Keterangan, trn_Angsuran_Detail.[Kode Pegawai], Val([trn_Angsuran_Detail]![Angsuran Ke]) AS Angsuran, mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, trn_Permohonan_Detail.[Kode Barang],  " & _
                     "[mst_Barang]![Nama Barang] + ' ' + [mst_Barang]![Merk] + ' ' + [mst_Barang]![Type] AS [Nama Brg], trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Awal Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], mst_Barang.Satuan, trn_Permohonan_Head.[Kode Pegawai] AS [Kode Salesmen], trn_Permohonan_Head.[Kode Inspektur] " & _
                     "FROM ((mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) INNER JOIN (trn_Angsuran_Head LEFT JOIN trn_Angsuran_Detail ON trn_Angsuran_Head.[No Angsuran] = trn_Angsuran_Detail.[No Angsuran]) ON trn_Permohonan_Head.[No Permohonan] = trn_Angsuran_Head.[No Permohonan]) INNER JOIN (mst_Barang RIGHT JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] =  " & _
                     "trn_Permohonan_Detail.[Kode Barang]) ON (trn_Angsuran_Head.[No Barang] = trn_Permohonan_Detail.[No Barang]) AND (trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) <!where> " & _
                     "ORDER BY trn_Angsuran_Head.[No Angsuran], Val([trn_Angsuran_Detail]![Angsuran Ke]);"


            IsReportName = "lap_angsuran_tagihan|" & StrSql & "|TGL BAYARþ=þ<!date>þANDÿTGL DIBAYARþ IS NULL"
            isPrint = True
           End If
       Case 9
           If CekUser("13", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT trn_Angsuran_Head.[No Angsuran], trn_Angsuran_Head.[No Permohonan], trn_Angsuran_Head.[No Barang], " & _
                     "trn_Angsuran_Head.[Kode Pegawai], trn_Angsuran_Head.Keterangan, trn_Angsuran_Head.Status, trn_Angsuran_Head.[Hari Libur], trn_Angsuran_Head.[Hari Sabtu], trn_Angsuran_Detail.[No Bayar], trn_Angsuran_Detail.[Tgl Bayar], trn_Angsuran_Detail.[Tgl Dibayar], trn_Angsuran_Detail.[Jumlah Bayar], trn_Angsuran_Detail.Keterangan, trn_Angsuran_Detail.[Kode Pegawai], Val([trn_Angsuran_Detail]![Angsuran Ke]) AS Angsuran, mst_Pelanggan.Nama, mst_Pelanggan.Alamat, mst_Pelanggan.Kota, trn_Permohonan_Detail.[Kode Barang], [mst_Barang]![Nama Barang] + ' ' + [mst_Barang]![Merk] + ' ' + [mst_Barang]![Type] AS [Nama Brg], trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Awal Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], mst_Barang.Satuan, trn_Permohonan_Head.[Kode Pegawai] AS [Kode Salesmen], trn_Permohonan_Head.[Kode Inspektur] , mst_Pelanggan.Telp, mst_Pelanggan.[No HP] " & _
                     "FROM (mst_Pelanggan RIGHT JOIN trn_Permohonan_Head ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) INNER JOIN ((mst_Barang RIGHT JOIN (trn_Angsuran_Head INNER JOIN trn_Permohonan_Detail ON trn_Angsuran_Head.[No Barang] = trn_Permohonan_Detail.[No Barang]) ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang]) LEFT JOIN trn_Angsuran_Detail ON trn_Angsuran_Head.[No Angsuran] = trn_Angsuran_Detail.[No Angsuran]) ON (trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) AND (trn_Permohonan_Head.[No Permohonan] = trn_Angsuran_Head.[No Permohonan]) <!where> ORDER BY trn_Angsuran_Head.[No Angsuran], Val([trn_Angsuran_Detail]![Angsuran Ke]);"

            IsReportName = "lap_penerimaan_angsuran|" & StrSql & "|TGL BAYARþ=þ<!date>þANDÿTGL DIBAYARþ IS NULL"
            isPrint = True
           End If
       Case 11
           If CekUser("20", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT trn_Permohonan_Head.[No Permohonan], trn_Permohonan_Head.[Tgl Permohonan], trn_Permohonan_Head.[Kode Inspektur], mst_Pegawai.[Nama Pegawai], trn_Permohonan_Head.[Kode Pelanggan], mst_Pelanggan.Nama, trn_Permohonan_Detail.[Kode Barang], [mst_Barang]![Nama Barang]+' '+[mst_Barang]![Merk]+' '+[mst_Barang]![Type] AS NamaBrg, mst_Barang.Satuan, trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, [trn_Permohonan_Detail]![Qty]*[trn_Permohonan_Detail]![Harga Kredit] AS Jumlah_Kredit, trn_Permohonan_Head.[Uang Muka], trn_Permohonan_Head.[Biaya Adm], trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Angsuran JT], trn_Permohonan_Detail.[Awal Angsuran], trn_Perjanjian.[Tgl Perjanjian], trn_Permohonan_Head.Keterangan, trn_Permohonan_Head.[Kode Pegawai] AS [Kode Salesmen] " & _
                     "FROM ((mst_Pelanggan RIGHT JOIN (mst_Barang RIGHT JOIN (trn_Permohonan_Head LEFT JOIN trn_Permohonan_Detail ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) LEFT JOIN trn_Perjanjian ON trn_Permohonan_Head.[No Permohonan] = trn_Perjanjian.[No Permohonan]) LEFT JOIN mst_Pegawai ON trn_Permohonan_Head.[Kode Inspektur] = mst_Pegawai.[Kode Pegawai] <!where> ORDER BY trn_Permohonan_Head.[No Permohonan];"
            
            IsReportName = "lap_penjualan_inspektur|" & StrSql & ""
            isPrint = True
           End If
       Case 12
           If CekUser("22", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT trn_Permohonan_Head.[No Permohonan], trn_Permohonan_Head.[Tgl Permohonan], trn_Permohonan_Head.[Kode Pegawai], mst_Pegawai.[Nama Pegawai], trn_Permohonan_Head.[Kode Pelanggan], mst_Pelanggan.Nama, trn_Permohonan_Detail.[Kode Barang], [mst_Barang]![Nama Barang]+' '+[mst_Barang]![Merk]+' '+[mst_Barang]![Type] AS NamaBrg, mst_Barang.Satuan, trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, [trn_Permohonan_Detail]![Qty]*[trn_Permohonan_Detail]![Harga Kredit] AS Jumlah_Kredit, trn_Permohonan_Head.[Uang Muka], trn_Permohonan_Head.[Biaya Adm], trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Angsuran JT], trn_Permohonan_Detail.[Awal Angsuran], trn_Perjanjian.[Tgl Perjanjian], trn_Permohonan_Head.Keterangan " & _
                     "FROM (mst_Pelanggan RIGHT JOIN ((mst_Barang RIGHT JOIN (trn_Permohonan_Head LEFT JOIN trn_Permohonan_Detail ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang]) LEFT JOIN mst_Pegawai ON trn_Permohonan_Head.[Kode Pegawai] = mst_Pegawai.[Kode Pegawai]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) LEFT JOIN trn_Perjanjian ON trn_Permohonan_Head.[No Permohonan] = trn_Perjanjian.[No Permohonan] <!where> ORDER BY trn_Permohonan_Head.[No Permohonan];"

            IsReportName = "lap_penjualan_salesmen|" & StrSql & ""
            isPrint = True
           End If
       Case 13
           If CekUser("21", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            StrSql = "SELECT trn_Permohonan_Head.[No Permohonan], trn_Permohonan_Head.[Tgl Permohonan], trn_Permohonan_Head.[Kode Inspektur], mst_Pegawai.[Nama Pegawai], trn_Permohonan_Head.[Kode Pelanggan], mst_Pelanggan.Nama, trn_Permohonan_Detail.[Kode Barang], [mst_Barang]![Nama Barang]+' '+[mst_Barang]![Merk]+' '+[mst_Barang]![Type] AS NamaBrg, mst_Barang.Satuan, trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, [trn_Permohonan_Detail]![Qty]*[trn_Permohonan_Detail]![Harga Kredit] AS Jumlah_Kredit, trn_Permohonan_Head.[Uang Muka], trn_Permohonan_Head.[Biaya Adm], trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Angsuran JT], trn_Permohonan_Detail.[Awal Angsuran], trn_Perjanjian.[Tgl Perjanjian], trn_Permohonan_Head.Keterangan, trn_Permohonan_Head.[Kode Pegawai] AS [Kode Salesmen] " & _
                     "FROM ((mst_Pelanggan RIGHT JOIN (mst_Barang RIGHT JOIN (trn_Permohonan_Head LEFT JOIN trn_Permohonan_Detail ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]) ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang]) ON mst_Pelanggan.[Kode Pelanggan] = trn_Permohonan_Head.[Kode Pelanggan]) LEFT JOIN trn_Perjanjian ON trn_Permohonan_Head.[No Permohonan] = trn_Perjanjian.[No Permohonan]) LEFT JOIN mst_Pegawai ON trn_Permohonan_Head.[Kode Inspektur] = mst_Pegawai.[Kode Pegawai] <!where> ORDER BY trn_Permohonan_Head.[No Permohonan];"

            IsReportName = "lap_penjualan_pelanggan|" & StrSql & ""
            isPrint = True
           End If
       Case 15
           If CekUser("23", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
            Form10.Show 1, Me
            If SelectMsg = vbOK Then
               Load Form4
               Set Form4.ARView.ReportSource = Nothing
                   Form4.ARView.Zoom = 100
                   Set Form4.ARView.ReportSource = New lap_RekapPendapatan

               Form4.ARView.Tag = "lap_pendapatan"
               Form4.Grid.Enabled = False
               Form4.ShowField StrSql
               Form4.Show
               Form4.Left = 0
               Form4.Top = 0
               Form4.ZOrder 0
            End If
           End If
      Case 16
           If CekUser("24", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
               Load Form4
               Set Form4.ARView.ReportSource = Nothing
                   Form4.ARView.Zoom = 100
                   Set Form4.ARView.ReportSource = New lap_top10_barang

               Form4.Grid.Enabled = False
               Form4.ShowField StrSql
               Form4.Show
               Form4.Left = 0
               Form4.Top = 0
               Form4.ZOrder 0
          End If
     Case 18
           If CekUser("25", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
           StrSql = "SELECT mst_hari_libur.[Tanggal], mst_hari_libur.Keterangan From mst_hari_libur <!where> ORDER BY mst_hari_libur.[Tanggal];"
           IsReportName = "lap_daftar_tanggal|" & StrSql & ""
           isPrint = True
           End If

End Select
If isPrint Then
   Load Form4
   Form4.ARView.Tag = IsReportName
   Form4.ShowField StrSql
   Form4.Show
   Form4.Left = 0
   Form4.Top = 0
   Form4.ZOrder 0
End If
End Sub

Private Sub mnuRefs_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            Form4.Show
            Form4.ZOrder 0
       Case 1
            Form8.Show
            Form8.ZOrder 0
       Case 2
            Form7.Show
            Form7.ZOrder 0
       Case 3
            Form3.Show
            Form3.ZOrder 0
       Case 4
            Form2.Show
            Form2.ZOrder 0
       Case 6
            Form9.Show
            Form9.ZOrder 0
End Select
End Sub

Private Sub Mnutran_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            Form5.Show
            Form5.ZOrder 0
       Case 1
            Form9.Show
            Form9.ZOrder 0
       Case 3
            Form6.Show
            Form6.ZOrder 0
       Case 4
            Form14.Show
            Form14.ZOrder 0
            
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.index
       Case 1
            Form4.Show
            Form4.ZOrder 0
       Case 2
            Form8.Show
            Form8.ZOrder 0
       Case 3
            Form7.Show
            Form7.ZOrder 0
       Case 4
            Form3.Show
            Form3.ZOrder 0
       Case 5
            Form2.Show
            Form2.ZOrder 0
       Case 6
            Form11.Show
            Form11.ZOrder 0
       Case 7
            Form9.Show
            Form9.ZOrder 0
       Case 8
            Form5.Show
            Form5.ZOrder 0
       Case 9
            Form9.Show
            Form9.ZOrder 0
       Case 10
            Form14.Show
            Form14.ZOrder 0
       Case 11
            Form6.Show
            Form6.ZOrder 0
       Case 16
            If GlobalAdmin Then
               Form12.Show 1
            Else
                ShowDlgMsg Me, "Hanya Administrator yang dapat menggunakan setting ini", vbOK, , True, False
            End If
       Case 18
            ShowDlgMsg Me, "Copyright by vbBeGo Team 2006<br>Sistem Informasi Sewa Beli Barang<br><br>Team:<br>- Puji Susanto, ST<br>- Ghany Darmawan<br><br>For More Information:<br>- support@vbbego.com<br>- http://www.vb-bego.com", vbOK, , True, False
       Case 17
            ShowDlgMsg Me, "File bantuan tidak ada, kemungkinan hilang atau rusak<br><br>" & StripPath(App.Path), vbOK, "", True, False
       Case 20
        Unload Me
End Select
End Sub


Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
PopupMenu mnuLaps, 4
End Sub

Function BackUpDB(Optional showdlg As Boolean = True) As Boolean
On Error GoTo salah
   If srvLogon.State = 1 Then srvLogon.Close
    Dim JRO As New DAO.DBEngine
    Dim hFile As String
    hFile = StripPath(App.Path) & "_dba\_defbasis.xdb"
    LockUnlock hFile, False
    
    Dim myFile As String
    myFile = StripPath(CekBackUp) & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & ".mbf"
    JRO.CompactDatabase hFile, myFile, , , ";pwd=" & syslog
    LockUnlock myFile, True
   If showdlg Then ShowDlgMsg Me, "Backup database selesai?", vbOK, "", True, False
   If srvLogon.State = 0 Then
      srvLogon.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_dba\_defbasis.xdb;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=true"
   End If
   BackUpDB = True
   LockUnlock hFile, True
   Exit Function
salah:
On Error Resume Next
ShowDlgMsg Me, "Tidak dapat membackup database?<br>Silahkan tutup semua form yang aktif<br>" & Error, vbOK, Error, True, False
If srvLogon.State = 0 Then
   srvLogon.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_dba\_defbasis.xdb;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=true"
   LockUnlock StripPath(App.Path) & "_dba\_defbasis.xdb", True
End If
BackUpDB = False
End Function

