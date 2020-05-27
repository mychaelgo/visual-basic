VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_util_report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Laporan"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_util_report.frx":0000
   LinkTopic       =   "_frm_util_report"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11835
   Begin SysInfo_Nardhika.net_Resize net_Resize1 
      Left            =   5670
      Top             =   3045
      _ExtentX        =   847
      _ExtentY        =   847
      KeepRatio       =   -1  'True
   End
   Begin VB.PictureBox piclap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   735
      ScaleHeight     =   690
      ScaleWidth      =   4065
      TabIndex        =   14
      Top             =   1230
      Visible         =   0   'False
      Width           =   4065
      Begin SysInfo_Nardhika.vbButton vbButton1 
         Height          =   690
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1217
         BTYPE           =   7
         TX              =   "Sedang pembuatan laporan, tunggu..."
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
         BCOL            =   -2147483644
         BCOLO           =   -2147483644
         FCOL            =   16576
         FCOLO           =   16576
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "_frm_util_report.frx":058A
         PICN            =   "_frm_util_report.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5130
      Top             =   3735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   45
      TabIndex        =   1
      Top             =   4905
      Width           =   11730
      Begin VB.ComboBox cboGroup 
         Height          =   315
         Left            =   5475
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1260
         Width           =   2250
      End
      Begin VB.TextBox tpriode 
         Height          =   285
         Left            =   1395
         TabIndex        =   13
         Top             =   1260
         Width           =   3570
      End
      Begin VB.PictureBox Grid 
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   90
         ScaleHeight     =   945
         ScaleWidth      =   7575
         TabIndex        =   2
         Top             =   195
         Width           =   7635
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   0
         Left            =   7815
         TabIndex        =   3
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Tambah"
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
         MICON           =   "_frm_util_report.frx":0940
         PICN            =   "_frm_util_report.frx":0C5A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   4
         Left            =   7815
         TabIndex        =   4
         Top             =   675
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Hapus"
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
         MICON           =   "_frm_util_report.frx":0FF4
         PICN            =   "_frm_util_report.frx":130E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   1
         Left            =   7815
         TabIndex        =   5
         Top             =   1110
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Reset"
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
         MICON           =   "_frm_util_report.frx":16A8
         PICN            =   "_frm_util_report.frx":19C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   2
         Left            =   9075
         TabIndex        =   6
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Muat"
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
         MICON           =   "_frm_util_report.frx":1D5C
         PICN            =   "_frm_util_report.frx":2076
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   3
         Left            =   9075
         TabIndex        =   7
         Top             =   675
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Simpan"
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
         MICON           =   "_frm_util_report.frx":2410
         PICN            =   "_frm_util_report.frx":272A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   5
         Left            =   9075
         TabIndex        =   8
         Top             =   1110
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Cetak"
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
         MICON           =   "_frm_util_report.frx":2AC4
         PICN            =   "_frm_util_report.frx":2DDE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   6
         Left            =   10440
         TabIndex        =   9
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Proses"
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
         MICON           =   "_frm_util_report.frx":3378
         PICN            =   "_frm_util_report.frx":3692
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   7
         Left            =   10440
         TabIndex        =   10
         Top             =   675
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Export"
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
         MICON           =   "_frm_util_report.frx":3A2C
         PICN            =   "_frm_util_report.frx":3D46
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   8
         Left            =   10440
         TabIndex        =   11
         Top             =   1110
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&Keluar"
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
         MICON           =   "_frm_util_report.frx":41A0
         PICN            =   "_frm_util_report.frx":44BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grup"
         Height          =   195
         Left            =   5040
         TabIndex        =   16
         Top             =   1305
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periode Laporan:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   1290
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   1215
         Left            =   10335
         Top             =   255
         Width           =   15
      End
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARView 
      Height          =   4860
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   8573
      SectionData     =   "_frm_util_report.frx":4A54
   End
End
Attribute VB_Name = "frm_util_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ccFieldsDump As New Collection
Dim ccFields As New Collection
Dim SQL As String
Dim AddDesc As String
Sub LoadCurrentReport()
On Error Resume Next
If ARView.Tag <> "" Then
   Dim jLap, RepObj As Object
   jLap = Split(ARView.Tag, "|")
    PeriodeLap = tpriode
   Select Case LCase(CStr(jLap(0)))
          Case "lap_pegawai"
                Set RepObj = New lap_Daftar_Pegawai
                GoSub loadObject
          Case "lap_pelanggan"
                Set RepObj = New lap_Daftar_Pelanggan
                GoSub loadObject
          Case "lap_penjamin"
                Set RepObj = New lap_Daftar_Penjamin
                GoSub loadObject
          Case "lap_divisi"
                Set RepObj = New lap_Daftar_Divisi
                GoSub loadObject
          Case "lap_barang"
                Set RepObj = New lap_Daftar_Barang
                GoSub loadObject
          Case "lap_angsuran_konsumen"
                Set RepObj = New lap_Angsuran_Konsumen
                GoSub loadObject
          Case "lap_angsuran_jt"
                Set RepObj = New lap_Angsuran_JT
                GoSub loadObject
          Case "lap_angsuran_tagihan"
                Set RepObj = New lap_Angsuran_Tagihan
                GoSub loadObject
          Case "lap_penerimaan_angsuran"
                Set RepObj = New lap_Angsuran_terimaA
                GoSub loadObject
          Case "lap_penjualan_pelanggan"
                Set RepObj = New lap_Penjualan_pelanggan
                GoSub loadObject
          Case "lap_penjualan_inspektur"
                Set RepObj = New lap_Penjualan_Inspek
                GoSub loadObject
          Case "lap_penjualan_salesmen"
                Set RepObj = New lap_Penjualan_Salesmen
                GoSub loadObject
          Case "lap_pendapatan"
                Set RepObj = New lap_RekapPendapatan
                GoSub loadObject
          Case "lap_daftar_tanggal"
                Set RepObj = New Lap_Daftar_tanggal
                GoSub loadObject
                
                                
          Case Else
                Unload Me
   End Select
   
End If
Exit Sub
loadObject:
On Error Resume Next
      Set ARView.ReportSource = Nothing
      ARView.Zoom = 100
      usedRep = DecodeSQL
      Set ARView.ReportSource = RepObj

      Call LockUnlock(StripPath(App.Path) & "_dba\_defbasis.xdb", False)
      RepObj.Sections("Detail").Controls("DataControl1").ConnectionString = srvLogon.ConnectionString
      RepObj.Sections("Detail").Controls("DataControl1").Source = usedRep
      On Error Resume Next
      RepObj.Sections("PageHeader").Controls("lblPeriode") = tpriode

End Sub
Sub simpanGrid()
On Error Resume Next
With CD
     .CancelError = True
     On Error GoTo X
     .Filter = "Kriteria file (*.kfg)|*.kfg"
     .ShowSave
     Grid.SaveGrid .Filename, flexFileAll
     Exit Sub
X:
     'ShowDlgMsg Me, "Data tidak dapat disimpan", vbOK, , True, False
End With
End Sub
Sub LoadGrid()
On Error Resume Next
With CD
     .CancelError = True
     On Error GoTo X
     .Filter = "Kriteria file (*.kfg)|*.kfg"
     .ShowOpen
     Grid.LoadGrid .Filename, flexFileAll
     Exit Sub
X:
     'ShowDlgMsg Me, "Data tidak dapat disimpan", vbOK, , True, False
End With
End Sub

Sub InitGrid()
On Error Resume Next
Grid.Clear
Grid.Rows = 23
Grid.TextMatrix(0, 0) = "Field"
Grid.TextMatrix(0, 1) = "Opt 1"
Grid.TextMatrix(0, 2) = "Value"
Grid.TextMatrix(0, 3) = "Opt 2"
End Sub

Private Sub ARView_LoadCompleted()
On Error Resume Next
piclap.Visible = False
Call LockUnlock(StripPath(App.Path) & "_dba\_defbasis.xdb", True)
End Sub

Private Sub ARView_ToolbarClick(ByVal Tool As DDActiveReportsViewer2Ctl.IDDTool)
Select Case Tool.ID
       Case 23
            Unload Me
End Select
End Sub


Sub btnExec_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            Grid.AddItem ""
       Case 1
            Grid.Rows = 1
            Grid.Rows = 5
       Case 2
            LoadGrid
       Case 3
            simpanGrid
       Case 4
            If Grid.Rows > 2 Then
               Grid.RemoveItem Grid.Row
            Else
               Grid.TextMatrix(Grid.Row, 0) = ""
               Grid.TextMatrix(Grid.Row, 1) = ""
               Grid.TextMatrix(Grid.Row, 2) = ""
               Grid.TextMatrix(Grid.Row, 3) = ""
            End If
       Case 5
            ARView.PrintReport True
       Case 6
            piclap.Visible = True
            LoadCurrentReport
       Case 7
               frm_util_report_pop.SetOBJ Me
            If Me.WindowState <> vbMinimized Then
               PopupMenu frm_util_report_pop.mnuExport, 1, (btnExec(index).Left + Frame1.Left) + 1150, btnExec(index).Top + Frame1.Top + btnExec(index).Height
            Else
                If Me.Left < 0 Then
                   PopupMenu frm_util_report_pop.mnuExport, 1, (btnExec(index).Left + Frame1.Left), btnExec(index).Top + Frame1.Top + btnExec(index).Height
                Else
                   PopupMenu frm_util_report_pop.mnuExport, 1, (btnExec(index).Left + Frame1.Left) + 1150, btnExec(index).Top + Frame1.Top + btnExec(index).Height
                End If
            End If
       Case 8
            Unload Me
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
tpriode.Text = "Periode: " & Format(Date, "dd Mmm yyyy")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Set Me.Picture = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next
'ARView.Width = Me.ScaleWidth
'ARView.Height = Me.ScaleHeight

End Sub

Sub ShowField(StrSql As String)
On Error Resume Next
SQL = StrSql
Set ccFields = Nothing
Set ccFieldsDump = Nothing
Dim Pos1 As Long, Pos2 As Long, strRes1 As New Collection, strRes2 As New Collection
Dim strSelect As String
Pos1 = InStr(1, SQL, "SELECT", vbTextCompare) 'cari kata select pada string sebagai acuan
If Pos1 Then
   Pos2 = InStr(Pos1 + 6, SQL, " FROM ", vbTextCompare) 'dan diakhiri dengan kata from
   If Pos2 Then
      strSelect = Mid(SQL, Pos1 + 6, Pos2 - 7)
      Dim nFields, X As Integer
      nFields = Split(strSelect, ",") 'pisahkan dengan menggunakan seperator koma (,)
      For X = 0 To UBound(nFields)
         If InStr(1, Trim(nFields(X)), ".", vbTextCompare) Then 'Pisahkan antara nama table dan nama field
            Dim myff, mygg
            myff = Split(Trim(nFields(X)), ".")
            
            If InStr(1, myff(0), " ", vbTextCompare) Then 'Seleksi untuk Nama Table
               If Left(myff(0), 1) = "[" And Right(myff(0), 1) = "]" Then 'Cari jika ada spasi tanpa kurung buka siku
                  strRes1.Add myff(0) & ""
               Else
                  strRes1.Add "[" & myff(0) & "]"
               End If
            Else
               strRes1.Add "[" & myff(0) & "]"
            End If
            
            If InStr(1, myff(1), " AS ", vbTextCompare) Then 'Seleksi untuk Nama Field
                mygg = Split(myff(1), " AS ", , vbTextCompare)
                If UBound(mygg) > 0 Then
                   If InStr(1, mygg(1), " ", vbTextCompare) Then
                      If Left(mygg(1), 1) = "[" And Right(mygg(1), 1) = "]" Then   'Cari jika ada spasi tanpa kurung buka siku
                         strRes2.Add mygg(1) & "»"
                         ccFieldsDump.Add myff(0) & "." & myff(1), mygg(1) & "»"
                      Else
                         strRes2.Add "[" & mygg(1) & "]»"
                         ccFieldsDump.Add myff(0) & "." & myff(1), "[" & mygg(1) & "]»"
                      End If
                   Else
                      strRes2.Add "[" & Trim(mygg(1)) & "]»"
                      ccFieldsDump.Add myff(0) & "." & myff(1), "[" & Trim(mygg(1)) & "]»"
                   End If
                End If
            Else
                strRes2.Add myff(1)
            End If
         Else
         
         End If
      Next X
      Dim i As Integer, C As String, d As String
      For i = 1 To strRes1.Count
      
           C = C & strRes2(i) & "|"
           
           d = Replace(strRes2(i), "]", "")
           
           d = Replace(d, "[", "")
           d = Replace(d, "]", "")
           d = Replace(d, "_", " ")
           d = Replace(d, "»", "")
           d = Replace(d, "$", "") 'Currency
           d = Replace(d, "#", "") 'Numeric
           d = Replace(d, "&", "") 'String
           d = Replace(d, "@", "") 'Date/Time
           cboGroup.AddItem StrConv(d, vbUpperCase)
           'lstCheck.AddItem UCase(d)
           'lstCheck.Selected(lstCheck.ListCount - 1) = True
           If InStr(1, strRes2(i), "»") > 0 Then
              ccFields.Add strRes2(i), UCase(d)
           Else
              ccFields.Add strRes1(i) & "." & strRes2(i), UCase(d)
           End If
      Next i
      cboGroup.AddItem " "
      C = Replace(C, "]", "")
      C = Replace(C, "[", "")
      C = Replace(C, "_", " ")
      C = Replace(C, "»", "")
      C = Replace(C, "$", "") 'Currency
      C = Replace(C, "#", "") 'Numeric
      C = Replace(C, "&", "") 'String
      C = Replace(C, "@", "") 'Date/Time
      Grid.ColComboList(0) = StrConv(C, vbUpperCase)
      
   End If
   Set strRes1 = Nothing
   Set strRes2 = Nothing
   C = ""
   d = ""
   Dim jLap
   jLap = Split(ARView.Tag, "|")
    Dim inSet
    'TGL^=^1/2/2002ÿ
    jLap(2) = Replace(jLap(2), "<!date>", "@" & (Date) & "@", , , vbTextCompare)
    inSet = Split(jLap(2), "ÿ")
    If UBound(inSet) >= 0 Then
       Dim inSide
       For i = 0 To UBound(inSet)
          inSide = Split(inSet(i), "þ")
          Grid.TextMatrix(i + 1, 0) = inSide(0)
          Grid.TextMatrix(i + 1, 1) = inSide(1)
          Grid.TextMatrix(i + 1, 2) = inSide(2)
          Grid.TextMatrix(i + 1, 3) = inSide(3)
       Next i
    End If
End If
End Sub

Function DecodeSQL() As String
On Error Resume Next
          Dim i As Integer, cSQL As String, jjField As String, jjvalue As String
          For i = 1 To Grid.Rows - 1
              If Trim(Grid.TextMatrix(i, 0)) <> "" And Trim(Grid.TextMatrix(i, 1)) <> "" Then
                    jjField = ccFields(Grid.TextMatrix(i, 0))
                    If InStr(1, jjField, "$") Then
                        jjField = Replace(jjField, "$", "")
                        jjvalue = AllowChar(FDec(Grid.TextMatrix(i, 2)))
                    ElseIf InStr(1, jjField, "#") Then
                        jjField = Replace(jjField, "#", "")
                        jjvalue = AllowChar(Grid.TextMatrix(i, 2))
                    ElseIf InStr(1, jjField, "@") Then
                        jjField = Replace(jjField, "@", "")
                        jjvalue = "'" & AllowChar(strToDate(Grid.TextMatrix(i, 2))) & "'"
                    ElseIf InStr(1, jjField, "&") Then
                        jjField = Replace(jjField, "&", "")
                        jjvalue = "'" & AllowChar(Grid.TextMatrix(i, 2)) & "'"
                    Else
                        If InStr(1, Grid.TextMatrix(i, 2), "@", vbTextCompare) Then
                           jjvalue = Replace(Grid.TextMatrix(i, 2), "@", "")
                           jjvalue = "#" & ReverseDate(jjvalue) & "#"
                        ElseIf InStr(1, Grid.TextMatrix(i, 2), "|", vbTextCompare) Then
                           jjvalue = Replace(Grid.TextMatrix(i, 2), "|", "'")
                        Else
                           If UCase(Grid.TextMatrix(i, 1)) = "BETWEEN" Then
                              jjvalue = Date2Between(Grid.TextMatrix(i, 2))
                           ElseIf Trim(UCase(Grid.TextMatrix(i, 1))) = "IS NULL" Or Trim(UCase(Grid.TextMatrix(i, 1))) = "NULL" Then
                              jjvalue = Grid.TextMatrix(i, 2)
                           Else
                              jjvalue = "'" & AllowChar(Grid.TextMatrix(i, 2)) & "'"
                           End If
                        End If
                    End If
                    Grid.AddItem ""
                    
                    If InStr(1, jjField, "»", vbTextCompare) Then
                       Dim kAS
                       kAS = Split(CStr(ccFieldsDump(jjField)), " AS ", , vbTextCompare)
                       jjField = kAS(0)
                    End If
                    
                    If Trim(Grid.TextMatrix(i + 1, 0)) <> "" And Trim(Grid.TextMatrix(i + 1, 1)) <> "" Then
                       cSQL = cSQL & "(" & jjField & " " & _
                              Grid.TextMatrix(i, 1) & " " & _
                              jjvalue & ") " & _
                              Grid.TextMatrix(i, 3) & " "
                    Else
                       cSQL = cSQL & "(" & jjField & " " & _
                              Grid.TextMatrix(i, 1) & " " & _
                              jjvalue & ") "
                    
                    End If
                    Grid.RemoveItem Grid.Rows - 1
              Else
              End If
          Next i
             
           Dim nSource As String
           'nSource = Replace(SQL, "*", "")
           nSource = Replace(SQL, "$", "") 'Currency
           nSource = Replace(nSource, "#", "") 'Numeric
           nSource = Replace(nSource, "&", "") 'String
           nSource = Replace(nSource, "@", "") 'Date/Time
          
          Dim inmyStr As String
             inmyStr = cSQL
          If Trim(inmyStr) <> "" Then
             If InStr(1, nSource, "<!having>") > 0 Then
                 nSource = Replace(nSource, "<!having>", " HAVING " & inmyStr & " " & AddDesc)
             ElseIf InStr(1, nSource, "<!where>") > 0 Then
                 nSource = Replace(nSource, "<!where>", " WHERE " & inmyStr & " " & AddDesc)
             End If
             DecodeSQL = nSource
          Else
             nSource = Replace(nSource, "<!having>", "") & " " & AddDesc
             nSource = Replace(nSource, "<!where>", "") & " " & AddDesc
             DecodeSQL = nSource
          End If
End Function


Function dlgExportRpt(nType As String) As String
On Error Resume Next
With CD
     .CancelError = True
     On Error GoTo X
        Select Case LCase(nType)
               Case "pdf"
                    Dim pdf As New ActiveReportsPDFExport.ARExportPDF
                    CD.Filter = "Portable Document Format (*.PDF)| *.PDF"
                    CD.DialogTitle = "Export ke PDF"
                    .ShowSave
                    pdf.Filename = CD.Filename
                    Exporits pdf
               Case "html"
                    Dim html As New ActiveReportsHTMLExport.HTMLexport
                    Dim sFolder As String
                    Dim iPos As Long
                    Dim iLen As Long, sFile As String
                
                    CD.Filter = "Hyper Text Mark-Up Language (*.HTML)| *.HTML"
                    CD.DialogTitle = "Export ke HTML"
                    .ShowSave
                    sFile = .Filename
                    iLen = InStr(1, sFile, vbNullChar)
                    For iPos = (iLen - 1) To 0 Step -1
                   
                    If Mid(sFile, iPos, 1) = "\" Then
                        sFolder = Left(sFile, iPos)
                        Exit For
                    End If
                    Next iPos
                   
                    html.HTMLOutputPath = sFolder
                    html.FileNamePrefix = sFile
                    Exporits html
               Case "rtf"
                    Dim rtf As New ActiveReportsRTFExport.ARExportRTF
                    CD.Filter = "Rich Text Format (*.doc)| *.doc"
                    CD.DialogTitle = "Export ke dokumen"
                    .ShowSave
                    rtf.Filename = CD.Filename
                    Exporits rtf
               Case "xls"
                    'Dim xls As New ActiveReportsHTMLExport.HTMLexport
                    Dim xls As New ActiveReportsExcelExport.ARExportExcel
                    CD.Filter = "Excel Format (*.xls)| *.xls"
                    CD.DialogTitle = "Export ke Excel"
                    .ShowSave
                    xls.Filename = CD.Filename
                    Exporits xls
                    
        End Select
                        
     Exit Function
X:
     
End With
End Function

Sub Exporits(obj As Object)
On Error Resume Next
    If ARView.Pages.Count > 0 Then
        obj.Export ARView.Pages
    ElseIf Not ARView.ReportSource Is Nothing Then
        If ARView.ReportSource.Pages.Count > 0 Then
            obj.Export ARView.ReportSource.Pages
        End If
    End If
    Set obj = Nothing
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
Select Case Col
       Case 0
            If Grid.TextMatrix(Row - 1, 3) = "" Then
               Cancel = True
            End If
       Case 1
            If Grid.TextMatrix(Row, 0) = "" Then
               Cancel = True
            End If
       Case 2
            If Grid.TextMatrix(Row, 1) = "" Then
               Cancel = True
            End If
       Case 3
            If Grid.TextMatrix(Row, 1) = "" Then
               Cancel = True
            End If
End Select
End Sub
