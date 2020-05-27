VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Hari Libur"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_mst_tgl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4.551
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   11.192
   Begin VB.Frame Frame1 
      Caption         =   "DATA HARI LIBUR"
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   6090
      Begin SysInfo_Nardhika.vbTextBox txtFields 
         Height          =   315
         Index           =   0
         Left            =   1815
         TabIndex        =   2
         Top             =   330
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         Icon            =   "_frm_mst_tgl.frx":038A
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   -1  'True
         BorderColor     =   33023
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
         Index           =   1
         Left            =   1815
         TabIndex        =   4
         Top             =   720
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
         Index           =   1
         Left            =   225
         TabIndex        =   3
         Top             =   795
         Width           =   945
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Libur"
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
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   1500
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
            Picture         =   "_frm_mst_tgl.frx":07D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":0B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":0F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":12A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":1640
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":19DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":1D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":210E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":24A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":2842
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":2BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":2F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":3310
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":36AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_mst_tgl.frx":3C44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1710
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   3016
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
      TabIndex        =   6
      Top             =   2250
      Width           =   6345
      _ExtentX        =   11192
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
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   34.387
      Y1              =   1.058
      Y2              =   1.058
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   -0.026
      X2              =   34.361
      Y1              =   1.085
      Y2              =   1.085
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   -0.397
      X2              =   33.99
      Y1              =   3.863
      Y2              =   3.863
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   -0.397
      X2              =   33.99
      Y1              =   3.889
      Y2              =   3.889
   End
End
Attribute VB_Name = "Form11"
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.index
       Case 1
           If CekUser("23", "N") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk mencetak laporan ini", vbOK, , True, False
           Else
                ClearControl Me
                Me.Caption = Replace(Me.Caption, "*", "")
                Me.Tag = "*"
                txtFields(0).Tag = ""
                Me.Caption = Me.Caption & Me.Tag
                txtFields(0).SetFocus
           End If
       Case 2
           If CekUser("23", "S") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else
               SimpanData (txtFields(0).Text)
           End If
       Case 4
           If CekUser("23", "D") = False Then
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
           If CekUser("23", "P") = False Then
              ShowDlgMsg Me, "Anda tidak diperkenankan untuk melakukan pengeditan data", vbOK, , True, False
           Else

            Dim StrSql As String, Form4 As New frm_util_report
            Load Form4
            StrSql = "SELECT mst_hari_libur.[Tanggal], mst_hari_libur.Keterangan From mst_hari_libur <!where> ORDER BY mst_hari_libur.[Tanggal];"

            Form4.ARView.Tag = "lap_daftar_tanggal|" & StrSql
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
On Error GoTo salah
Dim h As String, hErr As String, i As Integer
h = FindRecord("SELECT mst_hari_libur.[Tanggal] From mst_hari_libur WHERE mst_hari_libur.[Tanggal]=#" & ReverseDate(strToDate(nKey)) & "#;")

If h = "0" Then
                                                  
   h = SaveRecord("mst_hari_libur", Array("@Tanggal=" & txtFields(0).Text, _
                                          "Keterangan=" & txtFields(1).Text))
                                                  
  If h = "" Then
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
         h = UpdateRecord("mst_hari_libur", Array("@Tanggal=" & txtFields(0).Text, _
                          "Keterangan=" & txtFields(1).Text), " WHERE [Tanggal]=#" & strToDate(txtFields(0).Text) & "# ")
                                                          
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
Exit Sub
salah:
CreateLog Error
End Sub

Sub HapusData(hKey As String)
On Error GoTo salah
Dim hErr, h As String
hErr = FindRecord("SELECT mst_hari_libur.[Tanggal] From mst_hari_libur WHERE (((mst_hari_libur.[Tanggal])=#" & strToDate(hKey) & "#));")
If hErr = "1" Then
   If ShowDlgMsg(Me, "Hapus Data Hari Libur?", vbYesNo, Error, False, True, , , , , Me.name & "_deleted") = False Then
      GoSub Delete_Label
   Else
      If SelectMsg = vbYes Then
Delete_Label:
         hErr = ExecQuery("DELETE mst_hari_libur.[Tanggal] From mst_hari_libur WHERE (((mst_hari_libur.[Tanggal])=#" & strToDate(hKey) & "#));")
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
Exit Sub
salah:
CreateLog Error
End Sub

Sub ShowDivisi(nKey As String)
On Error Resume Next
Dim hKey, rc As New ADODB.Recordset, hErr As String
hKey = Split(nKey, "|")
hErr = SelectQuery(rc, "SELECT mst_hari_libur.[Tanggal],  mst_hari_libur.[Keterangan] from mst_hari_libur where [Tanggal]=#" & ReverseDate(CStr(hKey(0))) & "# ORDER BY mst_hari_libur.[Tanggal]; ")
If hErr = "" Then
    If Not rc.EOF Then
        txtFields(0).Text = NotNull(rc("Tanggal"))
        txtFields(1).Text = NotNull(rc("Keterangan"))
    Else
kembali:
       txtFields(0).Text = ""
        txtFields(1).Text = ""
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
            CurRec.MoveFirst
            ShowDivisi NotNull(CurRec("Tanggal")) & "|"
       Case 2
            If Not CurRec.BOF Then
               CurRec.MovePrevious
               If CurRec.BOF Then
                  CurRec.MoveNext
               End If
               ShowDivisi NotNull(CurRec("Tanggal")) & "|"
            End If
       Case 3
            If Not CurRec.EOF Then
               CurRec.MoveNext
               If CurRec.EOF Then
                  CurRec.MovePrevious
               End If
               ShowDivisi NotNull(CurRec("Tanggal")) & "|"
            End If
       Case 4
            'If Not CurRec.EOF Then
               CurRec.MoveLast
               ShowDivisi NotNull(CurRec("Tanggal")) & "|"
            'End If
       Case 7
            If CurRec.State = 1 Then CurRec.Close
            GoSub subLoadDB
            CurRec.MoveFirst
            ShowDivisi NotNull(CurRec("Tanggal")) & "|"
End Select
MainMenu.StatusBar1.Panels(2).Text = "Record " & CurRec.AbsolutePosition & " dari " & CurRec.RecordCount
Exit Sub
subLoadDB:
SelectQuery CurRec, "SELECT [Tanggal] From mst_hari_libur ORDER BY Tanggal"
Return

End Sub

Private Sub txtFields_DownButtonClick(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            ShowFindForm "SELECT mst_hari_libur.[Tanggal], mst_hari_libur.[Keterangan] " & _
                         " FROM mst_hari_libur <!where> ORDER BY mst_hari_libur.[Tanggal]; ", "#" & txtFields(index).Hwnd1, Me, "ShowDivisi"
       End Select
End Sub

Private Sub txtFields_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case 0
            Select Case index
                   Case 0
                      ShowDivisi txtFields(index).Text & "|"
            End Select
       Case Else
            If Me.Tag = "" Then
               Me.Tag = "*"
               Me.Caption = Me.Caption & Me.Tag
            End If
End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
PostFindForm "#" & txtFields(0).Hwnd1
CurRec.Close
Set CurRec = Nothing
End Sub

