VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Management"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_util_user.frx":0000
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   90
      TabIndex        =   18
      Top             =   90
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "User Properties"
      TabPicture(0)   =   "_frm_util_user.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnExec(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnExec(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnExec(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnExec(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "btnExec(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btnExec(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btnExec(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Grid"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Check5(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check5(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Check5(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Check5(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Check5(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   195
         Index           =   6
         Left            =   5040
         TabIndex        =   23
         Top             =   2385
         Width           =   210
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   195
         Index           =   5
         Left            =   4770
         TabIndex        =   22
         Top             =   2385
         Width           =   210
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   195
         Index           =   4
         Left            =   4500
         TabIndex        =   21
         Top             =   2385
         Width           =   210
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   195
         Index           =   3
         Left            =   4215
         TabIndex        =   20
         Top             =   2385
         Width           =   210
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   19
         Top             =   2385
         Width           =   210
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Top             =   495
         Width           =   2835
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1290
         PasswordChar    =   "l"
         TabIndex        =   3
         Top             =   870
         Width           =   2835
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1290
         PasswordChar    =   "l"
         TabIndex        =   5
         Top             =   1245
         Width           =   2835
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Administrator"
         Height          =   225
         Left            =   1275
         TabIndex        =   6
         Top             =   1620
         Width           =   1320
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Account Aktif"
         Height          =   225
         Left            =   2745
         TabIndex        =   9
         Top             =   1605
         Width           =   1515
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Restore Database"
         Height          =   225
         Left            =   1275
         TabIndex        =   8
         Top             =   2145
         Width           =   1860
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Backup Database"
         Height          =   225
         Left            =   1275
         TabIndex        =   7
         Top             =   1875
         Width           =   1680
      End
      Begin VB.PictureBox Grid 
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   2610
         Left            =   120
         ScaleHeight     =   2550
         ScaleWidth      =   5355
         TabIndex        =   10
         Top             =   2640
         Width           =   5415
      End
      Begin SysInfo_Nardhika.vbButton btnExec 
         Height          =   375
         Index           =   0
         Left            =   4365
         TabIndex        =   14
         Top             =   465
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Baru"
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
         MICON           =   "_frm_util_user.frx":03A6
         PICN            =   "_frm_util_user.frx":06C0
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
         Left            =   4365
         TabIndex        =   16
         Top             =   1335
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   5
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
         MICON           =   "_frm_util_user.frx":0A5A
         PICN            =   "_frm_util_user.frx":0D74
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
         Left            =   4365
         TabIndex        =   15
         Top             =   900
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   5
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
         MICON           =   "_frm_util_user.frx":130E
         PICN            =   "_frm_util_user.frx":1628
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
         Left            =   105
         TabIndex        =   11
         Top             =   5385
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   ""
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
         MICON           =   "_frm_util_user.frx":1BC2
         PICN            =   "_frm_util_user.frx":1EDC
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
         Left            =   810
         TabIndex        =   12
         Top             =   5385
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   ""
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
         MICON           =   "_frm_util_user.frx":2276
         PICN            =   "_frm_util_user.frx":2590
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
         Left            =   1530
         TabIndex        =   13
         Top             =   5385
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Refresh"
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
         MICON           =   "_frm_util_user.frx":292A
         PICN            =   "_frm_util_user.frx":2C44
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
         Index           =   6
         Left            =   4365
         TabIndex        =   17
         Top             =   5385
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         BTYPE           =   5
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
         MICON           =   "_frm_util_user.frx":2FDE
         PICN            =   "_frm_util_user.frx":32F8
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
         Caption         =   "Nama User"
         Height          =   195
         Left            =   135
         TabIndex        =   0
         Top             =   525
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   915
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Password"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   1275
         Width           =   945
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hUser As New ADODB.Recordset
Sub UpdateGrid()
On Error Resume Next
Dim lErr As String
lErr = FindRecord("SELECT login From users WHERE (login = '" & AllowChar(Text1.Text) & "')", True, srvUSER)
If lErr = "0" Then
kembali:
   lErr = SaveRecord("users", Array("login=" & Text1.Text, _
                                    "[password]=" & Text2.Text, _
                                    "admin=" & Check1.Value, _
                                    "backup=" & Check4.Value, _
                                    "restore=" & Check3.Value, _
                                    "aktif=" & Check2.Value), True, srvUSER)
    Dim i As Integer
    For i = 1 To Grid.Rows - 1
        SaveRecord "manage", Array("ID=" & Grid.TextMatrix(i, 0), _
                                   "Login=" & Text1.Text, _
                                   "N=" & Grid.TextMatrix(i, 2), _
                                   "S=" & Grid.TextMatrix(i, 3), _
                                   "E=" & Grid.TextMatrix(i, 4), _
                                   "D=" & Grid.TextMatrix(i, 5), _
                                   "P=" & Grid.TextMatrix(i, 6)), True, srvUSER
    Next i
Else
   srvUSER.Execute "DELETE FROM users WHERE (login = '" & AllowChar(Text1) & "')"
   GoSub kembali:
End If
ClearControl Me
Grid.Rows = 1
GetFormData
hUser.Requery
End Sub
Sub GetFormData()
With Grid
     Dim rc As New ADODB.Recordset
     Dim lErr As String
     rc.Open "SELECT ID, Description From IDProg ORDER BY ID", srvUSER, LockType1, LockType2
     If Not rc.EOF Then
        Grid.Rows = 1
        While Not rc.EOF
             Grid.AddItem NotNull(rc("ID"))
             Grid.TextMatrix(Grid.Rows - 1, 1) = NotNull(rc("Description"))
             rc.MoveNext
        Wend
     End If
End With
End Sub
Private Sub btnExec_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            ClearControl Me
            Grid.Rows = 1
            GetFormData
            Text1.SetFocus
            Check1.Value = 0
            Check2.Value = 0
            Check3.Value = 0
            Check4.Value = 0
       Case 1
                If Text2 = Text3 Then
                   UpdateGrid
                Else
                   MsgBox "Password tidak cocok dengan re-Password, coba lagi donk", 16
                   Text2.SetFocus
                End If
       Case 6
            Unload Me
       Case 2
            If LCase(Text1) <> "admin" Then
               srvUSER.Execute "DELETE FROM users WHERE (login = '" & AllowChar(Text1) & "')"
               ClearControl Me
               Grid.Rows = 1
               GetFormData
            Else
               ShowDlgMsg Me, "Administrator tidak dapat dihapus", vbOK, , True, False
            End If
       Case 3
            On Error Resume Next
            Check1.Value = 0
            Check2.Value = 0
            Check3.Value = 0
            Check4.Value = 0
            
            If Not hUser.BOF Then
               hUser.MovePrevious
               If Not hUser.BOF Then ShowData hUser.Fields("login").Value
            End If
       Case 4
            On Error Resume Next
            Check1.Value = 0
            Check2.Value = 0
            Check3.Value = 0
            Check4.Value = 0
            If Not hUser.EOF Then
               hUser.MoveNext
               If Not hUser.EOF Then ShowData hUser.Fields("login").Value
               
            End If
       Case 5
            On Error Resume Next
            hUser.Requery
End Select
End Sub

Sub ShowData(nKey As String)
On Error Resume Next
Dim rc As New ADODB.Recordset
rc.Open "Select * from users where login='" & nKey & "'", srvUSER, LockType1, LockType2
If Not rc.EOF Then
   Text1 = NotNull(rc("login"))
   Text2 = NotNull(rc("password"))
   Text3 = NotNull(rc("password"))
   Check1.Value = Val(NotNull(rc("admin")))
   Check2.Value = Val(NotNull(rc("aktif")))
   Check3.Value = Val(NotNull(rc("restore")))
   Check4.Value = Val(NotNull(rc("backup")))
   Dim i As Integer
   For i = 1 To Grid.Rows - 1
      Dim RC2 As New ADODB.Recordset
      RC2.Open "SELECT N, S, E, D, P From manage WHERE (Login = '" & Text1 & "') AND (ID = '" & Grid.TextMatrix(i, 0) & "')", srvUSER, LockType1, LockType2
      If Not RC2.EOF Then
         Grid.TextMatrix(i, 2) = NotNull(RC2("N"))
         Grid.TextMatrix(i, 3) = NotNull(RC2("S"))
         Grid.TextMatrix(i, 4) = NotNull(RC2("E"))
         Grid.TextMatrix(i, 5) = NotNull(RC2("D"))
         Grid.TextMatrix(i, 6) = NotNull(RC2("P"))
      End If
      RC2.Close
   Next i
End If
rc.Close
End Sub

Private Sub Check5_Click(index As Integer)
On Error Resume Next
Dim i As Integer
For i = 1 To Grid.Rows - 1
    Grid.TextMatrix(i, index) = IIf(Check5(index).Value = 1, -1, 0)
Next i
End Sub

Private Sub Form_Load()
On Error Resume Next
    hUser.Open "SELECT login From users ORDER BY login", srvUSER, LockType1, LockType2
    GetFormData
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
hUser.Close
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
Select Case Col
       Case 0, 1
            Cancel = True
End Select
End Sub



