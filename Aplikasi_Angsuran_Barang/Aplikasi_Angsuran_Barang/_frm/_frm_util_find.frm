VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_util_find 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_util_find.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7020
   Begin SysInfo_Nardhika.vbButton cmdExec 
      Height          =   345
      Index           =   11
      Left            =   5670
      TabIndex        =   16
      Top             =   2685
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
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
      MICON           =   "_frm_util_find.frx":038A
      PICN            =   "_frm_util_find.frx":06A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfo_Nardhika.vbButton cmdExec 
      Height          =   345
      Index           =   0
      Left            =   4485
      TabIndex        =   15
      Top             =   2685
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
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
      MICON           =   "_frm_util_find.frx":0A3E
      PICN            =   "_frm_util_find.frx":0D58
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab objTAB 
      Height          =   2025
      Left            =   120
      TabIndex        =   7
      Top             =   585
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   3572
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "Standard Pencarian"
      TabPicture(0)   =   "_frm_util_find.frx":10F2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "O1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "O2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCari"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Pencarian Custom"
      TabPicture(1)   =   "_frm_util_find.frx":110E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdExec(10)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdExec(6)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdExec(8)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Grid"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Tampilan Field"
      TabPicture(2)   =   "_frm_util_find.frx":112A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lstCheck"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00EFF7F7&
         Height          =   315
         Left            =   -74340
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   405
         Width           =   6015
      End
      Begin VB.TextBox txtCari 
         BackColor       =   &H00EFF7F7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -74880
         TabIndex        =   1
         Top             =   1035
         Width           =   6540
      End
      Begin VB.OptionButton O2 
         Caption         =   "0&2. Mengandung Kata"
         Height          =   240
         Left            =   -73245
         TabIndex        =   3
         Top             =   1665
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton O1 
         Caption         =   "0&1. Kata Awalan"
         Height          =   240
         Left            =   -74895
         TabIndex        =   2
         Top             =   1665
         Width           =   1920
      End
      Begin VB.ListBox lstCheck 
         BackColor       =   &H00FFFFFF&
         Columns         =   2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1530
         IntegralHeight  =   0   'False
         ItemData        =   "_frm_util_find.frx":1146
         Left            =   75
         List            =   "_frm_util_find.frx":1148
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   405
         Width           =   6600
      End
      Begin VB.PictureBox Grid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1290
         Left            =   -74895
         ScaleHeight     =   1260
         ScaleWidth      =   5385
         TabIndex        =   9
         Top             =   615
         Width           =   5415
      End
      Begin SysInfo_Nardhika.vbButton cmdExec 
         Height          =   345
         Index           =   8
         Left            =   -69390
         TabIndex        =   10
         Top             =   615
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   609
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "_frm_util_find.frx":114A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton cmdExec 
         Height          =   345
         Index           =   6
         Left            =   -69390
         TabIndex        =   11
         Top             =   1020
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   609
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "_frm_util_find.frx":1464
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SysInfo_Nardhika.vbButton cmdExec 
         Height          =   345
         Index           =   10
         Left            =   -69390
         TabIndex        =   12
         Top             =   1425
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   609
         BTYPE           =   8
         TX              =   "&Reset"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "_frm_util_find.frx":177E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Field:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   5
         Top             =   465
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Kata yang akan dicari:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -74850
         TabIndex        =   0
         Top             =   780
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria Pencarian"
         Height          =   195
         Left            =   -74850
         TabIndex        =   13
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.PictureBox DrgData 
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   120
      ScaleHeight     =   2010
      ScaleWidth      =   6720
      TabIndex        =   4
      Top             =   3105
      Width           =   6780
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7425
      Top             =   3810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":1A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":1E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":21CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":2566
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":2900
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":2C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":3034
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":33CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":3768
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":3D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":429C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":4636
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_util_find.frx":49D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   -165
      X2              =   19330
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   -165
      X2              =   19330
      Y1              =   375
      Y2              =   375
   End
End
Attribute VB_Name = "frm_util_find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ccFields As New Collection
Dim ccFieldsDump As New Collection
'Dim SQL As String
Dim CallObj As Object, procName As String
Dim RCxx As New ADODB.Recordset
Dim SQL As String

Sub DecodeSQL()
On Error Resume Next
          Dim i As Integer, cSQL As String, jjField As String, jjvalue As String
          Dim oAsField As String
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
                        If InStr(1, Grid.TextMatrix(i, 2), "@", vbTextCompare) Then                        'jjvalue = "'" & AllowChar(Grid.TextMatrix(I, 2)) & "'"
                        jjvalue = Replace(Grid.TextMatrix(i, 2), "@", "")
                        Else
                        jjvalue = "'" & AllowChar(Grid.TextMatrix(i, 2)) & "'"
                        End If
                    End If
                    Grid.AddItem ""
                    
                    If InStr(1, jjField, "�", vbTextCompare) Then
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
           nSource = Replace(arrQueryForm.GetItem(Me.Tag), "�", "")
           nSource = Replace(nSource, "$", "") 'Currency
           nSource = Replace(nSource, "#", "") 'Numeric
           nSource = Replace(nSource, "&", "") 'String
           nSource = Replace(nSource, "@", "") 'Date/Time
          
          Dim inmyStr As String
          If objTAB.Tab = 0 Then
             oAsField = ccFields.Item(Combo1.Text)
             If InStr(1, oAsField, "�", vbTextCompare) Then
                kAS = Split(CStr(ccFieldsDump(oAsField)), " AS ", , vbTextCompare)
                inmyStr = kAS(0)
             Else
               inmyStr = oAsField
             End If
             If O1.Value Then
                inmyStr = "LEFT(" & inmyStr & "," & Len(AllowChar(txtCari)) & ")='" & AllowChar(txtCari) & "'"
             Else
                inmyStr = inmyStr & " LIKE '%" & AllowChar(txtCari) & "%'"
             End If
          ElseIf objTAB.Tab = 1 Then
             inmyStr = cSQL
          End If
          Dim customSQL As String, Pos1 As Long
          For i = 0 To lstCheck.ListCount - 1
              If lstCheck.Selected(i) = True Then
                 oAsField = ccFields.Item(lstCheck.List(i))
                 If InStr(1, oAsField, "�", vbTextCompare) Then
                    customSQL = customSQL & ccFieldsDump.Item(oAsField) & ", "
                 Else
                    customSQL = customSQL & ccFields.Item(lstCheck.List(i)) & ", "
                 End If
              End If
          Next i
          customSQL = Replace(customSQL, "$", "") 'Currency
          customSQL = Replace(customSQL, "#", "") 'Numeric
          customSQL = Replace(customSQL, "&", "") 'String
          customSQL = Replace(customSQL, "@", "") 'Date/Time

          Pos1 = InStr(Pos1 + 1, nSource, "from", vbTextCompare)
          If Pos1 Then
             nSource = "SELECT " & Mid(customSQL, 1, Len(customSQL) - 2) & _
             Mid(nSource, Pos1 - 1)
          End If
                    
          If Trim(inmyStr) <> "" Then
             If InStr(1, nSource, "<!having>") > 0 Then
                nSource = Replace(nSource, "<!having>", " HAVING " & DrgData.Tag & "  " & inmyStr, , , vbTextCompare)
             ElseIf InStr(1, nSource, "<!where>") > 0 Then
                 nSource = Replace(nSource, "<!where>", " WHERE " & DrgData.Tag & "  " & inmyStr, , , vbTextCompare)
             End If
             ShowRecord nSource
          Else
             nSource = Replace(nSource, "<!having>", "")
             nSource = Replace(nSource, "<!where>", "")
             ShowRecord nSource
          End If
          
End Sub

Sub ShowRecord(nSource As String)
On Error GoTo salah
   Dim lErr As String
   Set RCxx = New ADODB.Recordset
   lErr = SelectQuery(RCxx, nSource)
   If lErr = "" Then
      Set DrgData.DataSource = RCxx
      DrgData.DataRefresh
   End If
'   Dim colStr As String, I As Integer
'   For I = 0 To DrgData.Cols - 1
'       colStr = Replace(DrgData.TextMatrix(0, I), "_", " ")
'       colStr = StrConv(colStr, vbProperCase)
'       DrgData.TextMatrix(0, I) = colStr
'   Next I
Exit Sub
salah:
ShowDlgMsg Me, "Ada kesalahan sewaktu pengisian data yang akan dicari.", vbOK, Error, True, False
CreateLog Error
End Sub

Sub ShowField(obj As Object, Proc As String)
On Error Resume Next
SQL = arrQueryForm.GetItem(Me.Tag)
Dim Pos1 As Long, Pos2 As Long, strRes1 As New Collection, strRes2 As New Collection
Dim strSelect As String
Set CallObj = obj
procName = Proc
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
                         strRes2.Add mygg(1) & "�"
                         ccFieldsDump.Add myff(0) & "." & myff(1), mygg(1) & "�"
                      Else
                         strRes2.Add "[" & mygg(1) & "]�"
                         ccFieldsDump.Add myff(0) & "." & myff(1), "[" & mygg(1) & "]�"
                      End If
                   Else
                      strRes2.Add "[" & Trim(mygg(1)) & "]�"
                      ccFieldsDump.Add myff(0) & "." & myff(1), "[" & Trim(mygg(1)) & "]�"
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
           d = Replace(d, "_", " ")
           d = Replace(d, "�", "")
           d = Replace(d, "$", "") 'Currency
           d = Replace(d, "#", "") 'Numeric
           d = Replace(d, "&", "") 'String
           d = Replace(d, "@", "") 'Date/Time
           Combo1.AddItem UCase(d)
           lstCheck.AddItem UCase(d)
           lstCheck.Selected(lstCheck.ListCount - 1) = True
           If InStr(1, strRes2(i), "�", vbTextCompare) > 0 Then
              ccFields.Add strRes2(i), UCase(d)
           Else
              ccFields.Add strRes1(i) & "." & strRes2(i), UCase(d)
           End If
      Next i
      C = Replace(C, "]", "")
      C = Replace(C, "[", "")
      C = Replace(C, "_", " ")
      C = Replace(C, "�", "")
      C = Replace(C, "$", "") 'Currency
      C = Replace(C, "#", "") 'Numeric
      C = Replace(C, "&", "") 'String
      C = Replace(C, "@", "") 'Date/Time
      Grid.ColComboList(0) = StrConv(C, vbUpperCase)
      Combo1.ListIndex = 0
   End If
   Set strRes1 = Nothing
   Set strRes2 = Nothing
   C = ""
   d = ""
End If
End Sub

Private Sub btnMenu_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            Unload Me
       Case 1
            Me.WindowState = vbMinimized
            Me.Hide
End Select
End Sub

Private Sub cmdExec_Click(index As Integer)
On Error Resume Next
Select Case index
       Case 0
            Dim i As Integer
            For i = 0 To lstCheck.ListCount - 1
                If lstCheck.Selected(i) = True Then
                    DecodeSQL
                    Exit For
                End If
            Next i
            'DrgData.SetFocus
       Case 1
            On Error Resume Next
            If DrgData.Rows > 1 Then
            If DrgData.TextMatrix(1, 0) <> "" Then
            Dim m As String
            For i = 0 To DrgData.Cols - 1
               m = m & DrgData.TextMatrix(DrgData.Row, i) & "|"
            Next i
            DrgData.SetFocus
            Me.Hide
            CallByName CallObj, procName, VbMethod, m
            CallObj.SetFocus
            End If
            End If
       Case 2
            On Error Resume Next
            DrgData.Row = DrgData.Row + 1
            DrgData_MouseDown 2, 0, 0, 0
       Case 11
            Me.Hide
       Case 4
            On Error Resume Next
            DrgData.Row = DrgData.Row - 1
            DrgData_MouseDown 2, 0, 0, 0
       Case 5
            Set ccFields = Nothing
            ShowField CallObj, procName
       Case 8
            Grid.AddItem ""
       Case 6
            If Grid.Rows > 2 Then Grid.RemoveItem (Grid.Rows - 1)
       Case 10
            Grid.Rows = 1
            Grid.Rows = 6
End Select
End Sub

Private Sub DrgData_DblClick()
cmdExec_Click 1
End Sub

Private Sub DrgData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
       Case vbKeyUp
            If DrgData.Row = 1 Then
               If objTAB.Tab = 0 Then
                  txtCari.SetFocus
               ElseIf objTAB.Tab = 1 Then
                  Grid.SetFocus
               End If
            End If
End Select
End Sub

Private Sub DrgData_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdExec_Click 1
End Sub

Private Sub DrgData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    If DrgData.Rows > 1 Then
        If DrgData.TextMatrix(1, 0) <> "" Then
            Dim m As String, i As Long
            For i = 0 To DrgData.Cols - 1
               m = m & DrgData.TextMatrix(DrgData.Row, i) & "|"
            Next i
            DrgData.SetFocus
            CallByName CallObj, procName, VbMethod, m
            'CallObj.SetFocus
        End If
    End If
End If
End Sub

Private Sub Form_Activate()
Me.Show
End Sub

Private Sub Form_Deactivate()
'Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RemoveFindItem Me.Tag
Set DrgData.DataSource = Nothing
Set RCxx = Nothing
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
If Col = 0 Then
   If Trim(Grid.TextMatrix(Row - 1, 3)) = "" Then
      Cancel = True
   End If
ElseIf Row > 2 Then
   If Trim(Grid.TextMatrix(Row - 1, 0)) = "" Then
      Cancel = True
   End If
End If
Select Case Col
       Case 1, 2
          If Trim(Grid.TextMatrix(Row, Col - 1)) = "" Then Cancel = True
      Case 3
          If Trim(Grid.TextMatrix(Row, 1)) = "" Then Cancel = True
End Select
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyInsert Then
   Grid.AddItem ""
ElseIf KeyCode = vbKeyDelete Then
   If Grid.Rows > 2 Then Grid.RemoveItem Grid.Row
ElseIf KeyCode = vbKeyC And Shift = 2 Then
    Clipboard.Clear
    Clipboard.SetText Chr(255) & vbKeyTab & Grid.TextMatrix(Grid.Row, 0) & vbKeyTab & Grid.TextMatrix(Grid.Row, 1) & vbKeyTab & Grid.TextMatrix(Grid.Row, 2) & vbKeyTab & Grid.TextMatrix(Grid.Row, 3)
ElseIf KeyCode = vbKeyV And Shift = 2 Then
    Dim cc
    cc = Split(Clipboard.GetText, vbKeyTab)
    If UBound(cc) > 0 Then
       If cc(0) = Chr(255) Then
          If Grid.TextMatrix(Grid.Row - 1, 3) <> "" Then
                Grid.TextMatrix(Grid.Row, 0) = cc(1)
                Grid.TextMatrix(Grid.Row, 1) = cc(2)
                Grid.TextMatrix(Grid.Row, 2) = cc(3)
                Grid.TextMatrix(Grid.Row, 3) = cc(4)
                If Grid.Row < Grid.Rows - 1 Then Grid.Row = Grid.Row + 1
          End If
       End If
    End If
End If
End Sub

Private Sub objTAB_Click(PreviousTab As Integer)
On Error Resume Next
Select Case objTAB.Tab
       Case 0
            txtCari.SetFocus
       Case 1
            Grid.SetFocus
End Select
End Sub

Private Sub txtCari_GotFocus()
BlokX txtCari, 0
End Sub

Private Sub txtCari_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyDown
            DrgData.SetFocus
End Select
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdExec_Click 0: KeyAscii = 0
End Sub


