VERSION 5.00
Begin VB.Form frm_util_print_redirect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_util_print_fak.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Awal posisi"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   2460
      Width           =   1290
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mulai Dari:"
      Height          =   690
      Left            =   165
      TabIndex        =   9
      Top             =   1665
      Width           =   3495
      Begin VB.OptionButton Option4 
         Caption         =   "Kiri Kertas"
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
         Left            =   1800
         TabIndex        =   11
         Top             =   315
         Width           =   1410
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Kanan Kertas"
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
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Custom"
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
      Left            =   225
      TabIndex        =   1
      Top             =   435
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1200
      Left            =   165
      TabIndex        =   8
      Top             =   435
      Width           =   3495
      Begin SysInfo_Nardhika.vbTextBox txtMulai 
         Height          =   315
         Index           =   0
         Left            =   615
         TabIndex        =   3
         Top             =   615
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Alignment       =   2
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
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
      Begin SysInfo_Nardhika.vbTextBox txtMulai 
         Height          =   315
         Index           =   1
         Left            =   2100
         TabIndex        =   5
         Top             =   615
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Alignment       =   2
         BackColor       =   16777215
         BackColorMain   =   14737632
         DownButton      =   0   'False
         BorderColor     =   33023
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sampai"
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
         Left            =   1425
         TabIndex        =   4
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dari"
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
         Left            =   165
         TabIndex        =   2
         Top             =   660
         Width           =   315
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Angsuran aktif"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin SysInfo_Nardhika.vbButton vbButton1 
      Default         =   -1  'True
      Height          =   375
      Left            =   3810
      TabIndex        =   6
      Top             =   525
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   661
      BTYPE           =   5
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "_frm_util_print_fak.frx":038A
      PICN            =   "_frm_util_print_fak.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfo_Nardhika.vbButton vbButton2 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3810
      TabIndex        =   7
      Top             =   960
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Batal"
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
      MPTR            =   1
      MICON           =   "_frm_util_print_fak.frx":0740
      PICN            =   "_frm_util_print_fak.frx":075C
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
Attribute VB_Name = "frm_util_print_redirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Option1_Click()
On Error Resume Next
Frame1.Enabled = False
End Sub

Private Sub Option2_Click()
On Error Resume Next
Frame1.Enabled = True
txtMulai(0).SetFocus
txtMulai(0).Text = 1
txtMulai(1).Text = 1
End Sub

Private Sub vbButton1_Click()
On Error Resume Next
Dim hText As String, i As Integer
Dim inPos As Boolean, hAlamat As String
Dim isNext As String

           If Option1.Value Then
               txtMulai(0).Text = Form6.GridMe.Row
               txtMulai(1).Text = Form6.GridMe.Row
           Else
               txtMulai(0).Text = Val(txtMulai(0).Text) '+ 1
               txtMulai(1).Text = Val(txtMulai(1).Text) ' + 1
           End If
            
           If Option3.Value Then
              isNext = "kanan_atas"
           Else
              isNext = "kiri_atas"
           End If
            
            With Form6

                If Option1.Value Then 'Aktif Angsuran
CetakSatuan:
                   If Option3.Value Then 'Kanan Atas
                      If Check1.Value = 1 Then
                         cetakKePrinter IIf(Option1.Value, Val(txtMulai(0)), .GridMe.Row), "kanan_atas"
                      Else
                         cetakKePrinter IIf(Option1.Value, Val(txtMulai(0)), .GridMe.Row), "kanan_bawah"
                      End If
                   Else 'Kiri Atas
                      If Check1.Value = 1 Then
                         cetakKePrinter IIf(Option1.Value, Val(txtMulai(0)), .GridMe.Row), "kiri_atas"
                      Else
                         cetakKePrinter IIf(Option1.Value, Val(txtMulai(0)), .GridMe.Row), "kiri_bawah"
                      End If
                   End If
                Else 'Custom
                 If Val(txtMulai(1).Text) > Val(txtMulai(0).Text) Then
                    For i = Val(txtMulai(0).Text) To Val(txtMulai(1).Text)
                        If isNext = "kanan_atas" Then
                           cetakKePrinter i + 1, isNext
                           
                           isNext = "kiri_atas"
                        ElseIf isNext = "kiri_atas" Then
                          If i + 1 <= Val(txtMulai(1).Text) Then
                             If (i <= 2) Then
                                If Option4.Value Then
                                   'MsgBox "tengah_atas " & i & "  " & i + 1
                                   If Check1.Value = 1 Then
                                      cetakKePrinter i + 1, "tengah_atas"
                                   Else
                                      cetakKePrinter i + 1, "tengah_bawah" 'tengah_atas
                                   End If
                                Else
                                   'MsgBox "tengah_bawah " & i & "  " & i + 1
                                   cetakKePrinter i + 1, "tengah_bawah"
                                End If
                             Else
                                'MsgBox "tengah_bawah " & i & "  " & i + 1
                                cetakKePrinter i + 1, "tengah_bawah"
                             End If
                             i = i + 1
                          Else
                            isNext = "kiri_bawah"
                            'MsgBox isNext & i
                            cetakKePrinter i + 1, isNext
                          End If
                        End If
                        
                    Next i
                 ElseIf Val(txtMulai(1).Text) = Val(txtMulai(0).Text) Then
                    GoSub CetakSatuan
                 End If
                End If
            End With
            Unload Me
End Sub

Sub cetakKePrinter(index As Integer, posPrint As String)
On Error Resume Next
Dim hText As String, Filename As String, myText As String, hAlamat As String
Dim deftgl As String
With Form6
     '
    Select Case LCase(posPrint)
           Case "kanan_atas", "kanan_bawah"
                deftgl = .GridMe.TextMatrix(index, 2)
                If Trim(deftgl) = "" Then
                   If GetSetting("vbbego.com\SISRent", "Setting", "opt5", "") = "1" Then
                      deftgl = Format(Date, "dd-mm-yyyy")
                   End If
                End If
           
                If LCase(posPrint) = "kanan_atas" Then
                   Filename = StripPath(App.Path) & "_support\page_right_top.dll"
                Else
                   Filename = StripPath(App.Path) & "_support\page_right_down.dll"
                End If
                
                myText = String(FileLen(Filename), 0)
                Open Filename For Binary As #1
                    Get #1, , myText
                Close #1
                hText = myText
                hAlamat = Replace(.txtFieldsLine, vbCrLf, "")
                hText = Replace(hText, String(17, "j"), AddSpace(.txtFields(0).Text, 17))
                hText = Replace(hText, String(37, "k"), AddSpace("Rp. " & fNum(.GridMe.TextMatrix(index, 3), True) & "         No." & .GridMe.TextMatrix(index, 0), 37))
                hText = Replace(hText, String(37, "l"), AddSpace(.txtFields(15), 37))
                hText = Replace(hText, String(37, "t"), AddSpace(hAlamat, 37))
                hText = Replace(hText, String(37, "n"), AddSpace(Mid(hAlamat, 38), 37))
                hText = Replace(hText, String(37, "o"), AddSpace(.txtFields(5), 37))
                hText = Replace(hText, String(37, "!"), AddSpace(.txtFields(12), 37))
                hText = Replace(hText, String(37, "|"), AddSpace(.txtFields(16) & " - " & .txtFields(19), 37))
                hText = Replace(hText, String(10, "r"), AddSpace(deftgl, 10))

           Case "kiri_atas", "kiri_bawah"
                deftgl = .GridMe.TextMatrix(index, 2)
                If Trim(deftgl) = "" Then
                   If GetSetting("vbbego.com\SISRent", "Setting", "opt5", "") = "1" Then
                      deftgl = Format(Date, "dd-mm-yyyy")
                   End If
                End If
           
                If LCase(posPrint) = "kiri_atas" Then
                   Filename = StripPath(App.Path) & "_support\page_left_top.dll"
                Else
                   Filename = StripPath(App.Path) & "_support\page_left_down.dll"
                End If

                myText = String(FileLen(Filename), 0)
                Open Filename For Binary As #1
                    Get #1, , myText
                Close #1
                hText = myText
                hAlamat = Replace(.txtFieldsLine, vbCrLf, "")
                hText = Replace(hText, String(17, "a"), AddSpace(.txtFields(0).Text & " - " & .GridMe.TextMatrix(index, 0), 17))
                hText = Replace(hText, String(37, "b"), AddSpace("Rp. " & fNum(.GridMe.TextMatrix(index, 3), True) & "         No." & .GridMe.TextMatrix(index, 0), 37))
                hText = Replace(hText, String(37, "c"), AddSpace(.txtFields(15), 37))
                hText = Replace(hText, String(37, "d"), AddSpace(hAlamat, 37))
                hText = Replace(hText, String(37, "e"), AddSpace(Mid(hAlamat, 38), 37))
                hText = Replace(hText, String(37, "f"), AddSpace(.txtFields(5), 37))
                hText = Replace(hText, String(37, "g"), AddSpace(.txtFields(12), 37))
                hText = Replace(hText, String(37, "h"), AddSpace(.txtFields(16) & " - " & .txtFields(19), 37))
                hText = Replace(hText, String(10, "i"), AddSpace(deftgl, 10))
           Case "tengah_atas", "tengah_bawah"
                deftgl = .GridMe.TextMatrix(index, 2)
                If Trim(deftgl) = "" Then
                   If GetSetting("vbbego.com\SISRent", "Setting", "opt5", "") = "1" Then
                      deftgl = Format(Date, "dd-mm-yyyy")
                   End If
                End If
           
                If LCase(posPrint) = "tengah_atas" Then
                   Filename = StripPath(App.Path) & "_support\page_two_face_top.dll"
                Else
                   Filename = StripPath(App.Path) & "_support\page_two_face_down.dll"
                End If

                myText = String(FileLen(Filename), 0)
                Open Filename For Binary As #1
                    Get #1, , myText
                Close #1
                hText = myText
                
                hAlamat = Replace(.txtFieldsLine, vbCrLf, "")
                hText = Replace(hText, String(17, "a"), AddSpace(.txtFields(0).Text, 17))
                hText = Replace(hText, String(37, "b"), AddSpace("Rp. " & fNum(.GridMe.TextMatrix(index, 3), True) & "         No." & .GridMe.TextMatrix(index, 0), 37))
                hText = Replace(hText, String(37, "c"), AddSpace(.txtFields(15), 37))
                hText = Replace(hText, String(37, "d"), AddSpace(hAlamat, 37))
                hText = Replace(hText, String(37, "e"), AddSpace(Mid(hAlamat, 38), 37))
                hText = Replace(hText, String(37, "f"), AddSpace(.txtFields(5), 37))
                hText = Replace(hText, String(37, "g"), AddSpace(.txtFields(12), 37))
                hText = Replace(hText, String(37, "h"), AddSpace(.txtFields(16) & " - " & .txtFields(19), 37))
                hText = Replace(hText, String(10, "i"), AddSpace(deftgl, 10))
                
                deftgl = .GridMe.TextMatrix(index + 1, 2)
                If Trim(deftgl) = "" Then
                   If GetSetting("vbbego.com\SISRent", "Setting", "opt5", "") = "1" Then
                      deftgl = Format(Date, "dd-mm-yyyy")
                   End If
                End If
                
                hText = Replace(hText, String(17, "j"), AddSpace(.txtFields(0).Text, 17))
                hText = Replace(hText, String(37, "k"), AddSpace("Rp. " & fNum(.GridMe.TextMatrix(index + 1, 3), True) & "         No." & .GridMe.TextMatrix(index + 1, 0), 37))
                hText = Replace(hText, String(37, "l"), AddSpace(.txtFields(15), 37))
                hText = Replace(hText, String(37, "t"), AddSpace(hAlamat, 37))
                hText = Replace(hText, String(37, "n"), AddSpace(Mid(hAlamat, 38), 37))
                hText = Replace(hText, String(37, "o"), AddSpace(.txtFields(5), 37))
                hText = Replace(hText, String(37, "!"), AddSpace(.txtFields(12), 37))
                hText = Replace(hText, String(37, "|"), AddSpace(.txtFields(16) & " - " & .txtFields(19), 37))
                hText = Replace(hText, String(10, "r"), AddSpace(deftgl, 10))
                            
    End Select
    
    If Trim(hText) <> "" Then
        LoadPrintRedirect
        WriteToPrinter hText
        ClosePrintRedirect
        'Open "C:\test\" & Format(Time, "HHMMSS") & Int(Rnd(255)) & posPrint & "_" & index & ".txt" For Output As #1
        '     Print #1, hText
        'Close #1
    End If
End With
End Sub

Private Sub vbButton2_Click()
Unload Me
End Sub
