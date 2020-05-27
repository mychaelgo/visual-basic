VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.1#0"; "osenvistasuite2009.ocx"
Begin VB.Form frm_Upgrade 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "OsenVistaSuite 2009 Pro Upgrade Utility"
   ClientHeight    =   5055
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Upgrade.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaTextBox TxtFiles 
      Height          =   345
      Left            =   210
      TabIndex        =   0
      Top             =   840
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8421504
      ButtonCaption   =   ""
      ButtonPicture   =   "frm_Upgrade.frx":617A
      ButtonVisible   =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      BackColorOver   =   12648447
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaCheckBox ChkPublic 
      Height          =   255
      Left            =   4110
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   450
      BackColor       =   16767935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Alignment       =   1
      Caption         =   "Downgrade to v12.24.0.12."
      Style           =   1
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OsenVistaSuite 2009 Pro Upgrade Utility"
      TitleTop        =   7
      icon            =   "frm_Upgrade.frx":6714
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   10
      Top             =   4515
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   953
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   0   'False
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   2
      HaveXPForm      =   -1  'True
      BorderStyle     =   6
      WindowColor     =   3
      PWidth1         =   158
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "http://www.osenxpsuite.net"
      pTextAlignment1 =   0
      PanelPicture1   =   "frm_Upgrade.frx":6AAE
      PanelPicAlignment1=   0
      PWidth2         =   1000
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Copyright (c) 2009 Osen Kusnadi | All Right Reserved."
      pTextAlignment2 =   0
      PanelPicture2   =   "frm_Upgrade.frx":6ACA
      PanelPicAlignment2=   0
      GradientColor1  =   15382160
      GradientColor2  =   12752244
   End
   Begin VistaSuitePro.OsenVistaProgressBar pBar 
      Height          =   225
      Left            =   1770
      TabIndex        =   8
      Top             =   4170
      Visible         =   0   'False
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   2871848
      Value           =   100
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtData 
      Height          =   540
      Left            =   2880
      TabIndex        =   7
      Top             =   2820
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   953
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8370596
      MultiLine       =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaButton CmdSelectAll 
      Height          =   345
      Left            =   210
      TabIndex        =   6
      Top             =   4110
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Caption         =   "&Select All"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   99
      MICON           =   "frm_Upgrade.frx":6AE6
      PICN            =   "frm_Upgrade.frx":6C48
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frm_Upgrade.frx":6FE2
      BinaryImageOver =   "frm_Upgrade.frx":6FFA
   End
   Begin VistaSuitePro.OsenVistaButton cmdUpgrade 
      Height          =   345
      Left            =   4980
      TabIndex        =   5
      Top             =   4110
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Caption         =   "&Upgrade Now"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   99
      MICON           =   "frm_Upgrade.frx":7012
      PICN            =   "frm_Upgrade.frx":7174
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frm_Upgrade.frx":770E
      BinaryImageOver =   "frm_Upgrade.frx":7726
   End
   Begin VistaSuitePro.OsenVistaListBox LstFiles 
      Height          =   2775
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4895
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   0
      BackSelected    =   7381139
      BackSelectedG1  =   16777215
      BackSelectedG2  =   8632490
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      ItemTextLeft    =   20
      BorderColor     =   13603685
      Lstyle          =   1
      ShowHeader      =   -1  'True
      HeaderFormatString=   "Select your file(s) which would you like to upgrade;410;0;0;;-1"
      Columns         =   1
      ShowGridLines   =   -1  'True
      MaxAllColumnWidth=   410
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientTop=   14737632
      HeaderGradientBottom=   4210752
      BinaryImage     =   "frm_Upgrade.frx":773E
      WindowColor     =   3
      Begin VistaSuitePro.OsenVistaListBox lVBP 
         Height          =   1545
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Visible         =   0   'False
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2725
         Appearance      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontNormal      =   16777215
         BackSelected    =   7381139
         BackSelectedG1  =   16777215
         BackSelectedG2  =   8632490
         WordWrap        =   0   'False
         ItemHeightAuto  =   0   'False
         ItemOffset      =   2
         SelectModeStyle =   2
         BorderColor     =   13603685
         ShowHeader      =   -1  'True
         ShowGridLines   =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IMGLIST         =   ""
         ForeColorSelected=   16576
         BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderGradientTop=   8569007
         HeaderGradientBottom=   4487779
         HeaderGradientAllow=   -1  'True
         HeaderForeColor =   16777215
         BinaryImage     =   "frm_Upgrade.frx":7756
         WindowColor     =   3
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your project directory:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   540
      Width           =   3585
   End
End
Attribute VB_Name = "frm_Upgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyGUI1        As String
Private MyGUI2        As String
Private Const MyGUI11       As String = "Object={198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.1#0; OSENVISTASUITE2009.ocx"
Private Const MyGUI21       As String = "Object = ""{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.1#0""; ""OSENVISTASUITE2009.ocx"""

Private Const OLD_OCX_REF   As String = " OSENXPSUITE2007."
Private Const OLD_OCX_REV   As String = " OSENXPSUITE2008."
Private Const OLD_OCX_REV2   As String = " OSENXPSUITE2009OCX."
Private Const OLD_OCX_REP1   As String = " VistaSuiteXE."
Private Const OLD_OCX_REP   As String = " VistaSuiteProX."
Private Const NEW_OCX_REF   As String = " VistaSuitePro."

Private StrData, oxpDATA
Private bEndTask As Boolean
Private WithEvents SysTray As CLS_SysTray
Attribute SysTray.VB_VarHelpID = -1

Private Sub ChangeReferenceObject(sFile As String, _
                                  IsProjectFile As Boolean)
    On Error Resume Next
    Dim Izx As Long
    Dim sxStr As String

    If IsProjectFile Then
        If ChkPublic.Value Then
            StrData = Replace$(StrData, MyGUI11, MyGUI1, , , vbTextCompare)
        Else
            StrData = Replace$(StrData, MyGUI1, MyGUI11, , , vbTextCompare)
        End If

        Izx = InStr(1, StrData, "[MS Transaction Server]", vbTextCompare)

        If Izx Then
            StrData = Left$(StrData, Izx - 2)
        End If

        Open sFile For Output As #1
        Print #1, StrData
        Close #1
    Else
        oxpDATA = Replace$(StrData, MyGUI2, MyGUI21, , , vbTextCompare)
        sxStr = Replace$(oxpDATA, OLD_OCX_REF & "OsenXP", NEW_OCX_REF & "OsenVista", , , vbTextCompare)
        sxStr = Replace$(sxStr, OLD_OCX_REV & "OsenXP", NEW_OCX_REF & "OsenVista", , , vbTextCompare)
        sxStr = Replace$(sxStr, OLD_OCX_REV2 & "OsenXP", NEW_OCX_REF & "OsenVista", , , vbTextCompare)
        sxStr = Replace$(sxStr, OLD_OCX_REF, NEW_OCX_REF, , , vbTextCompare)
        sxStr = Replace$(sxStr, OLD_OCX_REV, NEW_OCX_REF, , , vbTextCompare)
        sxStr = Replace$(sxStr, OLD_OCX_REP, NEW_OCX_REF, , , vbTextCompare)
        sxStr = Replace$(sxStr, OLD_OCX_REP1, NEW_OCX_REF, , , vbTextCompare)
        sxStr = Replace$(sxStr, OLD_OCX_REV2, NEW_OCX_REF, , , vbTextCompare)
        Open sFile For Output As #1
        Print #1, sxStr
        Close #1
    End If

End Sub

Private Function GetOldGUID(sFile As String) As Boolean

    Dim myoldguidx As String
    myoldguidx = GetStringFromFile(sFile)
    Dim adata
    Dim n As Long
    Dim l As Long
    Dim z As Long
    Dim IsVBP As Boolean
        
    GetOldGUID = False
        
    Const frm As String = "Object = ""{"
    Const vbp As String = "Object={"
    Dim ext As String
        
    If Right$(LCase$(sFile), 3) = "vbp" Then
        ext = vbp
        IsVBP = True
    Else
        ext = frm
        IsVBP = False
    End If
        
    adata = Split(myoldguidx, vbCrLf)
    n = UBound(adata)
        
    If n > 50 Then
        n = 50
    End If
        
    For l = 0 To n

        If InStr(1, adata(l), ext) Then
            If InStr(1, LCase$(adata(l)), "osen") Then
                If InStr(1, LCase$(adata(l)), "ocx") Then
                    If IsVBP Then
                        MyGUI1 = adata(l)
                        z = InStr(1, MyGUI1, MyGUI11, vbTextCompare)
                        GetOldGUID = z
                    Else
                        MyGUI2 = adata(l)
                        z = InStr(1, MyGUI2, MyGUI21, vbTextCompare)
                        GetOldGUID = z
                    End If

                    Exit For
                End If
            End If
        End If

    Next

End Function

Private Function ConvertbyFile(sfilename As String) As Boolean
    On Error Resume Next
    Dim bForm As Boolean
    bForm = IIf(Right$(UCase$(sfilename), 3) = "VBP", -1, 0)

    If Not GetOldGUID(sfilename) Then
        StrData = GetStringFromFile(sfilename)
        ChangeReferenceObject sfilename, bForm
        ConvertbyFile = True
    End If

    DoEvents
End Function

Private Sub CmdSelectAll_Click()
    On Error Resume Next

    If LstFiles.ListCount > 0 Then
        Dim J As Long

        For J = 0 To LstFiles.ListCount - 1
            LstFiles.Selected(J) = True
        Next J

        LstFiles.LockUpdate = False
    End If

    On Error GoTo 0
End Sub

Private Sub cmdUpgrade_Click()
    On Error Resume Next
    Dim h As Long
    Dim i As Long

    If MyGUI11 = MyGUI1 Then
        MsgBoxGT "You doesn't need for converting your project!" & vbCrLf & "Your project has already been used a newest version of OsenVistaSuite" & vbCrLf & vbCrLf & "GUID Information:" & vbCrLf & MyGUI11, vbInformation
        Exit Sub
    End If

    pBar.Visible = True
    Screen.MousePointer = vbHourglass
    i = 0
    pBar.Max = LstFiles.SelectedCount

    For h = 0 To LstFiles.ListCount - 1

        If LstFiles.Selected(h) Then
            If ConvertbyFile(LstFiles.List(h)) Then
                LstFiles.Item(h).lColor = vbBlue
            Else
                LstFiles.Item(h).BackColor = 0&
                LstFiles.Item(h).lColor = vbGreen
                
            End If
            
            i = i + 1
            LstFiles.Selected(h) = False
            LstFiles.ListIndex = h
            pBar.Value = i
        End If

        If bEndTask Then Exit Sub

        DoEvents
    Next h

    MsgBoxGT "Upgrading your project successfull.", vbInformation, "Upgrade Utility", 5
    pBar.StopSearch
    pBar.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.OsenXPForm1.Init Me
    MyGUI1 = "Object={7862F185-E724-42A7-8512-9B99A4C10583}#1.0#0; osenxpsuite2007.ocx"
    MyGUI2 = "Object = ""{7862F185-E724-42A7-8512-9B99A4C10583}#1.0#0""; ""osenxpsuite2007.ocx"""
    DefaultWindowColor = 3
    
End Sub

Private Sub CheckOldGUID()

    If ChkPublic.Value Then
        MyGUI1 = "Object={7862F185-E724-42A7-8512-9B99A4C10583}#1.0#0; osenxpsuite2007.ocx"
        MyGUI2 = "Object = ""{7862F185-E724-42A7-8512-9B99A4C10583}#1.0#0""; ""osenxpsuite2007.ocx"""
    Else

        If lVBP.ListCount Then
            Dim sFile As String
            Dim x As Integer
            Dim y As Integer
            y = lVBP.ListCount
            x = 0

            Do While (x < y)
                sFile = lVBP.List(x)
                Dim myoldguidx As String
                myoldguidx = GetStringFromFile(sFile)
                Dim adata
                Dim n As Long
                Dim l As Long
                Dim z As Long
                adata = Split(myoldguidx, vbCrLf)
                n = UBound(adata)

                For l = 0 To n

                    If InStr(1, adata(l), "Object={") Then
                        If InStr(1, LCase$(adata(l)), "osen") Then
                            If LCase$(Right$(LCase$(adata(l)), 3)) = "ocx" Then
                                MyGUI1 = adata(l)
                                MyGUI2 = "Object = ""{"
                                z = InStr(1, adata(l), ";")
                                MyGUI2 = MyGUI2 & Mid$(adata(l), 9, z - 9) & """; """ & Mid$(adata(l), z + 2) & """"
                                Exit For
                            End If
                        End If
                    End If

                Next

                x = x + 1

                If MyGUI2 <> "" Then
                    Exit Do
                End If

            Loop
        
        End If
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    bEndTask = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not IsRegistered Then
        Set SysTray = Nothing
    End If

End Sub

Private Sub OsenXPStatusBar1_Click(iPanelNumber As Variant)
    OpenBrowser Me.hWnd

    DoEvents
End Sub

Private Sub SysTray_BalloonClick()
    OpenBrowser 0
End Sub

Private Sub TxtFiles_ButtonClick()
    On Error Resume Next
    Dim StrData
    TxtFiles.Text = ""
    TxtFiles.Text = StrGetFolder(Me.hWnd, "Select your project directory ...")

    If LenB(TxtFiles.Text) > 0 Then
        StrData = GetAllFiles(TxtFiles.Text)
        Dim i As Long
        Dim sFile As String

        If UBound(StrData) Then

            With LstFiles
                .Clear
                .LockUpdate = True
                lVBP.Clear

                For i = 1 To UBound(StrData)
                    sFile = StrData(i)

                    If Right$(UCase$(sFile), 3) = "VBP" Then
                        .AddItem sFile
                        lVBP.AddItem GetDosPath(sFile)
                        CheckOldGUID
                    ElseIf Right$(UCase$(sFile), 3) = "FRM" Then
                        .AddItem sFile
                    ElseIf Right$(UCase$(sFile), 3) = "CTL" Then
                        .AddItem sFile
                    End If

                Next i

                .AutoColumnWidth
                .LockUpdate = False

                If .ListCount Then
                    CmdSelectAll_Click
                End If

            End With

        End If
    End If

End Sub

Private Sub TxtFiles_OnEnter()
    On Error Resume Next

    If LenB(TxtFiles.Text) > 0 Then
        StrData = GetAllFiles(TxtFiles.Text)
        Dim i As Long
        Dim sFile As String

        If UBound(StrData) Then

            With LstFiles
                .Clear
                .LockUpdate = True
                lVBP.Clear

                For i = 1 To UBound(StrData)
                    sFile = StrData(i)

                    If Right$(UCase$(sFile), 3) = "VBP" Then
                        .AddItem sFile
                        lVBP.AddItem GetDosPath(sFile)
                        CheckOldGUID
                    End If

                    If Right$(UCase$(sFile), 3) = "FRM" Then
                        .AddItem sFile
                    End If

                Next i

                .AutoColumnWidth
                .LockUpdate = False
            End With

        End If
    End If

End Sub
