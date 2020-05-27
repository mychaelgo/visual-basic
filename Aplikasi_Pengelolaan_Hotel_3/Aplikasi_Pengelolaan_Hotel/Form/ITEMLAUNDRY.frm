VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form NEW_LAUNDRY 
   Caption         =   "NEW_LAUNDRY"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "ITEMLAUNDRY.frx":0000
   ScaleHeight     =   6840
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   240
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.PictureBox Tabel 
      Height          =   4455
      Left            =   120
      Picture         =   "ITEMLAUNDRY.frx":2B01
      ScaleHeight     =   4395
      ScaleWidth      =   5715
      TabIndex        =   19
      Top             =   2160
      Width           =   5775
      Begin MSComctlLib.ListView Lv 
         Height          =   3615
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IDITEM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Laundry"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DryClean"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox cmbIDItem 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      ItemData        =   "ITEMLAUNDRY.frx":5602
      Left            =   2280
      List            =   "ITEMLAUNDRY.frx":560F
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox tname 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox tlaundry 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox tdry 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   2160
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox tIDItem 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin LvButtons.lvButtons_H btnEXIT 
      Height          =   855
      Left            =   4560
      TabIndex        =   6
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      Caption         =   "&EXIT"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "ITEMLAUNDRY.frx":563D
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnAdd 
      Height          =   960
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "ADD"
      Top             =   3360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Caption         =   "&SAVE"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "ITEMLAUNDRY.frx":72F1
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnNew 
      Height          =   915
      Left            =   4560
      TabIndex        =   0
      ToolTipText     =   "&NEW"
      Top             =   2280
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Caption         =   "&NEW"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "ITEMLAUNDRY.frx":7EE4
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnDel 
      Height          =   855
      Left            =   4560
      TabIndex        =   15
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      Caption         =   "&DeLete"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "ITEMLAUNDRY.frx":91D1
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H cmdSearch 
      Height          =   615
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "SEARCH"
      Top             =   3240
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1085
      Caption         =   "&SEARCH"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   65535
      cBhover         =   255
      LockHover       =   3
      cGradient       =   255
      Gradient        =   4
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "ITEMLAUNDRY.frx":ADF8
      ImgSize         =   48
      cBack           =   -2147483633
      mPointer        =   99
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "&NEW"
            Object.ToolTipText     =   "&NEW"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "&Delete"
            Object.ToolTipText     =   "&DeLeTe"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "&Fresh"
            Object.ToolTipText     =   "&Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "&Exit"
            Object.ToolTipText     =   "&EXIT"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   -480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITEMLAUNDRY.frx":C03C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITEMLAUNDRY.frx":1289E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITEMLAUNDRY.frx":13F21
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITEMLAUNDRY.frx":15B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITEMLAUNDRY.frx":16542
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   1095
      Left            =   360
      TabIndex        =   21
      Top             =   840
      Width           =   5295
      _cx             =   4203644
      _cy             =   4196235
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4200
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   3960
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   1920
      TabIndex        =   14
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   1920
      TabIndex        =   13
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   1920
      TabIndex        =   12
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   4320
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "laundry"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dry clean"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   9
      Top             =   5640
      Width           =   1230
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   3960
      Y2              =   6360
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   1920
      TabIndex        =   8
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDitem"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   765
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4200
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   5760
      X2              =   5760
      Y1              =   2160
      Y2              =   6360
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4320
      X2              =   4320
      Y1              =   2160
      Y2              =   6360
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4320
      X2              =   5760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4320
      X2              =   5760
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "NEW_LAUNDRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
Call btnEXIT_Click
End Sub

Sub isitabel()
Lv.ListItems.Clear
 If Rs.State = 1 Then Rs.Close
        Rs.Open "Select * From itemlaundry", KOneKsi, 3, 3
            If Not Rs.EOF Then
                While Not Rs.EOF
                    Set List = Lv.ListItems.Add(, , CEKNULL(Rs!idItem))
                        List.SubItems(1) = CEKNULL(Rs!Name)
                        List.SubItems(2) = CEKNULL(Rs!laundry)
                        List.SubItems(3) = CEKNULL(Rs!dRYClean)
                
                        
                    Rs.MoveNext
                Wend
            End If
End Sub
Sub auto()
Dim x As String
    x = Format(Date, "yymm")
With Rs

If .State = 1 Then .Close
.Open "select * from ItemLaundry ORDER BY iditem ASC ", KOneKsi, adOpenKeyset, adLockReadOnly

    If .RecordCount = 0 Then

        tIDItem = "JL" + Format(Date, "yymm") + "001"
    Else
        .MoveLast

        If Left(Rs!idItem, 6) = "JL" + x Then
        tIDItem = Right(Rs!idItem, 3) + 1
        tIDItem = "JL" + Format(Date, "yymm") + Left("000", 3 - Len(tIDItem)) + tIDItem
        Else

         tIDItem = "JL" + Format(Date, "yymm") + "001"
        End If
    End If
End With
End Sub


Private Sub btnDel_Click()
If Rs.State = 1 Then Rs.Close

If cmbIDItem <> "" Then
    Rs.Open "select * from ItemLAundry where idItem='" & cmbIDItem & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            
            pesan = MsgBox("Are You Sure Delete?" + " " + cmbIDItem, vbQuestion + vbYesNo, "DELETE")
                
                If pesan = vbYes Then
                    KOneKsi.Execute "Delete * from ItemLaundry where idItem='" & cmbIDItem & "'"
                        MsgBox "100 % Sucessfully Delete" + " " + cmbIDItem, vbExclamation, "DELETE"
                 End If
        
        End If
        
Else
    cmbIDItem.Visible = True
    tIDItem.Visible = False
    Call cmdSearch_Click
                    
End If
                    
End Sub

Private Sub cmbIDItem_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from ItemLAundry where IdItem='" & cmbIDItem & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            tname = Rs!Name
            tlaundry = Rs!laundry
            tdry = Rs!dRYClean
         End If

End Sub

Private Sub cmdSearch_Click()
cmbIDItem.Visible = True
tIDItem.Visible = False
cmbIDItem.Clear
kunci Me
cmbIDItem.SetFocus
cmbIDItem.Locked = False
cmbIDItem.Clear
If Rs.State = 1 Then Rs.Close

    Rs.Open "select * from ItemLAundry", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbIDItem.AddItem Rs!idItem
                Rs.MoveNext
            Wend
        End If
End Sub





Private Sub tdry_Change()
If Not IsNumeric(tdry) Then tdry = 0
End Sub

Private Sub tlaundry_Change()
If Not IsNumeric(tlaundry) Then tlaundry = 0
End Sub



Private Sub btnAdd_Click()

If tname.Text = "" Then
    MsgBox "Please Enter Name", vbExclamation, "mYHoTEL"
    tname.SetFocus
Else
If tlaundry.Text = "" Then
    MsgBox "Please Enter LAundry", vbExclamation, "mYHoTEL"
    tlaundry.SetFocus
Else
If tdry.Text = "" Then
    MsgBox "Please Enter DryClean", vbExclamation, "mYHoTEL"
    tlaundry.SetFocus
Else

If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from itemlaundry where IDItem='" & tIDItem & "'", KOneKsi, 3, 3
        If Rs.EOF Then
            KOneKsi.Execute "Insert into itemlaundry(Iditem,Name,Laundry,DryClean)values('" & tIDItem & "','" & Replace(tname, "'", "''") & "' ,'" & Replace(tlaundry, "'", "''") & "','" & Replace(tdry, "'", "''") & "')"
             MsgBox ("Data added. for User") + " " + tname.Text, vbInformation, "mYHoTEL"
        Else
            MsgBox "Sorry is Have IDItem" + " " + Rs!idItem, vbExclamation, "mYHoTEL"
        End If
End If
End If
End If
End Sub

Private Sub btnEXIT_Click()
pesan = MsgBox("Are You Sure?", vbQuestion + vbYesNo, "TurnOff")
If pesan = vbYes Then Unload Me
splashHidup
End Sub

Private Sub btnNew_Click()

bersih Me

Buka Me

tIDItem.Locked = True
auto
cmbIDItem.Visible = False
tIDItem.Visible = True
tname.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then SendKeys "+{TAB}", True
If KeyAscii = 13 Then SendKeys ("{tab}")

End Sub

Private Sub Form_Load()
OPENDATA
about.Movie = App.Path & "\Document\New_LAundry.swf"
about.Play
awal Me
splashMati
isitabel
Me.Height = 7350
Me.Width = 6420
kunci Me
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "&NEW" Then
    Tabel.Visible = False
    bersih Me
    Buka Me
    cmbIDItem.Visible = False
    tIDItem.Visible = True
    tIDItem.Locked = True
    auto
    tname.SetFocus
    
End If

If Button.Key = "&Fresh" Then
    Tabel.Visible = True
    isitabel
End If
    
Dim Hapus As String


If Button.Key = "&Delete" Then
    Hapus = InputBox("Please INSERT IDUSER for Delete", "DELETE")
    If Hapus <> "" Then
            If Rs.State = 1 Then Rs.Close
            Rs.Open "select * from itemlaundry where idItem='" & Hapus & "'", KOneKsi, 3, 3
                If Not Rs.EOF Then
                    
                    pesan = MsgBox("Are You Sure Delete?" + " " + pesan, vbQuestion + vbYesNo, "DELETE")
                        
                        If pesan = vbYes Then
                            KOneKsi.Execute "Delete * from ItemLAundry where idItem='" & Hapus & "'"
                                MsgBox "100 % Sucessfully Delete" + " " + Hapus, vbExclamation, "DELETE"
                         End If
                
                    Else
                        MsgBox "Sorry  Not Data", vbExclamation, "No Data"
                        
                
                End If
    End If
End If
    
If Button.Key = "&Exit" Then Unload Me
End Sub
