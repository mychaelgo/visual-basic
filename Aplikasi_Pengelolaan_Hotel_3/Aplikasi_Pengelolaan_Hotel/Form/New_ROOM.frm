VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form New_ROOM 
   Caption         =   "NEW_ROOM"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "New_ROOM.frx":0000
   ScaleHeight     =   6975
   ScaleWidth      =   6120
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   360
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox Tabel 
      Height          =   4935
      Left            =   240
      Picture         =   "New_ROOM.frx":2B01
      ScaleHeight     =   4875
      ScaleWidth      =   5715
      TabIndex        =   17
      Top             =   2160
      Width           =   5775
      Begin MSComctlLib.ListView Lv 
         Height          =   4095
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   7223
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ROOMNO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TYPE_ROOM"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PRICE"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "New_ROOM.frx":5602
      Left            =   2280
      List            =   "New_ROOM.frx":5612
      TabIndex        =   16
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox tType 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox tPrice 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      TabIndex        =   1
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox tRoom 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ComboBox cmbRoom 
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
      ItemData        =   "New_ROOM.frx":5651
      Left            =   2280
      List            =   "New_ROOM.frx":5653
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1080
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
            Picture         =   "New_ROOM.frx":5655
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New_ROOM.frx":BEB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New_ROOM.frx":D53A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New_ROOM.frx":F161
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New_ROOM.frx":FB5B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LvButtons.lvButtons_H btnEXIT 
      Height          =   855
      Left            =   4680
      TabIndex        =   6
      Top             =   5640
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
      Image           =   "New_ROOM.frx":10435
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnAdd 
      Height          =   960
      Left            =   4680
      TabIndex        =   2
      ToolTipText     =   "ADD"
      Top             =   3480
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
      Image           =   "New_ROOM.frx":120E9
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnNew 
      Height          =   915
      Left            =   4680
      TabIndex        =   7
      ToolTipText     =   "NEW"
      Top             =   2400
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
      Image           =   "New_ROOM.frx":12CDC
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnDel 
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      ToolTipText     =   "Delete"
      Top             =   4680
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
      Image           =   "New_ROOM.frx":13FC9
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H cmdSearch 
      Height          =   615
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "SEARCH"
      Top             =   3360
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
      Image           =   "New_ROOM.frx":15BF0
      ImgSize         =   48
      cBack           =   -2147483633
      mPointer        =   99
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6120
      _ExtentX        =   10795
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   1095
      Left            =   360
      TabIndex        =   19
      Top             =   960
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
      X1              =   360
      X2              =   4320
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4320
      X2              =   4320
      Y1              =   4200
      Y2              =   6600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   360
      X2              =   4320
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   1080
      Y2              =   4200
   End
   Begin VB.Label Label31 
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
      Left            =   2040
      TabIndex        =   14
      Top             =   5280
      Width           =   135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "price"
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
      Left            =   600
      TabIndex        =   13
      Top             =   5520
      Width           =   600
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
      Left            =   2040
      TabIndex        =   12
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type_room"
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
      Left            =   480
      TabIndex        =   11
      Top             =   4320
      Width           =   1320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   360
      X2              =   360
      Y1              =   4200
      Y2              =   6600
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROOMNO"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4320
      X2              =   4320
      Y1              =   2280
      Y2              =   3240
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   360
      X2              =   4320
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   360
      X2              =   4320
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   360
      X2              =   360
      Y1              =   2280
      Y2              =   3240
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   2280
      Y2              =   6600
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   2280
      Y2              =   6600
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4440
      X2              =   5880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4440
      X2              =   5880
      Y1              =   6600
      Y2              =   6600
   End
End
Attribute VB_Name = "New_ROOM"
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
        Rs.Open "Select * From Room", KOneKsi, 3, 3
            If Not Rs.EOF Then
                While Not Rs.EOF
                    Set List = Lv.ListItems.Add(, , CEKNULL(Rs!Roomno))
                        List.SubItems(1) = CEKNULL(Rs!Type_Room)
                        List.SubItems(2) = CEKNULL(Rs!amount)
                        
                        
                    Rs.MoveNext
                Wend
            End If
End Sub
'Sub auto()
'Dim x As String
'x = "101"
'With Rs
'
'If .State = 1 Then .Close
'.Open "select * from Room ORDER BY RoomNo ASC ", KOneKsi, 3, 3
'
'
'    If .RecordCount = 0 Then
'
'        tRoom = "101"
'    Else
'        .MoveLast
'
'        tRoom = Right(Rs!Roomno, 3) + 1
'
'
'    End If
'End With
'End Sub
Sub autoHome()
Dim x As String
x = "101"
With Rs

If .State = 1 Then .Close
.Open "select * from Room  where  Roomno between 101 and  199 order by Roomno asc", KOneKsi, 3, 3


    If .RecordCount = 0 Then

        tRoom = "101"
    Else
        .MoveLast

        tRoom = Right(Rs!Roomno, 3) + 1
        
        
    End If
End With


End Sub
Sub autoExecutive()
Dim x As String
x = "201"
With Rs

If .State = 1 Then .Close
.Open "select * from Room  where  Roomno between 201 and  299 order by Roomno asc", KOneKsi, 3, 3


    If .RecordCount = 0 Then

        tRoom = "201"
    Else
        .MoveLast

        tRoom = Right(Rs!Roomno, 3) + 1
        
        
    End If
End With


End Sub

Sub autoVIP()
Dim x As String
x = "301"
With Rs

If .State = 1 Then .Close
.Open "select * from Room  where  Roomno between 301 and  399 order by Roomno asc", KOneKsi, 3, 3


    If .RecordCount = 0 Then

        tRoom = "301"
    Else
        .MoveLast

        tRoom = Right(Rs!Roomno, 3) + 1
        
        
    End If
End With


End Sub

Sub autoPresident()
Dim x As String
x = "401"
With Rs

If .State = 1 Then .Close
.Open "select * from Room  where  Roomno between 401 and  499 order by Roomno asc", KOneKsi, 3, 3


    If .RecordCount = 0 Then

        tRoom = "401"
    Else
        .MoveLast

        tRoom = Right(Rs!Roomno, 3) + 1
        
        
    End If
End With


End Sub

Private Sub btnAdd_Click()

If cmbType.Text = "" Then
    MsgBox "Please Enter Type_Room", vbExclamation, "mYHoTEL"
    cmbType.SetFocus
Else
If tPrice.Text = "" Then
    MsgBox "Please Enter Price", vbExclamation, "mYHoTEL"
    tPrice.SetFocus
Else

If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from room where roomno=" & tRoom & "", KOneKsi, 3, 3
        If Rs.EOF Then
            KOneKsi.Execute "Insert into Room(Roomno,Type_Room,Amount)values(" & tRoom & ",'" & Replace(cmbType, "'", "''") & "' ," & Replace(tPrice, "'", "''") & ")"
             MsgBox ("Data added. for RoomNo") + " " + tRoom.Text, vbInformation, "mYHoTEL"
        Else
            MsgBox "Sorry is Have ROOMNO" + " " + Str(Rs!Roomno), vbExclamation, "mYHoTEL"
        End If
End If
End If

End Sub

Private Sub btnDel_Click()
If Rs.State = 1 Then Rs.Close
If cmbRoom <> "" Then
    Rs.Open "select * from Room where Roomno =" & cmbRoom & "  order by Roomno Asc", KOneKsi, 3, 3
        If Not Rs.EOF Then
            
            pesan = MsgBox("Are You Sure Delete?" + " " + cmbRoom, vbQuestion + vbYesNo, "DELETE")
                
                If pesan = vbYes Then
                    KOneKsi.Execute "Delete * from Room where Roomno=" & cmbRoom & ""
                        MsgBox "100 % Sucessfully Delete" + " " + cmbRoom, vbExclamation, "DELETE"
                 End If
        End If
        
Else
    cmbRoom.Visible = True
    tRoom.Visible = False
    Call cmdSearch_Click
                    
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
cmbRoom.Visible = False
tRoom.Visible = True
tRoom.Locked = True
cmbType.SetFocus
End Sub



Private Sub cmbRoom_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Room where Roomno=" & cmbRoom & "", KOneKsi, 3, 3
        If Not Rs.EOF Then
            cmbType = CEKNULL(Rs!Type_Room)
            tPrice = CEKNULL(Rs!amount)
            
         End If
End Sub




Private Sub cmbType_Click()
If cmbType.ListIndex = 0 Then autoHome
If cmbType.ListIndex = 1 Then autoExecutive
If cmbType.ListIndex = 2 Then autoVIP
If cmbType.ListIndex = 3 Then autoPresident
kunci Me
cmbType.Locked = False


If Rs.State = 1 Then Rs.Close

    Rs.Open "select * from Room where Type_Room='" & cmbType & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            
                tPrice = Rs!amount
         End If

End Sub

Private Sub cmdSearch_Click()
cmbRoom.Visible = True
tRoom.Visible = False
cmbRoom.SetFocus
cmbRoom.Clear
kunci Me
cmbRoom.Locked = False
cmbRoom.Clear
If Rs.State = 1 Then Rs.Close

    Rs.Open "select * from Room where status=false order by Roomno", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbRoom.AddItem Rs!Roomno
                Rs.MoveNext
            Wend
        End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then SendKeys "+{TAB}", True
If KeyAscii = 13 Then SendKeys ("{tab}")

End Sub

Private Sub Form_Load()
OPENDATA
about.Movie = App.Path & "\Document\New_Room.swf"
about.Play
awal Me
splashMati
isitabel

Me.Height = 7425
Me.Width = 6255
kunci Me
End Sub






Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "&NEW" Then
    Tabel.Visible = False
    bersih Me
    Buka Me
    cmbRoom.Visible = False
    tRoom.Visible = True
    tRoom.Locked = True
    cmbType.SetFocus
    
End If

If Button.Key = "&Fresh" Then
    Tabel.Visible = True
    isitabel
End If
    
Dim Hapus As String


If Button.Key = "&Delete" Then
    Hapus = InputBox("Please INSERT ROOMNO for Delete", "DELETE")
    If Hapus <> "" Then
        If Not IsNumeric(Hapus) Then Hapus = 0
            If Rs.State = 1 Then Rs.Close
            
                Rs.Open "select *  from Room where roomno=" & CEKNULL(Hapus) & "  and status=0 ", KOneKsi, 3, 3
                If Not Rs.EOF Then
                   

                         If RsTemp.State = 1 Then RsTemp.Close
                                RsTemp.Open "select * from Room Where  Roomno=" & CEKNULL(Hapus) & "", KOneKsi, 3, 3
                               Rs.MoveNext
                      
                             If Not RsTemp.EOF Then
                                 pesan = MsgBox("Are You Sure Delete ROOMNO?" + " " + Str(RsTemp!Roomno), vbQuestion + vbYesNo, "DELETE")

                                     If pesan = vbYes Then
                                         KOneKsi.Execute "Delete * from Room where Roomno=" & Hapus & ""
                                             MsgBox "100 % Sucessfully Delete" + " " + Hapus, vbExclamation, "DELETE"
                                      End If

                             Else
                                 MsgBox "Sorry  Not Data", vbExclamation, "No Data"
                             End If
                     
                Else
                   
                        
                        If RsTemp.State = 1 Then RsTemp.Close
                            RsTemp.Open "select * from Room Where status= true and  Roomno=" & CEKNULL(Hapus) & "", KOneKsi, 3, 3
                                If Not RsTemp.EOF Then
                                    MsgBox "Roomno" + " " + Hapus + " Is Occupied", vbExclamation, "mYHoTEL"
                                Else
                                    MsgBox "Sorry  Not Data", vbExclamation, "No Data"
                                End If
                
                End If
    End If
End If
    
If Button.Key = "&Exit" Then Unload Me
End Sub

Private Sub tPrice_Change()
If Not IsNumeric(tPrice) Then tPrice = 0
End Sub
