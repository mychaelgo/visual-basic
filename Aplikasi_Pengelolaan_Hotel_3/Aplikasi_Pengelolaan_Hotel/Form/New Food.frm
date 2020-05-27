VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form New_Food 
   Caption         =   "NEW_FOOD"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "New Food.frx":0000
   Picture         =   "New Food.frx":030A
   ScaleHeight     =   7185
   ScaleWidth      =   6165
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   360
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.PictureBox Tabel 
      Height          =   4575
      Left            =   240
      Picture         =   "New Food.frx":2E0B
      ScaleHeight     =   4515
      ScaleWidth      =   5715
      TabIndex        =   22
      Top             =   2280
      Width           =   5775
      Begin MSComctlLib.ListView Lv 
         Height          =   3975
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   7011
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IDFOOD"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dish"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Kinds"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin LvButtons.lvButtons_H btnDel 
      Height          =   735
      Left            =   4920
      TabIndex        =   8
      Top             =   4680
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
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
      Image           =   "New Food.frx":590C
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cmbIDFood 
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
      ItemData        =   "New Food.frx":7533
      Left            =   2160
      List            =   "New Food.frx":7540
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cmbDish 
      BackColor       =   &H80000008&
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
      Height          =   360
      ItemData        =   "New Food.frx":756E
      Left            =   2160
      List            =   "New Food.frx":757B
      TabIndex        =   2
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox tPrice 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
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
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox tKinds 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
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
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox tIDFood 
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox tname 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
   Begin LvButtons.lvButtons_H btnEXIT 
      Height          =   855
      Left            =   4800
      TabIndex        =   6
      Top             =   5760
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
      Image           =   "New Food.frx":75A9
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnAdd 
      Height          =   960
      Left            =   4800
      TabIndex        =   5
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
      Image           =   "New Food.frx":925D
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnNew 
      Height          =   915
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "&NEW"
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
      Image           =   "New Food.frx":9E50
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H cmdSearch 
      Height          =   615
      Left            =   240
      TabIndex        =   20
      ToolTipText     =   "SEARCH"
      Top             =   3360
      Width           =   4245
      _ExtentX        =   7488
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
      Image           =   "New Food.frx":B13D
      ImgSize         =   48
      cBack           =   -2147483633
      mPointer        =   99
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   120
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
            Picture         =   "New Food.frx":C381
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New Food.frx":12BE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New Food.frx":14266
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New Food.frx":15E8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "New Food.frx":16887
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6165
      _ExtentX        =   10874
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
      Left            =   480
      TabIndex        =   24
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      TabIndex        =   18
      Top             =   6120
      Width           =   600
   End
   Begin VB.Label Label4 
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
      TabIndex        =   17
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Kinds 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kinds"
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
      TabIndex        =   16
      Top             =   5520
      Width           =   675
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
      TabIndex        =   15
      Top             =   5280
      Width           =   135
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4560
      X2              =   6000
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4560
      X2              =   6000
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   2280
      Y2              =   6720
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6000
      X2              =   6000
      Y1              =   2280
      Y2              =   6720
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   2280
      Y2              =   3240
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4440
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   2280
      Y2              =   3240
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDfood"
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
      TabIndex        =   14
      Top             =   2640
      Width           =   840
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
      TabIndex        =   13
      Top             =   2400
      Width           =   135
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   4080
      Y2              =   6720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dish"
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
      TabIndex        =   12
      Top             =   4800
      Width           =   510
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
      Top             =   4200
      Width           =   630
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
      TabIndex        =   10
      Top             =   4560
      Width           =   135
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
      TabIndex        =   9
      Top             =   3960
      Width           =   135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4440
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   4080
      Y2              =   6720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   240
      X2              =   4440
      Y1              =   6720
      Y2              =   6720
   End
End
Attribute VB_Name = "New_Food"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub isitabel()
Lv.ListItems.Clear
 If Rs.State = 1 Then Rs.Close
        Rs.Open "Select * From Food", KOneKsi, 3, 3
            If Not Rs.EOF Then
                While Not Rs.EOF
                    Set List = Lv.ListItems.Add(, , CEKNULL(Rs!IdFood))
                        List.SubItems(1) = CEKNULL(Rs!Name)
                        List.SubItems(2) = CEKNULL(Rs!Dish)
                        List.SubItems(3) = CEKNULL(Rs!Kinds)
                        List.SubItems(4) = CEKNULL(Rs!Price)
                    Rs.MoveNext
                Wend
            End If
End Sub

Sub auto()
Dim x As String
    x = Format(Date, "yymm")
With Rs

If .State = 1 Then .Close
.Open "select * from Food ORDER BY idfood ASC ", KOneKsi, adOpenKeyset, adLockReadOnly

    If .RecordCount = 0 Then

        tIDFood = "FD" + Format(Date, "yymm") + "001"
    Else
        .MoveLast

        If Left(Rs!IdFood, 6) = "FD" + x Then
        tIDFood = Right(Rs!IdFood, 3) + 1
        tIDFood = "FD" + Format(Date, "yymm") + Left("000", 3 - Len(tIDFood)) + tIDFood
        Else

         tIDFood = "FD" + Format(Date, "yymm") + "001"
        End If
    End If
End With
End Sub



Private Sub btnAdd_Click()

If tname.Text = "" Then
    MsgBox "Please Enter Name", vbExclamation, "mYHoTEL"
    tname.SetFocus
Else
If cmbDish.Text = "" Then
    MsgBox "Please Enter Dish", vbExclamation, "mYHoTEL"
    cmbDish.SetFocus
Else
If tKinds.Text = "" Then
    MsgBox "Please Enter KINDS", vbExclamation, "mYHoTEL"
    tKinds.SetFocus
Else
If tPrice.Text = "" Then
    MsgBox "Please Enter PRICE", vbExclamation, "mYHoTEL"
    tPrice.SetFocus
Else
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Food where IDfood='" & tIDFood & "'", KOneKsi, 3, 3
        If Rs.EOF Then
            KOneKsi.Execute "Insert into Food(IDFood,Dish,Name,Kinds,Price)values('" & tIDFood & "','" & Replace(cmbDish, "'", "''") & "' ,'" & Replace(tname, "'", "''") & "','" & Replace(tKinds, "'", "''") & "'," & Replace(tPrice, "'", "''") & ")"
             MsgBox ("Data added. for User") + " " + tname.Text, vbInformation, "mYHoTEL"
        Else
            MsgBox "Sorry is Have IDFood" + " " + Rs!IdFood, vbExclamation, "mYHoTEL"
        End If
End If
End If
End If
End If
End Sub

Private Sub btnDel_Click()

If cmbIDFood <> "" Then
        
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Food where idFood='" & cmbIDFood & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            
            pesan = MsgBox("Are You Sure Delete?" + " " + cmbIDFood, vbQuestion + vbYesNo, "DELETE")
                
                If pesan = vbYes Then
                    KOneKsi.Execute "Delete * from Food where idfood='" & cmbIDFood & "'"
                        MsgBox "100 % Sucessfully Delete" + " " + cmbIDFood, vbExclamation, "DELETE"
                 End If
        
        End If
Else
    cmbIDFood.Visible = True
    tIDFood.Visible = False
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
tIDFood.Locked = True
auto
cmbIDFood.Visible = False
tIDFood.Visible = True

tname.SetFocus
End Sub



Private Sub cmbIDFood_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Food where Idfood='" & cmbIDFood & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            tname = Rs!Name
            cmbDish = Rs!Dish
            tKinds = Rs!Kinds
            tPrice = Rs!Price
        End If
End Sub

Private Sub cmdSearch_Click()
cmbIDFood.Visible = True
tIDFood.Visible = False
cmbIDFood.Clear
cmbIDFood.SetFocus
kunci Me
cmbIDFood.Locked = False
cmbIDFood.Clear
If Rs.State = 1 Then Rs.Close

    Rs.Open "select * from Food", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbIDFood.AddItem Rs!IdFood
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
isitabel
about.Movie = App.Path & "\Document\New_Food.swf"
about.Play
splashMati
awal Me
Me.Height = 7695
Me.Width = 6285
kunci Me
End Sub








Private Sub Form_Unload(Cancel As Integer)
Call btnEXIT_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "&NEW" Then
    Tabel.Visible = False
    bersih Me
    Buka Me
    cmbIDFood.Visible = False
    tIDFood.Visible = True
    tIDFood.Locked = True
    auto
    tname.SetFocus
    
End If

If Button.Key = "&Fresh" Then
    Tabel.Visible = True
    isitabel
End If
    
Dim Hapus As String


If Button.Key = "&Delete" Then
    Hapus = InputBox("Please INSERT IDFOOD for Delete", "DELETE")
    If Hapus <> "" Then
            If Rs.State = 1 Then Rs.Close
            Rs.Open "select * from FOOD where idFOOD='" & Hapus & "'", KOneKsi, 3, 3
                If Not Rs.EOF Then
                    
                    pesan = MsgBox("Are You Sure Delete?" + " " + pesan, vbQuestion + vbYesNo, "DELETE")
                        
                        If pesan = vbYes Then
                            KOneKsi.Execute "Delete * from FOOD where idFOOD='" & Hapus & "'"
                                MsgBox "100 % Sucessfully Delete" + " " + Hapus, vbExclamation, "DELETE"
                         End If
                
                    Else
                        MsgBox "Sorry  Not Data", vbExclamation, "No Data"
                        
                
                End If
    End If
End If
    
If Button.Key = "&Exit" Then Unload Me
End Sub

Private Sub tPrice_Change()
If Not IsNumeric(tPrice) = True Then tPrice = 0
End Sub
