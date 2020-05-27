VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Restaurant 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "Restaurant.frx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   2  'Custom
   Picture         =   "Restaurant.frx":030A
   ScaleHeight     =   11010
   ScaleLeft       =   1000
   ScaleMode       =   0  'User
   ScaleWidth      =   138.032
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   1800
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   5895
   End
   Begin LvButtons.lvButtons_H btnclose 
      Height          =   735
      Left            =   840
      TabIndex        =   25
      Top             =   6720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Caption         =   "&CLOSE"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Restaurant.frx":33140
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnNew 
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   6720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Caption         =   "&NEW"
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
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "Restaurant.frx":34DE6
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.TextBox tKinds 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   480
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1080
      Top             =   960
   End
   Begin VB.TextBox tIDRestaurant 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox tDate 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox cmbIDguEST 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox tname 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox taddr 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox tPhone 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox tNameFood 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.ComboBox cmbIDFood 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox tTotal 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   960
   End
   Begin VB.TextBox tPrice 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvItem 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IDFood"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kinds"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin LvButtons.lvButtons_H btnADD 
      Height          =   1335
      Left            =   7680
      TabIndex        =   26
      ToolTipText     =   "ADD"
      Top             =   3600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2355
      CapAlign        =   2
      BackStyle       =   7
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
      Image           =   "Restaurant.frx":36A8C
      cBack           =   -2147483633
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   975
      Left            =   2040
      TabIndex        =   27
      Top             =   360
      Width           =   5295
      _cx             =   4203644
      _cy             =   4196024
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kinds"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4200
      TabIndex        =   23
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4440
      TabIndex        =   21
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4440
      TabIndex        =   20
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4440
      TabIndex        =   19
      Top             =   2400
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDRestaurant"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDGUEST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Food"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2040
      TabIndex        =   15
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDFood"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   570
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   240
      TabIndex        =   13
      Top             =   5760
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1575
      Left            =   0
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1575
      Left            =   3360
      Top             =   1560
      Width           =   5775
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      Top             =   5640
      Width           =   8055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      Top             =   3240
      Width           =   9135
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label lTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6120
      TabIndex        =   11
      Top             =   3360
      Width           =   450
   End
End
Attribute VB_Name = "Restaurant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim laundry As Integer
Dim dRY As Integer
Dim Lv As ListItem
Dim money As Variant
Sub isi()

cmbIDFood.Clear

If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from food", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbIDFood.AddItem Rs!IdFood
                Rs.MoveNext
            Wend
        End If
cmbIDGuest.Clear
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from checkin", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbIDGuest.AddItem Rs!idguest
                Rs.MoveNext
            Wend
        End If

End Sub
Sub auto()
Dim x As String
    x = Format(Date, "yymm")
With Rs

If .State = 1 Then .Close
.Open "select * from restaurant ORDER BY idrestaurant ASC ", KOneKsi, 3, 3

If Rs.EOF Then
  

        tIDRestaurant = "R" + Format(Date, "yymm") + "001"
    Else
        .MoveLast

        If Left(Rs!idrestaurant, 5) = "R" + x Then
      tIDRestaurant = Right(Rs!idrestaurant, 3) + 1
       tIDRestaurant = "R" + Format(Date, "yymm") + Left("000", 3 - Len(tIDRestaurant)) + tIDRestaurant
        Else

         tIDRestaurant = "R" + Format(Date, "yymm") + "001"
        End If
    End If
End With

End Sub


Private Sub btnAdd_Click()

If cmbIDGuest = "" Then
    MsgBox "Please Enter ID Guest,", vbExclamation, "mYHoTEL"
cmbIDGuest.SetFocus
Else
If cmbIDFood = "" Then
    MsgBox "Please Enter ID Food,", vbExclamation, "mYHoTEL"
cmbIDFood.SetFocus
Else
    
 
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from food where idfood='" & cmbIDFood & "'", KOneKsi, 3, 3
  If Not Rs.EOF Then
    While Not Rs.EOF
     
        Set Lv = lvItem.ListItems.Add(, , cmbIDFood)
            Lv.SubItems(1) = tNameFood
            Lv.SubItems(2) = tKinds
            Lv.SubItems(3) = Val(tPrice)
            Rs.MoveNext
tTotal = Lv.ListSubItems(3) + Val(tTotal)
    Wend
 KOneKsi.Execute "iNSERT INTO restaurant(idguest,idrestaurant,idfood,datetrans,timetrans,totalprice)values('" & cmbIDGuest & "','" & tIDRestaurant & "','" & cmbIDFood & "','" & tdate & "','" & lTime & "' ," & Lv.ListSubItems(3) & ")"
    
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from checkin", KOneKsi, 3, 3
    
    If Not Rs.EOF Then
   
             KOneKsi.Execute "Update Checkin set [restaurant]=" & Val(tTotal) + Val(CEKNULL((Rs!Restaurant))) & " where idguest ='" & cmbIDGuest & "'"
   money = Val(CEKNULL(Rs!service)) - Val(CEKNULL(Rs!laundry)) - Val(CEKNULL(Rs!Restaurant))
            KOneKsi.Execute "Update Checkin set [return money]=" & money & " where idguest ='" & cmbIDGuest & "'"
    End If
    
    
End If
End If

   
End If

 

End Sub

Private Sub btnCLOSE_Click()

Unload Me
splashHidup
End Sub

Private Sub btnNew_Click()
bersih Me
kunci Me
auto
isi
cmbIDGuest.Locked = False
cmbIDGuest.SetFocus
cmbIDFood.Locked = False
lvItem.ListItems.Clear
End Sub















Private Sub cmbIDFood_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Food where idfood='" & cmbIDFood & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            tNameFood = Rs!Name
            tKinds = Rs!Kinds
            tPrice = Rs!Price
        End If

End Sub

Private Sub cmbIDguEST_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from gUEST where idGUEST='" & cmbIDGuest & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            tname = Rs!Name
            tAddr = Rs!address
            tPhone = Rs!phone
            
        End If
End Sub



















Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then SendKeys "+ (tab)", True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub
Private Sub Timer1_Timer()
lTime = Time
tdate = Date
End Sub




Private Sub Form_Load()
OPENDATA
about.Movie = App.Path & "\Document\RESTAUTANT.swf"
about.Play
splashMati
Me.Height = 0
Me.Width = 0
Me.Top = 10
End Sub

Private Sub Timer2_Timer()
Me.Left = Me.Left + 100

If Me.Left >= 3000 Then

Timer2.Enabled = False
Timer3.Enabled = True
End If

End Sub

Private Sub Timer3_Timer()
Me.Width = Me.Width + 100
Me.Height = Me.Height + 100
If Me.Width >= 9375 And Me.Height >= 7365 Then Timer3.Enabled = False
End Sub






