VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Laundry 
   Caption         =   "LAUNDRY"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmLAundry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "frmLAundry.frx":038A
   MousePointer    =   99  'Custom
   PaletteMode     =   2  'Custom
   Picture         =   "frmLAundry.frx":0694
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   1920
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin LvButtons.lvButtons_H btnADD 
      Height          =   1455
      Left            =   8400
      TabIndex        =   29
      ToolTipText     =   "ADD"
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2566
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
      Image           =   "frmLAundry.frx":3195
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1560
      Top             =   720
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1080
      Top             =   600
   End
   Begin VB.TextBox tQTY 
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
      Left            =   7200
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   3240
      Width           =   3135
      Begin VB.CheckBox ChkDry 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "cLAUNDRy"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   24
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox ChkLaun 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "cLAUNDRy"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   22
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DRYCLEAN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2040
         TabIndex        =   25
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LAUNDRY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   480
         TabIndex        =   23
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   720
   End
   Begin VB.TextBox tPrice 
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
      Left            =   1800
      TabIndex        =   18
      Top             =   5520
      Width           =   1815
   End
   Begin VB.ComboBox cmbIDitem 
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
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox tNameLaun 
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
      TabIndex        =   15
      Top             =   3360
      Width           =   1815
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
      Left            =   5520
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
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
      Left            =   5520
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
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
      Left            =   5520
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
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
      Left            =   1800
      TabIndex        =   0
      Top             =   2160
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox tIDLaundry 
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvItem 
      Height          =   1335
      Left            =   240
      TabIndex        =   14
      Top             =   3840
      Width           =   7935
      _ExtentX        =   13996
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Laundry"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dry Clean"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin LvButtons.lvButtons_H btnclose 
      Height          =   735
      Left            =   960
      TabIndex        =   27
      Top             =   6480
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
      Image           =   "frmLAundry.frx":3B8F
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnNew 
      Height          =   735
      Left            =   240
      TabIndex        =   28
      Top             =   6480
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
      Image           =   "frmLAundry.frx":5835
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   975
      Left            =   2160
      TabIndex        =   30
      Top             =   120
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
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Left            =   7320
      TabIndex        =   26
      Top             =   3120
      Width           =   720
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
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   675
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      Top             =   3000
      Width           =   9135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1575
      Left            =   3480
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      Top             =   1320
      Width           =   3255
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
      Left            =   360
      TabIndex        =   19
      Top             =   5520
      Width           =   960
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Item"
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
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Laundry"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   3120
      Width           =   1200
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
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   795
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
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Laundry"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   930
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
      Left            =   4560
      TabIndex        =   10
      Top             =   2160
      Width           =   510
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
      Left            =   4560
      TabIndex        =   9
      Top             =   1800
      Width           =   675
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
      Left            =   4560
      TabIndex        =   8
      Top             =   1440
      Width           =   465
   End
End
Attribute VB_Name = "Laundry"
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


cmbIDItem.Clear
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from itemlaundry", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbIDItem.AddItem Rs!idItem
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
Dim X As String
    X = Format(Date, "yymm")
With Rs

If .State = 1 Then .Close
.Open "select * from LAUNDRY ORDER BY IDLaundry ASC ", KOneKsi, 3, 3

If Rs.EOF Then


        tIDLaundry = "L" + Format(Date, "yymm") + "001"
    Else
        .MoveLast

        If Left(Rs!IDLaundry, 5) = "L" + X Then
      tIDLaundry = Right(Rs!IDLaundry, 3) + 1
       tIDLaundry = "L" + Format(Date, "yymm") + Left("000", 3 - Len(tIDLaundry)) + tIDLaundry
        Else

         tIDLaundry = "L" + Format(Date, "yymm") + "001"
        End If
    End If
End With

End Sub


Private Sub btnAdd_Click()
If tQTY = "" Or tQTY = "0" Then
    MsgBox "Please Enter Quantity,", vbExclamation, "mYHoTEL"
tQTY.SetFocus
Else
If cmbIDGuest = "" Then
    MsgBox "Please Enter ID Guest,", vbExclamation, "mYHoTEL"
cmbIDGuest.SetFocus
Else
If cmbIDItem = "" Then
    MsgBox "Please Enter ID Item,", vbExclamation, "mYHoTEL"
cmbIDItem.SetFocus
Else
    If ChkDry.Value = 0 And ChkLaun.Value = 0 Then
        MsgBox "Please Enter Laundry Or Dry Clean,", vbExclamation, "mYHoTEL"
Else
Dim jml As Integer
 
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from itemLAundry where iditem='" & cmbIDItem & "'", KOneKsi, 3, 3
  If Not Rs.EOF Then
    While Not Rs.EOF
     
        Set Lv = lvItem.ListItems.Add(, , tNameLaun)
            Lv.SubItems(1) = tQTY
            Lv.SubItems(2) = laundry
            Lv.SubItems(3) = dRY
            Lv.SubItems(4) = Val(tQTY) * (dRY + laundry)
            Rs.MoveNext
tPrice = Lv.ListSubItems(4) + Val(tPrice)
    Wend
 KOneKsi.Execute "iNSERT INTO LAUNDRY(idguest,idlaundry,iditem,datetrans,timetrans,totalqty,price)values('" & cmbIDGuest & "','" & tIDLaundry & "','" & cmbIDItem & "','" & tdate & "','" & lTime & "' ," & tQTY & "," & Lv.ListSubItems(4) & ")"
    
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from checkin", KOneKsi, 3, 3
    
    If Not Rs.EOF Then
   
             KOneKsi.Execute "Update Checkin set [Laundry]=" & Val(tPrice) + Val(CEKNULL((Rs!laundry))) & " where idguest ='" & cmbIDGuest & "'"
   money = Val(CEKNULL(Rs!service)) - Val(CEKNULL(Rs!laundry)) - Val(CEKNULL(Rs!Restaurant))
                   
        
            KOneKsi.Execute "Update Checkin set [return money]=" & money & " where idguest ='" & cmbIDGuest & "'"
    End If
    
    
End If
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
lvItem.ListItems.Clear
cmbIDGuest.Locked = False
cmbIDGuest.SetFocus
cmbIDItem.Locked = False
tQTY.Locked = False
End Sub







Private Sub ChkDry_Click()

If Rs.State = 1 Then Rs.Close
Rs.Open "select * from itemLAundry where iditem='" & cmbIDItem & "'", KOneKsi, 3, 3
If Not Rs.EOF Then
    If ChkDry.Value = Checked Then dRY = Rs!dRYClean
    If ChkDry.Value = Unchecked Then dRY = 0
End If
End Sub

Private Sub ChkLaun_Click()
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from itemLAundry where iditem='" & cmbIDItem & "'", KOneKsi, 3, 3
If Not Rs.EOF Then
    If ChkLaun.Value = Checked Then laundry = Rs!laundry
    If ChkLaun.Value = Unchecked Then laundry = 0
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

Private Sub cmbIDItem_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from itemlaundry where iditem='" & cmbIDItem & "'", KOneKsi, 3, 3
        If Not Rs.EOF Then
            tNameLaun = Rs!Name
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
about.Movie = App.Path & "\Document\LAundry.swf"
about.Play
splashMati
OPENDATA
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
If Me.Height >= 7515 And Me.Width >= 9375 Then Timer3.Enabled = False
End Sub



Private Sub tQTY_Change()
If Not IsNumeric(tQTY) Then tQTY = 0
End Sub


