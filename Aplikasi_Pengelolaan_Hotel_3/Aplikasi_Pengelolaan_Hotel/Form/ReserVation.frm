VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ReserVation 
   BackColor       =   &H80000007&
   Caption         =   "Form2"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "ReserVation.frx":0000
   ScaleHeight     =   8265
   ScaleWidth      =   11940
   Begin VB.TextBox ttYpe 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   39
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin LvButtons.lvButtons_H cmdResLIST 
      Height          =   735
      Left            =   6480
      TabIndex        =   34
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      Caption         =   "Reserved  &list"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "ReserVation.frx":2B01
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H cmdUpdate 
      Height          =   480
      Left            =   10080
      TabIndex        =   33
      Top             =   3600
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   847
      Caption         =   "&Confirm"
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
      Image           =   "ReserVation.frx":3C2E
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnEXIT 
      Height          =   450
      Left            =   9720
      TabIndex        =   32
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3175
      _ExtentY        =   794
      Caption         =   "&EXIT"
      CapAlign        =   2
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
      Image           =   "ReserVation.frx":464F
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H cmdADD 
      Height          =   480
      Left            =   9960
      TabIndex        =   31
      Top             =   2760
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   847
      Caption         =   "&SUBMIT"
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
      Image           =   "ReserVation.frx":4E06
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H cmdClear 
      Height          =   465
      Left            =   9960
      TabIndex        =   0
      Top             =   1920
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   820
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
      cBhover         =   0
      cGradient       =   0
      Gradient        =   3
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "ReserVation.frx":59CA
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.ListBox LdaTE 
      BackColor       =   &H80000006&
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
      Height          =   2220
      ItemData        =   "ReserVation.frx":6526
      Left            =   7080
      List            =   "ReserVation.frx":6528
      TabIndex        =   27
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ListBox Lid 
      BackColor       =   &H80000006&
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
      Height          =   2220
      ItemData        =   "ReserVation.frx":652A
      Left            =   5280
      List            =   "ReserVation.frx":652C
      TabIndex        =   26
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox cmbRoomno 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2160
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tIDReservation 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox chkConfirm 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   6120
      Width           =   255
   End
   Begin VB.TextBox tPHone 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox tADDR 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox tName 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox tID 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9120
      Top             =   960
   End
   Begin MSComCtl2.DTPicker tgl 
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483642
      CalendarForeColor=   65280
      CalendarTitleBackColor=   0
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   0
      Format          =   54788097
      CurrentDate     =   39484
   End
   Begin LvButtons.lvButtons_H cmdResConfirmed 
      Height          =   735
      Left            =   6120
      TabIndex        =   35
      Top             =   6360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      Caption         =   "Reserved &CONFIRMED"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "ReserVation.frx":652E
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H cmdREsDeL 
      Height          =   735
      Left            =   6120
      TabIndex        =   36
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      Caption         =   "&DElete expired reserved"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "ReserVation.frx":6F5E
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   975
      Left            =   3480
      TabIndex        =   40
      Top             =   240
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
   Begin VB.Label lType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "type room"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   720
      TabIndex        =   38
      Top             =   7320
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label titik2 
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
      TabIndex        =   37
      Top             =   7080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line29 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9600
      X2              =   11760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9600
      X2              =   11760
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9600
      X2              =   11760
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9600
      X2              =   9600
      Y1              =   1800
      Y2              =   5280
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   11760
      X2              =   11760
      Y1              =   1800
      Y2              =   5280
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9600
      X2              =   11760
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9600
      X2              =   11760
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9000
      X2              =   5760
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9000
      X2              =   5760
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9000
      X2              =   9000
      Y1              =   5280
      Y2              =   8160
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   5760
      X2              =   5760
      Y1              =   5280
      Y2              =   8160
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9480
      X2              =   5160
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   5160
      X2              =   9480
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   6960
      X2              =   6960
      Y1              =   2400
      Y2              =   5280
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   5160
      X2              =   9480
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9480
      X2              =   9480
      Y1              =   1800
      Y2              =   8160
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reserved date"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7440
      TabIndex        =   30
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDRESERVATIOn"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   29
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation  List"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   6240
      TabIndex        =   28
      Top             =   1920
      Width           =   2325
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   5160
      X2              =   9480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   5160
      X2              =   5160
      Y1              =   1800
      Y2              =   8160
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   4920
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   4920
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   4920
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4920
      X2              =   4920
      Y1              =   5040
      Y2              =   7800
   End
   Begin VB.Label Titik 
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
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lroom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Roomno"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   720
      TabIndex        =   24
      Top             =   6720
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmed Arrival"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   960
      TabIndex        =   22
      Top             =   6120
      Width           =   2145
   End
   Begin VB.Label Label8 
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
      Left            =   2760
      TabIndex        =   21
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label7 
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
      Left            =   2760
      TabIndex        =   20
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label6 
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
      Left            =   2760
      TabIndex        =   19
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label5 
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
      Left            =   2760
      TabIndex        =   18
      Top             =   2520
      Width           =   135
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
      Left            =   2760
      TabIndex        =   17
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   2760
      TabIndex        =   16
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDreservation"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label TDATE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3120
      TabIndex        =   13
      Top             =   1320
      Width           =   180
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   4920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   4920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4920
      X2              =   4920
      Y1              =   1800
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   600
      Y1              =   5040
      Y2              =   7800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   4920
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   960
      TabIndex        =   11
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated  Arrival"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   720
      TabIndex        =   10
      Top             =   5160
      Width           =   2610
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   9
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   8
      Top             =   3840
      Width           =   885
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   7
      Top             =   3240
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDCARD"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   6
      Top             =   2760
      Width           =   795
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   600
      Y1              =   1800
      Y2              =   4800
   End
End
Attribute VB_Name = "ReserVation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTemp As New ADODB.Recordset
Option Explicit
Sub auto()
Dim x As String
    x = Format(Date, "yymm")

With Rs
    If .State = 1 Then .Close
        .Open "select * from  rESERVATION order by idRESERVATION asc", KOneKsi, 3, 3
            If .EOF Then
                tIDReservation = "S" + x + "001"
            Else
                .MoveLast
                    If Left(Rs!idreservation, 5) = "S" + x Then
                        x = Right(Rs!idreservation, 3) + 1
                        tIDReservation = "S" + Format(Date, "yymm") + Left("000", 3 - Len(x)) + x
                    Else
                        tIDReservation = "S" + x + "001"
                    End If
    End If
End With
End Sub

Private Sub btnEXIT_Click()
Unload Me
splashHidup
End Sub

Private Sub chkConfirm_Click()
If chkConfirm.Value Then
    cmbRoomno.Visible = True
    Titik.Visible = True
    lroom.Visible = True
    cmbRoomno.Locked = False
    cmbRoomno.Clear
    ttYpe.Visible = True
    titik2.Visible = True
    lType.Visible = True
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Room where status=false", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbRoomno.AddItem Rs!Roomno
                Rs.MoveNext
            Wend
        End If
Else
If chkConfirm.Value = Unchecked Then
    cmbRoomno.Visible = False
    Titik.Visible = False
    lroom.Visible = False
End If
End If
End Sub





Private Sub cmbRoomno_Click()

With Rs
    If .State = 1 Then .Close
        .Open "select * from room where roomno=" & cmbRoomno & "", KOneKsi, 3, 3
            If Not Rs.EOF Then
                ttYpe = CEKNULL(Rs!Type_Room)
            End If
End With


                               

Call tgl_Change



End Sub

Private Sub cmdADD_Click()
Dim RsTemp As New ADODB.Recordset
If tID.Text = "" Then
MsgBox "Please Enter ID Card", vbExclamation, "mYHoTEL"
tID.SetFocus
Else
If tName.Text = "" Then
MsgBox "Please Enter name", vbExclamation, "mYHoTEL"
tName.SetFocus
Else
    If tADDR.Text = "" Then
    MsgBox "Please Enter Address", vbExclamation, "mYHoTEL"
    tADDR.SetFocus
Else
    If tPHone.Text = "" Then
    MsgBox "Please Enter Phone Number", vbExclamation, "mYHoTEL"
    tPHone.SetFocus
Else
    

    If Rs.State = 1 Then Rs.Close
        Rs.Open "select * from reservation where idreservation='" & tIDReservation & "'", KOneKsi, 3, 3
            If Rs.EOF Then
                
'                If rstemp.State = 1 Then rstemp.Close
'                  rstemp.Open "select * from reservation where =#" & tgl & "# ", KOneKsi, 3, 3
'                        If rstemp.EOF Then
                        
                                KOneKsi.Execute "insert into reservation(Typeroom,idreservation,ondate,idcard,name,address,phone,arrivaldate,confirmed)values('" & ttYpe & "','" & Replace(tIDReservation, "'", "") & "','" & TDATE & "','" & Replace(tID, "'", "") & "','" & Replace(tName, "'", "") & "','" & Replace(tADDR, "'", "") & "'," & tPHone & ",'" & tgl & "'," & chkConfirm & ")"
                            
                            If cmbRoomno <> "" Then KOneKsi.Execute " Update reservation set roomno=" & cmbRoomno & " where idreservation='" & tIDReservation & "'"
                            
                            MsgBox ("Data added. Room alloted for Reservation") + " " + tName, vbInformation, "mYHoTEL"
'                        Else
'                            MsgBox "SORRY, Room is Have Reservation  !!!", vbExclamation, "mYHoTEL"""
'                        End If
                        
            Else
                 MsgBox "SORRY, You Have Reservation  !!!", vbExclamation, "mYHoTEL"""
            End If

End If
End If
End If
End If
End Sub

Private Sub cmdClear_Click()
Lid.Clear
LdaTE.Clear
Buka Me
bersih Me
auto
tgl.Enabled = True
chkConfirm.Enabled = True
cmdUpdate.Visible = False
cmdADD.Enabled = True
tID.SetFocus




End Sub

Private Sub cmdResConfirmed_Click()
LdaTE.Clear
Lid.Clear
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from reservation where confirmed= true ", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                Lid.AddItem Rs!idreservation
                LdaTE.AddItem Rs!arrivaldate
                Rs.MoveNext
            Wend
        End If
cmdREsDeL.Visible = False

End Sub

Private Sub cmdREsDeL_Click()
 pesan = MsgBox("Are You Sure Delete?" + " " + Lid, vbQuestion + vbYesNo, "DELETE")
                
                If pesan = vbYes Then

KOneKsi.Execute "delete * from reservation where idreservation='" & Lid & "'"
MsgBox Lid + " " + "Expired reservation deleted sucessfuly...", vbInformation, "mYHoTEL"
End If
End Sub

Private Sub cmdResLIST_Click()
Lid.Clear
LdaTE.Clear
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from reservation where confirmed= false ", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                Lid.AddItem Rs!idreservation
                LdaTE.AddItem Rs!onDate
                Rs.MoveNext
            Wend
        End If
        cmdREsDeL.Visible = False

End Sub




Private Sub cmdUpdate_Click()
If Rs.State = 1 Then Rs.Close
If chkConfirm = 1 Then
    Rs.Open "select * from reservation where idreservation='" & Lid & "' and confirmed = false", KOneKsi, 3, 3
        If Not Rs.EOF Then
            KOneKsi.Execute "update reservation set confirmed=" & chkConfirm & " , roomno=" & cmbRoomno & " , typeroom='" & ttYpe & "' where idreservation='" & tIDReservation & "'"
            MsgBox ("Confirmation. Room alloted for Reservation") + " " + tName, vbInformation, "mYHoTEL"
            cmdUpdate.Visible = False
            cmdADD.Enabled = True
            cmdClear.Enabled = True
        End If
Else
    MsgBox "Please Enter Confirmed ArrivaL", vbExclamation, "mYHoTEL"
End If
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Form_Load()
OPENDATA
about.Movie = App.Path & "\Document\new_Reservation.swf"
about.Play
splashMati
awal Me
kunci Me
bersih Me
tgl = Date
cmdUpdate.Visible = False

Me.Height = 8775

Me.Width = 12060
End Sub




Private Sub lid_Click()
kunci Me
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from reservation where idreservation='" & Lid & "' and confirmed = false ", KOneKsi, 3, 3
        If Not Rs.EOF Then
            
            tIDReservation = CEKNULL(Rs!idreservation)
            tID = CEKNULL(Rs!idcard)
            TDATE = CEKNULL(Rs!onDate)
            tName = CEKNULL(Rs!Name)
            tADDR = CEKNULL(Rs!address)
            tPHone = CEKNULL(Rs!phone)
            tgl = CEKNULL(Rs!arrivaldate)
            chkConfirm.Value = Rs!confirmed
            chkConfirm.Enabled = True
            cmdUpdate.Visible = True
            cmdADD.Enabled = False
            cmdREsDeL.Visible = True
        End If

    If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from reservation where idreservation='" & Lid & "' And confirmed = true ", KOneKsi, 3, 3
        If Not Rs.EOF Then
            cmbRoomno = CEKNULL(Rs!Roomno)
            ttYpe = CEKNULL(Rs!typeroom)
            tIDReservation = CEKNULL(Rs!idreservation)
            tID = CEKNULL(Rs!idcard)
            TDATE = CEKNULL(Rs!onDate)
            tName = CEKNULL(Rs!Name)
            tADDR = CEKNULL(Rs!address)
            tPHone = CEKNULL(Rs!phone)
            tgl = CEKNULL(Rs!arrivaldate)
            chkConfirm.Value = 1
            chkConfirm.Enabled = False
            cmdUpdate.Visible = False
            cmdADD.Enabled = False
            cmdREsDeL.Visible = True
        End If
        
    

        
        
End Sub


Private Sub tgl_Change()
With RsTemp
        If .State = 1 Then .Close
        If cmbRoomno <> "" Then
            .Open "Select * from reservation where roomno=" & cmbRoomno & " order by Arrivaldate asc", KOneKsi, 3, 3
                  
                If Not .EOF Then
                  
                        If tgl >= RsTemp!arrivaldate Then
                            MsgBox "Sorry Is Have Reservation", vbExclamation, "mYHoTEL"
                            tgl = DateValue(RsTemp!arrivaldate) - 1
                           
                       
                    End If
                    
'                    If tgl <= rstemp!ArrivalDate Then
'                        MsgBox "Sorry Is Have Reservation", vbExclamation, "mYHoTEL"
'                       tgl = DateValue(rstemp!ArrivalDate) - 1
'                    End If
                    
                       If tgl = RsTemp!arrivaldate Then
                            MsgBox "Sorry Is Have Reservation", vbExclamation, "mYHoTEL"
                            tgl = DateValue(RsTemp!arrivaldate) - 1
                           
                       
                    End If
                    
                End If
        End If
End With

        If tgl < TDATE Then tgl = Date



End Sub

Private Sub Timer1_Timer()
TDATE = Date
End Sub

Private Sub tPhone_Change()
If Not IsNumeric(tPHone) Then tPHone = 0
End Sub
