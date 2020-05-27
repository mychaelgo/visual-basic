VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CHECK_IN 
   BackColor       =   &H00400000&
   Caption         =   "CHECK_IN"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   11565
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Transaksi.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11565
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   2160
      TabIndex        =   71
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
   End
   Begin LvButtons.lvButtons_H btnSave 
      Height          =   975
      Left            =   9840
      TabIndex        =   5
      ToolTipText     =   "BACK_UP"
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&SAVE"
      CapAlign        =   2
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   0
      LockHover       =   3
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "Transaksi.frx":2B01
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnPrint 
      Height          =   975
      Left            =   9840
      TabIndex        =   65
      ToolTipText     =   "PRINT"
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Caption         =   "&PRINT"
      CapAlign        =   2
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "Transaksi.frx":3B9F
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnExit 
      Height          =   855
      Left            =   9960
      TabIndex        =   64
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      Caption         =   "&EXIT"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   0
      LockHover       =   3
      cGradient       =   0
      Gradient        =   1
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "Transaksi.frx":4DE3
      ImgSize         =   48
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "Transaksi.frx":6A97
   End
   Begin LvButtons.lvButtons_H btnNEW 
      Height          =   735
      Left            =   9720
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      Caption         =   "&NEW"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cBhover         =   0
      cGradient       =   0
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      CustomClick     =   1
      Image           =   "Transaksi.frx":6DB1
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin VB.TextBox tIDTransaction 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cmbIDGuest 
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
      ForeColor       =   &H0000FFFF&
      Height          =   405
      ItemData        =   "Transaksi.frx":77AB
      Left            =   7560
      List            =   "Transaksi.frx":77AD
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox tcity 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox tNAme 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox tAddr 
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
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox tPhone 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox tIDCARD 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox tNational 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox tTotal 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox tTAx 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   6
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox tAmount 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox tService 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox tNett2 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox tday 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox tNett 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox tdIscount 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox tpersen 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox tharga 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox ttype 
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4920
      Width           =   3255
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
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   2400
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   16080
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   16320
      Top             =   600
   End
   Begin VB.PictureBox btnCHECKIN 
      Height          =   1215
      Left            =   15000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker tglout 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   6360
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
      Format          =   19726337
      CurrentDate     =   39484
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   975
      Left            =   2640
      TabIndex        =   70
      Top             =   120
      Width           =   6735
      _cx             =   4206184
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
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   6720
      TabIndex        =   69
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   2640
      TabIndex        =   68
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of arrival"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   360
      TabIndex        =   67
      Top             =   1200
      Width           =   2190
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Time"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   4680
      TabIndex        =   66
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   9600
      X2              =   11160
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FFFF80&
      BorderWidth     =   3
      X1              =   9600
      X2              =   11160
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   11160
      X2              =   11160
      Y1              =   1680
      Y2              =   5760
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   9600
      X2              =   9600
      Y1              =   1680
      Y2              =   5760
   End
   Begin VB.Label taRrival 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   7080
      TabIndex        =   63
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label tdate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   3000
      TabIndex        =   62
      Top             =   1080
      Width           =   270
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   9240
      X2              =   9240
      Y1              =   2400
      Y2              =   1800
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   9240
      X2              =   360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   9240
      X2              =   360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2280
      TabIndex        =   61
      Top             =   1920
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDTrancation"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   60
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   7200
      TabIndex        =   59
      Top             =   1920
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDGUEST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5520
      TabIndex        =   58
      Top             =   1920
      Width           =   915
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   360
      X2              =   360
      Y1              =   1800
      Y2              =   2400
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   9000
      X2              =   5160
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   9000
      X2              =   9000
      Y1              =   2640
      Y2              =   4200
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   5160
      X2              =   9000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   5160
      X2              =   5160
      Y1              =   2640
      Y2              =   4200
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   5040
      X2              =   5040
      Y1              =   2640
      Y2              =   4200
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   360
      X2              =   5040
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   360
      X2              =   5040
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "city"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5280
      TabIndex        =   56
      Top             =   2760
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDcard"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   55
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nationality"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5280
      TabIndex        =   54
      Top             =   3240
      Width           =   1500
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   53
      Top             =   2880
      Width           =   60
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   6960
      TabIndex        =   52
      Top             =   2760
      Width           =   60
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   6960
      TabIndex        =   51
      Top             =   3240
      Width           =   60
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   50
      Top             =   3600
      Width           =   960
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "phone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5280
      TabIndex        =   49
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   48
      Top             =   3600
      Width           =   60
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   6960
      TabIndex        =   47
      Top             =   3720
      Width           =   60
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   46
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   45
      Top             =   3240
      Width           =   60
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   360
      X2              =   360
      Y1              =   2640
      Y2              =   4200
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   8400
      X2              =   360
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   8400
      X2              =   10800
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   10800
      X2              =   8400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   360
      X2              =   10800
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   10800
      X2              =   10800
      Y1              =   6120
      Y2              =   8880
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   38
      Top             =   7560
      Width           =   480
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   37
      Top             =   7200
      Width           =   450
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Roomno"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   36
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   10560
      TabIndex        =   35
      Top             =   7200
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   8640
      X2              =   10680
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   34
      Top             =   7680
      Width           =   60
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   33
      Top             =   7200
      Width           =   60
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   32
      Top             =   6840
      Width           =   60
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "service"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   31
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6120
      TabIndex        =   30
      Top             =   6240
      Width           =   165
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5640
      TabIndex        =   29
      Top             =   6360
      Width           =   450
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4320
      TabIndex        =   28
      Top             =   6360
      Width           =   105
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   27
      Top             =   6360
      Width           =   60
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "checkOUT_date"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   26
      Top             =   6360
      Width           =   1500
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "price rent"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   25
      Top             =   5880
      Width           =   1230
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "         NETT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6480
      TabIndex        =   24
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "        DISCOUNT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4200
      TabIndex        =   23
      Top             =   5520
      Width           =   2145
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   22
      Top             =   5880
      Width           =   60
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "          PRICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2280
      TabIndex        =   21
      Top             =   5520
      Width           =   1665
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3600
      TabIndex        =   20
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   2160
      TabIndex        =   19
      Top             =   4920
      Width           =   60
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   300
      Left            =   4920
      TabIndex        =   18
      Top             =   4920
      Width           =   60
   End
   Begin VB.Label TERBILANG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TERBILANG"
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
      TabIndex        =   17
      Top             =   8280
      Width           =   1200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   360
      X2              =   8400
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   360
      X2              =   8400
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   8400
      X2              =   8400
      Y1              =   4680
      Y2              =   8040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   5
      X1              =   360
      X2              =   360
      Y1              =   4680
      Y2              =   8880
   End
End
Attribute VB_Name = "CHECK_IN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub auto()
Dim x As String
    x = Format(Date, "yymm")

With Rs
    If .State = 1 Then .Close
        .Open "select * from  checkIN order by idcheckin asc", KOneKsi, 3, 3
            If .EOF Then
                tIDTransaction = "IN" + x + "001"
            Else
                .MoveLast
                    If Left(Rs!idcheckin, 6) = "IN" + x Then
                        x = Right(Rs!idcheckin, 3) + 1
                        tIDTransaction = "IN" + Format(Date, "yymm") + Left("000", 3 - Len(x)) + x
                    Else
                        tIDTransaction = "IN" + x + "001"
                    End If
    End If
End With
End Sub

Private Sub btnSave_Click()
With Rs
    
If cmbIDGuest = "" Then
MsgBox "Please Enter ID Guest,", vbExclamation, "mYHoTEL"
cmbIDGuest.SetFocus
Else
    If cmbRoomno = "" Then
        MsgBox "Please Enter Room No,", vbExclamation, "mYHoTEL"
        cmbRoomno.SetFocus
Else
    If tService = "" Then
        MsgBox "Please Enter  Service,", vbExclamation, "mYHoTEL"
        tService.SetFocus
Else
    If tTAx = "" Then
        MsgBox "Please Enter Tax,", vbExclamation, "mYHoTEL"
        tTAx.SetFocus
Else
    If .State = 1 Then .Close
    .Open "select * from checkIN where idguest='" & cmbIDGuest & "' ", KOneKsi, 3, 3
        If .EOF Then
        
           KOneKsi.Execute "Insert Into checkin (days,idCheckIN,IdGuest,Idcard,Name,Address,City,Nationality,phone,roomno,Typeroom,Arrival_Date,Arrival_Time,Out_Date,Price,Discount,service,tax,Amount)values('" & Replace(tday, "'", "") & "','" & Replace(tIDTransaction, "'", "") & "' ,'" & Replace(cmbIDGuest, "'", "") & "','" & Replace(tIDCARD, "'", "") & "','" & Replace(tname, "'", "") & "','" & Replace(tAddr, "'", "") & "','" & Replace(tcity, "'", "") & "', '" & Replace(tNational, "'", "") & "'," & tPhone & ",'" & Replace(cmbRoomno, "'", "") & "','" & Replace(ttype, "'", "") & "', '" & tdate & " ','" & taRrival & "','" & tglout & "','" & Replace(tharga, "'", "") & "','" & Replace(tdIscount, "'", "") & "','" & Replace(tService, "'", "") & "','" & Replace(tTAx, "'", "") & "','" & Replace(tTotal, "'", "") & "')"
            
            KOneKsi.Execute "Update Guest set checkIN=true where IdGuest='" & cmbIDGuest & "'"
            
            KOneKsi.Execute "update room set status=true where roomno=" & cmbRoomno & ""
           MsgBox ("Data added. Room alloted for visitor") + " " + tname, vbInformation, "mYHoTEL"
          Else
            MsgBox "Sorry,You Have IDCheckIN" + " " + Rs!idcheckin, vbExclamation, "mYHoTEL"
         End If
 
 
 
 
 
 
 
 
 
End If
End If
    End If
    End If
    
End With
End Sub

Private Sub btnEXIT_Click()
keluar Me
splashHidup
End Sub

Private Sub btnNew_Click()
Timer1.Enabled = True
cmbIDGuest.Locked = False
cmbRoomno.Locked = False
tglout.Enabled = True
tAmount.Locked = False
tService.Locked = False
bersih Me
tday = 1

cmbIDGuest.SetFocus
cmbIDGuest.Clear
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Guest where CheckIn= false", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbIDGuest.AddItem Rs!idguest
                Rs.MoveNext
            Wend
        Else
        MsgBox "Sorry,Is Not Have Guest", vbExclamation, "mYHoTEL"
        Unload Me
        End If
cmbRoomno.Clear
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Room where status=false", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
                cmbRoomno.AddItem Rs!Roomno
                Rs.MoveNext
            Wend
        End If
    auto
    

End Sub






Private Sub btnPrint_Click()
On Error Resume Next
BILL_CHECK_IN.Show
Unload Me
BILL_CHECK_IN.PrintForm


End Sub





Private Sub cmbIDguEST_Click()
With Rs
    If .State = 1 Then .Close
        .Open "select * from guest where idguest='" & cmbIDGuest & "'", KOneKsi, 3, 3
            If Not Rs.EOF Then
                tIDCARD = CEKNULL(Rs!idcard)
                tname = CEKNULL(Rs!Name)
                tAddr = CEKNULL(Rs!address)
                tcity = CEKNULL(Rs!City)
                tNational = CEKNULL(Rs!NATIONALITY)
                tPhone = CEKNULL(Rs!phone)
            End If
End With
End Sub


Private Sub cmbRoomno_Click()

With Rs
    If .State = 1 Then .Close
        .Open "select * from room where roomno=" & cmbRoomno & "", KOneKsi, 3, 3
            If Not Rs.EOF Then
                    ttype = CEKNULL(Rs!Type_Room)
                tharga = CEKNULL(Rs!amount)
                Call tday_Change
            End If
                          
                                        
                                        
                                            If .State = 1 Then .Close
                                                .Open "select * from reservation where  Arrivaldate= #" & tdate & "# and Roomno = " & cmbRoomno & " ", KOneKsi, 3, 3
                                                If Not .EOF Then
                                                    If tdate = Rs!arrivaldate Then
                                                          
                                                             If Not .EOF Then
                                                              
                                                                 MsgBox "Room is Have Reservation,Please Enter Information Guest ", vbInformation, "myHoTEL"
                                                                  bersih Me
                                                              Else
                                                                 tglout.Enabled = True
                                                           End If

                                                    End If
                                                End If
                                               
                        
  End With
  Call tglout_Change

          

End Sub










Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
about.Movie = App.Path & "\Document\Checkin.swf"
about.Play
awal Me
splashMati
Me.Width = 11685
Me.Height = 9510

kunci Me
OPENDATA
tglout = Date

End Sub

Private Sub tAmount_Change()
Dim x As Double
Dim Y As Double
Dim z As Double
Dim total As Double


x = Val(tAmount)
Y = Val(tService)
z = Val(tTAx)
total = x + Y + z
tTotal = total
tTAx = 10 * total \ 110

End Sub




Private Sub tday_Change()
Dim x As Double
Dim z As Double
Dim r As Double
Dim harga As Double
harga = Val(tharga)


x = Val(tday)

            If x >= 1 And x < 7 Then
                z = 0
            Else
                If x >= 7 And x <= 14 Then
                    z = 0.1
            Else
                If x >= 15 And x <= 21 Then
                    z = 0.2
            Else
                If x >= 22 And x <= 29 Then
                    z = 0.3
            Else
                If x >= 30 And x <= 50 Then
                    z = 0.4
            Else
                If x >= 51 Then
                    z = 0.5
            End If
            End If
            End If
            End If
            End If
            End If
            
r = z * harga
tpersen = z * 100 & " %"
tdIscount = r

End Sub





Private Sub tdIscount_Change()
Dim x As Double
Dim z As Double
Dim nett As Double

x = Val(tharga)
z = Val(tdIscount)
nett = x - z

tNett = nett
End Sub

Private Sub tglout_Change()
tpersen.Locked = True
Dim x As Date
Dim z As Date
Dim day As Double

x = CDate(tdate)
z = CDate(tglout)

If tglout = tdate Then tday = 1



If x <> z Then
    
    If x > z Then
        tday = 0
    Else
        day = DateValue(z) - DateValue(x)
        tday = day
    End If
Else
     tday = 1

    
End If
With Rs
If .State = 1 Then .Close
If cmbRoomno <> "" Then
                            .Open "select * from reservation where roomno=" & CEKNULL(cmbRoomno) & " order by Arrivaldate", KOneKsi, 3, 3
                                If Not Rs.EOF Then
                                        If tglout >= Rs!arrivaldate Then
                                            cmbRoomno.Text = ""
                                      
                                                MsgBox "Sorry Is Have Reservation", vbExclamation, "mYHoTEL"
                                                tglout = DateValue(Rs!arrivaldate) - 1
                                                    If tdate > tglout Then tglout = tdate
                                            End If
                                            
                                                               
                                                                        
                                Else
                                    tglout.Enabled = True
                                End If

End If
End With
Call tdIscount_Change
Call tNett2_Change
Call tday_Change

End Sub

Private Sub tharga_Change()
Call tdIscount_Change
End Sub

Private Sub Timer1_Timer()
tdate = Date
taRrival = Time()
Me.Caption = Right(Me.Caption, 1) + Left(Me.Caption, Len(Me.Caption) - 1)

End Sub





Private Sub tNett_Change()

 tNett2 = tNett
End Sub

Private Sub tNett2_Change()
Dim x As Double
Dim z As Double
Dim amount As Double

x = Val(tday)
z = Val(tNett2)
                amount = x * z
        tAmount = amount


 
End Sub



Private Sub tService_Change()
Call tAmount_Change
If Not IsNumeric(tService) Then tService = 0

End Sub

Private Sub tTAx_Change()
Call tAmount_Change
If Not IsNumeric(tTAx) Then tTAx = 0
End Sub

Private Sub tTotal_Change()
TERBILANG.Caption = angka(CEKNULL(Val(tTotal))) & " Rupiah"
End Sub

