VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Check_OUT 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Check_OUT"
   ClientHeight    =   10425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13890
   ForeColor       =   &H8000000C&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "ChecK_OuT.frx":0000
   ScaleHeight     =   10425
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1680
      TabIndex        =   100
      Top             =   240
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.TextBox tGrand 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   9840
      Width           =   2055
   End
   Begin VB.TextBox tservice2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox tIDcheckout 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   2760
      TabIndex        =   86
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox tTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   83
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox tIN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox tRest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox tReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox tLAundry 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox cmbRoomno 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   2400
      TabIndex        =   70
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox cmbIDGuest 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   7320
      TabIndex        =   69
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ListBox LOccupied 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2310
      ItemData        =   "ChecK_OuT.frx":2B01
      Left            =   9840
      List            =   "ChecK_OuT.frx":2B03
      TabIndex        =   68
      Top             =   2760
      Width           =   615
   End
   Begin VB.ListBox LVacant 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2310
      ItemData        =   "ChecK_OuT.frx":2B05
      Left            =   10680
      List            =   "ChecK_OuT.frx":2B07
      TabIndex        =   67
      Top             =   2760
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   16440
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   16200
      Top             =   1440
   End
   Begin VB.TextBox ttype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox tharga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   17
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox tpersen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox tdIscount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox tNett 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox tday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox tNett2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox tService 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox tAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox tTAx 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
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
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox tTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox tNational 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox tIDCARD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox tPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox tAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox tNAme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox tcity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox tIDcheckIN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin LvButtons.lvButtons_H cmdStatus 
      Height          =   735
      Left            =   10200
      TabIndex        =   0
      ToolTipText     =   "STATUS"
      Top             =   5280
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1296
      Caption         =   "&STATUS"
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
      cBhover         =   -2147483633
      LockHover       =   3
      cGradient       =   0
      Gradient        =   1
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "ChecK_OuT.frx":2B09
      ImgSize         =   48
      cBack           =   -2147483635
   End
   Begin MSComCtl2.DTPicker tglout 
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   6600
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
      Format          =   52232193
      CurrentDate     =   39484
   End
   Begin LvButtons.lvButtons_H btnSave 
      Height          =   1095
      Left            =   11880
      TabIndex        =   93
      ToolTipText     =   "Check_Out"
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "&SAVE"
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
      Image           =   "ChecK_OuT.frx":3755
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnPrint 
      Height          =   1095
      Left            =   11880
      TabIndex        =   94
      ToolTipText     =   "PRINT"
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
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
      Image           =   "ChecK_OuT.frx":3FC9
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnExit 
      Height          =   1095
      Left            =   11880
      TabIndex        =   95
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
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
      Image           =   "ChecK_OuT.frx":520D
      ImgSize         =   48
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "ChecK_OuT.frx":6EC1
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   1095
      Left            =   1680
      TabIndex        =   99
      Top             =   240
      Width           =   9015
      _cx             =   4210205
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
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   97
      Top             =   9840
      Width           =   45
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GRAND TOTAL"
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
      Left            =   480
      TabIndex        =   96
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Line Line29 
      BorderColor     =   &H00FFFF00&
      X1              =   11760
      X2              =   11760
      Y1              =   2280
      Y2              =   6720
   End
   Begin VB.Line Line32 
      BorderColor     =   &H00FFFF00&
      X1              =   13080
      X2              =   13080
      Y1              =   2280
      Y2              =   6720
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00FFFF00&
      X1              =   0
      X2              =   0
      Y1              =   480
      Y2              =   4920
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00FFFF00&
      X1              =   11760
      X2              =   13080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FFFF80&
      X1              =   11760
      X2              =   13080
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   11880
      Picture         =   "ChecK_OuT.frx":71DB
      Top             =   -240
      Width           =   12000
   End
   Begin VB.Label LmiN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11040
      TabIndex        =   92
      Top             =   9240
      Width           =   180
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   8760
      X2              =   11040
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   91
      Top             =   8280
      Width           =   45
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERVICE"
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
      Left            =   480
      TabIndex        =   90
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label Label63 
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
      Left            =   2400
      TabIndex        =   88
      Top             =   2040
      Width           =   60
   End
   Begin VB.Label Label62 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDCheck_Out"
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
      TabIndex        =   87
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Check_IN"
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
      TabIndex        =   85
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Label Label60 
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
      Left            =   2400
      TabIndex        =   84
      Top             =   3840
      Width           =   60
   End
   Begin VB.Label Label59 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Check_IN"
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
      TabIndex        =   82
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label Label58 
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
      Left            =   2400
      TabIndex        =   81
      Top             =   3360
      Width           =   60
   End
   Begin VB.Line Line35 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   16200
      X2              =   16200
      Y1              =   8520
      Y2              =   10080
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LAUNDRY"
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
      Left            =   480
      TabIndex        =   79
      Top             =   8640
      Width           =   810
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   78
      Top             =   8640
      Width           =   45
   End
   Begin VB.Label lmoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETURN MONEY"
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
      Left            =   480
      TabIndex        =   77
      Top             =   9480
      Width           =   1350
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   76
      Top             =   9480
      Width           =   45
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RESTAURANT"
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
      Left            =   480
      TabIndex        =   75
      Top             =   9000
      Width           =   1155
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   74
      Top             =   9000
      Width           =   45
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      X1              =   360
      X2              =   360
      Y1              =   5040
      Y2              =   10200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      X1              =   8520
      X2              =   8520
      Y1              =   5040
      Y2              =   8160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFF00&
      X1              =   360
      X2              =   8520
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF00&
      X1              =   360
      X2              =   8520
      Y1              =   5040
      Y2              =   5040
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
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   66
      Top             =   10320
      Width           =   1200
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
      TabIndex        =   65
      Top             =   5160
      Width           =   60
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
      TabIndex        =   64
      Top             =   5160
      Width           =   60
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
      TabIndex        =   63
      Top             =   5160
      Width           =   1275
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
      TabIndex        =   62
      Top             =   5760
      Width           =   1665
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
      TabIndex        =   61
      Top             =   6120
      Width           =   60
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
      TabIndex        =   60
      Top             =   5760
      Width           =   2145
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
      TabIndex        =   59
      Top             =   5760
      Width           =   1695
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
      TabIndex        =   58
      Top             =   6120
      Width           =   1230
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
      TabIndex        =   57
      Top             =   6600
      Width           =   1500
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
      TabIndex        =   56
      Top             =   6600
      Width           =   60
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
      TabIndex        =   55
      Top             =   6600
      Width           =   105
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
      TabIndex        =   54
      Top             =   6600
      Width           =   450
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
      TabIndex        =   53
      Top             =   6480
      Width           =   165
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
      TabIndex        =   52
      Top             =   7080
      Width           =   855
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
      TabIndex        =   51
      Top             =   7080
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
      TabIndex        =   50
      Top             =   7440
      Width           =   60
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
      TabIndex        =   49
      Top             =   7800
      Width           =   60
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8760
      X2              =   10800
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   330
      Left            =   10680
      TabIndex        =   48
      Top             =   7320
      Width           =   165
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
      Left            =   480
      TabIndex        =   47
      Top             =   5280
      Width           =   1005
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
      TabIndex        =   46
      Top             =   7440
      Width           =   450
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
      TabIndex        =   45
      Top             =   7800
      Width           =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      X1              =   11400
      X2              =   11400
      Y1              =   6480
      Y2              =   10200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFF00&
      X1              =   360
      X2              =   11400
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFF00&
      X1              =   11400
      X2              =   8520
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFF00&
      X1              =   11400
      X2              =   360
      Y1              =   10200
      Y2              =   10200
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFF00&
      X1              =   480
      X2              =   480
      Y1              =   2760
      Y2              =   4680
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
      Left            =   6480
      TabIndex        =   44
      Top             =   2880
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
      Left            =   5280
      TabIndex        =   43
      Top             =   3000
      Width           =   630
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
      Left            =   6480
      TabIndex        =   42
      Top             =   4560
      Width           =   60
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
      Left            =   6480
      TabIndex        =   41
      Top             =   3360
      Width           =   60
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      TabIndex        =   40
      Top             =   4560
      Width           =   615
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
      Left            =   5280
      TabIndex        =   39
      Top             =   3480
      Width           =   960
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
      Left            =   6480
      TabIndex        =   38
      Top             =   4200
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
      Left            =   6480
      TabIndex        =   37
      Top             =   3840
      Width           =   60
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
      Left            =   2400
      TabIndex        =   36
      Top             =   4200
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
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
      TabIndex        =   35
      Top             =   4200
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDCard"
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
      TabIndex        =   34
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      TabIndex        =   33
      Top             =   3840
      Width           =   360
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFF00&
      X1              =   480
      X2              =   4920
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFF00&
      X1              =   480
      X2              =   4920
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFF00&
      X1              =   4920
      X2              =   4920
      Y1              =   2760
      Y2              =   4680
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFF00&
      X1              =   5160
      X2              =   5160
      Y1              =   2760
      Y2              =   4920
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FFFF00&
      X1              =   5160
      X2              =   9360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FFFF00&
      X1              =   9360
      X2              =   9360
      Y1              =   2760
      Y2              =   4920
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFF00&
      X1              =   9360
      X2              =   5160
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFF00&
      X1              =   480
      X2              =   480
      Y1              =   1920
      Y2              =   2520
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
      Left            =   5640
      TabIndex        =   32
      Top             =   2040
      Width           =   915
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
      Left            =   7080
      TabIndex        =   31
      Top             =   2040
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDCheck_In"
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
      TabIndex        =   30
      Top             =   3000
      Width           =   1065
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
      Left            =   2400
      TabIndex        =   29
      Top             =   3000
      Width           =   60
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFF00&
      X1              =   9240
      X2              =   480
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FFFF00&
      X1              =   9240
      X2              =   480
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FFFF00&
      X1              =   9240
      X2              =   9240
      Y1              =   2520
      Y2              =   1920
   End
   Begin VB.Label tdate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   3000
      TabIndex        =   28
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label taRrival 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   7200
      TabIndex        =   27
      Top             =   1560
      Width           =   240
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FFFF00&
      X1              =   9600
      X2              =   9600
      Y1              =   2400
      Y2              =   6120
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupied"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   26
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "Vacant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   25
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FFFF00&
      X1              =   9600
      X2              =   11520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FFFF00&
      X1              =   11520
      X2              =   11520
      Y1              =   2400
      Y2              =   6120
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FFFF00&
      X1              =   9600
      X2              =   11520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00FFFF00&
      X1              =   9600
      X2              =   11520
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room status"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9360
      TabIndex        =   24
      Top             =   1920
      Width           =   2235
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Time"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5160
      TabIndex        =   23
      Top             =   1560
      Width           =   1710
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of arrival"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   480
      TabIndex        =   22
      Top             =   1560
      Width           =   1965
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   2760
      TabIndex        =   21
      Top             =   1440
      Width           =   150
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   6960
      TabIndex        =   20
      Top             =   1440
      Width           =   150
   End
End
Attribute VB_Name = "Check_OUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Double
Sub auto()
Dim x As String
    x = Format(Date, "yymm")

With Rs
    If .State = 1 Then .Close
        .Open "select * from  checkOut order by idcheckOut asc", KOneKsi, 3, 3
            If .EOF Then
                tIDcheckout = "OT" + x + "001"
            Else
                .MoveLast
                    If Left(Rs!idcheckOuT, 6) = "OT" + x Then
                        x = Right(Rs!idcheckOuT, 3) + 1
                        tIDcheckout = "OT" + Format(Date, "yymm") + Left("000", 3 - Len(x)) + x
                    Else
                        tIDcheckout = "OT" + x + "001"
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
    .Open "select * from checkOut where idCheckout='" & tIDcheckout & "'  ", KOneKsi, 3, 3
        If .EOF Then
            KOneKsi.Execute "Insert Into checkOut (GrandTotal,[return money],restaurant,laundry,idcheckout,days,idCheckIN,IdGuest,Idcard,Name,Address,City,Nationality,phone,roomno,Typeroom,Arrival_Date,Arrival_Time,Out_Date,Price,Discount,service,tax,Amount)values( '" & tGrand & "','" & tReturn & "','" & tRest & "','" & tlaundry & "','" & tIDcheckout & "','" & tday & "','" & tIDcheckIN & "' ,'" & cmbIDGuest & "','" & tIDCARD & "','" & tname & "','" & tAddr & "','" & tcity & "', '" & tNational & "'," & tPhone & ",'" & cmbRoomno & "','" & ttype & "', '" & tdate & " ','" & taRrival & "','" & tglout & "','" & tharga & "','" & tdIscount & "','" & tService & "','" & tTAx & "','" & tTotal & "')"
            KOneKsi.Execute "delete from checkIn where idcheckin='" & tIDcheckIN & "'"
            KOneKsi.Execute "Update Guest set CheckIN=false where IDGuest= '" & cmbIDGuest & "'"
            KOneKsi.Execute "update room set status=false where roomno=" & cmbRoomno & ""
           MsgBox ("Visitor") + " " + tname + " Is Sucessfully Check_Out", vbInformation, "mYHoTEL"
          Else
            MsgBox "Sorry,You Have IDCheckIN" + " " + tIDcheckIN, vbExclamation, "mYHoTEL"
         End If
 
 End If
End If
    End If
    End If
    
End With
End Sub

Private Sub btnEXIT_Click()
keluar Me
splashMati
End Sub



Private Sub btnPrint_Click()
On Error Resume Next
BiLL_CheckOUT.Show
Unload Me
BiLL_CheckOUT.PrintForm


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

Dim rs2 As New ADODB.Recordset
With Rs
    If .State = 1 Then .Close
        .Open "select * from room where roomno=" & cmbRoomno & "", KOneKsi, 3, 3
            If Not Rs.EOF Then
                    ttype = CEKNULL(Rs!Type_Room)
                tharga = CEKNULL(Rs!amount)
                Call tday_Change
            End If
                         If .State = 1 Then .Close
                            .Open "select * from reservation where roomno=" & cmbRoomno & "", KOneKsi, 3, 3
                                If Not Rs.EOF Then
                                    tglout.Enabled = False
                                    tglout = Rs!onDate
                                Else
                                    tglout.Enabled = False
                                End If
                                        
                                            If .State = 1 Then .Close
                                                .Open "select * from reservation where  ondate= #" & tdate & "#", KOneKsi, 3, 3
                                                If Not .EOF Then
                                                    If tdate = Rs!onDate Then
                                                            rs2.Open "select * from reservation where  roomno= " & cmbRoomno & " and idcard <>'" & tIDCARD & "'", KOneKsi, 3, 3
                                                             If Not rs2.EOF Then
                                                                 MsgBox "Room is Have Reservation,Please Enter Information Guest ", vbInformation, "myHoTEL"
                                                              Else
                                                                 tglout.Enabled = False
                                                           End If

                                                    End If
                                                End If
                        
  End With
             
Call tglout_Change

End Sub



Private Sub cmdStatus_Click()
kunci Me
If Rs.State = 1 Then Rs.Close
    LVacant.Clear
    LOccupied.Clear
    Rs.Open "Select * from Room where status=false", KOneKsi, 3, 3
        If Not Rs.EOF Then
            While Not Rs.EOF
               LVacant.AddItem Rs!Roomno
               Rs.MoveNext
            Wend
        End If
        
If Rs.State = 1 Then Rs.Close
    LOccupied.Clear
    Rs.Open "Select * from Room where status=true", KOneKsi, 3, 3
        If Not Rs.EOF Then
             While Not Rs.EOF
               LOccupied.AddItem Rs!Roomno
               Rs.MoveNext
              Wend
        Else
           MsgBox "Sorry,Is Not Have CheckIn", vbExclamation, "mYHoTEL"
            keluar Me
        End If
        

bersih Me
auto
End Sub








Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
about.Movie = App.Path & "\Document\CheckOUt.swf"
about.Play
awal Me
splashMati
Me.Width = 14010
Me.Height = 10665

kunci Me
tglout = Now()
OPENDATA
End Sub

















Private Sub lmoney_Change()
Dim tot As Variant
Dim ret As Variant
If lmoney.Caption = "RETURN MONEY" Then
    
tot = Val(tTotal)
ret = Val(tReturn)
    tGrand = tot - ret
Else
tot = Val(tTotal)
ret = Val(tReturn)
    tGrand = tot + ret
    
     'tGrand = Val(tTotal) - Val(-tReturn)
End If
End Sub


Private Sub LOccupied_Click()
If Rs.State = 1 Then Rs.Close
    Rs.Open "Select * from checkin where roomno= " & LOccupied & "", KOneKsi, 3, 3
        If Not Rs.EOF Then
            tIDcheckIN = CEKNULL(Rs!idcheckin)
            cmbIDGuest = CEKNULL(Rs!idguest)
            tIDCARD = CEKNULL(Rs!idcard)
            tIN = CEKNULL(Rs!arrival_date)
            tname = CEKNULL(Rs!Name)
            tAddr = CEKNULL(Rs!address)
            tcity = CEKNULL(Rs!City)
            tPhone = CEKNULL(Rs!phone)
            tTime = CEKNULL(Rs!Arrival_Time)
            tglout = CEKNULL(Rs!Out_Date)
            cmbRoomno = CEKNULL(Rs!Roomno)
            tService = CEKNULL(Rs!service)
            tservice2 = CEKNULL(Rs!service)
            ttype = CEKNULL(Rs!typeroom)
            tharga = CEKNULL(Rs!Price)
            tlaundry = CEKNULL(Rs!laundry)
            tRest = CEKNULL(Rs!Restaurant)
            tReturn = Val(CEKNULL(Rs!service)) - Val(CEKNULL(Rs!laundry)) - Val(CEKNULL(Rs!Restaurant))
            
            tday = Rs!days
             tNational = Rs!NATIONALITY
             tglout.Enabled = True
             
            
             
             
          
            kunci Me
            
            kunci Me
            tglout.Enabled = False
        End If
Call tNett2_Change

End Sub





Private Sub LVacant_Click()
bersih Me
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


    
End If


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





Private Sub tReturn_Change()

If Val(tReturn) >= 0 Then
    lmoney.Caption = "RETURN MONEY"
    LmiN.Caption = "-"
    'MsgBox tTotal + " - " + tReturn
    

Else
   i = Val(tReturn) + (-Val(tReturn) * 2)
    tReturn = i
    lmoney.Caption = "PAY BILL"
    LmiN.Caption = "+"
    'tGrand = Val(tTotal) - Val(tReturn)
End If
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
TERBILANG.Caption = angka(CEKNULL(Val(tGrand))) & " Rupiah"
Call lmoney_Change
End Sub


