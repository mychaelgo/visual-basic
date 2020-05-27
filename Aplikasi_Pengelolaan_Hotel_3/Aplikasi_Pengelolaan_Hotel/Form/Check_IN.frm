VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Begin VB.Form INFORMATION_GUEST 
   BackColor       =   &H00404040&
   Caption         =   "INFORMATION GUEST"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "Check_IN.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Check_IN.frx":030A
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   651
   Begin VB.ComboBox cmbIDcard 
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
      ForeColor       =   &H0000FF00&
      Height          =   405
      ItemData        =   "Check_IN.frx":2E0B
      Left            =   7920
      List            =   "Check_IN.frx":2E0D
      TabIndex        =   53
      Top             =   1680
      Width           =   1455
   End
   Begin LvButtons.lvButtons_H cmdReservation 
      Height          =   855
      Left            =   5640
      TabIndex        =   52
      ToolTipText     =   "Reservation List"
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      Caption         =   "&Reservation List"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   0
      cGradient       =   0
      Gradient        =   4
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "Check_IN.frx":2E0F
      ImgSize         =   48
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "Check_IN.frx":4053
   End
   Begin LvButtons.lvButtons_H btnPreV 
      Height          =   855
      Left            =   6960
      TabIndex        =   51
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1508
      Caption         =   "&<<"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   3
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
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnNext 
      Height          =   855
      Left            =   5880
      TabIndex        =   50
      Top             =   6120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      Caption         =   "&>>"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   3
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
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnLast 
      Height          =   855
      Left            =   7920
      TabIndex        =   49
      Top             =   6120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1508
      Caption         =   "&<"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   1
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
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnFirst 
      Height          =   855
      Left            =   5640
      TabIndex        =   48
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1508
      Caption         =   "&>"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
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
      Gradient        =   2
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnEXIT 
      Height          =   855
      Left            =   8520
      TabIndex        =   47
      Top             =   4800
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
      Image           =   "Check_IN.frx":436D
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnAdd 
      Height          =   960
      Left            =   7080
      TabIndex        =   46
      ToolTipText     =   "ADD"
      Top             =   4800
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
      Image           =   "Check_IN.frx":6021
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin LvButtons.lvButtons_H btnNew 
      Height          =   915
      Left            =   5640
      TabIndex        =   0
      ToolTipText     =   "&NEW"
      Top             =   4800
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
      Image           =   "Check_IN.frx":6C14
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.OptionButton optBUDHA 
      BackColor       =   &H80000012&
      Height          =   255
      Left            =   7320
      TabIndex        =   35
      Top             =   3360
      Width           =   255
   End
   Begin VB.OptionButton optKatholik 
      BackColor       =   &H80000012&
      Height          =   255
      Left            =   7320
      TabIndex        =   34
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton OPTHINDU 
      BackColor       =   &H80000012&
      Height          =   255
      Left            =   5880
      TabIndex        =   33
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton optISLAM 
      BackColor       =   &H80000012&
      Height          =   255
      Left            =   5880
      TabIndex        =   32
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton OptKristen 
      BackColor       =   &H80000012&
      Height          =   255
      Left            =   5880
      TabIndex        =   31
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox tIDCard 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox tID 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox tPhone 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   4680
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
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox tAge 
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
      TabIndex        =   3
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox tAddR 
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
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox tCity 
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
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox tNational 
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
      TabIndex        =   7
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   2295
      Begin VB.OptionButton optFemale 
         BackColor       =   &H00000000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H00000000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   41
         Top             =   240
         Width           =   255
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   10
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   3240
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   10
         X1              =   0
         X2              =   0
         Y1              =   360
         Y2              =   3360
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   10
         X1              =   0
         X2              =   3960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   10
         X1              =   480
         X2              =   4440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FEMALE"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   1080
         TabIndex        =   44
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MALE"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.Frame TEXT1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   735
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   6975
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash EDANbgt 
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   240
      Width           =   6975
      _cx             =   4206607
      _cy             =   4195600
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
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10320
      Top             =   9600
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   632
      X2              =   632
      Y1              =   96
      Y2              =   168
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   368
      X2              =   632
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "reservation list"
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
      Left            =   5520
      TabIndex        =   56
      Top             =   1080
      Width           =   1980
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDcard"
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
      Left            =   6720
      TabIndex        =   55
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label1 
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
      Left            =   7680
      TabIndex        =   54
      Top             =   1440
      Width           =   135
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   368
      X2              =   632
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   440
      X2              =   440
      Y1              =   96
      Y2              =   168
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   368
      X2              =   368
      Y1              =   96
      Y2              =   168
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   632
      X2              =   368
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   552
      X2              =   552
      Y1              =   312
      Y2              =   392
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   456
      X2              =   456
      Y1              =   312
      Y2              =   392
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   368
      X2              =   368
      Y1              =   312
      Y2              =   480
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   368
      X2              =   632
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   632
      X2              =   632
      Y1              =   312
      Y2              =   480
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   368
      X2              =   632
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   5520
      TabIndex        =   45
      Top             =   2640
      Width           =   1020
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   280
      X2              =   280
      Y1              =   104
      Y2              =   168
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   16
      X2              =   280
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   16
      X2              =   280
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   16
      X2              =   16
      Y1              =   104
      Y2              =   168
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BUDHA"
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
      Left            =   7800
      TabIndex        =   40
      Top             =   3360
      Width           =   780
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KATHOLIK"
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
      Left            =   7800
      TabIndex        =   39
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HINDU"
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
      Left            =   6240
      TabIndex        =   38
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ISLAM"
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
      Left            =   6240
      TabIndex        =   37
      Top             =   3000
      Width           =   645
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KRISTEN"
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
      Left            =   6240
      TabIndex        =   36
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label Label25 
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
      Left            =   1440
      TabIndex        =   30
      Top             =   1920
      Width           =   135
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
      Left            =   1440
      TabIndex        =   29
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDCARD"
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
      TabIndex        =   28
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDGUEST"
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
      TabIndex        =   27
      Top             =   1680
      Width           =   915
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   368
      X2              =   368
      Y1              =   200
      Y2              =   296
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   440
      X2              =   632
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   632
      X2              =   632
      Y1              =   184
      Y2              =   296
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   368
      X2              =   632
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Label tArriVAL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
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
      Left            =   240
      TabIndex        =   25
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label tdate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
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
      Left            =   360
      TabIndex        =   24
      Top             =   1200
      Width           =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   16
      X2              =   280
      Y1              =   408
      Y2              =   408
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   280
      X2              =   280
      Y1              =   200
      Y2              =   408
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   16
      X2              =   280
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   16
      X2              =   16
      Y1              =   200
      Y2              =   408
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
      Left            =   1920
      TabIndex        =   23
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      TabIndex        =   22
      Top             =   4680
      Width           =   735
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
      TabIndex        =   21
      Top             =   2880
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
      TabIndex        =   20
      Top             =   3480
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
      TabIndex        =   19
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label27 
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
      TabIndex        =   18
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label Label26 
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
      TabIndex        =   16
      Top             =   3120
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      TabIndex        =   15
      Top             =   3720
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Top             =   4200
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      TabIndex        =   13
      Top             =   5160
      Width           =   510
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NationaLity"
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
      Top             =   5640
      Width           =   1500
   End
End
Attribute VB_Name = "INFORMATION_GUEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim JK As String
Dim Din As String


Sub TAMPIL()
tID = CEKNULL(RsMove!idguest)
            tIDCard = CEKNULL(RsMove!idcard)
            tdate = CEKNULL(RsMove!arrivaldate)
            tname = CEKNULL(RsMove!Name)
            tAge = CEKNULL(RsMove!age)
            tAddR = CEKNULL(RsMove!address)
            tCity = CEKNULL(RsMove!City)
            'tPINCODE = CEKNULL(RsMove!PinCode)
            tPhone = CEKNULL(RsMove!phone)
            tArriVAL = CEKNULL(RsMove!ArrivalTime)
            'CmbRoom = CEKNULL(RsMove!ROOMNO)
            tNational = CEKNULL(RsMove!NATIONALITY)
           ' ttype = CEKNULL(RsMove!typeRoom)
            If (RsMove!sex = "MALE") Then
             optMale = 1
            Else
                optFemale = 1
            End If
            If (RsMove!religion = "ISLAM") Then optISLAM = 1
            If (RsMove!religion = "KRISTEN") Then OptKristen = 1
            If (RsMove!religion = "KATHOLIK") Then optKatholik = 1
            If (RsMove!religion = "HINDU") Then OPTHINDU = 1
            If (RsMove!religion = "BUDHA") Then optBUDHA = 1
            
            kunci Me
            Timer1.Enabled = False
End Sub
Sub auto()
Dim x As String
    x = Format(Date, "yymm")
With Rs

If .State = 1 Then .Close
.Open "select * from Guest ORDER BY IDGUEST ASC ", KOneKsi, adOpenKeyset, adLockReadOnly

    If .RecordCount = 0 Then

        tID = "G" + Format(Date, "yymm") + "001"
    Else
        .MoveLast

        If Left(Rs!idguest, 5) = "G" + x Then
        tID = Right(Rs!idguest, 3) + 1
        tID = "G" + Format(Date, "yymm") + Left("000", 3 - Len(tID)) + tID
        Else

         tID = "G" + Format(Date, "yymm") + "001"
        End If
    End If
End With
End Sub

Private Sub btnAdd_Click()


If tIDCard.Text = "" Then
MsgBox "Please Enter IDCARD", vbExclamation, "mYHoTEL"
tIDCard.SetFocus
Else

If tname.Text = "" Then
MsgBox "Please Enter name", vbExclamation, "mYHoTEL"
tname.SetFocus
Else
If tAge.Text = "" Then
MsgBox "Please Enter age", vbExclamation, "mYHoTEL"
tAge.SetFocus
Else
If tAddR.Text = "" Then
MsgBox "Please Enter address", vbExclamation, "mYHoTEL"
tAddR.SetFocus
Else
If tCity.Text = "" Then
MsgBox "Please Enter city", vbExclamation, "mYHoTEL"
tCity.SetFocus
Else
If tPhone.Text = "" Then
MsgBox "Please Enter phone", vbExclamation, "mYHoTEL"
tPhone.SetFocus
Else
If tNational.Text = "" Then
MsgBox "Please Enter Nationality", vbExclamation, "mYHoTEL"
tNational.SetFocus
Else
If optMale.Value = False And optFemale.Value = False Then
MsgBox "Please Enter sex", vbExclamation, "mYHoTEL"
Else
If optISLAM.Value = False And optKatholik.Value = False And OptKristen.Value = False And OPTHINDU.Value = False And optBUDHA.Value = False Then
MsgBox "Please Enter Religion", vbExclamation, "mYHoTEL"
Else
If Rs.State = 1 Then Rs.Close
    Rs.Open "select * from Guest where idGuest='" & tID & "'", KOneKsi, 3, 3
        If Rs.EOF Then
       
         KOneKsi.Execute "insert into Guest(idcard,Religion,nationality,idGuest,arrivaldate,arrivaltime,name,sex,age,address,city,phone)values('" & Replace(tIDCard.Text, "'", "") & "','" & Replace(Din, "'", "") & "','" & Replace(tNational, "'", "") & "','" & Replace(tID, "'", "") & "',#" & tdate & "#,#" & tArriVAL & "#,'" & Replace(tname, "'", "''") & "','" & Replace(JK, "'", "") & "'," & Replace(tAge, "'", "") & ",'" & Replace(tAddR, "'", "") & "','" & Replace(tCity, "'", "") & "'," & Replace(tPhone, "'", "") & ")"
                   
        MsgBox ("Data added. for visitor") + " " + tname.Text, vbInformation, "mYHoTEL"
        Else
         MsgBox "SORRY, Is You Have IDGUEST !!!", vbExclamation, "mYHoTEL"
        End If


End If
End If
End If
End If
End If
End If
End If
End If
End If

End Sub





Private Sub btnEXIT_Click()
keluar Me
splashHidup
End Sub

Private Sub btnFirst_Click()

RsMove.MoveFirst
If Not RsMove.EOF Then
    
            TAMPIL
  End If
End Sub

Private Sub btnLast_Click()

RsMove.MoveLast
If Not RsMove.EOF Then
           TAMPIL
  End If
End Sub

Private Sub btnNew_Click()

bersih Me
Buka Me
tID.Locked = True
Timer1.Enabled = True
auto
tIDCard.SetFocus


End Sub

Private Sub btnNext_Click()


RsMove.MoveNext
  
  If RsMove.EOF = True Then
  MsgBox "END OF FILE"
  RsMove.MoveLast
 Else
 TAMPIL
End If
End Sub

Private Sub btnPreV_Click()
'
RsMove.MovePrevious
  
  If RsMove.BOF = True Then
  MsgBox "BEGIN OF FILE"
  RsMove.MoveFirst
  Else
      TAMPIL
End If
End Sub







Private Sub Command1_Click()

'BILL.Show
End Sub





Private Sub cmbIDcard_Click()
If Rs.State = 1 Then Rs.Close

    Rs.Open "select * from reservation where idcard='" & cmbIDcard & "'", KOneKsi, 3, 3
    If Not Rs.EOF Then
            tname = CEKNULL(Rs!Name)
            tAddR = CEKNULL(Rs!address)
            tPhone = CEKNULL(Rs!phone)
            tIDCard = CEKNULL(Rs!idcard)
            tAge.SetFocus
    End If
End Sub

Private Sub cmdReservation_Click()
If Rs.State = 1 Then Rs.Close
 cmbIDcard.Clear
    Rs.Open "select * from reservation where arrivaldate=#" & CDate(tdate) & "#", KOneKsi, 3, 3
        If Not Rs.EOF Then
            Buka Me
            While Not Rs.EOF
                    cmbIDcard.AddItem Rs!idcard
                    Rs.MoveNext
              Wend
        
        auto
        
         Else
            MsgBox "Not Reservation List", vbExclamation, "mYHoTEL"
      
        
        End If
        
        
        
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{TAB}")

End Sub

Private Sub Form_Load()
'ReZiseForm Me, EDANbgt
awal Me
Me.Height = 7980
Me.Width = 10320
splashMati
TEXT1.Visible = False
EDANbgt.Movie = App.Path & "\Document\gUEST1.SWF"
EDANbgt.Play
kunci Me
btnNew.Enabled = True
btnAdd.Enabled = True

 OPENDATA
'If Rs.State = 1 Then Rs.Close
''    CmbRoom.Clear
''    CmbRoom.Clear
'    Rs.Open "Select * from Room where status=false", KOneKsi, 3, 3
'        If Not Rs.EOF Then
'            While Not Rs.EOF
'               CmbRoom.AddItem Rs!ROOMNO
'               Rs.MoveNext
'            Wend
'        End If
If RsMove.State = 1 Then RsMove.Close
RsMove.Open "select * from Guest", KOneKsi, 3, 3


End Sub






Private Sub Form_Unload(Cancel As Integer)
Call btnEXIT_Click
End Sub

Private Sub optBUDHA_Click()
If optBUDHA = True Then Din = "BUDHA"
End Sub

Private Sub optFemale_Click()
If optFemale = True Then JK = "FEMALE"
End Sub

Private Sub OPTHINDU_Click()
If OPTHINDU = True Then Din = "HINDU"
End Sub

Private Sub optISLAM_Click()
If optISLAM = True Then Din = "ISLAM"
End Sub

Private Sub optKatholik_Click()
If optKatholik = True Then Din = "KATHOLIK"
End Sub

Private Sub OptKristen_Click()
If OptKristen = True Then Din = "KRISTEN"
End Sub

Private Sub optMale_Click()
If optMale = True Then JK = "MALE"
End Sub

Private Sub tAge_Change()
If Not IsNumeric(tAge) Then tAge = ""
End Sub




Private Sub Timer1_Timer()
tdate = Date
tArriVAL = Time()
End Sub

Private Sub tPhone_Change()
If Not IsNumeric(tPhone) Then tPhone = ""
End Sub




