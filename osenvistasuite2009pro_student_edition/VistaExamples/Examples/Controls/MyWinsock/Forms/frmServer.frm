VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frmServer 
   BackColor       =   &H00EAF7F7&
   BorderStyle     =   0  'None
   Caption         =   "Server"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11625
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyWinsock Server 
      Left            =   10560
      Top             =   4230
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel3 
      Height          =   405
      Left            =   8910
      TabIndex        =   17
      Top             =   4530
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmServer.frx":058A
      Caption         =   "Total BytesRcv:"
      ForeColor       =   0
      AutoSize        =   0   'False
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaLabel OsenXPLabel2 
      Height          =   435
      Left            =   8880
      TabIndex        =   16
      Top             =   3840
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmServer.frx":0B24
      Caption         =   "Total BytesSent:"
      ForeColor       =   0
      AutoSize        =   0   'False
      BackStyle       =   0
   End
   Begin VistaSuitePro.OsenVistaButton cmdExit 
      Height          =   375
      Left            =   10260
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Exit"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frmServer.frx":10BE
      PICN            =   "frmServer.frx":10DA
      UMCOL           =   -1  'True
      BinaryImageNormal=   "frmServer.frx":1474
      BinaryImageOver =   "frmServer.frx":148C
   End
   Begin VistaSuitePro.OsenVistaButton cmdAddClient 
      Height          =   375
      Left            =   8940
      TabIndex        =   12
      Top             =   5160
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      Caption         =   "&Add Client"
      Enabled         =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frmServer.frx":14A4
      PICN            =   "frmServer.frx":14C0
      UMCOL           =   -1  'True
      BinaryImageNormal=   "frmServer.frx":185A
      BinaryImageOver =   "frmServer.frx":1872
   End
   Begin VistaSuitePro.OsenVistaFrame FraMsg 
      Height          =   1755
      Left            =   2280
      TabIndex        =   5
      Top             =   3810
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   3096
      Caption         =   "Send Message"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14215660
      ForeColor       =   0
      BorderColor     =   12164479
      Appearance      =   1
      BinaryImage     =   "frmServer.frx":188A
      WindowColor     =   2
      GradientColor1  =   16777215
      GradientColor2  =   10522143
      Begin VistaSuitePro.OsenVistaButton cmdSendAll 
         Height          =   345
         Left            =   5190
         TabIndex        =   20
         Top             =   780
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         Caption         =   "Send &All"
         Enabled         =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "frmServer.frx":18A2
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         BinaryImageNormal=   "frmServer.frx":18BE
         BinaryImageOver =   "frmServer.frx":18D6
      End
      Begin VistaSuitePro.OsenVistaButton cmdSend 
         Height          =   345
         Left            =   5190
         TabIndex        =   19
         Top             =   390
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         Caption         =   "Send"
         Enabled         =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "frmServer.frx":18EE
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         BinaryImageNormal=   "frmServer.frx":190A
         BinaryImageOver =   "frmServer.frx":1922
      End
      Begin VistaSuitePro.OsenVistaTextBox txtMsg 
         Height          =   1305
         Left            =   60
         TabIndex        =   18
         Top             =   390
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2302
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MultiLine       =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VistaSuitePro.OsenVistaFrame FrameLOG 
      Height          =   2805
      Left            =   180
      TabIndex        =   4
      Top             =   5610
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   4948
      Caption         =   "Log"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14215660
      ForeColor       =   0
      BorderColor     =   12164479
      Appearance      =   1
      BinaryImage     =   "frmServer.frx":193A
      WindowColor     =   2
      GradientColor1  =   16777215
      GradientColor2  =   10522143
      Begin VistaSuitePro.OsenVistaTextBox txtLog 
         Height          =   2415
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4260
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MultiLine       =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaTreeView tvwLog 
         Height          =   2415
         Left            =   30
         TabIndex        =   6
         Top             =   360
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   4260
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SelectedBackColor=   9471874
         SelectedColor   =   16777215
         ShowNumber      =   -1  'True
         BorderStyle     =   0
         HeaderCaption   =   "OsenXPTreeView1"
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   15779735
         HeaderForeColor =   16777215
         WindowColor     =   2
      End
   End
   Begin VistaSuitePro.OsenVistaFrame FraCFG 
      Height          =   1755
      Left            =   180
      TabIndex        =   3
      Top             =   3810
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3096
      Caption         =   "Configuration"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14215660
      ForeColor       =   0
      BorderColor     =   12164479
      Appearance      =   1
      BinaryImage     =   "frmServer.frx":1952
      WindowColor     =   2
      GradientColor1  =   16777215
      GradientColor2  =   10522143
      Begin VistaSuitePro.OsenVistaButton cmdDisconnect 
         Height          =   375
         Left            =   150
         TabIndex        =   11
         Top             =   1260
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "frmServer.frx":196A
         PICN            =   "frmServer.frx":1986
         UMCOL           =   -1  'True
         BinaryImageNormal=   "frmServer.frx":1D20
         BinaryImageOver =   "frmServer.frx":1D38
      End
      Begin VistaSuitePro.OsenVistaButton cmdListen 
         Height          =   375
         Left            =   150
         TabIndex        =   10
         Top             =   840
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         Caption         =   "&Listen"
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "frmServer.frx":1D50
         PICN            =   "frmServer.frx":1D6C
         UMCOL           =   -1  'True
         BinaryImageNormal=   "frmServer.frx":2106
         BinaryImageOver =   "frmServer.frx":211E
      End
      Begin VistaSuitePro.OsenVistaLabel lbPort 
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   450
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LocalPort:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaTextBox txtPort 
         Height          =   315
         Left            =   1050
         TabIndex        =   8
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Text            =   "1022"
         Alignment       =   2
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
         ForeColor       =   0
         Value           =   1022
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOver   =   12648447
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VistaSuitePro.MyImageList MyImageList1 
      Left            =   10260
      Top             =   2640
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   5740
      Images          =   "frmServer.frx":2136
      Version         =   131072
      KeyCount        =   5
      Keys            =   "ÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaListBox LstConnection 
      Height          =   2205
      Left            =   180
      TabIndex        =   2
      Top             =   1530
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   3889
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   16777215
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      BorderColor     =   9471874
      ShowHeader      =   -1  'True
      HeaderFormatString=   $"frmServer.frx":37C2
      Columns         =   8
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      MaxAllColumnWidth=   720
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   "MyImageList1"
      ForeColorSelected=   16576
      HeaderCaption   =   "OsenXPListBox1"
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "frmServer.frx":3856
      WindowColor     =   2
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1720
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmServer.frx":386E
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   13089392
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Space           =   18
      Description     =   "<Description here ... xxxxxxxxx xxxxxxxxxxx xxxxxxx xxx.xxx.xxx.xxx>"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "MyWinsock ActiveX Control - Demo"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "frmServer.frx":53C0
      WindowColor     =   2
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Server"
      TitleTop        =   7
      icon            =   "frmServer.frx":53D8
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   2
   End
   Begin VB.Label lbRcv 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 MB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10845
      TabIndex        =   15
      Top             =   4680
      Width           =   450
   End
   Begin VB.Label lbSent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 MB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10845
      TabIndex        =   14
      Top             =   3900
      Width           =   450
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
    tvwLog.ImageList = MyImageList1.hIml
    tvwLog.Nodes.Add "SERVER", "Connection", 4
    tvwLog.Nodes(1).Expanded = True

End Sub

Private Sub cmdAddClient_Click()

  Dim c As New frmClient

    c.Connect txtPort.Value
    c.Show

End Sub

Private Sub cmdDisconnect_Click()

  ' Stop server

    Server.Disconnect

    ' Clear list
    LstConnection.Clear

    lbSent.Caption = ""
    lbRcv.Caption = ""

    cmdListen.Enabled = True
    txtPort.Enabled = True
    cmdAddClient.Enabled = False
    cmdDisconnect.Enabled = False
    tvwLog.Clear
    tvwLog.Nodes.Add "SERVER", "Connection", 4
    tvwLog.Nodes(1).Expanded = True

End Sub

Private Sub cmdExit_Click()

    Server.Disconnect
    Unload Me

End Sub

Private Sub cmdListen_Click()

  'Listen

    Server.Listen txtPort.Value
    txtPort.Enabled = False
    cmdListen.Enabled = False

    cmdDisconnect.Enabled = True
    cmdAddClient.Enabled = True

End Sub

Private Sub cmdSend_Click()

    Server(LstConnection.ItemData(LstConnection.ListIndex)).SendData txtMsg
    txtMsg.Text = ""

End Sub

Private Sub cmdSendAll_Click()

  Dim l As Long
  Dim n As Long

    n = Server.SocketCount
    For l = 1 To n
        If Server(l).State = SckIdle Or Server(l).State = SckConnecting Then
            Server(l).SendData txtMsg.Text
        End If
    Next
    txtMsg.Text = ""

End Sub


Private Sub LstConnection_ListClick(CurrentIndex As Long)

    FraMsg.Caption = "Send message to " & LstConnection.Cell(CurrentIndex, 3) & ":" & LstConnection.Cell(CurrentIndex, 4)
    cmdSend.Enabled = True
    cmdSendAll.Enabled = True

End Sub

Private Sub Server_ConnectionRequest(RequestID As Long)

  Dim l As Long, M As Long

    l = Server.Accept(RequestID)

    M = LstConnection.AddItem(l & vbTab & Server(l).SocketHandle & vbTab & Server(l).RemoteHostName & vbTab & Server(l).RemoteIP & vbTab & Server(l).RemotePort & vbTab & Now() & vbTab & 0 & vbTab & 0, , , l)
    LstConnection.SetReportIcon M, 0, 0
    LstConnection.SetReportIcon M, 6, 2
    LstConnection.SetReportIcon M, 7, 3

    tvwLog.Nodes.Add "KEY_" & Server(l).SocketHandle, Server(l).RemoteIP & ":" & Server(l).RemotePort & "(" & Server(l).SocketHandle & ")", 0, , "SERVER"

    Server(l).Value = M
    tvwLog.Nodes("KEY_" & Server(l).SocketHandle).Expanded = True
    tvwLog.Nodes("server").ShowChildCount
    tvwLog.Refresh

End Sub

Private Sub Server_DataArrival(Index As Long, Data As String, Sck As VistaSuitePro.IWinsock)

  Dim stra As String

    stra = "KEY_" & Server(Index).SocketHandle

    LstConnection.Cell(Sck.Value, 6) = Sck.BytesReceived
    lbRcv.Caption = Format$(Server.TotalBytesReceived / 1024 / 1024, "0.00 MB")
    tvwLog.Nodes.Add "DATA_" & stra & MD5(Timer), Now(), 3, , stra, True, Data, , ItemData:=Len(Data)
    tvwLog.Nodes(stra).ShowChildCount
    tvwLog.Refresh

End Sub

Private Sub Server_OnClose(Index As Long, Sck As VistaSuitePro.IWinsock)

    LstConnection.SetReportIcon Sck.Value, 0, 1

End Sub

Private Sub Server_SendComplete(Index As Long, Sck As VistaSuitePro.IWinsock)

    LstConnection.Cell(Sck.Value, 7) = Sck.BytesSent
    lbSent.Caption = Format$(Server.TotalBytesSent / 1024 / 1024, "0.00 MB")

End Sub

Private Sub tvwLog_NodeClick(Node As VistaSuitePro.CLS_xpNode)

    If Left$(Node.Key, 4) = "DATA" Then
        txtLog.Text = Node.Data
    End If

End Sub






















