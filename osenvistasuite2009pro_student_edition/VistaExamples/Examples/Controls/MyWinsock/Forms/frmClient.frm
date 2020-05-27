VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frmClient 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Client"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   588
   StartUpPosition =   3  'Windows Default
   Begin VistaSuitePro.OsenVistaStatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   4140
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   979
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   3
      HaveXPForm      =   -1  'True
      WindowColor     =   3
      PWidth1         =   200
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Localhostname:"
      pTextAlignment1 =   0
      PanelPicture1   =   "frmClient.frx":038A
      PanelPicAlignment1=   0
      PWidth2         =   140
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "LocalIP:"
      pTextAlignment2 =   0
      PanelPicture2   =   "frmClient.frx":03A6
      PanelPicAlignment2=   0
      PWidth3         =   108
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "LocalPort:"
      pTextAlignment3 =   0
      PanelPicture3   =   "frmClient.frx":03C2
      PanelPicAlignment3=   0
      GradientColor1  =   15382160
      GradientColor2  =   12752244
   End
   Begin VistaSuitePro.MyWinsock Client 
      Left            =   4470
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VistaSuitePro.OsenVistaTextBox txtLOG 
      Height          =   2715
      Left            =   3090
      TabIndex        =   5
      Top             =   1350
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   4789
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
   Begin VistaSuitePro.OsenVistaListBox lstLog 
      Height          =   2715
      Left            =   270
      TabIndex        =   4
      Top             =   1350
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4789
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
      BorderColor     =   13603685
      ShowHeader      =   -1  'True
      HeaderFormatString=   "Message;180;0;0;"
      Columns         =   1
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      MaxAllColumnWidth=   180
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
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
      BinaryImage     =   "frmClient.frx":03DE
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaButton cmdDisconnect 
      Height          =   345
      Left            =   7470
      TabIndex        =   3
      Top             =   960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Caption         =   "&Disconnet"
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
      MICON           =   "frmClient.frx":03F6
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "frmClient.frx":0412
      BinaryImageOver =   "frmClient.frx":042A
   End
   Begin VistaSuitePro.OsenVistaButton cmdSend 
      Height          =   345
      Left            =   7470
      TabIndex        =   2
      Top             =   540
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Caption         =   "Send"
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
      MICON           =   "frmClient.frx":0442
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "frmClient.frx":045E
      BinaryImageOver =   "frmClient.frx":0476
   End
   Begin VistaSuitePro.OsenVistaTextBox txtMsg 
      Height          =   795
      Left            =   270
      TabIndex        =   1
      Top             =   510
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   1402
      Text            =   $"frmClient.frx":048E
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
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
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
      Caption         =   "Client"
      TitleTop        =   7
      icon            =   "frmClient.frx":22B7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
   
End Sub

Private Sub Client_DataArrival(Index As Long, Data As String, Sck As VistaSuitePro.IWinsock)

    lstLog.AddItem Now() & " [" & Len(Data) & " bytes ]", StrKey:=Data
    txtLog.Text = Data

End Sub

Private Sub Client_OnClose(Index As Long, Sck As VistaSuitePro.IWinsock)

    Unload Me

End Sub

Private Sub Client_OnConnect(Index As Long, Sck As VistaSuitePro.IWinsock)

    Me.OsenXPForm1.Caption = "Client connected to " & Client.RemoteIP & ":" & Client.RemotePort
    sBar.PanelCaption(1) = "Localhostname: " & Client.LocalHost
    sBar.PanelCaption(2) = "LocalIP: " & Client.LocalIP
    sBar.PanelCaption(3) = "LocalPort: " & Client.LocalPort

End Sub

Private Sub cmdDisconnect_Click()

    Client.Disconnect
    Unload Me

End Sub

Private Sub cmdSend_Click()

    If txtMsg.Text <> "" Then
        Client.SendData txtMsg.Text
    End If

End Sub

Public Sub Connect(Port As Integer)

    Client.Connect "localhost", Port
    DoEvents
    Client.SendData "Test .... sent by client :)"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Client.Disconnect

End Sub

Private Sub lstLog_ListClick(CurrentIndex As Long)

    txtLog.Text = lstLog.Item(CurrentIndex).Key

End Sub




