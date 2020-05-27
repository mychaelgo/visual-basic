VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_ping 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "PING && Trace route"
   ClientHeight    =   9330
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPING.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton2 
      Height          =   315
      Left            =   5790
      TabIndex        =   6
      Top             =   1380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "Tracert"
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
      MICON           =   "frmPING.frx":000C
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VB.TextBox txtresult 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   7185
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmPING.frx":0028
      Top             =   1800
      Width           =   7425
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   1380
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "&Ping"
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
      MICON           =   "frmPING.frx":002E
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
   End
   Begin VistaSuitePro.MyWinsock MyWinsock1 
      Left            =   7110
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VistaSuitePro.OsenVistaTextBox txtIP 
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Top             =   1380
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      Text            =   "osenxpsuite.net"
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
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   1482
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   14854529
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "MyWinsock sample usage (PING)."
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "OsenXPSuite 2006 Enterprise Edition"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "frmPING.frx":004A
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PING && Trace route"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hostname/IP address:"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   1410
      Width           =   1605
   End
End
Attribute VB_Name = "frm_ping"
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

Private Sub MyWinsock1_PingResponce(IpAddress As String, DataSize As Long, TripTime As Long, TTL As Byte, ErrorDesc As String)
    If ErrorDesc <> "" Then
        txtresult.Text = txtresult.Text & ErrorDesc & vbCrLf
    Else
        txtresult.Text = txtresult.Text & "Reply from " & IpAddress & ":   Bytes=" & DataSize & "   time=" & TripTime & "   TTL=" & TTL & vbCrLf
    End If

End Sub

Private Sub MyWinsock1_TracertInfo(Info As String)
    txtresult.Text = txtresult.Text & vbCrLf & Info & vbCrLf & vbCrLf
End Sub

' Tracert responce
Private Sub MyWinsock1_TracertResponce(IpAddress As String, HopNumber As Long, TripTime As Long, TTL As Byte, ErrorDesc As String)
    
    If ErrorDesc <> "" Then
        txtresult.Text = txtresult.Text & HopNumber & vbTab & "*" & vbTab & vbTab & ErrorDesc & vbCrLf
    Else
        txtresult.Text = txtresult.Text & HopNumber & vbTab & TripTime & " ms" & vbTab & vbTab & IpAddress & vbCrLf
    End If
    
End Sub

' PING
Private Sub OsenXPButton1_Click()
    Dim l As Byte
    
    txtresult.Text = ""
    
    For l = 1 To 4
        Me.MyWinsock1.Ping txtIP
        DoEvents
    Next
    
    txtresult.Text = txtresult.Text & vbCrLf & "Done."
    
End Sub

' Tracert
Private Sub OsenXPButton2_Click()

    txtresult.Text = ""
    MyWinsock1.Tracert txtIP
    
End Sub























