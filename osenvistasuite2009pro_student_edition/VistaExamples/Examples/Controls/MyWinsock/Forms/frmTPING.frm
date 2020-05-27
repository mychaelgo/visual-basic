VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Ping responce ..."
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   LinkTopic       =   "Form2"
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.MyWinsock MyWinsock1 
      Left            =   3210
      Top             =   2490
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VistaSuitePro.OsenVistaTextBox TxtResult 
      Height          =   1395
      Left            =   150
      TabIndex        =   1
      Top             =   450
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2461
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      BackColor       =   0
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   0
      ForeColorOver   =   65280
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
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ping responce ..."
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Form_Unload(Cancel As Integer)
    MyWinsock1.Disconnect
End Sub

Private Sub MyWinsock1_PingResponce(IpAddress As String, DataSize As Long, TripTime As Long, TTL As Byte, ErrorDesc As String)
    If ErrorDesc <> "" Then
        txtresult.Text = txtresult.Text & ErrorDesc & vbCrLf
    Else
        txtresult.Text = txtresult.Text & "Reply from " & IpAddress & ":   Bytes=" & DataSize & "   time=" & TripTime & "   TTL=" & TTL & vbCrLf
    End If
End Sub

Public Sub Ping(IpAddess As String)

    Dim l As Byte
    Me.Show
    txtresult.Text = ""
    DoEvents
    
    For l = 1 To 4
        Me.MyWinsock1.Ping IpAddess
        DoEvents
    Next
    
    txtresult.Text = txtresult.Text & vbCrLf & "Done."
    
    ' waiting ...
    WaitTimes 7000
    
    Unload Me
    
End Sub























