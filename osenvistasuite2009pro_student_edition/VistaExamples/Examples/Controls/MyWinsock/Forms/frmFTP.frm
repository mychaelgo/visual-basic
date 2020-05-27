VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_upload 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "File Transfer ..."
   ClientHeight    =   5505
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton cmdUpload 
      Height          =   345
      Left            =   4920
      TabIndex        =   6
      Top             =   4800
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      Caption         =   "Upload now!"
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
      MPTR            =   99
      MICON           =   "frmFTP.frx":038A
      PICN            =   "frmFTP.frx":04EC
      UMCOL           =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frmFTP.frx":0A86
      BinaryImageOver =   "frmFTP.frx":0A9E
   End
   Begin VistaSuitePro.OsenVistaProgressBar PBar 
      Height          =   315
      Left            =   210
      TabIndex        =   17
      Top             =   4830
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   2871848
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VistaSuitePro.MyWinsock MyWinsock1 
      Left            =   4770
      Top             =   5940
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame2 
      Height          =   1575
      Left            =   210
      TabIndex        =   13
      Top             =   3150
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   2778
      Caption         =   "File Info:"
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
      BorderColor     =   14396553
      Appearance      =   1
      image           =   "frmFTP.frx":0AB6
      BinaryImage     =   "frmFTP.frx":0E50
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaTextBox txtRFN 
         Height          =   345
         Left            =   2430
         TabIndex        =   5
         Top             =   1170
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   609
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
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
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
      Begin VistaSuitePro.OsenVistaTextBox txtFolder 
         Height          =   345
         Left            =   2430
         TabIndex        =   4
         Top             =   780
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   609
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
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
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
      Begin VistaSuitePro.OsenVistaTextBox txtLCF 
         Height          =   345
         Left            =   2430
         TabIndex        =   3
         Top             =   390
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   609
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
         ButtonCaption   =   ""
         ButtonPicture   =   "frmFTP.frx":0E68
         ButtonVisible   =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel4 
         Height          =   255
         Left            =   270
         TabIndex        =   14
         Top             =   1200
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Remote filename*:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel5 
         Height          =   255
         Left            =   270
         TabIndex        =   15
         Top             =   810
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Destination directory*:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel6 
         Height          =   285
         Left            =   270
         TabIndex        =   16
         Top             =   420
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Local filename:"
         ForeColor       =   0
         BackStyle       =   0
      End
   End
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
      Height          =   1665
      Left            =   210
      TabIndex        =   9
      Top             =   1410
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   2937
      Caption         =   "FTP server information:"
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
      BorderColor     =   14396553
      Appearance      =   1
      image           =   "frmFTP.frx":1202
      BinaryImage     =   "frmFTP.frx":179C
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaTextBox txtPass 
         Height          =   345
         Left            =   2430
         TabIndex        =   2
         Top             =   1230
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   609
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
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
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
      Begin VistaSuitePro.OsenVistaTextBox txtUser 
         Height          =   345
         Left            =   2430
         TabIndex        =   1
         Top             =   840
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   609
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
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
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
      Begin VistaSuitePro.OsenVistaTextBox SvrName 
         Height          =   345
         Left            =   2430
         TabIndex        =   0
         Top             =   450
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   609
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
         ButtonCaption   =   "Ping"
         ButtonPicture   =   "frmFTP.frx":17B4
         ButtonVisible   =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
         BackColorOver   =   12648447
         AutoTab         =   -1  'True
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
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel3 
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   1260
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Password:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel2 
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   870
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Username:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VistaSuitePro.OsenVistaLabel OsenXPLabel1 
         Height          =   285
         Left            =   270
         TabIndex        =   10
         Top             =   480
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Server name (hostname):"
         ForeColor       =   0
         BackStyle       =   0
      End
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   420
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmFTP.frx":1D4E
      BorderColor     =   14854529
      PictureAlignment=   7
      BackStyle       =   0
      GradientBackGround=   -1  'True
      GradientColor2  =   16310477
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "Upload your file into FTP server ..."
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "FTP Utility"
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
      BinaryImage     =   "frmFTP.frx":38A0
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
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
      Caption         =   "File Transfer ..."
      TitleTop        =   7
      icon            =   "frmFTP.frx":38B8
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
End
Attribute VB_Name = "frm_upload"
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
    Me.MyWinsock1.Disconnect
End Sub

Private Sub MyWinsock1_DataArrival(Index As Long, Data As String, Sck As VistaSuitePro.IWinsock)
    Debug.Print Data
End Sub

Private Sub MyWinsock1_FTPUploadProgress(Progress As Long)
    PBar.Value = Progress
    DoEvents
End Sub

Private Sub SvrName_ButtonClick()
    If SvrName.Text <> "" Then
        Form2.Ping SvrName.Text
    End If
End Sub

Private Sub txtLCF_ButtonClick()
    ' Show std open diagog
    txtLCF.ShowOpenDialog
End Sub

Private Sub cmdUpload_Click()
    If MyWinsock1.UploadFile(SvrName, txtUser, txtPass, txtLCF, txtFolder, txtRFN) Then
        MsgBoxGT "Your file has been successfull transfered!", vbExclamation
    Else
        MsgBoxGT "Could not upload your file!", vbCritical, "Error found"
    End If
End Sub























