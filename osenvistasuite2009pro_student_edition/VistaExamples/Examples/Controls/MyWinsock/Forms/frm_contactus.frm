VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_contactUs 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Contact Us"
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   Icon            =   "frm_contactus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   578
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture2 
      Height          =   495
      Left            =   4590
      TabIndex        =   11
      Top             =   8010
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
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
      BorderColor     =   14854529
      GradientColor2  =   14854529
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BinaryImage     =   "frm_contactus.frx":058A
   End
   Begin VistaSuitePro.OsenVistaProgressBar PbAR 
      Height          =   315
      Left            =   270
      TabIndex        =   9
      Top             =   8100
      Visible         =   0   'False
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Value           =   100
   End
   Begin VistaSuitePro.OsenVistaButton cmdSend 
      Height          =   345
      Left            =   7170
      TabIndex        =   5
      Top             =   8070
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   "&Send"
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
      MICON           =   "frm_contactus.frx":05A2
      PICN            =   "frm_contactus.frx":05BE
      UMCOL           =   -1  'True
      BinaryImageNormal=   "frm_contactus.frx":0958
      BinaryImageOver =   "frm_contactus.frx":0970
   End
   Begin VistaSuitePro.MyWinsock MyWinsock1 
      Left            =   5910
      Top             =   5580
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VistaSuitePro.OsenVistaListBox lstInfo 
      Height          =   1875
      Left            =   240
      TabIndex        =   8
      Top             =   6030
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3307
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontNormal      =   0
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      ShowHeader      =   -1  'True
      HeaderFormatString=   "Status;545;0;0;;-1"
      Columns         =   1
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      RowColor1       =   12648447
      RowColor2       =   16777152
      MaxAllColumnWidth=   545
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
      AllowSortItem   =   0   'False
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridHeaderMode  =   0   'False
      BinaryImage     =   "frm_contactus.frx":0988
   End
   Begin VistaSuitePro.OsenVistaTextBox txtName 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   582
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      Required        =   -1  'True
      LabelBackColor  =   15790320
      LabelCaption    =   "Name:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   1035
      Left            =   0
      TabIndex        =   7
      Top             =   420
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   1826
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
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "OsenVistaSuite Bug Reports, Suggestions and Feature Requests Form."
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Contact us, Bug Reports, Suggestions and Feature Requests "
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
      BinaryImage     =   "frm_contactus.frx":09A0
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
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
      Caption         =   "Contact Us"
      TitleTop        =   7
      icon            =   "frm_contactus.frx":09B8
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaTextBox txtFrom 
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   1950
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   582
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      Required        =   -1  'True
      LabelBackColor  =   15790320
      LabelCaption    =   "Email:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtSubject 
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   2340
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   582
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      Required        =   -1  'True
      LabelBackColor  =   15790320
      LabelCaption    =   "Subject:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtAttach 
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   2730
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   582
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
      Locked          =   -1  'True
      ButtonCaption   =   ""
      ButtonPicture   =   "frm_contactus.frx":0F52
      ButtonVisible   =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      BackColorOver   =   12648447
      LabelBackColor  =   15790320
      LabelCaption    =   "Attachment:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtMessage 
      Height          =   2835
      Left            =   270
      TabIndex        =   4
      Top             =   3150
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   5001
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      MultiLine       =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      Required        =   -1  'True
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Priority:"
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
      Left            =   6030
      TabIndex        =   10
      Top             =   1650
      Width           =   660
   End
End
Attribute VB_Name = "frm_contactUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Xttip As CLS_XToolTip

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
    
    Set Xttip = New CLS_XToolTip
    With Xttip
    
        .CreateTooltip txtName.hWnd, "Enter your name here.", , TTBalloonIfActive, , TTIconInfo
        
        .CreateTooltip txtFrom.hWnd, "Please enter your valid email address here.", , TTBalloonIfActive, , TTIconWarning, "Email address"
        
        .CreateTooltip txtSubject.hWnd, "Please enter a subject title that summarizes your problem.", , TTBalloonIfActive, , TTIconWarning, "Subject"
        
        .CreateTooltip txtAttach.hWnd, "Attach a screenshot or sample project that demonstrates the problem you are having (Zip file is recommended).", , TTBalloonIfActive, , TTIconWarning, "Attachment"
        
        
    End With
    
    ' Init last value from registry
    txtName = GetSetting("OsenXPSuite 2006", "Contact Us", "Name")
    txtFrom = GetSetting("OsenXPSuite 2006", "Contact Us", "email")
    txtSubject = GetSetting("OsenXPSuite 2006", "Contact Us", "subject")
    txtAttach = GetSetting("OsenXPSuite 2006", "Contact Us", "attachment")
    txtMessage = GetSetting("OsenXPSuite 2006", "Contact Us", "message")
    
    ' Loading Animate GIF
    OsenXPPicture1.OpenPicture App.Path & "\at3.gif"
    OsenXPPicture2.OpenPicture App.Path & "\mail1.gif"
    
End Sub

' SendMail
Private Sub cmdSend_Click()

    lstInfo.Clear
    PBar.Visible = True
    OsenXPPicture2.Visible = True
    If MyWinsock1.SendMail("osenxpsuite.net", txtName & "<" & txtFrom & ">", "support@osenxpsuite.net", txtSubject, txtMessage, 1, , txtAttach) Then
        MsgBoxGT "Message successfull sent.", vbInformation, "Success", 3
    Else
        MsgBoxGT "Cannot sent message.", vbCritical, "Fail", 3
    End If
    OsenXPPicture2.Visible = False
    OsenXPPicture1.Description = "OsenXPSuite 2006's Bug Reports, Suggestions and Feature Requests Form."
    PBar.Visible = False
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Set Xttip = Nothing
    Me.MyWinsock1.Disconnect
End Sub

Private Sub MyWinsock1_DataArrival(Index As Long, Data As String, Sck As VistaSuitePro.IWinsock)
OsenXPPicture1.Description = Data

End Sub

' Trace sendmail status ...
Private Sub MyWinsock1_SendMailInfo(Info As String)

   Call lstInfo.AddItem(Info, CustomColor:=IIf(LCase$(Left$(Info, 7)) = "command", &HFF0000, &HC000&))
    
End Sub

' Progress ...
Private Sub MyWinsock1_SendMailProgress(Progress As Long)
    PBar.Value = Progress
End Sub

' Show open dialog from txtAttach
Private Sub txtAttach_ButtonClick()
     txtAttach.ShowOpenDialog
End Sub
 
Private Sub txtAttach_Change()
    SaveSetting "OsenXPSuite 2006", "Contact Us", "Attachment", txtAttach
End Sub

Private Sub txtFrom_Change()
    SaveSetting "OsenXPSuite 2006", "Contact Us", "EMail", txtFrom
End Sub

Private Sub txtMessage_Change()
    SaveSetting "OsenXPSuite 2006", "Contact Us", "Message", txtMessage
End Sub

Private Sub txtName_Change()
    SaveSetting "OsenXPSuite 2006", "Contact Us", "Name", txtName
End Sub

Private Sub txtSubject_Change()
    SaveSetting "OsenXPSuite 2006", "Contact Us", "Subject", txtSubject.Text
End Sub






















