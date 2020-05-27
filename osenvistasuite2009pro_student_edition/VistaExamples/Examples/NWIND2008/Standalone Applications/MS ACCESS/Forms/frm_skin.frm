VERSION 5.00
Object = "{55B64FA0-A16C-4C43-8608-1F302BBBC189}#1.0#0"; "VISTASUITEXE.ocx"
Begin VB.Form frm_splash 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "XP ListBox Demo"
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_skin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaProgressBar OsenXPProgressBar1 
      Height          =   195
      Left            =   1410
      TabIndex        =   0
      Top             =   1020
      Width           =   2175
      _ExtentX        =   3836
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
      BrushStyle      =   0
      Color           =   2871848
      Max             =   50
      Value           =   100
      ColorScheme     =   0
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

    FadeIn Me.hWnd
    OsenXPProgressBar1.Left = 94
    OsenXPProgressBar1.Top = 68
    Me.OsenXPProgressBar1.StartSearch
    WaitTimes 3000
    Me.OsenXPProgressBar1.StopSearch
    FadeOut Me.hWnd
    Unload Me
    
End Sub

Private Sub Form_Load()

    CreateFormSkin Me, LoadPicture(App.Path & "\resources\mynwind2005.bmp"), vbWhite

End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

    MoveForm hWnd, Button

End Sub


















