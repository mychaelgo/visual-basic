VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   0  'None
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6855
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FadeIn Me.hWnd
    WaitTimes 2500
    FadeOut Me.hWnd
    Unload Me
    login.Show
    FadeIn login.hWnd
End Sub

Private Sub Form_Load()
    CreateFormSkin Me, LoadPicture(App.Path & "\in.bmp"), vbWhite
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm hWnd, Button
End Sub
