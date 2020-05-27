VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.MDIForm MAIN 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MAIN"
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9330
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "MAIN.frx":0ECA
   MousePointer    =   99  'Custom
   Picture         =   "MAIN.frx":11D4
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   9270
      TabIndex        =   0
      Top             =   5850
      Width           =   9330
      Begin VB.Frame Frame1 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2775
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15480
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash hotel 
         Height          =   5055
         Left            =   0
         TabIndex        =   2
         Top             =   -1080
         Width           =   15255
         _cx             =   4221212
         _cy             =   4203220
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2280
      Top             =   720
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "MAIN.frx":1E3E6
      Top             =   2040
   End
   Begin VB.Menu master 
      Caption         =   "&Master"
      Begin VB.Menu USR 
         Caption         =   "&NEW_USER"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu ADDLAUNDRY 
         Caption         =   "&NEW_LAUNDRY"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu ADDREST 
         Caption         =   "&NEW_FOOD"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu AddRoom 
         Caption         =   "&NEW_ROOM"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu PWD 
         Caption         =   "&CHANGE PASSWORD"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu LOG 
         Caption         =   "&LOGOFF"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu GRS 
         Caption         =   "-"
      End
      Begin VB.Menu EXT 
         Caption         =   "&EXIT"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu TRS 
      Caption         =   "TRANSACTION"
      Begin VB.Menu RES 
         Caption         =   "&RESERVATION"
         Shortcut        =   {F1}
      End
      Begin VB.Menu GUEST 
         Caption         =   "&GUEST"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnCHECK_IN 
         Caption         =   "&CHECK_IN"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnCHECK_OUT 
         Caption         =   "&CHECK_OUT"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu SRV 
      Caption         =   "&SERVICE"
      Begin VB.Menu dRY 
         Caption         =   "&LAUNDRY"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Rest 
         Caption         =   "&RESTAURANT"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu Rpt 
      Caption         =   "&REPORT"
      Begin VB.Menu Rport 
         Caption         =   "&SEARCH"
         Shortcut        =   {F7}
      End
      Begin VB.Menu HStatic 
         Caption         =   "&Hotel_Statistic"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu ab 
      Caption         =   "&ABOUT"
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim g As Byte
Dim i As Byte



Private Sub ab_Click()
about.Show
End Sub

Private Sub ADDLAUNDRY_Click()
NEW_LAUNDRY.Show
End Sub

Private Sub ADDREST_Click()
New_Food.Show
End Sub

Private Sub AddRoom_Click()
New_ROOM.Show
End Sub

Private Sub dRY_Click()
laundry.Show
End Sub

Private Sub EXT_Click()
pesan = MsgBox("Are You Sure?", vbQuestion + vbYesNo, "TurnOff")
If pesan = vbYes Then Timer2.Enabled = True
End Sub


Private Sub GUEST_Click()
INFORMATION_GUEST.Show
End Sub

Private Sub HStatic_Click()
HoTEL_staTIstics.Show
End Sub

Private Sub LOG_Click()
Me.Hide
pesan = MsgBox("Are You Sure?", vbQuestion + vbYesNo, "LogOff")
If pesan = vbYes Then
    LOGIN.Show
Else
    MAIN.Show
End If
End Sub



Private Sub MDIForm_Load()
splashHidup
If LeveL = "User" Then
    ADDLAUNDRY.Enabled = False
    ADDREST.Enabled = False
     USR.Enabled = False
End If

g = 255
 Timer1.Enabled = True
    Timer1.Interval = 20
    MakeTaskbarTransparent MAIN.hWnd, 0
Skin1.LoadSkin App.Path & "\Document\paper.skn"
    Skin1.ApplySkin hWnd
    
    hotel.Movie = App.Path & "\Document\hotel1.SWF"
hotel.Play
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Dim x As VbMsgBoxResult
'
'x = MsgBox("are you sure?", vbYesNo + vbQuestion, "confirm")
'If x = vbYes Then
'  Timer2.Enabled = True
'Else
'   Cancel = True
'End If
Call EXT_Click
End Sub

Private Sub mnCHECK_IN_Click()
CHECK_IN.Show
End Sub

Private Sub mnCHECK_OUT_Click()
Check_OUT.Show
End Sub

Private Sub PWD_Click()
ChangePwd.Show
End Sub


Private Sub RES_Click()
ReserVation.Show
End Sub

Private Sub Rest_Click()
Restaurant.Show
End Sub

Private Sub Rport_Click()
Report.Show
End Sub

Private Sub Timer1_Timer()

    For i = 0 To 255
        MakeTaskbarTransparent MAIN.hWnd, i
        If i = 254 Then
            MakeTaskbarTransparent MAIN.hWnd, 255
            Timer1.Enabled = False
            Exit Sub
        End If
    Next

End Sub

Private Sub Timer2_Timer()
On Error Resume Next

If g >= 101 Then
TranslucentForm Me, g
Else
If g <= 0 Then



TranslucentForm Me, 5
TranslucentForm Me, 3
TranslucentForm Me, 1
TranslucentForm Me, 0
Unload Me
Timer2.Enabled = False

End If
End If
g = g - 1
End Sub



Private Sub USR_Click()
New_USer.Show
End Sub
