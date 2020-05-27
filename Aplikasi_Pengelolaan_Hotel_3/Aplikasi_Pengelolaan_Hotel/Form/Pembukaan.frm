VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form Pembukaan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   4800
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7215
      _cx             =   4207030
      _cy             =   4203855
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
End
Attribute VB_Name = "Pembukaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Byte

Option Explicit
Dim g As Byte
Private Sub Form_Load()
about.Movie = App.Path & "\Document\Pembukaan.swf"
about.Play
g = 255
End Sub

Private Sub Timer1_Timer()
x = x + 1
    If x >= 50 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
       
    End If
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
Timer2.Enabled = False
frmTip.Show
Unload Me
Me.Hide

 
End If
End If
g = g - 1
End Sub
