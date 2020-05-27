VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form HoTEL_staTIstics 
   Caption         =   "HoTEL_staTIstics"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "HoTEL_StaTIstics.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   600
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   975
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   4095
      _cx             =   4201527
      _cy             =   4196024
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
   Begin VB.Label lTime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lDate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Guest Checkin Today"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Reservations Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Guest Checkout Today"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Reservations Confirmed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupied Rooms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Vacant Rooms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lTotCheckIN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   240
   End
   Begin VB.Label lResFalse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   3480
      TabIndex        =   4
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label LCheckOUTTODAY 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   240
   End
   Begin VB.Label lResConfim 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   240
   End
   Begin VB.Label LOccupied 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   3960
      Width           =   240
   End
   Begin VB.Label LVacant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      Top             =   4560
      Width           =   240
   End
End
Attribute VB_Name = "HoTEL_staTIstics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub NoL()
 If lTotCheckIN.Caption = "X" Then lResFalse.Caption = "0"
 If lResFalse.Caption = "X" Then lResFalse.Caption = "0"
  If lResConfim.Caption = "X" Then lResFalse.Caption = "0"
    If LCheckOUTTODAY.Caption = "X" Then LCheckOUTTODAY.Caption = "0"
        If LVacant.Caption = "X" Then LCheckOUTTODAY.Caption = "0"
            If LOccupied.Caption = "X" Then LCheckOUTTODAY.Caption = "0"
    
End Sub

Sub auto()
    If Rs.State = 1 Then Rs.Close
        Rs.Open "Select * from CheckIN", KOneKsi, 3, 3
            If Not Rs.EOF Then
                While Not Rs.EOF
                    lTotCheckIN.Caption = Val(lTotCheckIN) + 1
                Rs.MoveNext
                Wend
            End If
            
    If Rs.State = 1 Then Rs.Close
        Rs.Open "Select * from ReserVation ", KOneKsi, 3, 3
            If Not Rs.EOF Then
                If Rs.State = 1 Then Rs.Close
                     Rs.Open "Select * from ReserVation where confirmed=false", KOneKsi, 3, 3
                        While Not Rs.EOF
                            lResFalse.Caption = Val(lResFalse) + 1
                           
                        Rs.MoveNext
                        Wend
                        
                If Rs.State = 1 Then Rs.Close
                     Rs.Open "Select * from ReserVation where confirmed=true", KOneKsi, 3, 3
       
                While Not Rs.EOF
                    lResConfim.Caption = Val(lResConfim) + 1
                Rs.MoveNext
                Wend
                
            End If
            
If Rs.State = 1 Then Rs.Close
        Rs.Open "Select * from CheckIN where Out_Date=#" & Date & "#", KOneKsi, 3, 3
            If Not Rs.EOF Then
                While Not Rs.EOF
                    LCheckOUTTODAY.Caption = Val(LCheckOUTTODAY) + 1
                Rs.MoveNext
                Wend
            End If
                        
If Rs.State = 1 Then Rs.Close
        Rs.Open "Select * from Room ", KOneKsi, 3, 3
           
           If Not Rs.EOF Then
                If Rs.State = 1 Then Rs.Close
                     Rs.Open "Select * from Room where status=false", KOneKsi, 3, 3
                        While Not Rs.EOF
                            LVacant.Caption = Val(LVacant) + 1
                           
                        Rs.MoveNext
                        Wend
                        
                If Rs.State = 1 Then Rs.Close
                     Rs.Open "Select * from Room where status=true", KOneKsi, 3, 3
       
                While Not Rs.EOF
                    LOccupied.Caption = Val(LOccupied) + 1
                Rs.MoveNext
                Wend
                
            End If
            

    
End Sub



Private Sub Form_Load()
OPENDATA
about.Movie = App.Path & "\Document\HOTEL_STATISTIC.swf"
about.Play
awal Me
Me.Width = 4800
Me.Height = 5745
splashMati
 auto
 NoL
 
End Sub








Private Sub Form_Unload(Cancel As Integer)
splashHidup
End Sub

Private Sub Timer1_Timer()
lDate.Caption = Date
lTime.Caption = Time

End Sub
