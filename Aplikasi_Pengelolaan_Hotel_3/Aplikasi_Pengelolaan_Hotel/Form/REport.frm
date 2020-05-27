VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{A5B7E513-C349-4AF2-8648-C419AE687AEA}#2.0#0"; "lvButtons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Report 
   Caption         =   "REPORT"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "REport.frx":0000
   ScaleHeight     =   2895
   ScaleWidth      =   4215
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker TgLFor 
      Height          =   375
      Left            =   2280
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483642
      CalendarForeColor=   65280
      CalendarTitleBackColor=   0
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   0
      Format          =   19791873
      CurrentDate     =   39484
   End
   Begin MSComCtl2.DTPicker tgLFrom 
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483642
      CalendarForeColor=   65280
      CalendarTitleBackColor=   0
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   0
      Format          =   19791873
      CurrentDate     =   39484
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "REport.frx":10DC0
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox chkBetween 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   255
   End
   Begin LvButtons.lvButtons_H btnProses 
      Height          =   405
      Left            =   1800
      TabIndex        =   16
      Top             =   2400
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
      Caption         =   "&PROSES"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "REport.frx":10E32
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4800
      OleObjectBlob   =   "REport.frx":10EA0
      Top             =   480
   End
   Begin VB.TextBox tsearch 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox cmbRecord 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "REport.frx":110D4
      Left            =   1680
      List            =   "REport.frx":110D6
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ComboBox cmbField 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "REport.frx":110D8
      Left            =   1680
      List            =   "REport.frx":110FA
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker tgl 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483642
      CalendarForeColor=   65280
      CalendarTitleBackColor=   0
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   0
      Format          =   19791873
      CurrentDate     =   39484
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "REport.frx":1115B
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "REport.frx":111C1
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "REport.frx":11231
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "REport.frx":11297
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "REport.frx":11307
      TabIndex        =   15
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel lFrom 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "REport.frx":1136D
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel ltik 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   1800
      OleObjectBlob   =   "REport.frx":113E3
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel Lfor 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "REport.frx":11449
      TabIndex        =   21
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel lTik2 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   1800
      OleObjectBlob   =   "REport.frx":114BD
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash about 
      Height          =   495
      Left            =   720
      TabIndex        =   25
      Top             =   0
      Width           =   2655
      _cx             =   4198987
      _cy             =   4195177
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "search"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   825
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   600
      Y2              =   1560
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "field"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   135
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   1560
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   4080
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub hidup()
chkBetween.Enabled = True
tsearch.Visible = False
tgl.Visible = True
End Sub

Sub mati()
chkBetween.Enabled = False
Me.Height = 3405
tsearch.Visible = True
tgl.Visible = False

End Sub

Private Sub btnProses_Click()

If cmbField.ListIndex = 0 And cmbRecord = "" Then

RCheckIN.DataControl1.Source = _
   "SELECT *from checkin"""
 RCheckIN.Show
Else
If cmbField.ListIndex = 0 Then


    If cmbRecord.ListIndex = 8 Or cmbRecord.ListIndex = 14 Or cmbRecord.ListIndex = 15 Or cmbRecord.ListIndex = 16 Or cmbRecord.ListIndex = 17 Or cmbRecord.ListIndex = 18 Or cmbRecord.ListIndex = 19 Or cmbRecord.ListIndex = 20 Or cmbRecord.ListIndex = 21 Then
    
    RCheckIN.DataControl1.Source = _
        "Select * from checkin where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
          If cmbRecord.ListIndex = 10 Or cmbRecord.ListIndex = 12 Then
              RCheckIN.DataControl1.Source = _
                    "Select * from checkin where " & cmbRecord & "=#" & tgl & "#"
                            
                            If chkBetween.Value Then
                                        RCheckIN.DataControl1.Source = _
                                        "Select * from CheckIN where " & cmbRecord & " Between #" & tgLFrom & "# and #" & TgLFor & "#"
                                End If
                                
              
            Else
          
                If cmbRecord.ListIndex = 11 Then
                
                RCheckIN.DataControl1.Source = _
                    "Select * from checkin where " & cmbRecord & "=#" & tsearch & "#"
                Else
                    If cmbRecord.ListIndex = 0 Or cmbRecord.ListIndex = 1 Or cmbRecord.ListIndex = 2 Or cmbRecord.ListIndex = 3 Or cmbRecord.ListIndex = 4 Or cmbRecord.ListIndex = 5 Or cmbRecord.ListIndex = 6 Or cmbRecord.ListIndex = 7 Or cmbRecord.ListIndex = 9 Or cmbRecord.ListIndex = 13 Then
                    RCheckIN.DataControl1.Source = _
                    "Select * from checkin where " & cmbRecord & " LIKE '%" & tsearch & "%'"
                    End If
                    
                End If
            End If
    End If
    RCheckIN.Show
End If
 
End If
   

If cmbField.ListIndex = 1 And cmbRecord = "" Then
RCheckIN.DataControl1.Source = _
   "SELECT *from checkout"""
RCheckOut.Show
Else
If cmbField.ListIndex = 1 Then
   If cmbRecord.ListIndex = 9 Or cmbRecord.ListIndex = 15 Or cmbRecord.ListIndex = 16 Or cmbRecord.ListIndex = 17 Or cmbRecord.ListIndex = 18 Or cmbRecord.ListIndex = 19 Or cmbRecord.ListIndex = 20 Or cmbRecord.ListIndex = 21 Or cmbRecord.ListIndex = 22 Then
    
    RCheckOut.DataControl1.Source = _
        "Select * from checkout where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
    
            If cmbRecord.ListIndex = 11 Or cmbRecord.ListIndex = 13 Then
               
                RCheckOut.DataControl1.Source = _
                    "Select * from checkout where " & cmbRecord & "=#" & tgl & "#"
                        
                        If chkBetween.Value Then
                                RCheckOut.DataControl1.Source = _
                                "Select * from CheckOut where " & cmbRecord & " Between #" & tgLFrom & "# and #" & TgLFor & "#"
                        End If
                        
                        
                    
                    
                    
                Else
                If cmbRecord.ListIndex = 12 Then
                
                RCheckOut.DataControl1.Source = _
                    "Select * from checkout where " & cmbRecord & "=#" & tsearch & "#"
                Else
                    If cmbRecord.ListIndex = 0 Or cmbRecord.ListIndex = 1 Or cmbRecord.ListIndex = 2 Or cmbRecord.ListIndex = 3 Or cmbRecord.ListIndex = 4 Or cmbRecord.ListIndex = 5 Or cmbRecord.ListIndex = 6 Or cmbRecord.ListIndex = 7 Or cmbRecord.ListIndex = 8 Or cmbRecord.ListIndex = 10 Or cmbRecord.ListIndex = 14 Then
                    RCheckOut.DataControl1.Source = _
                    "Select * from checkout where " & cmbRecord & " LIKE '%" & tsearch & "%'"
                End If
                End If
     End If
    End If
    RCheckOut.Show
End If
End If


If cmbField.ListIndex = 2 Then
    If cmbRecord = "" Then
        rFOOd.DataControl1.Source = _
        "Select * from Food"
    Else
        If cmbRecord.ListIndex = 4 Then
            rFOOd.DataControl1.Source = _
            "select * from food where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
            rFOOd.DataControl1.Source = _
            "select * from food where " & cmbRecord & " LIKE '%" & tsearch & "%'"
        
  
    End If
    End If
rFOOd.Show
End If



If cmbField.ListIndex = 3 Then
    If cmbRecord = "" Then
        rGuest.DataControl1.Source = _
        "Select * from guest"
    Else
     If cmbRecord.ListIndex = 5 Or cmbRecord.ListIndex = 11 Then
        rGuest.DataControl1.Source = _
            "select * from guest where " & cmbRecord & " LIKE '%" & tsearch & "%'"
     
     Else
        If cmbRecord.ListIndex = 0 Then
          rGuest.DataControl1.Source = _
                    "Select * from guest where " & cmbRecord & "=#" & tgl & "#"
                        
                        If chkBetween.Value Then
                                rGuest.DataControl1.Source = _
                                "Select * from Guest where " & cmbRecord & " Between #" & tgLFrom & "# and #" & TgLFor & "#"
                        End If
                        
     
     Else
            rGuest.DataControl1.Source = _
            "select * from guest where " & cmbRecord & " LIKE '%" & tsearch & "%'"
     End If
     End If
     End If

rGuest.Show
End If


If cmbField.ListIndex = 4 Then
    If cmbRecord = "" Then
        rLaundry.DataControl1.Source = _
        "Select * from itemlaundry"
    Else
        If cmbRecord.ListIndex = 2 Or cmbRecord.ListIndex = 3 Then
             rLaundry.DataControl1.Source = _
            "select * from itemlaundry where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
             rLaundry.DataControl1.Source = _
            "select * from itemlaundry where " & cmbRecord & " LIKE '%" & tsearch & "%'"
        
  
    End If
    End If
rLaundry.Show
End If


If cmbField.ListIndex = 5 Then
    If cmbRecord = "" Then
        rLaundrytrans.DataControl1.Source = _
        "Select * from Laundry"
    Else
        If cmbRecord.ListIndex = 5 Or cmbRecord.ListIndex = 6 Then
            rLaundrytrans.DataControl1.Source = _
            "select * from Laundry where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
        If cmbRecord.ListIndex = 3 Then
          rLaundrytrans.DataControl1.Source = _
                    "Select * from laundry where " & cmbRecord & "=#" & tgl & "#"
                         
                        If chkBetween.Value Then
                                rLaundrytrans.DataControl1.Source = _
                                "Select * from Laundry where " & cmbRecord & " Between #" & tgLFrom & "# and #" & TgLFor & "#"
                        End If
                         
      
      
      Else
        If cmbRecord.ListIndex = 4 Then
        rLaundrytrans.DataControl1.Source = _
                    "Select * from laundry where " & cmbRecord & "=#" & tsearch & "#"
     Else
             rLaundrytrans.DataControl1.Source = _
            "select * from Laundry where " & cmbRecord & " LIKE '%" & tsearch & "%'"
     End If
    End If
    End If
    End If
rLaundrytrans.Show
End If

If cmbField.ListIndex = 6 Then
    If cmbRecord = "" Then
        rReservation.DataControl1.Source = _
        "Select * from reservation"
    Else
        If cmbRecord.ListIndex = 5 Or cmbRecord.ListIndex = 7 Or cmbRecord.ListIndex = 9 Then
            rReservation.DataControl1.Source = _
            "select * from reservation where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
        If cmbRecord.ListIndex = 0 Or cmbRecord.ListIndex = 6 Then
          rReservation.DataControl1.Source = _
                    "Select * from reservation where " & cmbRecord & "=#" & tgl & "#"
                        
                         If chkBetween.Value Then
                                rReservation.DataControl1.Source = _
                                "Select * from reservation where " & cmbRecord & " Between #" & tgLFrom & "# and #" & TgLFor & "#"
                        End If
                         
      
      
      Else
             rReservation.DataControl1.Source = _
            "select * from reservation where " & cmbRecord & " LIKE '%" & tsearch & "%'"
        
    End If
    End If
    End If
rReservation.Show
End If

If cmbField.ListIndex = 7 Then
    If cmbRecord = "" Then
        rRestaurantTrans.DataControl1.Source = _
        "Select * from Restaurant"
    Else
        If cmbRecord.ListIndex = 5 Then
            rRestaurantTrans.DataControl1.Source = _
            "select * from Restaurant where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
        If cmbRecord.ListIndex = 3 Then
          rRestaurantTrans.DataControl1.Source = _
                    "Select * from Restaurant where " & cmbRecord & "=#" & tgl & "#"
                         
                         If chkBetween.Value Then
                                rRestaurantTrans.DataControl1.Source = _
                                "Select * from false where " & cmbRecord & " Between #" & tgLFrom & "# and #" & TgLFor & "#"
                        End If
                          
                         
      Else
          If cmbRecord.ListIndex = 4 Then
          rRestaurantTrans.DataControl1.Source = _
                    "Select * from Restaurant where " & cmbRecord & "=#" & tsearch & "#"
      
      Else
             rRestaurantTrans.DataControl1.Source = _
            "select * from Restaurant where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    End If
    End If
    End If
    End If
rRestaurantTrans.Show
End If

If cmbField.ListIndex = 8 Then
    If cmbRecord = "" Then
        rRoom.DataControl1.Source = _
        "Select * from room"
    Else
        If cmbRecord.ListIndex = 0 Or cmbRecord.ListIndex = 1 Or cmbRecord.ListIndex = 3 Then
            rRoom.DataControl1.Source = _
            "select * from room where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    Else
      
            
  
             rRoom.DataControl1.Source = _
            "select * from room where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    
 
    End If
    End If
rRoom.Show
End If



If cmbField.ListIndex = 9 Then
    If cmbRecord = "" Then
        rUSER.DataControl1.Source = _
        "Select * from tuser"
    Else
        
            
  
             rUSER.DataControl1.Source = _
            "select * from tuser where " & cmbRecord & " LIKE '%" & tsearch & "%'"
    
 

    End If
rUSER.Show
End If



End Sub




Private Sub chkBetween_Click()
If chkBetween.Value Then
    TgLFor.Visible = True
    tgLFrom.Visible = True
    lFrom.Visible = True
    Lfor.Visible = True
    lTik2.Visible = True
    ltik.Visible = True
    Me.Height = 5040
Else
    TgLFor.Visible = False
    tgLFrom.Visible = False
    lFrom.Visible = False
    Lfor.Visible = False
    ltik.Visible = False
    lTik2.Visible = False
    Me.Height = 3405
End If


End Sub


Private Sub cmbField_Click()
If cmbField.ListIndex = 0 Then
        cmbRecord.Clear
        cmbRecord.AddItem "IDcheckin"
         cmbRecord.AddItem "IDguest"
          cmbRecord.AddItem "IDcard"
           cmbRecord.AddItem "Name"
            cmbRecord.AddItem "Address"
             cmbRecord.AddItem "City"
              cmbRecord.AddItem "Nationality"
               cmbRecord.AddItem "Phone"
                cmbRecord.AddItem "Roomno"
                 cmbRecord.AddItem "TypeRoom"
                  cmbRecord.AddItem "Arrival_date"
                   cmbRecord.AddItem "Arrival_time"
                     cmbRecord.AddItem "Out_date"
                      cmbRecord.AddItem "days"
                       cmbRecord.AddItem "Price"
                        cmbRecord.AddItem "Discount"
                         cmbRecord.AddItem "Service"
                          cmbRecord.AddItem "Tax"
                           cmbRecord.AddItem "Amount"
                            cmbRecord.AddItem "Laundry"
                             cmbRecord.AddItem "Restaurant"
                              cmbRecord.AddItem "Return Money"
Else
If cmbField.ListIndex = 1 Then
        cmbRecord.Clear
       cmbRecord.AddItem "IDCheckOut"
        cmbRecord.AddItem "IDcheckin"
         cmbRecord.AddItem "IDguest"
          cmbRecord.AddItem "IDcard"
           cmbRecord.AddItem "Name"
            cmbRecord.AddItem "Address"
             cmbRecord.AddItem "City"
              cmbRecord.AddItem "Nationality"
               cmbRecord.AddItem "Phone"
                cmbRecord.AddItem "Roomno"
                 cmbRecord.AddItem "TypeRoom"
                  cmbRecord.AddItem "Arrival_date"
                   cmbRecord.AddItem "Arrival_time"
                     cmbRecord.AddItem "Out_date"
                      cmbRecord.AddItem "days"
                       cmbRecord.AddItem "Price"
                        cmbRecord.AddItem "Discount"
                         cmbRecord.AddItem "Service"
                          cmbRecord.AddItem "Tax"
                           cmbRecord.AddItem "Amount"
                            cmbRecord.AddItem "Laundry"
                             cmbRecord.AddItem "Restaurant"
                              cmbRecord.AddItem "Return Money"
Else

        If cmbField.ListIndex = 2 Then
        cmbRecord.Clear
              cmbRecord.AddItem "IDFood"
                cmbRecord.AddItem "Dish"
                 cmbRecord.AddItem "Name"
                  cmbRecord.AddItem "Kinds"
                   cmbRecord.AddItem "Price"
                
Else
        If cmbField.ListIndex = 3 Then
        cmbRecord.Clear
              cmbRecord.AddItem "Arrivaldate"
                cmbRecord.AddItem "Arrivaltime"
                 cmbRecord.AddItem "Idguest"
                  cmbRecord.AddItem "IDcard"
                   cmbRecord.AddItem "Name"
                    cmbRecord.AddItem "Age"
                     cmbRecord.AddItem "Address"
                      cmbRecord.AddItem "City"
                       cmbRecord.AddItem "Nationality"
                        cmbRecord.AddItem "Sex"
                         cmbRecord.AddItem "Religion"
                          cmbRecord.AddItem "Phone"
                          
                         
                   
Else
         If cmbField.ListIndex = 4 Then
        cmbRecord.Clear
              cmbRecord.AddItem "IDitem"
                 cmbRecord.AddItem "Name"
                  cmbRecord.AddItem "Laundry"
                   cmbRecord.AddItem "DryClean"


Else
         If cmbField.ListIndex = 5 Then
        cmbRecord.Clear
              cmbRecord.AddItem "IDLAUNDRY"
                 cmbRecord.AddItem "Idguest"
                  cmbRecord.AddItem "IDitem"
                   cmbRecord.AddItem "DateTrans"
                    cmbRecord.AddItem "TimeTrans"
                     cmbRecord.AddItem "TotalQty"
                      cmbRecord.AddItem "Price"
                      


Else
        If cmbField.ListIndex = 6 Then
        cmbRecord.Clear
              cmbRecord.AddItem "ondate"
                 cmbRecord.AddItem "Idreservation"
                  cmbRecord.AddItem "IDcard"
                   cmbRecord.AddItem "name"
                    cmbRecord.AddItem "Address"
                     cmbRecord.AddItem "Phone"
                      cmbRecord.AddItem "Arrivaldate"
                       cmbRecord.AddItem "RoomNo"
                        cmbRecord.AddItem "Typeroom"
                         cmbRecord.AddItem "Confirmed"
                      


Else
        If cmbField.ListIndex = 7 Then
        cmbRecord.Clear
              cmbRecord.AddItem "IDrestaurant"
                 cmbRecord.AddItem "Idguest"
                  cmbRecord.AddItem "IDFood"
                   cmbRecord.AddItem "DateTrans"
                    cmbRecord.AddItem "TimeTrans"
                     cmbRecord.AddItem "TotalPrice"
                      
                      
Else
        If cmbField.ListIndex = 8 Then
        cmbRecord.Clear
              cmbRecord.AddItem "Roomno"
                 cmbRecord.AddItem "Status"
                  cmbRecord.AddItem "Type_Room"
                   cmbRecord.AddItem "Amount"
        

Else
    If cmbField.ListIndex = 9 Then
        cmbRecord.Clear
              cmbRecord.AddItem "Iduser"
                 cmbRecord.AddItem "Name"
                  cmbRecord.AddItem "Level"
                   cmbRecord.AddItem "NIP"
                    cmbRecord.AddItem "PWD"
Else
cmbRecord.Clear
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub








Private Sub cmbRecord_Click()
If cmbField.ListIndex = 0 Then
              If cmbRecord.ListIndex = 10 Or cmbRecord.ListIndex = 12 Then
                hidup
                 
              Else
                mati
              
              End If
Else
If cmbField.ListIndex = 1 Then
             If cmbRecord.ListIndex = 11 Or cmbRecord.ListIndex = 13 Then
                hidup
              
             Else
                mati
                
            End If
End If




If cmbField.ListIndex = 3 Then
             If cmbRecord.ListIndex = 0 Then
                hidup
             Else
                mati
            End If
End If
End If

If cmbField.ListIndex = 5 Then
             If cmbRecord.ListIndex = 3 Then
                hidup
             Else
                mati
            End If
End If

If cmbField.ListIndex = 6 Then
             If cmbRecord.ListIndex = 0 Or cmbRecord.ListIndex = 6 Then
                hidup
             Else
                mati
            End If
End If


If cmbField.ListIndex = 7 Then
             If cmbRecord.ListIndex = 3 Then
                hidup
             Else
                mati
            End If
End If





End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
Me.Top = 2000
Me.Width = 4335
Me.Left = MAIN.Width / 4
Me.Height = 3405
about.Movie = App.Path & "\Document\search.swf"
about.Play

Skin1.LoadSkin App.Path & "\Document\paper.skn"
Skin1.ApplySkin hWnd
splashHidup
tgl = Date
TgLFor = Date
tgLFrom = Date
End Sub











Private Sub TgLFor_Change()
If TgLFor < tgLFrom Then TgLFor = tgLFrom
End Sub
