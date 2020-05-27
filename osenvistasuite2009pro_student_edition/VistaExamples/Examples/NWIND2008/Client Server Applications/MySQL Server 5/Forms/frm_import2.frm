VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_import 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Configuration..."
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   Icon            =   "frm_import2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   214
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton cmdImport 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2550
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   661
      Caption         =   "&Load Sample Data"
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
      MICON           =   "frm_import2.frx":038A
      PICN            =   "frm_import2.frx":04EC
      UMCOL           =   -1  'True
      Style           =   1
      BinaryImageNormal=   "frm_import2.frx":0886
      BinaryImageOver =   "frm_import2.frx":089E
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Configuration..."
      TitleTop        =   7
      icon            =   "frm_import2.frx":08B6
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
      Height          =   1995
      Left            =   240
      TabIndex        =   1
      Top             =   510
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   3519
      Caption         =   "Configuration MySQL  Connection"
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
      BinaryImage     =   "frm_import2.frx":0C50
      GradientColor1  =   16773607
      GradientColor2  =   16768452
      Begin VistaSuitePro.OsenVistaTextBox txtHost 
         Height          =   285
         Left            =   1140
         TabIndex        =   2
         Top             =   450
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Text            =   "localhost"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaTextBox txtUser 
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   810
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Text            =   "root"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaTextBox txtPwd 
         Height          =   285
         Left            =   1140
         TabIndex        =   4
         Top             =   1170
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
         PasswordChar    =   "*"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaTextBox txtPort 
         Height          =   285
         Left            =   1140
         TabIndex        =   5
         Top             =   1530
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Text            =   "3306"
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
         Value           =   3306
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hostname:"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   1560
         Width           =   360
      End
   End
End
Attribute VB_Name = "frm_import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImport_Click()

    On Error GoTo Err_X
    
    Dim SQL         As String
    
    ' get all query from txt file
    SQL = App.Path & "\nwind2008.sql"
    
    ' if file not empty
    If FileLen(SQL) Then
        
        ' connect to specify mysql server ...
        MyCN.OpenConnection txtHost, txtUser, txtPwd, txtPort.Value
        
        ' check connection status
        If MyCN.State Then ' connectioin is active
        
            ' execute all query into active connection
            MyCN.Restore SQL
            
            
            mStrSQL = "UPDATE connection_info set host='" & txtHost & "', uid='" & txtUser & "', pwd='" & txtPwd & "', port=" & txtPort & ", active=1, dbname='" & MyCN.DBName & "'"
            
            sCN.Execute mStrSQL
            
            ' display message
            MsgBoxGT "All query/data was successful execute into " & MyCN.SQL_Result("select database()") & "@" & txtHost & vbCrLf & vbCrLf & _
                    "ConnectionID : " & MyCN.ConnectionID & vbCrLf & _
                    "Client version : " & MyCN.ClientVersion & vbCrLf & _
                    "Server Info: " & MyCN.HostInfo & " (" & MyCN.ServerVersionInfo & ")", vbInformation + vbSystemModal
                    
            
            Unload Me
            
        Else
            MsgBoxGT "Could not connect to the server ...", vbExclamation + vbSystemModal, "Connection failed", 3
        End If
        
       
    End If
    
    Exit Sub
    
Err_X:

    ' Display error message
    MsgBoxGT Err.Description, vbCritical, "Error", 5
    
    On Error GoTo 0
    
End Sub

Private Sub Form_Activate()
    Me.OsenXPForm1.FormOnTop True

End Sub

Private Sub Form_Load()
    
    Me.OsenXPForm1.Init Me
    
    Dim DATA() As String
    
    sCN.GetArrayFromSQL "select * from connection_info", DATA
    
    txtHost.Text = DATA(0)
    txtUser.Text = DATA(1)
    txtPwd.Text = DATA(2)
    txtPort.Text = DATA(3)
    
    
End Sub

















