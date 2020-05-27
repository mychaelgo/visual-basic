VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frmMDI 
   BackColor       =   &H00EAF7F7&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9345
   ClientLeft      =   3105
   ClientTop       =   1320
   ClientWidth     =   12555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   623
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   837
   ShowInTaskbar   =   0   'False
   Begin VistaSuitePro.MyContainerCtl picMain 
      Align           =   1  'Align Top
      Height          =   5445
      Left            =   0
      TabIndex        =   3
      Top             =   795
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   9604
      BackColor       =   8421504
      ScaleWidth      =   837
      ScaleHeight     =   363
      ClientWidth     =   20025
      ClientHeight    =   12945
      OffsetLR        =   4
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   12690
         Left            =   60
         ScaleHeight     =   12690
         ScaleWidth      =   19650
         TabIndex        =   4
         Top             =   0
         Width           =   19650
      End
   End
   Begin VistaSuitePro.OsenVistaStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   8925
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   741
      BackColor       =   14936810
      ForeColor       =   0
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   3
      HaveXPForm      =   -1  'True
      WindowColor     =   2
      PWidth1         =   100
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "OsenXPSuite 2006"
      pTextAlignment1 =   0
      PanelPicAlignment1=   0
      PWidth2         =   100
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Enterprise Edition"
      pTextAlignment2 =   0
      PanelPicAlignment2=   0
      PWidth3         =   108
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Version 11.24.0.1"
      pTextAlignment3 =   0
      PanelPicAlignment3=   0
      GradientColor1  =   16777215
      GradientColor2  =   10522143
   End
   Begin VistaSuitePro.OsenVistaHookMenu OsenXPHookMenu1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   688
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GripperLeft     =   8
      MCountMenu      =   1
      XMenuA1         =   "File "
      XMenuACS1       =   ""
      XMenuC1         =   "mnuFile"
      XMenuE1         =   -1  'True
      XMenuH1         =   0   'False
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
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
      Caption         =   "Form1"
      TitleTop        =   7
      BorderStyle     =   1
      UseDefaultTheme =   0   'False
      WindowColor     =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnusp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenXPSuite 2006 - OsenXPHookMenu sample                             *'
'*  Copyright (c) 2006 Osen Kusnadi.                                     *'
'*                                                                       *'
'*  This file is part of the OsenXPSuite 2006 sample applications.       *'
'*  The source code in this file is only intended as a supplement to     *'
'*  OsenXPSuite 2006 documentation, and is provided "as is", without     *'
'*  warranty of any kind, either expressed or implied.                   *'
'*************************************************************************'

Private Sub Form_Load()
    
    ' We recommend that you call the Init method in the main entry point of form object
    ' (which is specified in Form_Load event OR Form_Initialize)
    Me.OsenXPForm1.Init Me
       
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Frm_Main
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        picMain.Height = Me.ScaleHeight - 83
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Frm_Main.Show
    
End Sub

Private Sub mnuOpen_Click()
    Frm_Main.Show
End Sub

Private Sub mnuSave_Click()
    Frm_Main.Show
End Sub






















