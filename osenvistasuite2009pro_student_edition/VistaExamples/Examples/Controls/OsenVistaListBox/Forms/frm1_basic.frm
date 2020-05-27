VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm1_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Basic ListBox Demo #1"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   Icon            =   "frm1_basic.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   228
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaFrame OsenXPFrame1 
      Height          =   1785
      Left            =   210
      TabIndex        =   3
      Top             =   5010
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3149
      Caption         =   "Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BorderColor     =   12017457
      BinaryImage     =   "frm1_basic.frx":038A
      GradientColor2  =   15779735
      Begin VistaSuitePro.OsenVistaComboBox cboMS 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1320
         Width           =   2685
         _ExtentX        =   4736
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
         ForeColor       =   0
         ComboStyle      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSGL            =   -1  'True
         LLID            =   3
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         HoverSelection  =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFontColor =   0
         ASURC           =   0   'False
         TextColumn      =   0
         Required        =   0   'False
         Unicode         =   0   'False
         DataList        =   "Standard|Dither|G_Vertical|G_Horizontal"
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
      Begin VistaSuitePro.OsenVistaComboBox cboMode 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   630
         Width           =   1545
         _ExtentX        =   2725
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
         ForeColor       =   0
         ComboStyle      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSGL            =   -1  'True
         LLID            =   1
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         HoverSelection  =   -1  'True
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFontColor =   0
         ASURC           =   0   'False
         TextColumn      =   0
         Required        =   0   'False
         Unicode         =   0   'False
         DataList        =   "Single|Multiple"
         BorderColor     =   12164479
         BorderColorOver =   12164479
      End
      Begin VistaSuitePro.OsenVistaCheckBox chkCheck 
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Caption         =   "Checked"
      End
      Begin VistaSuitePro.OsenVistaCheckBox chkGrid 
         Height          =   285
         Left            =   1500
         TabIndex        =   5
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Value           =   1
         Caption         =   "Show Gridlines"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Mode Style:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Mode:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   690
         Width           =   945
      End
   End
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   4470
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   767
      Caption         =   "Populate 100,000 Items"
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
      MICON           =   "frm1_basic.frx":03A2
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      BinaryImageNormal=   "frm1_basic.frx":03BE
      BinaryImageOver =   "frm1_basic.frx":03D6
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   3795
      Left            =   180
      TabIndex        =   1
      Top             =   570
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   6694
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
      PicturePosition =   2
      TransparencyLevel=   22
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "frm1_basic.frx":03EE
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3420
      _ExtentX        =   6033
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
      Caption         =   "Basic ListBox Demo #1"
      TitleTop        =   7
      icon            =   "frm1_basic.frx":0406
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
   End
End
Attribute VB_Name = "frm1_basic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************'
'*  OsenXPSuite 2006 - OsenXPListBox sample                              *'
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

Private Sub OsenXPButton1_Click()
    Dim L As Long
    Dim z   As Long
    
    OsenXPListBox1.Clear
    
    OsenXPListBox1.LockUpdate = True
    
    z = GTick
    
    For L = 1 To 10000
        Me.OsenXPListBox1.AddItem "Osen Kusnadi: " & L
    Next
    
    z = GTick - z
    
    OsenXPListBox1.LockUpdate = False
    DoEvents
    
    MsgBoxGT "10,000 Items was successful inserted" & vbCrLf & z & " ms taken", vbInformation
    
    
End Sub

Private Sub cboMode_Click()
    Me.OsenXPListBox1.SelectMode = cboMode.ListIndex
End Sub

Private Sub cboMS_Click()
    Me.OsenXPListBox1.SelectModeStyle = cboMS.ListIndex
End Sub

Private Sub chkCheck_Click()
    Me.OsenXPListBox1.Style = chkCheck.Value
End Sub

Private Sub chkGrid_Click()
    Me.OsenXPListBox1.ShowGridLines = chkGrid.Value
End Sub























