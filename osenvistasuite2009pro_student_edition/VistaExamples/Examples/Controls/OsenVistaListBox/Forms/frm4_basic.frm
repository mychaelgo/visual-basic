VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "osenvistasuite2009.ocx"
Begin VB.Form frm4_basic 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Multiple Checkbox"
   ClientHeight    =   5805
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm4_basic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   2  'CenterScreen
   Begin VistaSuitePro.OsenVistaButton OsenXPButton1 
      Height          =   405
      Left            =   3060
      TabIndex        =   2
      Top             =   5220
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   714
      Caption         =   "Populate List"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MPTR            =   0
      MICON           =   "frm4_basic.frx":000C
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      BinaryImageNormal=   "frm4_basic.frx":0028
      BinaryImageOver =   "frm4_basic.frx":0040
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   4545
      Left            =   210
      TabIndex        =   1
      Top             =   570
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   8017
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      SelectModeStyle =   5
      ShowHeader      =   -1  'True
      HeaderFormatString=   $"frm4_basic.frx":0058
      Columns         =   10
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      MaxAllColumnWidth=   755
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      ForeColorSelected=   16576
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "frm4_basic.frx":00F6
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
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
      Caption         =   "Multiple Checkbox"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
   End
End
Attribute VB_Name = "frm4_basic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
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
    
    OsenXPButton1_Click
End Sub

Private Sub OsenXPButton1_Click()

    With Me.OsenXPListBox1
        .Clear
        .LockUpdate = True
        For I = 1 To 50
            .AddItem "Table " & I & vbTab & GetRandomValue & vbTab & GetRandomValue & vbTab & GetRandomValue & vbTab & GetRandomValue & vbTab & GetRandomValue & vbTab & GetRandomValue & vbTab & GetRandomValue & vbTab & GetRandomValue & vbTab
        Next I
        .LockUpdate = False
    End With

End Sub

Private Function GetRandomValue() As Long

    GetRandomValue = IIf(Rnd() * 10 > 5, 1, 0)

End Function

Private Sub OsenXPListBox1_CellClick(lrow As Long, iCol As Integer, lLeft As Long, lTop As Long, lWidth As Long, LHeight As Long, Value As String)
    If iCol Then
        MsgBoxGT "Allow " & Me.OsenXPListBox1.ColumnText(iCol + 1) & " to " & Me.OsenXPListBox1.ColumnValue(0) & " = " & Value, vbInformation, "Value", 2
    End If
End Sub























