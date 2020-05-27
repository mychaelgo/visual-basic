VERSION 5.00
Object = "{EE17B266-A61D-48F0-BB3E-5C4EC9EE2D1D}#1.1#0"; "osenxpsuite2009.ocx"
Begin VB.Form satuan 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Input Satuan"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin OSENXPSUITE2009OCX.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Input Satuan"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      BorderStyle     =   2
      ToolTipClose    =   "Tutup"
      UseDefaultTheme =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Satuan"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   510
   End
End
Attribute VB_Name = "satuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
