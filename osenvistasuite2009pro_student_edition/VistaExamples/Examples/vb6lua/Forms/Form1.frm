VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Running Lua Script in VB6"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.OsenVistaButton OsenVistaButton1 
      Height          =   375
      Left            =   7830
      TabIndex        =   4
      Top             =   4800
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      Caption         =   "Test (Run Script)"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "Form1.frx":0000
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      Style           =   1
      BinaryImageNormal=   "Form1.frx":001C
      BinaryImageOver =   "Form1.frx":0034
   End
   Begin VistaSuitePro.OsenVistaTextBox txtVar 
      Height          =   330
      Left            =   300
      TabIndex        =   2
      Top             =   4800
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      Text            =   "X"
      Alignment       =   2
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
      ForeColor       =   0
      BorderColor     =   12624503
      BorderColorOver =   12624503
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      LabelBackColor  =   16767935
      LabelCaption    =   "VarName:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelForeColor  =   16576
      LabelWidth      =   65
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtScript 
      Height          =   4095
      Left            =   300
      TabIndex        =   1
      Top             =   600
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   7223
      Text            =   "TextBox1"
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
      BorderColor     =   12624503
      MultiLine       =   -1  'True
      BackColor       =   0
      BorderColorOver =   12624503
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      BackColorOver   =   0
      ForeColorOver   =   65280
      LabelBackColor  =   16767935
      LabelCaption    =   "Lua Script:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelForeColor  =   33023
      LabelStyle      =   1
   End
   Begin VistaSuitePro.OsenVistaForm OsenVistaForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Running Lua Script in VB6"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaTextBox txtReturn 
      Height          =   330
      Left            =   3990
      TabIndex        =   3
      Top             =   4830
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   582
      Alignment       =   2
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
      ForeColor       =   0
      BorderColor     =   12624503
      BorderColorOver =   12624503
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonGradient  =   -1  'True
      LabelAlignment  =   2
      LabelBackColor  =   16767935
      LabelCaption    =   "VarValue (Return):"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelForeColor  =   12582912
      LabelWidth      =   120
      LabelStyle      =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    OsenVistaForm1.Init Me
    Me.txtScript.LoadFile App.Path & "\test.txt"
    
    'lua starting ...
    LuaOpen
    
    '-- simple lua script example
    Debug.Print LuaToString("_VERSION")
    
    '-- set variable and value ...
    LuaDoString ("A=10")
    
    ' get value from existing variable
    Debug.Print "A="; LuaToString("A")
    
    '-- math function
    LuaDoString ("B=77 C=A*B") ' c=10*77
    Debug.Print "C="; LuaToString("C")
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'lua stop
    LuaClose
    
End Sub

Private Sub OsenVistaButton1_Click()
    txtReturn.Text = LuaToString2(txtVar.Text, txtScript.Text)
End Sub


