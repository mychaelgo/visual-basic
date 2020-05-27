VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_Category 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Categories"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   Icon            =   "frm_Category.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyADODC MyADODC1 
      Height          =   555
      Left            =   210
      TabIndex        =   6
      Top             =   3600
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   979
      GradientColor1  =   16777215
      GradientColor2  =   12752244
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
      BeginProperty FontButton {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BorderStyle     =   6
      Gradient        =   -1  'True
      ShowFindButton  =   0   'False
      ShowFilterButton=   0   'False
      ShowRefreshButton=   0   'False
      ShowPrinterButton=   0   'False
      AutoConfirmBeforeDelete=   -1  'True
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaPicture picX 
      Height          =   2085
      Left            =   3870
      TabIndex        =   5
      Top             =   1440
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   3678
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
      FieldName       =   "picture"
      BorderColor     =   14854529
      GradientColor2  =   16310477
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "frm_Category.frx":058A
   End
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   1508
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
      Picture         =   "frm_Category.frx":05A2
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   16310477
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "detail of a categories is as follow ..."
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Categories records"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "frm_Category.frx":20F4
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
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
      Caption         =   "Categories"
      TitleTop        =   7
      icon            =   "frm_Category.frx":210C
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   330
      Index           =   0
      Left            =   210
      TabIndex        =   2
      Top             =   1470
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   582
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
      Enabled         =   0   'False
      Locked          =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      LabelBackColor  =   16767935
      LabelCaption    =   "Category ID:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   80
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   330
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   1890
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   582
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      LabelBackColor  =   16767935
      LabelCaption    =   "Category Name:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   80
      LabelStyle      =   2
   End
   Begin VistaSuitePro.OsenVistaTextBox txtData 
      Height          =   1200
      Index           =   2
      Left            =   210
      TabIndex        =   4
      Top             =   2340
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   2117
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
      MultiLine       =   -1  'True
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOver   =   12648447
      LabelBackColor  =   15790320
      LabelCaption    =   "Description:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelForeColor  =   8388608
      LabelStyle      =   1
      LabelGradient   =   -1  'True
      LabelGradientColor1=   16767935
      LabelGradientColor2=   16767935
   End
End
Attribute VB_Name = "frm_Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.OsenXPForm1.Init Me

    ' Open recordset and Bind txtdata and PicX into MyADODC1
    If IsNew Then
        MyADODC1.OpenMySQLTable "categories", "categoryid", MyCN, "where categoryid=-1", txtdata, picX
    Else
        MyADODC1.OpenMySQLTable "categories", "categoryid", MyCN, "where categoryid=" & KeyValue, txtdata, picX
    End If

    If IsNew Then
        MyADODC1.SendAction 5 ' Addnew
    Else
        MyADODC1.SendAction 6 ' Edit
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.RefreshView
    frmMain.Show
On Error GoTo 0
End Sub





