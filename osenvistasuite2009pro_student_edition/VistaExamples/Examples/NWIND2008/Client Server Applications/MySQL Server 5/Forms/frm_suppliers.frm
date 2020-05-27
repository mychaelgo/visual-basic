VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frm_suppliers 
   BackColor       =   &H00FFDBBF&
   BorderStyle     =   0  'None
   Caption         =   "Supplier Properties"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   Icon            =   "frm_suppliers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   379
   StartUpPosition =   1  'CenterOwner
   Begin VistaSuitePro.MyADODC MyADODC1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   14
      Top             =   6315
      Width           =   5685
      _ExtentX        =   10028
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
      ShowPrinterButton=   0   'False
      ShowGriper      =   0   'False
      AutoConfirmBeforeDelete=   -1  'True
      CaptionWidth    =   80
      WindowColor     =   3
   End
   Begin VistaSuitePro.OsenVistaTab OsenXPTab1 
      Height          =   5775
      Left            =   150
      TabIndex        =   13
      Top             =   480
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   10186
      TabHeight       =   22
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FrameColor      =   12164479
      MaskColor       =   16711935
      SelectedTab     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   1
      BackColorParent =   16767935
      TabWidth1       =   55
      TabText1        =   "General"
      TabEnabled1     =   -1  'True
      TabVisible1     =   0   'False
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   570
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         Required        =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Supplier Id:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   960
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         Required        =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Company Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   2
         Left            =   150
         TabIndex        =   2
         Top             =   1350
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Contact Name:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   3
         Left            =   150
         TabIndex        =   3
         Top             =   1740
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Contact Title:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   645
         Index           =   4
         Left            =   150
         TabIndex        =   4
         Top             =   2130
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   1138
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Address:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   5
         Left            =   150
         TabIndex        =   5
         Top             =   2850
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "City:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   6
         Left            =   150
         TabIndex        =   6
         Top             =   3270
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Region:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   7
         Left            =   150
         TabIndex        =   7
         Top             =   3660
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Postal Code:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   8
         Left            =   150
         TabIndex        =   8
         Top             =   4050
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Country:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   9
         Left            =   150
         TabIndex        =   9
         Top             =   4440
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Phone:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   10
         Left            =   150
         TabIndex        =   10
         Top             =   4830
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Fax:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
      Begin VistaSuitePro.OsenVistaTextBox txtData 
         Height          =   330
         Index           =   11
         Left            =   150
         TabIndex        =   11
         Top             =   5220
         Width           =   5100
         _ExtentX        =   8996
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
         AutoTab         =   -1  'True
         LabelAlignment  =   2
         LabelCaption    =   "Homepage:"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelForeColor  =   8388608
         LabelWidth      =   80
         LabelStyle      =   2
      End
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
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
      Caption         =   "Supplier Properties"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      AllowFadeIn     =   -1  'True
      WindowColor     =   3
   End
End
Attribute VB_Name = "frm_suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 
    Me.OsenXPForm1.Init Me

    If IsNew Then
    
        ' Open recordset
        MyADODC1.OpenMySQLTable "suppliers", "supplierid", MyCN, "where supplierid=-1", txtdata
        
        ' Set MyADODC action --> Addnew Button CLick
        MyADODC1.SendAction 5 ' AddNew
        
    Else
    
        ' Open recordset
        MyADODC1.OpenMySQLTable "suppliers", "supplierid", MyCN, "where supplierid=" & KeyValue, txtdata
        
        ' Set MyADODC action --> Update Button CLick
        MyADODC1.SendAction 6 ' Update
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    frmMain.RefreshView
    frmMain.Show
End Sub





