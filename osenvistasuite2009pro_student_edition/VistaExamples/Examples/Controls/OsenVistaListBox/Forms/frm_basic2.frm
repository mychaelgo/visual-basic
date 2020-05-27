VERSION 5.00
Object="{BDB4D61A-DD9F-45C7-91D8-432B91ADEDF5}#1.0#0"; "osenxpsuite2007.ocx"
Begin VB.Form frm_basic2 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Advance Editable Cell"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   Icon            =   "frm_basic2.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   3  'Windows Default
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Advance Editable Cell"
      TitleTop        =   7
      icon            =   "frm_basic2.frx":038A
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      UseDefaultTheme =   0   'False
      AllowFadeIn     =   -1  'True
   End
   Begin VistaSuitePro.OsenVistaListBox OsenXPListBox1 
      Height          =   5865
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10345
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
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      ShowHeader      =   -1  'True
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
      HeaderCaption   =   "OsenXPListBox1"
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
      Begin VistaSuitePro.OsenVistaTextBox txtdata 
         Height          =   345
         Left            =   4140
         TabIndex        =   4
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         Text            =   "TextBox1"
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
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VistaSuitePro.OsenVistaComboBox cboCategories 
         Height          =   345
         Left            =   5880
         TabIndex        =   3
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Text            =   "ComboBox2"
         ComboStyle      =   1
         TextColumn      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSH             =   -1  'True
         LSGL            =   -1  'True
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFontColor =   16777215
         ASURC           =   0   'False
         TextColumn      =   1
         Required        =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   12164479
         HeaderGradientAllow=   -1  'True
      End
      Begin VistaSuitePro.OsenVistaComboBox cboSuppliers 
         Height          =   345
         Left            =   7530
         TabIndex        =   2
         Top             =   5400
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
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
         Text            =   "ComboBox1"
         ComboStyle      =   1
         TextColumn      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSH             =   -1  'True
         LSGL            =   -1  'True
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFontColor =   16777215
         ASURC           =   0   'False
         TextColumn      =   1
         Required        =   0   'False
         BorderColor     =   12164479
         BorderColorOver =   12164479
         HeaderGradientAllow=   -1  'True
      End
   End
End
Attribute VB_Name = "frm_basic2"
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
    
    ' Insert the suppliers record into cboSuppliers
    cboSuppliers.InsertItemBySQL AdoCN, "Select SupplierID,CompanyName   from suppliers", True
    
    ' hide the supplierID
    cboSuppliers.ColumnWidth(1) = 0
    
    ' Insert the categories record into cbocategories
    cboCategories.InsertItemBySQL AdoCN, "select CategoryId,CategoryName   from categories", True
    
    ' Hide the CategoryID
    cboCategories.ColumnWidth(1) = 0
    
    
    ' Open recordset
    Me.OsenXPListBox1.InsertItemBySQL AdoCN, "Select * from products", , , True
    
    ' Bind OsenXPTExtBox, OsenXPComboBox into OsenXPListbox for allowing user to edit data on runtime
    Me.OsenXPListBox1.BindObjectArray txtdata, txtdata, cboSuppliers, cboCategories, txtdata, txtdata, txtdata, txtdata, txtdata
    
    ' set column alignment
    Me.OsenXPListBox1.ColumnAlignment(3) = 0
    Me.OsenXPListBox1.ColumnAlignment(4) = 0
    
    ' set columnwidth
    Me.OsenXPListBox1.ColumnWidth(3) = cboSuppliers.ColumnWidth(2)
    Me.OsenXPListBox1.ColumnWidth(4) = cboCategories.ColumnWidth(2)
    
    Me.OsenXPListBox1.LockUpdate = False
    
End Sub
 
 
 
 
 
 
 
 
 
 
 
 
 
 








