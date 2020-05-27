VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form vbbego_007 
   BackColor       =   &H00BD6342&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00BD6342&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8115
      TabIndex        =   17
      Top             =   5340
      Width           =   8115
      Begin SISPAN.ButtonEx cmdExec 
         Cancel          =   -1  'True
         Height          =   345
         Index           =   3
         Left            =   6765
         TabIndex        =   18
         Top             =   105
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Appearance      =   2
         BorderStyle     =   3
         Caption         =   "&Keluar"
         CaptionOffsetX  =   -5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "vbbego_007.frx":0000
         PictureOffsetX  =   10
         PictureOffsetY  =   1
         TransparentColor=   16711935
         SkinDown        =   "vbbego_007.frx":0352
         SkinFocus       =   "vbbego_007.frx":1990
         SkinUp          =   "vbbego_007.frx":2FCE
         TransparentColor=   16711935
      End
      Begin SISPAN.ButtonEx cmdExec 
         Height          =   345
         Index           =   2
         Left            =   195
         TabIndex        =   19
         Top             =   105
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Appearance      =   2
         BorderStyle     =   3
         Caption         =   "&Bantuan"
         CaptionOffsetX  =   -5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "vbbego_007.frx":460C
         PictureOffsetX  =   5
         PictureOffsetY  =   1
         TransparentColor=   16711935
         SkinDown        =   "vbbego_007.frx":495E
         SkinFocus       =   "vbbego_007.frx":5F9C
         SkinUp          =   "vbbego_007.frx":75DA
         TransparentColor=   16711935
      End
      Begin SISPAN.ButtonEx cmdExec 
         Height          =   345
         Index           =   1
         Left            =   5520
         TabIndex        =   20
         Top             =   105
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Appearance      =   2
         BorderStyle     =   3
         Caption         =   "Pili&h Data"
         CaptionOffsetX  =   -5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "vbbego_007.frx":8C18
         PictureOffsetX  =   5
         PictureOffsetY  =   1
         TransparentColor=   16711935
         SkinDown        =   "vbbego_007.frx":8F6A
         SkinFocus       =   "vbbego_007.frx":A5A8
         SkinUp          =   "vbbego_007.frx":BBE6
         TransparentColor=   16711935
      End
      Begin SISPAN.ButtonEx cmdExec 
         Height          =   345
         Index           =   4
         Left            =   4260
         TabIndex        =   22
         Top             =   105
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Appearance      =   2
         BorderStyle     =   3
         Caption         =   "&Fields >>"
         CaptionOffsetX  =   -5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "vbbego_007.frx":D224
         PictureOffsetX  =   5
         PictureOffsetY  =   1
         TransparentColor=   16711935
         SkinDown        =   "vbbego_007.frx":D576
         SkinFocus       =   "vbbego_007.frx":EBB4
         SkinUp          =   "vbbego_007.frx":101F2
         TransparentColor=   16711935
      End
      Begin SISPAN.ButtonEx cmdExec 
         Height          =   345
         Index           =   5
         Left            =   1440
         TabIndex        =   23
         Top             =   105
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Appearance      =   2
         Enabled         =   0   'False
         BorderStyle     =   3
         Caption         =   "&Refresh"
         CaptionOffsetX  =   -2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "vbbego_007.frx":11830
         PictureOffsetX  =   8
         PictureOffsetY  =   1
         TransparentColor=   16711935
         SkinDisabled    =   "vbbego_007.frx":11B82
         SkinDown        =   "vbbego_007.frx":131C0
         SkinFocus       =   "vbbego_007.frx":147FE
         SkinUp          =   "vbbego_007.frx":15E3C
         TransparentColor=   16711935
      End
   End
   Begin VB.ListBox lstCheck 
      BackColor       =   &H00EFF7F7&
      Columns         =   2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1260
      IntegralHeight  =   0   'False
      ItemData        =   "vbbego_007.frx":1747A
      Left            =   180
      List            =   "vbbego_007.frx":1747C
      Style           =   1  'Checkbox
      TabIndex        =   16
      Top             =   5595
      Width           =   7770
   End
   Begin SISPAN.HyperLabel Tabed 
      Height          =   255
      Index           =   1
      Left            =   2295
      TabIndex        =   7
      Top             =   675
      Width           =   1605
      _extentx        =   2831
      _extenty        =   450
      caption         =   "Pencarian &Custom"
      captionoffsety  =   1
      font            =   "vbbego_007.frx":1747E
      forecolor       =   12582912
      backcolor       =   12632256
   End
   Begin SISPAN.HyperLabel Tabed 
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   6
      Top             =   675
      Width           =   1845
      _extentx        =   3254
      _extenty        =   450
      caption         =   "Pencarian Sederha&na"
      captionoffsety  =   1
      font            =   "vbbego_007.frx":174A2
      forecolor       =   12582912
      backcolor       =   15726583
   End
   Begin SISPAN.ButtonEx cmdExec 
      Height          =   345
      Index           =   0
      Left            =   6750
      TabIndex        =   9
      Top             =   2085
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   2
      BorderStyle     =   3
      Caption         =   "&Proses"
      CaptionOffsetX  =   -5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "vbbego_007.frx":174C6
      PictureOffsetX  =   10
      PictureOffsetY  =   1
      TransparentColor=   16711935
      SkinDown        =   "vbbego_007.frx":17818
      SkinFocus       =   "vbbego_007.frx":18E56
      SkinUp          =   "vbbego_007.frx":1A494
      TransparentColor=   16711935
   End
   Begin SISPAN.ButtonEx cmdExec 
      Height          =   345
      Index           =   10
      Left            =   6750
      TabIndex        =   10
      Top             =   1695
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   2
      Enabled         =   0   'False
      BorderStyle     =   3
      Caption         =   "&Reset"
      CaptionOffsetX  =   -5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "vbbego_007.frx":1BAD2
      PictureOffsetX  =   10
      PictureOffsetY  =   1
      TransparentColor=   16711935
      SkinDisabled    =   "vbbego_007.frx":1BE24
      SkinDown        =   "vbbego_007.frx":1D462
      SkinFocus       =   "vbbego_007.frx":1EAA0
      SkinUp          =   "vbbego_007.frx":200DE
      TransparentColor=   16711935
   End
   Begin SISPAN.ButtonEx cmdExec 
      Height          =   345
      Index           =   6
      Left            =   6750
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   2
      Enabled         =   0   'False
      BorderStyle     =   3
      Caption         =   "&Del Row"
      CaptionOffsetX  =   -3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "vbbego_007.frx":2171C
      PictureOffsetX  =   10
      PictureOffsetY  =   1
      TransparentColor=   16711935
      SkinDisabled    =   "vbbego_007.frx":21A6E
      SkinDown        =   "vbbego_007.frx":230AC
      SkinFocus       =   "vbbego_007.frx":246EA
      SkinUp          =   "vbbego_007.frx":25D28
      TransparentColor=   16711935
   End
   Begin SISPAN.ButtonEx cmdExec 
      Height          =   345
      Index           =   8
      Left            =   6750
      TabIndex        =   12
      Top             =   930
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   2
      Enabled         =   0   'False
      BorderStyle     =   3
      Caption         =   "&Add Row"
      CaptionOffsetX  =   -2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "vbbego_007.frx":27366
      PictureOffsetX  =   8
      PictureOffsetY  =   1
      TransparentColor=   16711935
      SkinDisabled    =   "vbbego_007.frx":276B8
      SkinDown        =   "vbbego_007.frx":28CF6
      SkinFocus       =   "vbbego_007.frx":2A334
      SkinUp          =   "vbbego_007.frx":2B972
      TransparentColor=   16711935
   End
   Begin VSFlex8Ctl.VSFlexGrid DrgData 
      Height          =   2805
      Left            =   210
      TabIndex        =   8
      Top             =   2475
      Width           =   7740
      _cx             =   13652
      _cy             =   4948
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   33023
      BackColorFixed  =   13534307
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14058595
      BackColorAlternate=   16246750
      GridColor       =   -2147483633
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   7
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin SISPAN.GradientLabel lCaption 
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   688
      GradientType    =   2
      Caption         =   "      Pencarian Data"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color1          =   12411714
      Color2          =   12582912
      Color3          =   9987120
      Color4          =   16636349
      HighlightColour =   16777215
      Begin VB.Shape Shape3 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   135
         Left            =   7245
         Shape           =   2  'Oval
         Top             =   135
         Width           =   135
      End
      Begin VB.Image imgProp 
         Height          =   135
         Left            =   7245
         Top             =   135
         Width           =   135
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   90
         Picture         =   "vbbego_007.frx":2CFB0
         Stretch         =   -1  'True
         Top             =   75
         Width           =   240
      End
      Begin VB.Image btnMenu 
         Height          =   285
         Index           =   0
         Left            =   7755
         Picture         =   "vbbego_007.frx":2D33A
         Top             =   60
         Width           =   285
      End
      Begin VB.Image btnMenu 
         Height          =   285
         Index           =   1
         Left            =   7440
         Picture         =   "vbbego_007.frx":2D844
         Top             =   60
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00986430&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   210
      ScaleHeight     =   1530
      ScaleWidth      =   6480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   945
      Width           =   6480
      Begin VB.TextBox txtCari 
         BackColor       =   &H00EFF7F7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   255
         TabIndex        =   0
         Top             =   570
         Width           =   6090
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00EFF7F7&
         Height          =   315
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   150
         Width           =   2370
      End
      Begin VB.OptionButton O2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF7F7&
         Caption         =   "0&2. Mengandung Kata"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   1185
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton O1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF7F7&
         Caption         =   "0&1. Kata Awalan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   930
         Width           =   1920
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Field:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3435
         TabIndex        =   3
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Kata yang akan dicari:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   285
         TabIndex        =   14
         Top             =   300
         Width           =   1755
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00EFF7F7&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00EFF7F7&
         Height          =   1530
         Left            =   -15
         Top             =   0
         Width           =   6420
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00986430&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   210
      ScaleHeight     =   1530
      ScaleWidth      =   6465
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   945
      Visible         =   0   'False
      Width           =   6465
      Begin VSFlex8LCtl.VSFlexGrid Grid 
         Height          =   1530
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6465
         _cx             =   11404
         _cy             =   2699
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   15726583
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15726583
         BackColorAlternate=   15269375
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   4
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"vbbego_007.frx":2DD4E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field Yang Akan Ditampilkan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   21
      Top             =   5370
      Width           =   2310
   End
   Begin VB.Shape shpTab 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   2265
      Shape           =   4  'Rounded Rectangle
      Top             =   660
      Width           =   1695
   End
   Begin VB.Shape shpTab 
      BackColor       =   &H00EFF7F7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   660
      Width           =   1995
   End
End
Attribute VB_Name = "vbbego_007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ccFields As New Collection
Dim ccFieldsDump As New Collection
'Dim SQL As String
Dim CallObj As Object, procName As String
Dim RCxx As New ADODB.Recordset
Dim SQL As String

Sub DecodeSQL()
          Dim I As Integer, cSQL As String, jjField As String, jjvalue As String
          Dim oAsField As String
          For I = 1 To Grid.Rows - 1
              If Trim(Grid.TextMatrix(I, 0)) <> "" And Trim(Grid.TextMatrix(I, 1)) <> "" Then
                    jjField = ccFields(Grid.TextMatrix(I, 0))
                    If InStr(1, jjField, "$") Then
                        jjField = Replace(jjField, "$", "")
                        jjvalue = AllowChar(FDec(Grid.TextMatrix(I, 2)))
                    ElseIf InStr(1, jjField, "#") Then
                        jjField = Replace(jjField, "#", "")
                        jjvalue = AllowChar(Grid.TextMatrix(I, 2))
                    ElseIf InStr(1, jjField, "@") Then
                        jjField = Replace(jjField, "@", "")
                        jjvalue = "'" & AllowChar(strToDate(Grid.TextMatrix(I, 2))) & "'"
                    ElseIf InStr(1, jjField, "&") Then
                        jjField = Replace(jjField, "&", "")
                        jjvalue = "'" & AllowChar(Grid.TextMatrix(I, 2)) & "'"
                    Else
                        If InStr(1, Grid.TextMatrix(I, 2), "@", vbTextCompare) Then                        'jjvalue = "'" & AllowChar(Grid.TextMatrix(I, 2)) & "'"
                        jjvalue = Replace(Grid.TextMatrix(I, 2), "@", "")
                        Else
                        jjvalue = "'" & AllowChar(Grid.TextMatrix(I, 2)) & "'"
                        End If
                    End If
                    Grid.AddItem ""
                    
                    If InStr(1, jjField, "�", vbTextCompare) Then
                       Dim kAS
                       kAS = Split(CStr(ccFieldsDump(jjField)), " AS ", , vbTextCompare)
                       jjField = kAS(0)
                    End If

                    If Trim(Grid.TextMatrix(I + 1, 0)) <> "" And Trim(Grid.TextMatrix(I + 1, 1)) <> "" Then
                       cSQL = cSQL & "(" & jjField & " " & _
                              Grid.TextMatrix(I, 1) & " " & _
                              jjvalue & ") " & _
                              Grid.TextMatrix(I, 3) & " "
                    Else
                       cSQL = cSQL & "(" & jjField & " " & _
                              Grid.TextMatrix(I, 1) & " " & _
                              jjvalue & ") "
                    
                    End If
                    Grid.RemoveItem Grid.Rows - 1
              Else
              End If
          Next I
             
           Dim nSource As String
           nSource = Replace(arrQueryForm.GetItem(Me.Tag), "�", "")
           nSource = Replace(nSource, "$", "") 'Currency
           nSource = Replace(nSource, "#", "") 'Numeric
           nSource = Replace(nSource, "&", "") 'String
           nSource = Replace(nSource, "@", "") 'Date/Time
          
          Dim inmyStr As String
          If Picture2.Visible Then
             oAsField = ccFields.Item(Combo1.Text)
             If InStr(1, oAsField, "�", vbTextCompare) Then
                kAS = Split(CStr(ccFieldsDump(oAsField)), " AS ", , vbTextCompare)
                inmyStr = kAS(0)
             Else
               inmyStr = oAsField
             End If
             If O1.Value Then
                inmyStr = "LEFT(" & inmyStr & "," & Len(AllowChar(txtCari)) & ")='" & AllowChar(txtCari) & "'"
             Else
                inmyStr = inmyStr & " LIKE '%" & AllowChar(txtCari) & "%'"
             End If
          Else
             inmyStr = cSQL
          End If
          Dim customSQL As String, Pos1 As Long
          For I = 0 To lstCheck.ListCount - 1
              If lstCheck.Selected(I) = True Then
                 oAsField = ccFields.Item(lstCheck.List(I))
                 If InStr(1, oAsField, "�", vbTextCompare) Then
                    customSQL = customSQL & ccFieldsDump.Item(oAsField) & ", "
                 Else
                    customSQL = customSQL & ccFields.Item(lstCheck.List(I)) & ", "
                 End If
              End If
          Next I
          customSQL = Replace(customSQL, "$", "") 'Currency
          customSQL = Replace(customSQL, "#", "") 'Numeric
          customSQL = Replace(customSQL, "&", "") 'String
          customSQL = Replace(customSQL, "@", "") 'Date/Time

          Pos1 = InStr(Pos1 + 1, nSource, "from", vbTextCompare)
          If Pos1 Then
             nSource = "SELECT " & Mid(customSQL, 1, Len(customSQL) - 2) & _
             Mid(nSource, Pos1 - 1)
          End If
                    
          If Trim(inmyStr) <> "" Then
             If InStr(1, nSource, "<!having>") > 0 Then
                nSource = Replace(nSource, "<!having>", " HAVING " & DrgData.Tag & "  " & inmyStr, , , vbTextCompare)
             ElseIf InStr(1, nSource, "<!where>") > 0 Then
                 nSource = Replace(nSource, "<!where>", " WHERE " & DrgData.Tag & "  " & inmyStr, , , vbTextCompare)
             End If
             ShowRecord nSource
          Else
             nSource = Replace(nSource, "<!having>", "")
             nSource = Replace(nSource, "<!where>", "")
             ShowRecord nSource
          End If
          
End Sub

Sub ShowRecord(nSource As String)
On Error GoTo Salah
   Dim lErr As String
   Set RCxx = New ADODB.Recordset
   lErr = SelectQuery(RCxx, nSource)
   If lErr = "" Then
      Set DrgData.DataSource = RCxx
      DrgData.DataRefresh
   End If
'   Dim colStr As String, I As Integer
'   For I = 0 To DrgData.Cols - 1
'       colStr = Replace(DrgData.TextMatrix(0, I), "_", " ")
'       colStr = StrConv(colStr, vbProperCase)
'       DrgData.TextMatrix(0, I) = colStr
'   Next I
Exit Sub
Salah:
ShowDlgMsg Me, "Ada kesalahan sewaktu pengisian data yang akan dicari.", vbOK, Error, True, False
End Sub

Sub ShowField(Obj As Object, Proc As String)
'Dim SQL As String
SQL = arrQueryForm.GetItem(Me.Tag)
Dim Pos1 As Long, Pos2 As Long, strRes1 As New Collection, strRes2 As New Collection
Dim strSelect As String
Set CallObj = Obj
procName = Proc
Pos1 = InStr(1, SQL, "SELECT", vbTextCompare) 'cari kata select pada string sebagai acuan
If Pos1 Then
   Pos2 = InStr(Pos1 + 6, SQL, " FROM ", vbTextCompare) 'dan diakhiri dengan kata from
   If Pos2 Then
      strSelect = Mid(SQL, Pos1 + 6, Pos2 - 7)
      Dim nFields, X As Integer
      nFields = Split(strSelect, ",") 'pisahkan dengan menggunakan seperator koma (,)
      For X = 0 To UBound(nFields)
         If InStr(1, Trim(nFields(X)), ".", vbTextCompare) Then 'Pisahkan antara nama table dan nama field
            Dim myff, mygg
            myff = Split(Trim(nFields(X)), ".")
            
            If InStr(1, myff(0), " ", vbTextCompare) Then 'Seleksi untuk Nama Table
               If Left(myff(0), 1) = "[" And Right(myff(0), 1) = "]" Then 'Cari jika ada spasi tanpa kurung buka siku
                  strRes1.Add myff(0) & ""
               Else
                  strRes1.Add "[" & myff(0) & "]"
               End If
            Else
               strRes1.Add "[" & myff(0) & "]"
            End If
            
            If InStr(1, myff(1), " AS ", vbTextCompare) Then 'Seleksi untuk Nama Field
                mygg = Split(myff(1), " AS ", , vbTextCompare)
                If UBound(mygg) > 0 Then
                   If InStr(1, mygg(1), " ", vbTextCompare) Then
                      If Left(mygg(1), 1) = "[" And Right(mygg(1), 1) = "]" Then   'Cari jika ada spasi tanpa kurung buka siku
                         strRes2.Add mygg(1) & "�"
                         ccFieldsDump.Add myff(0) & "." & myff(1), mygg(1) & "�"
                      Else
                         strRes2.Add "[" & mygg(1) & "]�"
                         ccFieldsDump.Add myff(0) & "." & myff(1), "[" & mygg(1) & "]�"
                      End If
                   Else
                      strRes2.Add "[" & Trim(mygg(1)) & "]�"
                      ccFieldsDump.Add myff(0) & "." & myff(1), "[" & Trim(mygg(1)) & "]�"
                   End If
                End If
            Else
                strRes2.Add myff(1)
            End If
         Else
         
         End If
      Next X
      Dim I As Integer, C As String, d As String
      For I = 1 To strRes1.Count

           C = C & strRes2(I) & "|"
           d = Replace(strRes2(I), "]", "")

           d = Replace(d, "[", "")
           d = Replace(d, "_", " ")
           d = Replace(d, "�", "")
           d = Replace(d, "$", "") 'Currency
           d = Replace(d, "#", "") 'Numeric
           d = Replace(d, "&", "") 'String
           d = Replace(d, "@", "") 'Date/Time
           Combo1.AddItem UCase(d)
           lstCheck.AddItem UCase(d)
           lstCheck.Selected(lstCheck.ListCount - 1) = True
           If InStr(1, strRes2(I), "�", vbTextCompare) > 0 Then
              ccFields.Add strRes2(I), UCase(d)
           Else
              ccFields.Add strRes1(I) & "." & strRes2(I), UCase(d)
           End If
      Next I
      C = Replace(C, "]", "")
      C = Replace(C, "[", "")
      C = Replace(C, "_", " ")
      C = Replace(C, "�", "")
      C = Replace(C, "$", "") 'Currency
      C = Replace(C, "#", "") 'Numeric
      C = Replace(C, "&", "") 'String
      C = Replace(C, "@", "") 'Date/Time
      Grid.ColComboList(0) = StrConv(C, vbUpperCase)
      Combo1.ListIndex = 0
   End If
   Set strRes1 = Nothing
   Set strRes2 = Nothing
   C = ""
   d = ""
End If
End Sub

Private Sub btnMenu_Click(Index As Integer)
Select Case Index
       Case 0
            Unload Me
       Case 1
            Me.WindowState = vbMinimized
            Me.Hide
End Select
End Sub

Private Sub cmdExec_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case 0
            Dim I As Integer
            For I = 0 To lstCheck.ListCount - 1
                If lstCheck.Selected(I) = True Then
                    DecodeSQL
                    Exit For
                End If
            Next I
            DrgData.SetFocus
       Case 1
            If DrgData.Rows > 1 Then
            If DrgData.TextMatrix(1, 0) <> "" Then
            Dim m As String
            For I = 0 To DrgData.Cols - 1
               m = m & DrgData.TextMatrix(DrgData.Row, I) & "|"
            Next I
            DrgData.SetFocus
            Me.Hide
            CallByName CallObj, procName, VbMethod, m
            CallObj.SetFocus
            End If
            End If
       
       Case 3
            Me.Hide
       Case 4
            If cmdExec(4).Caption = "&Fields <<" Then
               cmdExec(4).Caption = "&Fields >>"
               lstCheck.Enabled = False
               Me.Height = 5940
            Else
               cmdExec(4).Caption = "&Fields <<"
               Me.Height = 7605
               lstCheck.Enabled = True
            End If
       Case 5
            Set ccFields = Nothing
            ShowField CallObj, procName
       Case 6
            If Grid.Rows > 1 Then Grid.RemoveItem Grid.Row
       Case 8
            Grid.AddItem ""
       Case 10
            Grid.Rows = 1
            Grid.Rows = 6
End Select
End Sub

Private Sub DrgData_DblClick()
cmdExec_Click 1
End Sub

Private Sub DrgData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyUp
            If DrgData.Row = 1 Then
               If Picture2.Visible Then
                  txtCari.SetFocus
               Else
                  Grid.SetFocus
               End If
            End If
End Select
End Sub

Private Sub DrgData_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdExec_Click 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKey1 And Shift = 2 Then
   Tabed_Click 0
ElseIf KeyCode = vbKey2 And Shift = 2 Then
   Tabed_Click 1
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RemoveFindItem Me.Tag
Set DrgData.DataSource = Nothing
Set RCxx = Nothing
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
   If Trim(Grid.TextMatrix(Row - 1, 3)) = "" Then
      Cancel = True
   End If
ElseIf Row > 2 Then
   If Trim(Grid.TextMatrix(Row - 1, 0)) = "" Then
      Cancel = True
   End If
End If
Select Case Col
       Case 1, 2
          If Trim(Grid.TextMatrix(Row, Col - 1)) = "" Then Cancel = True
      Case 3
          If Trim(Grid.TextMatrix(Row, 1)) = "" Then Cancel = True
End Select
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
   Grid.AddItem ""
ElseIf KeyCode = vbKeyDelete Then
   If Grid.Rows > 2 Then Grid.RemoveItem Grid.Row
ElseIf KeyCode = vbKeyC And Shift = 2 Then
    Clipboard.Clear
    Clipboard.SetText Chr(255) & vbKeyTab & Grid.TextMatrix(Grid.Row, 0) & vbKeyTab & Grid.TextMatrix(Grid.Row, 1) & vbKeyTab & Grid.TextMatrix(Grid.Row, 2) & vbKeyTab & Grid.TextMatrix(Grid.Row, 3)
ElseIf KeyCode = vbKeyV And Shift = 2 Then
    Dim cc
    cc = Split(Clipboard.GetText, vbKeyTab)
    If UBound(cc) > 0 Then
       If cc(0) = Chr(255) Then
          If Grid.TextMatrix(Grid.Row - 1, 3) <> "" Then
                Grid.TextMatrix(Grid.Row, 0) = cc(1)
                Grid.TextMatrix(Grid.Row, 1) = cc(2)
                Grid.TextMatrix(Grid.Row, 2) = cc(3)
                Grid.TextMatrix(Grid.Row, 3) = cc(4)
                If Grid.Row < Grid.Rows - 1 Then Grid.Row = Grid.Row + 1
          End If
       End If
    End If
End If
End Sub

Private Sub Form_Load()
'Set Picture = LoadResPicture("POLA02", vbResBitmap)
'HideCaption hwnd, False
On Error Resume Next
If GetSetting(vbReg, "Setting\View", "AllTransparent", "") <> "All" Then
   'MakeTransparent GetSetting(vbReg, "Setting\View\" & Me.Name, "Transparent", 255), Me.hWnd
Else
   'MakeTransparent GetSetting(vbReg, "Setting\View", "Transparent", 255), Me.hWnd
End If

End Sub

Private Sub lCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then MoveIt Me.hWnd
End Sub

Private Sub Tabed_Click(Index As Integer)
Select Case Index
       Case 0
            Tabed(Index).BackColor = &HEFF7F7
            shpTab(Index).BackColor = &HEFF7F7
            Tabed(1).BackColor = &HC0C0C0
            shpTab(1).BackColor = &HC0C0C0
            Picture2.Visible = True
            Picture1.Visible = False
            cmdExec(8).Enabled = False
            cmdExec(6).Enabled = False
            cmdExec(10).Enabled = False
            txtCari.SetFocus
       Case 1
            Tabed(Index).BackColor = &HEFF7F7
            shpTab(Index).BackColor = &HEFF7F7
            Tabed(0).BackColor = &HC0C0C0
            shpTab(0).BackColor = &HC0C0C0
            Picture1.Visible = True
            Picture2.Visible = False
            cmdExec(8).Enabled = True
            cmdExec(6).Enabled = True
            cmdExec(10).Enabled = True
            Grid.SetFocus
End Select
End Sub

Private Sub txtCari_GotFocus()
Blok txtCari
End Sub

Private Sub txtCari_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyDown
            DrgData.SetFocus
End Select
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdExec_Click 0: KeyAscii = 0
End Sub

Private Sub imgProp_Click()
Form17.Tag = Me.hWnd & ";" & Me.Name
Form17.Show 1, Me
End Sub

