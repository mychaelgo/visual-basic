VERSION 5.00
Object = "{198EAA50-71CD-47FE-888B-89B2BE177BB3}#1.0#0"; "OSENVISTASUITE2009.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Northwind Traders"
   ClientHeight    =   8370
   ClientLeft      =   2940
   ClientTop       =   1770
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VistaSuitePro.OsenVistaPicture OsenXPPicture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   420
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1720
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
      Picture         =   "frm_main.frx":0ECA
      BorderColor     =   14854529
      PictureAlignment=   7
      GradientBackGround=   -1  'True
      GradientColor2  =   12632256
      GradientOrientation=   1
      UseBottomLine   =   -1  'True
      UseBorderColor  =   0   'False
      Description     =   "One Portals Way, Twin Points WA  98156"
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   "Northwind Traders"
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   4194304
      DescriptionLeft =   24
      BinaryImage     =   "frm_main.frx":2A1C
      WindowColor     =   0
   End
   Begin VistaSuitePro.MyImageList SmallIcons 
      Left            =   1020
      Top             =   3600
      _ExtentX        =   900
      _ExtentY        =   767
      Size            =   60844
      Images          =   "frm_main.frx":2A34
      Version         =   131072
      KeyCount        =   53
      Keys            =   $"frm_main.frx":11800
   End
   Begin VistaSuitePro.MyImageList LargeIcons 
      Left            =   1860
      Top             =   3660
      _ExtentX        =   900
      _ExtentY        =   767
      IconSizeX       =   32
      IconSizeY       =   32
      Iconsize        =   2
      Size            =   154420
      Images          =   "frm_main.frx":118A0
      Version         =   131072
      KeyCount        =   35
      Keys            =   "????????????????????????????????????????????????????????????????????ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VistaSuitePro.OsenVistaToolBar OsenXPToolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   1395
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XPBlend         =   0   'False
      TotalButton     =   19
      ImageListName   =   "SmallIcons"
      Bname1          =   "System"
      Btype1          =   1
      Bwidth1         =   0
      Bchecked1       =   0   'False
      Bvalue1         =   0   'False
      BNI1            =   35
      BSI1            =   35
      Bname2          =   "TreeView"
      Btip2           =   "Show/Hide TreeView"
      Btype2          =   0
      Bwidth2         =   0
      Bchecked2       =   -1  'True
      Bvalue2         =   0   'False
      BNI2            =   51
      BSI2            =   51
      Bname3          =   "Button21"
      Btype3          =   2
      Bwidth3         =   0
      Bchecked3       =   -1  'True
      Bvalue3         =   0   'False
      BNI3            =   51
      BSI3            =   6
      Bname4          =   "Back"
      Btype4          =   0
      Bwidth4         =   0
      Bchecked4       =   0   'False
      Bvalue4         =   0   'False
      BNI4            =   36
      BSI4            =   36
      Bname5          =   "Forward"
      Btype5          =   0
      Bwidth5         =   0
      Bchecked5       =   0   'False
      Bvalue5         =   0   'False
      BNI5            =   37
      BSI5            =   37
      Bname6          =   "Up one level"
      Btype6          =   0
      Bwidth6         =   0
      Bchecked6       =   0   'False
      Bvalue6         =   0   'False
      BNI6            =   38
      BSI6            =   38
      Bname7          =   "Button5"
      Btype7          =   2
      Bwidth7         =   0
      Bchecked7       =   0   'False
      Bvalue7         =   0   'False
      Bname8          =   "New"
      Btype8          =   0
      Bwidth8         =   0
      Bchecked8       =   0   'False
      Bvalue8         =   0   'False
      BNI8            =   39
      BSI8            =   39
      Bname9          =   "Edit"
      Btype9          =   0
      Bwidth9         =   0
      Bchecked9       =   0   'False
      Bvalue9         =   0   'False
      BNI9            =   40
      BSI9            =   40
      Bname10         =   "Delete"
      Btype10         =   0
      Bwidth10        =   0
      Bchecked10      =   0   'False
      Bvalue10        =   0   'False
      BNI10           =   41
      BSI10           =   41
      Bname11         =   "Button9"
      Btype11         =   2
      Bwidth11        =   0
      Bchecked11      =   0   'False
      Bvalue11        =   0   'False
      Bname12         =   "Search"
      Btype12         =   0
      Bwidth12        =   0
      Bchecked12      =   0   'False
      Bvalue12        =   0   'False
      BNI12           =   42
      BSI12           =   42
      Bname13         =   "Filter"
      Btype13         =   0
      Bwidth13        =   0
      Bchecked13      =   0   'False
      Bvalue13        =   0   'False
      BNI13           =   43
      BSI13           =   43
      Bname14         =   "Refresh"
      Btype14         =   0
      Bwidth14        =   0
      Bchecked14      =   0   'False
      Bvalue14        =   0   'False
      BNI14           =   44
      BSI14           =   44
      Bname15         =   "Button13"
      Btype15         =   2
      Bwidth15        =   0
      Bchecked15      =   0   'False
      Bvalue15        =   0   'False
      Bname16         =   "Preview"
      Btype16         =   0
      Bwidth16        =   0
      Bchecked16      =   0   'False
      Bvalue16        =   0   'False
      BNI16           =   45
      BSI16           =   45
      Bname17         =   "Export To Excel"
      Btype17         =   0
      Bwidth17        =   0
      Bchecked17      =   0   'False
      Bvalue17        =   0   'False
      BNI17           =   46
      BSI17           =   46
      Bname18         =   "Button16"
      Btype18         =   2
      Bwidth18        =   0
      Bchecked18      =   0   'False
      Bvalue18        =   0   'False
      Bname19         =   "Help"
      Btype19         =   1
      Bwidth19        =   0
      Bchecked19      =   0   'False
      Bvalue19        =   0   'False
      BNI19           =   47
      BSI19           =   47
   End
   Begin VB.PictureBox pResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
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
      Height          =   1725
      Left            =   780
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1725
      ScaleWidth      =   30
      TabIndex        =   3
      Top             =   3600
      Width           =   30
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   1815
      Left            =   450
      ScaleHeight     =   1815
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   3540
      Visible         =   0   'False
      Width           =   30
   End
   Begin VistaSuitePro.OsenVistaTreeView tvwMenus 
      Height          =   4665
      Left            =   150
      TabIndex        =   4
      Top             =   1920
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   8229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectedBackColor=   12958375
      SelectedColor   =   16777215
      LostFocusSelectedBackColor=   12632256
      MouseIcon       =   "frm_main.frx":373F4
      MousePointer    =   99
      ShowNumber      =   -1  'True
      BorderStyle     =   0
      HeaderCaption   =   "Northwind Traders"
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowHeader      =   -1  'True
      GradientColor2  =   12632256
      HeaderForeColor =   16777215
      LineColor       =   16576
      BackSelection   =   8438015
      WindowColor     =   0
   End
   Begin VistaSuitePro.OsenVistaListBox vList 
      Height          =   4695
      Left            =   3120
      TabIndex        =   1
      Top             =   1950
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   8281
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
      FontNormal      =   0
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      AllowEdit       =   0   'False
      WordWrap        =   0   'False
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      MousePointer    =   99
      MouseIcon       =   "frm_main.frx":37556
      BorderColor     =   12958375
      ShowHeader      =   -1  'True
      ShowGridLines   =   -1  'True
      AlternateRowColors=   -1  'True
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFontColor =   16777215
      ASURC           =   -1  'True
      IMGLIST         =   "SmallIcons"
      FormatDateTime  =   "mmm dd,yyyy"
      ViewMode        =   1
      ForeColorSelected=   16576
      Picture         =   "frm_main.frx":376B8
      LargeImageList  =   "LargeIcons"
      ReadOnDemand    =   -1  'True
      BeginProperty LargeIconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorOver =   12164479
      HeaderGradientAllow=   -1  'True
      HeaderForeColor =   16777215
      BinaryImage     =   "frm_main.frx":3B0EC
   End
   Begin VistaSuitePro.OsenVistaForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
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
      Caption         =   "Northwind Traders"
      TitleTop        =   7
      icon            =   "frm_main.frx":3B104
      BorderStyle     =   1
      AutoBackColor   =   0   'False
   End
   Begin VistaSuitePro.OsenVistaStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   8
      Top             =   7935
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   767
      BackColor       =   14936810
      ForeColor       =   16777215
      ForeColorDissabled=   8421504
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
      NumberOfPanels  =   8
      HaveXPForm      =   -1  'True
      PWidth1         =   300
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   ""
      pTextAlignment1 =   0
      PanelPicture1   =   "frm_main.frx":3B69E
      PanelPicAlignment1=   0
      PWidth2         =   150
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   ""
      pTextAlignment2 =   0
      PanelPicture2   =   "frm_main.frx":3B6BA
      PanelPicAlignment2=   0
      PWidth3         =   240
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Powered By http://osenxpsuite.net"
      pTextAlignment3 =   0
      pTextBold3      =   -1  'True
      PanelPicture3   =   "frm_main.frx":3B6D6
      PanelPicAlignment3=   0
      PWidth4         =   55
      PMinWidth4      =   0
      pTTText4        =   ""
      pType4          =   5
      pText4          =   "CAPS"
      pTextAlignment4 =   0
      PanelPicture4   =   "frm_main.frx":3BA28
      PanelPicAlignment4=   0
      PWidth5         =   50
      PMinWidth5      =   0
      pTTText5        =   ""
      pType5          =   6
      pText5          =   "NUM"
      pTextAlignment5 =   0
      PanelPicture5   =   "frm_main.frx":3BA44
      PanelPicAlignment5=   0
      PWidth6         =   60
      PMinWidth6      =   0
      pTTText6        =   ""
      pType6          =   7
      pText6          =   "SCROLL"
      pTextAlignment6 =   0
      PanelPicture6   =   "frm_main.frx":3BA60
      PanelPicAlignment6=   0
      PWidth7         =   75
      PMinWidth7      =   0
      pTTText7        =   ""
      pType7          =   3
      pText7          =   "2005-06-27"
      pTextAlignment7 =   0
      PanelPicture7   =   "frm_main.frx":3BA7C
      PanelPicAlignment7=   0
      PWidth8         =   65
      PMinWidth8      =   0
      pTTText8        =   ""
      pType8          =   2
      pText8          =   "01:21:58"
      pTextAlignment8 =   0
      PanelPicture8   =   "frm_main.frx":3BA98
      PanelPicAlignment8=   0
      GradientColor1  =   10000535
      GradientColor2  =   5460819
      Begin VistaSuitePro.OsenVistaProgressBar pBar 
         Height          =   225
         Left            =   4590
         TabIndex        =   5
         Top             =   90
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   2871848
         Value           =   100
      End
   End
   Begin VistaSuitePro.OsenVistaHookMenu OsenXPHookMenu1 
      Height          =   375
      Left            =   3780
      TabIndex        =   6
      Top             =   4290
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      MCountMenu      =   3
      XMenuA1         =   "System "
      XMenuACS1       =   ""
      XMenuC1         =   "mnu_System"
      XMenuE1         =   -1  'True
      XMenuH1         =   0   'False
      XMenuA2         =   "Right Event "
      XMenuACS2       =   ""
      XMenuC2         =   "Mnu_Right"
      XMenuE2         =   -1  'True
      XMenuH2         =   0   'False
      XMenuA3         =   "Help "
      XMenuACS3       =   ""
      XMenuC3         =   "mnu_Help"
      XMenuE3         =   -1  'True
      XMenuH3         =   0   'False
   End
   Begin VB.Menu mnu_System 
      Caption         =   "System"
      Visible         =   0   'False
      Begin VB.Menu MnuSysChild 
         Caption         =   "Available"
         Index           =   0
      End
   End
   Begin VB.Menu Mnu_Right 
      Caption         =   "Right Event"
      Visible         =   0   'False
      Begin VB.Menu MnuAction 
         Caption         =   "User Action"
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Hlp_Child 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCrLeft             As Long     ' Resize
Private xL                  As Long     ' Resize
Private mWidth              As Long     ' Resize
Private bMove               As Boolean  ' Resize
Private lngTime             As Long
Private m_back              As Collection
Private m_Forward           As Collection
Private b_Allow_Back        As Boolean
Private b_Allow_Forward     As Boolean
Private strPrivileges       As String
Private m_LastNode          As String
Private m_HaveSystray       As Boolean
Private WithEvents Systray  As CLS_SysTray
Attribute Systray.VB_VarHelpID = -1


' Purpose : Insert Node item by recordset >> From table "Nodes" <<

Public Sub CreateNode()

    On Error Resume Next

    ' Prepared query to get User Privileges from users table
    mStrSQL = "select privileges from users where userid='" & StrUserID & "' "

    ' Get .....
    strPrivileges = ADO_SQL_RESULT(mStrSQL)

    vList.Clear

    With tvwMenus
        ' Clean Up
        .Clear

        .LockUpdate = True

    
        ' Provide node collection
        Set .TableNodes = GetRST("select * from nodes order by nodekey")

        ' Provide privileges information
        .Privileges = strPrivileges
        
        ' Create node by recordset ...
        .CreateNodeByCurrentRecordset "parent", "NWIND", 0, 1, 3, 4

        ' Check Count of nodes to expand
        If .Nodes.Count Then
            .Nodes(1).Expanded = True
        End If

        ' Unlock , and draw all nodes
        .LockUpdate = False

    End With

    ' Now prepare Back and Forward collection
    ' Clean Up and create new collection
    Set m_back = Nothing
    Set m_back = New Collection
    Set m_Forward = Nothing
    Set m_Forward = New Collection

    b_Allow_Back = True
    b_Allow_Forward = True

End Sub

' Purpose: Call Add/Edit form [Like Property Page on Windows System]

Private Sub DisplayForm()

    On Error Resume Next

    Select Case FrmName

        Case "order"
            Unload frm_orders
            frm_orders.Show 0, Me
            '<<Revisi>> 2005-11-10 <Set Parent>
            SetMyParent hWnd, frm_orders.hWnd

        Case "user"
            Unload frm_user_mgmt
            frm_user_mgmt.Show 0, Me
            '<<Revisi>> 2005-11-10 <Set Parent>
            SetMyParent hWnd, frm_user_mgmt.hWnd

        Case "employee"
            Unload frm_employees
            frm_employees.Show 0, Me
            '<<Revisi>> 2005-11-10 <Set Parent>
            SetMyParent hWnd, frm_employees.hWnd

        Case "category"
            Unload frm_Category
            frm_Category.Show 0, Me
            '<<Revisi>> 2005-11-10 <Set Parent>
            SetMyParent hWnd, frm_Category.hWnd

        Case "customer"
            Unload frm_customers
            frm_customers.Show 0, Me
            '<<Revisi>> 2005-11-10 <Set Parent>
            SetMyParent hWnd, frm_customers.hWnd

        Case "product"
            Unload frm_products
            frm_products.Show 0, Me
            '<<Revisi>> 2005-11-10 <Set Parent>
            SetMyParent hWnd, frm_products.hWnd

        Case "supplier"
            Unload frm_suppliers
            frm_suppliers.Show 0, Me
            '<<Revisi>> 2005-11-10 <Set Parent>
            SetMyParent hWnd, frm_suppliers.hWnd

        Case Else

            DoEvents
    End Select

End Sub

'Purpose: Display Report

Private Sub DisplayReport()

    On Error Resume Next

    ' Check resultset view
    If vList.ViewMode = lvwDetail Then

        ' Check have record(s) or not
        If vList.ListCount Then

            #If ACTIVEREPORT = 1 Then
            
                ' show activereport
                ' please make sure that you have activereport installed on your system
                
                vList.LockRecordset = False
                
                Select Case RptName
                    Case "customer"
                        DisplayARV vList.ActiveRst, "Customers", App.Path & "\reports\customers.rpx"
                    Case "supplier"
                        DisplayARV vList.ActiveRst, "Suppliers", App.Path & "\reports\suppliers.rpx"
                    Case "product"
                        DisplayARV vList.ActiveRst, "Products", App.Path & "\reports\products.rpx"
                    Case Else
                End Select
                
                vList.LockRecordset = True
            #Else
                ' show DataReport, and bound it with recordset in ListBox
                Select Case RptName
    
                    Case "customer"
                        vList.ShowReport rpt_customers
    
                    Case "supplier"
                        vList.ShowReport rpt_suppliers
    
                    Case "product"
                        vList.ShowReport rpt_products
                        
                    Case Else
                End Select
            #End If

        End If

    End If

End Sub

' Purpose: Initialize system ...

Private Sub Form_Load()

    On Error Resume Next

    ' Make sure the osenxpform controls work fine ....
    Me.OsenXPForm1.Init Me

    OsenXPToolBar1.ButtonValue(2) = True

    'Set the Sound for TreeView and ListBox
    tvwMenus.SoundClick = App.Path & "\resources\click.wav"
    tvwMenus.SoundExpand = App.Path & "\resources\expand.wav"
    tvwMenus.SoundHover = App.Path & "\resources\hover.wav"

    vList.SoundClick = App.Path & "\resources\click.wav"
    vList.SoundHover = App.Path & "\resources\hover.wav"

    ' Prepared icons collection
    PreparedImage


    ' Reposition
    mWidth = 200
    ResizePosition

    ' Prepared item of treeView (Nodes)
    CreateNode
    AlreadyExist = True

    vList.InsertViewFromNode tvwMenus.CurrentNode

    Set Systray = New CLS_SysTray
        
    ' Create systray
    '========== ************* WARNING ************** ============
    ' DON'T USE ME.HWND for pHwnd parameter ( OsenXPPicture1.Hwnd is recommended )
    Systray.Create "Northwind Traders 2008", Me.OsenXPPicture1.hWnd, OsenXPForm1.Icon
    '============================================================
        
    m_HaveSystray = True
            
    ' Check registration status
    If Not IsRegistered Then
            
        ' Show baloon popup
        Systray.BalloonShow "You have " & RemainDays & " remaining days." & vbLf & "Please support us by purchasing OsenVistaSuite" & vbLf & vbLf & "Click here to buy OsenVistaSuite now!", "Unregistered User", xpTrayIcon
    Else
        
        Dim sUser As String, sComp As String, sKey As String
            
        GetRegistrationInfo sUser, sComp, sKey
            
        ' Show baloon popup
        Systray.BalloonShow "Thank you very much for registering OsenVistaSuite 2008." & vbLf & vbLf & "Registration Name: " & sUser & vbLf & "Company Name: " & sComp & vbLf & "Serial Number: " & sKey & vbLf, "Registration Info", xpTrayIcon
    End If

End Sub

' Purpose: Resize ....

Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState <> 1 Then ResizePosition

End Sub

' Purpose : Exit App?

Private Sub Form_Unload(Cancel As Integer)

    ' exit application

    If Not CloseProgram Then
        Cancel = 1
    Else

        If m_HaveSystray Then
            ' Remove systray from taskbar
            Systray.Remove
            ' clean up
            Set Systray = Nothing
            
            On Error Resume Next
            ADOCN.Close
            
        End If
    End If

End Sub

' Purpose : Create Dynamic toolbar button and menus

Private Sub InitMenus()

    On Error Resume Next

    Dim i As Long
    Dim StrA

    ' Set Up ImageList for Hook Menu
    Set Me.OsenXPHookMenu1.ImageList = SmallIcons

    ' Now trying to create dynamic menus for User Activity
    For i = 1 To 10
        Load MnuAction(i)
        StrA = Split(LoadResString(104 + i), "|")
        MnuAction(i).Caption = StrA(0)
        OsenXPHookMenu1.SetIconIndex MnuAction(i), CLng(StrA(2)) + 1
    Next i

    MnuAction(0).Visible = False

    ' Create Child of System Menu
    For i = 1 To 4

        Load MnuSysChild(i)
        StrA = Split(LoadResString(118 + i), "|")
        MnuSysChild(i).Caption = StrA(0)
        OsenXPHookMenu1.SetIconIndex MnuSysChild(i), CLng(StrA(1)) + 1
    Next i

    MnuSysChild(0).Visible = False

    ' Create Child of Help Menu
    For i = 1 To 4
        Load Mnu_Hlp_Child(i)
        Mnu_Hlp_Child(i).Caption = LoadResString(122 + i)
    Next i

    Mnu_Hlp_Child(0).Visible = False

End Sub

Private Sub Mnu_Hlp_Child_Click(Index As Integer)

    If Index = 4 Then
        MsgBoxGT "Northwind Database Enterprise System v 0.9", vbInformation, "Northwind Traders", 2

    Else
        MsgBoxGT "Sorry the help system does not exist.", vbExclamation, "Help"
    End If

End Sub

'Purpose: Do it as sama as toolbar do

Private Sub MnuAction_Click(Index As Integer)

    ' Calll Procedure Toolbar_Button_Click

    OsenXPToolBar1_ButtonClick Index + 7, ""

End Sub

Private Sub MnuSysChild_Click(Index As Integer)

    Select Case Index

        Case 1
            frm_login.Show 1

        Case 2

        Case 4
            Unload Me
    End Select

End Sub



Private Sub OsenXPStatusBar1_MouseDownInPanel(iPanel As Long)

    If iPanel = 3 Then

        ' Open My Homepage
        OpenBrowser hWnd

    End If

End Sub

'purpose: User want to change the records or other ....

Private Sub OsenXPToolBar1_ButtonClick(Index As Integer, _
                                       sText As String)

    On Error GoTo Err_MSG
    Dim vt As CLS_xpNode

    Select Case Index

        Case 2
            ResizePosition

        Case 4
            tvwMenus.Back

        Case 5
            tvwMenus.Forward

        Case 6
            tvwMenus.UpOneLevel

        Case 8, 9
            mStrSQL = "select form_name from nodes where nodekey='" & tvwMenus.CurrentID & "' "
            FrmName = ADO_SQL_RESULT(mStrSQL)
            KeyValue = vList.GetItemKey

            If FrmName <> "" Then
                IsNew = (Index = 8)
                DisplayForm
            End If

        Case 10

            ' user allow to delete record if viewmode=detail not LargeIcon
            If vList.ViewMode = lvwDetail Then

                ' Check listview have data or not
                If vList.ListCount Then

                    If vList.ListIndex > -1 Then
                    
                        mStrSQL = "select form_name from nodes where nodekey='" & tvwMenus.CurrentID & "' "
                        FrmName = ADO_SQL_RESULT(mStrSQL)
                        KeyValue = vList.GetItemKey
                        
                        If Len(FrmName) Then
                        
                            ' Confirmation before delete
                            If MsgBoxGT("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                                DeleteRecord FrmName
                            End If
                        
                        End If

                    End If

                Else

                    MsgBoxGT "There are no record(s) to delete.", vbExclamation, "Delete", 3

                End If

            End If

        Case 12

            If vList.ViewMode = lvwDetail Then
                vList.List_Search
            End If

        Case 13

            If vList.ViewMode = lvwDetail Then
                vList.List_Filter
            End If

        Case 14

            If vList.ViewMode = lvwDetail Then
                tvwMenus.SetNodeClick tvwMenus.CurrentNode
            End If

        Case 16
            mStrSQL = "select report_name from nodes where nodekey='" & tvwMenus.CurrentID & "' "
            RptName = ADO_SQL_RESULT(mStrSQL)

            If RptName <> "" Then DisplayReport

        Case 17

            If vList.ViewMode = lvwDetail Then
                vList.ExportToExcel
            End If

        Case 19
            MsgBoxGT "Sorry the help system does not exist.", vbExclamation, "Help"

    End Select

    Exit Sub

Err_MSG:
    MsgBoxGT "Error No. " & Err.Number & vbLf & Err.Description, vbCritical, "Northwind traders 2006", 3

End Sub

' Purpose: Show Popup Menu

Private Sub OsenXPToolBar1_PopUpMainMenu(Index As Integer, _
                                         sText As String, _
                                         X As Long, _
                                         Y As Long)

    On Error Resume Next

    Select Case Index

        Case 1 ' Connection >> Login,ChangePassword,Exit
            PopupMenu mnu_System, , X, Y

        Case 19 ' Help,About
            PopupMenu mnu_Help, , X, Y
    End Select

End Sub

' Purpose : Insert Image Collection to Current MyImageList

Private Sub PreparedImage()

    ' Set Image handle

    tvwMenus.ImageList = SmallIcons

    ' Create Menus
    InitMenus

End Sub

' Purpose : Resize event handle on runtime

Private Sub pResize_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    If Button = 1 Then bMove = True

End Sub

' Purpose : Resize event handle on runtime

Private Sub pResize_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    On Error Resume Next

    If bMove Then
        If Button = 1 Then
            xL = CLng((X / 15))
            Picture2.Left = tvwMenus.Width + xL
            Picture2.Visible = True
        End If
    End If

End Sub

' Purpose : Resize event handle on runtime

Private Sub pResize_MouseUp(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

    On Error Resume Next
    Picture2.Visible = False
    mCrLeft = xL
    mWidth = tvwMenus.Width
    ResizePosition

End Sub

' Purpose : Reposition Controls

Private Sub ResizePosition()

    On Error Resume Next
    pResize.Visible = OsenXPToolBar1.ButtonValue(2)
    tvwMenus.Visible = OsenXPToolBar1.ButtonValue(2)
        
    ' Reposition TreeView and ListBox(ListView:))
    If OsenXPToolBar1.ButtonValue(2) Then
        tvwMenus.Move 5, OsenXPToolBar1.Top + OsenXPToolBar1.Height + 1, mWidth + mCrLeft, Me.ScaleHeight - (OsenXPToolBar1.Top + OsenXPToolBar1.Height + 1) - OsenXPStatusBar1.Height - 2
        vList.Move tvwMenus.Left + tvwMenus.Width + 2, tvwMenus.Top, ScaleWidth - tvwMenus.Width - 13, tvwMenus.Height
    Else
        vList.Move 4, OsenXPToolBar1.Top + OsenXPToolBar1.Height + 1, ScaleWidth - 10, Me.ScaleHeight - (OsenXPToolBar1.Top + OsenXPToolBar1.Height + 1) - OsenXPStatusBar1.Height - 2
    End If
        
    ' Repos PicHandle for Spliter
    pResize.Move tvwMenus.Left + tvwMenus.Width, 124, 2, tvwMenus.Height
    Picture2.Move tvwMenus.Left + tvwMenus.Width, 124, 2, tvwMenus.Height

End Sub

Private Sub Systray_BalloonClick()

    OpenBrowser Me.hWnd, "http://osenxpsuite.net/buy"
    
End Sub

Private Sub Systray_RightButtonClick()
    PopupMenu mnu_System
End Sub

Private Sub tvwMenus_CloseClick()

    'When user clicked CloseButton on this treeview

    OsenXPToolBar1.ButtonValue(2) = False
    ResizePosition
    tvwMenus.Visible = False

End Sub

' Purpose: Create dynamic nodes (child dynamic) if availables, when first time selected

Private Sub tvwMenus_NodeChange(Node As VistaSuitePro.CLS_xpNode)

    On Error Resume Next

    ' Make sure this node have not dynamic child and first time to select
    If Node.FirstSelected = False Then

        ' Check from database, is nodekey or dynamic nodekey available to create dynamic child or not
        ' Prepare SQL to do it
        mStrSQL = "select * from nodes where nodekey='" & Node.DynamicParentKey & "' and havechild >= " & Node.DynamicLevel

        ' Check from the resultset
        If GetRST(mStrSQL).RecordCount Then
        
            If GetRST("vw" & Node.DynamicLevel & "_" & Node.sp_SQL).State Then
                
                tvwMenus.AddNodeByRecordset GetRST("vw" & Node.DynamicLevel & "_" & Node.sp_SQL), 0, 1, , , Node.Key, 1, , , 2, 3, True
                
                If Node.Level > 1 Then
                    Node.ItemData = Node.ChildCount
                End If
            
            End If
            
        Else

            Node.HaveChild = False

        End If

        ' Done
        Node.FirstSelected = True

    End If
    
    WaitTimes 77
    
End Sub

' Purpose: Display resultset into listview if available or display large icon of child into listview

Private Sub tvwMenus_NodeClick(Node As VistaSuitePro.CLS_xpNode)

    On Error Resume Next

    'Ignore if user click current node
    'If m_LastNode = Node.Key Then Exit Sub
    m_LastNode = Node.Key

    ' Display FullPath of Node
    Me.OsenXPForm1.Caption = "Northwind Traders [MS ACCESS] - " & Node.FullPath

    ' Check total child from this node
    If Node.ChildCount > 0 Then

        vList.HeaderAlignment = enAlignLeft
        vList.InsertViewFromNode Node
        vList.ViewMode = 1
    Else

        ' Now prepare query from View of database base on node.key
        mStrSQL = "vw" & Node.DynamicLevel & "_" & Node.sp_SQL
        Debug.Print mStrSQL
        
        If Not GetRST(mStrSQL) Is Nothing Then

            Dim StrA As String
            StrA = ADO_SQL_RESULT("select flags from nodes where nodekey='" & Node.DynamicParentKey & "'")

            If StrA <> "" Then
                ' Conditional formating function here ....
                Dim StrB
                StrB = Split(StrA, "|")
                vList.InsertItemByRecordset GetRST(mStrSQL), , , True, lngTime, CInt(StrB(0)), , CInt(StrB(1)), CInt(StrB(2))
            Else
                vList.InsertItemByRecordset GetRST(mStrSQL), , , True, lngTime

            End If

        Else
            vList.Clear True
        End If

        ' If record(s) not found, show message on the header of listview
        If vList.ListCount = 0 Then

            vList.HeaderAlignment = enAlignCenter
            vList.HeaderCaption = "There are no items to show in this view."

        End If

        ' Get Privileges setting on current node [Menu] '00000'
        Me.OsenXPToolBar1.EnabledButton(8) = Node.CurrentPrivileges(1) ' Check Access to AddNew Records
        Me.OsenXPToolBar1.EnabledButton(9) = Node.CurrentPrivileges(2) ' Check Access to Edit Records
        Me.OsenXPToolBar1.EnabledButton(10) = Node.CurrentPrivileges(3) ' Check Access to Delete Records
        Me.OsenXPToolBar1.EnabledButton(16) = Node.CurrentPrivileges(4) ' Check Access to Print Preview Records
        Me.OsenXPToolBar1.EnabledButton(17) = Node.CurrentPrivileges(5) ' Check Access to Export to Excel
        tvwMenus.HistoryAdd Node.Parent

    End If
 
    WaitTimes 99
    
End Sub

' Prepared History of user navigation

Private Sub tvwMenus_UpdateHistory(EnableBack As Boolean, _
                                   EnableForward As Boolean, _
                                   EnableUpOneLevel As Boolean)

    On Error Resume Next

    OsenXPToolBar1.EnabledButton(4) = EnableBack
    OsenXPToolBar1.EnabledButton(5) = EnableForward
    OsenXPToolBar1.EnabledButton(6) = EnableUpOneLevel

End Sub

' Purpose: Display Current Row and Column

Private Sub vList_CellClick(lrow As Long, _
                            iCol As Integer, _
                            lLeft As Long, _
                            lTop As Long, _
                            lWidth As Long, _
                            lHeight As Long, _
                            Value As String)

    ' Display current cell position [Row,Col]

    OsenXPStatusBar1.PanelCaption(2) = "Ln " & lrow + 1 & " , Col " & iCol + 1

End Sub

'Purpose: there are function as same as edit command

Private Sub vList_DblClick()

    ' This function just available if type of Viewmode=lvwDetail , not LargeIcons
    
    If vList.ViewMode = lvwDetail And tvwMenus.CurrentNode.CurrentPrivileges(2) Then
        OsenXPToolBar1_ButtonClick 9, ""
    End If

End Sub

'Purpose: Hide the progressbar

Private Sub vList_EndProgress()

    pBar.Visible = False

    ' Display total of record(s) and taken time
    OsenXPStatusBar1.PanelCaption(1) = vList.ListCount & " row(s) retrieved [ " & lngTime & " ms taken ]"

End Sub

' Purpose: send message into treeview when icon selected

Private Sub vList_IconClick(Index As Long)

    '    On Error Resume Next
    '
    tvwMenus.SetNodeClick tvwMenus.CurrentNode.Child(Index)

End Sub

'Private Sub vList_IconDblClick(Index As Long)
'
'    On Error Resume Next
'
'        tvwMenus.SetNodeClick tvwMenus.CurrentNode.Child(Index)
'
'End Sub

Private Sub vList_IconEnter(Index As Long)

    On Error Resume Next

    tvwMenus.SetNodeClick tvwMenus.CurrentNode.Child(Index)

End Sub

Private Sub vList_MouseDown(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

    If Button = 2 And vList.ViewMode = lvwDetail Then

        ' Enable this menu by user privileges setting
        MnuAction(1).Enabled = Me.OsenXPToolBar1.EnabledButton(7)
        MnuAction(2).Enabled = Me.OsenXPToolBar1.EnabledButton(8)
        MnuAction(3).Enabled = Me.OsenXPToolBar1.EnabledButton(9)
        MnuAction(9).Enabled = Me.OsenXPToolBar1.EnabledButton(15)
        MnuAction(10).Enabled = Me.OsenXPToolBar1.EnabledButton(16)

        ' Show popup menu
        PopupMenu Mnu_Right

    End If

End Sub

' Purpose: Show progress ...

Private Sub vList_ProgressStatus(ByVal lngProgress As Long)

    pBar.Value = lngProgress

End Sub

' Purpose: raise when start insert item into list view

Private Sub vList_StartProgress()

    pBar.Value = 0
    pBar.Visible = True

End Sub

Public Sub RefreshView()
    If vList.ViewMode = lvwDetail Then
        tvwMenus.SetNodeClick tvwMenus.CurrentNode
    End If
End Sub

Private Sub DeleteRecord(ByVal mTable As String)

    On Error Resume Next

    Select Case mTable

        Case "order"
            
            ADOCN.Execute "delete from `order details` where orderid=" & KeyValue
            ADOCN.Execute "delete from `orders` where orderid=" & KeyValue

        Case "user"
        
            ADOCN.Execute "delete from users where userid='" & KeyValue & "'"

        Case "employee"
        
            ADOCN.Execute "delete from employees where employeeid=" & KeyValue
            
        Case "category"
        
            ADOCN.Execute "delete from categories where categoryid=" & KeyValue
            
        Case "customer"
        
            ADOCN.Execute "delete from customers where customerid='" & KeyValue & "'"
            
        Case "product"
        
            ADOCN.Execute "delete from products where productid=" & KeyValue
            
        Case "supplier"
        
            ADOCN.Execute "delete from suppliers where supplierid=" & KeyValue
            
        Case Else

            DoEvents
    End Select

    ' Refresh current view
    RefreshView
    
End Sub


':) Ulli's VB Code Formatter V2.19.3 (2005-Aug-16 18:05)  Decl: 12  Code: 717  Total: 729 Lines
':) CommentOnly: 94 (12.9%)  Commented: 11 (1.5%)  Empty: 202 (27.7%)  Max Logic Depth: 6


















