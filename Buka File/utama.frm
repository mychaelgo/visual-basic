VERSION 5.00
Begin VB.Form utama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silahkan Pilih Program Yang Ingin Anda Install..."
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13785
   Icon            =   "utama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cd 
      Caption         =   "Jelajahi DVD"
      DisabledPicture =   "utama.frx":1272
      DownPicture     =   "utama.frx":13F6
      Height          =   375
      Left            =   120
      Picture         =   "utama.frx":157A
      TabIndex        =   41
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Frame Mal 
      Caption         =   "Mal"
      Height          =   3495
      Left            =   3120
      TabIndex        =   34
      Top             =   6720
      Width           =   2655
      Begin VB.CommandButton atm 
         Caption         =   "Copy Folder ATM FOTO"
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton system 
         Caption         =   "Copy Folder SYSTEM"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton copy_dongbin 
         Caption         =   "Copy Folder Dongbin"
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton back 
         Caption         =   "Copy Background"
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton copy 
         Caption         =   "Copy Folder Data"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame software 
      Caption         =   "Software"
      Height          =   8295
      Left            =   5760
      TabIndex        =   13
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton s_player 
         Caption         =   "Install SPlayer"
         Height          =   495
         Left            =   5040
         TabIndex        =   74
         Top             =   5880
         Width           =   2295
      End
      Begin VB.CommandButton nero7 
         Caption         =   "Install Nero 7"
         Height          =   495
         Left            =   5040
         TabIndex        =   73
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton s_2003 
         Caption         =   "Serial"
         Height          =   495
         Left            =   6720
         TabIndex        =   72
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton o_2003 
         Caption         =   "Install Office2003"
         Height          =   495
         Left            =   5040
         TabIndex        =   71
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton r230 
         Caption         =   "Install Driver R230 Windows 7"
         Height          =   495
         Left            =   5040
         TabIndex        =   70
         Top             =   4080
         Width           =   2295
      End
      Begin VB.CommandButton transtool 
         Caption         =   "Install TransTool 9"
         Height          =   495
         Left            =   5040
         TabIndex        =   69
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton teracopy 
         Caption         =   "Install TeraCopy Pro 2.12"
         Height          =   495
         Left            =   5040
         TabIndex        =   68
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton notepad 
         Caption         =   "Install NotePad ++"
         Height          =   495
         Left            =   5040
         TabIndex        =   67
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton idm 
         Caption         =   "Install IDM 5.18"
         Height          =   495
         Left            =   5040
         TabIndex        =   64
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton easeus 
         Caption         =   "Install EASEUS Partion Magic"
         Height          =   495
         Left            =   5040
         TabIndex        =   63
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton daemon 
         Caption         =   "Install Daemon Tools"
         Height          =   495
         Left            =   5040
         TabIndex        =   62
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton evr_soft 
         Caption         =   "Install Evr Soft 2006"
         Height          =   495
         Left            =   2520
         TabIndex        =   61
         Top             =   7680
         Width           =   2295
      End
      Begin VB.CommandButton visual 
         Caption         =   "Install Visual Task Tips"
         Height          =   495
         Left            =   120
         TabIndex        =   56
         Top             =   7680
         Width           =   2295
      End
      Begin VB.CommandButton indo 
         Caption         =   "Install Windows XP Indonesia"
         Height          =   495
         Left            =   2520
         TabIndex        =   55
         Top             =   7080
         Width           =   2295
      End
      Begin VB.CommandButton pdf 
         Caption         =   "Install Nitro PDF"
         Height          =   495
         Left            =   2520
         TabIndex        =   53
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton pdf_creatorplus 
         Caption         =   "Install PDF Creator Plus"
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   7080
         Width           =   2295
      End
      Begin VB.CommandButton video_convert 
         Caption         =   "Install McFunSoft Video Convert Master 8.0.24"
         Height          =   495
         Left            =   2520
         TabIndex        =   47
         Top             =   6480
         Width           =   2295
      End
      Begin VB.CommandButton avira 
         Caption         =   "Install Avira Antivirus 9"
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   6480
         Width           =   2295
      End
      Begin VB.CommandButton zip_repair 
         Caption         =   "Install ZIP TAR RAR Repair"
         Height          =   495
         Left            =   2520
         TabIndex        =   45
         Top             =   5880
         Width           =   2295
      End
      Begin VB.CommandButton pdf_unlock 
         Caption         =   "Install PDF Unlocker"
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   5880
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Install nLite"
         Height          =   495
         Left            =   2520
         TabIndex        =   43
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton wucd 
         Caption         =   "Install Windows Unattend CD"
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton cifree 
         Caption         =   "Install Create Install Free"
         Height          =   495
         Left            =   2520
         TabIndex        =   40
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CommandButton ok 
         Caption         =   "Install O&&K Print Watch"
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CommandButton office 
         Caption         =   "Install Office 2007"
         Height          =   495
         Left            =   2520
         TabIndex        =   32
         Top             =   4080
         Width           =   2295
      End
      Begin VB.CommandButton flash3d 
         Caption         =   "Install 3D Flash Animator"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   4080
         Width           =   2295
      End
      Begin VB.CommandButton cd_burner 
         Caption         =   "Install CD Burner XP"
         Height          =   495
         Left            =   2520
         TabIndex        =   30
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton deep 
         Caption         =   "Install Deep Freeze 6"
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton ssc 
         Caption         =   "Install SSC Service 4.30"
         Height          =   495
         Left            =   2520
         TabIndex        =   28
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton tema 
         Caption         =   "Install Tema Royale"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton flash_player 
         Caption         =   "Install Flash Player"
         Height          =   495
         Left            =   2520
         TabIndex        =   20
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton mozilla 
         Caption         =   "Install Mozilla Firefox 3.5"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton izarc4 
         Caption         =   "Install Izarc 4"
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton reader9 
         Caption         =   "Install Adobe Reader 9"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton tuneup2010 
         Caption         =   "Install Tune Up 2010"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton dotnet35 
         Caption         =   "Install .NET Framework 3.5"
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton install45 
         Caption         =   "Install Windows Installer 4.5"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame design 
      Caption         =   "Design"
      Height          =   6495
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton html_creator 
         Caption         =   "Install HTML Creator"
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton dream_cs4 
         Caption         =   "Install DreamWeaver CS4"
         Height          =   495
         Left            =   120
         TabIndex        =   57
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Install Adobe Photoshop CS2"
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton photoshop80 
         Caption         =   "Install Adobe Photoshop 8.0"
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   5040
         Width           =   2295
      End
      Begin VB.CommandButton photo70 
         Caption         =   "Install Adobe Photoshop 7.0"
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   4440
         Width           =   2295
      End
      Begin VB.CommandButton ilustrasi 
         Caption         =   "Install Adobe Illustrator CS3"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton flash 
         Caption         =   "Install Macromedia Flash 8"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton corel 
         Caption         =   "Install Corel Draw X4"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cs4 
         Caption         =   "Install Photoshop CS 4"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label ins 
         AutoSize        =   -1  'True
         Caption         =   "Install:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   4200
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Portable:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame programming 
      Caption         =   "Programming"
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton c 
         Caption         =   "Install Borland C++"
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   5880
         Width           =   2775
      End
      Begin VB.CommandButton sql_server 
         Caption         =   "Install SQL SERVER 2000"
         Height          =   495
         Left            =   120
         TabIndex        =   66
         Top             =   5280
         Width           =   2775
      End
      Begin VB.CommandButton filezilla 
         Caption         =   "Install FileZilla Client"
         Height          =   495
         Left            =   120
         TabIndex        =   65
         Top             =   4680
         Width           =   2775
      End
      Begin VB.CommandButton cdkey_fox 
         Caption         =   "CD KEY"
         Height          =   495
         Left            =   2280
         TabIndex        =   60
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton fox9 
         Caption         =   "Install Microsoft Visual Fox Pro 9"
         Height          =   495
         Left            =   120
         TabIndex        =   59
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton virus_vb 
         Caption         =   "Ekstrak Kode Virus VB"
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   6480
         Width           =   2775
      End
      Begin VB.CommandButton vbscroll 
         Caption         =   "VB Scroll"
         Height          =   495
         Left            =   2280
         TabIndex        =   26
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cdkey8 
         Caption         =   "CD KEY"
         Height          =   495
         Left            =   2160
         TabIndex        =   25
         Top             =   3480
         Width           =   735
      End
      Begin VB.CommandButton delphixp 
         Caption         =   "Delphi XP"
         Height          =   495
         Left            =   2160
         TabIndex        =   24
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cdkeydelphi 
         Caption         =   "CD KEY"
         Height          =   495
         Left            =   1440
         TabIndex        =   23
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton vbxp 
         Caption         =   "VB XP"
         Height          =   495
         Left            =   1560
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cdkeyvb 
         Caption         =   "CD KEY"
         Height          =   495
         Left            =   840
         TabIndex        =   21
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton xampp 
         Caption         =   "Install XAMPP"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   2775
      End
      Begin VB.CommandButton delphi7 
         Caption         =   "Install Delphi 7"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton dream8 
         Caption         =   "Install Macromedia Dreamweaver 8"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CommandButton java_neatbeans 
         Caption         =   "Install Java dan Netbeans 6.7"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton vb 
         Caption         =   "Install VB6"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton pascal 
         Caption         =   "Install Turbo Pascal 7"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
  "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
  String, ByVal lpszFile As String, ByVal lpszParams As String, _
  ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Const SW_SHOWNORMAL = 1
Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&
Public Sub OpenDirectory(Directory As String)
      ShellExecute 0, "Open", Directory, vbNullString, _
        vbNullString, SW_SHOWNORMAL
End Sub
Function OpenDocument(ByVal DocName As String) As Long
   Dim Scr_hDC As Long
   'Scr_hDC = GetDesktopWindow()
   OpenDocument = ShellExecute(Me.hwnd, "Open", DocName, _
          "", "C:\", SW_SHOWNORMAL)
End Function


Private Sub atm_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFolder App.Path & "\mal\atm foto", "e:\ATM FOTO", True
MsgBox "Folder Sudah Ter-Copy", vbInformation, "Sukses..."
End Sub

Private Sub avira_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\avira.zip")
End Sub

Private Sub back_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFolder App.Path & "\mal\background", "e:\BACKGROUND", True
MsgBox "Background Sudah Ter-Copy", vbInformation, "Sukses..."
End Sub

Private Sub c_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\Borland C++ 5.02.zip")
End Sub

Private Sub cd_burner_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\cd burner xp.zip")
End Sub

Private Sub cd_Click()

OpenDirectory (App.Path)
End Sub

Private Sub cdkey_fox_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Programming\Microsoft Visual FoxPro 9.0\InstallNotes.txt")
End Sub

Private Sub cdkey8_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\macromedia\keygen.exe")
End Sub

Private Sub cdkeydelphi_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\delphi7\key.txt")
End Sub



Private Sub cdkeyvb_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\vb6\cd key.txt")
End Sub
Private Sub cifree_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\cifree.zip")
End Sub

Private Sub Command1_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\nLite-1.4.9.1.zip")
End Sub

Private Sub Command2_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Adobe\photoshop\Adobe Photoshop CS2.7z")
End Sub



Private Sub copy_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFolder App.Path & "\mal\data", "D:\ps_digital\data", True
MsgBox "File Sudah Ter-Copy", vbInformation, "Sukses..."
End Sub

Private Sub copy_dongbin_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFolder App.Path & "\mal\dongbin", "e:\DONGBIN", True
MsgBox "Folder Sudah Ter-Copy", vbInformation, "Sukses..."
End Sub

Private Sub corel_Click()
Dim r As Long
r = OpenDocument(App.Path & "\portable\corel draw x4.zip")
End Sub



Private Sub cs3_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Adobe\photoshop\Adobe Photoshop CS3.7z")
End Sub

Private Sub cs4_Click()
Dim r As Long
r = OpenDocument(App.Path & "\portable\setup photoshop cs4.zip")
End Sub

Private Sub daemon_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Daemon Tools.zip")
End Sub

Private Sub deep_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\deep freeze.zip")
End Sub

Private Sub delphi7_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\delphi7\install.exe")
End Sub

Private Sub delphixp_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFile App.Path & "\programming\delphi7\delphi32.exe.manifest", "C:\Program Files\Borland\Delphi7\Bin\delphi32.exe.manifest"
MsgBox "File Sudah Ter-Copy", vbInformation, "Sukses..."
End Sub

Private Sub dotnet35_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\dotnetfx35.zip")
End Sub

Private Sub dream_cs4_Click()
Dim r As Long
r = OpenDocument(App.Path & "\portable\Dreamweaver CS4.zip")
End Sub

Private Sub dream8_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\macromedia\dreamweaver8.exe")
End Sub

Private Sub easeus_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\EASEUS Partition Master 4.1.1 Professional Edition.zip")
End Sub

Private Sub evr_soft_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\evrsoft.zip")
End Sub

Private Sub filezilla_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Programming\FileZilla 3.3.1.zip")
End Sub

Private Sub flash_Click()
Dim r As Long
r = OpenDocument(App.Path & "\portable\setup flash.zip")
End Sub

Private Sub flash_player_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\flash player.zip")
End Sub

Private Sub flash3d_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\3DFlashAnimatorsetup.zip")
End Sub

Private Sub Form_Load()
On Error Resume Next
End Sub

Private Sub fox9_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Programming\Microsoft Visual FoxPro 9.0\setup.exe")
End Sub

Private Sub html_creator_Click()
Dim r As Long
r = OpenDocument(App.Path & "\portable\HTML CREATOR.zip")
End Sub

Private Sub idm_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\idm 5.18 build 5.zip")
End Sub

Private Sub ilustrasi_Click()
Dim r As Long
r = OpenDocument(App.Path & "\portable\ilustrasi.zip")
End Sub

Private Sub indo_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\windows xp indonesia.zip")
End Sub

Private Sub install45_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\windowsinstaller 4.5.zip")
End Sub

Private Sub izarc4_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\izarc.zip")
End Sub

Private Sub java_neatbeans_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\java & netbeans.exe")
End Sub

Private Sub mozilla_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\firefox.zip")
End Sub

Private Sub nero7_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\nero 7\setupx.exe")
End Sub

Private Sub notepad_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Notepad++5.6.2.zip")
End Sub

Private Sub o_2003_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\2003\setup.exe")
End Sub

Private Sub office_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\2007\setup.exe")
End Sub

Private Sub ok_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\O&K Print Watch.zip")
End Sub

Private Sub pascal_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\setup pascal.zip")
End Sub

Private Sub pdf_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Nitro PDF Professional.zip")
End Sub


Private Sub pdf_creatorplus_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Membuat file PDF.rar")
End Sub

Private Sub pdf_unlock_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\PDF_Unlocker_v2.0.rar")
End Sub

Private Sub photo70_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Adobe\photoshop\Adobe Photoshop 7.0.7z")
End Sub

Private Sub photoshop80_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Adobe\photoshop\Adobe Photoshop 8.0.7z")
End Sub

Private Sub r230_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\R230 WIN 7.zip")
End Sub

Private Sub reader9_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Adobe\Reader\ADBERDR9.zip")
End Sub

Private Sub s_2003_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\2003\cd key.txt")
End Sub

Private Sub s_player_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\splayer.zip")
End Sub



Private Sub sql_server_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\Microsoft SQL 2000 Personal Edition.zip")
End Sub

Private Sub ssc_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\sscserve 4.30.zip")
End Sub

Private Sub system_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFolder App.Path & "\mal\system", "e:\SYSTEM", True
MsgBox "Folder Sudah Ter-Copy", vbInformation, "Sukses..."
End Sub

Private Sub tema_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFolder App.Path & "\royale noir", "C:\WINDOWS\Resources\Themes\royale", True
MsgBox "File Sudah Ter-Copy", vbInformation, "Sukses..."
Dim r As Long
r = OpenDocument("C:\WINDOWS\Resources\Themes\royale\luna.msstyles")
End Sub

Private Sub teracopy_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\TeraCopy Pro v2.12.zip")
End Sub

Private Sub transtool_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Software\Transtool9.zip")
End Sub



Private Sub tuneup2010_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\Tune up 2010.zip")
End Sub

Private Sub vb_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\vb6\setup.exe")
End Sub

Private Sub vbscroll_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\vb6\scroll.exe")
End Sub

Private Sub vbxp_Click()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fso.CopyFile App.Path & "\programming\vb6\vb6.exe.manifest", "C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe.manifest"
MsgBox "File Sudah Ter-Copy", vbInformation, "Sukses..."
End Sub

Private Sub video_convert_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\McFunSoft Video Convert Master 8.0.24.zip")
End Sub

Private Sub virus_vb_Click()
Dim r As Long
r = OpenDocument(App.Path & "\Programming\Coding Virus in VB.zip")
End Sub

Private Sub visual_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\VisualTaskTips.zip")
End Sub

Private Sub wucd_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\wucd.zip")
End Sub

Private Sub xampp_Click()
Dim r As Long
r = OpenDocument(App.Path & "\programming\xampp.zip")
End Sub

Private Sub zip_repair_Click()
Dim r As Long
r = OpenDocument(App.Path & "\software\ZIP CAB Tar RAR Repair.rar")

End Sub
