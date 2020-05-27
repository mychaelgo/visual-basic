VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3480
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i  As Integer
Dim f, a, s
Dim fso As FileSystemObject
Set fso = New FileSystemObject

Drive1.Drive = "d:"
Dir1.Path = "\ps_digital\photosaved\"

For i = 0 To Dir1.ListCount
    Set f = fso.GetFolder(Dir1.List(i))
    s = f.Size
    a = f.DateCreated
    
        If s = 0 Then
            fso.DeleteFolder Dir1.List(i), True
        ElseIf Format(a, "dd") = Format(Now, "dd") And Format(a, "mm") <> Format(Now, "mm") Then
            fso.DeleteFolder Dir1.List(i), True
        End If
Next

Drive1.Drive = "d:"
Dir1.Path = "\ps_digital\photosaved\"

For i = 0 To Dir1.ListCount
    Set f = fso.GetFolder(Dir1.List(i))
    s = f.Size
    a = f.DateCreated
    
        If s = 0 Then
            fso.DeleteFolder Dir1.List(i), True
        ElseIf Format(a, "dd") = Format(Now, "dd") And Format(a, "mm") <> Format(Now, "mm") Then
            fso.DeleteFolder Dir1.List(i), True
        End If
Next


Drive1.Drive = "d:"
Dir1.Path = "\foto\"

For i = 0 To Dir1.ListCount
    Set f = fso.GetFolder(Dir1.List(i))
    s = f.Size
    a = f.DateCreated
    
        If s = 0 Then
            fso.DeleteFolder Dir1.List(i), True
        ElseIf Format(a, "dd") = Format(Now, "dd") And Format(a, "mm") <> Format(Now, "mm") Then
            fso.DeleteFolder Dir1.List(i), True
        End If
Next
End
End Sub
