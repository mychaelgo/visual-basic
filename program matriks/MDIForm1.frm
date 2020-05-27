VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000006&
   Caption         =   "PrOgRaM MatRiks [MN]"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7590
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
MDIForm1.Visible = False
Form1.Show
End Sub


