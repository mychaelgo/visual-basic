VERSION 5.00
Begin VB.Form frm_util_report_pop 
   Caption         =   "Form1"
   ClientHeight    =   -285
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   -285
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuExport 
      Caption         =   "&Export"
      Begin VB.Menu mnuExports 
         Caption         =   "Ms. Word Format"
         Index           =   0
      End
      Begin VB.Menu mnuExports 
         Caption         =   "Ms. Excel Format"
         Index           =   1
      End
      Begin VB.Menu mnuExports 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuExports 
         Caption         =   "Hyper Text Format (HTML)"
         Index           =   3
      End
      Begin VB.Menu mnuExports 
         Caption         =   "Portable Document (PDF)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frm_util_report_pop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj As Object
Sub SetOBJ(myOBJ As Object)
Set obj = myOBJ
End Sub

Private Sub mnuExports_Click(Index As Integer)
Select Case Index
       Case 0
            obj.dlgExportRpt "rtf"
       Case 1
            obj.dlgExportRpt "xls"
       Case 3
            obj.dlgExportRpt "html"
       Case 4
            obj.dlgExportRpt "pdf"
End Select
End Sub

