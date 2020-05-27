Attribute VB_Name = "mod_Find"
Public rep_Bulan As Byte
Public pubMenu As Boolean
'----------------------------------- Ref to Report anggaran
Public varUserGlobal As String
Public Enum ChiperAlgorithm
       RC4 = 1
       rc2 = 2
       DES = 3
End Enum
Public arrFindForm  As New clsArray
Public arrQueryForm  As New clsArray
Public UnloadStatus As Boolean

Sub RemoveFindItem(nKey As String)
On Error Resume Next
   arrFindForm.RemoveItem nKey
   arrQueryForm.RemoveItem nKey
End Sub

Sub PostFindForm(nKey As String)
On Error Resume Next
   Dim ret As String
   ret = Replace(arrFindForm.GetItem(nKey), "#", "")
   PostMessage Val(ret), &H10, 0, 0
   arrFindForm.RemoveItem nKey
   arrQueryForm.RemoveItem nKey
End Sub

Function ShowFindForm(StrSql As String, nKey As String, obj As Object, Proc As String, Optional SQLadd As String = "")
Dim ret As String
With arrFindForm
    If .CekKey(nKey) Then
       arrQueryForm.SetItem nKey, StrSql
       nKey = Replace(.GetItem(nKey), "#", "")
       ShowWindow Val(nKey), 1
       BringWindowToTop Val(nKey)
       Putfocus Val(nKey)
    Else
       Dim kFind As New frm_util_find
       arrQueryForm.AddItem StrSql, nKey
       .AddItem kFind.hWnd, nKey
       
       With kFind
            .Caption = "Pencarian: " & obj.Caption
            .Tag = nKey
            .ShowField obj, Proc
            .DrgData.Tag = SQLadd
            .Show '0, Obj
            On Error Resume Next
            .Left = 0
            .Top = 0
            .ZOrder 0
       End With
       
    End If
End With
End Function

Function FindInGrid(flex As VSFlexGrid, str2Find As String, col2Find As Long, Optional PassRow As Long = -1) As Boolean
On Error Resume Next
Dim i As Integer, p
For i = 0 To flex.Rows ' - 1
  If flex.TextMatrix(i, col2Find) <> "" Then
      If PassRow = -1 Then
        If flex.TextMatrix(i, col2Find) = str2Find Then
          p = p + 1
        End If
      Else
        If flex.TextMatrix(i, col2Find) = str2Find And i <> PassRow Then
          p = p + 1
        End If
      End If
        DoEvents
  Else
     Exit For
  End If
Next i
If p > 0 Then
    FindInGrid = True
    Exit Function
End If
End Function



Function ShowDlgMsg(Frm As Form, Msg As String, Button As VbMsgBoxStyle, Optional AddDesc As String = "", Optional byPass As Boolean, Optional showCheck As Boolean = True, Optional FontSize As Integer = 8, Optional IconEX As String = "RANDOM", Optional ParentTo As String = "MENU", Optional fForeColor As Long = 0, Optional KeyMe As String) As Boolean
Dim myForm As New frm_util_msg
Dim getKey As String
If showCheck = False Then myForm.Check1.Visible = False
myForm.Status.FontSize = FontSize
If byPass Then
    getKey = 1
Else
   If KeyMe = "" Then
      getKey = GetSetting(vbReg & Frm.name, "Message", "showme", "1")
   Else
      getKey = GetSetting(vbReg & Frm.name, "Message", KeyMe, "1")
   End If
End If
With myForm
    If getKey = 1 Then
        ShowDlgMsg = True
        Load myForm
        .cmd(0).Default = False
        .cmd(1).Default = False
        .cmd(2).Default = False
        .cmd(0).Cancel = False
        .cmd(1).Cancel = False
        .cmd(2).Cancel = False
        .Status.ForeColor = fForeColor
        On Error Resume Next
        If IconEX = "RANDOM" Then
           '.AniGif1.LoadFile App.Path & "\Media\Animation\Mode2\" & Format(Int(33 * Rnd), "0#") & ".gif", False
        Else
          '.AniGif1.LoadFile App.Path & "\Media\Animation\" & IconEX & ".gif", False
        End If
        Randomize
        '.AniGif1.LoadFile StripPath(App.Path) & "Media\Animation\Mode2\sc" & Format(Int(Rnd * 33), "0#") & ".gif", False
        '.AniGif1.SetInterval 100
        
       .Status = Replace(Msg, "<br>", vbCrLf, , , vbTextCompare)
       .Status.Tag = Replace(Msg, "<br>", vbCrLf, , , vbTextCompare)
       .Check1.Value = 1
       If AddDesc <> "" Then
          .Image1.Tag = "Error!" & vbCrLf & vbCrLf & AddDesc
          .cgmHyperLabel1.Visible = True
       Else
          .cgmHyperLabel1.Visible = False
       End If
       
       Select Case Button
              Case vbYesNo
                   .cmd(1).Caption = "&Ya"
                   .cmd(2).Caption = "&Tidak"
                   
                   .cmd(0).Visible = False
                   .cmd(1).Visible = True
                   .cmd(2).Visible = True
                   
                   .cmd(1).Default = True
                   .cmd(2).Cancel = True
                   
              Case vbYesNoCancel
                   .cmd(0).Caption = "&Ya"
                   .cmd(1).Caption = "&Tidak"
                   .cmd(2).Caption = "&Batal"
                   
                   .cmd(0).Visible = True
                   .cmd(1).Visible = True
                   .cmd(2).Visible = True
                   
                   .cmd(0).Default = True
                   .cmd(2).Cancel = True
              Case vbOKOnly Or vbOK
                   .cmd(0).Caption = ""
                   .cmd(1).Caption = ""
                   .cmd(2).Caption = "&OK"
                   
                   .cmd(0).Visible = False
                   .cmd(1).Visible = False
                   .cmd(2).Visible = True
                   
                   .cmd(2).Default = True
                   .cmd(2).Cancel = True
              Case vbOKCancel
                   .cmd(0).Caption = ""
                   .cmd(1).Caption = "&OK"
                   .cmd(2).Caption = "&Batal"
                   
                   .cmd(0).Visible = False
                   .cmd(1).Visible = True
                   .cmd(2).Visible = True
                   
                   .cmd(1).Default = True
                   .cmd(2).Cancel = True
              Case vbRetryCancel
                   .cmd(0).Caption = ""
                   .cmd(1).Caption = "&Coba lagi"
                   .cmd(2).Caption = "&Batal"
                   
                   .cmd(0).Visible = False
                   .cmd(1).Visible = True
                   .cmd(2).Visible = True
                   
                   .cmd(1).Default = True
                   .cmd(2).Cancel = True
              Case vbAbortRetryIgnore
                   .cmd(0).Caption = "&Batalkan"
                   .cmd(1).Caption = "&Coba lagi"
                   .cmd(2).Caption = "&Lanjut"
                   
                   .cmd(0).Visible = True
                   .cmd(1).Visible = True
                   .cmd(2).Visible = True
                   
                   .cmd(1).Default = True
                   .cmd(0).Cancel = True
       End Select
       .Tag = vbReg & Frm.name
       .cgmHyperLabel1.Tag = KeyMe
       'If MENU.Visible Then
       '  .Show 1, mn_005
       'Else
       If ParentTo = "MENU" Then
         .Show 1, MainMenu
       Else
         .Show 1, Frm
       End If
       'End If
    Else
        Unload myForm
    End If
End With
End Function
