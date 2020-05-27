VERSION 5.00
Begin VB.UserControl vbTextBoxMulti 
   BackColor       =   &H000040C0&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   2280
   ToolboxBitmap   =   "cgmMultiTextBox.ctx":0000
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   1980
      Top             =   60
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgDown 
      Height          =   210
      Left            =   855
      MouseIcon       =   "cgmMultiTextBox.ctx":0312
      MousePointer    =   99  'Custom
      Picture         =   "cgmMultiTextBox.ctx":0464
      Top             =   30
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   600
      Left            =   645
      Top             =   645
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1470
      Left            =   0
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "vbTextBoxMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iValue

Private vMask As Boolean
Private vMaskSeperator As String
Private vMaskFormat As String
Private vMaskSpace As String

Private aDataLabel As DataBinding
Private oValue As String

Private fFocusBackColor As OLE_COLOR
Private fFocusForeColor As OLE_COLOR
Private fFocusBackMainColor As OLE_COLOR
Private fFocusBorderColor As OLE_COLOR

Private fBackColor As OLE_COLOR
Private fBackColorMain As OLE_COLOR
Private fForeColor As OLE_COLOR
Private oldColor As OLE_COLOR
Private fAutoTab As Boolean

Private minVal As Long
Private maxVal As Long

Private aLabel As String

Private mDataType As DataTypes

'Public Enum FontStyle
'       [tNormal] = 0
'       [tUpperCase] = 1
'       [tLowerCase] = 2
'       [tProperCase] = 3
'       [tReverseCase] = 4
'End Enum

Public Enum FontFormat
       [tNormal] = 0
       [tDate] = 1
       [tCurrency2] = 2
       [tNumeric] = 3
End Enum


Private fUpperCase As FontStyle
Private fFFormat As FontFormat

'Public Enum DataTypes
'    [tText] = 0
'    [tInteger] = 1
'    [tLong] = 2
'    [tPositive] = 3
'    [tNegative] = 4
'    [tCurrency] = 5
'    [tSingle] = 6
'    [tDouble] = 7
'    [tDecimal] = 8
'    [tByte] = 9
'End Enum
Private m_Picture As Picture

'EVENTS
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event DownButtonClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fFontBold As Boolean

Private Sub imgDown_Click()
RaiseEvent DownButtonClick
End Sub

Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDown.BorderStyle = 1
End Sub

Private Sub imgDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDown.BorderStyle = 0
End Sub

Private Sub UserControl_Initialize()
    Text1.ZOrder 0
    Text1.Text = ""
    fFocusBackColor = Text1.BackColor
    fFocusForeColor = Text1.ForeColor
    aLabel = ""
    oldColor = vbWhite
End Sub

Private Sub Text1_GotFocus()
    If fFFormat = tCurrency2 Then
       Text1.Text = rNum(Text1.Text)
    ElseIf fFFormat = tDate Then
       Text1.Text = fDate(Text1.Text)
    ElseIf fFFormat = tNumeric Then
       If Not IsNumeric(Text1.Text) And Len(Text1.Text) > 0 Then
          Text1.Text = Val(Text1.Text)
       Else
          Text1.Text = Replace(Text1.Text, ",", ".")
       End If
    End If

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Text1.BackColor = FocusBackColor
    Text1.ForeColor = FocusForeColor
    Text1.FontBold = True
    UserControl.BackColor = FocusBackMainColor
    Shape1.BorderColor = FocusBorderColor
    Shape1.BackColor = FocusBorderColor
    Shape2.BorderColor = FocusBorderColor
    Shape2.BackColor = FocusBorderColor
    
    oValue = Text1.Text
    AssociatedLabelFontBold True
    BlokX Text1, 0, True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If fAutoTab And KeyAscii = 13 Then keybd_event &H9, 0, 0, 0
    If InStr("'`|~", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        
    Select Case mDataType
        ' ignores case [0 - Text] but validates any other data type
        Case 1 To 9: If Not IsValidCharacter(KeyAscii) Then KeyAscii = 0
    End Select
    
    RaiseEvent KeyPress(KeyAscii)
    If Mask Then
        Select Case KeyAscii
               Case vbKeyDelete
                    If Chr(KeyAscii) = MaskSeperator Then
                        Text1.SelStart = Text1.SelStart + 1
                        Text1.SelLength = 1
                        Text1.SelText = MaskSeperator
                        Text1.SelStart = Text1.SelStart + 1
                        KeyAscii = 0
                    Else
                        If Text1.SelStart > Len(MaskFormat) - 1 Then: KeyAscii = 0: Exit Sub
                        Text1.SelStart = Text1.SelStart + 1
                        Text1.SelLength = 1
                        Text1.SelText = MaskSpace
                        Text1.SelStart = Text1.SelStart + 1
                        KeyAscii = 0
                               
                    End If
               Case vbKeyBack
                    If Text1.SelStart = 0 Then Exit Sub
                    If Mid(Text1.Text, Text1.SelStart, 1) = MaskSeperator Then
                        Text1.SelStart = Text1.SelStart - 1
                        Text1.SelLength = 1
                        Text1.SelText = MaskSeperator
                        Text1.SelStart = Text1.SelStart - 1
                        KeyAscii = 0
                    Else
                        If Text1.SelStart = 0 Then Exit Sub
                        Text1.SelStart = Text1.SelStart - 1
                        Text1.SelLength = 1
                        Text1.SelText = MaskSpace
                        Text1.SelStart = Text1.SelStart - 1
                        KeyAscii = 0
                    End If
               Case Else
                    If KeyAscii <> 13 Then
                    If Text1.SelStart > Len(MaskFormat) - 1 Then KeyAscii = 0: Exit Sub
                    If Trim(Text1) = "" Then Text1 = Chr(KeyAscii) & Mid(MaskFormat, 2)
                    If Mid(Text1, Text1.SelStart + 1, 1) = MaskSeperator Then Text1.SelStart = Text1.SelStart + 1
                    Text1.SelLength = 1
                    Text1.SelText = Chr(KeyAscii)
                    If Mid(Text1, Text1.SelStart + 1, 1) = MaskSeperator Then
                       Text1.SelStart = Text1.SelStart + 1
                    End If
                    KeyAscii = 0
                    End If
        End Select
    End If
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Text1_LostFocus()
Dim isValidEntry As Boolean
Dim I As Boolean
Dim Msg As String
    '
    ' Validates input
    ' If Min/Max is set to 0(default) the validation is performed acording to the datatype
    ' otherwise validation is based on Min/Max values
    BlokX Text1, 0, False
    On Error GoTo Text1_LostFocus
    ' ignore empty stings
    If Trim(Text1.Text) <> "" Then
        If fUpperCase = tLowerCase Then
           Text1.Text = LCase(Text1.Text)
        ElseIf fUpperCase = tUpperCase Then
           Text1.Text = UCase(Text1.Text)
        ElseIf fUpperCase = tReverseCase Then
           Text1.Text = StrReverse(Text1.Text)
        ElseIf fUpperCase = tProperCase Then
           Text1.Text = StrConv(Text1.Text, vbProperCase)
        End If
        If fFFormat = tCurrency2 Then
           Text1.Text = fNum(Text1.Text)
        ElseIf fFFormat = tDate Then
           Text1.Text = fDate(Text1.Text)
        ElseIf fFFormat = tNumeric And Len(Text1.Text) > 0 Then
           If Not IsNumeric(Text1.Text) Then
              Text1.Text = Val(Text1.Text)
           Else
              Text1.Text = Replace(Text1.Text, ",", ".")
           End If
        End If
        isValidEntry = True ' be positive
        ' is min/max set to anything other then 0 ?
        If (MinValue = 0 And MaxValue = 0) And (mDataType <> 0) Then
            ' check input based on current data type
            ' note. cPos and cNeg are user defined function that check
            ' for positive and negative and raises an error if required
            Select Case mDataType
                Case 1: I = CInt(Text1.Text)
                Case 2: I = CLng(Text1.Text)
                Case 3: I = cPos(Text1.Text)
                Case 4: I = cNeg(Text1.Text)
                Case 5: I = CCur(Text1.Text)
                Case 6: I = CSng(Text1.Text)
                Case 7: I = CDbl(Text1.Text)
                Case 8: I = CDec(Text1.Text)
                Case 9: I = CByte(Text1.Text)
            End Select
        Else
            ' check min/max range
            Msg = ""
            If mDataType <> 0 Then
                If (CLng(Text1.Text) < minVal) Then
                    Msg = "Invalid entry. Minimum allowed is " & minVal & " !"
                ElseIf (CLng(Text1.Text) > maxVal) Then
                    Msg = "Invalid entry. Max allowed is " & maxVal & " !"
                End If
                If Msg <> "" Then
                    'MsgBox Msg, vbExclamation
                    Text1.SetFocus
                    GoTo Text1_LostFocus_Exit
                End If
            End If
        End If
    End If
    ' reset fore and backcolor
    Text1.BackColor = fBackColor
    Text1.ForeColor = fForeColor
    Text1.FontBold = 1
    UserControl.BackColor = fBackColorMain
    Shape1.BorderColor = BorderColor
    Shape1.BackColor = BorderColor
    Shape2.BorderColor = BorderColor
    Shape2.BackColor = BorderColor
    'Text1.BackColor = fForeColor
    AssociatedLabelFontBold False
Text1_LostFocus_Exit:
    Exit Sub
    
Text1_LostFocus:
    If Err.Number <> 0 Then
        If Err.Number = 6 Or Err.Number = 13 Then
            'MsgBox "Invalid entry! " & str(mDataType), vbExclamation
            Text1.SetFocus
            Resume Text1_LostFocus_Exit
        Else
            'MsgBox Err.Number & ":" & Err.Description
            Text1.SetFocus
            Resume Text1_LostFocus_Exit
        End If
    End If
End Sub

Private Sub AssociatedLabelFontBold(newValue As Boolean)
On Error Resume Next
Dim I As Long
    For I = 0 To Parent.Controls.Count - 1
        If LCase(Parent.Controls(I).Name) = LCase(aLabel) Then
            Parent.Controls(I).FontBold = newValue
            Exit Sub
        End If
    Next
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

' resise to fit
Private Sub UserControl_Resize()
On Error Resume Next
Shape1.Move 0, 0, UserControl.Width, UserControl.Height
Shape2.Move 0, 0, 250, UserControl.Height
Text1.Move 30, 50, UserControl.Width - (70 + IIf(imgDown.Visible, imgDown.Width + 10, 0)), UserControl.Height - 90
imgDown.Left = UserControl.Width - (imgDown.Width + 30)
End Sub

' Define PROPERTIES for custom control
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    If newValue Then
       Shape1.BorderColor = oldColor
       Shape1.BackColor = oldColor
       Shape2.BorderColor = oldColor
       Shape2.BackColor = oldColor
    Else
       Shape1.BorderColor = &HE0E0E0
       Shape1.BackColor = &HE0E0E0
       Shape2.BorderColor = &HE0E0E0
       Shape2.BackColor = &HE0E0E0
    End If
    PropertyChanged "Enabled"
    'Text1.Enabled = NewValue
End Property

Public Function Hwnd1() As Long
   Hwnd1 = Text1.hWnd
End Function
Public Function Hwnd2() As Long
   Hwnd2 = UserControl.hWnd
End Function

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = Text1.Text
End Property
Public Property Let Text(ByVal newValue As String)
    PropertyChanged "Text"
    Text1.Text = newValue
End Property

Public Property Set Font(ByVal newValue As StdFont)
    Set Text1.Font = newValue
    PropertyChanged "Font"
    fFontBold = newValue.Bold
End Property

Public Property Get Font() As StdFont
    Set Font = Text1.Font
End Property

Public Property Get FontName() As String
    FontName = Text1.FontName
    
End Property

Public Property Let FontName(ByVal newValue As String)
    Text1.FontName = newValue
    PropertyChanged "FontName"
End Property

Public Property Get DownButton() As Boolean
    DownButton = imgDown.Visible
End Property

Public Property Let DownButton(ByVal newValue As Boolean)
    imgDown.Visible = newValue
    PropertyChanged "DownButton"
    UserControl_Resize
End Property
Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property
Public Property Let Locked(ByVal newValue As Boolean)
    Text1.Locked = newValue
    PropertyChanged "Locked"
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = Text1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal newValue As Boolean)
    Text1.FontStrikethru = newValue
    PropertyChanged "FontStrikethru"
End Property

Public Property Let FontSize(newValue As Single)
    Text1.FontSize = newValue
    PropertyChanged "FontSize"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = fForeColor
End Property

Public Property Let ForeColor(ByVal newValue As OLE_COLOR)
    Text1.ForeColor = newValue
    fForeColor = newValue
    PropertyChanged "ForeColor"
End Property

Public Property Get FontSize() As Single
    FontSize = Text1.FontSize
End Property

Public Property Get MaxLength() As Long
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal newValue As Long)
    Text1.MaxLength = newValue
    PropertyChanged "MaxLength"
End Property

Public Property Get SelStart() As Long
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal newValue As Long)
    Text1.SelStart = newValue
End Property

Public Property Get SelLength() As Long
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal newValue As Long)
    Text1.SelLength = newValue
End Property

Public Property Get SelText() As String
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal newValue As String)
    Text1.SelText = newValue
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = oldColor
    Shape1.BackColor = oldColor
End Property
Public Property Let BorderColor(ByVal newValue As OLE_COLOR)
    oldColor = newValue
    Shape1.BorderColor = newValue
    Shape1.BackColor = newValue
    Shape2.BorderColor = newValue
    Shape2.BackColor = newValue
    PropertyChanged "BorderColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = fBackColor 'Text1.BackColor
End Property
Public Property Let BackColor(ByVal newValue As OLE_COLOR)
    Text1.BackColor = newValue
    fBackColor = newValue
    PropertyChanged "BackColor"
End Property

Public Property Get BackColorMain() As OLE_COLOR
    BackColorMain = fBackColorMain 'Text1.BackColor
End Property
Public Property Let BackColorMain(ByVal newValue As OLE_COLOR)
    UserControl.BackColor = newValue
    fBackColorMain = newValue
    PropertyChanged "BackColorMain"
End Property

Public Property Get FocusBorderColor() As OLE_COLOR
    FocusBorderColor = fFocusBorderColor
End Property
Public Property Let FocusBorderColor(ByVal newValue As OLE_COLOR)
    fFocusBorderColor = newValue
    PropertyChanged "FocusBorderColor"
End Property

Public Property Get FocusBackMainColor() As OLE_COLOR
    FocusBackMainColor = fFocusBackMainColor
End Property
Public Property Let FocusBackMainColor(ByVal newValue As OLE_COLOR)
    fFocusBackMainColor = newValue
    PropertyChanged "FocusBackMainColor"
End Property

Public Property Get FocusBackColor() As OLE_COLOR
    FocusBackColor = fFocusBackColor
End Property
Public Property Let FocusBackColor(ByVal newValue As OLE_COLOR)
    fFocusBackColor = newValue
    PropertyChanged "FocusBackColor"
End Property

Public Property Get FocusForeColor() As OLE_COLOR
    FocusForeColor = fFocusForeColor
End Property

Public Property Let FocusForeColor(ByVal newValue As OLE_COLOR)
    fFocusForeColor = newValue
    PropertyChanged "FocusForeColor"
End Property

Public Property Get AssociatedLabel() As String
    AssociatedLabel = aLabel
End Property
Public Property Let AssociatedLabel(ByVal newValue As String)
    aLabel = newValue
    PropertyChanged "AssociatedLabel"
End Property

Public Property Get PasswordChar() As String
    PasswordChar = Text1.PasswordChar
End Property
Public Property Let PasswordChar(ByVal newValue As String)
    Text1.PasswordChar = newValue
    PropertyChanged "PasswordChar"
End Property

Private Property Get OldValue() As String
    oValue = Text1.Text
End Property

Private Property Let OldValue(ByVal newValue As String)
    oValue = newValue
    PropertyChanged "OldValue"
End Property

Public Property Get AutoTab() As Boolean
    AutoTab = fAutoTab
End Property
Public Property Let AutoTab(ByVal New_Alignment As Boolean)
    fAutoTab = New_Alignment
    PropertyChanged "AutoTab"
End Property
Public Property Get FontFormat() As FontFormat
    FontFormat = fFFormat
End Property
Public Property Let FontFormat(New_Alignment As FontFormat)
    fFFormat = New_Alignment
    PropertyChanged "FontFormat"
End Property

Public Property Get UpperCase() As FontStyle
    UpperCase = fUpperCase
End Property
Public Property Let UpperCase(New_Alignment As FontStyle)
    fUpperCase = New_Alignment
    PropertyChanged "UpperCase"
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = Text1.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get MultiLine() As Boolean
       MultiLine = Text1.MultiLine
End Property

Public Property Let MultiLine(ByVal newValue As Boolean)
    'Text1.MultiLine() = NewValue
    PropertyChanged "MultiLine"
End Property

Public Property Get MinValue() As Long
    MinValue = minVal
End Property

Public Property Let MinValue(ByVal newValue As Long)
    ' checks that the NewValue is Less or equal to MaxValue
    If DataType <> 0 Then
        If newValue > MaxValue And maxVal <> 0 Then
            'MsgBox "Invalid Property Value. Min > Max", vbExclamation
        Else
            minVal = newValue
        End If
        PropertyChanged "MinValue"
    End If
End Property

Public Property Get MaxValue() As Long
    MaxValue = maxVal
End Property

Public Property Let MaxValue(ByVal newValue As Long)
    ' checks that the NewValue is greater or equal to MinValue
    ' ignores Min/Max if datatype is set to Text
    If DataType <> 0 Then
        If newValue < MinValue And MinValue <> 0 Then
            'MsgBox "Invalid Property Value. Max < Min", vbExclamation
        Else
            maxVal = newValue
        End If
        
        PropertyChanged "MaxValue"
    End If
End Property

Public Property Get DataType() As DataTypes
    DataType = mDataType
End Property

Public Property Let DataType(ByVal vNewValue As DataTypes)
    ' if datatype is Text sets the Allignment and resets the Min/Max values
    ' ignores Min/Max if datatype is set to Text
    mDataType = vNewValue
    UserControl.PropertyChanged "DataType"
    If mDataType <> 0 Then
        'Text1.Alignment = 1 ' Right justify
    Else
        'Text1.Alignment = 0 ' Left justify for text only
        MinValue = 0
        MaxValue = 0
    End If
End Property

' Read and Write Properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set Icon = .ReadProperty("Icon", Image1.Picture)
        Text1.Alignment = .ReadProperty("Alignment", 0)
        Text1.PasswordChar = .ReadProperty("PasswordChar", "")
        AssociatedLabel = .ReadProperty("AssociatedLabel", aLabel)
        BackColor = .ReadProperty("BackColor", Text1.BackColor)
        BackColorMain = .ReadProperty("BackColorMain", UserControl.BackColor)
        BorderColor = .ReadProperty("BorderColor", Shape1.BackColor)
        DataType = .ReadProperty("DataType", 0)
        Enabled = .ReadProperty("Enabled", True)
        AutoTab = .ReadProperty("AutoTab", False)
        UpperCase = .ReadProperty("UpperCase", 0)
        FontFormat = .ReadProperty("FontFormat", 0)
        DownButton = .ReadProperty("DownButton", False)
        Locked = .ReadProperty("Locked", False)
        FocusBackColor = .ReadProperty("FocusBackColor", Text1.BackColor)
        FocusForeColor = .ReadProperty("FocusForeColor", Text1.ForeColor)
        FocusBackMainColor = .ReadProperty("FocusBackMainColor", UserControl.BackColor)
        FocusBorderColor = .ReadProperty("FocusBorderColor", Shape1.BorderColor)
        FontBold = .ReadProperty("FontBold", False)
        FontItalic = .ReadProperty("FontItalic", False)
        FontName = .ReadProperty("FontName", "Ms Sans Serif")
        FontSize = .ReadProperty("FontSize", 8)
        FontStrikethru = .ReadProperty("FontStrikethru", False)
        FontUnderline = .ReadProperty("FontUnderline", False)
        ForeColor = .ReadProperty("ForeColor", Text1.ForeColor)
        MaxLength = .ReadProperty("MaxLength", 0)
        
        Mask = .ReadProperty("Mask", False)
        MaskSeperator = .ReadProperty("MaskSeperator", "")
        MaskFormat = .ReadProperty("MaskFormat", "")
        MaskSpace = .ReadProperty("MaskSpace", "")
        
        MinValue = .ReadProperty("MinValue", 0)
        
        'MultiLine = .ReadProperty("MultiLine", Text1.MultiLine)
        MultiLine = .ReadProperty("MultiLine", False)
        OldValue = .ReadProperty("AssociatedLabel", oValue)
        Text = .ReadProperty("Text", Text1.Text)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Icon", m_Picture, Image1.Picture
        .WriteProperty "Alignment", Text1.Alignment, 0
        .WriteProperty "PasswordChar", Text1.PasswordChar, ""
        .WriteProperty "AssociatedLabel", aLabel, ""
        .WriteProperty "BackColor", fBackColor
        .WriteProperty "BackColorMain", fBackColorMain
        .WriteProperty "DownButton", DownButton
        .WriteProperty "BorderColor", BorderColor
        .WriteProperty "DataType", DataType, 0
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "Locked", Locked, False
        .WriteProperty "AutoTab", fAutoTab, False
        .WriteProperty "UpperCase", fUpperCase, 0
        .WriteProperty "FontFormat", fFFormat, 0
        .WriteProperty "FocusBackColor", fFocusBackColor
        .WriteProperty "FocusForeColor", fFocusForeColor
        .WriteProperty "FocusBackMainColor", fFocusBackMainColor
        .WriteProperty "FocusBorderColor", fFocusBorderColor
        .WriteProperty "FontBold", FontBold, False
        .WriteProperty "FontItalic", FontItalic
        .WriteProperty "FontName", FontName, "Ms Sans Serif"
        .WriteProperty "FontSize", FontSize, 8
        .WriteProperty "FontStrikethru", FontStrikethru, False
        .WriteProperty "FontUnderline", FontUnderline, False
        .WriteProperty "ForeColor", ForeColor
        .WriteProperty "MaxLength", MaxLength, 0
        .WriteProperty "MinValue", minVal, 0
        .WriteProperty "MaxValue", maxVal, 0
        
        .WriteProperty "Mask", vMask, False
        
        .WriteProperty "MaskSeperator", vMaskSeperator, ""
        .WriteProperty "MaskFormat", vMaskFormat, ""
        .WriteProperty "MaskSpace", vMaskSpace, ""
        
        .WriteProperty "MultiLine", MultiLine, False
        .WriteProperty "Text", Text, 0
    End With
End Sub

' Add Events Handlers

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    'If Button = 2 Then PopupMenu frmMenu.mnufile
    ''|
    ''|
    ''/
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    Select Case KeyCode
           Case vbKeyTab, vbKeyInsert
                If KeyCode = vbKeyTab Then
                    If imgDown.Visible And Shift = 2 Then
                       imgDown.BorderStyle = 0
                       RaiseEvent DownButtonClick
                    End If
                ElseIf KeyCode = vbKeyInsert Then
                    If imgDown.Visible Then
                       imgDown.BorderStyle = 0
                       RaiseEvent DownButtonClick
                    End If
                End If
            Case vbKeyUp
                 'SendKeys "+{TAB}"
                 keybd_event vbKeyTab, 0, 1, 0
            Case vbKeyDown
                 keybd_event vbKeyTab, 0, 0, 0
    End Select
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case vbKeyEscape
                Text1.Text = oValue: Exit Sub
           Case vbKeyTab, vbKeyInsert
                If KeyCode = vbKeyTab Then
                    If imgDown.Visible And Shift = 2 Then
                       imgDown.BorderStyle = 1
                    End If
                ElseIf KeyCode = vbKeyInsert Then
                    If imgDown.Visible Then
                       imgDown.BorderStyle = 1
                    End If
                End If
                
    End Select
    RaiseEvent KeyDown(KeyCode, Shift)
    If Mask Then
        Select Case KeyCode
               Case vbKeyDelete
                    If Text1.SelStart > Len(MaskFormat) - 1 Then Exit Sub
                    If Text1.SelLength = Len(Text1) Then
                       Text1 = MaskFormat
                       KeyCode = 0
                    Else
                        If Mid(Text1.Text, Text1.SelStart + 1, 1) = MaskSeperator Then
                              Text1.SelLength = 1
                              Text1.SelText = MaskSeperator
                              KeyCode = 0
                          Else
                              Text1.SelLength = 1
                              Text1.SelText = MaskSpace
                              KeyCode = 0
                        End If
                    End If
        End Select
    End If
End Sub
Private Sub Text1_Change()
    RaiseEvent Change
End Sub
Private Sub Text1_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub Text1_Click()
    RaiseEvent Click
End Sub


' Private Function
Private Function IsValidCharacter(KeyAscii As Integer) As Boolean
' check key pressed. returns false if not in range of numbers$
    Const Numbers$ = "0123456789-.,"
    IsValidCharacter = True
    If KeyAscii <> 8 Then
        If (InStr(Numbers, Chr(KeyAscii)) = 0) Then
            IsValidCharacter = False
            Exit Function
        End If
    End If
End Function
Private Function cPos(txt) As Boolean
    ' checks that the value of txt is positive. if not raised an error
    If Not CLng(txt) >= 0 Then Err.Raise 6
End Function
Private Function cNeg(txt) As Boolean
    ' checks that the value of txt is negative. if not raised an error
    If Not CLng(txt) <= 0 Then Err.Raise 6
End Function
Private Sub UserControl_GotFocus()
If Text1.Enabled And Text1.Visible Then
   Text1.SetFocus
End If
End Sub
Public Property Get Icon() As Picture
    Set Icon = m_Picture
End Property
Public Property Set Icon(ByVal newValue As Picture)
    Set m_Picture = newValue
    Set imgDown.Picture = newValue
    PropertyChanged "Icon"
End Property

Public Property Get Mask() As Boolean
    Mask = vMask
End Property
Public Property Let Mask(ByVal New_MAsk As Boolean)
    vMask = New_MAsk
    PropertyChanged "Mask"
End Property
Public Property Get MaskSeperator() As String
    MaskSeperator = vMaskSeperator
End Property
Public Property Let MaskSeperator(ByVal New_MAsk As String)
    vMaskSeperator = New_MAsk
    PropertyChanged "MaskSeperator"
End Property
Public Property Get MaskFormat() As String
    MaskFormat = vMaskFormat
End Property
Public Property Let MaskFormat(ByVal New_MAsk As String)
    vMaskFormat = New_MAsk
    PropertyChanged "MaskFormat"
End Property
Public Property Get MaskSpace() As String
    MaskSpace = vMaskSpace
End Property
Public Property Let MaskSpace(ByVal New_MAsk As String)
    vMaskSpace = New_MAsk
    PropertyChanged "MaskSpace"
End Property

