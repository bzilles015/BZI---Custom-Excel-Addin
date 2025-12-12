Attribute VB_Name = "modStyles"
 Option Explicit
 
 '==============================================================================
' Module: modStyles (Sxx)
' Purpose:
'   - Font/fill/text/font-color cycles
'   - Input styles & header styles
'   - Layout helpers (zoom, indent, center across)
'   - Zero-check conditional formatting
'==============================================================================

 '------------------------------------------------------------------------------
' S01 - AutoColorSelection  (Ctrl+Alt+A)
    'Colors formulas vs constants in the current selection.
'------------------------------------------------------------------------------
Sub AutoColorSelection()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Application.ScreenUpdating = False

    Dim rF As Range, rN As Range
    On Error Resume Next
    Set rF = Selection.SpecialCells(xlCellTypeFormulas)
    Set rN = Selection.SpecialCells(xlCellTypeConstants, xlNumbers)
    On Error GoTo 0

    Dim rLinks As Range, rSheet As Range, rOps As Range, rPlain As Range, c As Range, f As String

    If Not rF Is Nothing Then
        For Each c In rF.Cells
            f = c.Formula
            If InStr(f, "[") > 0 Then
                Set rLinks = AddToUnion(rLinks, c)
            ElseIf InStr(f, "!") > 0 Then
                Set rSheet = AddToUnion(rSheet, c)
            ElseIf InStr(f, "+") > 0 Or InStr(f, "-") > 0 Or InStr(f, "*") > 0 Or InStr(f, "/") > 0 Or InStr(f, "^") > 0 Then
                Set rOps = AddToUnion(rOps, c)
            Else
                Set rPlain = AddToUnion(rPlain, c)
            End If
        Next c
    End If

    If Not rLinks Is Nothing Then rLinks.Font.Color = RGB(255, 0, 0)   ' external links
    If Not rSheet Is Nothing Then rSheet.Font.Color = RGB(0, 128, 0)   ' sheet refs
    If Not rOps Is Nothing Then rOps.Font.Color = RGB(0, 0, 0)         ' math
    If Not rPlain Is Nothing Then rPlain.Font.Color = RGB(0, 0, 0)     ' other formulas
    If Not rN Is Nothing Then rN.Font.Color = RGB(0, 0, 255)           ' numbers

    Application.ScreenUpdating = True
    LogAction "AutoColor", Selection.Address(False, False)
    RegisterUndo "Auto Color"
End Sub

'helper for AutoColor
Private Function AddToUnion(ByVal acc As Range, ByVal c As Range) As Range
    If acc Is Nothing Then
        Set AddToUnion = c
    Else
        Set AddToUnion = Application.Union(acc, c)
    End If
End Function


'------------------------------------------------------------------------------
' S02 – CycleFont  (Ctrl+')
'     Cycles font family/style for the current selection.
'------------------------------------------------------------------------------

Sub CycleFont()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevFontAddress Then FontCycleIndex = 0
    PrevFontAddress = Selection.Address(False, False)
    Dim fonts As Variant, ni As Long
    fonts = Array("Aptos Narrow", "Poppins", "Times New Roman")
    ni = FontCycleIndex Mod (UBound(fonts) + 1)
    Selection.Font.Name = fonts(ni)
    FontCycleIndex = FontCycleIndex + 1
    LogAction "FontCyc" & (ni + 1), PrevFontAddress
    RegisterUndo "Font Cycle"
End Sub

'------------------------------------------------------------------------------
' S03 - CycleFill  (Ctrl+Shift+K)
   '-Cycle through a list of cell fills
'------------------------------------------------------------------------------
Sub CycleFill()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevFillAddress Then FillCycleIndex = 0
    PrevFillAddress = Selection.Address(False, False)
    Dim items As Variant, ni As Long, cl As Variant
    items = Array( _
      "NoFill", _
      RGB(255, 242, 204), _
      RGB(217, 217, 217), _
      RGB(14, 40, 65), _
      RGB(0, 0, 0), _
      RGB(198, 239, 206), _
      RGB(255, 199, 206), _
      RGB(255, 0, 255) _
    )
    ni = FillCycleIndex Mod (UBound(items) + 1)
    cl = items(ni)
    If VarType(cl) = vbString Then
        Selection.Interior.Pattern = xlNone
    Else
        With Selection.Interior
            .Pattern = xlSolid
            .Color = cl
        End With
    End If
    FillCycleIndex = FillCycleIndex + 1
    LogAction "FillCyc" & (ni + 1), PrevFillAddress
    RegisterUndo "Fill Cycle"
End Sub

'------------------------------------------------------------------------------
' S04 - CycleTextCase  (Ctrl+Alt+Shift+I)
   '-Cycle through cases on selection - easily change things from upper to lower or vice versa
'------------------------------------------------------------------------------
Sub CycleTextCase()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevTextAddress Then TextCaseIndex = 0
    PrevTextAddress = Selection.Address(False, False)
    Dim modes As Variant, ni As Long, c As Range
    modes = Array(vbProperCase, vbLowerCase, vbUpperCase)
    ni = TextCaseIndex Mod (UBound(modes) + 1)
    For Each c In Selection: c.Value = VBA.StrConv(c.Value, modes(ni)): Next
    TextCaseIndex = TextCaseIndex + 1
    LogAction "TextCase" & (ni + 1), PrevTextAddress
    RegisterUndo "Text Case"
End Sub

'------------------------------------------------------------------------------
' S05 - CycleFontColor  (Ctrl+Shift+C)
   '-Cycle through font colors - normal financial colors
'------------------------------------------------------------------------------
Sub CycleFontColor()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevFontColorAddress Then FontColorCycleIndex = 0
    PrevFontColorAddress = Selection.Address(False, False)
    Dim cols As Variant, ni As Long
    cols = Array( _
      RGB(0, 0, 255), _
      RGB(0, 128, 0), _
      RGB(0, 0, 0), _
      RGB(255, 0, 0), _
      RGB(127, 127, 127), _
      RGB(112, 48, 160), _
      RGB(255, 255, 255) _
    )
    ni = FontColorCycleIndex Mod (UBound(cols) + 1)
    Selection.Font.Color = cols(ni)
    FontColorCycleIndex = FontColorCycleIndex + 1
    LogAction "FontColorCyc" & (ni + 1), PrevFontColorAddress
    RegisterUndo "Font Color Cycle"
End Sub


'------------------------------------------------------------------------------
' S06 – ZoomIn  (Ctrl+Alt+Shift+=)
'     Increase active window zoom level.
'------------------------------------------------------------------------------

Sub ZoomIn()
    BeginMacroWithUndo
    With ActiveWindow: .Zoom = Application.Min(.Zoom + 10, 400): End With
    LogAction "ZoomIn", ""
    RegisterUndo "Zoom In"
End Sub

'------------------------------------------------------------------------------
' S07 – ZoomOut  (Ctrl+Alt+Shift+-)
'     Decrease active window zoom level.
'------------------------------------------------------------------------------

Sub ZoomOut()
    BeginMacroWithUndo
    With ActiveWindow: .Zoom = Application.Max(.Zoom - 10, 10): End With
    LogAction "ZoomOut", ""
    RegisterUndo "Zoom Out"
End Sub


'------------------------------------------------------------------------------
' S08 - CycleInputStyle  (Ctrl+Alt+Shift+U)
'     Cycles selection through Yellow / Gray / Peach input styles.
'------------------------------------------------------------------------------

Sub CycleInputStyle()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Static idx As Integer, prev As String
    If Selection.Address(False, False) <> prev Then idx = 0
    prev = Selection.Address(False, False)

    Select Case (idx Mod 3)
        Case 0: ApplyInputStyle Selection, RGB(255, 242, 204), RGB(0, 0, 255), RGB(0, 0, 0)          ' Yellow + Blue font + Black border
        Case 1: ApplyInputStyle Selection, RGB(217, 217, 217), RGB(0, 0, 255), RGB(0, 0, 0)          ' Gray + Blue font + Black border
        Case 2: ApplyInputStyle Selection, RGB(255, 204, 153), RGB(0, 133, 178), RGB(127, 127, 127)  ' Peach + Teal font + Gray border
    End Select

    idx = idx + 1
    LogAction "InputStyle", prev
    RegisterUndo "Input Style"
End Sub

' Helper - Input - Yellow
Sub ApplyInputYellow()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    ApplyInputStyle Selection, RGB(255, 242, 204), RGB(0, 0, 255), RGB(0, 0, 0)
    LogAction "InputYellow", Selection.Address(False, False)
    RegisterUndo "Input Yellow"
End Sub

' Helper - Input - Gray
Sub ApplyInputGray()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    ApplyInputStyle Selection, RGB(217, 217, 217), RGB(0, 0, 255), RGB(0, 0, 0)
    LogAction "InputGray", Selection.Address(False, False)
    RegisterUndo "Input Gray"
End Sub

' Helper - Input - Peach
Sub ApplyInputPeach()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    ApplyInputStyle Selection, RGB(255, 204, 153), RGB(0, 133, 178), RGB(127, 127, 127)
    LogAction "InputPeach", Selection.Address(False, False)
    RegisterUndo "Input Peach"
End Sub


' ==== Core Input Style ====
Private Sub ApplyInputStyle(ByVal rng As Range, ByVal fillColor As Long, ByVal fontColor As Long, ByVal borderColor As Long)
    Dim parts As Variant, i As Long

    If rng Is Nothing Then Exit Sub

    ' Font + Fill (leave NumberFormat alone)
    rng.Font.Color = fontColor
    With rng.Interior
        .Pattern = xlSolid
        .Color = fillColor
        .TintAndShade = 0
    End With

    ' Clear existing borders
    rng.Borders.LineStyle = xlNone

    ' Apply dotted hairline box (outline + inside)
    parts = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                  xlInsideVertical, xlInsideHorizontal)

    For i = LBound(parts) To UBound(parts)
        With rng.Borders(parts(i))
            .LineStyle = xlDot
            .Weight = xlHairline
            .Color = borderColor
            .TintAndShade = 0
        End With
    Next i
End Sub


'------------------------------------------------------------------------------
' S10 - CycleHeaderStyle  (Ctrl+Alt+Shift+H)
'     Insert headers for Financial Modeling Easily
'------------------------------------------------------------------------------
Sub CycleHeaderStyle()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Static idx As Integer, prev As String
    If Selection.Address(False, False) <> prev Then idx = 0
    prev = Selection.Address(False, False)
    Select Case (idx Mod 4)
        Case 0: ApplyHeaderStyle Selection, RGB(14, 40, 65)
        Case 1: ApplyHeaderStyle Selection, RGB(68, 84, 106)
        Case 2: ApplyHeaderStyle Selection, RGB(0, 0, 0)
        Case 3: ApplyHeaderStyle Selection, RGB(68, 114, 196)
    End Select
    idx = idx + 1
    LogAction "HeaderStyle", prev
    RegisterUndo "Header Style"
End Sub

Private Sub ApplyHeaderStyle(ByVal rng As Range, ByVal fillColor As Long)
    With rng
        .Interior.Pattern = xlSolid
        .Interior.Color = fillColor
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .NumberFormat = "General"
    End With
End Sub

'------------------------------------------------------------------------------
' S11 - InsertHeadersFromPrompt  (Ctrl+Alt+Shift+Y)
'     Insert headers from prompt
'------------------------------------------------------------------------------

Sub InsertHeadersFromPrompt()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim s As String, parts() As String, i As Long
    s = InputBox("Enter comma-separated headers (e.g., 2024A,2025B,2026E):", _
                 "Insert Headers", "2024A,2025B,2026E")
    If Len(Trim$(s)) = 0 Then Exit Sub
    parts = Split(s, ",")
    For i = LBound(parts) To UBound(parts)
        Selection.Cells(1, i + 1).Value = Trim$(parts(i))
    Next i
    ApplyHeaderStyle Selection.Resize(1, UBound(parts) - LBound(parts) + 1), RGB(14, 40, 65)
    LogAction "InsertHeaders", Selection.Address(False, False)
    RegisterUndo "Insert Headers"
End Sub

'------------------------------------------------------------------------------
' S12 - InsertVarianceHeaders  (Ctrl+Alt+Shift+D)
'     Insert normal variance headers
'------------------------------------------------------------------------------
Sub InsertVarianceHeaders()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim labels As Variant
    labels = Array("AvF %?", "AvB%?", "Var AvB")
    Selection.Resize(1, 3).Value = labels
    ApplyHeaderStyle Selection.Resize(1, 3), RGB(14, 40, 65)
    LogAction "InsertVarHdrs", Selection.Address(False, False)
    RegisterUndo "Insert Variance Headers"
End Sub

'------------------------------------------------------------------------------
' S13 - CenterAcrossSelection  (Ctrl+Alt+E)
'        -Best practice for centering - makes it a quick keyboardshort cut. the menu is burried otherwise
'------------------------------------------------------------------------------
Sub CenterAcrossSelection()
    BeginMacroWithUndo
    Selection.HorizontalAlignment = xlCenterAcrossSelection
    LogAction "CenterAcross", Selection.Address(False, False)
    RegisterUndo "Center Across"
End Sub



'------------------------------------------------------------------------------
' S14 - IncreaseFontSize  (Ctrl+Shift+F)
'------------------------------------------------------------------------------
Sub IncreaseFontSize()
    BeginMacroWithUndo
    Selection.Font.Size = Selection.Font.Size + 1
    LogAction "FontInc", Selection.Address(False, False)
    RegisterUndo "Increase Font Size"
End Sub

'------------------------------------------------------------------------------
' S15 - DecreaseFontSize  (Ctrl+Shift+G)
'------------------------------------------------------------------------------
Sub DecreaseFontSize()
    BeginMacroWithUndo
    Selection.Font.Size = Application.Max(1, Selection.Font.Size - 1)
    LogAction "FontDec", Selection.Address(False, False)
    RegisterUndo "Decrease Font Size"
End Sub


'------------------------------------------------------------------------------
' S16 - IndentIn  (Ctrl+Shift+])
'------------------------------------------------------------------------------
Sub IndentIn()
    BeginMacroWithUndo
    Selection.IndentLevel = Application.Min(Selection.IndentLevel + 1, 15)
    LogAction "IndentIn", Selection.Address(False, False)
    RegisterUndo "Indent In"
End Sub

'------------------------------------------------------------------------------
' S17 - IndentOut  (Ctrl+Shift+[)
'------------------------------------------------------------------------------
Sub IndentOut()
    BeginMacroWithUndo
    Selection.IndentLevel = Application.Max(Selection.IndentLevel - 1, 0)
    LogAction "IndentOut", Selection.Address(False, False)
    RegisterUndo "Indent Out"
End Sub

'------------------------------------------------------------------------------
' S18 - InsertStaticNow  (Ctrl+Shift+N)
'------------------------------------------------------------------------------
Sub InsertStaticNow()
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection: c.Value = Now: Next c
    LogAction "InsertNow", Selection.Address(False, False)
    RegisterUndo "Insert Static Now"
End Sub

'------------------------------------------------------------------------------
' S19 - PasteValuesKeepFormat  (Ctrl+Alt+Shift+V)
'------------------------------------------------------------------------------
Sub PasteValuesKeepFormat()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Selection.Value = Selection.Value
    LogAction "ValsKeepFmt", Selection.Address(False, False)
    RegisterUndo "Values (Keep Format)"
End Sub

'------------------------------------------------------------------------------
' S20 ApplyZeroCheckCF  – Green if =0, Red if <>0 (selection only)
'------------------------------------------------------------------------------
Sub ApplyZeroCheckCF()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim area As Range
    Application.ScreenUpdating = False
    For Each area In Selection.Areas
        ApplyZeroCheckToArea area
    Next area
    Application.ScreenUpdating = True
    LogAction "ZeroCheckCF_Apply", Selection.Address(False, False)
    RegisterUndo "Zero-Check CF"
End Sub


'helper for check zero function
Private Sub ApplyZeroCheckToArea(ByVal rng As Range)
    Dim tl As Range, fEq As String, fNe As String, i As Long
    Set tl = rng.Cells(1, 1)
    fEq = "=" & tl.Address(False, False) & "=0"
    fNe = "=" & tl.Address(False, False) & "<>0"
    For i = rng.FormatConditions.Count To 1 Step -1
        With rng.FormatConditions(i)
            If .Type = xlExpression Then
                If .Formula1 = fEq Or .Formula1 = fNe Then .Delete
            End If
        End With
    Next i
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fEq)
        .Interior.Color = RGB(198, 239, 206)
        .StopIfTrue = True
    End With
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fNe)
        .Interior.Color = RGB(255, 199, 206)
    End With
End Sub

'------------------------------------------------------------------------------
' S21 - ClearZeroCheckCF  – Removes the two zero-check rules from selection
'------------------------------------------------------------------------------
Sub ClearZeroCheckCF()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim area As Range, tl As Range, fEq As String, fNe As String, i As Long
    Application.ScreenUpdating = False
    For Each area In Selection.Areas
        Set tl = area.Cells(1, 1)
        fEq = "=" & tl.Address(False, False) & "=0"
        fNe = "=" & tl.Address(False, False) & "<>0"
        For i = area.FormatConditions.Count To 1 Step -1
            With area.FormatConditions(i)
                If .Type = xlExpression Then
                    If .Formula1 = fEq Or .Formula1 = fNe Then .Delete
                End If
            End With
        Next i
    Next area
    Application.ScreenUpdating = True
    LogAction "ZeroCheckCF_Clear", Selection.Address(False, False)
    RegisterUndo "Clear Zero-Check CF"
End Sub

