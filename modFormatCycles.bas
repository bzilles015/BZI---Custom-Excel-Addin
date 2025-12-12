Attribute VB_Name = "modFormatCycles"
 Option Explicit

'==============================================================================
' Module: modFormatCycles (Fxx)
' Purpose:
'   Number/date/percent/currency/other cycles, decimals, and scale in/out.
' Bound shortcuts (see modBindings):
'   F01 – CycleNumberFormat      Ctrl+Shift+1
'   F02 – CycleDateFormat        Ctrl+Shift+3
'   F03 – CyclePercentFormat     Ctrl+Shift+5
'   F04 – CycleCurrencyFormat    Ctrl+Shift+4
'   F05 – CycleOtherNumbers      Ctrl+Shift+8
'   F06 – IncreaseDecimal        Ctrl+Shift+.
'   F07 – DecreaseDecimal        Ctrl+Shift+,
'   F08 – ScaleUp                Alt+Shift+<
'   F09 – ScaleDown              Alt+Shift+>
'==============================================================================



'------------------------------------------------------------------------------
' F01 CycleNumberFormat  (Ctrl+Shift+1)
'------------------------------------------------------------------------------
Sub CycleNumberFormat()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevNumAddress Then FormatNumIndex = 0
    PrevNumAddress = Selection.Address(False, False)
    Dim fmts As Variant, ni As Long
    fmts = Array( _
      "#,##0_);(#,##0);""--"";@", _
      "#,##0,_);(#,##0,);""--"";@", _
      "#,##0,""K""_);(#,##0,""K"");""--"";@", _
      "#,##0.0,,_);(#,##0.0,,);""--"";@", _
      "#,##0.0,,""M""_);(#,##0.0,,""M"");""--"";@" _
    )
    ni = FormatNumIndex Mod (UBound(fmts) + 1)
    Selection.NumberFormat = fmts(ni)
    FormatNumIndex = FormatNumIndex + 1
    LogAction "NumFmt" & (ni + 1), PrevNumAddress
    RegisterUndo "Number Format"
End Sub

'------------------------------------------------------------------------------
' F02 CycleDateFormat  (Ctrl+Shift+3)
'------------------------------------------------------------------------------
Sub CycleDateFormat()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevDateAddress Then FormatDateIndex = 0
    PrevDateAddress = Selection.Address(False, False)
    Dim dates As Variant, ni As Long
    dates = Array("m/d/yyyy", "m/d/yy", "mmm-yy", "d-mmm-yy;d-mmm-yy;-")
    ni = FormatDateIndex Mod (UBound(dates) + 1)
    Selection.NumberFormat = dates(ni)
    FormatDateIndex = FormatDateIndex + 1
    LogAction "DateFmt" & (ni + 1), PrevDateAddress
    RegisterUndo "Date Format"
End Sub

'------------------------------------------------------------------------------
' F03 CyclePercentFormat  (Ctrl+Shift+5)
'------------------------------------------------------------------------------
Sub CyclePercentFormat()
    On Error GoTo CleanFail
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevPctAddress Then FormatPctIndex = 0
    PrevPctAddress = Selection.Address(False, False)
    Dim fmts(1 To 5) As String
    fmts(1) = "0.0%;(0.0%);""—"";@"
    fmts(2) = "0%;(0%);""—"";@"
    fmts(3) = "+0.0%;-0.0%;""—"";@"
    fmts(4) = "[<=-0.0005](0.0%);[>=0.0005]0.0%;"""";@"
    fmts(5) = "0.0%;(0.0%);"""";@"
    Dim ni As Long
    ni = (FormatPctIndex Mod UBound(fmts)) + 1
    Selection.NumberFormat = fmts(ni)
    FormatPctIndex = FormatPctIndex + 1
    LogAction "PctFmt:" & CStr(ni), PrevPctAddress
    RegisterUndo "Percent Format"
CleanExit:
    Exit Sub
CleanFail:
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------
' F04 CycleCurrencyFormat  (Ctrl+Shift+4)
'------------------------------------------------------------------------------
Sub CycleCurrencyFormat()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevCurAddress Then FormatCurIndex = 0
    PrevCurAddress = Selection.Address(False, False)
    Dim fmts As Variant, ni As Long
    fmts = Array( _
      "$#,##0_);($#,##0);""--"";@", _
      "$#,##0,_);($#,##0,);""--"";@", _
      "$#,##0,""K""_);($#,##0,""K"");""--"";@", _
      "$#,##0.0,,_);($#,##0.0,,);""--"";@", _
      "$#,##0.0,,""M""_);($#,##0.0,,""M"");""--"";@" _
    )
    ni = FormatCurIndex Mod (UBound(fmts) + 1)
    Selection.NumberFormat = fmts(ni)
    FormatCurIndex = FormatCurIndex + 1
    LogAction "CurFmt" & (ni + 1), PrevCurAddress
    RegisterUndo "Currency Format"
End Sub

'------------------------------------------------------------------------------
' F05 CycleOtherNumbers  (Ctrl+Shift+8)
'------------------------------------------------------------------------------
Sub CycleOtherNumbers()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevOtherAddress Then FormatOtherIndex = 0
    PrevOtherAddress = Selection.Address(False, False)
    Dim fmts As Variant, ni As Long
    fmts = Array("0\A", "0\B", "0\F", """Q""#", "0\P", "0\E", "0.0""x""")
    ni = FormatOtherIndex Mod (UBound(fmts) + 1)
    Selection.NumberFormat = fmts(ni)
    FormatOtherIndex = FormatOtherIndex + 1
    LogAction "OtherFmt" & (ni + 1), PrevOtherAddress
    RegisterUndo "Other Numbers Format"
End Sub

'------------------------------------------------------------------------------
' F06 IncreaseDecimal (Ctrl+Shift+.)
'------------------------------------------------------------------------------

Sub IncreaseDecimal()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Application.ScreenUpdating = False
    AdjustDecimalsInSelection Selection, 1
    Application.ScreenUpdating = True
    LogAction "IncreaseDecimal", Selection.Address(False, False)
    RegisterUndo "Increase Decimal"
End Sub

'------------------------------------------------------------------------------
' F07 DecreaseDecimal (Ctrl+Shift+,)
'------------------------------------------------------------------------------
Sub DecreaseDecimal()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Application.ScreenUpdating = False
    AdjustDecimalsInSelection Selection, -1
    Application.ScreenUpdating = True
    LogAction "DecreaseDecimal", Selection.Address(False, False)
    RegisterUndo "Decrease Decimal"
End Sub

' Helper: cache formats so we don't recompute for every single cell
Private Sub AdjustDecimalsInSelection(ByVal rng As Range, ByVal delta As Long)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim c As Range, fmt As String, newFmt As String

    For Each c In rng.Cells
        fmt = CStr(c.NumberFormat)
        If Not dict.Exists(fmt) Then dict(fmt) = AdjustSectionDecimalsOne(fmt, delta)
        newFmt = dict(fmt)
        If c.NumberFormat <> newFmt Then c.NumberFormat = newFmt
    Next c
End Sub


'------------------------------------------------------------------------------
' F08 ScaleUp  (Alt+Shift+<)
'------------------------------------------------------------------------------
Sub ScaleUp()
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection
        If c.HasFormula Then
            c.Formula = "=(" & Mid(c.Formula, 2) & ")/1000"
        ElseIf IsNumeric(c.Value) Then
            c.Value = c.Value / 1000
        End If
    Next c
    LogAction "ScaleUp", Selection.Address(False, False)
    RegisterUndo "Scale Up"
End Sub

'------------------------------------------------------------------------------
' F09 ScaleDown  (Alt+Shift+>)
'------------------------------------------------------------------------------
Sub ScaleDown()
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection
        If c.HasFormula Then
            c.Formula = "=(" & Mid(c.Formula, 2) & ")*1000"
        ElseIf IsNumeric(c.Value) Then
            c.Value = c.Value * 1000
        End If
    Next c
    LogAction "ScaleDown", Selection.Address(False, False)
    RegisterUndo "Scale Down"
End Sub



'------------------------------------------------------------------------------
' IncreaseDecimalFormat / DecreaseDecimalFormat  (legacy helpers)
'------------------------------------------------------------------------------
Private Function IncreaseDecimalFormat(fmt As String) As String
    Dim parts As Variant, mainFmt As String, rest As String
    Dim dp As Long, pos As Long, i As Long
    parts = Split(fmt, ";")
    mainFmt = parts(0)
    rest = ""
    For i = 1 To UBound(parts)
        rest = rest & ";" & parts(i)
    Next i
    dp = InStr(mainFmt, ".")
    If dp > 0 Then
        pos = dp + 1
        Do While pos <= Len(mainFmt) And (Mid(mainFmt, pos, 1) = "0" Or Mid(mainFmt, pos, 1) = "#")
            pos = pos + 1
        Loop
        mainFmt = Left(mainFmt, pos - 1) & "0" & Mid(mainFmt, pos)
    Else
        mainFmt = mainFmt & ".0"
    End If
    IncreaseDecimalFormat = mainFmt & rest
End Function

Private Function DecreaseDecimalFormat(fmt As String) As String
    Dim parts As Variant, mainFmt As String, rest As String
    Dim dp As Long, pos As Long, i As Long
    parts = Split(fmt, ";")
    mainFmt = parts(0)
    rest = ""
    For i = 1 To UBound(parts)
        rest = rest & ";" & parts(i)
    Next i
    dp = InStr(mainFmt, ".")
    If dp > 0 Then
        pos = dp + 1
        Do While pos <= Len(mainFmt) And (Mid(mainFmt, pos, 1) = "0" Or Mid(mainFmt, pos, 1) = "#")
            pos = pos + 1
        Loop
        If pos - dp - 1 > 0 Then
            mainFmt = Left(mainFmt, pos - 2) & Mid(mainFmt, pos)
            If Right(mainFmt, 1) = "." Then mainFmt = Left(mainFmt, Len(mainFmt) - 1)
        End If
    End If
    DecreaseDecimalFormat = mainFmt & rest
End Function


' Helper: cache formats so we don't recompute for every single cell
Private Function AdjustSectionDecimalsOne(ByVal sec As String, ByVal delta As Long) As String
    Dim lastDigit As Long, dotPos As Long, p As Long, lastDec As Long
    lastDigit = LastDigitIndex(sec)
    If lastDigit = 0 Then
        AdjustSectionDecimalsOne = sec
        Exit Function
    End If
    dotPos = InStrRev(sec, ".", lastDigit)
    If delta > 0 Then
        If dotPos > 0 Then
            AdjustSectionDecimalsOne = Left$(sec, lastDigit) & "0" & Mid$(sec, lastDigit + 1)
        Else
            AdjustSectionDecimalsOne = Left$(sec, lastDigit) & ".0" & Mid$(sec, lastDigit + 1)
        End If
    Else
        If dotPos = 0 Then
            AdjustSectionDecimalsOne = sec
        Else
            p = dotPos + 1
            Do While p <= Len(sec) And (Mid$(sec, p, 1) = "0" Or Mid$(sec, p, 1) = "#")
                lastDec = p
                p = p + 1
            Loop
            If lastDec = 0 Then
                AdjustSectionDecimalsOne = sec
            ElseIf lastDec = dotPos + 1 Then
                AdjustSectionDecimalsOne = Left$(sec, dotPos - 1) & Mid$(sec, dotPos + 2)
            Else
                AdjustSectionDecimalsOne = Left$(sec, lastDec - 1) & Mid$(sec, lastDec + 1)
            End If
        End If
    End If
End Function

Private Function LastDigitIndex(ByVal sec As String) As Long
    Dim p0 As Long, pH As Long
    p0 = InStrRev(sec, "0")
    pH = InStrRev(sec, "#")
    If p0 >= pH Then LastDigitIndex = p0 Else LastDigitIndex = pH
End Function


