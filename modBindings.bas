Attribute VB_Name = "modBindings"

Option Explicit

'==============================================================================
' Module: modBindings
' Purpose:
'   Auto-load and bind all keyboard shortcuts for the add-in.
'   Keep bindings grouped by module so maintenance is painless.
'==============================================================================

' Runs when the add-in loads
Sub Auto_Open()
    ResetCycleState    ' Ensure all cycles start from their first item
    BindAllKeys        ' Bind every shortcut
End Sub

Private Sub Workbook_Activate()
    BindAllKeys        ' Reassert mappings when the add-in regains focus
End Sub

' Binds every shortcut (and disables nuisance keys)
Sub BindAllKeys()
    On Error Resume Next

    '==== Disable nuisance keys ===============================================
    Application.OnKey "{F1}", ""            ' Disable F1 (Help)
    Application.OnKey "{SCROLLLOCK}", ""    ' Disable Scroll Lock
    Application.OnKey "{NUMLOCK}", ""       ' Disable Num Lock
    Application.OnKey "{INSERT}", ""        ' Disable Insert

    '==== modCore (Cxx) – performance + helpers ===============================
    Application.OnKey "^%+M", "'" & ThisWorkbook.Name & "'!TogglePerformanceMode"          ' Ctrl+Alt+Shift+M
    Application.OnKey "^%+A", "'" & ThisWorkbook.Name & "'!MakeRefsAbsolute"              ' Ctrl+Alt+Shift+A
    Application.OnKey "^%+R", "'" & ThisWorkbook.Name & "'!MakeRefsRelative"              ' Ctrl+Alt+Shift+R
    Application.OnKey "^%+N", "'" & ThisWorkbook.Name & "'!GoToNextBlank"                 ' Ctrl+Alt+Shift+N
    Application.OnKey "^%+E", "'" & ThisWorkbook.Name & "'!GoToNextError"                 ' Ctrl+Alt+Shift+E
    Application.OnKey "^%+L", "'" & ThisWorkbook.Name & "'!BreakExternalLinksInSelection" ' Ctrl+Alt+Shift+L

    '==== modFormatCycles (Fxx) ===============================================
    ' Format cycles
    Application.OnKey "^+1", "'" & ThisWorkbook.Name & "'!CycleNumberFormat"              ' Ctrl+Shift+1
    Application.OnKey "^+3", "'" & ThisWorkbook.Name & "'!CycleDateFormat"                ' Ctrl+Shift+3
    Application.OnKey "^+4", "'" & ThisWorkbook.Name & "'!CycleCurrencyFormat"            ' Ctrl+Shift+4
    Application.OnKey "^+5", "'" & ThisWorkbook.Name & "'!CyclePercentFormat"             ' Ctrl+Shift+5
    Application.OnKey "^+8", "'" & ThisWorkbook.Name & "'!CycleOtherNumbers"              ' Ctrl+Shift+8

    ' Decimal places
    Application.OnKey "^+.", "'" & ThisWorkbook.Name & "'!IncreaseDecimal"                ' Ctrl+Shift+.
    Application.OnKey "^+,", "'" & ThisWorkbook.Name & "'!DecreaseDecimal"                ' Ctrl+Shift+,

    ' Scale (unit conversion)
    Application.OnKey "+%<", "'" & ThisWorkbook.Name & "'!ScaleUp"                        ' Alt+Shift+<
    Application.OnKey "+%>", "'" & ThisWorkbook.Name & "'!ScaleDown"                      ' Alt+Shift+>

    '==== modStyles (Sxx) – colors/styles/layout/CF ===========================
    ' AutoColor
    Application.OnKey "^%a", "'" & ThisWorkbook.Name & "'!AutoColorSelection"             ' Ctrl+Alt+A

    ' Font / fill / text case / font color
    Application.OnKey "^'", "'" & ThisWorkbook.Name & "'!CycleFont"                       ' Ctrl+'
    Application.OnKey "^+K", "'" & ThisWorkbook.Name & "'!CycleFill"                      ' Ctrl+Shift+K
    Application.OnKey "^%+I", "'" & ThisWorkbook.Name & "'!CycleTextCase"                 ' Ctrl+Alt+Shift+I
    Application.OnKey "^+C", "'" & ThisWorkbook.Name & "'!CycleFontColor"                 ' Ctrl+Shift+C

    ' Zoom (NOTE: your binding strings include Shift)
    Application.OnKey "^%+=", "'" & ThisWorkbook.Name & "'!ZoomIn"                        ' Ctrl+Alt+Shift+=
    Application.OnKey "^%+-", "'" & ThisWorkbook.Name & "'!ZoomOut"                       ' Ctrl+Alt+Shift+-

    ' Font size
    Application.OnKey "^+F", "'" & ThisWorkbook.Name & "'!IncreaseFontSize"               ' Ctrl+Shift+F
    Application.OnKey "^+G", "'" & ThisWorkbook.Name & "'!DecreaseFontSize"               ' Ctrl+Shift+G

    ' Indent
    Application.OnKey "^+]", "'" & ThisWorkbook.Name & "'!IndentIn"                       ' Ctrl+Shift+]
    Application.OnKey "^+[", "'" & ThisWorkbook.Name & "'!IndentOut"                      ' Ctrl+Shift+[

    ' Layout helpers
    Application.OnKey "^%e", "'" & ThisWorkbook.Name & "'!CenterAcrossSelection"          ' Ctrl+Alt+E

    ' Misc
    Application.OnKey "^+N", "'" & ThisWorkbook.Name & "'!InsertStaticNow"                ' Ctrl+Shift+N
    Application.OnKey "^%+V", "'" & ThisWorkbook.Name & "'!PasteValuesKeepFormat"         ' Ctrl+Alt+Shift+V

    ' Input / Header tools
    Application.OnKey "^%+U", "'" & ThisWorkbook.Name & "'!CycleInputStyle"               ' Ctrl+Alt+Shift+U
    'Application.OnKey "^%+1", "'" & ThisWorkbook.Name & "'!ApplyInputYellow"             ' Ctrl+Alt+Shift+1 (optional)
    'Application.OnKey "^%+2", "'" & ThisWorkbook.Name & "'!ApplyInputGray"               ' Ctrl+Alt+Shift+2 (optional)
    Application.OnKey "^%+H", "'" & ThisWorkbook.Name & "'!CycleHeaderStyle"              ' Ctrl+Alt+Shift+H
    Application.OnKey "^%+Y", "'" & ThisWorkbook.Name & "'!InsertHeadersFromPrompt"       ' Ctrl+Alt+Shift+Y
    Application.OnKey "^%+D", "'" & ThisWorkbook.Name & "'!InsertVarianceHeaders"         ' Ctrl+Alt+Shift+D

    ' Zero-check CF
    Application.OnKey "^%+Z", "'" & ThisWorkbook.Name & "'!ApplyZeroCheckCF"              ' Ctrl+Alt+Shift+Z
    Application.OnKey "^%+X", "'" & ThisWorkbook.Name & "'!ClearZeroCheckCF"              ' Ctrl+Alt+Shift+X

    '==== modBorders (Bxx) ====================================================
    Application.OnKey "^%+{UP}", "'" & ThisWorkbook.Name & "'!BorderTop"                  ' Ctrl+Alt+Shift+Up
    Application.OnKey "^%+{DOWN}", "'" & ThisWorkbook.Name & "'!BorderBottom"             ' Ctrl+Alt+Shift+Down
    Application.OnKey "^%+{LEFT}", "'" & ThisWorkbook.Name & "'!BorderLeft"               ' Ctrl+Alt+Shift+Left
    Application.OnKey "^%+{RIGHT}", "'" & ThisWorkbook.Name & "'!BorderRight"             ' Ctrl+Alt+Shift+Right
    Application.OnKey "^%+B", "'" & ThisWorkbook.Name & "'!BordersOutlineInside"          ' Ctrl+Alt+Shift+B

    '==== modUnitTags (Uxx) ===================================================
    Application.OnKey "^%+T", "'" & ThisWorkbook.Name & "'!CycleUnitTag_Value_Uniform"       ' Ctrl+Alt+Shift+T
    Application.OnKey "^%+O", "'" & ThisWorkbook.Name & "'!CycleUnitTag_Duration_Uniform"    ' Ctrl+Alt+Shift+O
    Application.OnKey "^%+P", "'" & ThisWorkbook.Name & "'!CycleUnitTag_Rate_Uniform"        ' Ctrl+Alt+Shift+P
    Application.OnKey "^%+{BACKSPACE}", "'" & ThisWorkbook.Name & "'!RemoveUnitTag"          ' Ctrl+Alt+Shift+Backspace

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Performance Mode quick note (for humans, not Excel)
'   Turn ON before heavy actions: mass formatting, filling/copying big ranges,
'   applying/removing lots of CF, duplicating sheets.
'   It sets Manual calc, turns off screen updates & events, and skips undo/logging.
'
'   Turn OFF for review and normal use: automatic calc, Ctrl+Z history, and logging.
'------------------------------------------------------------------------------

