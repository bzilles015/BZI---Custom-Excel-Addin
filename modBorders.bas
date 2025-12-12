Attribute VB_Name = "modBorders"
 Option Explicit


'------------------------------------------------------------------------------
' B01 – BorderTop  (Ctrl+Alt+Shift+Up)
'     Cycles top edge border for the selection.
'------------------------------------------------------------------------------

Sub BorderTop()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevTopAddress Then BorderTopIndex = 0
    PrevTopAddress = Selection.Address(False, False)
    Dim idx As Integer: idx = BorderTopIndex Mod 4
    With Selection.Borders(xlEdgeTop)
        Select Case idx
          Case 0: .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
          Case 1: .LineStyle = xlNone
          Case 2: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlMedium
          Case 3: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlHairline
        End Select
    End With
    BorderTopIndex = BorderTopIndex + 1
    LogAction "BdrTopCyc" & (idx + 1), PrevTopAddress
    RegisterUndo "Top Border"
End Sub

'------------------------------------------------------------------------------
' B02 – BorderBottom  (Ctrl+Alt+Shift+Down)
'     Cycles bottom edge border for the selection.
'------------------------------------------------------------------------------

Sub BorderBottom()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevBottomAddress Then BorderBottomIndex = 0
    PrevBottomAddress = Selection.Address(False, False)
    Dim idx As Integer: idx = BorderBottomIndex Mod 4
    With Selection.Borders(xlEdgeBottom)
        Select Case idx
          Case 0: .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
          Case 1: .LineStyle = xlNone
          Case 2: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlMedium
          Case 3: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlHairline
        End Select
    End With
    BorderBottomIndex = BorderBottomIndex + 1
    LogAction "BdrBotCyc" & (idx + 1), PrevBottomAddress
    RegisterUndo "Bottom Border"
End Sub

'------------------------------------------------------------------------------
' B03 – BorderLeft  (Ctrl+Alt+Shift+Left)
'     Cycles left edge border for the selection.
'------------------------------------------------------------------------------

Sub BorderLeft()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevLeftAddress Then BorderLeftIndex = 0
    PrevLeftAddress = Selection.Address(False, False)
    Dim idx As Integer: idx = BorderLeftIndex Mod 4
    With Selection.Borders(xlEdgeLeft)
        Select Case idx
          Case 0: .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
          Case 1: .LineStyle = xlNone
          Case 2: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlMedium
          Case 3: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlHairline
        End Select
    End With
    BorderLeftIndex = BorderLeftIndex + 1
    LogAction "BdrLftCyc" & (idx + 1), PrevLeftAddress
    RegisterUndo "Left Border"
End Sub

'------------------------------------------------------------------------------
' B04 – BorderRight  (Ctrl+Alt+Shift+Right)
'     Cycles right edge border for the selection.
'------------------------------------------------------------------------------

Sub BorderRight()
    BeginMacroWithUndo
    If Selection.Address(False, False) <> PrevRightAddress Then BorderRightIndex = 0
    PrevRightAddress = Selection.Address(False, False)
    Dim idx As Integer: idx = BorderRightIndex Mod 4
    With Selection.Borders(xlEdgeRight)
        Select Case idx
          Case 0: .LineStyle = xlContinuous: .ColorIndex = 0: .TintAndShade = 0: .Weight = xlThin
          Case 1: .LineStyle = xlNone
          Case 2: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlMedium
          Case 3: .LineStyle = xlContinuous: .ColorIndex = xlAutomatic: .TintAndShade = 0: .Weight = xlHairline
        End Select
    End With
    BorderRightIndex = BorderRightIndex + 1
    LogAction "BdrRgtCyc" & (idx + 1), PrevRightAddress
    RegisterUndo "Right Border"
End Sub

'------------------------------------------------------------------------------
' B05 – BordersOutlineInside  (Ctrl+Alt+Shift+B)
'     Applies outline + inside borders using the current border cycle state.
'------------------------------------------------------------------------------

Sub BordersOutlineInside()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    With Selection
        .Borders.LineStyle = xlNone
        With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: .ColorIndex = 0: End With
        With .Borders(xlInsideVertical):   .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 0: End With
        With .Borders(xlInsideHorizontal): .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 0: End With
    End With
    LogAction "BdrOutlineInside", Selection.Address(False, False)
    RegisterUndo "Outline Borders"
End Sub

