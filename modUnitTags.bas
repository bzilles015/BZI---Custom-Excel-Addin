Attribute VB_Name = "modUnitTags"
'==============================================================================
' Module: modUnitTags (Uxx)
' Purpose:
'   - Apply uniform bracket unit tags (e.g., [#], [%], [mln $]) to selections.
'   - Detect/replace/remove the last [...] tag in cell text.
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' U01 – CycleUnitTag_Value_Uniform  (Ctrl+Alt+Shift+T)
'     Cycles selection through: [#], [%], [mln $], [thd $], [bn $], [x], [pp], [bps].
'------------------------------------------------------------------------------

Public Sub CycleUnitTag_Value_Uniform()
    ApplyUniformTagCycle Array("[#]", "[%]", "[mln $]", "[thd $]", "[bn $]", "[x]", "[pp]", "[bps]"), _
                         "UnitTag_Value_Uniform"
End Sub


'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' U02 – CycleUnitTag_Duration_Uniform  (Ctrl+Alt+Shift+O)
'     Cycles duration-style tags uniformly across the whole selection.
'------------------------------------------------------------------------------

Public Sub CycleUnitTag_Duration_Uniform()
    ApplyUniformTagCycle Array("[d]", "[m]", "[q]", "[y]"), _
                         "UnitTag_Duration_Uniform"
End Sub


'------------------------------------------------------------------------------
' U03 – CycleUnitTag_Rate_Uniform  (Ctrl+Alt+Shift+P)
'     Cycles rate-style tags uniformly across the whole selection.
'------------------------------------------------------------------------------


Public Sub CycleUnitTag_Rate_Uniform()
    ApplyUniformTagCycle Array("[%/y]", "[$/unit]", "[$/FTE]", "[$/yr]"), _
                         "UnitTag_Rate_Uniform"
End Sub


'------------------------------------------------------------------------------
' U04 – RemoveUnitTag  - (Ctrl+Alt+Shift+Backspace)
'     ' Remove the final [...] tag from each selected cell
'------------------------------------------------------------------------------

Public Sub RemoveUnitTag()
    If TypeName(Selection) <> "Range" Then Exit Sub
    BeginMacroWithUndo
    Dim c As Range
    For Each c In Selection.Cells
        If Not c.HasFormula Then c.Value = StripLastBracketTag(CStr(c.Value))
    Next c
    LogAction "UnitTag_Remove", Selection.Address(False, False)
    RegisterUndo "Remove Unit Tag"
End Sub


'================ CORE ================
Private Sub ApplyUniformTagCycle(ByVal tags As Variant, ByVal actionName As String)
    If TypeName(Selection) <> "Range" Then Exit Sub

    BeginMacroWithUndo

    ' 1) Determine the current tag from first nonblank, non-formula cell in selection
    Dim cur As String, idx As Long
    cur = DetectSelectionTag(Selection)
    idx = FindTagIndex(cur, tags)

    ' 2) Compute next tag for ENTIRE selection
    Dim n As Long, nextTag As String
    n = UBound(tags) - LBound(tags) + 1
    If idx = -1 Then
        nextTag = CStr(tags(LBound(tags)))
    Else
        nextTag = CStr(tags((idx - LBound(tags) + 1) Mod n + LBound(tags)))
    End If

    ' 3) Apply nextTag to all non-formula cells
    Dim c As Range
    For Each c In Selection.Cells
        If Not c.HasFormula Then
            c.Value = ReplaceOrAppendTag(CStr(c.Value), nextTag)
        End If
    Next c

    LogAction actionName & ":" & nextTag, Selection.Address(False, False)
    RegisterUndo "Cycle Unit Tag (Uniform)"
End Sub

' Return index of tag in list; -1 if not found
Private Function FindTagIndex(ByVal tag As String, ByVal tags As Variant) As Long
    Dim i As Long
    If Len(tag) = 0 Then FindTagIndex = -1: Exit Function
    For i = LBound(tags) To UBound(tags)
        If StrComp(tag, CStr(tags(i)), vbTextCompare) = 0 Then
            FindTagIndex = i
            Exit Function
        End If
    Next i
    FindTagIndex = -1
End Function

' Detect tag from first nonblank, non-formula cell; "" if none
Private Function DetectSelectionTag(ByVal rg As Range) As String
    Dim c As Range, s As String
    For Each c In rg.Cells
        If Len(c.Value2) > 0 And Not c.HasFormula Then
            s = CStr(c.Value2)
            DetectSelectionTag = ExtractLastBracketTag(s)
            Exit Function
        End If
    Next c
    DetectSelectionTag = ""
End Function

' Replace last [...] or append new tag
Private Function ReplaceOrAppendTag(ByVal s As String, ByVal newTag As String) As String
    Dim lb As Long, rb As Long
    lb = InStrRev(s, "["): rb = InStrRev(s, "]")
    If lb > 0 And rb > lb Then
        ReplaceOrAppendTag = Trim$(Left$(s, lb - 1) & newTag & Mid$(s, rb + 1))
    Else
        If Len(Trim$(s)) = 0 Then
            ReplaceOrAppendTag = newTag
        Else
            ReplaceOrAppendTag = Trim$(s) & " " & newTag
        End If
    End If
End Function

' Extract last [...] tag; "" if none
Private Function ExtractLastBracketTag(ByVal s As String) As String
    Dim lb As Long, rb As Long
    lb = InStrRev(s, "["): rb = InStrRev(s, "]")
    If lb > 0 And rb > lb Then
        ExtractLastBracketTag = Mid$(s, lb, rb - lb + 1)
    Else
        ExtractLastBracketTag = ""
    End If
End Function

' Remove last [...] tag
Private Function StripLastBracketTag(ByVal s As String) As String
    Dim lb As Long, rb As Long
    lb = InStrRev(s, "["): rb = InStrRev(s, "]")
    If lb > 0 And rb > lb Then
        StripLastBracketTag = Trim$(Left$(s, lb - 1) & Mid$(s, rb + 1))
    Else
        StripLastBracketTag = s
    End If
End Function


