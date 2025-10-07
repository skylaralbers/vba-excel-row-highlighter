# Excel Row Highlighter Macro

This VBA script highlights the last clicked row in Excel with a light blue fill color and keeps it highlighted even after clicking away or switching windows.

Excel doesn’t have this built in—even though it should.  
This is extremely useful when you’re entering data across multiple windows or referencing Excel while using other programs.  
It keeps your position visually locked so you never lose track of the active row.

Color can be changed by editing VBA macro's code in RGB values or specific name. 

---

## 💻 Macro Code (fully explained inline)

```vba
' === Excel Row Highlighter Macro ===
' Highlights the last clicked row with a persistent color (default: light blue)
' Useful for data entry across multiple windows; Excel lacks this natively.

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next

    ' Store the number of the currently selected row
    ' "LastRow" becomes a named reference that tracks which row to highlight
    If Me.Names("LastRow") Is Nothing Then
        Me.Names.Add Name:="LastRow", RefersTo:="=" & Target.Row
    Else
        Me.Names("LastRow").RefersTo = "=" & Target.Row
    End If

    ' Define which part of the sheet will highlight
    ' Adjust this range (A:Z) to fit your actual data width
    Dim rng As Range
    Set rng = Me.Range("A:Z")

    ' Prevent duplicate conditional formatting rules
    Dim fc As FormatCondition, hasRule As Boolean
    hasRule = False
    For Each fc In rng.FormatConditions
        If fc.Type = xlExpression Then
            If fc.Formula1 = "=ROW()=LastRow" Then hasRule = True
        End If
    Next fc

    ' If the format rule doesn’t exist, add it
    If Not hasRule Then
        rng.FormatConditions.Add Type:=xlExpression, Formula1:="=ROW()=LastRow"
        With rng.FormatConditions(rng.FormatConditions.Count).Interior
            .Pattern = xlSolid
            .Color = RGB(173, 216, 230) ' Default light blue color
        End With
    End If
End Sub
