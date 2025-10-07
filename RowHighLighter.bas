Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next

    If Me.Names("LastRow") Is Nothing Then
        Me.Names.Add Name:="LastRow", RefersTo:="=" & Target.Row
    Else
        Me.Names("LastRow").RefersTo = "=" & Target.Row
    End If

    Dim rng As Range
    Set rng = Me.Range("A:Z")  ' adjust to your data

    Dim fc As FormatCondition, hasRule As Boolean
    hasRule = False
    For Each fc In rng.FormatConditions
        If fc.Type = xlExpression Then
            If fc.Formula1 = "=ROW()=LastRow" Then hasRule = True
        End If
    Next fc
    If Not hasRule Then
        rng.FormatConditions.Add Type:=xlExpression, Formula1:="=ROW()=LastRow"
        With rng.FormatConditions(rng.FormatConditions.Count).Interior
            .Pattern = xlSolid
            .Color = RGB(173, 216, 230) ' light blue
        End With
    End If
End Sub
