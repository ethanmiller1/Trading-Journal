Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  If Target.Column <> 18 Or Target.Row < 9 Then Exit Sub
  Dim MyCell As String
  MyCell = Range("A" & Target.Row).text
  Clipboard (GetOptionSignature(MyCell))
End Sub