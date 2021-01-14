
Private Sub Workbook_Open()
'Instantiate singletons every time the workbook is opened. 
'This allows VLookup only to be executed one time when the
'workbook is opened instead of one time per row.
  Call InstantiateSingletons
End Sub
