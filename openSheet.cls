Private Sub Workbook_Open()
 If (ActiveSheet.Name <> "sheet") Then
    ThisWorkbook.Worksheets.Item("sheet").Activate
  End If
End Sub
