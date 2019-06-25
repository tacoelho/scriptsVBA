Private Sub SaveAsPdf()

Dim PDFFilename As Variant
PDFFilename = Application.GetSaveAsFilename(InitialFileName:="yourfilename", _
FileFilter:="PDF, *.pdf", _
Title:="Salvar como PDF")
ChDir "YOUR DIRECTORY"
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
      PDFFilename, Quality:=xlQualityStandard, _
      IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
      True
End Sub