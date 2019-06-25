Option Explicit
Option Private Module

Private Sub LimpaConteudo()

Application.ScreenUpdating = False                        
    
    Dim sheets      As Variant
    Dim sheet         As Variant
    
    sheets = (Array("sheet1", "sheet2", "sheet3"))     
    
    For Each sheet In sheets                               
        Sheets(sheet).Rows("PUT THE NUMBER OF THE STARTING ROW HERE:" & Rows.Count).ClearContents
    Next sheet

Application.ScreenUpdating = True

    MsgBox ("All sheets are cleared!")

End Sub