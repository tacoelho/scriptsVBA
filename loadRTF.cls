'******************************************
'Last update 22.08.2018
'Por tcoelho83@gmail.com
'******************************************

Option Explicit
Option Private Module

Private Sub loadRTF()

Dim wdApp As Word.Application
Dim wdDoc As Word.Document
Dim rngDestination As Excel.Range
Dim sPath As String
Dim sFile As String


sPath = "YOUR .RTF FILE PATH HERE"
sFile = Dir(sFile & "*.RTF")

Set wdApp = CreateObject("Word.Application")
With wdApp
    .Visible = True
    .ScreenUpdating = False
    .DisplayAlerts = wdAlertsNone
End With

Set wdDoc = wdApp.Documents.Open(Filename:=sPath & sFile)
wdDoc.Content.Copy


Set rngDestination = Application.Range("YOUR DEFINED RANGE")
With rngDestination
    .PasteSpecial Paste:=xlPasteValues
End With

ActiveWorkbook.Sheets("YOUR DEFINED SHEET").Activate

wdDoc.Close (True)
Set wdDoc = Nothing
wdApp.Quit (True)
Set wdApp = Nothing

End Sub