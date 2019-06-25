Option Explicit
Option Private Module

Private Sub createPath()

    Dim strPath         As String: strPath = "C:\YOUR PATH HERE"
    Dim strYear         As String: strYear = Year(Date) & "\"
    Dim strMonth        As String: strMonth = MonthName(Month(Date)) & "\"
    Dim strDay          As String: strDay = Day(Date) & "\"
    Dim strFile		As String: strFile = "YOUR FILE NAME"

With ActiveWorkbook
 
        ' Checks for year path
        If Len(Dir(strPath & strYear, vbDirectory)) = 0 Then
            MkDir strPath & strYear
        End If
        ' Checks for moth path
        If Len(Dir(strPath & strYear & strMonth, vbDirectory)) = 0 Then
            MkDir strPath & strYear & strMonth
        End If
        ' Checks for day path
        If Len(Dir(strPath & strYear & strMonth & strDay, vbDirectory)) = 0 Then
            MkDir strPath & strYear & strMonth & strDay
        End If
	Workbooks.Add (strFile)

            ActiveWorkbook.SaveAs Filename:= _
            strPath & strYear & strMonth & strDay & strFile, _
            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            
            ActiveWorkbook.Close

End With

Application.ScreenUpdating = True

MsgBox ("Workbook saved succesfully.")



End Sub