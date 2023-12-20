' From multisheets excel file get one subExcel file for each sheet '
Sub SplitEachSheet() 
    Dim FilePath As String 

    'Get Current Dir of your multiSheet excel File'
    FilePath = Application.ActiveWorkbook.FilePath

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False 

    'For each sheet i have to record sheet data to new excel file'
    For Each sheet IN ThisWorkbook.Sheets
        sheet.Copy 
        Application.ActiveWorkbook.SaveAs Filename:=FilePath & "\" & sheet.Name & ".xlsx"
        Application.ActiveWorkbook.Close False
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
