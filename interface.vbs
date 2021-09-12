Sub compare_csvs()
'
' compare_csvs Macro
' Compares multiple CSV files visually.
'
' Keyboard Shortcut: Ctrl+Shift+C
'
    Dim d_files As FileDialog       ' Dialog to ask for CSVs
    
    ' Ask for multiple CSV files, starting in this directory
    Set d_files = Application.FileDialog(msoFileDialogOpen)
    d_files.AllowMultiSelect = True
    d_files.Filters.Add "Comma Separated Value files", "*.csv", 1
    d_files.InitialFileName = ThisWorkbook.Path
    d_files.Show
    
    ' Loop through and print the file names
    For k = 1 To d_files.SelectedItems.Count
        MsgBox d_files.SelectedItems.Item(k)
    Next k
End Sub 'compare_csvs()
