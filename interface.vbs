Sub compare_csvs()
'
' compare_csvs Macro
' Compares multiple CSV files visually.
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Dim d_in_files As FileDialog        ' Dialog to ask for input CSVs
    Dim d_out_file As FileDialog       ' Dialog to ask for location to save
    Dim this_path As String             ' The directory that this workbook runs from

    ' Build the path
    this_path = (ThisWorkbook.Path & "\")

    ' Ask for multiple CSV files, starting in this directory
    MsgBox "Please select the CSV files for input."
    Set d_in_files = Application.FileDialog(msoFileDialogOpen)          ' Save file dialog
    d_in_files.AllowMultiSelect = True                                  ' multiple files selectable
    d_in_files.Filters.Add "Comma Separated Value files", "*.csv", 1    ' filter out CSV files
    d_in_files.InitialFileName = this_path                              ' start in this directory
    d_in_files.Show                                                     ' show the dialog after building it

    ' Loop through and print the file names
    For k = 1 To d_in_files.SelectedItems.Count
        MsgBox d_in_files.SelectedItems.Item(k)
    Next k

    ' Ask for path to save to
    MsgBox "Please choose a path to save the charts."
    Set d_out_file = Application.FileDialog(msoFileDialogSaveAs)        ' Open file dialog
    d_out_file.InitialFileName = this_path                              ' start in this directory
    d_out_file.Show                                                     ' show the dialog after building it

    ' Create and save the output file
    Workbooks.Add
    ActiveWorkbook.SaveAs d_out_file.SelectedItems.Item(1)

End Sub 'compare_csvs()

