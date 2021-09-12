Sub compare_csvs()
'
' compare_csvs Macro
' Compares multiple CSV files visually.
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Dim in_filenames() As String            ' Input CSV filenames
    Dim out_filenames() As String           ' Output files to save to (should be <= 1)
    Dim an_out_filename As Variant
    Dim this_path As String                 ' The directory that this workbook runs from
    Dim curr_in_file_subdirs() As String    ' The subdirectories of the current input file
    Dim curr_sheet_name As String           ' The name of the corresponding worksheet

    ' Build the path
    this_path = (ThisWorkbook.path & "\")
    
    ' Ask for multiple CSV files and save path, starting in this directory
    in_filenames = InputCsvFiles(this_path)
    out_filenames = InputSaveFile(this_path)

    ' Create the output file
    Workbooks.Add
    ' Add a sheet for the charts
    Sheets.Item(1).Name = "Chart Summary"
    ' Add each CSV to the workbook
    For k = 1 To UBound(in_filenames)
        ' Split the current file's path
        curr_in_file_subdirs = Split(in_filenames(k), "\")
        ' Create the name of the sheet "(k) file_name.csv"
        curr_sheet_name = "(" & k & ") " & curr_in_file_subdirs(UBound(curr_in_file_subdirs))
        ' Add the sheet
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = curr_sheet_name
    Next k

    ' Loop through the "save filenames" (there should be <= 1)
    For k = 1 To UBound(out_filenames)
        ' Save the output file
        ActiveWorkbook.SaveAs out_filenames(k)
    Next

End Sub 'compare_csvs()
' --------------------------------------------------------------------

Function InputCsvFiles(path As String) As Variant
    Dim d_in_files As FileDialog            ' Dialog to ask for input CSVs
    
    ' Ask for multiple CSV files, starting in this directory
    MsgBox "Please select the CSV files for input."
    Set d_in_files = Application.FileDialog(msoFileDialogOpen)          ' Save file dialog
    d_in_files.AllowMultiSelect = True                                  ' multiple files selectable
    d_in_files.Filters.Add "Comma Separated Value files", "*.csv", 1    ' filter out CSV files
    d_in_files.InitialFileName = path                                   ' start in this directory
    d_in_files.Show                                                     ' show the dialog after building it
    
    ' Return the filenames
    InputCsvFiles = CopySelectedFiles(d_in_files.SelectedItems)
End Function 'InputCsvFiles(path As String) As Variant
' --------------------------------------------------------------------

Function InputSaveFile(path As String) As Variant
    Dim d_out_file As FileDialog            ' Dialog to ask for location to save
    
    ' Ask for path to save to
    MsgBox "Please choose a path to save the charts."
    Set d_out_file = Application.FileDialog(msoFileDialogSaveAs)        ' Open file dialog
    d_out_file.InitialFileName = path                                   ' start in this directory
    d_out_file.Show                                                     ' show the dialog after building it

    InputSaveFile = CopySelectedFiles(d_out_file.SelectedItems)
End Function 'InputSaveFile(path As String) As Variant
' --------------------------------------------------------------------

Function CopySelectedFiles(items As FileDialogSelectedItems) As Variant
    Dim filenames() As String               ' The selected files to copy
    
    ' Copy the filenames selected
    ReDim filenames(items.Count)
    For k = 1 To items.Count
        filenames(k) = items.Item(k)
    Next k
    
    ' Return the copy
    CopySelectedFiles = filenames
End Function 'CopySelectedFiles(items As FileDialogSelectedItems) As Variant
' --------------------------------------------------------------------
