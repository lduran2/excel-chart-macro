Sub compare_csvs()
'
' compare_csvs Macro
' Compares multiple CSV files visually.
'
' Keyboard Shortcut: Ctrl+Shift+R
'
' by      : Leomar Duran <https://github.com/lduran2/>
' when    : 2021-09-12 t20:45
' self    : https://github.com/lduran2/excel-chart-macro
' version : 1.3
'
' changelog :
'     v1.3 -- 2021-09-12 t20:45
'         successfully copying CSV files into the worksheets
'
'     v1.2.2 -- 2021-09-12 t18:36
'         the save file input section is its own function, as is copy
'             to array, and change error checking for saving
'
'     v1.2.1 -- 2021-09-12 t18:01
'         the input CSV file input and copy to array section are their own function
'
'     v1.2 -- 2021-09-12 t17:49
'         creates all of the worksheets in the output workbook
'
'     v1.1 -- 2021-09-12 t17:00
'         asks User Agent for the save path, creates and saves a file
'
'     v1.0 -- 2021-09-12 t16:38
'         asks User Agent for input CSV files and echos them to the user
'
    Dim in_filenames() As String            ' Input CSV filenames
    Dim out_filenames() As String           ' Output files to save to (should be <= 1)
    Dim out_workbook As Workbook            ' Points to the current sheet in the output workbook
    Dim this_path As String                 ' The directory that this workbook runs from
    Dim curr_in_file_subdirs() As String    ' The subdirectories of the current input file
    Dim curr_sheet_name As String           ' The name of the corresponding worksheet
    Dim curr_workbook_nsame As String       ' The name of the current workbook
    
    ' Build the path
    this_path = (ThisWorkbook.path & "\")
    
    ' Ask for multiple CSV files and save path, starting in this directory
    in_filenames = input_csv_files(this_path)
    out_filenames = input_save_fle(this_path)

    ' Create and store the output file workbook
    Set out_workbook = Workbooks.Add
    ' Add a sheet for the charts
    Sheets.Item(1).Name = "Chart Summary"
    ' Add each CSV to the workbook
    For k = 1 To UBound(in_filenames)
        ' Split the current file's path
        curr_in_file_subdirs = Split(in_filenames(k), "\")
        ' Create the name of the workbook and the corresponding sheet "(k) file_name.csv"
        curr_csv_workbook_name = curr_in_file_subdirs(UBound(curr_in_file_subdirs))
        curr_sheet_name = "(" & k & ") " & curr_csv_workbook_name

        ' Add the sheet to the output workbook
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = curr_sheet_name

        ' Open the corresponding CSV file
        Workbooks.Open(in_filenames(k)).Activate

        ' Copy the cells of the CSV workbook into the output worksheet
        Sheets(1).Cells.Copy _
            out_workbook.Sheets(out_workbook.Sheets.Count).Cells
        '

        ' Close the CSV file
        Workbooks(curr_csv_workbook_name).Close
        ' Set the output workbook to active
        out_workbook.Activate
    Next k

    ' Loop through the "save filenames" (there should be <= 1)
    For k = 1 To UBound(out_filenames)
        ' Save the output file
        ActiveWorkbook.SaveAs out_filenames(k)
    Next

End Sub 'compare_csvs()
' --------------------------------------------------------------------

Function input_csv_files(path As String) As Variant
'
' input_csv_files Function
' Asks the user for CSV data files for input, returned as an array.
'
    Dim d_in_files As FileDialog            ' Dialog to ask for input CSVs
    
    ' Ask for multiple CSV files, starting in this directory
    MsgBox "Please select the CSV data files for input."
    Set d_in_files = Application.FileDialog(msoFileDialogOpen)          ' Save file dialog
    d_in_files.AllowMultiSelect = True                                  ' multiple files selectable
    d_in_files.Filters.Add "Comma Separated Value files", "*.csv", 1    ' filter out CSV files
    d_in_files.InitialFileName = path                                   ' start in this directory
    d_in_files.Show                                                     ' show the dialog after building it
    
    ' Return the filenames
    input_csv_files = copy_selected_files(d_in_files.SelectedItems)
End Function 'input_csv_files(path As String) As Variant
' --------------------------------------------------------------------

Function input_save_fle(path As String) As Variant
'
' input_save_fle Function
' Asks the user for 1 path to save to, returned as an array.
'
    Dim d_out_file As FileDialog            ' Dialog to ask for location to save
    
    ' Ask for path to save to
    MsgBox "Please choose a path to save the charts."
    Set d_out_file = Application.FileDialog(msoFileDialogSaveAs)        ' Open file dialog
    d_out_file.InitialFileName = path                                   ' start in this directory
    d_out_file.Show                                                     ' show the dialog after building it

    input_save_fle = copy_selected_files(d_out_file.SelectedItems)
End Function 'input_save_fle(path As String) As Variant
' --------------------------------------------------------------------

Function copy_selected_files(items As FileDialogSelectedItems) As Variant
'
' copy_selected_files Function
' Copies a FileDialogSelectedItems object into a standard Visual Basic
' array.
'
    Dim filenames() As String               ' The selected files to copy
    
    ' Copy the filenames selected
    ReDim filenames(items.Count)
    For k = 1 To items.Count
        filenames(k) = items.Item(k)
    Next k
    
    ' Return the copy
    copy_selected_files = filenames
End Function 'copy_selected_files(items As FileDialogSelectedItems) As Variant

