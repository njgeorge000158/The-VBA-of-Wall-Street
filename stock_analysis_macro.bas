Attribute VB_Name = "Module1"
'*******************************************************************************************
 '
 '  File Name:  stock_analysis_macro.bas
 '
 '  File Description:
 '      The file contains the macro, stock_analysis_macro, which formats the active
 '      worksheet then generates summary tables from raw stock data. Here is a list
 '      of the support functions:
 '
 '      format_stock_data
 '      format_summary_data
 '      format_titles
 '      format_worksheet
 '
 '      create_summary_table
 '      create_change_table
 '      convert_data_and_summary_ranges_to_tables
 '      create_format_analysis_worksheet
 '
 '      change_string_to_date_in_date_column
 '      set_up_titles_for_summary_data
 '      create_summary_data_row
 '      format_change_data_titles
 '      set_up_change_data_titles
 '      calculate_and_write_change_data
 '      convert_range_into_table
 '      format_yearly_change_cell
 '      create_worksheet
 '      format_analysis_worksheet
 '      format_local_summary_data
 '      insert_analysis_worksheet_row_and_titles
 '
 '      set_up_sorted_tables
 '      set_up_greatest_increase_table
 '      set_up_greatest_decrease_table
 '      set_up_greatest_total_volume_table
 '
 '      copy_table
 '      sort_table
 '
 '      calculate_yearly_change_function
 '      calculate_percent_change_function
 '      return_analysis_worksheet_name_function
 '
 '
 '  Date               Description                             Programmer
 '  ---------------    ------------------------------------    ------------------
 '  07/19/2023         Initial Development                     Nicholas J. George
 '
'*******************************************************************************************/

' These are the global enumerations that identify the rows and columns in the original
' and summary data.
Enum row_global_enumeration
    
    ENUM_K_TITLE = 1
    
    ENUM_K_FIRST_DATA = 2
    
    ENUM_K_PERCENT_DECREASE = 3
    
    ENUM_K_GREATEST_TOTAL_VOLUME = 4

End Enum

Enum column_global_enumeration
    
    ENUM_K_STOCK_TICKER = 1
    
    ENUM_K_STOCK_DATE = 2
    
    ENUM_K_STOCK_OPEN = 3
    
    ENUM_K_STOCK_HIGH = 4
    
    ENUM_K_STOCK_LOW = 5
    
    ENUM_K_STOCK_CLOSE = 6
    
    ENUM_K_STOCK_VOL = 7
    
    ENUM_K_BLANK_1 = 8

    ENUM_K_SUMMARY_TICKER = 9
    
    ENUM_K_SUMMARY_YEARLY_CHANGE = 10
    
    ENUM_K_SUMMARY_PERCENT_CHANGE = 11
    
    ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME = 12
    
    ENUM_K_BLANK_2 = 13
    
    ENUM_K_BLANK_3 = 14
    
    ENUM_K_CHANGE_ROW_TITLES = 15
    
    ENUM_K_CHANGE_TICKERS = 16
    
    ENUM_K_CHANGE_VALUES = 17

End Enum


' These global constants specify substring lengths in the original data's date strings:
' the date string format is YYYYMMDD.
Global Const GLOBAL_CONSTANT_YEAR_LENGTH As Integer = 4

Global Const GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH As Integer = 2


' This global variable holds the value of the number of rows in the raw stock data.
Global last_data_row_global_long As Long

'*******************************************************************************************
 '
 '  Macro Name:  stock_analysis_macro
 '
 '  Macro Description:
 '      This macro formats the active worksheet then generates summary tables
 '      from raw stock data.
 '
 '  Macro Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Sub stock_analysis_macro()

    ' This line of code assigns the last row index to the appropriate global variable.
    last_data_row_global_long _
        = Worksheets _
                (ActiveSheet.Name) _
                    .Cells _
                        (Rows.Count, _
                         column_global_enumeration.ENUM_K_STOCK_TICKER) _
                    .End(xlUp) _
                    .Row

    If InStr(1, ActiveSheet.Name, "Analysis", vbTextCompare) = 0 _
        And InStr(1, ActiveSheet.Name, "analysis", vbTextCompare) = 0 Then
    
        ' These subroutines format the active worksheet.
        format_worksheet (ActiveWorkbook.ActiveSheet.Name)

        format_stock_data

        format_summary_data
    
        format_titles
    
        
        ' This subroutine summarizes the raw stock data and writes it to the summary table.
        create_summary_table
    
    
        ' This subroutine creates a second summary table for the tickers with the greatest
        ' changes in percentage and greatest total stock volume.
        create_change_table
    
    
        ' This subroutine converts the data and summary ranges to tables.
        convert_data_and_summary_ranges_to_tables
    
    
        ' This subroutine creates, formats, and populates the Analysis Worksheet.
        create_format_analysis_worksheet

    Else
    
        MsgBox ("Please select the Summary Data Worksheet not the Analysis Worksheet!")
    
    End If

End Sub ' This statement ends the macro, stock_analysis_macro.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_stock_data
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the formats of the stock data's various columns.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_stock_data()

    ' If the first value in the date column is a string, the subroutine converts all its values
    ' to a Date type.
    If VarType _
            (Worksheets(ActiveSheet.Name) _
                .Cells _
                    (row_global_enumeration.ENUM_K_FIRST_DATA, _
                     column_global_enumeration.ENUM_K_STOCK_DATE) _
                .Value) = vbString Then
        
        change_string_to_date_in_date_column
    
    End If
    

    ' These lines of code change the column formats.
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_STOCK_TICKER) _
        .NumberFormat _
            = "General"
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_STOCK_DATE) _
        .NumberFormat _
            = "mm/dd/yyyy"
            
    For i = column_global_enumeration.ENUM_K_STOCK_OPEN _
        To column_global_enumeration.ENUM_K_STOCK_CLOSE
    
        Worksheets(ActiveSheet.Name).Columns(i).NumberFormat = "$#,##0.00"
    
    Next i  ' This statement ends the first repetition loop.
    
            
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_STOCK_VOL) _
        .NumberFormat _
            = "#,##0"


    ' These lines of code change the column widths.
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_STOCK_TICKER) _
        .ColumnWidth _
            = 10
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_STOCK_DATE) _
        .ColumnWidth _
            = 14
    
    
    For i = column_global_enumeration.ENUM_K_STOCK_OPEN _
        To column_global_enumeration.ENUM_K_STOCK_CLOSE

        Worksheets(ActiveSheet.Name).Columns(i).ColumnWidth = 12
            
    Next i ' This statement ends the second repetition loop.
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_STOCK_VOL) _
        .ColumnWidth _
            = 15
    
End Sub ' This stastement ends the private subroutine, format_stock_data.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_summary_data
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the format for the summary table's columns.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_summary_data()

    ' These lines of code set the formats for the various columns.
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
        .NumberFormat _
            = "General"
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
        .NumberFormat _
            = "#,##0.00"
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
        .NumberFormat _
            = "0.00%"
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
        .NumberFormat _
            = "#,##0"
    
        
    ' These lines of code set the column widths for the various columns.
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
        .ColumnWidth _
            = 10
    
    
    For i = column_global_enumeration.ENUM_K_SUMMARY_YEARLY_CHANGE _
        To column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE

        Worksheets(ActiveSheet.Name).Columns(i).ColumnWidth = 16
            
    Next i ' This statement ends the repetition loop.
    
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
        .ColumnWidth _
            = 25

End Sub ' This statement ends the private subroutine, format_summary_data.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_titles
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the format for the row containing titles for both
 '       the stock and the summary data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_titles()

    Worksheets(ActiveSheet.Name) _
        .Rows(row_global_enumeration.ENUM_K_TITLE) _
        .NumberFormat _
            = "General"
    
    Worksheets(ActiveSheet.Name) _
        .Rows(row_global_enumeration.ENUM_K_TITLE) _
        .Font.Bold _
            = True
    
    Worksheets(ActiveSheet.Name) _
        .Rows(row_global_enumeration.ENUM_K_TITLE) _
        .HorizontalAlignment _
            = xlCenter

End Sub ' This statement ends the private subroutine, format_titles.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_worksheet
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets the font type and font size for a worksheet.
 '
 '  Subroutine Parameters:
 '
 '  Type     Name                  Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '              input_worksheet_name_string
 '                          This parameter is the input worksheet name.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_worksheet(ByVal input_worksheet_name_string As String)
    
    ActiveWorkbook.Sheets(input_worksheet_name_string).Cells.Font.Name = "Garamond"
    
    ActiveWorkbook.Sheets(input_worksheet_name_string).Cells.Font.Size = 14
    
End Sub ' This statement ends the private subroutine, format_worksheet.

'*******************************************************************************************
 '
 '  Subroutine Name:  create_summary_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine creates the summary table by analyzing the raw stock data.
 '
 '  Subroutine Parameters:
 '
 '  Type     Name           Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub create_summary_table()
    
    ' This line of code declares a variables for the stock data's first row index
    ' in the repetition loop.
    Dim first_row_long As Long
    

    ' These lines of code declare variables for the stock record.  The program uses
    ' these values to calculate the summary table record.
    Dim current_ticker_name_string As String
    
    Dim opening_price_currency As Currency
    
    Dim closing_price_currency As Currency
    
    Dim total_stock_volume_variant As Variant
    
    
    ' This line of code declares the variable for the row index in the summary table.
    Dim summary_table_row_long As Long
    
    
    ' This subroutine places the titles in the appropriate cells.
    set_up_titles_for_summary_data
    
    
    ' These lines of code initialize variables with information from the first row
    ' of the raw stock.
    current_ticker_name_string _
        = Worksheets(ActiveSheet.Name) _
                .Cells _
                    (row_global_enumeration.ENUM_K_FIRST_DATA, _
                     column_global_enumeration.ENUM_K_STOCK_TICKER) _
                .Value
    
    opening_price_currency _
        = Worksheets(ActiveSheet.Name) _
                .Cells _
                    (row_global_enumeration.ENUM_K_FIRST_DATA, _
                     column_global_enumeration.ENUM_K_STOCK_OPEN) _
                .Value
                
    total_stock_volume_variant _
        = Worksheets(ActiveSheet.Name) _
                .Cells _
                    (row_global_enumeration.ENUM_K_FIRST_DATA, _
                     column_global_enumeration.ENUM_K_STOCK_VOL) _
                .Value
                
    
    ' These lines of code set the initial row indices for the original data
    ' and summary tables.
    first_row_long = row_global_enumeration.ENUM_K_FIRST_DATA + 1
    
    summary_table_row_long = row_global_enumeration.ENUM_K_FIRST_DATA
 
 
    ' This repetition loop runs through all the rows of the original data
    ' and generates the summary table: the loop starts with the second
    ' row of original data.
    For rw = first_row_long To last_data_row_global_long
    
        If Worksheets(ActiveSheet.Name) _
                .Cells _
                    (rw, _
                     column_global_enumeration.ENUM_K_STOCK_TICKER) _
                .Value _
            = current_ticker_name_string Then
        
        
            ' If the ticker name is the same, this line of code adds the current stock volume
            ' to the total.
            total_stock_volume_variant _
                = total_stock_volume_variant _
                    + Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_VOL) _
                            .Value
                            
                            
            ' If the loop has reached the last row the program creates a summary table record.
            If rw = last_data_row_global_long Then
                                           
                create_summary_data_row _
                    current_ticker_name_string, _
                    opening_price_currency, _
                    total_stock_volume_variant, _
                    summary_table_row_long, _
                    rw, True
                    
            End If
            
        Else
        
            ' This selection statement executes if the repetition loop has not reached
            ' the end of the data.
            If rw <> last_data_row_global_long Then
            
                ' If the current ticker does not match the previous ticker,
                ' the script creates a record.
                create_summary_data_row _
                    current_ticker_name_string, _
                    opening_price_currency, _
                    total_stock_volume_variant, _
                    summary_table_row_long, _
                    rw, False
                    
                
                ' These lines of code assign new values to the stock data variables.
                current_ticker_name_string _
                    = Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_TICKER) _
                            .Value
                
                opening_price_currency _
                    = Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_OPEN) _
                            .Value
                
                total_stock_volume_variant _
                    = Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_VOL) _
                            .Value
            
                
               ' This line of code increases the summary table row index
               ' by one for the next record.
                summary_table_row_long = summary_table_row_long + 1
                        
            Else
            
                ' These lines of code initialize variables with information
                ' from the stock data's last row.
                current_ticker_name_string _
                    = Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_TICKER) _
                            .Value
                    
                opening_price_currency _
                    = Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_OPEN) _
                            .Value
                            
                total_stock_volume_variant _
                    = total_stock_volume_variant _
                        + Worksheets(ActiveSheet.Name) _
                                .Cells _
                                    (rw, _
                                     column_global_enumeration.ENUM_K_STOCK_VOL) _
                                .Value
                                
                                
                ' The program then creates a record with this information.
                create_summary_data_row _
                        current_ticker_name_string, _
                        opening_price_currency, _
                        total_stock_volume_variant, _
                        summary_table_row_long, _
                        rw, True
            
            End If
            
        End If
        
    Next rw ' This statement ends the repetition loop.
    
End Sub ' This statement ends the private subroutine, create_summary_table.

'*******************************************************************************************
 '
 '  Subroutine Name:  create_change_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine creates a table that lists the tickers with the greatest percent change
 '      and the ticker with the greatest total stock volume.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Public Sub create_change_table()

    format_change_data_titles
    
    set_up_change_data_titles
    
    calculate_and_write_change_data

End Sub ' This statement ends the private subroutine, create_change_table.

'*******************************************************************************************
 '
 '  Subroutine Name:  convert_data_and_summary_ranges_to_tables
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine converts the stock data and summary ranges to tables.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub convert_data_and_summary_ranges_to_tables()
    
    convert_range_into_table _
        row_global_enumeration.ENUM_K_TITLE, _
        column_global_enumeration.ENUM_K_STOCK_TICKER, _
        "StockData"

    convert_range_into_table _
        row_global_enumeration.ENUM_K_TITLE, _
        column_global_enumeration.ENUM_K_SUMMARY_TICKER, _
        "Summary"
    
End Sub ' This statement ends the private subroutine, convert_data_and_summary_ranges_to_tables.

'*******************************************************************************************
 '
 '  Subroutine Name:  create_format_analysis_worksheet
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This function creates the Analysis Worksheet, formats it, copies three summary
 '      summary tables over to it, and sorts those tables.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub create_format_analysis_worksheet()
    
    Dim primary_sheet_name_string As String
    
    Dim analysis_sheet_name_string As String
            
    Dim analysis_worksheet_exists_boolean As Boolean

  
    ' This line of code saves the primary worksheet's name to a variable.
    primary_sheet_name_string = ActiveWorkbook.ActiveSheet.Name
                 
                
    ' This line of code saves the primary worksheet's name to a variable.
                
    analysis_sheet_name_string _
        = return_analysis_worksheet_name_function(primary_sheet_name_string)
        
    
    ' This line of code checks if the analysis worksheet exists.
    On Error Resume Next
    
    analysis_worksheet_exists_boolean _
        = (ActiveWorkbook.Sheets(analysis_sheet_name_string).Index > 0)
    
       
    If analysis_worksheet_exists_boolean = False Then
       
        ' This subroutine creates and activates the analysis worksheet.
        create_worksheet (analysis_sheet_name_string)
        
    
        ' This subroutine formats the Analysis Worksheet.
        format_analysis_worksheet (analysis_sheet_name_string)
        
        
        ' This subroutine copies, sorts, and renames the summary table three times.
        set_up_sorted_tables _
            primary_sheet_name_string, _
            analysis_sheet_name_string
        
        ActiveWorkbook.Worksheets(primary_sheet_name_string).Activate
    
    Else
    
          MsgBox _
            ("Please delete the Analysis Worksheet and select the Summary Table Worksheet before proceeding!")
      
    End If
        
End Sub ' This statement ends the private subroutine, create_format_analysis_worksheet.

'*******************************************************************************************
 '
 '  Subroutine Name:  change_string_to_date_in_date_column
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine changes the text strings in the stock data's date column
 '        to a Date type.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub change_string_to_date_in_date_column()

    '  This line of code declares a variable for the current date.
    Dim current_date As Date
    
    ' These lines of code declare variables for the current year, month, and day.
    Dim current_year_integer As Integer
    
    Dim current_month_integer As Integer
    
    Dim current_day_integer As Integer
       
    ' These lines of code declare variables for the start indexes in the date string.
    Dim year_index_integer As Integer
    
    Dim month_index_integer As Integer
    
    Dim day_index_integer As Integer
    
    
    ' These lines of code initialize variables for the start indices.
    year_index_integer = 1
        
    month_index_integer = year_index_integer + GLOBAL_CONSTANT_YEAR_LENGTH
        
    day_index_integer = month_index_integer + GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH
    
    
    ' These lines of code loop through all the values in the specified column
    ' and converts them to a Date type.
    For rw = row_global_enumeration.ENUM_K_FIRST_DATA To last_data_row_global_long
    
        ' These lines of code parse out the date from the string, YYYYMMDD,
        ' in the current cell and converts it to a Date type.
        current_year_integer _
            = Mid(Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_DATE) _
                            .Value, _
                      year_index_integer, _
                      GLOBAL_CONSTANT_YEAR_LENGTH)
        
        current_month_integer _
            = Mid(Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_DATE) _
                            .Value, _
                      month_index_integer, _
                      GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH)
        
        current_day_integer _
            = Mid(Worksheets(ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_STOCK_DATE) _
                            .Value, _
                      day_index_integer, _
                      GLOBAL_CONSTANT_MONTH_OR_DAY_LENGTH)
        
        ' This line of code takes the values for year, month, and day, converts them
        ' to a Date type, then  assigns them to the appropriate variable
        current_date _
            = DateSerial(current_year_integer, current_month_integer, current_day_integer)
    
    
        ' This line of code assigns the new date value to the current cell.
        Worksheets(ActiveSheet.Name) _
            .Cells _
                (rw, _
                 column_global_enumeration.ENUM_K_STOCK_DATE) _
            .Value _
                = current_date
    
    Next rw ' This statement ends the repetition loop.

End Sub ' This statement ends the private subroutine, change_string_to_date_in_date_column.

'*******************************************************************************************
 '
 '  Subroutine Name:  set_up_titles_for_summary_data
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sets up the titles for the summary data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub set_up_titles_for_summary_data()

    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_TITLE, _
             column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
        .Value _
            = "Ticker"
    
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_TITLE, _
             column_global_enumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
        .Value _
            = "Yearly Change"
    
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_TITLE, _
             column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
        .Value _
            = "Percent Change"
    
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_TITLE, _
             column_global_enumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
        .Value _
            = "Total Stock Volume"
    
End Sub  ' This statement ends the public subroutine, set_up_titles_for_summary_data.

'*******************************************************************************************
 '
 '  Subroutine Name:  create_summary_data_row
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine creates a summary data record based on the parameters.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          tickerNameStringParameter
 '                          This parameter holds the name of the stock ticker.
 '  Currency
 '          openingPriceCurrencyParameter
 '                          This parameter is the first opening price for this
 '                          stock ticker.
 '  Variant
 '          totalStockVolumeVariantParameter
 '                          This parameter is the total stock volume for this
 '                          stock ticker.
 '  Long Integer
 '          summaryRowLongParameter
 '                          This parameter is the current summary table row index.
 '  Long Integer
 '          originalRowLongParameter
 '                          This parameter is the current original data row index.
 '  Boolean
 '          lastRowFlagBooleanParameter
 '                          This parameter indicates whether the program
 '                          has reached the last record or not.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub create_summary_data_row _
    (ByVal ticker_name_string As String, _
     ByVal opening_price_currency As Currency, _
     ByVal total_stock_volume_variant As Variant, _
     ByVal summary_row_long As Long, _
     ByVal original_row_long As Long, _
     ByVal last_row_flag_boolean As Boolean)

    ' This line of code declares a variable for the closing price which is different
    ' based on whether the program has reached the last row or not in the
    ' raw stock data.
    Dim closing_price_currency As Currency


    ' If the script has not reached the last row, the closing price comes
    ' from the previous row in the raw stock data; otherwise, the closing
    ' price comes from the current row.
    If last_row_flag_boolean = False Then
            
        closing_price_currency _
            = Worksheets(ActiveSheet.Name) _
                    .Cells _
                        (original_row_long - 1, _
                         column_global_enumeration.ENUM_K_STOCK_CLOSE) _
                    .Value
            
    Else
            
        closing_price_currency _
            = Worksheets(ActiveSheet.Name) _
                    .Cells _
                        (original_row_long, _
                         column_global_enumeration.ENUM_K_STOCK_CLOSE) _
                    .Value
            
    End If
            
            
    ' These lines of code create a record in the summary data.
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (summary_row_long, _
             column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
        .Value _
            = ticker_name_string
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (summary_row_long, _
             column_global_enumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
        .Value _
            = calculate_yearly_change_function _
                    (CDbl(opening_price_currency), _
                     CDbl(closing_price_currency))
                            
    format_yearly_change_cell (summary_row_long)
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (summary_row_long, _
             column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
        .Value _
            = calculate_percent_change_function _
                    (CDbl(opening_price_currency), _
                     CDbl(closing_price_currency))
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (summary_row_long, _
             column_global_enumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
        .Value _
            = total_stock_volume_variant

End Sub ' This statement ends the private subroutine,
' create_summary_data_row.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_change_data_titles
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine formats the row and column titles in the change data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_change_data_titles()

    ' These lines of code format the columns and cells of the change data.
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_CHANGE_ROW_TITLES) _
        .NumberFormat _
            = "General"
    
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_CHANGE_TICKERS) _
        .NumberFormat _
            = "General"
    
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_FIRST_DATA, _
             column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .NumberFormat _
            = "0.00%"
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_PERCENT_DECREASE, _
             column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .NumberFormat _
            = "0.00%"
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
             column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .NumberFormat _
            = "#,##0"
            
           
    ' These lines of code set the column widths for the change data.
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_CHANGE_ROW_TITLES) _
        .ColumnWidth _
            = 25
            
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_CHANGE_TICKERS) _
        .ColumnWidth _
            = 10
            
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .ColumnWidth _
            = 25
            
            
    ' This line of code sets the font style for the row titles to bold.
    Worksheets(ActiveSheet.Name) _
        .Columns(column_global_enumeration.ENUM_K_CHANGE_ROW_TITLES) _
        .Font _
        .Bold _
            = True

End Sub ' This statement ends the private subroutine, format_change_data_titles.

'*******************************************************************************************
 '
 '  Subroutine Name:  set_up_change_data_titles
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine writes the column and row titles to the appropriate cells
 '      for the change data.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub set_up_change_data_titles()

    ' These lines of code set the column titles in the change table.
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_TITLE, _
             column_global_enumeration.ENUM_K_CHANGE_TICKERS) _
        .Value _
            = "Ticker"
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_TITLE, _
             column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .Value _
            = "Value"
            
            
    ' These lines of code set the row titles in the change table.
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_FIRST_DATA, _
             column_global_enumeration.ENUM_K_CHANGE_ROW_TITLES) _
        .Value _
            = "Greatest % Increase"
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_PERCENT_DECREASE, _
             column_global_enumeration.ENUM_K_CHANGE_ROW_TITLES) _
        .Value _
            = "Greatest % Decrease"
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
             column_global_enumeration.ENUM_K_CHANGE_ROW_TITLES) _
        .Value _
            = "Greatest Total Volume"

End Sub ' This statement ends the private subroutine, set_up_change_data_titles.

'*******************************************************************************************
 '
 '  Subroutine Name:  calculate_and_write_change_data
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine calculates the values for the change data based on raw stock
 '      data and writes the results to the change table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  n/a     n/a             n/a
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub calculate_and_write_change_data()
    
    Dim increase_ticker_string As String
    
    Dim decrease_ticker_string As String
    
    Dim volume_ticker_string As String
    
    Dim increase_percentage_double As Double
    
    Dim decrease_percentage_double As Double
    
    Dim volume_variant As Variant

    Dim first_row_long As Long
    
    Dim last_row_long As Long
    
    
    ' These lines of code initialize the variables with the first record
    ' in the summary data.
    increase_ticker_string _
        = Worksheets(ActiveSheet.Name) _
                .Cells _
                    (row_global_enumeration.ENUM_K_FIRST_DATA, _
                     column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
                .Value
         
    decrease_ticker_string = increase_ticker_string
        
    volume_ticker_string = increase_ticker_string
        
    increase_percentage_double _
        = Worksheets(ActiveSheet.Name) _
                .Cells _
                    (row_global_enumeration.ENUM_K_FIRST_DATA, _
                     column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                .Value
             
    decrease_ticker_string = increase_percentage_double
        
    volume_variant _
        = Worksheets(ActiveSheet.Name) _
                .Cells _
                    (row_global_enumeration.ENUM_K_FIRST_DATA, _
                     column_global_enumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                .Value
    
    
    ' These lines of code initialize the first and last index of the repetition loop.
    first_row_long = row_global_enumeration.ENUM_K_FIRST_DATA + 1
        
    last_row_long _
        = Worksheets(ActiveSheet.Name) _
                .Cells _
                    (Rows.Count, _
                     column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
                .End(xlUp) _
                .Row
        
        
    ' This repetition loop starts at the second record of the summary data and,
    ' through comparisons, finds the tickers with the greatest increase, greatest
    ' decrease, and greatest total stock volume.
    For rw = first_row_long To last_row_long
    
        ' If a record has a larger change in percentage than the previous holder,
        ' set it as the new leader in percentage increase.
        If Worksheets(ActiveSheet.Name) _
                .Cells _
                    (rw, _
                     column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                .Value _
            > increase_percentage_double Then
        
            increase_ticker_string _
                = Worksheets(ActiveSheet.Name) _
                        .Cells _
                            (rw, _
                             column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
                        .Value
                        
            increase_percentage_double _
                = Worksheets(ActiveSheet.Name) _
                        .Cells _
                            (rw, _
                             column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                        .Value
        
        End If
        
        
        ' If a record has a smaller change in percentage than the previous holder,
        ' set it as the new leader in percentage decrease.
        If Worksheets(ActiveSheet.Name) _
                .Cells _
                    (rw, _
                     column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                .Value _
            < decrease_percentage_double Then
        
            decrease_ticker_string _
                = Worksheets(ActiveSheet.Name) _
                        .Cells _
                            (rw, _
                             column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
                        .Value
                        
            decrease_percentage_double _
                = Worksheets(ActiveSheet.Name) _
                        .Cells _
                            (rw, _
                             column_global_enumeration.ENUM_K_SUMMARY_PERCENT_CHANGE) _
                        .Value
        
        End If
        
        
        ' If a record has a larger total stock volume than the previous holder,
        ' set it as the new leader in total stock volume.
        If Worksheets(ActiveSheet.Name) _
                .Cells _
                    (rw, _
                     column_global_enumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                .Value _
            > volume_variant Then
        
            volume_ticker_string _
                = Worksheets _
                        (ActiveSheet.Name) _
                            .Cells _
                                (rw, _
                                 column_global_enumeration.ENUM_K_SUMMARY_TICKER) _
                            .Value
                        
            volume_variant _
                = Worksheets(ActiveSheet.Name) _
                        .Cells _
                            (rw, _
                             column_global_enumeration.ENUM_K_SUMMARY_TOTAL_STOCK_VOLUME) _
                        .Value
        
        End If
    
    Next rw ' This statement ends the repetition loop.
             
    
    ' These lines of code write the results to the change data.
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_FIRST_DATA, _
             column_global_enumeration.ENUM_K_CHANGE_TICKERS) _
        .Value _
            = increase_ticker_string
    
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_FIRST_DATA, _
             column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .Value _
            = increase_percentage_double
            
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_PERCENT_DECREASE, _
             column_global_enumeration.ENUM_K_CHANGE_TICKERS) _
        .Value _
            = decrease_ticker_string
    
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_PERCENT_DECREASE, _
             column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .Value _
            = decrease_percentage_double
                
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
             column_global_enumeration.ENUM_K_CHANGE_TICKERS) _
        .Value _
            = volume_ticker_string
    
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_global_enumeration.ENUM_K_GREATEST_TOTAL_VOLUME, _
             column_global_enumeration.ENUM_K_CHANGE_VALUES) _
        .Value _
            = volume_variant
    
End Sub ' This statement ends the private subroutine, calculate_and_write_change_data.

'*******************************************************************************************
 '
 '  Subroutine Name:  convert_range_into_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This subroutine converts a range of data into a table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  Integer
 '          row_integer
 '                          This is the row number of the upper left corner of the range.
 '  Integer
 '          column_integer
 '                          This is the column number of the upper left corner of the range.
 '  String
 '          table_name_string
 '                          This is the name of the new table.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/20/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub convert_range_into_table _
    (ByVal row_integer As Integer, _
     ByVal column_integer As Integer, _
     ByVal table_name_string As String)
    
    Dim temp_list_object As ListObject
    
    Dim current_table_name_string As String
    
    
    ' This line of code selects the range of data.
    Worksheets(ActiveSheet.Name) _
        .Cells _
            (row_integer, _
             column_integer) _
        .Select
    
    
    ' This line of code assigns the selected range of data to a ListObject.
    On Error Resume Next
        
        Set temp_list_object _
            = Worksheets(ActiveSheet.Name) _
                    .Cells _
                        (row_integer, _
                         column_integer) _
                    .ListObject
    
    On Error GoTo 0
    
    
    ' If there is no ListObject, the script converts the range to a table.
    If temp_list_object Is Nothing Then
    
        current_table_name_string = table_name_string & ActiveSheet.Name & "Table"
            
        ActiveSheet.ListObjects _
            .Add _
                (xlSrcRange, _
                 Selection.CurrentRegion, , _
                 xlYes) _
            .Name _
                = current_table_name_string
    
    End If
                
End Sub ' This statement ends the private subroutine, convert_range_into_table.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_yearly_change_cell
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '      This function formats the newly assigned yearly change cell in the summary data
 '      based on the row index.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  Integer
 '          row_index_integer
 '                          This parameter holds the row index for the current record
 '                          in the summary table.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_yearly_change_cell _
    (ByVal row_index_integer As Integer)
    
    ' If the yearly change is zero or positive, the script changes the background color
    ' to green.
    If Worksheets(ActiveSheet.Name) _
            .Cells _
                (row_index_integer, _
                 column_global_enumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
            .Value _
        >= 0 Then
        
        Worksheets(ActiveSheet.Name) _
            .Cells _
                (row_index_integer, _
                 column_global_enumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
            .Interior _
            .ColorIndex _
                = 4
        
    Else ' If the yearly change is negative, the script changes the background color to red.
    
        Worksheets(ActiveSheet.Name) _
            .Cells _
                (row_index_integer, _
                 column_global_enumeration.ENUM_K_SUMMARY_YEARLY_CHANGE) _
            .Interior _
            .ColorIndex _
                = 3
    
    End If
    
End Sub ' This statement ends the private subroutine, format_yearly_change_cell.

'*******************************************************************************************
 '
 '  Subroutine Name:  create_worksheet
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine creates a worksheet if it does not already exist.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          worksheet_name_string
 '                          This parameter is the name of the input worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub create_worksheet _
    (ByVal worksheet_name_string As String)
                
    Dim worksheet_exists_boolean As Boolean
    
    
    ' This line of code checks if the worksheet exists.
    On Error Resume Next
    
    worksheet_exists_boolean _
        = ActiveWorkbook.Sheets(workSheetNameStringParameter).Index > 0
    
    
    ' This selection statement creates the worksheet if it does not exist.
    If worksheet_exists_boolean = False Then
    
        Sheets.Add.Name = worksheet_name_string
      
    End If

End Sub ' This statement ends the private subroutine, create_worksheet.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_analysis_worksheet
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine formats the Analysis Worksheet.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          worksheet_name_string
 '                          This parameter is the name of the input worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_analysis_worksheet _
    (ByVal worksheet_name_string As String)
    
    '  This repetition loop formats the columns for three summary tables.
    For column_index = 1 To 11 Step 5
    
        format_local_summary_data _
            worksheet_name_string, _
            column_index
    
    Next column_index ' This statement ends the first repetition loop.
    
    
    ' This subroutine formats the table title row.
    format_titles
  
  
    ' This subroutine inserts the table titles.
    insert_analysis_worksheet_row_and_titles (worksheet_name_string)
  
  
    ' This subroutine formats the whole worksheet.
    format_worksheet (worksheet_name_string)
    
End Sub ' This statement ends the private subroutine, format_analysis_worksheet.

'*******************************************************************************************
 '
 '  Subroutine Name:  format_local_summary_data
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine formats a section of a worksheet for a summary table.
 '
 '  Subroutine Parameters:
 '
 '  Type     Name                  Description
 '  -----   -------------   ----------------------------------------------
 '  -----   -------------   ----------------------------------------------
 '  String
 '          worksheet_name_string
 '                          This parameter is the name of the input worksheet.
 '  Integer
 '          column_number_integer
 '                          This parameter is the column number.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub format_local_summary_data _
    (ByVal worksheet_name_string As String, _
     ByVal column_number_integer As Integer)
    
     ' These lines of code set the formats for the various columns.
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer) _
        .NumberFormat _
            = "General"
    
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer + 1) _
        .NumberFormat _
            = "#,##0.00"
    
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer + 2) _
        .NumberFormat _
            = "0.00%"
    
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer + 3) _
        .NumberFormat _
            = "#,##0"
    
        
    ' These lines of code set the column widths for the various columns.
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer) _
        .ColumnWidth _
            = 10
                
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer + 1) _
        .ColumnWidth _
            = 16
    
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer + 2) _
        .ColumnWidth _
            = 16
    
    Worksheets(worksheet_name_string) _
        .Columns _
            (column_number_integer + 3) _
        .ColumnWidth _
            = 25
    
End Sub ' This statement ends the private subroutine, format_local_summary_data.

'*******************************************************************************************
 '
 '  Subroutine Name:  insert_analysis_worksheet_row_and_titles
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine inserts a row at the top of the worksheet, formats it,
 '       and writes titles to it.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          worksheet_name_string
 '                          This parameter is the name of the input worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub insert_analysis_worksheet_row_and_titles _
    (ByVal worksheet_name_string As String)
    
    ' This line of code inserts a row for the table titles.
    Worksheets(worksheet_name_string) _
        .Range("A1") _
        .EntireRow _
        .Insert
            
            
    ' This line of code formats the font for the first row as bold.
     Worksheets(worksheet_name_string) _
        .Range("A1") _
        .EntireRow _
        .Font _
        .Bold _
            = True
             
             
    ' These lines of code write the table titles to the appropriate cells.
    Worksheets(worksheet_name_string) _
        .Cells _
            (1, 1) _
        .Value _
            = "Greatest % Increase"
            
    Worksheets(worksheet_name_string) _
        .Cells _
            (1, 6) _
        .Value _
            = "Greatest % Decrease"
            
    Worksheets(worksheet_name_string) _
        .Cells _
            (1, 11) _
        .Value _
            = "Greatest Total Volume"

End Sub ' This statement ends the private subroutine, insert_analysis_worksheet_row_and_titles.

'*******************************************************************************************
 '
 '  Subroutine Name:  set_up_sorted_tables
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine insert a row at the top of the worksheet, formats it,
 '       and writes titles to it.
 '
 '  Subroutine Parameters:
 '
 '  Type     Name                  Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          primary_sheet_name_string
 '                          This parameter is the name of the Summary Worksheet.
 '  String
 '          analysis_sheet_name_string
 '                          This parameter is the name of the Analysis Worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub set_up_sorted_tables _
    (ByVal primary_sheet_name_string As String, _
     ByVal analysis_sheet_name_string As String)
                
    set_up_greatest_increase_table _
        primary_sheet_name_string, _
        analysis_sheet_name_string
    
    set_up_greatest_decrease_table _
        primary_sheet_name_string, _
        analysis_sheet_name_string

    set_up_greatest_total_volume_table _
        primary_sheet_name_string, _
        analysis_sheet_name_string
        
End Sub ' This statement ends the private subroutine, set_up_sorted_tables.

'*******************************************************************************************
 '
 '  Subroutine Name:  set_up_greatest_increase_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine copies a summary table to the Analysis Worksheet and sorts it
 '       by the percent change column in descending order.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          primary_sheet_name_string
 '                          This parameter is the name of the Summary Worksheet.
 '  String
 '          analysis_sheet_name_string
 '                          This parameter is the name of the Analysis Worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub set_up_greatest_increase_table _
    (ByVal primary_sheet_name_string As String, _
     ByVal analysis_sheet_name_string As String)

    ' This line of code copies the first summary table from the Summary Worksheet
    ' to the Analysis Spreadsheet.
    copy_table _
        primary_sheet_name_string, _
        analysis_sheet_name_string, _
         "Summary" & primary_sheet_name_string & "Table[#All]", _
        "A2"
 
 
    ' This line of code renames the copied table.
    Worksheets(analysis_sheet_name_string) _
        .ListObjects(1) _
        .Name _
            = "GreatestIncreaseTable"
     
     
     ' This line of code sorts the copied table, Greatest Increase Table.
     sort_table _
        analysis_sheet_name_string, _
        "GreatestIncreaseTable", _
        "Percent Change", _
        True

End Sub ' This statement ends the private subroutine, set_up_greatest_increase_table.
        
 '*******************************************************************************************
 '
 '  Subroutine Name:  set_up_greatest_decrease_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine copies a summary table to the Analysis Worksheet and sorts it
 '       by the percent change column in ascending order.
 '
 '  Subroutine Parameters:
 '
 '  Type     Name                  Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          primary_sheet_name_string
 '                          This parameter is the name of the Summary Worksheet.
 '  String
 '          analysis_sheet_name_string
 '                          This parameter is the name of the Analysis Worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/
         
Private Sub set_up_greatest_decrease_table _
    (ByVal primary_sheet_name_string As String, _
     ByVal analysis_sheet_name_string As String)

    ' This line of code copies the first summary table.
    copy_table _
        primary_sheet_name_string, _
        analysis_sheet_name_string, _
         "Summary" & primary_sheet_name_string & "Table[#All]", _
        "F2"
 
 
    ' This line of code renames the copied table.
    Worksheets(analysis_sheet_name_string) _
        .ListObjects(2) _
        .Name _
            = "GreatestDecreaseTable"
     
     
     ' This line of code sorts the Greatest Increase Table.
     sort_table _
        analysis_sheet_name_string, _
        "GreatestDecreaseTable", _
        "Percent Change", _
        False

End Sub ' This statement ends the private subroutine, set_up_greatest_decrease_table.
        
 '*******************************************************************************************
 '
 '  Subroutine Name:  set_up_greatest_total_volume_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine copies a summary table to the Analysis Worksheet and sorts it
 '       by the total volume column in descending order.
 '
 '  Subroutine Parameters:
 '
 '  Type     Name                  Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          primary_sheet_name_string
 '                          This parameter is the name of the Summary Worksheet.
 '  String
 '          analysis_sheet_name_string
 '                          This parameter is the name of the Analysis Worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub set_up_greatest_total_volume_table _
    (ByVal primary_sheet_name_string As String, _
     ByVal analysis_sheet_name_string As String)

    ' This line of code copies the first summary table.
    copy_table _
        primary_sheet_name_string, _
        analysis_sheet_name_string, _
         "Summary" & primary_sheet_name_string & "Table[#All]", _
        "K2"
 
 
    ' This line of code renames the copied table.
    Worksheets(analysis_sheet_name_string) _
        .ListObjects(3) _
        .Name _
            = "GreatestTotalVolumeTable"
     
     
     ' This line of code sorts the Greatest Increase Table.
     sort_table _
        analysis_sheet_name_string, _
        "GreatestTotalVolumeTable", _
        "Total Stock Volume", _
        True

End Sub ' This statement ends the private subroutine, set_up_greatest_total_volume_table.

'*******************************************************************************************
 '
 '  Subroutine Name:  copy_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine copies a table from one location to another.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          primary_sheet_name_string
 '                          This parameter is the name of the Summary Worksheet.
 '  String
 '          analysis_sheet_name_string
 '                          This parameter is the name of the Analysis Worksheet.
 '  String
 '          table_name_string
 '                          This parameter is the name of the input table.
 '  String
 '          destination_string
 '                          This parameter is the location of the copied table (i.e., "A1").
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub copy_table _
    (ByVal primary_sheet_name_string As String, _
     ByVal analysis_sheet_name_string As String, _
     ByVal table_name_string As String, _
     ByVal destination_string As String)
                
    Dim current_range As Range
 
 
    ' This repetition loopiterates through the table and assigns the rows
    ' to a Range object.
    For Each Row In Worksheets(primary_sheet_name_string).Range(table_name_string).Rows
 
        If Row.EntireRow.Hidden = False Then
    
            If current_range Is Nothing Then
                
                Set current_range = Row
            
            End If
            
            Set current_range = Union(Row, current_range)
        
        End If
 
    Next Row ' This statement ends the first repetition loop.
 
 
    ' This line of code copies the Range Object to the destination.
    current_range.Copy _
        Destination _
            :=Worksheets(analysis_sheet_name_string).Range(destination_string)
                                                
End Sub ' This statement ends the private subroutine, copy_table.

'*******************************************************************************************
 '
 '  Subroutine Name:  sort_table
 '
 '  Subroutine Type: Private
 '
 '  Subroutine Description:
 '       This subroutine sorts a table.
 '
 '  Subroutine Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          worksheet_name_string
 '                          This parameter is the name of the input worksheet.
 '  String
 '          table_name_string
 '                          This parameter is the name of the input table.
 '  String
 '          column_name_stringParameter
 '                          This parameter is the name of the table column for sorting.
 '  Boolean
 '          descending_flag_boolean
 '                          This parameter indicates whether the sorting is in descending
 '                          or ascending order.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Sub sort_table _
    (ByVal worksheet_name_string As String, _
     ByVal table_name_string As String, _
     ByVal column_name_string As String, _
     ByVal descending_flag_boolean As Boolean)

    Dim table_list_object As ListObject
    
    Dim sort_column_range As Range


    Set table_list_object _
        = Worksheets(worksheet_name_string) _
                .ListObjects(table_name_string)

    Set sort_column_range _
        = Range(table_name_string & "[" & column_name_string & "]")
                
    
    If descending_flag_boolean = True Then
                
        With table_list_object.Sort
        
            .SortFields.Clear
        
            .SortFields _
                .Add _
                    Key:=sort_column_range, _
                    SortOn:=xlSortOnValues, _
                    Order:=xlDescending
        
            .Header = xlYes
        
            .Apply

        End With
        
    Else
    
        With table_list_object.Sort
        
            .SortFields.Clear
        
            .SortFields _
                .Add _
                    Key:=sort_column_range, _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending
        
            .Header = xlYes
        
            .Apply

        End With
    
    End If

End Sub ' This statement ends the private subroutine, sort_table.

'*******************************************************************************************
 '
 '  Function Name:  calculate_yearly_change_function
 '
 '  Function Type: Private
 '
 '  Function Description:
 '      This function calculates the yearly change between the first opening price
 '      and the last closing price of the year for a single ticker.
 '
 '  Function Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  Double
 '          opening_price_double
 '                          This parameter holds the first opening price of a ticker.
 '  Double
 '          closing_price_double
 '                          This parameter holds the last closing price of a ticker.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Function calculate_yearly_change_function _
    (ByVal opening_price_double As Double, _
     ByVal closing_price_double As Double) _
As Double
    
    calculate_yearly_change_function = closing_price_double - opening_price_double
            
End Function ' This statement ends the private function, calculate_yearly_change_function.

'*******************************************************************************************
 '
 '  Function Name:  calculate_percent_change_function
 '
 '  Function Type: Private
 '
 '  Function Description:
 '      This function calculates the percent change between the first opening price
 '      and the last closing price of the year for a ticker.
 '
 '  Function Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  Double
 '          opening_price_double
 '                          This parameter holds the first opening price of a ticker.
 '  Double
 '          closing_price_double
 '                          This parameter holds the last closing price of a ticker.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Function calculate_percent_change_function _
    (ByVal opening_price_double As Double, _
     ByVal closing_price_double As Double) _
As Double

    calculate_percent_change_function _
        = (closing_price_double - opening_price_double) / opening_price_double

End Function ' This statement ends the private function, calculate_percent_change_function.

'*******************************************************************************************
 '
 '  Function Name:  return_analysis_worksheet_name_function
 '
 '  Function Type: Private
 '
 '  Function Description:
 '      This function returns the Analysis Worksheet name based on the Summary Table
 '      Worksheet.
 '
 '  Function Parameters:
 '
 '  Type    Name            Description
 '  -----   -------------   ----------------------------------------------
 '  String
 '          worksheet_name_string
 '                          This parameter is the name of the Summary Table Worksheet.
 '
 '
 '  Date               Description                              Programmer
 '  ---------------    ------------------------------------     ------------------
 '  07/19/2023         Initial Development                      Nicholas J. George
 '
 '******************************************************************************************/

Private Function return_analysis_worksheet_name_function _
    (ByVal worksheet_name_string As String) _
As String

    return_analysis_worksheet_name_function _
        = "Analysis " & worksheet_name_string

End Function ' This statement ends the private function, return_analysis_worksheet_name_function.
