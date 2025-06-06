# Excel Date to Weekday VBA Macro

## Overview
This Excel VBA macro automatically converts dates entered in Column A to their corresponding Japanese weekday names (e.g., "金曜日" for Friday) in Column B. It handles a wide variety of date formats, including Japanese-specific formats like "令和7年6月6日" or "6月6日", using the `Worksheet_Change` event for real-time updates without requiring a button.

## Features
- **Automatic Execution**: Updates the weekday in Column B whenever a date is entered or modified in Column A.
- **Flexible Date Parsing**: Supports multiple date formats, including:
  - Standard: `2025/6/6`, `2025-06-06`, `2025.6.6`
  - Japanese: `2025年6月6日`, `令和7年6月6日`, `6月6日` (assumes current year if omitted)
  - Others: `20250606`, `6/6`, full-width characters (e.g., `２０２５年`)
- **Error Handling**: Displays "無効な日付" (Invalid date) for non-date inputs and handles errors gracefully.
- **Year Completion**: If the year is omitted (e.g., `6月6日`), it defaults to the current year (e.g., 2025).

## Prerequisites
- Microsoft Excel 2007 or later (tested on Windows with Japanese regional settings).
- Excel file saved as `.xlsm` (macro-enabled workbook).

## Installation
1. **Open Excel and Enable Developer Tab**:
   - Go to `File` → `Options` → `Customize Ribbon` → Check `Developer` checkbox.
2. **Access VBA Editor**:
   - Press `Alt + F11` to open the VBA Editor.
3. **Insert Code into Sheet Module**:
   - In the Project Explorer, find `ThisWorkbook` → `Microsoft Excel Objects` → `Sheet1` (or your target sheet).
   - Double-click the sheet and paste the following VBA code:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Monitor Column A (1st column)
    If Not Intersect(Target, Me.Columns(1)) Is Nothing Then
        Application.EnableEvents = False ' Prevent infinite loop
        Dim cell As Range
        Dim inputValue As String
        Dim parsedDate As Date
        Dim outputCell As Range
        
        For Each cell In Target
            ' Process only cells in Column A
            If cell.Column = 1 Then
                Set outputCell = cell.Offset(0, 1) ' Adjacent cell in Column B
                inputValue = Trim(cell.Value) ' Remove leading/trailing spaces
                
                ' Clear output for empty input
                If inputValue = "" Then
                    outputCell.Value = ""
                    GoTo NextCell
                End If
                
                ' Try parsing the date
                If IsDate(inputValue) Then
                    parsedDate = CDate(inputValue)
                    outputCell.Value = Format(parsedDate, "aaaa") ' Japanese weekday
                Else
                    ' Handle Japanese-specific formats
                    If ParseJapaneseDate(inputValue, parsedDate) Then
                        outputCell.Value = Format(parsedDate, "aaaa")
                    Else
                        outputCell.Value = "無効な日付"
                    End If
                End If
            End If
NextCell:
        Next cell
        Application.EnableEvents = True ' Re-enable events
    End If
End Sub

Private Function ParseJapaneseDate(ByVal input As String, ByRef outDate As Date) As Boolean
    ' Parse Japanese-specific date formats (e.g., 令和7年6月6日, 6月6日)
    Dim reiwaYear As Integer
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim temp As String
    Dim parts() As String
    
    On Error GoTo ErrorHandler
    temp = Replace(input, " ", "") ' Remove spaces
    temp = Replace(temp, "年", "/") ' Replace 年 with /
    temp = Replace(temp, "月", "/") ' Replace 月 with /
    temp = Replace(temp, "日", "") ' Remove 日
    
    ' Convert full-width to half-width characters
    temp = StrConv(temp, vbNarrow)
    
    ' Handle Reiwa era
    If InStr(temp, "令和") > 0 Then
        temp = Replace(temp, "令和", "")
        parts = Split(temp, "/")
        If UBound(parts) >= 2 Then
            reiwaYear = CInt(parts(0))
            year = 2018 + reiwaYear ' Reiwa 1 = 2019
            month = CInt(parts(1))
            day = CInt(parts(2))
            outDate = DateSerial(year, month, day)
            ParseJapaneseDate = True
            Exit Function
        End If
    End If
    
    ' Handle year-omitted dates (e.g., 6月6日) with current year
    If InStr(temp, "/") = 0 And InStr(input, "月") > 0 Then
        temp = Year(Date) & "/" & temp
    End If
    
    ' Retry parsing with slash-separated format
    If IsDate(temp) Then
        outDate = CDate(temp)
        ParseJapaneseDate = True
        Exit Function
    End If
    
ErrorHandler:
    ParseJapaneseDate = False
End Function
```

4. **Save the File**:
   - Save the workbook as an `.xlsm` file (`File` → `Save As` → `Excel Macro-Enabled Workbook`).
   - Enable macros when opening the file (click "Enable Content" if prompted).

## Usage
1. Open the `.xlsm` file in Excel.
2. Enter a date in any cell in Column A (e.g., A1, A2). Examples:
   - `2025/6/6`
   - `2025-06-06`
   - `令和7年6月6日`
   - `6月6日` (uses current year, e.g., 2025)
   - `20250606`
3. The corresponding weekday (e.g., `金曜日`) will automatically appear in Column B (e.g., B1, B2).
4. If an invalid date is entered (e.g., `abc`), Column B will show `無効な日付`.

## Supported Date Formats
- **Standard Formats**: `2025/6/6`, `2025-06-06`, `2025.6.6`
- **Japanese Formats**: `2025年6月6日`, `令和7年6月6日`, `6月6日`
- **Other Formats**: `20250606`, `6/6` (year assumed as current year)
- **Full-width Characters**: e.g., `２０２５年６月６日`
- **Edge Cases**: Handles spaces, year omission, and some ambiguous formats (dependent on system regional settings).

## Customization
- **Change Output Column**: Modify `cell.Offset(0, 1)` to another offset (e.g., `cell.Offset(0, 2)` for Column C).
- **English Weekdays**: Change `Format(parsedDate, "aaaa")` to `"dddd"` (e.g., "Friday") or `"ddd"` (e.g., "Fri").
- **Default Year**: Change `Year(Date)` in `ParseJapaneseDate` to a fixed year (e.g., `2025`) for year-omitted dates.
- **Target Sheet**: Ensure the code is in the correct sheet module (e.g., `Sheet1`) or update references.

## Notes
- **Regional Settings**: Works best with Japanese regional settings in Windows for proper handling of formats like `2025年6月6日`. Non-Japanese settings may interpret ambiguous formats (e.g., `6/6`) differently.
- **Sheet Name**: The code assumes the target sheet is `Sheet1`. Adjust the sheet module or references if using a different sheet.
- **Performance**: Handles multiple cell changes (e.g., copy-paste) efficiently with `Application.EnableEvents` to prevent infinite loops.
- **Limitations**: Extremely ambiguous formats (e.g., `6-6`) may depend on regional settings. Highly non-standard formats (e.g., `六月六`) are not supported but can be added with regex parsing.

## License
MIT License. Feel free to use, modify, and distribute this code.

## Contributing
Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

## Contact
For questions or issues, create an issue on this GitHub repository.
