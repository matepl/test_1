# Excel Material Number Counter

This project contains VBA code to count rows with material numbers in an Excel file and automatically update a counter when the data changes.

## Files

- `test.xlsm` - Excel file with material data (Materialnummer column)
- `MaterialnummerCounter.bas` - Main VBA module with counting functions
- `Sheet1_Code.bas` - Worksheet event handlers for automatic updates

## VBA Components

### MaterialnummerCounter.bas (Module)
- `UpdateMaterialnummerCount()` - Updates count in Sheet2 cell B1
- `CountMaterialnummer()` - Displays count in message box
- `CountMaterialnummerRows()` - Returns count as function

### Sheet1_Code.bas (Worksheet Events)
- `Worksheet_Change()` - Triggers on cell changes
- `Worksheet_Calculate()` - Triggers on recalculation

## Setup Instructions

1. Copy `MaterialnummerCounter.bas` code to a new VBA module
2. Copy `Sheet1_Code.bas` code to Sheet1's worksheet code module
3. The count will automatically update in Sheet2 cell B1 when Sheet1 data changes