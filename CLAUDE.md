# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This project contains VBA code for counting material numbers in Excel files. The main functionality automatically counts rows with material numbers and updates a display cell when data changes.

## Project Structure

- `test.xlsm` - Excel workbook with material data in Materialnummer column
- `MaterialnummerCounter.bas` - VBA module with counting functions
- `Sheet1_Code.bas` - Worksheet event handlers for automatic updates
- `README.md` - Project documentation
- `.gitignore` - Contains .venv

## VBA Architecture

The project uses two VBA components:
1. **Module code** (`MaterialnummerCounter.bas`) - Contains the counting logic
2. **Worksheet events** (`Sheet1_Code.bas`) - Triggers automatic updates when data changes

## Development Notes

- The Excel file has 59 rows of material data with numeric material numbers
- Counting logic checks for non-empty numeric values in column A (Materialnummer)
- Results are automatically updated in Sheet2 cell B1 when Sheet1 changes
- Event handlers must be placed in worksheet code module, not standard module