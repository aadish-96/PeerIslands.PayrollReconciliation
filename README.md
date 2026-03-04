# Payroll Reconciliation Tool — C# Console Application

## Overview
Compares HR salary data vs Finance disbursement data from Excel files,
flags discrepancies, and produces a colour-coded reconciliation report.

## Project Structure
```
PayrollReconciliation/
├── PayrollReconciliation.csproj
├── Program.cs                        ← Entry point & CLI
├── Models/
│   ├── Constants.cs                  ← Constants such as Column Names, Color Hex Codes, etc.
│   ├── HRRecord.cs                   ← HR data model
│   ├── FinanceRecord.cs              ← Finance data model
│   └── ReconciliationResult.cs       ← Result + status enum
└── Services/
    ├── Logger.cs                     ← Timestamped file + console logger
    ├── ExcelReader.cs                ← Reads HR & Finance sheets (ClosedXML)
    ├── ReconciliationEngine.cs       ← Comparison logic (field-by-field)
    └── ReportWriter.cs               ← Writes .xlsx or .csv output (ClosedXML)
```

## Prerequisites
- .NET 8 SDK
- Run `dotnet restore` to pull NuGet packages (ClosedXML, CsvHelper)

## Build & Run
```bash
# Build
dotnet build

# Run — same workbook (two sheets)
dotnet run -- payroll.xlsx payroll.xlsx

# Run — separate files, explicit output and log paths
dotnet run -- hr.xlsx finance.xlsx output/report.xlsx logs/recon.log

# Override sheet names
dotnet run -- payroll.xlsx payroll.xlsx report.xlsx run.log "HR_Salary_Data" "Finance_Disbursement"
```

## CLI Arguments
| Pos | Argument        | Required | Default                          |
|-----|-----------------|----------|----------------------------------|
| 1   | hr-file         | Yes      | —                                |
| 2   | finance-file    | Yes      | —                                |
| 3   | output-file     | No       | reconciliation_output_<ts>.xlsx  |
| 4   | log-file        | No       | reconciliation_<ts>.log          |
| 5   | hr-sheet        | No       | HR_Salary_Data                   |
| 6   | finance-sheet   | No       | Finance_Disbursement             |

## Output Report (3-tab Excel)
| Sheet               | Contents                                          |
|---------------------|---------------------------------------------------|
| Reconciliation Report | All records colour-coded by status              |
| Issues Only           | Only mismatched / missing records               |
| Summary               | Counts per category with percentages            |

## Status Codes
| Status         | Meaning                                        | Colour  |
|----------------|------------------------------------------------|---------|
| Matched        | All fields agree                               | Green   |
| Mismatched     | One or more field differences                  | Red     |
| HR_Only        | In HR but no Finance disbursement found        | Yellow  |
| Finance_Only   | Disbursed but not found in HR records          | Orange  |

## Fields Compared
- Gross Salary
- PF Deduction
- Professional Tax
- Other Deductions
- Net Pay

## Exit Codes
| Code | Meaning                      |
|------|------------------------------|
| 0    | Success                      |
| 1    | Usage / no args              |
| 2    | Input file not found         |
| 99   | Unexpected runtime exception |
