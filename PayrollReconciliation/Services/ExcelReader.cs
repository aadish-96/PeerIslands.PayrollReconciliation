using ClosedXML.Excel;
using PayrollReconciliation.Models;

namespace PayrollReconciliation.Services;

public class ExcelReader(Logger logger)
{
    public List<HRRecord> ReadHRData(string filePath, string sheetName)
    {
        logger.Info($"Opening HR file: {filePath}, sheet: {sheetName}");
        var records = new List<HRRecord>();

        using var wb = new XLWorkbook(filePath);

        if (!wb.TryGetWorksheet(sheetName, out var ws))
            throw new InvalidOperationException(
                $"Sheet '{sheetName}' not found. Available: {string.Join(", ", wb.Worksheets.Select(s => s.Name))}");

        int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;
        logger.Info($"HR sheet has {rowCount} rows (including header)");

        var colMap = BuildColumnMap(ws);
        logger.Info($"HR columns detected: {string.Join(", ", colMap.Keys)}");

        int skipped = 0;

        for (int row = 2; row <= rowCount; row++)
        {
            string empId = GetCellTextAsString(ws, row, colMap, ColumnNames.EMPLOYEE_ID).Trim();

            if (string.IsNullOrWhiteSpace(empId))
            {
                skipped++;
                continue;
            }

            try
            {
                records.Add(new HRRecord()
                {
                    EmployeeId = empId,
                    EmployeeName = GetCellTextAsString(ws, row, colMap, ColumnNames.EMPLOYEE_NAME),
                    Designation = GetCellTextAsString(ws, row, colMap, ColumnNames.DESIGNATION),
                    Department = GetCellTextAsString(ws, row, colMap, ColumnNames.DEPARTMENT),
                    GrossSalary = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.GROSS_SALARY),
                    PFDeduction = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.PF_DEDUCTION),
                    ProfessionalTax = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.PROFESSIONAL_TAX),
                    OtherDeductions = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.OTHER_DEDUCTIONS),
                    NetPay = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.NET_PAY),
                    PayMonth = GetCellTextAsString(ws, row, colMap, ColumnNames.PAY_MONTH),
                    Remarks = GetCellTextAsString(ws, row, colMap, ColumnNames.REMARKS),
                });
            }
            catch (Exception ex)
            {
                logger.Warn($"HR row {row} skipped: {ex.Message}");
                skipped++;
            }
        }

        if (skipped > 0)
            logger.Warn($"Skipped {skipped} HR rows");

        return records;
    }

    public List<FinanceRecord> ReadFinanceData(string filePath, string sheetName)
    {
        logger.Info($"Opening Finance file: {filePath}, sheet: {sheetName}");
        var records = new List<FinanceRecord>();

        using var wb = new XLWorkbook(filePath);

        if (!wb.TryGetWorksheet(sheetName, out var ws))
            throw new InvalidOperationException(
                $"Sheet '{sheetName}' not found. Available: {string.Join(", ", wb.Worksheets.Select(s => s.Name))}");

        int rowCount = ws.LastRowUsed()?.RowNumber() ?? 0;
        logger.Info($"Finance sheet has {rowCount} rows (including header)");

        var colMap = BuildColumnMap(ws);
        logger.Info($"Finance columns detected: {string.Join(", ", colMap.Keys)}");

        int skipped = 0;

        for (int row = 2; row <= rowCount; row++)
        {
            string empId = GetCellTextAsString(ws, row, colMap, ColumnNames.EMPLOYEE_ID);
            if (string.IsNullOrWhiteSpace(empId)) { skipped++; continue; }

            try
            {
                records.Add(new FinanceRecord()
                {
                    EmployeeId = empId.Trim(),
                    EmployeeName = GetCellTextAsString(ws, row, colMap, ColumnNames.EMPLOYEE_NAME),
                    Designation = GetCellTextAsString(ws, row, colMap, ColumnNames.DESIGNATION),
                    Department = GetCellTextAsString(ws, row, colMap, ColumnNames.DEPARTMENT),
                    GrossSalary = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.GROSS_SALARY),
                    PFDeduction = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.PF_DEDUCTION),
                    ProfessionalTax = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.PROFESSIONAL_TAX),
                    OtherDeductions = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.OTHER_DEDUCTIONS),
                    NetPayDisbursed = GetCellValueAsDecimal(ws, row, colMap, ColumnNames.NET_PAY),
                    DisbursementDate = GetCellTextAsString(ws, row, colMap, ColumnNames.DISBURSEMENT_DATE),
                    BankRefNo = GetCellTextAsString(ws, row, colMap, ColumnNames.BANK_REF_NO),
                    Remarks = GetCellTextAsString(ws, row, colMap, ColumnNames.REMARKS),
                });
            }
            catch (Exception ex)
            {
                logger.Warn($"Finance row {row} skipped: {ex.Message}");
                skipped++;
            }
        }

        if (skipped > 0)
            logger.Warn($"Skipped {skipped} Finance rows");

        return records;
    }

    // Build header→column-number map using partial matching
    private static Dictionary<string, int> BuildColumnMap(IXLWorksheet ws)
    {
        var map = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);
        var headerRow = ws.Row(1);

        foreach (var cell in headerRow.CellsUsed())
        {
            var text = cell.GetString().Trim();

            if (!string.IsNullOrEmpty(text))
                map[text] = cell.Address.ColumnNumber;
        }

        return map;
    }

    private static string GetCellTextAsString(IXLWorksheet ws, int row, Dictionary<string, int> colMap, string keyFragment)
    {
        int col = FindColumn(colMap, keyFragment);
        return col == -1 ? string.Empty : ws.Cell(row, col).GetString().Trim();
    }

    private static decimal GetCellValueAsDecimal(IXLWorksheet ws, int row, Dictionary<string, int> colMap, string keyFragment)
    {
        int col = FindColumn(colMap, keyFragment);

        if (col == -1)
            return 0;

        var cell = ws.Cell(row, col);
        return cell.TryGetValue(out decimal d) ? d : 0;
    }

    private static int FindColumn(Dictionary<string, int> colMap, string fragment)
    {
        if (colMap.TryGetValue(fragment, out int idx))
            return idx;

        foreach (var kv in colMap)
            if (kv.Key.Contains(fragment, StringComparison.InvariantCultureIgnoreCase))
                return kv.Value;

        return -1;
    }
}
