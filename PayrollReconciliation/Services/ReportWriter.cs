using ClosedXML.Excel;
using PayrollReconciliation.Models;
using System.Text;

namespace PayrollReconciliation.Services;

public class ReportWriter(Logger logger)
{
    public void WriteReport(List<ReconciliationResult> results, string outputPath)
    {
        if (outputPath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            WriteCsv(results, outputPath);
        else
            WriteExcel(results, outputPath);
    }

    // ── Excel ─────────────────────────────────────────────────────────────────
    private void WriteExcel(List<ReconciliationResult> results, string outputPath)
    {
        logger.Info($"Writing Excel report to {outputPath}...");

        var outDir = Path.GetDirectoryName(outputPath);

        if (!string.IsNullOrEmpty(outDir))
            Directory.CreateDirectory(outDir);

        using var wb = new XLWorkbook();

        WriteMainSheet(wb, results);
        WriteIssuesSheet(wb, results);
        WriteSummarySheet(wb, results);

        wb.SaveAs(outputPath);
        logger.Info("Excel report saved.");
    }

    private static void WriteMainSheet(XLWorkbook wb, List<ReconciliationResult> results)
    {
        var ws = wb.Worksheets.Add("Reconciliation Report");

        // Title
        ws.Cell("A1").Value = "PAYROLL RECONCILIATION REPORT";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml($"#{HexColorCodes.HEADER}");

        int matched = results.Count(r => r.Status == ReconciliationStatus.Matched);
        int mismatch = results.Count(r => r.Status == ReconciliationStatus.Mismatched);
        int hrOnly = results.Count(r => r.Status == ReconciliationStatus.HROnly);
        int finOnly = results.Count(r => r.Status == ReconciliationStatus.FinanceOnly);
        ws.Cell("A2").Value = $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}  |  Total: {results.Count}  |  Matched: {matched}  |  Mismatched: {mismatch}  |  HR-Only: {hrOnly}  |  Finance-Only: {finOnly}";
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Cell("A2").Style.Font.FontSize = 9;

        // Headers — row 4
        var headers = new[]
        {
            ColumnNames.EMPLOYEE_ID, ColumnNames.EMPLOYEE_NAME, ColumnNames.DEPARTMENT, ColumnNames.DESIGNATION, ColumnNames.PAY_MONTH, ColumnNames.HR_GROSS, ColumnNames.FIN_GROSS, ColumnNames.HR_PF_DED, ColumnNames.FIN_PF_DED, ColumnNames.HR_PROF_TAX, ColumnNames.FIN_PROF_TAX, ColumnNames.HR_OTHER_DED, ColumnNames.FIN_OTHER_DED, ColumnNames.HR_NET_PAY, ColumnNames.FIN_NET_PAY, ColumnNames.DIFFERENCE, ColumnNames.DISBURSEMENT_DATE, ColumnNames.BANK_REF_NO, ColumnNames.STATUS, ColumnNames.MISMATCH_REMARKS, ColumnNames.HR_NOTES
        };

        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(4, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml($"#{HexColorCodes.HEADER}");
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
        }

        ws.Row(4).Height = 30;

        // Status → background colour
        XLColor RowColor(ReconciliationStatus s) => s switch
        {
            ReconciliationStatus.Matched => XLColor.FromHtml($"#{HexColorCodes.MATCHED}"),
            ReconciliationStatus.Mismatched => XLColor.FromHtml($"#{HexColorCodes.MISMATCH}"),
            ReconciliationStatus.HROnly => XLColor.FromHtml($"#{HexColorCodes.HR_ONLY}"),
            ReconciliationStatus.FinanceOnly => XLColor.FromHtml($"#{HexColorCodes.FIN_ONLY}"),
            _ => XLColor.White
        };

        string StatusLabel(ReconciliationStatus s) => s switch
        {
            ReconciliationStatus.Matched => "Matched",
            ReconciliationStatus.Mismatched => "Mismatched",
            ReconciliationStatus.HROnly => "HR Only — Not Disbursed",
            ReconciliationStatus.FinanceOnly => "Finance Only — Not in HR",
            _ => s.ToString()
        };

        int row = 5;

        foreach (var r in results)
        {
            var rowRange = ws.Row(row);
            ws.Cell(row, 1).Value = r.EmployeeId;
            ws.Cell(row, 2).Value = r.EmployeeName;
            ws.Cell(row, 3).Value = r.Department;
            ws.Cell(row, 4).Value = r.Designation;
            ws.Cell(row, 5).Value = r.PayMonth;
            SetNum(ws, row, 6, r.HR_GrossSalary);
            SetNum(ws, row, 7, r.Fin_GrossSalary);
            SetNum(ws, row, 8, r.HR_PFDeduction);
            SetNum(ws, row, 9, r.Fin_PFDeduction);
            SetNum(ws, row, 10, r.HR_ProfTax);
            SetNum(ws, row, 11, r.Fin_ProfTax);
            SetNum(ws, row, 12, r.HR_OtherDeductions);
            SetNum(ws, row, 13, r.Fin_OtherDeductions);
            SetNum(ws, row, 14, r.HR_NetPay);
            SetNum(ws, row, 15, r.Fin_NetPay);
            var diff = (r.Fin_NetPay ?? 0) - (r.HR_NetPay ?? 0);

            if (diff != 0)
                SetNum(ws, row, 16, diff); else ws.Cell(row, 16).Value = "";

            ws.Cell(row, 17).Value = r.DisbursementDate;
            ws.Cell(row, 18).Value = r.BankRefNo;
            ws.Cell(row, 19).Value = StatusLabel(r.Status);
            ws.Cell(row, 19).Style.Font.Bold = true;
            ws.Cell(row, 20).Value = r.MismatchRemarks;
            ws.Cell(row, 20).Style.Alignment.WrapText = true;
            ws.Cell(row, 21).Value = r.HRRemarks;
            ws.Cell(row, 21).Style.Alignment.WrapText = true;

            // Colour entire row
            var bg = RowColor(r.Status);
            ws.Range(row, 1, row, headers.Length).Style.Fill.BackgroundColor = bg;
            ws.Range(row, 1, row, headers.Length).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(row, 1, row, headers.Length).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range(row, 1, row, headers.Length).Style.Border.OutsideBorderColor = XLColor.Gray;
            ws.Range(row, 1, row, headers.Length).Style.Border.InsideBorderColor = XLColor.Gray;

            row++;
        }

        //// Column widths
        //int[] widths = [12, 22, 16, 22, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 16, 18, 26, 60, 35];

        //for (int i = 0; i < widths.Length; i++)
        //    ws.Column(i + 1).Width = widths[i];

        ws.SheetView.FreezeRows(4);

        // Legend
        int legRow = row + 2;
        ws.Cell(legRow, 1).Value = "Colour Legend:";
        ws.Cell(legRow, 1).Style.Font.Bold = true;
        AddLegend(ws, legRow + 1, HexColorCodes.MATCHED, "Matched — all fields agree");
        AddLegend(ws, legRow + 2, HexColorCodes.MISMATCH, "Mismatched — one or more field differences");
        AddLegend(ws, legRow + 3, HexColorCodes.HR_ONLY, "HR Only — employee not found in Finance disbursement");
        AddLegend(ws, legRow + 4, HexColorCodes.FIN_ONLY, "Finance Only — disbursed but not in HR records");


        ws.Rows().AdjustToContents();
        ws.Columns().AdjustToContents();
    }

    private static void AddLegend(IXLWorksheet ws, int row, string hexColor, string label)
    {
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml($"#{hexColor}");
        ws.Cell(row, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Cell(row, 2).Value = label;
        ws.Cell(row, 2).Style.Font.FontSize = 9;
    }

    private static void SetNum(IXLWorksheet ws, int row, int col, decimal? val)
    {
        if (val.HasValue)
        {
            ws.Cell(row, col).Value = val.Value;
            ws.Cell(row, col).Style.NumberFormat.Format = "#,##0";
        }
    }

    // ── Issues-Only Sheet ─────────────────────────────────────────────────────
    private static void WriteIssuesSheet(XLWorkbook wb, List<ReconciliationResult> results)
    {
        var issues = results.Where(r => r.Status != ReconciliationStatus.Matched).ToList();
        var ws = wb.Worksheets.Add("Issues Only");

        ws.Cell("A1").Value = $"ISSUES REQUIRING ATTENTION — {issues.Count} records";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 13;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml("#C00000");

        var hdrs = new[] { ColumnNames.EMPLOYEE_ID, ColumnNames.EMPLOYEE_NAME, ColumnNames.DEPARTMENT, ColumnNames.STATUS, ColumnNames.HR_NET_PAY, ColumnNames.FIN_NET_PAY, ColumnNames.DIFFERENCE, ColumnNames.MISMATCH_REMARKS, ColumnNames.HR_NOTES };

        for (int i = 0; i < hdrs.Length; i++)
        {
            var c = ws.Cell(3, i + 1);
            c.Value = hdrs[i];
            c.Style.Font.Bold = true;
            c.Style.Font.FontColor = XLColor.White;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#C00000");
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = 4;

        foreach (var r in issues)
        {
            XLColor bg = r.Status switch
            {
                ReconciliationStatus.Mismatched => XLColor.FromHtml($"#{HexColorCodes.MISMATCH}"),
                ReconciliationStatus.HROnly => XLColor.FromHtml($"#{HexColorCodes.HR_ONLY}"),
                ReconciliationStatus.FinanceOnly => XLColor.FromHtml($"#{HexColorCodes.FIN_ONLY}"),
                _ => XLColor.White
            };

            decimal diff = (r.Fin_NetPay ?? 0) - (r.HR_NetPay ?? 0);
            ws.Cell(row, 1).Value = r.EmployeeId;
            ws.Cell(row, 2).Value = r.EmployeeName;
            ws.Cell(row, 3).Value = r.Department;
            ws.Cell(row, 4).Value = r.Status.ToString();
            ws.Cell(row, 4).Style.Font.Bold = true;
            SetNum(ws, row, 5, r.HR_NetPay);
            SetNum(ws, row, 6, r.Fin_NetPay);
            if (diff != 0) SetNum(ws, row, 7, diff);
            ws.Cell(row, 8).Value = r.MismatchRemarks;
            ws.Cell(row, 8).Style.Alignment.WrapText = true;
            ws.Cell(row, 9).Value = r.HRRemarks;
            ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = bg;
            ws.Range(row, 1, row, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(row, 1, row, 9).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            row++;
        }

        //int[] w = [12, 22, 16, 26, 16, 16, 16, 65, 35];

        //for (int i = 0; i < w.Length; i++)
        //    ws.Column(i + 1).Width = w[i];

        ws.SheetView.FreezeRows(3);

        ws.Rows().AdjustToContents();
        ws.Columns().AdjustToContents();
    }

    // ── Summary Sheet ─────────────────────────────────────────────────────────
    private static void WriteSummarySheet(XLWorkbook wb, List<ReconciliationResult> results)
    {
        var ws = wb.Worksheets.Add("Summary");
        ws.Cell("A1").Value = "RECONCILIATION SUMMARY";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml($"#{HexColorCodes.HEADER}");
        ws.Cell("A2").Value = $"Pay Period: {results.FirstOrDefault()?.PayMonth ?? "—"}  |  Run Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        ws.Cell("A2").Style.Font.Italic = true;

        int total = results.Count;
        var rows = new[]
        {
            ("Category", "Count", "% of Total", HexColorCodes.HEADER),
            ("Total Records Processed", total.ToString(), "100.0%", HexColorCodes.WHITE),
            ("Matched (No Issues)", results.Count(r => r.Status == ReconciliationStatus.Matched).ToString(), $"{results.Count(r => r.Status == ReconciliationStatus.Matched) * 100.0 / total:F1}%", HexColorCodes.MATCHED),
            ("Mismatched (Field Differences)",results.Count(r => r.Status == ReconciliationStatus.Mismatched).ToString(), $"{results.Count(r => r.Status == ReconciliationStatus.Mismatched) * 100.0 / total:F1}%", HexColorCodes.MISMATCH),
            ("HR Only (Not Disbursed)", results.Count(r => r.Status == ReconciliationStatus.HROnly).ToString(), $"{results.Count(r => r.Status == ReconciliationStatus.HROnly) * 100.0 / total:F1}%", HexColorCodes.HR_ONLY),
            ("Finance Only (Not in HR)", results.Count(r => r.Status == ReconciliationStatus.FinanceOnly).ToString(), $"{results.Count(r => r.Status == ReconciliationStatus.FinanceOnly) * 100.0 / total:F1}%", HexColorCodes.FIN_ONLY),
        };

        for (int i = 0; i < rows.Length; i++)
        {
            int rowNum = i + 4;
            var (label, count, pct, color) = rows[i];
            bool isHeader = i == 0;
            ws.Cell(rowNum, 1).Value = label;
            ws.Cell(rowNum, 2).Value = count;
            ws.Cell(rowNum, 3).Value = pct;
            ws.Cell(rowNum, 1).Style.Font.Bold = isHeader;
            ws.Cell(rowNum, 2).Style.Font.Bold = !isHeader;
            var bg = isHeader ? XLColor.FromHtml("#" + color) : XLColor.FromHtml("#" + color);
            ws.Range(rowNum, 1, rowNum, 3).Style.Fill.BackgroundColor = bg;
            if (isHeader) ws.Range(rowNum, 1, rowNum, 3).Style.Font.FontColor = XLColor.White;
            ws.Range(rowNum, 1, rowNum, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(rowNum, 1, rowNum, 3).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        //ws.Column(1).Width = 36;
        //ws.Column(2).Width = 12;
        //ws.Column(3).Width = 14;
        ws.Rows().AdjustToContents();
        ws.Columns().AdjustToContents();
    }

    // ── CSV fallback ──────────────────────────────────────────────────────────
    private void WriteCsv(List<ReconciliationResult> results, string outputPath)
    {
        logger.Info($"Writing CSV report to {outputPath}...");
        
        StringBuilder hdr = new();
        hdr.AppendJoin<string>(',', [
            ColumnNames.EMPLOYEE_ID,
            ColumnNames.EMPLOYEE_NAME,
            ColumnNames.DEPARTMENT,
            ColumnNames.DESIGNATION,
            ColumnNames.PAY_MONTH,
            ColumnNames.HR_GROSS,
            ColumnNames.FIN_GROSS,
            ColumnNames.HR_PF_DED,
            ColumnNames.FIN_PF_DED,
            ColumnNames.HR_PROF_TAX,
            ColumnNames.FIN_PROF_TAX,
            ColumnNames.HR_OTHER_DED,
            ColumnNames.FIN_OTHER_DED,
            ColumnNames.HR_NET_PAY,
            ColumnNames.FIN_NET_PAY,
            ColumnNames.DIFFERENCE,
            ColumnNames.DISBURSEMENT_DATE,
            ColumnNames.BANK_REF_NO,
            ColumnNames.STATUS,
            ColumnNames.MISMATCH_REMARKS,
            ColumnNames.HR_NOTES
            ]);

        var lines = new List<string>
        {
            hdr.ToString()
        };

        foreach (var r in results)
        {
            decimal diff = (r.Fin_NetPay ?? 0) - (r.HR_NetPay ?? 0);

            lines.Add(string.Join(",", Q(r.EmployeeId), Q(r.EmployeeName), Q(r.Department), Q(r.Designation), Q(r.PayMonth), r.HR_GrossSalary, r.Fin_GrossSalary, r.HR_PFDeduction, r.Fin_PFDeduction, r.HR_ProfTax, r.Fin_ProfTax, r.HR_OtherDeductions, r.Fin_OtherDeductions, r.HR_NetPay, r.Fin_NetPay, diff != 0 ? diff.ToString() : "", Q(r.DisbursementDate), Q(r.BankRefNo), r.Status, Q(r.MismatchRemarks), Q(r.HRRemarks)));
        }

        var dir = Path.GetDirectoryName(outputPath);

        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        
        File.WriteAllLines(outputPath, lines);
        logger.Info("CSV report saved.");
    }

    private static string Q(string? s) => $"\"{(s ?? string.Empty).Replace("\"", "\"\"")}\"";
}
