using PayrollReconciliation.Models;
using PayrollReconciliation.Services;

class Program
{
    static int Main(string[] args)
    {
        Console.WriteLine("=== Automated Payroll Data Reconciliation Tool ===");
        Console.WriteLine($"Started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        Console.WriteLine();

        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        string hrFilePath = args[0];
        string finFilePath = args[1];
        string outputPath = args.Length >= 3 ? args[2] : GenerateOutputPath();
        string logPath = args.Length >= 4 ? args[3] : GenerateLogPath();
        string hrSheet = args.Length >= 5 ? args[4] : "HR_Salary_Data";
        string financeSheet = args.Length >= 6 ? args[5] : "Finance_Disbursement";

        var logger = new Logger(logPath);

        try
        {
            logger.Info("Payroll Reconciliation Tool started");
            logger.Info($"HR File      : {hrFilePath}");
            logger.Info($"Finance File : {finFilePath}");
            logger.Info($"Output File  : {outputPath}");
            logger.Info($"HR Sheet     : {hrSheet}");
            logger.Info($"Finance Sheet: {financeSheet}");

            if (!File.Exists(hrFilePath))
            {
                logger.Error($"HR file not found: {hrFilePath}");
                Console.Error.WriteLine($"ERROR: HR file not found: {hrFilePath}");
                return 2;
            }
            if (!File.Exists(finFilePath))
            {
                logger.Error($"Finance file not found: {finFilePath}");
                Console.Error.WriteLine($"ERROR: Finance file not found: {finFilePath}");
                return 2;
            }

            logger.Info("Reading HR data...");
            var excelReader = new ExcelReader(logger);
            var hrRecords = excelReader.ReadHRData(hrFilePath, hrSheet);
            logger.Info($"Loaded {hrRecords.Count} HR records");

            logger.Info("Reading Finance data...");
            var financeRecords = excelReader.ReadFinanceData(finFilePath, financeSheet);
            logger.Info($"Loaded {financeRecords.Count} Finance records");

            logger.Info("Running reconciliation...");
            var engine = new ReconciliationEngine(logger);
            var results = engine.Reconcile(hrRecords, financeRecords);

            int matched = results.Count(r => r.Status == ReconciliationStatus.Matched);
            int mismatched = results.Count(r => r.Status == ReconciliationStatus.Mismatched);
            int hrOnly = results.Count(r => r.Status == ReconciliationStatus.HROnly);
            int finOnly = results.Count(r => r.Status == ReconciliationStatus.FinanceOnly);

            logger.Info($"Reconciliation complete — Total: {results.Count} | Matched: {matched} | Mismatched: {mismatched} | HR-Only: {hrOnly} | Finance-Only: {finOnly}");

            logger.Info("Writing output report...");
            var reporter = new ReportWriter(logger);
            reporter.WriteReport(results, outputPath);
            logger.Info($"Report written to: {outputPath}");

            PrintSummary(results, outputPath, logPath);

            logger.Info("Payroll Reconciliation Tool completed successfully");
            return 0;
        }
        catch (Exception ex)
        {
            logger.Error($"Fatal error: {ex.Message}");
            logger.Error(ex.StackTrace ?? string.Empty);
            Console.Error.WriteLine($"FATAL ERROR: {ex.Message}");
            return 99;
        }
    }

    static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  PayrollReconciliation <hr-file> <finance-file> [output-file] [log-file] [hr-sheet] [finance-sheet]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  hr-file        Path to HR Excel file (.xlsx)");
        Console.WriteLine("  finance-file   Path to Finance Excel file (.xlsx) — can be the same file");
        Console.WriteLine("  output-file    Output reconciliation report (.xlsx or .csv) [optional]");
        Console.WriteLine("  log-file       Path for the log file [optional]");
        Console.WriteLine("  hr-sheet       Sheet name for HR data [default: HR_Salary_Data]");
        Console.WriteLine("  finance-sheet  Sheet name for Finance data [default: Finance_Disbursement]");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  PayrollReconciliation payroll.xlsx payroll.xlsx");
        Console.WriteLine("  PayrollReconciliation hr.xlsx finance.xlsx output.xlsx recon.log HR Finance");
    }

    static string GenerateOutputPath() => $"output/reconciliation_output_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

    static string GenerateLogPath() => $"logs/reconciliation_{DateTime.Now:yyyyMMdd_HHmmss}.log";

    static void PrintSummary(List<ReconciliationResult> results, string outputPath, string logPath)
    {
        int total = results.Count;
        int matched = results.Count(r => r.Status == ReconciliationStatus.Matched);
        int mismatched = results.Count(r => r.Status == ReconciliationStatus.Mismatched);
        int hrOnly = results.Count(r => r.Status == ReconciliationStatus.HROnly);
        int finOnly = results.Count(r => r.Status == ReconciliationStatus.FinanceOnly);

        Console.WriteLine("╔══════════════════════════════════════════════╗");
        Console.WriteLine("║         RECONCILIATION SUMMARY               ║");
        Console.WriteLine("╠══════════════════════════════════════════════╣");
        Console.WriteLine($"║  Total Records Processed : {total,-17} ║");
        Console.WriteLine($"║  Matched                 : {matched,-17} ║");
        Console.WriteLine($"║  Mismatched              : {mismatched,-17} ║");
        Console.WriteLine($"║  HR Only (missing in Fin): {hrOnly,-17} ║");
        Console.WriteLine($"║  Finance Only            : {finOnly,-17} ║");
        Console.WriteLine("╚══════════════════════════════════════════════╝");
        Console.WriteLine($"Output: {outputPath,-36}");
        Console.WriteLine($"Log   : {logPath,-36}");

        if (mismatched > 0)
        {
            Console.WriteLine();
            Console.WriteLine("MISMATCH DETAILS:");

            foreach (var r in results.Where(x => x.Status == ReconciliationStatus.Mismatched))
                Console.WriteLine($"  [{r.EmployeeId}] {r.EmployeeName}: {r.MismatchRemarks}");
        }
    }
}
