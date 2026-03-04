using PayrollReconciliation.Models;

namespace PayrollReconciliation.Services;

public class ReconciliationEngine(Logger logger)
{
    private const decimal Tolerance = 0.01m; // 1 paisa tolerance for floating point

    public List<ReconciliationResult> Reconcile(List<HRRecord> hrRecords, List<FinanceRecord> financeRecords)
    {
        var results = new List<ReconciliationResult>();
        var finDict = financeRecords.ToDictionary(f => f.EmployeeId, StringComparer.InvariantCultureIgnoreCase);
        var hrIds = new HashSet<string>(hrRecords.Select(h => h.EmployeeId), StringComparer.InvariantCultureIgnoreCase);

        logger.Info($"Matching {hrRecords.Count} HR records against {financeRecords.Count} Finance records...");

        // Process all HR records
        foreach (var hr in hrRecords)
        {
            if (!finDict.TryGetValue(hr.EmployeeId, out var fin))
            {
                logger.Warn($"HR-Only: {hr.EmployeeId} ({hr.EmployeeName}) — not found in Finance data");
                results.Add(new ReconciliationResult()
                {
                    EmployeeId = hr.EmployeeId,
                    EmployeeName = hr.EmployeeName,
                    Department = hr.Department,
                    Designation = hr.Designation,
                    PayMonth = hr.PayMonth,
                    HR_GrossSalary = hr.GrossSalary,
                    HR_PFDeduction = hr.PFDeduction,
                    HR_ProfTax = hr.ProfessionalTax,
                    HR_OtherDeductions = hr.OtherDeductions,
                    HR_NetPay = hr.NetPay,
                    Status = ReconciliationStatus.HROnly,
                    MismatchRemarks = "Employee found in HR but missing in Finance disbursement",
                    HRRemarks = hr.Remarks,
                });

                continue;
            }

            var result = CompareRecords(hr, fin);
            results.Add(result);

            if (result.Status == ReconciliationStatus.Mismatched)
                logger.Warn($"Mismatch: {hr.EmployeeId} ({hr.EmployeeName}) — {result.MismatchRemarks}");
            else
                logger.Info($"Matched : {hr.EmployeeId} ({hr.EmployeeName})");
        }

        // Finance-only records (in Finance but not in HR)
        foreach (var fin in financeRecords.Where(f => !hrIds.Contains(f.EmployeeId)))
        {
            logger.Warn($"Finance-Only: {fin.EmployeeId} ({fin.EmployeeName}) — not found in HR data");
            results.Add(new ReconciliationResult()
            {
                EmployeeId = fin.EmployeeId,
                EmployeeName = fin.EmployeeName,
                Department = fin.Department,
                Designation = fin.Designation,
                Fin_GrossSalary = fin.GrossSalary,
                Fin_PFDeduction = fin.PFDeduction,
                Fin_ProfTax = fin.ProfessionalTax,
                Fin_OtherDeductions = fin.OtherDeductions,
                Fin_NetPay = fin.NetPayDisbursed,
                DisbursementDate = fin.DisbursementDate,
                BankRefNo = fin.BankRefNo,
                Status = ReconciliationStatus.FinanceOnly,
                MismatchRemarks = "Employee found in Finance disbursement but missing in HR records",
            });
        }

        return [.. results.OrderBy(r => r.EmployeeId)];
    }

    private static ReconciliationResult CompareRecords(HRRecord hr, FinanceRecord fin)
    {
        var mismatches = new List<string>();

        if (Math.Abs(hr.GrossSalary - fin.GrossSalary) > Tolerance)
            mismatches.Add($"Gross Salary differs: HR=₹{hr.GrossSalary:N0} vs Fin=₹{fin.GrossSalary:N0} (diff ₹{fin.GrossSalary - hr.GrossSalary:N0})");

        if (Math.Abs(hr.PFDeduction - fin.PFDeduction) > Tolerance)
            mismatches.Add($"PF Deduction differs: HR=₹{hr.PFDeduction:N0} vs Fin=₹{fin.PFDeduction:N0} (diff ₹{fin.PFDeduction - hr.PFDeduction:N0})");

        if (Math.Abs(hr.ProfessionalTax - fin.ProfessionalTax) > Tolerance)
            mismatches.Add($"Prof Tax differs: HR=₹{hr.ProfessionalTax:N0} vs Fin=₹{fin.ProfessionalTax:N0} (diff ₹{fin.ProfessionalTax - hr.ProfessionalTax:N0})");

        if (Math.Abs(hr.OtherDeductions - fin.OtherDeductions) > Tolerance)
            mismatches.Add($"Other Deductions differ: HR=₹{hr.OtherDeductions:N0} vs Fin=₹{fin.OtherDeductions:N0} (diff ₹{fin.OtherDeductions - hr.OtherDeductions:N0})");

        if (Math.Abs(hr.NetPay - fin.NetPayDisbursed) > Tolerance)
            mismatches.Add($"Net Pay differs by ₹{Math.Abs(fin.NetPayDisbursed - hr.NetPay):N0}: HR=₹{hr.NetPay:N0} vs Fin=₹{fin.NetPayDisbursed:N0}");

        return new ReconciliationResult()
        {
            EmployeeId = hr.EmployeeId,
            EmployeeName = hr.EmployeeName,
            Department = hr.Department,
            Designation = hr.Designation,
            PayMonth = hr.PayMonth,
            HR_GrossSalary = hr.GrossSalary,
            HR_PFDeduction = hr.PFDeduction,
            HR_ProfTax = hr.ProfessionalTax,
            HR_OtherDeductions = hr.OtherDeductions,
            HR_NetPay = hr.NetPay,
            Fin_GrossSalary = fin.GrossSalary,
            Fin_PFDeduction = fin.PFDeduction,
            Fin_ProfTax = fin.ProfessionalTax,
            Fin_OtherDeductions = fin.OtherDeductions,
            Fin_NetPay = fin.NetPayDisbursed,
            DisbursementDate = fin.DisbursementDate,
            BankRefNo = fin.BankRefNo,
            Status = mismatches.Count > 0 ? ReconciliationStatus.Mismatched : ReconciliationStatus.Matched,
            MismatchRemarks = string.Join("; ", mismatches),
            HRRemarks = hr.Remarks,
        };
    }
}
