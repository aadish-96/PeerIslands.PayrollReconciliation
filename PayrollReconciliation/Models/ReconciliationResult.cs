namespace PayrollReconciliation.Models;

public enum ReconciliationStatus
{
    Matched,
    Mismatched,
    HROnly,
    FinanceOnly
}

public class ReconciliationResult
{
    public string EmployeeId { get; set; } = string.Empty;
    public string EmployeeName { get; set; } = string.Empty;
    public string Department { get; set; } = string.Empty;
    public string Designation { get; set; } = string.Empty;
    public string PayMonth { get; set; } = string.Empty;

    // HR values
    public decimal? HR_GrossSalary { get; set; }
    public decimal? HR_PFDeduction { get; set; }
    public decimal? HR_ProfTax { get; set; }
    public decimal? HR_OtherDeductions { get; set; }
    public decimal? HR_NetPay { get; set; }

    // Finance values
    public decimal? Fin_GrossSalary { get; set; }
    public decimal? Fin_PFDeduction { get; set; }
    public decimal? Fin_ProfTax { get; set; }
    public decimal? Fin_OtherDeductions { get; set; }
    public decimal? Fin_NetPay { get; set; }
    public string DisbursementDate { get; set; } = string.Empty;
    public string BankRefNo { get; set; } = string.Empty;

    public ReconciliationStatus Status { get; set; }
    public string MismatchRemarks { get; set; } = string.Empty;
    public string HRRemarks { get; set; } = string.Empty;
}
