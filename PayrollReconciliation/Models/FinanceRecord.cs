namespace PayrollReconciliation.Models;

public class FinanceRecord
{
    public string EmployeeId { get; set; } = string.Empty;
    public string EmployeeName { get; set; } = string.Empty;
    public string Designation { get; set; } = string.Empty;
    public string Department { get; set; } = string.Empty;
    public decimal GrossSalary { get; set; }
    public decimal PFDeduction { get; set; }
    public decimal ProfessionalTax { get; set; }
    public decimal OtherDeductions { get; set; }
    public decimal NetPayDisbursed { get; set; }
    public string DisbursementDate { get; set; } = string.Empty;
    public string BankRefNo { get; set; } = string.Empty;
    public string Remarks { get; set; } = string.Empty;
}
