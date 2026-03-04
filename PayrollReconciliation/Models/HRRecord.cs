namespace PayrollReconciliation.Models;

public class HRRecord
{
    public string EmployeeId { get; set; } = string.Empty;
    public string EmployeeName { get; set; } = string.Empty;
    public string Designation { get; set; } = string.Empty;
    public string Department { get; set; } = string.Empty;
    public decimal GrossSalary { get; set; }
    public decimal PFDeduction { get; set; }
    public decimal ProfessionalTax { get; set; }
    public decimal OtherDeductions { get; set; }
    public decimal NetPay { get; set; }
    public string PayMonth { get; set; } = string.Empty;
    public string Remarks { get; set; } = string.Empty;
}
