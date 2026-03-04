namespace PayrollReconciliation.Models
{
    public class ColumnNames
    {
        public const string EMPLOYEE_ID = "Employee ID";
        public const string EMPLOYEE_NAME = "Employee Name";
        public const string DESIGNATION = "Designation";
        public const string DEPARTMENT = "Department";
        public const string STATUS = "Status";
        public const string REMARKS = "Remarks";
        public const string GROSS_SALARY = "Gross Salary";
        public const string PF_DEDUCTION = "PF Deduction";
        public const string PROFESSIONAL_TAX = "Professional Tax";
        public const string OTHER_DEDUCTIONS = "Other Deductions";
        public const string NET_PAY = "Net Pay";
        public const string HR_NET_PAY = "HR Net Pay (₹)";
        public const string FIN_NET_PAY = "Fin Net Pay (₹)";
        public const string DIFFERENCE = "Difference (₹)";
        public const string MISMATCH_REMARKS = "Mismatch Remarks";
        public const string HR_NOTES = "HR Notes";
        public const string PAY_MONTH = "Pay Month";
        public const string HR_GROSS = "HR Gross (₹)";
        public const string FIN_GROSS = "Fin Gross (₹)";
        public const string HR_PF_DED = "HR PF Ded (₹)";
        public const string FIN_PF_DED = "Fin PF Ded (₹)";
        public const string HR_PROF_TAX = "HR Prof Tax (₹)";
        public const string FIN_PROF_TAX = "Fin Prof Tax (₹)";
        public const string HR_OTHER_DED = "HR Other Ded (₹)";
        public const string FIN_OTHER_DED = "Fin Other Ded (₹)";
        public const string DISBURSEMENT_DATE = "Disbursement Date";
        public const string BANK_REF_NO = "Bank Ref No.";
    }

    public class HexColorCodes
    {
        // Hex colours (no # prefix for ClosedXML)
        public const string HEADER = "1F497D";
        public const string MATCHED = "C6EFCE";
        public const string MISMATCH = "FFC7CE";
        public const string HR_ONLY = "FFFF00";
        public const string FIN_ONLY = "FA9D05";
        public const string WHITE = "FFFFFF";
    }
}
