namespace SalaryHub.Models
{
    public class SalaryReportViewModel
    {
        public List<string> IncomeTitles { get; set; } = new();

        public List<string> PitTitles { get; set; } = new();

        public string? DuplicateName { get; set; }

        /// <summary>
        /// Mapping title -> màu CSS hex (vd: "#D9EAD3")
        /// </summary>
        public Dictionary<string, string> TitleColorMap { get; set; } = new();

        public List<UserSalaryRow> Rows { get; set; } = new();
    }

    public class UserSalaryRow
    {
        public string EmployeeCode { get; set; }
        public string FullName { get; set; }
        public string Department { get; set; }

        public Dictionary<string, decimal> Incomes { get; set; } = new();
        public Dictionary<string, decimal> Pits { get; set; } = new();

        public decimal Bhxh { get; set; }
    }
}
