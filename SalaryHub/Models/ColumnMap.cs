namespace SalaryHub.Models
{
    public class ColumnMap
    {
        public int No { get; set; }

        public int EmpCode { get; set; }

        public int Name { get; set; }

        public int Department { get; set; }

        public int TaxableIncome { get; set; } // Tổng thu nhập chịu thuế

        public int Pit { get; set; } = -1; // Trích thuế TNCN

        public int Bhxh { get; set; } = -1; // Trích BHXH
    }
}
