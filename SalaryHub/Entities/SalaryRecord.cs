namespace SalaryHub.Entities
{
    public class SalaryRecord
    {
        public int Id { get; set; }

        public int UserId { get; set; }

        public int Month { get; set; }

        public int Year { get; set; }

        public string Title { get; set; }

        public decimal TaxableIncome { get; set; } // Tổng thu nhập chịu thuế

        public decimal? Pit { get; set; } // Trích thuế TNCN

        public decimal? Bhxh { get; set; } // Trích BHXH

        public DateTime? CreatedDate { get; set; }

        public User User { get; set; }
    }
}
