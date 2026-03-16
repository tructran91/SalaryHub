namespace SalaryHub.Entities
{
    public class User
    {
        public int Id { get; set; }

        public string EmployeeCode { get; set; }

        public string FullName { get; set; }

        public string Department { get; set; }

        public DateTime? CreatedDate { get; set; }

        public ICollection<SalaryRecord> SalaryRecords { get; set; }
    }
}
