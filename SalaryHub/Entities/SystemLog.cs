namespace SalaryHub.Entities
{
    public class SystemLog
    {
        public int Id { get; set; }

        public string? ErrorMessage { get; set; }

        public string? ErrorInnerMessage { get; set; }

        public string? StackTrace { get; set; }

        public DateTime? CreatedDate { get; set; }
    }
}
