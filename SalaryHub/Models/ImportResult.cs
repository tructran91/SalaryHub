namespace SalaryHub.Models
{
    public class ImportResult
    {
        public bool Success { get; set; }

        public List<string> Errors { get; set; } = new();

        public List<string> Warnings { get; set; } = new();

        public int TotalInserted { get; set; }
    }
}
