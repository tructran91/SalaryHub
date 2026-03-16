using Microsoft.EntityFrameworkCore;
using SalaryHub.Entities;

namespace SalaryHub.Data
{
    public class AppDbContext : DbContext
    {
        public DbSet<User> Users { get; set; }

        public DbSet<SalaryRecord> SalaryRecords { get; set; }

        public DbSet<SystemLog> SystemLogs { get; set; }

        public AppDbContext(DbContextOptions<AppDbContext> options)
            : base(options)
        {
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<User>()
                .HasIndex(x => x.EmployeeCode)
                .IsUnique();

            //modelBuilder.Entity<SalaryRecord>()
            //    .HasIndex(x => new { x.UserId, x.Month, x.Year, x.Title })
            //    .IsUnique();
        }
    }
}
