using CodalResearcherMVC.Models;
using DocumentFormat.OpenXml.InkML;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

namespace CodalResearcherMVC.Entities
{
    public partial class ExcelTableDbContext : DbContext
    {
        public ExcelTableDbContext(DbContextOptions<ExcelTableDbContext> options) : base(options)
        {

        }
        public DbSet<ExcelTable> ExcelTables { get; set; }
        public DbSet<CleanedTableRowViewModel> CleanedTableRows { get; set; }
        public DbSet<MonthlyReportEntity> MonthlyReports { get; set; }
        public DbSet<MonthlyReportData> MonthlyReportData { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Configure relationships
            modelBuilder.Entity<MonthlyReportEntity>()
                .HasMany(r => r.DataEntries)
                .WithOne(d => d.Report)
                .HasForeignKey(d => d.MonthlyReportId)
                .OnDelete(DeleteBehavior.Cascade);

            // Optimize indices
            modelBuilder.Entity<MonthlyReportData>()
                .HasIndex(d => d.MonthlyReportId);

            modelBuilder.Entity<MonthlyReportData>()
                .HasIndex(d => d.ProductName);

            modelBuilder.Entity<MonthlyReportData>()
                .HasIndex(d => d.PeriodDate);

            modelBuilder.Entity<MonthlyReportData>()
                .HasIndex(d => d.ProductType);
        }
    }
}
