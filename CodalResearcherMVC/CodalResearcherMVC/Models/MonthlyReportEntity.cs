using System.ComponentModel.DataAnnotations;

namespace CodalResearcherMVC.Models
{
    public class MonthlyReportEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string Title { get; set; }

        [Required]
        public string Link { get; set; }

        public string CompanySymbol { get; set; }

        [Required]
        public DateTime ImportDate { get; set; }

        [Required]
        public DateTime ReportDate { get; set; }

        // Navigation property
        public virtual ICollection<MonthlyReportData> DataEntries { get; set; }
    }
}
