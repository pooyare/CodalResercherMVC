using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace CodalResearcherMVC.Models
{
    public class MonthlyReportData
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int MonthlyReportId { get; set; }

        [Required]
        public string ProductName { get; set; }

        [Required]
        public string UnitName { get; set; }

        [Required]
        public string ProductType { get; set; } // تولید/فروش داخلی/فروش صادراتی

        [Required]
        public string PeriodType { get; set; } // از ابتدای سال مالی / دوره یک ماهه منتهی / etc.

        [Required]
        public DateTime PeriodDate { get; set; }

        public int? ProductionQuantity { get; set; }

        public decimal? SalesQuantity { get; set; }

        public decimal? SalesRate { get; set; }

        public decimal? SalesAmount { get; set; }

        [ForeignKey("MonthlyReportId")]
        public virtual MonthlyReportEntity Report { get; set; }
    }
}
