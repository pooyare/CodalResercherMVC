using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace CodalResearcherMVC.Models
{
    public class CodalSearchViewModel
    {
        [Required]
        public string Symbol { get; set; }
        [Required]
        public string Keyword { get; set; }
        public List<ReportResult> Results { get; set; } = new List<ReportResult>();
    }

    public class ReportResult
    {
        public string Title { get; set; }
        public string Link { get; set; }
        public string ExcelLink { get; set; }
        public string CompanySymbol { get; set; } // نام شرکت اضافه شد
    }
}
