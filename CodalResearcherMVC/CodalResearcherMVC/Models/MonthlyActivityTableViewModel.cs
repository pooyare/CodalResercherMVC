namespace CodalResearcherMVC.Models
{
    public class MonthlyActivityTableViewModel
    {
        public string Title { get; set; }
        public string Link { get; set; }
        public List<string> Headers { get; set; } = new List<string>();
        public List<MonthlyActivityRow> Rows { get; set; } = new List<MonthlyActivityRow>();
        public string RawHtml { get; set; } // Added for debugging purposes
        public string CompanySymbol { get; set; } // پروپرتی جدید برای سمبل شرکت
    }
}
