namespace CodalResearcherMVC.Models
{
    public class MatchedTable
    {
        public int Id { get; set; }
        public string Keyword { get; set; }
        public string FileName { get; set; }
        public string HtmlTable { get; set; }
        public string CompanySymbol { get; set; } // اضافه کردن نام شرکت
    }
}
