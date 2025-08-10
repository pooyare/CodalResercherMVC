namespace CodalResearcherMVC.Models
{
    public class ExcelTable
    {
        public int? Id { get; set; }

        public string? Keyword { get; set; } // کلمه کلیدی که جدول بعدش پیدا شد

        public string? FileName { get; set; } // اسم فایل اصلی

        public string? HtmlContent { get; set; } // محتوای کامل جدول (HTML)

        public DateTime? CreatedAt { get; set; } = DateTime.Now;
    }
}
