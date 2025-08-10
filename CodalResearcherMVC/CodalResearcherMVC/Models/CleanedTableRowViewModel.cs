namespace CodalResearcherMVC.Models
{
    public class CleanedTableRowViewModel
    {
        public int Id { get; set; }
        public int LevelId { get; set; } // شماره جدول یا سطح
        public string FileName { get; set; }
        public string? MainCategory { get; set; } // مثل "دارایی‌ها" یا "حقوق مالکانه و بدهی‌ها"
        public string? SubCategory { get; set; } // مثل "دارایی‌های جاری" یا "بدهی‌های غیرجاری"
        public string? Description { get; set; } // شرح
        public string? Period { get; set; } // دوره منتهی به...
        public string? AuditStatus { get; set; } // حسابرسی شده / نشده
        public string? Value { get; set; } // عدد
        public string? ChangePercent { get; set; } // درصد تغییر
        public string? CompanySymbol { get; set; } // اضافه کردن نام شرکت
    }
}
