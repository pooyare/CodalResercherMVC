namespace CodalResearcherMVC.Models
{
    public class ErrorViewModel
    {
        public string Title { get; set; } = "یک خطا رخ داده است";
        public string Message { get; set; }
        public string ErrorDetails { get; set; }
        public bool ShowDetails { get; set; } = false;
        public string ReturnUrl { get; set; } = "/";
    }
}
