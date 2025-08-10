using CodalResearcherMVC.Entities;
using CodalResearcherMVC.Models;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodalResearcherMVC.Services
{
    public class ExcelTableService
    {
        private readonly ExcelTableDbContext _context;

        public ExcelTableService(ExcelTableDbContext context)
        {
            _context = context;
        }

        #region ذخیره اکسل های اطلاعات صورت های مالی
        public async Task SaveMatchedTablesAsync(List<CleanedTableRowViewModel> cleanedRows)
        {
            var entities = cleanedRows.Select(r => new CleanedTableRowViewModel
            {
                LevelId = r.LevelId,
                FileName = r.FileName,
                MainCategory = r.MainCategory,
                SubCategory = r.SubCategory,
                Description = r.Description,
                Period = r.Period,
                AuditStatus = r.AuditStatus,
                Value = r.Value,
                ChangePercent = r.ChangePercent,
                CompanySymbol = r.CompanySymbol, // اضافه کردن نام شرکت
            });

            _context.CleanedTableRows.AddRange(entities);
            await _context.SaveChangesAsync();
        }
        #endregion

        #region ذخیره جدول گزارش فعالیت ماهانه
        public async Task SaveMonthlyReportAsync(List<MonthlyActivityTableViewModel> reports)
        {
            foreach (var report in reports)
            {
                if (report?.Headers == null || report.Rows == null || !report.Headers.Any() || !report.Rows.Any())
                    continue;

                // Extract report date from the title
                DateTime reportDate = ExtractDateFromTitle(report.Title);

                // Create or update report entity
                var reportEntity = await _context.MonthlyReports.FirstOrDefaultAsync(r =>
                    r.Link == report.Link && r.ReportDate == reportDate);

                if (reportEntity == null)
                {
                    reportEntity = new MonthlyReportEntity
                    {
                        Title = report.Title,
                        Link = report.Link,
                        ImportDate = DateTime.Now,
                        ReportDate = reportDate,
                        CompanySymbol = report.CompanySymbol
                        
                    };
                    _context.MonthlyReports.Add(reportEntity);
                    await _context.SaveChangesAsync(); // Save to get the ID
                }

                // Process and categorize headers
                var headerCategories = CategorizeHeaders(report.Headers);

                // Process each row and store data
                foreach (var row in report.Rows)
                {
                    // Skip empty or header rows
                    if (row.Columns.Count == 0)
                        continue;

                    // Extract product name and unit
                    string productName = GetProductName(row, report.Headers);
                    string unitName = GetUnitName(row, report.Headers);
                    string productType = DetermineProductType(row, report.Headers);

                    if (string.IsNullOrWhiteSpace(productName))
                        continue; // Skip rows without product name

                    // Process each period category (e.g., "از ابتدای سال مالی تا تاریخ ...")
                    foreach (var category in headerCategories)
                    {
                        // Skip categories we're not interested in
                        if (string.IsNullOrWhiteSpace(category.PeriodType) || category.Headers.Count == 0)
                            continue;

                        // Extract date from period type
                        DateTime periodDate = ExtractDateFromPeriodType(category.PeriodType);

                        // Create data entry for this period
                        var dataEntry = new MonthlyReportData
                        {
                            MonthlyReportId = reportEntity.Id,
                            ProductName = productName,
                            UnitName = unitName,
                            ProductType = productType,
                            PeriodType = category.PeriodType,
                            PeriodDate = periodDate
                        };

                        // Process each metric in this period category
                        PopulateMetrics(dataEntry, row, category.Headers);

                        // Only add if we have at least one metric populated
                        if (dataEntry.ProductionQuantity.HasValue ||
                            dataEntry.SalesQuantity.HasValue ||
                            dataEntry.SalesRate.HasValue ||
                            dataEntry.SalesAmount.HasValue)
                        {
                            _context.Set<MonthlyReportData>().Add(dataEntry);
                        }
                    }
                }

                await _context.SaveChangesAsync();
            }
        }

        private List<HeaderCategory> CategorizeHeaders(List<string> headers)
        {
            var categories = new List<HeaderCategory>();

            // Define period patterns (date patterns to recognize in headers)
            var periodPatterns = new[]
            {
                @"از ابتدای سال مالی تا تاریخ \d{4}/\d{2}/\d{2}",
                @"دوره یک ماهه منتهی به \d{4}/\d{2}/\d{2}",
                @"دوره سه ماهه منتهی به \d{4}/\d{2}/\d{2}"
            };

            // Group headers by period type
            foreach (var header in headers)
            {
                string periodType = null;

                // Find which period pattern this header belongs to
                foreach (var pattern in periodPatterns)
                {
                    var match = Regex.Match(header, pattern);
                    if (match.Success)
                    {
                        periodType = match.Value;
                        break;
                    }
                }

                if (periodType != null)
                {
                    // Find or create the category
                    var category = categories.FirstOrDefault(c => c.PeriodType == periodType);
                    if (category == null)
                    {
                        category = new HeaderCategory { PeriodType = periodType };
                        categories.Add(category);
                    }

                    category.Headers.Add(header);
                }
                else if (header.Contains("نام محصول") || header.Contains("واحد") ||
                        header.Contains("وضعیت") || header.Contains("فروش داخلی") ||
                        header.Contains("فروش صادراتی"))
                {
                    // These are specific headers we don't categorize by period
                }
                else
                {
                    // Add to "Other" category for any uncategorized headers
                    var otherCategory = categories.FirstOrDefault(c => c.PeriodType == "Other");
                    if (otherCategory == null)
                    {
                        otherCategory = new HeaderCategory { PeriodType = "Other" };
                        categories.Add(otherCategory);
                    }

                    otherCategory.Headers.Add(header);
                }
            }

            return categories;
        }

        private string GetProductName(MonthlyActivityRow row, List<string> headers)
        {
            // Try common product name headers
            var productHeaders = new[] { "نام محصول", "محصول" };
            foreach (var header in productHeaders)
            {
                var productHeader = headers.FirstOrDefault(h => h.Contains(header));
                if (productHeader != null && row.Columns.TryGetValue(productHeader, out var value))
                    return value?.Trim() ?? "";
            }

            // If not found, return the first column as a fallback
            if (headers.Any() && row.Columns.TryGetValue(headers[0], out var firstValue))
                return firstValue?.Trim() ?? "";

            return "";
        }

        private string GetUnitName(MonthlyActivityRow row, List<string> headers)
        {
            // Try common unit headers
            var unitHeaders = new[] { "واحد" };
            foreach (var header in unitHeaders)
            {
                var unitHeader = headers.FirstOrDefault(h => h.Contains(header));
                if (unitHeader != null && row.Columns.TryGetValue(unitHeader, out var value))
                    return value?.Trim() ?? "";
            }

            // Default unit if not found
            return "تن";
        }

        private string DetermineProductType(MonthlyActivityRow row, List<string> headers)
        {
            // Check row for indicators of product type
            foreach (var header in headers)
            {
                if (header.Contains("وضعیت") && row.Columns.TryGetValue(header, out var value))
                {
                    if (value?.Contains("تولید") == true)
                        return "تولید";
                    if (value?.Contains("فروش داخلی") == true)
                        return "فروش داخلی";
                    if (value?.Contains("فروش صادراتی") == true)
                        return "فروش صادراتی";
                }
            }

            // Check if the row is in a specific section by context
            if (headers.Any(h => h.Contains("فروش داخلی")))
                return "فروش داخلی";
            if (headers.Any(h => h.Contains("فروش صادراتی")))
                return "فروش صادراتی";

            // Default if we can't determine
            return "تولید";
        }

        private void PopulateMetrics(MonthlyReportData dataEntry, MonthlyActivityRow row, List<string> periodHeaders)
        {
            foreach (var header in periodHeaders)
            {
                // Extract value if it exists
                if (!row.Columns.TryGetValue(header, out var value) || string.IsNullOrWhiteSpace(value))
                    continue;

                // Clean value - remove commas, convert Persian digits, etc.
                value = CleanNumericValue(value);

                // Check which metric this header represents
                if (header.Contains("تعداد تولید"))
                {
                    int.TryParse(value, out int quantity);
                    dataEntry.ProductionQuantity = quantity;
                }
                else if (header.Contains("تعداد فروش"))
                {
                    decimal.TryParse(value, out decimal quantity);
                    dataEntry.SalesQuantity = quantity;
                }
                else if (header.Contains("نرخ فروش"))
                {
                    decimal.TryParse(value, out decimal rate);
                    dataEntry.SalesRate = rate;
                }
                else if (header.Contains("مبلغ فروش"))
                {
                    decimal.TryParse(value, out decimal amount);
                    dataEntry.SalesAmount = amount;
                }
            }
        }

        private DateTime ExtractDateFromTitle(string title)
        {
            // Try to extract date from title
            var dateMatch = Regex.Match(title, @"\d{4}/\d{2}/\d{2}");
            if (dateMatch.Success)
            {
                if (DateTime.TryParseExact(dateMatch.Value, "yyyy/MM/dd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                {
                    return date;
                }
            }

            // Default to current date if no date found
            return DateTime.Now;
        }

        private DateTime ExtractDateFromPeriodType(string periodType)
        {
            // Extract date part from period string
            var dateMatch = Regex.Match(periodType, @"\d{4}/\d{2}/\d{2}");
            if (dateMatch.Success)
            {
                if (DateTime.TryParseExact(dateMatch.Value, "yyyy/MM/dd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                {
                    return date;
                }
            }

            // Default to current date if no date found
            return DateTime.Now;
        }

        private string CleanNumericValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "0";

            // Remove commas and convert Persian digits to Latin
            value = value.Replace(",", "")
                         .Replace("۰", "0")
                         .Replace("۱", "1")
                         .Replace("۲", "2")
                         .Replace("۳", "3")
                         .Replace("۴", "4")
                         .Replace("۵", "5")
                         .Replace("۶", "6")
                         .Replace("۷", "7")
                         .Replace("۸", "8")
                         .Replace("۹", "9");

            // Remove any remaining non-numeric characters
            return Regex.Replace(value, @"[^\d.-]", "");
        }
    }

    // Helper class for categorizing headers
    public class HeaderCategory
    {
        public string PeriodType { get; set; }
        public List<string> Headers { get; set; } = new List<string>();
    }

    public interface IExcelTableService
    {
        Task SaveMonthlyReportAsync(List<MonthlyActivityTableViewModel> reports);
    }
    #endregion

}
