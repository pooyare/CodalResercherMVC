using ClosedXML.Excel;
using CodalResearcherMVC.Models;
using ExcelDataReader;
using HtmlAgilityPack;
using Newtonsoft.Json;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;


namespace CodalResearcherMVC.Services
{
    public class HtmlExcelParser
    {
        public async Task<string> ConvertToHtmlAsync(string filePath)
        {
            var content = await File.ReadAllTextAsync(filePath);
            return content;
        }
        public List<(string keyword, int index)> FindKeywordsInHtml(string htmlContent, List<string> keywords)
        {
            var results = new List<(string keyword, int index)>();

            foreach (var keyword in keywords)
            {
                var matches = Regex.Matches(htmlContent, Regex.Escape(keyword), RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    results.Add((keyword, match.Index));
                }
            }

            return results.OrderBy(r => r.index).ToList();
        }
        public List<MatchedTable> ExtractTablesAfterKeywords(string htmlContent, List<string> keywords, string fileName, string companySymbol)
        {
            var results = new List<MatchedTable>();

            foreach (var keyword in keywords)
            {
                var match = Regex.Match(htmlContent, Regex.Escape(keyword), RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var remainingHtml = htmlContent.Substring(match.Index);

                    // پیدا کردن اولین جدول بعد از کلمه
                    var tableMatch = Regex.Match(remainingHtml, @"<table.*?</table>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    if (tableMatch.Success)
                    {
                        results.Add(new MatchedTable
                        {
                            Keyword = keyword,
                            FileName = fileName,
                            HtmlTable = tableMatch.Value,
                            CompanySymbol = companySymbol // اضافه کردن نام شرکت
                        });
                    }
                }
            }

            return results;
        }

        //public List<CleanedTableRowViewModel> ParseHtmlTable(string htmlContent,string keyword)
        //{
        //    var doc = new HtmlDocument();
        //    doc.LoadHtml(htmlContent);

        //    var tables = doc.DocumentNode.SelectNodes("//table");
        //    var result = new List<CleanedTableRowViewModel>();
        //    int levelId = 1;

        //    foreach (var table in tables)
        //    {
        //        var theadRows = table.SelectNodes(".//thead/tr");
        //        if (theadRows == null || theadRows.Count < 2) continue;

        //        // استخراج header های دوره و وضعیت حسابرسی
        //        var periodHeaders = theadRows[0].SelectNodes(".//th/span")?.Skip(1).TakeWhile(x => x.InnerText.Trim() != "درصد تغيير").ToList();
        //        var auditStatusHeaders = theadRows[1].SelectNodes(".//th/span")?.ToList();
        //        if (periodHeaders == null || auditStatusHeaders == null) continue;

        //        var tbodyRows = table.SelectNodes(".//tbody/tr");
        //        if (tbodyRows == null) continue;

        //        string? currentCategory = null;

        //        foreach (var row in tbodyRows)
        //        {
        //            var cells = row.SelectNodes(".//td/span")?.Select(x => x.InnerText.Trim()).ToList();
        //            if (cells == null || cells.Count == 0) continue;

        //            string firstCell = cells[0];
        //            bool isCategory = firstCell.Contains(":");

        //            if (isCategory)
        //            {
        //                // به‌روزرسانی دسته‌بندی جاری
        //                currentCategory = firstCell;
        //                continue; // خود این ردیف ذخیره نشه
        //            }

        //            string? changePercent = cells.Count > periodHeaders.Count ? cells.Last() : null;

        //            for (int i = 0; i < periodHeaders.Count && i + 1 < cells.Count; i++)
        //            {
        //                var model = new CleanedTableRowViewModel
        //                {
        //                    LevelId = levelId,
        //                    MainCategory = currentCategory,
        //                    SubCategory = null,
        //                    Description = firstCell,
        //                    Period = periodHeaders[i].InnerText.Trim(),
        //                    AuditStatus = i < auditStatusHeaders.Count ? auditStatusHeaders[i].InnerText.Trim() : null,
        //                    Value = cells[i + 1],
        //                    ChangePercent = changePercent // درج همون درصد تغییر در تمام رکوردها
        //                };

        //                result.Add(model);
        //            }
        //        }

        //        levelId++;
        //    }

        //    return result;
        //}


        public List<CleanedTableRowViewModel> ParseHtmlTable(string htmlContent, string keyword, string FileName, string companySymbol)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(htmlContent);

            var tables = doc.DocumentNode.SelectNodes("//table");
            var result = new List<CleanedTableRowViewModel>();
            int levelId = 1;

            bool isIncomeStatement = keyword.Contains("صورت سود و زیان");
            bool isBalanceSheet = keyword.Contains("صورت وضعیت مالی");

            foreach (var table in tables)
            {
                var theadRows = table.SelectNodes(".//thead/tr");
                if (theadRows == null || theadRows.Count < 2) continue;

                // استخراج هدرهای دوره از ردیف اول
                var periodHeaderCells = theadRows[0].SelectNodes(".//th/span");
                if (periodHeaderCells == null || periodHeaderCells.Count < 2) continue;

                // اولین سلول معمولا "شرح" است، آن را حذف می‌کنیم
                var periods = periodHeaderCells.Skip(1)
                                              .Where(x => !x.InnerText.Trim().Contains("درصد تغيير"))
                                              .Select(x => x.InnerText.Trim())
                                              .ToList();

                // استخراج وضعیت‌های حسابرسی از ردیف دوم
                var auditStatusCells = theadRows[1].SelectNodes(".//th/span");
                if (auditStatusCells == null) continue;

                var auditStatuses = auditStatusCells
                                  .Select(x => x.InnerText.Trim())
                                  .Where(x => x.Contains("حسابرس"))
                                  .ToList();

                var tbodyRows = table.SelectNodes(".//tbody/tr");
                if (tbodyRows == null) continue;

                string currentMainCategory = null;
                string currentSubCategory = null;

                foreach (var row in tbodyRows)
                {
                    var cells = row.SelectNodes(".//td/span")?.Select(x => x.InnerText.Trim()).ToList();
                    if (cells == null || cells.Count == 0) continue;

                    string description = cells[0];
                    string normalizedDescription = NormalizeText(description);

                    // بررسی و تعیین دسته‌بندی‌ها
                    if (isBalanceSheet)
                    {
                        // دسته‌بندی اصلی
                        if (normalizedDescription.Contains("دارایی‌ها") || normalizedDescription.Contains("داراییها"))
                        {
                            if (normalizedDescription == "دارایی‌ها" || normalizedDescription == "داراییها")
                            {
                                currentMainCategory = "دارایی‌ها";
                                continue; // این ردیف فقط هدر است، نیازی به ذخیره‌سازی ندارد
                            }
                        }
                        else if (normalizedDescription.Contains("حقوق مالکانه و بدهی‌ها") || normalizedDescription.Contains("حقوق مالکانه و بدهیها"))
                        {
                            if (normalizedDescription == "حقوق مالکانه و بدهی‌ها" || normalizedDescription == "حقوق مالکانه و بدهیها")
                            {
                                currentMainCategory = "حقوق مالکانه و بدهی‌ها";
                                continue;
                            }
                        }

                        // دسته‌بندی فرعی - استفاده از Contains به جای تطابق دقیق
                        if ((normalizedDescription.Contains("دارایی‌های غیرجاری") || normalizedDescription.Contains("داراییهای غیرجاری"))
                            && !normalizedDescription.Contains("جمع"))
                        {
                            currentSubCategory = "دارایی‌های غیرجاری";
                            continue;
                        }
                        else if ((normalizedDescription.Contains("دارایی‌های جاری") || normalizedDescription.Contains("داراییهای جاری"))
                            && !normalizedDescription.Contains("جمع"))
                        {
                            currentSubCategory = "دارایی‌های جاری";
                            continue;
                        }
                        else if (normalizedDescription.Contains("حقوق مالکانه") && !normalizedDescription.Contains("جمع") && !normalizedDescription.Contains("بدهی"))
                        {
                            currentSubCategory = "حقوق مالکانه";
                            continue;
                        }
                        else if ((normalizedDescription.Contains("بدهی‌ها") || normalizedDescription.Contains("بدهیها"))
                            && !normalizedDescription.Contains("جمع") && !normalizedDescription.Contains("حقوق"))
                        {
                            currentSubCategory = "بدهی‌ها";
                            continue;
                        }
                        else if ((normalizedDescription.Contains("بدهی‌های غیرجاری") || normalizedDescription.Contains("بدهیهای غیرجاری"))
                            && !normalizedDescription.Contains("جمع"))
                        {
                            currentSubCategory = "بدهی‌های غیرجاری";
                            continue;
                        }
                        else if ((normalizedDescription.Contains("بدهی‌های جاری") || normalizedDescription.Contains("بدهیهای جاری"))
                            && !normalizedDescription.Contains("جمع"))
                        {
                            currentSubCategory = "بدهی‌های جاری";
                            continue;
                        }
                    }
                    else if (isIncomeStatement)
                    {
                        // برای صورت سود و زیان، دسته‌بندی‌ها با ":" مشخص می‌شوند
                        if (normalizedDescription.Contains(":"))
                        {
                            currentMainCategory = description;
                            continue;
                        }
                    }

                    // در صورت وجود ستون درصد تغییر، آن را پیدا کنیم
                    string changePercent = null;
                    int valueCount = Math.Min(periods.Count, cells.Count - 1); // تعداد مقادیر (بدون ستون توضیحات)

                    // اگر تعداد سلول‌ها بیشتر از تعداد دوره‌ها + 1 (ستون توضیحات) باشد، آخرین سلول درصد تغییر است
                    if (cells.Count > periods.Count + 1)
                    {
                        changePercent = FormatNumberValue(cells.Last());
                    }

                    // ساخت مدل برای هر دوره
                    for (int i = 0; i < valueCount; i++)
                    {
                        string value = (i + 1 < cells.Count) ? cells[i + 1] : null;

                        // اصلاح فرمت عدد (تبدیل اعداد داخل پرانتز به اعداد منفی با علامت -)
                        value = FormatNumberValue(value);

                        var model = new CleanedTableRowViewModel
                        {
                            LevelId = levelId,
                            FileName = FileName,
                            MainCategory = currentMainCategory,
                            SubCategory = currentSubCategory,
                            Description = description,
                            Period = i < periods.Count ? periods[i] : null,
                            AuditStatus = i < auditStatuses.Count ? auditStatuses[i] : null,
                            Value = value,
                            ChangePercent = changePercent,
                            CompanySymbol = companySymbol
                        };

                        result.Add(model);
                    }

                    levelId++;
                }
            }

            return result;
        }

        // متد کمکی برای فرمت کردن اعداد
        private string FormatNumberValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "0";

            // اگر عدد داخل پرانتز باشد، آن را به عدد منفی با علامت - تبدیل می‌کنیم
            if (value.StartsWith("(") && value.EndsWith(")"))
            {
                string numberPart = value.Substring(1, value.Length - 2);
                return "-" + numberPart;
            }

            // تبدیل ۰ فارسی به 0 انگلیسی
            if (value == "۰")
                return "0";

            return value;
        }

        private string NormalizeText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // جایگزینی کاراکترهای عربی با معادل‌های فارسی
            text = text.Replace('ي', 'ی')
                       .Replace('ك', 'ک');

            // حذف فاصله‌های اضافی
            text = Regex.Replace(text, @"\s+", " ").Trim();

            return text;
        }

        #region گرفتن جدول صفحه گزارش فعالیت ماهانه
        public (List<string> Headers, List<MonthlyActivityRow> Rows) ParseMonthlyActivityTable(string html)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            // 1. پیدا کردن جدول اصلی
            var table = doc.DocumentNode.SelectSingleNode("//table[@class='rayanDynamicStatement']");
            if (table == null)
            {
                table = doc.DocumentNode.SelectSingleNode("//table"); // Fallback
            }

            if (table == null)
            {
                return (new List<string>(), new List<MonthlyActivityRow>());
            }

            // 2. استخراج ساختار کامل هدرها با در نظر گرفتن ادغام سلول‌ها
            List<string> flattenedHeaders = ExtractComplexHeaders(table);

            // 3. استخراج داده‌های سطرها
            var rows = new List<MonthlyActivityRow>();
            var dataRows = table.SelectNodes(".//tbody/tr");

            if (dataRows == null)
            {
                // Fallback: اگر tbody پیدا نشد، سطرهای بعد از هدرها را به عنوان داده در نظر بگیر
                var allRows = table.SelectNodes(".//tr");
                if (allRows != null && allRows.Count > 2)
                {
                    dataRows = new HtmlNodeCollection(table);
                    for (int i = 2; i < allRows.Count; i++) // فرض می‌کنیم دو سطر اول هدر هستند
                    {
                        dataRows.Add(allRows[i]);
                    }
                }
            }

            if (dataRows != null)
            {
                foreach (var tr in dataRows)
                {
                    // فقط سطرهایی که hidden نیستند را در نظر بگیر
                    if (HasHiddenAttribute(tr)) continue;

                    var cells = tr.SelectNodes(".//td[not(@hidden)]");
                    if (cells == null || cells.Count == 0) continue;

                    var row = new MonthlyActivityRow();
                    int columnIndex = 0;

                    for (int i = 0; i < cells.Count && columnIndex < flattenedHeaders.Count; i++)
                    {
                        var cell = cells[i];

                        // فقط سلول‌هایی که hidden نیستند را در نظر بگیر
                        if (HasHiddenAttribute(cell)) continue;

                        string cellValue = CleanHeaderText(cell.InnerText);

                        // بررسی colspan برای سلول‌های با عرض چندتایی
                        int colSpan = 1;
                        if (cell.Attributes["colspan"] != null && int.TryParse(cell.Attributes["colspan"].Value, out int parsedColSpan))
                        {
                            colSpan = parsedColSpan;
                        }

                        // مقدار را برای هر ستون اختصاص بده
                        for (int c = 0; c < colSpan && columnIndex < flattenedHeaders.Count; c++)
                        {
                            if (columnIndex < flattenedHeaders.Count)
                            {
                                row.Columns[flattenedHeaders[columnIndex]] = cellValue;
                                columnIndex++;
                            }
                        }
                    }

                    // بررسی اگر ردیف خالی نیست، آن را اضافه کن
                    if (row.Columns.Count > 0)
                    {
                        rows.Add(row);
                    }
                }
            }

            return (flattenedHeaders, rows);
        }

        private List<string> ExtractComplexHeaders(HtmlNode table)
        {
            // Find all header rows, with multiple fallback strategies
            var headerRows = FindHeaderRows(table);
            if (headerRows == null || headerRows.Count == 0)
            {
                return new List<string>();
            }

            // Enhanced column calculation that considers maximum potential spread
            int totalColumns = CalculateRobustTotalColumns(headerRows);

            // Create a more flexible grid for header tracking
            string[,] headerGrid = new string[headerRows.Count, totalColumns];
            bool[,] cellOccupied = new bool[headerRows.Count, totalColumns];

            // Robust header extraction with advanced tracking
            PopulateHeaderGrid(headerRows, headerGrid, cellOccupied, totalColumns);

            // Generate more intelligent combined headers
            return GenerateIntelligentHeaders(headerGrid, headerRows.Count, totalColumns);
        }

        private HtmlNodeCollection FindHeaderRows(HtmlNode table)
        {
            // Multiple strategies to find header rows
            var headerRows = table.SelectNodes(".//thead/tr");

            if (headerRows == null || headerRows.Count == 0)
            {
                // Strategy 2: Look for rows with mostly <th> elements
                var allRows = table.SelectNodes(".//tr");
                if (allRows != null)
                {
                    headerRows = new HtmlNodeCollection(table);
                    foreach (var row in allRows)
                    {
                        // Skip rows with hidden attribute
                        if (HasHiddenAttribute(row)) continue;

                        var thCells = row.SelectNodes(".//th[not(@hidden)]");
                        var tdCells = row.SelectNodes(".//td[not(@hidden)]");

                        // If row has more th elements than td, consider it a header row
                        if (thCells != null && thCells.Count > 0 &&
                            (tdCells == null || thCells.Count >= tdCells.Count))
                        {
                            headerRows.Add(row);
                        }

                        // Stop searching after finding reasonable header rows
                        if (headerRows.Count >= 2) break;
                    }
                }
            }

            return headerRows;
        }

        private int CalculateRobustTotalColumns(HtmlNodeCollection headerRows)
        {
            int maxColumns = 0;
            foreach (var row in headerRows)
            {
                var cells = row.SelectNodes(".//th[not(@hidden)]|.//td[not(@hidden)]");
                if (cells == null) continue;

                int rowColumns = 0;
                foreach (var cell in cells)
                {
                    // Skip hidden cells
                    if (HasHiddenAttribute(cell)) continue;

                    int colSpan = 1;

                    if (cell.Attributes["colspan"] != null &&
                        int.TryParse(cell.Attributes["colspan"].Value, out int parsedColSpan))
                    {
                        colSpan = parsedColSpan;
                    }

                    rowColumns += colSpan;
                }

                maxColumns = Math.Max(maxColumns, rowColumns);
            }

            return maxColumns;
        }

        private void PopulateHeaderGrid(
            HtmlNodeCollection headerRows,
            string[,] headerGrid,
            bool[,] cellOccupied,
            int totalColumns)
        {
            for (int rowIndex = 0; rowIndex < headerRows.Count; rowIndex++)
            {
                var headerCells = headerRows[rowIndex].SelectNodes(".//th[not(@hidden)]|.//td[not(@hidden)]");
                if (headerCells == null) continue;

                int colIndex = 0;
                foreach (var cell in headerCells)
                {
                    // Skip hidden cells
                    if (HasHiddenAttribute(cell)) continue;

                    // Find next available column
                    while (colIndex < totalColumns && cellOccupied[rowIndex, colIndex])
                    {
                        colIndex++;
                    }

                    if (colIndex >= totalColumns) break;

                    string headerText = CleanHeaderText(cell.InnerText);

                    int colSpan = 1, rowSpan = 1;

                    if (cell.Attributes["colspan"] != null &&
                        int.TryParse(cell.Attributes["colspan"].Value, out int parsedColSpan))
                    {
                        colSpan = parsedColSpan;
                    }

                    if (cell.Attributes["rowspan"] != null &&
                        int.TryParse(cell.Attributes["rowspan"].Value, out int parsedRowSpan))
                    {
                        rowSpan = parsedRowSpan;
                    }

                    // Mark cells as occupied and populate with header text
                    for (int r = 0; r < rowSpan && (rowIndex + r) < headerRows.Count; r++)
                    {
                        for (int c = 0; c < colSpan && (colIndex + c) < totalColumns; c++)
                        {
                            headerGrid[rowIndex + r, colIndex + c] = headerText;
                            cellOccupied[rowIndex + r, colIndex + c] = true;
                        }
                    }

                    colIndex += colSpan;
                }
            }
        }

        private List<string> GenerateIntelligentHeaders(
            string[,] headerGrid,
            int rowCount,
            int columnCount)
        {
            var finalHeaders = new List<string>();

            for (int col = 0; col < columnCount; col++)
            {
                var columnTexts = new List<string>();

                for (int row = 0; row < rowCount; row++)
                {
                    string currentText = headerGrid[row, col];
                    if (!string.IsNullOrEmpty(currentText) &&
                        !columnTexts.Contains(currentText))
                    {
                        columnTexts.Add(currentText);
                    }
                }

                // Intelligent header combination
                string combinedHeader = string.Join(" - ",
                    columnTexts.Where(t => !string.IsNullOrWhiteSpace(t)));

                finalHeaders.Add(combinedHeader);
            }

            return finalHeaders;
        }

        private string CleanHeaderText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            // Remove debug info and other Angular-related content
            string cleaned = Regex.Replace(text, "<app-debug-info[\\s\\S]*?</app-debug-info>", "");

            // Extract text from span tags if present
            var match = Regex.Match(cleaned, "<span[^>]*>([^<]*)</span>");
            if (match.Success && !string.IsNullOrWhiteSpace(match.Groups[1].Value))
            {
                cleaned = match.Groups[1].Value;
            }

            // Remove remaining HTML tags
            cleaned = Regex.Replace(cleaned, "<[^>]*>", "");

            return cleaned.Trim()
                .Replace("\r", "")
                .Replace("\n", " ")
                .Replace("  ", " ");
        }

        // Helper method to check if a node has hidden attribute
        private bool HasHiddenAttribute(HtmlNode node)
        {
            if (node == null) return true;

            if (node.Attributes["hidden"] != null)
            {
                return true;
            }

            var styleAttr = node.Attributes["style"]?.Value;
            return styleAttr != null &&
                   (styleAttr.Contains("display: none") ||
                    styleAttr.Contains("visibility: hidden"));
        }
    #endregion
}
}
