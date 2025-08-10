using CodalResearcherMVC.Models;
using CodalResearcherMVC.Services;
using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

public class CodalController : Controller
{
    private readonly CodalScraperSelenium _scraper;
    private readonly CodalExcelDownloader _excelDownloader;
    private readonly HtmlExcelParser _htmlParser;
    private readonly ExcelTableService _excelTableService;
    private readonly ILogger<CodalController> _logger;
    private readonly IHubContext<ProcessingHub> _hubContext; // SignalR Hub Context

    public CodalController(ExcelTableService excelTableService, ILogger<CodalController> logger, IHubContext<ProcessingHub> hubContext)
    {
        _scraper = new CodalScraperSelenium();
        _excelDownloader = new CodalExcelDownloader();
        _htmlParser = new HtmlExcelParser();
        _excelTableService = excelTableService;
        _logger = logger;
        _hubContext = hubContext;
    }

    [HttpGet]
    public IActionResult Index()
    {
        return View(new CodalSearchViewModel());
    }

    #region سرچ در صفحه codal
    [HttpPost]
    public async Task<IActionResult> Search(CodalSearchViewModel model)
    {
        try
        {
            // گزارش وضعیت: شروع جستجو
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "شروع عملیات جستجو در کدال");

            if (!ModelState.IsValid)
            {
                await _hubContext.Clients.All.SendAsync("UpdateStatus", "فرم جستجو معتبر نیست. لطفاً مجدداً تلاش کنید", "error");
                return View("Index", model);
            }

            // Define both Persian and English commas
            char[] commaDelimiters = new[] { ',', '،' };

            // Split the keyword input by commas (both Persian and English)
            var keywords = model.Keyword?.Split(commaDelimiters, StringSplitOptions.RemoveEmptyEntries)
                                .Select(k => k.Trim())
                                .Where(k => !string.IsNullOrWhiteSpace(k))
                                .ToList() ?? new List<string>();

            // Split the symbol input by commas (both Persian and English)
            var symbols = model.Symbol?.Split(commaDelimiters, StringSplitOptions.RemoveEmptyEntries)
                              .Select(s => s.Trim())
                              .Where(s => !string.IsNullOrWhiteSpace(s))
                              .ToList() ?? new List<string>();

            // If no valid keywords or symbols, return the view with empty results
            if (!keywords.Any() || !symbols.Any())
            {
                await _hubContext.Clients.All.SendAsync("UpdateStatus", "هیچ کلیدواژه یا نماد معتبری وارد نشده است", "warning");
                model.Results = new List<ReportResult>();
                return View("Results", model);
            }

            // Create a list to hold all search results
            var allResults = new List<ReportResult>();

            // Dictionary to track which symbol and keyword produced which results
            var searchResultsMap = new Dictionary<string, List<ReportResult>>();

            // گزارش تعداد نمادها و کلیدواژه‌ها
            await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال جستجوی {keywords.Count} کلیدواژه برای {symbols.Count} نماد");

            var totalCombinations = symbols.Count * keywords.Count;
            var completedCombinations = 0;

            // Process each symbol and keyword combination
            foreach (var symbol in symbols)
            {
                foreach (var keyword in keywords)
                {
                    // گزارش وضعیت: در حال جستجوی هر نماد و کلیدواژه
                    await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال جستجوی نماد {symbol} با کلیدواژه {keyword}");

                    // Call the scraper for each individual symbol and keyword combination
                    var results = await _scraper.ScrapeAsync(symbol, keyword);

                    completedCombinations++;
                    await _hubContext.Clients.All.SendAsync("UpdateProgress", (completedCombinations * 100) / totalCombinations);

                    // Add the results to our combined list if there are any
                    if (results != null && results.Any())
                    {
                        // اضافه کردن نام شرکت (سمبل) به هر نتیجه
                        foreach (var result in results)
                        {
                            result.CompanySymbol = symbol;
                        }

                        // Store which results came from which symbol-keyword combination
                        string key = $"{symbol}:{keyword}";
                        searchResultsMap[key] = results;

                        // Add to the overall results list
                        allResults.AddRange(results);

                        // گزارش تعداد نتایج یافت شده برای این ترکیب
                        await _hubContext.Clients.All.SendAsync("UpdateStatus", $"{results.Count} نتیجه برای نماد {symbol} با کلیدواژه {keyword} یافت شد");
                    }
                    else
                    {
                        await _hubContext.Clients.All.SendAsync("UpdateStatus", $"هیچ نتیجه‌ای برای نماد {symbol} با کلیدواژه {keyword} یافت نشد", "info");
                    }
                }
            }

            // Remove duplicates (if the same report appears in multiple searches)
            allResults = allResults.GroupBy(r => r.Link)
                                  .Select(g => g.First())
                                  .ToList();

            // گزارش وضعیت: اتمام جستجو
            await _hubContext.Clients.All.SendAsync("UpdateStatus", $"جستجو با موفقیت انجام شد. {allResults.Count} نتیجه یافت شد", "success");

            // Set the combined results to the model
            model.Results = allResults;

            // Store search data in ViewData for use in the view
            ViewData["SearchResultsMap"] = searchResultsMap;
            ViewData["SearchedSymbols"] = symbols;
            ViewData["SearchedKeywords"] = keywords;

            // Return the results view with the populated model
            return View("Results", model);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "خطا در جستجوی کدال");
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "خطا در جستجوی کدال: " + ex.Message, "error");
            return RedirectToAction("Error", new { message = "خطا در جستجوی کدال", details = ex.Message });
        }
    }
    #endregion

    #region دانلود و ذخیره اکسل های اطلاعات و صورت های مالی
    [HttpPost]
    public async Task<IActionResult> DownloadExcelFiles(List<string> excelLinks, List<string> titles, List<string> companySymbols, string searchKeywords)
    {
        try
        {
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "شروع فرآیند دانلود و پردازش فایل‌های اکسل");

            // فیلتر کردن فایل‌هایی که در عنوان آنها "اطلاعات و صورت های مالی" یا "صورت های مالی" وجود دارد
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "در حال فیلتر کردن فایل‌های مالی از لیست");

            var filteredExcelLinks = new List<string>();
            var filteredTitles = new List<string>();
            var filteredCompanySymbols = new List<string>(); // لیست نام شرکت‌های فیلتر شده

            for (int i = 0; i < titles.Count; i++)
            {
                var normalizedTitle = await _excelDownloader.NormalizePersian(titles[i]);

                if (i < excelLinks.Count &&
                    (normalizedTitle.Contains("اطلاعات و صورتهای مالی") || normalizedTitle.Contains("صورتهای مالی")))
                {
                    filteredExcelLinks.Add(excelLinks[i]);
                    filteredTitles.Add(titles[i]); // یا normalizedTitle بسته به نیاز
                                                   // اضافه کردن نام شرکت به لیست فیلتر شده
                    if (i < companySymbols.Count)
                    {
                        filteredCompanySymbols.Add(companySymbols[i]);
                    }
                    else
                    {
                        // اگر نام شرکت برای این ردیف موجود نبود، یک مقدار پیش‌فرض قرار می‌دهیم
                        filteredCompanySymbols.Add("نامشخص");
                    }
                }
            }

            // اگر هیچ فایلی با عنوان مورد نظر یافت نشد
            if (!filteredExcelLinks.Any())
            {
                await _hubContext.Clients.All.SendAsync("UpdateStatus", "هیچ فایل اطلاعات و صورت های مالی یافت نشد", "warning");
                return RedirectToAction("Error", new
                {
                    title = "فایلی یافت نشد",
                    message = "هیچ فایل اطلاعات و صورت های مالی یافت نشد."
                });
            }

            await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال دانلود {filteredExcelLinks.Count} فایل اکسل");

            var downloadedPaths = await DownloadExcelsWithProgress(filteredExcelLinks, filteredTitles, filteredCompanySymbols);

            if (!downloadedPaths.Any())
            {
                await _hubContext.Clients.All.SendAsync("UpdateStatus", "هیچ فایلی دانلود نشد", "error");
                return RedirectToAction("Error", new
                {
                    title = "دانلود ناموفق",
                    message = "هیچ فایلی دانلود نشد."
                });
            }

            await _hubContext.Clients.All.SendAsync("UpdateStatus", "در حال پردازش کلیدواژه‌ها");

            var keywords = searchKeywords?
                .Replace('،', ',')
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim())
                .ToList() ?? new List<string>();

            await _hubContext.Clients.All.SendAsync("UpdateStatus", $"استخراج جداول بر اساس {keywords.Count} کلیدواژه");

            var matchedTables = new List<MatchedTable>();

            for (int i = 0; i < downloadedPaths.Count; i++)
            {
                var path = downloadedPaths[i];
                var companySymbol = i < filteredCompanySymbols.Count ? filteredCompanySymbols[i] : "نامشخص";
                var fileName = Path.GetFileName(path);

                await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال تبدیل {fileName} به HTML");
                var html = await _htmlParser.ConvertToHtmlAsync(path);

                await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال جستجوی کلیدواژه‌ها در {fileName}");
                var matches = _htmlParser.ExtractTablesAfterKeywords(html, keywords, fileName, companySymbol);
                matchedTables.AddRange(matches);

                // گزارش پیشرفت
                await _hubContext.Clients.All.SendAsync("UpdateProgress", ((i + 1) * 100) / downloadedPaths.Count);
            }

            await _hubContext.Clients.All.SendAsync("UpdateStatus", $"{matchedTables.Count} جدول منطبق با کلیدواژه‌ها پیدا شد");
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "در حال تمیزسازی و پردازش جداول");

            var allCleanedRows = new List<CleanedTableRowViewModel>();
            foreach (var table in matchedTables)
            {
                var cleaned = _htmlParser.ParseHtmlTable(table.HtmlTable, table.Keyword, table.FileName, table.CompanySymbol);
                allCleanedRows.AddRange(cleaned);
            }

            await _hubContext.Clients.All.SendAsync("UpdateStatus", "در حال ذخیره‌سازی جداول در پایگاه داده");
            await _excelTableService.SaveMatchedTablesAsync(allCleanedRows);

            await _hubContext.Clients.All.SendAsync("UpdateStatus", "عملیات با موفقیت انجام شد", "success");

            return View("SelectTables", matchedTables);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "خطا در دانلود فایل‌های اکسل");
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "خطا در پردازش فایل‌ها: " + ex.Message, "error");
            return RedirectToAction("Error", new
            {
                title = "خطا در پردازش فایل‌ها",
                message = "مشکلی در دانلود یا پردازش فایل‌های اکسل رخ داده است.",
                details = ex.Message
            });
        }
    }

    // متد کمکی برای دانلود اکسل‌ها با گزارش پیشرفت
    private async Task<List<string>> DownloadExcelsWithProgress(List<string> excelLinks, List<string> titles, List<string> companySymbols)
    {
        var downloadedPaths = new List<string>();

        for (int i = 0; i < excelLinks.Count; i++)
        {
            var title = i < titles.Count ? titles[i] : $"فایل {i + 1}";
            await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال دانلود فایل {i + 1} از {excelLinks.Count}: {title}");

            try
            {
                var path = await _excelDownloader.DownloadExcelAsync(excelLinks[i], titles[i],
                    i < companySymbols.Count ? companySymbols[i] : "نامشخص");

                if (!string.IsNullOrEmpty(path))
                {
                    downloadedPaths.Add(path);
                    await _hubContext.Clients.All.SendAsync("UpdateStatus", $"دانلود {title} با موفقیت انجام شد");
                }
                else
                {
                    await _hubContext.Clients.All.SendAsync("UpdateStatus", $"دانلود {title} ناموفق بود", "warning");
                }
            }
            catch (Exception ex)
            {
                await _hubContext.Clients.All.SendAsync("UpdateStatus", $"خطا در دانلود {title}: {ex.Message}", "error");
                _logger.LogError(ex, $"خطا در دانلود فایل {title}");
            }

            // گزارش پیشرفت
            await _hubContext.Clients.All.SendAsync("UpdateProgress", ((i + 1) * 100) / excelLinks.Count);
        }

        return downloadedPaths;
    }
    #endregion

    #region گزارش فعالیت ماهانه
    [HttpPost]
    public async Task<IActionResult> ExtractInfoPage(List<string> link, List<string> titles, List<string> companySymbols)
    {
        try
        {
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "شروع استخراج گزارش‌های فعالیت ماهانه");

            var viewModelList = new List<MonthlyActivityTableViewModel>();
            int successCount = 0;
            int totalReports = titles.Count;
            int processedReports = 0;


            for (int i = 0; i < titles.Count; i++)
            {
                try
                {
                    var title = titles[i];
                    processedReports++;

                    if (!title.Contains("گزارش فعالیت ماهانه"))
                    {
                        await _hubContext.Clients.All.SendAsync("UpdateStatus", $"رد کردن {title} (گزارش فعالیت ماهانه نیست)");
                        continue;
                    }

                    await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال پردازش {title} ({processedReports} از {totalReports})");

                    var tableHtml = await _scraper.GetRenderedHtmlAsync(link[i]);

                    if (string.IsNullOrEmpty(tableHtml) || !tableHtml.Contains("<table"))
                    {
                        await _hubContext.Clients.All.SendAsync("UpdateStatus", $"هیچ جدولی در {title} یافت نشد", "warning");
                        continue;
                    }

                    await _hubContext.Clients.All.SendAsync("UpdateStatus", $"در حال تجزیه جدول {title}");
                    var (headers, rows) = _htmlParser.ParseMonthlyActivityTable(tableHtml);

                    if (headers.Count > 0 && rows.Count > 0)
                    {
                        string companySymbol = i < companySymbols.Count ? companySymbols[i] : "";

                        viewModelList.Add(new MonthlyActivityTableViewModel
                        {
                            Title = title,
                            Link = link[i],
                            Headers = headers,
                            Rows = rows,
                            RawHtml = tableHtml, // Optionally store raw HTML for debugging
                            CompanySymbol = companySymbol
                        });

                        successCount++;
                        await _hubContext.Clients.All.SendAsync("UpdateStatus", $"{title} با موفقیت پردازش شد");
                    }
                    else
                    {
                        await _hubContext.Clients.All.SendAsync("UpdateStatus", $"جدول معتبری در {title} یافت نشد", "warning");
                    }

                    // گزارش پیشرفت
                    await _hubContext.Clients.All.SendAsync("UpdateProgress", (processedReports * 100) / totalReports);
                }
                catch (Exception ex)
                {
                    // Log the error but continue processing other items
                    _logger.LogWarning(ex, $"خطا در پردازش {titles[i]}");
                    await _hubContext.Clients.All.SendAsync("UpdateStatus", $"خطا در پردازش {titles[i]}: {ex.Message}", "error");
                }
            }

            if (viewModelList.Any())
            {
                try
                {
                    // Save to database
                    await _hubContext.Clients.All.SendAsync("UpdateStatus", "در حال ذخیره‌سازی گزارش‌ها در پایگاه داده");
                    await _excelTableService.SaveMonthlyReportAsync(viewModelList);
                    await _hubContext.Clients.All.SendAsync("UpdateStatus", $"{successCount} گزارش با موفقیت ذخیره شد", "success");
                    TempData["Message"] = $"{successCount} گزارش با موفقیت پردازش و در پایگاه داده ذخیره شد.";
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "خطا در ذخیره‌سازی در پایگاه داده");
                    await _hubContext.Clients.All.SendAsync("UpdateStatus", "خطا در ذخیره‌سازی در پایگاه داده: " + ex.Message, "error");
                    return RedirectToAction("Error", new
                    {
                        title = "خطا در ذخیره‌سازی",
                        message = $"خطا در ذخیره سازی در پایگاه داده.",
                        details = ex.Message
                    });
                }
            }
            else
            {
                await _hubContext.Clients.All.SendAsync("UpdateStatus", "هیچ گزارش فعالیت ماهانه‌ای استخراج نشد", "warning");
                TempData["Warning"] = "هیچ گزارش فعالیت ماهانه‌ای استخراج نشد.";
            }

            TempData["Message"] = $"{successCount} گزارش با موفقیت پردازش شد.";
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "عملیات با موفقیت انجام شد", "success");
            return View("MonthlyActivityTable", viewModelList);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "خطا در استخراج صفحه اطلاعات");
            await _hubContext.Clients.All.SendAsync("UpdateStatus", "خطا در استخراج اطلاعات: " + ex.Message, "error");
            return RedirectToAction("Error", new
            {
                title = "خطا در استخراج اطلاعات",
                message = "مشکلی در پردازش گزارش‌های فعالیت ماهانه رخ داده است.",
                details = ex.Message
            });
        }
    }
    #endregion

    #region صفحه خطا
    [HttpGet]
    public IActionResult Error(string title = null, string message = null, string details = null, bool showDetails = false)
    {
        var model = new ErrorViewModel
        {
            Title = string.IsNullOrEmpty(title) ? "خطایی رخ داده است" : title,
            Message = message,
            ErrorDetails = details,
            ShowDetails = showDetails
        };

        // در محیط توسعه جزئیات خطا را نمایش می‌دهیم
#if DEBUG
        model.ShowDetails = true;
#endif

        return View(model);
    }
    #endregion
}