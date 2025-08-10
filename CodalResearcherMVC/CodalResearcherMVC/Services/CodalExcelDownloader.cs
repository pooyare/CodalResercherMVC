namespace CodalResearcherMVC.Services
{
    public class CodalExcelDownloader
    {
        private readonly string _excelDirectory;

        public CodalExcelDownloader()
        {
            _excelDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Excels");
            Directory.CreateDirectory(_excelDirectory);
        }

        // متد جدید برای دانلود تکی فایل اکسل
        public async Task<string> DownloadExcelAsync(string excelLink, string title, string companySymbol = "")
        {
            if (string.IsNullOrWhiteSpace(excelLink) || excelLink == "ندارد")
                return null;

            try
            {
                using var client = new HttpClient();
                var bytes = await client.GetByteArrayAsync(excelLink);

                // امن‌سازی نام فایل
                var safeTitle = string.Join("_", title.Split(Path.GetInvalidFileNameChars()));
                string fileName = $"{safeTitle}.xls";
                string filePath = Path.Combine(_excelDirectory, fileName);

                await File.WriteAllBytesAsync(filePath, bytes);
                Console.WriteLine($"فایل {title} با موفقیت دانلود شد.");
                return filePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"خطا در دانلود {title}: {ex.Message}");
                return null;
            }
        }

        public async Task<List<string>> DownloadExcelsAsync(List<string> excelLinks, List<string> titles, List<string> companySymbols)
        {
            // پاک کردن محتوای پوشه excels قبل از دانلود فایل‌های جدید
            try
            {
                if (Directory.Exists(_excelDirectory))
                {
                    // حذف تمام فایل‌های موجود در پوشه
                    foreach (var file in Directory.GetFiles(_excelDirectory))
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    // ایجاد پوشه اگر وجود نداشته باشد
                    Directory.CreateDirectory(_excelDirectory);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"خطا در پاک کردن پوشه excels: {ex.Message}");
            }

            var downloadedPaths = new List<string>();

            for (int i = 0; i < excelLinks.Count; i++)
            {
                string link = excelLinks[i];
                string title = titles[i];
                string companySymbol = companySymbols[i];

                if (string.IsNullOrWhiteSpace(link) || link == "ندارد")
                    continue;

                try
                {
                    using var client = new HttpClient();
                    var bytes = await client.GetByteArrayAsync(link);

                    // امن‌سازی نام فایل
                    var safeTitle = string.Join("_", title.Split(Path.GetInvalidFileNameChars()));
                    string fileName = $"{safeTitle}.xls";
                    string filePath = Path.Combine(_excelDirectory, fileName);

                    await File.WriteAllBytesAsync(filePath, bytes);
                    downloadedPaths.Add(filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"خطا در دانلود {title}: {ex.Message}");
                }
            }

            return downloadedPaths;
        }

        public async Task<string>  NormalizePersian(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            return input
                .Replace("ي", "ی")  // عربی به فارسی
                .Replace("ك", "ک")
                .Replace("‌", "")   // حذف نیم‌فاصله (U+200C)
                .Trim();
        }

    }
}
