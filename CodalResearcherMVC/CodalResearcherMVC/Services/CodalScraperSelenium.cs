using CodalResearcherMVC.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CodalResearcherMVC.Services
{
    public class CodalScraperSelenium
    {
        public async Task<List<ReportResult>> ScrapeAsync(string symbol, string keyword)
        {
            var results = new List<ReportResult>();
            string searchUrl = $"https://www.codal.ir/ReportList.aspx?search&Symbol={Uri.EscapeDataString(symbol)}";

            var options = new ChromeOptions();
            //برای عدم باز شدن مرورگر chrome
            options.AddArgument("--headless");
            options.AddArgument("--disable-gpu");

            using (var driver = new ChromeDriver(options))
            {
                driver.Navigate().GoToUrl(searchUrl);
                await Task.Delay(3000);

                bool hasNextPage = true;

                while (hasNextPage)
                {
                    var rows = driver.FindElements(By.CssSelector("table.search-result-table tbody tr"));

                    foreach (var row in rows)
                    {
                        var titleElement = row.FindElement(By.CssSelector("td:nth-child(4) a"));
                        var title = titleElement.Text;

                        // اگر عنوان شامل کلمه کلیدی باشد
                        if (title.Contains(keyword))
                        {
                            var link = titleElement.GetAttribute("href");
                            string excel = "ندارد";

                            try
                            {
                                var excelIcon = row.FindElement(By.CssSelector("td:last-child a.icon-excel"));
                                excel = excelIcon.GetAttribute("href");
                            }
                            catch { }

                            results.Add(new ReportResult { Title = title, Link = link, ExcelLink = excel });
                        }
                    }

                    try
                    {
                        var nextBtn = driver.FindElement(By.CssSelector("li[title='صفحه بعدی'] a"));
                        var parent = nextBtn.FindElement(By.XPath(".."));
                        if (!parent.GetAttribute("class").Contains("disabled"))
                        {
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", nextBtn);
                            await Task.Delay(3000);
                        }
                        else
                            hasNextPage = false;
                    }
                    catch
                    {
                        hasNextPage = false;
                    }
                }

                driver.Quit();
            }

            return results;
        }

        private readonly HttpClient _httpClient = new HttpClient();

        public async Task<string> GetRenderedHtmlAsync(string url)
        {
            var options = new ChromeOptions();
            options.AddArgument("--headless");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--window-size=1920,1080");

            using (var driver = new ChromeDriver(options))
            {
                driver.Navigate().GoToUrl(url);

                // Wait for the table to appear
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(d => d.FindElements(By.CssSelector("table.rayanDynamicStatement")).Count > 0);

                // Expand all possible expand buttons (if any)
                var expandButtons = driver.FindElements(By.CssSelector("button.expand-button"));
                foreach (var btn in expandButtons)
                {
                    try
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", btn);
                        await Task.Delay(200);
                        btn.Click();
                        await Task.Delay(500);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error clicking expand button: {ex.Message}");
                    }
                }

                // Wait for the table to be fully loaded
                await Task.Delay(1000);

                // Use JavaScript to get the complete table HTML to ensure we capture all attributes
                string tableHtml = (string)((IJavaScriptExecutor)driver).ExecuteScript(
                    "return document.querySelector('table.rayanDynamicStatement').outerHTML;");

                if (string.IsNullOrEmpty(tableHtml))
                {
                    // Fallback if the specific class isn't found
                    tableHtml = (string)((IJavaScriptExecutor)driver).ExecuteScript(
                        "return document.querySelector('table').outerHTML;");
                }

                return tableHtml;
            }
        }
    }
}
