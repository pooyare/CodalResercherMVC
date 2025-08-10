using Microsoft.AspNetCore.SignalR;
using System.Threading.Tasks;

namespace CodalResearcherMVC.Services
{
    public class ProcessingHub : Hub
    {
        public async Task UpdateStatus(string message, string type = "info")
        {
            await Clients.All.SendAsync("UpdateStatus", message, type);
        }

        public async Task UpdateProgress(int percentage)
        {
            await Clients.All.SendAsync("UpdateProgress", percentage);
        }
        // اضافه کردن متد جدید برای اطلاع از تکمیل پردازش همراه با URL انتقال
        public async Task NotifyProcessingComplete(bool success, string message, string redirectUrl = null)
        {
            await Clients.All.SendAsync("ProcessingComplete", success, message, redirectUrl);
        }
    }
}
