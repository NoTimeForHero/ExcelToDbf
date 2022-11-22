using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services.Preload;
using Newtonsoft.Json;

namespace ExcelToDbf.Core.Services
{
    public interface IWebService
    {
        Task<T> Get<T>(string url, CancellationToken? token = null);
        Task<string> GetFile(string url, CancellationToken? token = null);
    }

    internal class WebService : IDisposable, IWebService
    {
        private readonly HttpClient client = new HttpClient();

        public async Task<T> Get<T>(string url, CancellationToken? token = null)
        {
            var json = await GetFile(url, token);
            try
            {
                return JsonConvert.DeserializeObject<T>(json);
            }
            catch (FormatException ex)
            {
                throw new Exception("", ex);
            }
        }

        public async Task<string> GetFile(string url, CancellationToken? token = null)
        {
            var response = await client.GetAsync(url, token ?? new CancellationTokenSource().Token);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public void Dispose()
        {
            client.Dispose();
        }
    }
}
