using Microsoft.Extensions.Configuration;
using OutlookScheduleGetter.Domains;

namespace OutlookScheduleGetter
{
    /// <summary>
    /// Program
    /// </summary>
    internal class Program
    {
        private static void Main(string[] args)
        {
            var configuration = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            var appSettings = configuration.Get<AppSettings>();
        }
    }
}
