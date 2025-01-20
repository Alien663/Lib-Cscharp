using Microsoft.Extensions.Configuration;
using System.Reflection;

namespace Config.Extention
{
    public class AppConfig
    {
        private static string _filename = "appsettings.json";

        public AppConfig(string filename) 
        {
            _filename = filename;
        }

        public static IConfigurationRoot Config => LazyConfig.Value;

        private static readonly Lazy<IConfigurationRoot> LazyConfig = new Lazy<IConfigurationRoot>(() => new ConfigurationBuilder()
            .SetBasePath(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!)
            .AddJsonFile(_filename)
            .Build());
    }
}
