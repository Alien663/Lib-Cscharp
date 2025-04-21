using Microsoft.Extensions.Configuration;
using System.Reflection;

namespace Alien.Common.Config;

public class AppConfig
{
    private string _filename;
    private Lazy<IConfiguration> LazyConfig;
    public IConfiguration Configuration { get { return LazyConfig.Value; } }

    public AppConfig(string filename = "appsettings.json")
    {
        _filename = filename;
        LazyConfig = new Lazy<IConfiguration>(() =>
        {
            return new ConfigurationBuilder()
                .SetBasePath(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!)
                .AddJsonFile(_filename)
                .Build();
        });
    }
}
