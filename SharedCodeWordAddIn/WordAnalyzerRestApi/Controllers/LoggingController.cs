using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace WordAnalyzerRestApi.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class LoggingController : ControllerBase
    {
        public LoggingController()
        {
        }

        [HttpGet("log")]
        public void Log(string logMessage, string logType)
        {
            switch (logType)
            {
                case "Information":
                    Serilog.Log.Information(logMessage);
                    break;
                case "Fatal":
                    Serilog.Log.Fatal(logMessage);
                    break;
                case "Verbose":
                    Serilog.Log.Verbose(logMessage);
                    break;
                case "Warning":
                    Serilog.Log.Warning(logMessage);
                    break;
                case "Debug":
                    Serilog.Log.Debug(logMessage);
                    break;
                case "Error":
                    Serilog.Log.Error(logMessage);
                    break;
                default:
                    Serilog.Log.Information(logMessage);
                    break;
            }
        }
    }
}
