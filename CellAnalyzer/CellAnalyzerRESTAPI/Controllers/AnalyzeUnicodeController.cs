using Microsoft.AspNetCore.Mvc;

namespace CellAnalyzerRESTAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AnalyzeUnicodeController : ControllerBase
    {
        [HttpGet]
        public ActionResult<string> AnalyzeUnicode(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(value);
        }
    }
}
