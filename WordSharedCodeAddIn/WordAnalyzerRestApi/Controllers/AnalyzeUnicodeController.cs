using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace WordAnalyzerRestApi.Controllers
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
            return WordSharedCodeLibrary.WordOperations.GetHighlightedTextUnicode(value);
        }
    }
}
