using Microsoft.AspNetCore.Mvc;

namespace WordAnalyzerRestApi.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class WordAnalyzerController : ControllerBase
    {
        [HttpGet("unicode")]
        public ActionResult<string> GetUnicode(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return SharedCodeWordLibrary.WordOperations.GetUnicode(value);
        }

        [HttpGet("charcount")]
        public ActionResult<string> GetCharCount(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return SharedCodeWordLibrary.WordOperations.GetCharCount(value);
        }

        [HttpGet("wordcount")]
        public ActionResult<string> GetWordCount(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return SharedCodeWordLibrary.WordOperations.GetWordCount(value);
        }
    }
}
