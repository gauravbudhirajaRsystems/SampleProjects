using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace WordAnalyzerRestApi.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class WordAnalyzerController : ControllerBase
    {

        [HttpPost("document")]
        public ActionResult<string> SendDocument(dynamic docBody)
        {
            //var obj = JsonSerializer.Deserialize<dynamic>(docBody);
            if (docBody == null)
            {
                return BadRequest();
            }
            return null;
        }

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
