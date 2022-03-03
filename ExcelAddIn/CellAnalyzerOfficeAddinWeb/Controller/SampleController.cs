using System;
using System.Web.Http;

namespace CellAnalyzerOfficeAddinWeb.Controller
{
    public class SampleController : ApiController
    {
        public class Response
        {
            public string Status { get; set; }
            public string Message { get; set; }
        }


        [HttpGet()]
        public Response Sample()
        {
            try
            {
                return new Response()
                {
                    Status = "Success Call",
                    Message = "Your Call was Success"
                };
            }
            catch (Exception ex)
            {
                return new Response()
                {
                    Status = "Un-Success Call",
                    Message = ex.Message
                };
            }
        }

    }
}
