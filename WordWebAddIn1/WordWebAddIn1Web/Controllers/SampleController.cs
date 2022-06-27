using System;
using System.Net;
using System.Web.Http;

namespace WordWebAddIn1Web.Controllers
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

        //[HttpGet()]
        //public Response Sample()
        //{
        //    try
        //    {
        //        Ht

        //        return new Response()
        //        {
        //            Status = "Success Call",
        //            Message = "Your Call was Success"
        //        };
        //    }
        //    catch (Exception ex)
        //    {
        //        return new Response()
        //        {
        //            Status = "Un-Success Call",
        //            Message = ex.Message
        //        };
        //    }
        //}
    }
}
