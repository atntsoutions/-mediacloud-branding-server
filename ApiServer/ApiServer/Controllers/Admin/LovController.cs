using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;

using DataBase;
using BLAdmin;

namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Admin/Lov")]
    public class LovController : ApiController
    {
        [Route("Lov")]
        [HttpPost]
        public IHttpActionResult Lov(Dictionary<string, object> SearchData)
        {
            try
            {
                using (LovService obj = new LovService())
                    return Ok(obj.Lov(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [Route("SearchRecord")]
        [HttpPost]
        public IHttpActionResult SearchRecord(Dictionary<string, object> SearchData)
        {
            try
            {
                using (LovService obj = new LovService())
                    return Ok(obj.SearchRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [Route("List")]
        [HttpPost]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (LovService obj = new LovService())
                    return Ok(obj.List(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


    }
}
