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
    [RoutePrefix("api/Admin/Module")]
    public class ModuleController : ApiController
    {

        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ModuleService obj = new ModuleService())
                    return Ok(obj.List(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

       
        [HttpPost]
        [Route("GetRecord")]
        public IHttpActionResult GetRecord(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ModuleService obj = new ModuleService())
                    return Ok(obj.GetRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save(Modulem Record)
        {
            try
            {
                using (ModuleService obj = new ModuleService())
                    return Ok(obj.Save(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


    }
}
