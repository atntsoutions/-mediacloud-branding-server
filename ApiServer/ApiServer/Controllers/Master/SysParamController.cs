using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;

using DataBase;
using BLMaster;

//check itsm sdf


namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Master/SysParam")]
    public class SysParamController : ApiController
    {

        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (SysParamService obj = new SysParamService())
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
                using (SysParamService obj = new SysParamService())
                    return Ok(obj.GetRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save(paramvalues_vm Record)
        {
            try
            {
                using (SysParamService obj = new SysParamService())
                    return Ok(obj.Save(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

    }
}
