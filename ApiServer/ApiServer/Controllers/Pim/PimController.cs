using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;

using DataBase;
using BLPim;

namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Pim/Group")]
    public class PimController : ApiController
    {

        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (GroupService obj = new GroupService())
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
                using (GroupService obj = new GroupService())
                    return Ok(obj.GetRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        

        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save(pim_groupm Record)
        {
            try
            {
                using (GroupService obj = new GroupService())
                    return Ok(obj.Save(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("Delete")]
        public IHttpActionResult Delete(Dictionary<string, object> SearchData)
        {
            try
            {
                using (GroupService obj = new GroupService())
                    return Ok(obj.Delete(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


    }
}
