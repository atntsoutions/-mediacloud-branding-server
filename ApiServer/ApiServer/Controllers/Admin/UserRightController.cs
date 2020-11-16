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
    [RoutePrefix("api/Admin/UserRight")]
    public class UserRightController : ApiController
    {

        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserRightService obj = new UserRightService())
                    return Ok(obj.List(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("RightsList")]
        public IHttpActionResult RightsList(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserRightService obj = new UserRightService())
                    return Ok(obj.RightsList(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        


        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save(UserRights_VM Record)
        {
            try
            {
                using (UserRightService obj = new UserRightService())
                    return Ok(obj.Save(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("CopyRights")]
        public IHttpActionResult CopyRights(UserRights_VM Record)
        {
            try
            {
                using (UserRightService obj = new UserRightService())
                    return Ok(obj.CopyRights(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

    }
}
