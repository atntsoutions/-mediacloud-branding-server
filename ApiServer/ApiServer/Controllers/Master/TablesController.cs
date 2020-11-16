using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;

using DataBase;
using BLMaster;

namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Master/Tablesm")]
    public class PimTablesController : ApiController
    {

        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (TableService obj = new TableService())
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
                using (TableService obj = new TableService())
                    return Ok(obj.GetRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save(tablesm Record)
        {
            try
            {
                using (TableService obj = new TableService())
                    return Ok(obj.Save(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("SaveDetail")]
        public IHttpActionResult SaveDetail(tablesd Record)
        {
            try
            {
                using (TableService obj = new TableService())
                    return Ok(obj.SaveDetail(Record));
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
                using (TableService obj = new TableService())
                    return Ok(obj.Delete(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("DeleteDetail")]
        public IHttpActionResult DeleteDetail(Dictionary<string, object> SearchData)
        {
            try
            {
                using (TableService obj = new TableService())
                    return Ok(obj.DeleteDetail(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("DeleteTable")]
        public IHttpActionResult DeleteTable(Dictionary<string, object> SearchData)
        {
            try
            {
                using (TableService obj = new TableService())
                    return Ok(obj.DeleteTable(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



    }
}
