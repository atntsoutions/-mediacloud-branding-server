using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;

using DataBase;
using BLPim;

// test

namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Pim/Job")]
    public class JobController : ApiController
    {
        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (JobService obj = new JobService())
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
                using (JobService obj = new JobService())
                    return Ok(obj.GetRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



        [HttpPost]
        [Route("SaveStatus")]
        public IHttpActionResult SaveStatus(pim_spotd Record)
        {
            try
            {
                using (JobService obj = new JobService())
                    return Ok(obj.SaveStatus(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save()
        {

            int islno = 0;
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {

                Dictionary<string, object> SearchData = new Dictionary<string, object>();
                var model = HttpContext.Current.Request.Form["record"];
                pim_spot record = JsonConvert.DeserializeObject<pim_spot>(model);
                using (JobService obj = new JobService())
                {
                    RetData = obj.Save(record);
                    islno = Lib.Conv2Integer(RetData["slno"].ToString());
                }

                /*
                System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;
                for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                {
                    Lib.UploadFile(hfc[iCnt], record._globalvariables.comp_code, record.spot_pkid, hfc.Keys[iCnt]);
                }
                */

                return Ok(RetData);
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

                using (JobService obj = new JobService())
                    return Ok(obj.Delete(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



        [HttpPost]
        [Route("GetRecord_recce_user")]
        public IHttpActionResult GetRecord_recce_user(Dictionary<string, object> SearchData)
        {
            try
            {
                using (JobService obj = new JobService())
                    return Ok(obj.GetRecord_recce_user(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



        [HttpPost]
        [Route("Save_recce_user")]
        public IHttpActionResult Save_recce_user(pim_spot Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {
                using (JobService obj = new JobService())
                    return Ok(obj.Save_recce_user(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }





    }
}
