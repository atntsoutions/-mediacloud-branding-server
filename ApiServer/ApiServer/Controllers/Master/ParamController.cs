using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;

using DataBase;
using BLMaster;

using System.IO;

//helo


namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Master/Param")]
    public class ParamController : ApiController
    {
        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ParamService obj = new ParamService())
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
                using (ParamService obj = new ParamService())
                    return Ok(obj.GetRecord(SearchData));
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
                Param record = JsonConvert.DeserializeObject<Param>(model);

                string ServerImageURL = Lib.GetSeverImageURL(record._globalvariables.comp_code);
                string ServerReportPath = Lib.GetReportPath(record._globalvariables.comp_code);
                string ServerImagePath = Lib.GetImagePath(record._globalvariables.comp_code);


                using (ParamService obj = new ParamService())
                {
                    RetData = obj.Save(record);
                    islno =  Lib.Conv2Integer(RetData["slno"].ToString());
                }

                System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;

                
                for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                {
                    Lib.UploadFile(hfc[iCnt], record._globalvariables.comp_code, record.param_pkid, hfc.Keys[iCnt]);
                }

                return Ok(RetData);

            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


   



        [HttpPost]
        [Route("getSettings")]
        public IHttpActionResult getSettings(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.getSettings(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("SaveSettings")]
        public IHttpActionResult SaveSettings(Settings_VM Record)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.SaveSettings(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("DataTransfer")]
        public IHttpActionResult DataTransfer(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.DataTransfer(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("LoadDefault")]
        public IHttpActionResult LoadDefault(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.LoadDefault(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("SaveImportData")]
        public IHttpActionResult SaveImportData(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.SaveImportData(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("UpdateData")]
        public IHttpActionResult UpdateData(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.UpdateData(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("SaveLockings")]
        public IHttpActionResult SaveLockings(Lockingm Record)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.SaveLockings(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("Process")]
        public IHttpActionResult Process(Dictionary<string, object> SearchData)
        {
            try
            {
                using (ParamService obj = new ParamService())
                    return Ok(obj.Process(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }
    }
}
