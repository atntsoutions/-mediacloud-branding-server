using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;

using DataBase;
using BLAdmin;

namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Admin/Company")]
    public class CompanyController : ApiController
    {
        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using ( CompanyService obj = new CompanyService())
                    return Ok(obj.List(SearchData));
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
                using (CompanyService obj = new CompanyService())
                    return Ok(obj.LoadDefault(SearchData));
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
                using (CompanyService obj = new CompanyService())
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
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {
                var model = HttpContext.Current.Request.Form["record"];
                Companym record = JsonConvert.DeserializeObject<Companym>(model);

                using (CompanyService obj = new CompanyService())
                {
                    RetData  = obj.Save(record);
                }

                System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;

                for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                {
                    Lib.UploadFile(hfc[iCnt], record._globalvariables.comp_code, record.comp_pkid, hfc.Keys[iCnt]);
                }

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
                using (CompanyService obj = new CompanyService())
                    return Ok(obj.Delete(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("LoadUserStore")]
        public IHttpActionResult LoadUserStore(Dictionary<string, object> SearchData)
        {
            try
            {
                using (CompanyService obj = new CompanyService())
                    return Ok(obj.LoadUserStore(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("SaveUserStore")]
        public IHttpActionResult SaveUserStore(Companym Record)
        {
            try
            {
                using (CompanyService obj = new CompanyService())
                    return Ok(obj.SaveUserStore(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("ListApproval")]
        public IHttpActionResult ListApproval(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {
                using (CompanyService obj = new CompanyService())
                    return Ok(obj.ListApproval(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }

        }



        [HttpPost]
        [Route("SaveApproval")]
        public IHttpActionResult SaveApproval(approvald Record )
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {
                using (CompanyService obj = new CompanyService())
                    return Ok(obj.SaveApproval(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }

        }





    }
}
