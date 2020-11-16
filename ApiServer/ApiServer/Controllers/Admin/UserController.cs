using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;
using System.Web.Http.Cors;

using DataBase;
using BLAdmin;




namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Admin/User")]
    
    public class UserController : ApiController
    {
        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
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
                using (UserService obj = new UserService())
                    return Ok(obj.LoadDefault(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("NewUserDefault")]
        public IHttpActionResult NewUserDeDefault(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.NewUserDefault(SearchData));
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
                using (UserService obj = new UserService())
                    return Ok(obj.GetRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save(User Record)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.Save(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("LoadMenu")]
        public IHttpActionResult LoadMenu(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.LoadMenu(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [AllowAnonymous]
        [HttpPost]
        [Route("LoadCompany")]
        public IHttpActionResult LoadCompany(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.LoadCompany(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("LoadBranch")]
        public IHttpActionResult LoadBranch(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.LoadBranch(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }
        [HttpPost]
        [Route("LoadYear")]
        public IHttpActionResult LoadYear(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.LoadBranch(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [AllowAnonymous]
        [HttpGet]
        [Route("DownloadFile")]
        public HttpResponseMessage DownloadFile(string report_folder, string filename, string filetype, string filedisplayname = "")
        {
            string fextn = "";
            if (filetype == null)
            {
                filetype = "";
            }
            if (filetype.ToUpper() == "EXCEL")
                fextn = ".xls";
            if (filetype.ToUpper() == "PDF")
                fextn = ".pdf";
            if (filetype.ToUpper() == "SB")
                fextn = ".sb";
            if (filetype.ToUpper() == "CSV")
                fextn = ".csv";
            if (filetype.ToUpper() == "XML")
                fextn = ".xml";

            if (filedisplayname == "" || filedisplayname == "N")
            {
                report_folder = System.IO.Path.Combine(report_folder, filename);
                filename = System.IO.Path.Combine(report_folder, filename) + fextn;
            }

            HttpResponseMessage response = new HttpResponseMessage();
            response.StatusCode = HttpStatusCode.OK;

            response.Content = new StreamContent(new FileStream(filename, FileMode.Open, FileAccess.Read));

            if (filetype.ToUpper() == "EXCEL")
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-msexcel");
            if (filetype.ToUpper() == "PDF")
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
            if (filetype.ToUpper() == "XML")
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/xml");

            response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
            if (filedisplayname == "" || filedisplayname == "N")
                response.Content.Headers.ContentDisposition.FileName = "Report" + fextn;
            else
                response.Content.Headers.ContentDisposition.FileName = filedisplayname;

            return response;
        }





    }
}
