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

namespace WebApiServer.Controllers
{
    [Authorize]
    [RoutePrefix("api/Pim/spot")]
    public class SpotController : ApiController
    {
        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (SpotService obj = new SpotService())
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
                using (SpotService obj = new SpotService())
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
                pim_spot record = JsonConvert.DeserializeObject<pim_spot>(model);
                using (SpotService obj = new SpotService())
                {
                    RetData = obj.Save(record);
                    islno = Lib.Conv2Integer(RetData["slno"].ToString());

                }
                System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;
                for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                {
                    Lib.UploadFile(hfc[iCnt], record._globalvariables.comp_code, record.spot_pkid, hfc.Keys[iCnt]);
                }
                return Ok(RetData);
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("SaveDet")]
        public IHttpActionResult SaveDet()
        {

            int islno = 0;
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {

                Dictionary<string, object> SearchData = new Dictionary<string, object>();
                var model = HttpContext.Current.Request.Form["record"];
                pim_spotd record = JsonConvert.DeserializeObject<pim_spotd>(model);
                using (SpotService obj = new SpotService())
                {
                    RetData = obj.SaveDet(record);
                    islno = Lib.Conv2Integer(RetData["slno"].ToString());

                }
                System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;
                for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                {
                    Lib.UploadFile(hfc[iCnt], record._globalvariables.comp_code, record.spotd_pkid, hfc.Keys[iCnt]);
                }
                return Ok(RetData);
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



        [HttpPost]
        [Route("Spotmemo")]
        public IHttpActionResult SpotMemo(Dictionary<string, string> SearchData)
        {
            try
            {
                Dictionary<string, object> RetData = new Dictionary<string, object>();
                string pkid = SearchData["pkid"].ToString();
                string comp_code = SearchData["comp_code"].ToString();

                string user_code = SearchData["user_code"].ToString();

                string source = SearchData["source"].ToString();

                string ServerReportPath = Lib.GetReportPath(comp_code);

                string ServerImagePath = Lib.GetImagePath(comp_code);

                string REPID = System.Guid.NewGuid().ToString().ToUpper();

                string report_folder = Path.Combine(ServerReportPath, System.DateTime.Now.ToString("yyyy-MM-dd"), REPID);


                SpotMemo report = new SpotMemo();
                report.pkid = pkid;
                report.imagefolder = ServerImagePath;
                report.comp_code = comp_code;
                report.user_code = user_code;
                report.Report_Caption = "Spot Details";
                report.source = source;
                report.process();

                string File_Display_Name = "spot-" + report.slno + ".pdf";
                string File_Name = Path.Combine(report_folder, File_Display_Name);
                string File_Type = "pdf";

                if (report.ExportList != null)
                {
                    if (Lib.CreateFolder(report_folder))
                    {
                        Export2Pdf mypdf = new Export2Pdf();
                        mypdf.ExportList = report.ExportList;
                        mypdf.FileName = File_Name;
                        mypdf.Page_Height = report.Page_Height;
                        mypdf.Page_Width = report.Page_Width;
                        mypdf.Process();
                    }
                }

                RetData.Add("filename", File_Name);
                RetData.Add("filetype", File_Type);
                RetData.Add("filedisplayname", File_Display_Name);
                RetData.Add("email_to", report.emails_to);
                RetData.Add("email_cc", report.emails_cc);
                RetData.Add("html", report.html);

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
                using (SpotService obj = new SpotService())
                    return Ok(obj.Delete(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("DeleteDet")]
        public IHttpActionResult DeleteDet(Dictionary<string, object> SearchData)
        {
            try
            {
                using (SpotService obj = new SpotService())
                    return Ok(obj.DeleteDet(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



    }
}
