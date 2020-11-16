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
    [RoutePrefix("api/Pim/Campaign")]
    public class CampaignController : ApiController
    {
        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                using (CampaignService obj = new CampaignService())
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
                using (CampaignService obj = new CampaignService())
                    return Ok(obj.GetRecord(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        

        [HttpPost]
        [Route("Save")]
        public IHttpActionResult Save(Campaign Record)
        {
            try
            {
                using (CampaignService obj = new CampaignService())
                    return Ok(obj.Save(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("RunCampaign")]
        public IHttpActionResult RunCampaign(Campaign Record)
        {
            try
            {
                using (CampaignService obj = new CampaignService())
                    return Ok(obj.RunCampaign(Record));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("ResetCampaign")]
        public IHttpActionResult ResetCampaign(Campaign Record)
        {
            try
            {
                using (CampaignService obj = new CampaignService())
                    return Ok(obj.ResetCampaign(Record));
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
                using (CampaignService obj = new CampaignService())
                    return Ok(obj.Delete(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }
        


    }
}
