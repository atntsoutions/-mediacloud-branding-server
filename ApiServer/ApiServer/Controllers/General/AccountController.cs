
using System;
using System.IO;
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
    [RoutePrefix("api/General")]
    public class Accounts_Controller : ApiController
    {
        [HttpGet]
        [Route("GetString")]
        public IHttpActionResult GetString()
        {
            return Ok("Hello");
        }


        [HttpPost]
        [Route("DocumentList")]
        public IHttpActionResult DocumentList(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.DocumentList(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("ExtraList")]
        public IHttpActionResult ExtraList(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.ExtraList(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("CopyFiles")]
        public IHttpActionResult CopyFiles(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.CopyFiles(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }




        [HttpPost]
        [Route("DeleteDocument")]
        public IHttpActionResult DeleteDocument(Dictionary<string, object> SearchData)
        {
            try
            {
                using (UserService obj = new UserService())
                    return Ok(obj.DeleteDocument(SearchData));
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
                    return Ok(obj.LoadDocumentCategory(SearchData));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



        [HttpPost]
        [Route("UploadFiles")]
        public IHttpActionResult UploadFiles()
        {
            Boolean bRet = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            try
            {
                int iLen = 0;
                string sRootFolder = @"D://documents/alldocs/";
                string sLocalFolder = "2018-04";

                Dictionary<string, object> SearchData = new Dictionary<string, object>();


                int ContentLength = 0;
                int iUploadedCnt = 0;
                string sPath = "";
                string Comp_Code = HttpContext.Current.Request.Form["COMPCODE"];

                string Folder_ID = System.Guid.NewGuid().ToString().ToUpper();

                //sRootFolder = @"D://documents/alldocs/";
                //sLocalFolder = "2018-04";

                sRootFolder = HttpContext.Current.Request.Form["ROOT-FOLDER"];
                sLocalFolder = HttpContext.Current.Request.Form["SUB-FOLDER"];



                SearchData.Add("COMP_CODE", Comp_Code);
                SearchData.Add("BRANCH_CODE", HttpContext.Current.Request.Form["BRANCHCODE"]);
                SearchData.Add("PARENT_ID", HttpContext.Current.Request.Form["PARENTID"]);
                SearchData.Add("GROUP_ID", HttpContext.Current.Request.Form["GROUPID"]);
                SearchData.Add("TYPE", HttpContext.Current.Request.Form["TYPE"]);
                SearchData.Add("CATG_ID", HttpContext.Current.Request.Form["CATGID"]);
                SearchData.Add("CREATED_BY", HttpContext.Current.Request.Form["CREATEDBY"]);
                SearchData.Add("PATH", "");
                SearchData.Add("FILENAME", "");

                if (SearchData["GROUP_ID"].ToString() == "MAIL-FTP-ATTACHMENT"|| SearchData["GROUP_ID"].ToString() == "INCREMENT-LETTER")
                {
                    string File_Name = "";
                    string File_Type = "";
                    string File_Display_Name = "";
                    string Category = "";

                    sRootFolder = SearchData["PARENT_ID"].ToString();//set report folder - temp files
                    Category = SearchData["CATG_ID"].ToString();

                    System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;
                    // CHECK THE FILE COUNT.
                    ContentLength = 0;
                    for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                    {
                        System.Web.HttpPostedFile hpf = hfc[iCnt];
                        File_Name = Lib.GetFileName(sRootFolder, Folder_ID, Path.GetFileName(hpf.FileName).Trim().ToUpper().Replace(",", " "));
                        if (hpf.ContentLength > 0)
                        {
                            //SAVE THE FILES IN THE FOLDER.
                            hpf.SaveAs(File_Name);
                            iUploadedCnt = iUploadedCnt + 1;
                            File_Display_Name = Path.GetFileName(hpf.FileName).Trim().ToUpper().Replace(",", " ");
                            ContentLength += hpf.ContentLength;
                        }
                    }
                    RetData.Add("filesize", ContentLength);
                    RetData.Add("category", Category);
                    RetData.Add("filename", File_Name);
                    RetData.Add("filetype", File_Type);
                    RetData.Add("filedisplayname", File_Display_Name);
                }
                else
                {
                    //sPath = System.Web.Hosting.HostingEnvironment.MapPath("~/locker/");
                    sPath = @"{COMP_CODE}/{FOLDER}/{SUBFOLDER_ID}";
                    sPath = sPath.Replace("{COMP_CODE}", Comp_Code);
                    sPath = sPath.Replace("{FOLDER}", sLocalFolder);
                    sPath = sPath.Replace("{SUBFOLDER_ID}", Folder_ID);

                    System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;
                    // CHECK THE FILE COUNT.
                    for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                    {
                        Lib.CreateFolder(sRootFolder + sPath);
                        System.Web.HttpPostedFile hpf = hfc[iCnt];
                        if (hpf.ContentLength > 0)
                        {
                            //SAVE THE FILES IN THE FOLDER.
                            hpf.SaveAs(sRootFolder + sPath + "/" + Path.GetFileName(hpf.FileName).Trim().ToUpper());
                            iUploadedCnt = iUploadedCnt + 1;
                            SearchData["PKID"] = System.Guid.NewGuid().ToString().ToUpper();
                            SearchData["PATH"] = sPath;
                            SearchData["FILENAME"] = Path.GetFileName(hpf.FileName).Trim().ToUpper();
                            if (hpf.ContentLength < 1024)
                                SearchData["SIZE"] = hpf.ContentLength.ToString() + " B";
                            else
                            {
                                iLen = hpf.ContentLength / 1024;
                                SearchData["SIZE"] = iLen.ToString() + " KB" ;
                            }

                            using (UserService obj = new UserService())
                                bRet = obj.SaveDocuments(SearchData);
                        }
                    }
                    // RETURN A MESSAGE.
                    if (iUploadedCnt > 0)
                    {
                        RetData.Add("status", "OK");
                    }
                    else
                    {
                        RetData.Add("status", "FAILED");
                    }
                }
                return Ok(RetData);
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


    }





}
