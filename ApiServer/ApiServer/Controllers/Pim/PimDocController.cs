using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;

using System.Diagnostics;

using DataBase;
using BLPim;

namespace WebApiServer.Controllers
{

    [Authorize]
    [RoutePrefix("api/Pim/Doc")]
    public class PimDocController : ApiController
    {


        [HttpPost]
        [Route("List")]
        public IHttpActionResult List(Dictionary<string, object> SearchData)
        {
            try
            {
                string ServerImageURL = Lib.GetSeverImageURL(SearchData["comp_code"].ToString());
                using (DocService obj = new DocService())
                    return Ok(obj.List(SearchData, ServerImageURL));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }


        [HttpPost]
        [Route("Download")]
        public IHttpActionResult Download(Dictionary<string, object> SearchData)
        {
            try
            {
                string ServerReportPath = Lib.GetReportPath(SearchData["comp_code"].ToString());
                string ServerImagePath = Lib.GetImagePath(SearchData["comp_code"].ToString());
                using (DocService obj = new DocService())
                    return Ok(obj.Download(SearchData, ServerImagePath, ServerReportPath));
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
                string ServerImageURL = Lib.GetSeverImageURL(SearchData["comp_code"].ToString());
                using (DocService obj = new DocService())
                    return Ok(obj.GetRecord(SearchData, ServerImageURL));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }

        [HttpPost]
        [Route("LoadDefault")]
        public IHttpActionResult LoadDefault (Dictionary<string, object> SearchData)
        {
            try
            {
                using (DocService obj = new DocService())
                    return Ok(obj.LoadDefault(SearchData));
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
                string Folder = "";
                string slno = "";
                imageTools tools = new imageTools();

                Boolean isImageFile = false;
                Boolean isPdfFile = false;

                DataTable Dt_Record = new DataTable();
                DataTable Dt_Records = new DataTable();

                string uploadError = "";

                Dictionary<string, object> SearchData = new Dictionary<string, object>();

                Dictionary<string, string> Files2Remove = new Dictionary<string, string>();

                var model = HttpContext.Current.Request.Form["record"];
                pim_docm record = JsonConvert.DeserializeObject<pim_docm>(model);

                var columnList = HttpContext.Current.Request.Form["records"];
                tablesd [] records  = JsonConvert.DeserializeObject<tablesd[]>(columnList);


                string ServerImageURL = Lib.GetSeverImageURL(record._globalvariables.comp_code);
                string ServerReportPath = Lib.GetReportPath(record._globalvariables.comp_code);
                string ServerImagePath = Lib.GetImagePath(record._globalvariables.comp_code);

                using (DocService obj = new DocService())
                {
                    Dt_Record = obj.getDataTableRecord(record.doc_pkid);
                    RetData = obj.Save(record, records, ServerImageURL);
                    slno = RetData["slno"].ToString();
                }

                // first save to report/temp folder
                System.Web.HttpFileCollection hfc = System.Web.HttpContext.Current.Request.Files;
                string REPID = System.Guid.NewGuid().ToString().ToUpper();
                ServerReportPath = Path.Combine(ServerReportPath, System.DateTime.Now.ToString("yyyy-MM-dd"), REPID);
                string smallname = "";
                for (int iCnt = 0; iCnt <= hfc.Count - 1; iCnt++)
                {
                    isImageFile = false;
                    isPdfFile = false;
                    if ( record.doc_file_name.ToUpper() == hfc[iCnt].FileName.ToUpper())
                    {
                        if (hfc[iCnt].ContentType.ToString().ToUpper().StartsWith("IMAGE"))
                        {
                            isImageFile = true;
                            smallname = "ts.jpg";
                        }
                        if (Path.GetExtension(hfc[iCnt].FileName).ToUpper() == ".PDF") {
                            isPdfFile = true;
                            smallname = "ts.jpg";
                        }
                    }
                    SaveFile(ServerReportPath, smallname, hfc[iCnt], isImageFile, isPdfFile, record._globalvariables.comp_code);
                }

                if (smallname != "")
                {
                    using (DocService obj = new DocService())
                    {
                        obj.UpdateDocFileName(record, "doc_thumbnail", smallname);
                        RetData["thumbnail"] = smallname;
                    }
                }
                
                //then copy to original folder
                string sError = "";
                if (record.doc_file_name.ToString().Trim().Length >0)
                {
                    Folder = Lib.getPath(ServerImagePath, record._globalvariables.comp_code, record.doc_table_name, record.doc_slno.ToString(), true);
                    sError = Lib.CopyFile(Path.Combine(ServerReportPath, record.doc_file_name),  Path.Combine(Folder, record.doc_file_name));
                    if ( sError != "")
                    {
                        uploadError += "\n" + sError;
                        using (DocService obj = new DocService())
                        {
                            obj.UpdateDocFileName(record, "doc_file_name");
                            obj.UpdateDocFileName(record, "doc_thumbnail", "");
                        }
                    }
                    if (smallname != "")
                    {
                        sError = Lib.CopyFile(Path.Combine(ServerReportPath, smallname), Path.Combine(Folder, smallname));
                    }

                }

                foreach (tablesd mRow in records)
                {
                    if ( mRow.tabd_col_type == "FILE")
                    {
                        Folder = Lib.getPath(ServerImagePath, record._globalvariables.comp_code, record.doc_table_name, record.doc_slno.ToString(), true);
                        sError   = Lib.CopyFile(Path.Combine(ServerReportPath, mRow.tabd_col_value), Path.Combine(Folder,mRow.tabd_col_value));
                        if (sError != "")
                        {
                            uploadError += "\n" + sError;
                            using (DocService obj = new DocService())
                            {
                                obj.UpdateDocFileName(record, "COL_" + mRow.tabd_col_name);
                            }
                        }
                    }
                }

                RetData.Add("uploaderror", uploadError);
                
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
                string ServerReportPath = Lib.GetReportPath(SearchData["comp_code"].ToString());
                string ServerImagePath = Lib.GetImagePath(SearchData["comp_code"].ToString());

                using (DocService obj = new DocService())
                    return Ok(obj.Delete(SearchData, ServerImagePath));
            }
            catch (Exception Ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.BadRequest, Ex.Message.ToString()));
            }
        }



        private string SaveFile(string Folder, string tmbname, System.Web.HttpPostedFile hpf, Boolean isImageFile,Boolean isPdfFile, string comp_code )
        {
            string retval = "";
            string FileName = "";
            string thumbNail = "";
            imageTools tools = new imageTools();

            try
            {
                Lib.CreateFolder(Folder);

                FileName = Path.Combine(Folder, Path.GetFileName(hpf.FileName));
                hpf.SaveAs(FileName);
                
                if (isImageFile)
                {
                    thumbNail = Path.Combine(Folder, tmbname);
                    tools.Save(FileName, thumbNail);
                }
                if ( isPdfFile)
                {
                    thumbNail = Path.Combine(Folder, tmbname);
                    SavePdf2Image(FileName, thumbNail, comp_code);
                }
            }
            catch (Exception Ex)
            {
                retval = Ex.Message.ToString();
            }

            return retval;

        }

        Boolean SavePdf2Image(string sourcefile, string targetFile, string comp_code)
        {
            Boolean bRet = true;
            try
            {
                string gsexe = Lib.getSettings(comp_code, "GS-LOCATION", "NAME");

                //gswin64c -dNOPAUSE -sDEVICE=jpeg -dPDFFitPage=true -r96 -dDEVICEWIDTHPOINTS=50 -dDEVICEHEIGHTPOINTS=50 -sOutputFile=outputfile.jpg -ffirst.pdf
                string ghostScriptPath = gsexe;

                //if ( gsexe == "")
                  //  ghostScriptPath = @"C:\gs9.52\bin\gswin64c.exe";

                String ars = " -dNOPAUSE -sDEVICE=jpeg  -r96 -dDEVICEWIDTHPOINTS=10 -dDEVICEHEIGHTPOINTS=10 -q -dBATCH " + " -sOutputFile=" + targetFile + " -f" + sourcefile;
                Process proc = new Process();
                proc.StartInfo.FileName = ghostScriptPath;
                proc.StartInfo.Arguments = ars;
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc.Start();
                proc.WaitForExit();
            } catch ( Exception )
            {
                bRet = false;
            }
            return bRet;
        }




    }






    public class imageTools
    {

        public void SaveFile(string input)
        {
            Image image = Image.FromFile(input);
            image.Save(input);
        }



        public void Save(string input, string output)
        {
            //input = "C:\\background1.png";
            //output = "C:\\thumbnail.png";
            // Load image.

            

            Image image = Image.FromFile(input);
            // Compute thumbnail size.
            Size thumbnailSize = GetThumbnailSize(image);
            // Get thumbnail.
            Image thumbnail = image.GetThumbnailImage(thumbnailSize.Width,
                thumbnailSize.Height, null, IntPtr.Zero);
            // Save thumbnail.
            thumbnail.Save(output);
        }

        Size GetThumbnailSize(Image original)
        {
            // Maximum size of any dimension.
            const int maxPixels = 70;

            // Width and height.
            int originalWidth = original.Width;
            int originalHeight = original.Height;

            // Compute best factor to scale entire image based on larger dimension.
            double factor;
            if (originalWidth > originalHeight)
            {
                factor = (double)maxPixels / originalWidth;
            }
            else
            {
                factor = (double)maxPixels / originalHeight;
            }

            // Return thumbnail size.
            return new Size((int)(originalWidth * factor), (int)(originalHeight * factor));
        }


    }

}
