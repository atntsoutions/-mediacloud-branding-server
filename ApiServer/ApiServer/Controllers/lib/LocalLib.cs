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

    public static class LocalLib
    {
        private static string SaveFile(string Folder, System.Web.HttpPostedFile hpf, string comp_code)
        {
            string retval = "";
            string FileName = "";
            imageTools tools = new imageTools();
            try
            {
                Lib.CreateFolder(Folder);
                FileName = Path.Combine(Folder, Path.GetFileName(hpf.FileName));
                hpf.SaveAs(FileName);
            }
            catch (Exception Ex)
            {
                retval = Ex.Message.ToString();
            }
            return retval;
        }


    }

}