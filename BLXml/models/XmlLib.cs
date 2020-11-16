using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using System.IO;

namespace BLXml.models
{
    public class XmlLib
    {
        public static string FolderCode = null;
        public static string memberCode = null;
        public static string messageSenderField = null;
        public static string messageRecipientField = null;
        public static string RootFolder = null; // sentfolder
        public static string sentFolder = null; // sentfolder/1001

        public static string Branch_Code = "";
        public static string Company_Code = "";
        public static string Agent_Id = "";
        public static string Agent_Name = "";
        public static string Agent_Code = "";
        public static string MBL_IDS = "";
        public static string HBL_BL_NOS = "";

        public static string FolderId = "";
        public static string report_folder = "";
        public static string File_Name = "";
        public static string File_Type = "XML";
        public static string File_Display_Name = "myreport.xml";
        public static string File_Category = "myreport";
        public static string File_Processid = "1001";

        public static string MessageNumberSeq = "";

        public static Boolean SaveInTempFolder = false;

        public static string PROCESSID = "";

        public static Dictionary<int, string> XmlErrorDic = new Dictionary<int, string>();

        private static string GetNewFolder()
        {
            return System.DateTime.Now.ToString("yMMddHHmmss") + "\\";
        }
        private static string GetDayFolder()
        {
            return System.DateTime.Now.ToString("dd-MMM-yyyy").ToUpper() + "\\";
        }
        public static void CreateSentFolder()
        {
            try
            {
                if (sentFolder == null)
                {
                    sentFolder = RootFolder + GetNewFolder();
                    Directory.CreateDirectory(sentFolder);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static void CreateDayFolder()
        {
            try
            {
                if (sentFolder == null)
                {
                    sentFolder = RootFolder + GetDayFolder();
                    if (!Directory.Exists(sentFolder))
                        Directory.CreateDirectory(sentFolder);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static string GetNewMessageNumber()
        {
            /*
            DataTable Dt = null;
            string sRet = "";
            string sql = "";
            sql = "update edi_allnum set Next_Number = Next_Number + 1 where table_name = 'MESSAGE NUMBER'";
            try
            {
                StoredProcedure.CreateCommand(sql, true);
                StoredProcedure.RunText();
                StoredProcedure.CommitTrans();
                sql = "select * from edi_allnum where table_name = 'MESSAGE NUMBER'";
                Dt = new DataTable(); 
                StoredProcedure.CreateCommand(sql);
                StoredProcedure.Run(Dt);
                if (Dt.Rows.Count > 0)
                {
                    sRet = Dt.Rows[0]["NEXT_NUMBER"].ToString();
                }
            }
            catch (Exception)
            {
                StoredProcedure.RollBackTrans();
                throw;
            }
            return sRet;
            */

            //  return System.DateTime.Now.ToString("yMMddHHmmss"); changed on 01/04/2019

            return String.Concat(System.DateTime.Now.ToString("yMMddHHmmss"), MessageNumberSeq);
        }

        public static DateTime GetCreatedDate()
        {
            DateTime mDate;
            mDate = (DateTime)System.DateTime.Now;
            return mDate.ToUniversalTime();
            //return System.DateTime.Now.ToUniversalTime();   
        }
        public static string GetPortCode(string ml_code)
        {
            string Mcode = ml_code;
            int iLen = 0;
            //if (Lib.Agent_Id == "MOTHER1")
            if (Agent_Name == "MOTHER LINES")
            {
                if (ml_code != null)
                {
                    iLen = ml_code.Length;
                    if (iLen > 3)
                        return Mcode.Substring(2, 3);
                    else
                        return Mcode;
                }
            }
            return ml_code;
        }

        public static void AddToErrorList(string sType, string sRemarks)
        {
            string str = String.Concat(sType, ",", sRemarks);
            XmlErrorDic.Add(XmlErrorDic.Count, str);
        }
    }
}