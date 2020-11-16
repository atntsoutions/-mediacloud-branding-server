using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;

namespace BLXml
{
    public class BankEdiService : BL_Base
    {
        private string File_Name = "";
        private string File_Type = "XML";
        private string File_Display_Name = "myreport.txt";
        private string File_Extension = "";

        private string FolderId = "";
        private string report_folder = "";
        private string branch_code = "";
        private string company_code = "";
        private string user_code = "";
        private string SaveMessage = "Generate Complete ";
        private string DisplayMsg = "";
        private string GenerateType = "";
        private int TOTAL = 0;
        private string rowtype = "";
        private string mtlock = "";
        private string custrefno = "";
        private string PKID = "";
        Dictionary<int, string> ErrorDic = new Dictionary<int, string>();

        public Dictionary<string, object> GenerateEdi_Bank(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            report_folder = "";
            branch_code = "";
            SaveMessage = "";
            string Cust_Ref_Uniq_No = "";
            try
            {
                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();
                if (SearchData.ContainsKey("company_code"))
                    company_code = SearchData["company_code"].ToString();
                if (SearchData.ContainsKey("type"))
                    GenerateType = SearchData["type"].ToString();
                if (SearchData.ContainsKey("rowtype"))
                    rowtype = SearchData["rowtype"].ToString();
                if (SearchData.ContainsKey("pkid"))
                    PKID = SearchData["pkid"].ToString();
                if (SearchData.ContainsKey("filedisplayname"))
                    File_Display_Name = SearchData["filedisplayname"].ToString();

                if (SearchData.ContainsKey("user_code"))
                    user_code = SearchData["user_code"].ToString();

                sql = "select mt_cust_cfno from moneytransfer where mt_jv_id ='" + PKID + "'";
                Con_Oracle = new DBConnection();
                Object sVal = Con_Oracle.ExecuteScalar(sql);
                Con_Oracle.CloseConnection();
                Cust_Ref_Uniq_No = sVal.ToString();


                XmlLib.XmlErrorDic = new Dictionary<int, string>();
                XmlLib.Branch_Code = branch_code;
                XmlLib.Company_Code = company_code;

                //string yymmdd = DateTime.Now.ToString("yyyyMMdd");

                //if (rowtype == "CHECK-LIST")
                //    File_Extension = ".CSV";
                //else
                //    File_Extension = ".TXT";

                //File_Display_Name = yymmdd + Cust_Ref_Uniq_No + File_Extension;


                report_folder = @"C://BANK/";

                if (!System.IO.Directory.Exists(report_folder))
                    System.IO.Directory.CreateDirectory(report_folder);

               // File_Name = report_folder + File_Display_Name;
                Init();

                AxisMoneyTransfer atrans = new AxisMoneyTransfer();
                atrans.PKID = PKID;
                atrans.company_code = company_code;
                atrans.branch_code = branch_code;
                atrans.user_code = user_code;
                atrans.InvokeType = GenerateType;
                atrans.report_folder = report_folder;
                atrans.Generate();
                custrefno = atrans.cust_uniq_ref;
                mtlock = atrans.mt_lock;
                File_Name = atrans.File_Name;
                File_Display_Name = atrans.File_Display_Name;
                if (atrans.IsError)
                {
                    throw new Exception(atrans.ErrorMessage);
                }

            }
            catch (Exception Ex)
            {
                throw Ex;
            }

            SaveMessage = "Generate Complete ";
            SaveMessage += "\nCustomer Reference Number : " + custrefno;

            RetData.Add("savemsg", SaveMessage);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("mtlock", mtlock);
            RetData.Add("custrefno", custrefno);
            RetData.Add("processid", XmlLib.PROCESSID);
            return RetData;
        }


        private void Init()
        {

        }


        private void WriteErrorFile()
        {
            if (XmlLib.XmlErrorDic.Count <= 0)
                return;
            bool bOk = false;
            string fName = "";
            StringBuilder StrBld = new StringBuilder();
            for (int i = 0; i < XmlLib.XmlErrorDic.Count; i++)
            {
                bOk = true;
                StrBld.AppendLine();
                StrBld.Append(XmlLib.XmlErrorDic[i]);
            }

            XmlLib.CreateSentFolder();
            fName = XmlLib.sentFolder + "ERRORFILE.CSV";
            if (System.IO.File.Exists(fName))
                System.IO.File.Delete(fName);
            if (bOk)
                System.IO.File.AppendAllText(fName, StrBld.ToString());
        }

    }
}
