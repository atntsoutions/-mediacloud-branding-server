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
    public class XmlMexicoService : BL_Base
    {
        private string File_Name = "";
        private string File_Type = "XML";
        private string File_Display_Name = "myreport.xml";
        private string File_NameAcK = "";
        private string File_TypeAck = "XML";
        private string File_Display_NameAck = "ack-myreport.xml";
        private string File_Extension = "";
        private string File_ExtensionAck = "";

        private string FolderId = "";
        private string report_folder = "";
        private string branch_code = "";
        private string company_code = "";
        private string SaveMessage = "Generate Complete ";
        private string DisplayMsg = "";
        private string GenerateType = "CNTR";
        private int TOTAL = 0;
        private string rowtype = "";

        private string PKID = "";
        Dictionary<int, string> ErrorDic = new Dictionary<int, string>();

        public Dictionary<string, object> GenerateXmlEdi_Mexico(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            report_folder = "";
            branch_code = "";
            SaveMessage = "";
            string subject = "";
            string updtsql = "";
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

                XmlLib.XmlErrorDic = new Dictionary<int, string>();
                XmlLib.Branch_Code = branch_code;
                XmlLib.Company_Code = company_code;
                
                string yymmdd = DateTime.Now.ToString("yyyyMMdd");
                // string ProcessNum = Lib.getProcessNumber(company_code, "MEXICO-TMM", yymm);
                string ProcessNum = "";
                if (rowtype == "CHECK-LIST")
                    ProcessNum = "Checklist";
                else
                    ProcessNum = Lib.getProcessNumber(company_code, "MEXICO-TMM", "MEXICO-TMM");

                if (ProcessNum == "")
                {
                    throw new Exception("Invalid Process ID");
                }
                // XmlLib.PROCESSID = yymm + ProcessNum;

                if (GenerateType == "CONTAINER" || GenerateType == "TRACKING" || GenerateType == "MBL-SE")
                    XmlLib.PROCESSID = String.Concat(yymmdd, ProcessNum);
                else
                    XmlLib.PROCESSID = ProcessNum;
                /*
                if (File_Display_Name == "")
                    File_Display_Name = "CntrReport.xml";
                else
                    File_Display_Name = Lib.ProperFileName(File_Display_Name) + ".xml";


                if (File_Display_NameAck == "")
                    File_Display_NameAck = "Ack-CntrReport.xml";
                else
                    File_Display_NameAck = Lib.ProperFileName(File_Display_NameAck) + ".xml";
                    */

                if (GenerateType == "AGENTBOOKING" || GenerateType == "ORDERLIST" || rowtype == "CHECK-LIST")
                    File_Extension = ".CSV";
                else
                    File_Extension = ".XML";

                File_ExtensionAck = ".XML";
                File_Display_Name = XmlLib.PROCESSID + File_Extension;
               
                // File_Display_NameAck = XmlLib.PROCESSID + File_Extension;
                File_Display_NameAck = XmlLib.PROCESSID + File_ExtensionAck;

                FolderId = Guid.NewGuid().ToString().ToUpper();
                File_Name = Lib.GetFileName(report_folder, FolderId, File_Display_Name);

                FolderId = Guid.NewGuid().ToString().ToUpper();
                File_NameAcK = Lib.GetFileName(report_folder, FolderId, File_Display_NameAck);

                Init();

                if (GenerateType=="CONTAINER"|| GenerateType == "TRACKING")
                {
                    MexicoOrdersRpt OrdRpt = new MexicoOrdersRpt();//Cargo Process
                    OrdRpt.PKID = PKID;
                    OrdRpt.InvokeType = GenerateType;
                    OrdRpt.File_Name = File_Name;
                    OrdRpt.RowType = rowtype;
                    OrdRpt.Generate();
                    subject = OrdRpt.File_Subject;
                    updtsql = OrdRpt.Ftp_updtsql;
                    if (OrdRpt.IsError)
                    {
                        throw new Exception(OrdRpt.ErrorMessage);
                    }

                    /*
                    MexicoAckRpt aCKRpt = new MexicoAckRpt();
                    aCKRpt.PKID = PKID;
                    aCKRpt.GenerateType = GenerateType;
                    aCKRpt.ProcessCount = OrdRpt.ProcessOrdCount;
                    aCKRpt.File_Name = File_NameAcK;
                    aCKRpt.Generate();
                    if (aCKRpt.IsError)
                    {
                        throw new Exception(aCKRpt.ErrorMessage);
                    }*/
                }
                if (GenerateType == "MBL-SE")
                {
                    MexicoBLRpt BlRpt = new MexicoBLRpt();
                    BlRpt.PKID = PKID;
                    BlRpt.File_Name = File_Name;
                    BlRpt.RowType = rowtype;
                    BlRpt.Generate();
                    subject = BlRpt.File_Subject;
                    if (BlRpt.IsError)
                    {
                        throw new Exception(BlRpt.ErrorMessage);
                    }
                    
                    /*
                    MexicoAckRpt aCKRpt = new MexicoAckRpt();
                    aCKRpt.PKID = PKID;
                    aCKRpt.File_Name = File_NameAcK;
                    aCKRpt.Generate();
                    if (aCKRpt.IsError)
                    {
                        throw new Exception(aCKRpt.ErrorMessage);
                    }*/
                }
                if (GenerateType == "CUSTOMER")
                {
                    MexicoSupplierRpt SupRpt = new MexicoSupplierRpt();
                    SupRpt.PKID = PKID;
                    SupRpt.branch_code = branch_code;
                    SupRpt.File_Name = File_Name;
                    SupRpt.Generate();
                    if (SupRpt.IsError)
                    {
                        throw new Exception(SupRpt.ErrorMessage);
                    }
                }
                if (GenerateType == "AGENTBOOKING" || GenerateType == "ORDERLIST")
                {
                    MexicoOrdersCsv OrdCsv = new MexicoOrdersCsv();
                    OrdCsv.PKID = PKID;
                    OrdCsv.InvokeType = GenerateType;
                    OrdCsv.File_Name = File_Name;
                    OrdCsv.Generate();
                    subject = OrdCsv.File_Subject;
                    updtsql = OrdCsv.Ftp_updtsql;
                    if (OrdCsv.IsError)
                    {
                        throw new Exception(OrdCsv.ErrorMessage);
                    }

                    MexicoAckRpt aCKRpt = new MexicoAckRpt();
                    aCKRpt.PKID = PKID;
                    aCKRpt.GenerateType = GenerateType;
                    aCKRpt.ProcessCount = OrdCsv.ProcessOrdCount;
                    aCKRpt.File_Name = File_NameAcK;
                    aCKRpt.Generate();
                    if (aCKRpt.IsError)
                    {
                        throw new Exception(aCKRpt.ErrorMessage);
                    }
                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
            RetData.Add("subject", subject);
            RetData.Add("savemsg", SaveMessage);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("filenameack", File_NameAcK);
            RetData.Add("filetypeack", File_TypeAck);
            RetData.Add("filedisplaynameack", File_Display_NameAck);
            RetData.Add("processid", XmlLib.PROCESSID);
            RetData.Add("updatesql", updtsql);
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

