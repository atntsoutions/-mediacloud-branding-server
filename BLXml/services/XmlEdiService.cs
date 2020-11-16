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
    public class XmlEdiService : BL_Base
    {
        private int TOTAL = 0;
        private int TOTAL_VM = 0;
        private int TOTAL_VS = 0;
        private int TOTAL_BL = 0;
        private int TOTAL_ST = 0;
        private string report_folder = "";
        private string branch_code = "";
        private string branch_name = "";
        private string company_code = "";
        private string agent_id = "";
        private string agent_code = "";
        private string agent_name = "";
        private string SaveMessage = "Generate Complete ";
        private string DisplayMsg = "";
        private string cost_pkid = "";
        private string folder_sent_on = "";
        private string GenerateType = "SEA";
        private string hbl_nos = "";
        private string Subject = "";
        private string TradingPartner = "";
        List<FileDetails> fList = new List<FileDetails>();
        Dictionary<int, string> ErrorDic = new Dictionary<int, string>();
        public Dictionary<string, object> GenerateXmlEdi(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            fList = new List<FileDetails>();
            report_folder = "";
            branch_code = "";
            hbl_nos = "";
            Subject = "";
            try
            {
                XmlLib.SaveInTempFolder = false;
                XmlLib.MBL_IDS = "";

                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();
                if (SearchData.ContainsKey("branch_name"))
                    branch_name = SearchData["branch_name"].ToString();
                if (SearchData.ContainsKey("company_code"))
                    company_code = SearchData["company_code"].ToString();
                if (SearchData.ContainsKey("agent_id"))
                    agent_id = SearchData["agent_id"].ToString();
                if (SearchData.ContainsKey("agent_code"))
                    agent_code = SearchData["agent_code"].ToString();
                if (SearchData.ContainsKey("agent_name"))
                    agent_name = SearchData["agent_name"].ToString();
                if (SearchData.ContainsKey("hbl_nos"))
                    hbl_nos = SearchData["hbl_nos"].ToString();
                if (SearchData.ContainsKey("type"))
                    GenerateType = SearchData["type"].ToString();
                if (SearchData.ContainsKey("mbl_id"))
                    XmlLib.MBL_IDS = SearchData["mbl_id"].ToString();

                if (agent_id == "")
                    throw new Exception("Agent ID Not Found");

                sql = "select param_pkid,param_code from param where param_type='PARAM' and param_id1='" + agent_code + "'";
                Con_Oracle = new DBConnection();
                DataTable Dt_temp = new DataTable();
                Dt_temp = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                if (Dt_temp.Rows.Count > 0)
                    TradingPartner = Dt_temp.Rows[0]["param_code"].ToString();

                if (TradingPartner == "")
                    throw new Exception("Trading Partner Information not Found");

                XmlLib.XmlErrorDic = new Dictionary<int, string>();
                XmlLib.Agent_Id = agent_id;
                XmlLib.Agent_Code = agent_code;
                XmlLib.Agent_Name = agent_name;
               
                XmlLib.Branch_Code = branch_code;
                XmlLib.Company_Code = company_code;

                //if (agent_code.StartsWith("RITRA"))
                //    GenerateRitraXml();
                //else if (agent_code.StartsWith("MOTHERLINES"))
                //    GenerateMotherlinesXml();
                //else
                //    GenerateSeaBLXml();

                if (TradingPartner == "RITRA")
                    GenerateRitraXml();
                else if (TradingPartner == "MOTHERLINES-US")
                    GenerateMotherlinesXml();
                else
                    throw new Exception("FTP Transfer not Implemented for Agent " + agent_name);
                //GenerateSeaBLXml();
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
            RetData.Add("subject", Subject);
            RetData.Add("savemsg", SaveMessage);
            RetData.Add("filelist", fList);
            return RetData;
        }


        private void GenerateRitraXml()
        {
            XmlLib.messageSenderField = "CARGOMAR";
            XmlLib.messageRecipientField = "FECL";
            XmlLib.MessageNumberSeq = Lib.getProcessNumber(company_code, "XML-FILENAME-SEQ", "XML-FILENAME-SEQ");
            XmlLib.memberCode = "CARGOMAR";
            if (GenerateType == "AIR")
                XmlLib.RootFolder = report_folder + "\\xmldata\\AIR\\" + branch_code;
            else
                XmlLib.RootFolder = report_folder + "\\xmldata\\SEA\\" + branch_code;

            if (XmlLib.MBL_IDS != "") //Invoke from Booking Sea Page
            {
                XmlLib.MBL_IDS = "'" + XmlLib.MBL_IDS + "'";
                XmlLib.SaveInTempFolder = true;
                XmlLib.report_folder = report_folder;
                fList = new List<FileDetails>();
                Subject = "HBL-" + GetHBL_Nos();
                Init();

                CreateXml(new vesselMessage());
                fList.Add(GetFileDetails());
                CreateXml(new VesselSchedule());
                fList.Add(GetFileDetails());
                CreateXml(new companyMessage());
                fList.Add(GetFileDetails());
                CreateXml(new unlocodeMessage());
                fList.Add(GetFileDetails());
                CreateXml(new shipmentTracking());
                fList.Add(GetFileDetails());
                CreateXml(new BillLading());
                fList.Add(GetFileDetails());
            }
            else 
            {
                //Generate for a period 

                XmlLib.MBL_IDS = "";
                XmlLib.HBL_BL_NOS = "";
                if (hbl_nos.Trim().Length > 0)
                {
                    XmlLib.HBL_BL_NOS = "'" + hbl_nos.Replace(" ", "").Replace(",", "','") + "'";
                    XmlLib.MBL_IDS = GetMBL_Ids();
                }

                Init();
                XmlLib.CreateSentFolder();

                if (GenerateType == "AIR")
                {
                    DisplayMsg = " Sender : " + XmlLib.messageSenderField;
                    DisplayMsg += ", Recipient : " + XmlLib.messageRecipientField;
                    DisplayMsg += ", Files Generated : ";

                    AWBillLadingMsg AwBillMsg = new AWBillLadingMsg();
                    AwBillMsg.Generate();
                    if (AwBillMsg.IsError == false)
                    {
                        DisplayMsg += " Bill Of Lading";
                        TOTAL++;
                    }

                    if (TOTAL > 0)
                    {
                        AWCompanyMsg AwCompMsg = new AWCompanyMsg();
                        AwCompMsg.Generate();
                        if (AwCompMsg.IsError == false)
                        {
                            DisplayMsg += ", Company Message";
                            TOTAL++;
                        }
                    }
                    if (TOTAL > 1)
                    {
                        AWTrackingMsg AwTrackMsg = new AWTrackingMsg();
                        AwTrackMsg.Generate();
                        if (AwTrackMsg.IsError == false)
                        {
                            DisplayMsg += ", Shipment Tracking";
                            TOTAL++;
                        }
                    }
                    if (TOTAL > 0)
                        SaveMessage = "Generate Details  :- " + DisplayMsg;
                    else
                        SaveMessage = "No Changes Found, Xml EDI";
                }
                else
                {

                    DisplayMsg = " Sender : " + XmlLib.messageSenderField;
                    DisplayMsg += ", Recipient : " + XmlLib.messageRecipientField;
                    DisplayMsg += ", Files Generated : ";
                    CreateXml(new vesselMessage());
                    DisplayMsg += " Vessel Messages";
                    CreateXml(new VesselSchedule());
                    DisplayMsg += ",VesselSchedule";
                    CreateXml(new companyMessage());
                    DisplayMsg += ",companyMessage";
                    CreateXml(new unlocodeMessage());
                    DisplayMsg += ",unlocodeMessage";
                    CreateXml(new shipmentTracking());
                    DisplayMsg += ",shipmentTracking";
                    CreateXml(new BillLading());
                    DisplayMsg += ",BillLading";

                    if (TOTAL > 0)
                    {
                        SaveMessage = "Generate Details  :- " + DisplayMsg;
                        WriteErrorFile();
                    }
                    else
                        SaveMessage = "No Changes Found, Xml EDI";
                }
            }

        }


        private void GenerateSeaBLXml()
        {
            XmlLib.messageSenderField = "CARGOMAR";
            XmlLib.messageRecipientField = agent_code;
            XmlLib.MessageNumberSeq = Lib.getProcessNumber(company_code, "XML-FILENAME-SEQ", "XML-FILENAME-SEQ");
            XmlLib.memberCode = "CARGOMAR";
            XmlLib.RootFolder = report_folder + "\\xmldata\\SEA\\" + branch_code;
            if (XmlLib.MBL_IDS != "")
            {
                XmlLib.MBL_IDS = "'" + XmlLib.MBL_IDS + "'";
                XmlLib.SaveInTempFolder = true;
                XmlLib.report_folder = report_folder;
                fList = new List<FileDetails>();
                Subject = "HBL-" + GetHBL_Nos();
                Init();
                CreateXml(new BillLading());
                fList.Add(GetFileDetails());
            }
        }
        private void GenerateMotherlinesXml()
        {
            XmlLib.messageSenderField = "CARGOMAR";
            XmlLib.messageRecipientField = agent_code;
            XmlLib.MessageNumberSeq = Lib.getProcessNumber(company_code, "MOTHERLINES-US", "MOTHERLINES-US");
            XmlLib.memberCode = "CARGOMAR";
            XmlLib.RootFolder = report_folder + "\\xmldata\\SEA\\" + branch_code;
            if (XmlLib.MBL_IDS != "")
            {
                XmlLib.MBL_IDS = "'" + XmlLib.MBL_IDS + "'";
                XmlLib.SaveInTempFolder = true;
                XmlLib.report_folder = report_folder;
                fList = new List<FileDetails>();
                Subject = "HBL-" + GetHBL_Nos();
                Init();
               
                CmarBillLading CmarBl = new CmarBillLading();
                CmarBl.MessageBranchName = branch_name;
                CmarBl.Generate();
                if (CmarBl.IsError)
                {
                    throw new Exception(CmarBl.ErrorMessage);
                }
                fList.Add(GetFileDetails());
            }
        }
        private FileDetails GetFileDetails()
        {
            FileDetails fRec = new FileDetails();
            fRec.filetype = XmlLib.File_Type;
            fRec.filedisplayname = XmlLib.File_Display_Name;
            fRec.filename = XmlLib.File_Name;
            fRec.filecategory = XmlLib.File_Category;
            fRec.fileprocessid = XmlLib.File_Processid ;
            return fRec;
        }

        private void Init()
        {
            XmlLib.sentFolder = null;
            TOTAL = 0;
            TOTAL_VM = 0;
            TOTAL_VS = 0;
            TOTAL_ST = 0;
            TOTAL_BL = 0;
        }
        private Boolean CreateXml(XmlRoot T)
        {
            Boolean bRet = false;
            try
            {
                T.Generate();
                TOTAL += T.Total_Records;
                if (T.MODULE_ID == "VM")
                    TOTAL_VM = T.Total_Records;
                if (T.MODULE_ID == "VS")
                    TOTAL_VS = T.Total_Records;
                if (T.MODULE_ID == "ST")
                    TOTAL_ST = T.Total_Records;
                if (T.MODULE_ID == "BL")
                    TOTAL_BL = T.Total_Records;
                bRet = true;
            }
            catch (Exception)
            {
                bRet = false;
                throw;
            }
            return bRet;
        }



        public Dictionary<string, object> GenerateXmlCostingInvoice(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            report_folder = "";
            branch_code = "";
            try
            {
                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();
                if (SearchData.ContainsKey("company_code"))
                    company_code = SearchData["company_code"].ToString();
                if (SearchData.ContainsKey("cost_pkid"))
                    cost_pkid = SearchData["cost_pkid"].ToString();
                if (SearchData.ContainsKey("agent_id"))
                    agent_id = SearchData["agent_id"].ToString();

                if (SearchData.ContainsKey("sent_on"))
                    folder_sent_on = SearchData["sent_on"].ToString();

                if (cost_pkid != "")
                {
                    CreateCostInvoiceXml("SINGLE");
                }
                else if (folder_sent_on != "")
                {
                    TOTAL = 0;
                    ErrorDic = new Dictionary<int, string>();
                    folder_sent_on = Lib.StringToDate(folder_sent_on);
            
                    sql = "select cost_pkid,rec_branch_code,cost_sent_on,cost_refno from costingm ";
                    sql += " where cost_agent_id = '{AGENTID}' ";
                    sql += " and cost_sent_on between  '{FDATE}'  and  '{EDATE}' ";

                    sql = sql.Replace("{AGENTID}", agent_id);
                    sql = sql.Replace("{FDATE}", folder_sent_on);
                    sql = sql.Replace("{EDATE}", DateTime.Now.ToString(Lib.BACK_END_DATE_FORMAT));

                    Con_Oracle = new DBConnection();
                    DataTable Dt_Cost = new DataTable();
                    Dt_Cost = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    foreach (DataRow dr in Dt_Cost.Rows)
                    {
                        cost_pkid = dr["cost_pkid"].ToString();
                        branch_code = dr["rec_branch_code"].ToString();
                        CreateCostInvoiceXml("MULTIPLE",dr["cost_refno"].ToString());
                    }
                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
            RetData.Add("savemsg", SaveMessage);
            return RetData;
        }

        private void CreateCostInvoiceXml(string stype,string costRefNo="")
        {
            XmlLib.messageSenderField = "CARGOMAR";
            XmlLib.messageRecipientField = "FECL";
            XmlLib.memberCode = "CARGOMAR";
            XmlLib.RootFolder = report_folder + "\\xmldata\\COSTING\\" + branch_code + "\\";
            //XmlLib.Agent_Name = XmlLib.messageSenderField;
            XmlLib.Agent_Name = "RITRA";
            XmlLib.Agent_Id = agent_id;
            XmlLib.Branch_Code = branch_code;
            XmlLib.Company_Code = company_code;
            XmlLib.sentFolder = null;
            XmlLib.CreateDayFolder();

            CostInvoice cInv = new CostInvoice();
            cInv.COST_PKID = cost_pkid;
            cInv.Generate();
            if (stype == "SINGLE")
            {
                if (cInv.IsError)
                    SaveMessage = "Generate Details  :- " + cInv.ErrorMessage;
                else
                    SaveMessage = " Costing Invoice XML Generated Successfully.";
            }
            else
            {
                if (cInv.IsError)
                    ErrorDic.Add(ErrorDic.Count, costRefNo);
                else
                    TOTAL++;

                SaveMessage = "Generate Details  :- ";
                if (TOTAL > 0)
                    SaveMessage += TOTAL + " Files Generated Successfully.";
                if (ErrorDic.Count > 0)
                    SaveMessage += " Files Not Generated ";
                for (int i = 0; i < ErrorDic.Count; i++)
                    SaveMessage += " | " + ErrorDic[i];

            }
        }
        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            //Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "SALES EXECUTIVE");
            //RetData.Add("smanlist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "CITY");
            //RetData.Add("citylist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "STATE");
            //RetData.Add("statelist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;
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

        private string GetMBL_Ids()
        {
            string sIds = "";

            sql = "select hbl_mbl_id from hblm a ";
            sql += " where a.rec_company_code = '{COMPCODE}'";
            sql += " and a.rec_branch_code = '{BRCODE}'";
            if (GenerateType == "AIR")
                sql += " and a.hbl_type = 'HBL-AE'";
            else
                sql += " and a.hbl_type = 'HBL-SE'";
            sql += " and a.hbl_bl_no in (" + XmlLib.HBL_BL_NOS + ")";

            sql = sql.Replace("{COMPCODE}", company_code);
            sql = sql.Replace("{BRCODE}", branch_code);

            Con_Oracle = new DBConnection();
            DataTable Dt_MBL = new DataTable();
            Dt_MBL = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            sIds = "";
            foreach (DataRow dr in Dt_MBL.Rows)
            {
                if (sIds.Trim() != "")
                    sIds += ",";

                sIds += "'" + dr["hbl_mbl_id"].ToString() + "'";
            }
            if (sIds == "")
                sIds = "'NA'";

            Dt_MBL.Rows.Clear();

            return sIds;
        }
        private string GetHBL_Nos()
        {
            bool differentHblDate = false;
            string sNos = "";
            sql = "select hbl_bl_no,hbl_date from hblm a ";
            sql += " where a.hbl_mbl_id in (" + XmlLib.MBL_IDS + ") order by hbl_bl_no";

            Con_Oracle = new DBConnection();
            DataTable Dt_HBL = new DataTable();
            Dt_HBL = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            if (Dt_HBL.Rows.Count > 0)
            {
                DataTable DistinctBLDT = Dt_HBL.DefaultView.ToTable(true, "hbl_date");
                if (DistinctBLDT.Rows.Count > 1)
                    differentHblDate = true;

                string sDate = "";
                sNos = "";
                foreach (DataRow dr in Dt_HBL.Rows)
                {
                    if (sNos.Trim() != "")
                        sNos += ", ";
                    sNos += dr["hbl_bl_no"].ToString();
                    sDate = Lib.DatetoStringDisplayformat(dr["hbl_date"]);
                    if (differentHblDate)
                        sNos += "-" + sDate;
                }
                if (!differentHblDate && sNos.Trim() != "" && sDate.Trim() != "")
                    sNos += "-" + sDate;
            }
            Dt_HBL.Rows.Clear();
            return sNos;
        }
    }
}

