using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.CM;

namespace BLXml 
{
    public partial class companyMessage : BLXml.models.XmlRoot
    {
        DBConnection Con_Oracle = null;
        public override void Generate()
        {
            this.MODULE_ID = "CM";
            this.FileName = "CM";
            if (XmlLib.Agent_Name == "RITRA")
            {
                this.MODULE_ID = "CMCE";
                this.FileName = "CMCE";
            }
            CompanyMessage MyList = new CompanyMessage();
            MyList.CompanyRecord  = GetRecords();
            if (Total_Records > 0)
            {
                MyList.MessageInfo = GetMessageInfo();
                if (XmlLib.SaveInTempFolder)
                {
                    XmlLib.FolderId = Guid.NewGuid().ToString().ToUpper();
                    XmlLib.File_Type = "XML";
                    XmlLib.File_Category = "CMCE";
                    XmlLib.File_Processid = "CMCE" + this.MessageNumber;
                    XmlLib.File_Display_Name = FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlLib.File_Name = Lib.GetFileName(XmlLib.report_folder, XmlLib.FolderId, XmlLib.File_Display_Name);
                    XmlSerializer serializer =
                          new XmlSerializer(typeof(CompanyMessage));
                    TextWriter writer = new StreamWriter(XmlLib.File_Name);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
                else
                {

                    XmlLib.CreateSentFolder();
                    FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlSerializer serializer =
                          new XmlSerializer(typeof(CompanyMessage));
                    TextWriter writer = new StreamWriter(FileName);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
            }
        }
        public CompanyMessageMessageInfo GetMessageInfo()
        {
            CompanyMessageMessageInfo VMInfo = new CompanyMessageMessageInfo();
            this.MessageNumber = XmlLib.GetNewMessageNumber();
            VMInfo.MessageNumber = this.MessageNumber;
            VMInfo.MessageSender = XmlLib.messageSenderField;
            VMInfo.MessageRecipient = XmlLib.messageRecipientField;
            VMInfo.CreatedDateTime  = XmlLib.GetCreatedDate();
            VMInfo.CreatedDateTimeSpecified = true;
            return VMInfo;
        }
        public CompanyMessageCompanyRecord[] GetRecords()
        {
            int nCtr = 0;
            DataTable Dt = null;
            string sql = "";
            string sWhere = "";
            System.Collections.ArrayList aList = new ArrayList();
            CompanyMessageCompanyRecord Record;

            /*
            if (ActualShipper == "N" && HBLNO.Trim().Length > 0)
            {
                sql = " select  distinct AGENT_CODE,AGENT_NAME, 'AGENT' as type from TABLE_XMLEDI ";
                sql += " union all";
                sql += " select  distinct PARTY.ACC_CODE  ,PARTY.ACC_NAME , 'SHIPPER' as type  ";
                sql += " from hbl_summary a";
                sql += " inner join jobinvoicem inv on (a.hbls_hbl_id = inv.inv_parent_id)";
                sql += " left join acctm 	party on inv.inv_acc_id = party.acc_pk_id";
                sql += " where hbls_bl_no = '" + HBLNO + "'";
                sql += " and inv_source like 'SEADUMMY%'";
                sql += " union all";
                sql += " select  distinct CONSIGNEE_CODE,CONSIGNEE_NAME, 'CONSIGNEE' as type from TABLE_XMLEDI ";
            }
            else
            {
                sql = " select  distinct AGENT_CODE,AGENT_NAME, 'AGENT' as type from TABLE_XMLEDI ";
                sql += " union all";
                sql += " select  distinct SHIPPER_CODE,SHIPPER_NAME, 'SHIPPER' as type from TABLE_XMLEDI ";
                sql += " union all";
                sql += " select  distinct CONSIGNEE_CODE,CONSIGNEE_NAME, 'CONSIGNEE' as type from TABLE_XMLEDI ";
            }
            */

            sWhere = " where mbl.rec_company_code = '" + XmlLib.Company_Code + "' ";
            sWhere += " and mbl.rec_branch_code  = '" + XmlLib.Branch_Code + "' ";
            sWhere += " and mbl.hbl_agent_id = '" + XmlLib.Agent_Id + "' ";
            if (XmlLib.HBL_BL_NOS.Length > 0)
            {
                sWhere += " and hbl.hbl_bl_no in (" + XmlLib.HBL_BL_NOS + ")";
            }
            else if (XmlLib.MBL_IDS.Length > 0)
            {
                sql += " and mbl.hbl_pkid in (" + XmlLib.MBL_IDS + ")";
            }
            else
            {
                // sWhere += " and (sysdate - mbl.hbl_pol_etd) between 2 and  60 ";
                sWhere += " and (((sysdate - mbl.hbl_pol_etd) between 2 and  60) or (sysdate - mbl.hbl_prealert_date) <= 5 ) ";
            }

            sql = " select distinct agnt.cust_code as agent_code,agnt.cust_name as agent_name, 'AGENT' as type from hblm mbl ";
            sql += " inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += " left join customerm agnt on mbl.hbl_agent_id = agnt.cust_pkid";
            sql +=  sWhere;
            sql += " union all";
            sql += " select distinct shpr.cust_code as shipper_code,shpr.cust_name as shipper_name, 'SHIPPER' as type from hblm mbl ";
            sql += " inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += " left join customerm shpr on hbl.hbl_exp_id = shpr.cust_pkid";
            sql += sWhere;
            sql += " union all";
            sql += " select distinct cnge.cust_code as consignee_code,cnge.cust_name as consignee_name, 'CONSIGNEE' as type from hblm mbl ";
            sql += " inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += " left join customerm cnge on hbl.hbl_imp_id = cnge.cust_pkid";
            sql += sWhere;

            Dt = new DataTable();
            Con_Oracle = new DBConnection();
            Dt = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            foreach (DataRow Dr in Dt.Rows)
            {
                if (Dr["AGENT_CODE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new CompanyMessageCompanyRecord();
                    Record.MemberCode = XmlLib.memberCode;
                    Record.Sequence = nCtr.ToString();
                    Record.Action = CompanyMessageCompanyRecordAction.Replace;
                    Record.CompanyCode = Dr["AGENT_CODE"].ToString();
                    Record.CompanyName = Dr["AGENT_NAME"].ToString();

                    Record.IsShipper= false;
                    Record.IsConsignee = false;
                    Record.IsAgent = false;
                    Record.IsNotify = false;

                    if ( Dr["type"].ToString() == "AGENT")
                        Record.IsAgent = true;
                    if (Dr["type"].ToString() == "CONSIGNEE")
                        Record.IsConsignee  = true;
                    if (Dr["type"].ToString() == "NOTIFY")
                        Record.IsNotify = true;
                    if (Dr["type"].ToString() == "SHIPPER")
                        Record.IsShipper  = true;


                    aList.Add(Record);
                }
            }
            Total_Records = nCtr;
            return (CompanyMessageCompanyRecord[])aList.ToArray(typeof(CompanyMessageCompanyRecord));
        }
    }
}