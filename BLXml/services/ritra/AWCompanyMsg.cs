using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.AirCM;

namespace BLXml 
{
    public partial class AWCompanyMsg
    {
        private DataTable DT_Company = new DataTable();
        public Boolean IsError = false;
        private CompanyMessage CompMessage = null;
        Dictionary<int, string> ErrorDic = new Dictionary<int, string>();
        private string ErrorValues = "";
        private string sql = "";
        private string MessageNumber = "";
        private int CompSeq = 0;
        DBConnection Con_Oracle = null;
        public void Generate()
        {
            try
            {
                IsError = false;
                ReadData();
                if (DT_Company.Rows.Count <= 0)
                {
                    IsError = true;
                    return;
                }

                GenerateXmlFiles();
                WriteXmlFiles();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void ReadData()
        {
            string sWhere = "";
            sWhere = " where m.rec_company_code = '" + XmlLib.Company_Code + "' ";
            sWhere += " and m.rec_branch_code  = '" + XmlLib.Branch_Code + "' ";
            sWhere += " and m.hbl_agent_id = '" + XmlLib.Agent_Id + "' ";
            if (XmlLib.HBL_BL_NOS.Length > 0)
            {
                sWhere += " and h.hbl_bl_no in (" + XmlLib.HBL_BL_NOS + ")";
            }
            else
            {
                sWhere += " and ((sysdate - m.hbl_pol_etd) between 2 and  60 )";
            }

            sql = " select distinct CompanyCode,CompanyName,AddressLine1,AddressLine2,AddressLine3,AddressLine4 ";
            sql += "   from (";
            //Agent
            sql += " select agnt.cust_code as CompanyCode,agnt.cust_name as CompanyName,";
            sql += " agntaddr.add_line1 as AddressLine1,agntaddr.add_line2 as AddressLine2,agntaddr.add_line3 as AddressLine3,agntaddr.add_line4 as AddressLine4 ";
            sql += " from hblm m ";
            sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
            sql += " left join customerm agnt on m.hbl_agent_id = agnt.cust_pkid";
            sql += " left join addressm agntaddr on m.hbl_agent_br_id = agntaddr.add_pkid ";
            sql += sWhere;

            sql += "  union all";

            //Shipper
            sql += " select shpr.cust_code as CompanyCode,shpr.cust_name as CompanyName,";
            sql += " shpraddr.add_line1 as AddressLine1,shpraddr.add_line2 as AddressLine2,shpraddr.add_line3 as AddressLine3,shpraddr.add_line4 as AddressLine4 ";
            sql += " from hblm m ";
            sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
            sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";
            sql += " left join addressm shpraddr on m.hbl_exp_br_id = shpraddr.add_pkid ";
            sql += "  " + sWhere;

            sql += "   union all";

            //Consignee
            sql += " select cnge.cust_code as CompanyCode,cnge.cust_name as CompanyName,";
            sql += " cngeaddr.add_line1 as AddressLine1,cngeaddr.add_line2 as AddressLine2,cngeaddr.add_line3 as AddressLine3,cngeaddr.add_line4 as AddressLine4 ";
            sql += " from hblm m ";
            sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
            sql += " left join customerm cnge on h.hbl_imp_id = cnge.cust_pkid";
            sql += " left join addressm cngeaddr on m.hbl_imp_br_id = cngeaddr.add_pkid ";
            sql += "  " + sWhere;

            sql += "   union all";

            //Notify
            sql += " select nvl(nfy.cust_code,b.bl_notify_name) as CompanyCode,nvl(nfy.cust_name,b.bl_notify_name) as CompanyName,";
            sql += " nfyaddr.add_line1 as AddressLine1,nfyaddr.add_line2 as AddressLine2,nfyaddr.add_line3 as AddressLine3,nfyaddr.add_line4 as AddressLine4 ";
            sql += " from hblm m ";
            sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
            sql += " inner join bl b on h.hbl_pkid = b.bl_pkid";
            sql += " left join customerm nfy on b.bl_notify_id = nfy.cust_pkid ";
            sql += " left join addressm nfyaddr on b.bl_notify_br_id = nfyaddr.add_pkid ";
            sql += "  " + sWhere;

            sql += "  )a order by CompanyCode";

            DT_Company = new DataTable();
            Con_Oracle = new DBConnection();
            DT_Company = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
        }
        private void GenerateXmlFiles()
        {
            CompMessage = new CompanyMessage();
            CompMessage.Items = Generate_CompanyMessage();
        }
        private object[] Generate_CompanyMessage()
        {
            object[] Items = null;
            int iTotRows = 0;
            int ArrIndex = 0;
            try
            {
                CompSeq = 0;
                iTotRows = DT_Company.Rows.Count;
                Items = new object[iTotRows + 1];
                Items[ArrIndex++] = Generate_MessageInfo();
                foreach (DataRow dr in DT_Company.Rows)
                {
                    CompSeq++;
                    Items[ArrIndex++] = Generate_CompanyRecord(dr);
                    if (IsError)
                        break;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return Items;
        }
        private CompanyMessageMessageInfo Generate_MessageInfo()
        {
            CompanyMessageMessageInfo Rec = null;
            try
            {
                this.MessageNumber = XmlLib.GetNewMessageNumber();
                Rec = new CompanyMessageMessageInfo();
                Rec.MessageSender = XmlLib.messageSenderField;
                Rec.MessageNumber = this.MessageNumber;
                Rec.MessageRecipient = XmlLib.messageRecipientField;
                Rec.CreatedDateTime = ConvertYMDDate(DateTime.Now.ToString());
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return Rec;
        }


        private CompanyMessageCompanyRecord Generate_CompanyRecord(DataRow dr)
        {
            CompanyMessageCompanyRecord Rec = null;
            try
            {
                Rec = new CompanyMessageCompanyRecord();
                Rec.Action = "Replace";
                Rec.Sequence = CompSeq.ToString();
                Rec.MemberCode = XmlLib.messageSenderField;
                Rec.CompanyCode = dr["CompanyCode"].ToString();
                Rec.CompanyName = dr["CompanyName"].ToString();
                Rec.AddressLine1 = dr["AddressLine1"].ToString();
                Rec.AddressLine2 = dr["AddressLine2"].ToString();
                Rec.AddressLine3 = dr["AddressLine3"].ToString();
                Rec.AddressLine4 = dr["AddressLine4"].ToString();
                Rec.City = "";
                Rec.ZipCode = "";
                Rec.Country = "";
                Rec.ContactPerson = "";
                Rec.Department = "";
                Rec.TelephoneNumber = "";
                Rec.FaxNumber = "";
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return Rec;
        }
        private void WriteXmlFiles()
        {
            try
            {
                if (CompMessage == null || IsError)
                    return;

                string FileName = "AWC";
                FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(CompanyMessage));
                StreamWriter writer = new StreamWriter(FileName);
                mySerializer.Serialize(writer, CompMessage, ns);
                writer.Close();
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
        }
        private string ConvertYMDDate(string sDate)
        {
            if (sDate != null)
            {
                if (sDate.Trim().Length > 0)
                    sDate = Convert.ToDateTime(sDate).ToString("yyyy-MM-dd HH:mm:ss");
            }
            return sDate;
        }

    }
}
