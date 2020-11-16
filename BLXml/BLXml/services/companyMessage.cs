using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
 
namespace BLXml.service
{
 
    public partial class companyMessage : BLXml.models.XmlRoot
    {

        public override void Generate()
        {
            this.MODULE_ID = "CM";
            this.FileName = "CM";
            if (Lib.Agent_Name == "RITRA")
            {
                this.MODULE_ID = "CMCE";
                this.FileName = "CMCE";
            }
            CompanyMessage MyList = new CompanyMessage();

            MyList.CompanyRecord  = GetRecords();
            if (Total_Records > 0)
            {
                MyList.MessageInfo  = GetMessageInfo();

                Lib.CreateSentFolder();
                FileName = Lib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                XmlSerializer serializer =
                      new XmlSerializer(typeof(CompanyMessage));
                TextWriter writer = new StreamWriter(FileName);
                serializer.Serialize(writer, MyList);
                writer.Close();
            }
        }
        public CompanyMessageMessageInfo GetMessageInfo()
        {
            CompanyMessageMessageInfo VMInfo = new CompanyMessageMessageInfo();
            this.MessageNumber = Lib.GetNewMessageNumber();
            VMInfo.MessageNumber = this.MessageNumber;
            VMInfo.MessageSender = Lib.messageSenderField;
            VMInfo.MessageRecipient = Lib.messageRecipientField;
            VMInfo.CreatedDateTime  = Lib.GetCreatedDate();
            VMInfo.CreatedDateTimeSpecified = true;
            return VMInfo;
        }
        public CompanyMessageCompanyRecord[] GetRecords()
        {
            int nCtr = 0;
            DataTable Dt = null;
            string sql = "";
            System.Collections.ArrayList aList = new ArrayList();
            CompanyMessageCompanyRecord Record;


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

            Dt = new DataTable();
            StoredProcedure.CreateCommand(sql);
            StoredProcedure.Run(Dt);
            foreach (DataRow Dr in Dt.Rows)
            {
                if (Dr["AGENT_CODE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new CompanyMessageCompanyRecord();
                    Record.MemberCode = Lib.memberCode;
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