using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.VM;

namespace BLXml
{
    public partial class vesselMessage : BLXml.models.XmlRoot
    {
        DBConnection Con_Oracle = null;
        public override void Generate()
        {
            this.MODULE_ID = "VM";
            this.FileName = "VM";
            if (XmlLib.Agent_Name == "RITRA")
            {
                this.MODULE_ID = "VMCE";
                this.FileName = "VMCE";
            }
            VesselMessage MyList = new VesselMessage();

            MyList.Vessel = GetRecords();
            if (Total_Records > 0)
            {
                MyList.MessageInfo = GetMessageInfo();
                if (XmlLib.SaveInTempFolder)
                {
                    XmlLib.FolderId = Guid.NewGuid().ToString().ToUpper();
                    XmlLib.File_Type = "XML";
                    XmlLib.File_Category = "VMCE";
                    XmlLib.File_Processid = "VMCE" + this.MessageNumber;
                    XmlLib.File_Display_Name = FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlLib.File_Name = Lib.GetFileName(XmlLib.report_folder, XmlLib.FolderId, XmlLib.File_Display_Name);

                    //XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                    //ns.Add("", "");
                    XmlSerializer serializer = new XmlSerializer(typeof(VesselMessage));
                    TextWriter writer = new StreamWriter(XmlLib.File_Name);
                    serializer.Serialize(writer, MyList);
                    //serializer.Serialize(writer, MyList, ns);
                    writer.Close();
                }
                else
                {
                    XmlLib.CreateSentFolder();
                    FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlSerializer serializer = new XmlSerializer(typeof(VesselMessage));
                    TextWriter writer = new StreamWriter(FileName);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
            }
        }
        public VesselMessageMessageInfo GetMessageInfo()
        {
            VesselMessageMessageInfo VMInfo = new VesselMessageMessageInfo();
            this.MessageNumber = XmlLib.GetNewMessageNumber();
            VMInfo.MessageNumber = this.MessageNumber;
            VMInfo.MessageSender = XmlLib.messageSenderField;
            VMInfo.MessageRecipient = XmlLib.messageRecipientField;
            VMInfo.CreatedDateTime = XmlLib.GetCreatedDate();  
            VMInfo.CreatedDateTimeSpecified = true;
            return VMInfo;
        }
        public VesselMessageVessel[] GetRecords()
        {
            int nCtr = 0;
            DataTable Dt = null;
            string sql = "";
            System.Collections.ArrayList aList = new ArrayList();
            VesselMessageVessel Record;

            /*
            sql = "";
            sql += " SELECT DISTINCT VESSEL1_CODE,VESSEL1_NAME ";
            sql += " FROM TABLE_XMLEDI A  ";
            sql += " WHERE (VESSEL1_CODE IS NOT NULL )";
            sql += " union ";
            sql += " SELECT DISTINCT VESSEL2_CODE,VESSEL2_NAME ";
            sql += " FROM TABLE_XMLEDI A  ";
            sql += " WHERE (VESSEL2_CODE IS NOT NULL )";
            sql += " union ";
            sql += " SELECT DISTINCT VESSEL3_CODE,VESSEL3_NAME ";
            sql += " FROM TABLE_XMLEDI A  ";
            sql += " WHERE (VESSEL3_CODE IS NOT NULL )";
            sql += " union ";
            sql += " SELECT DISTINCT VESSEL4_CODE,VESSEL4_NAME ";
            sql += " FROM TABLE_XMLEDI A  ";
            sql += " WHERE (VESSEL4_CODE IS NOT NULL )";
            */

            sql = " select distinct vsl.param_code as vessel1_code,vsl.param_name as vessel1_name from hblm mbl";
            sql += " inner join trackingm trk on mbl.hbl_pkid = trk.trk_parent_id";
            sql += " left join param vsl on trk_vsl_id= vsl.param_pkid ";
            sql += " where mbl.rec_company_code = '" + XmlLib.Company_Code + "' ";
            sql += " and mbl.rec_branch_code  = '" + XmlLib.Branch_Code + "' ";
            sql += " and trk.rec_category  = 'SEA EXPORT' ";
            sql += " and mbl.hbl_agent_id = '" + XmlLib.Agent_Id + "' ";
            if (XmlLib.MBL_IDS.Length > 0)// HBL_BL_NOS wise
            {
                sql += " and mbl.hbl_pkid in (" + XmlLib.MBL_IDS + ")";
            }
            else
            {
                // sql += " and (sysdate - mbl.hbl_pol_etd) between 2 and  60 ";
                sql += " and (((sysdate - mbl.hbl_pol_etd) between 2 and  60) or (sysdate - mbl.hbl_prealert_date) <= 5 ) ";
            }

            Dt = new DataTable();
            Con_Oracle = new DBConnection();
            Dt = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Dr in Dt.Rows)
            {
                if (Dr["VESSEL1_CODE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new VesselMessageVessel();
                    Record.MemberCode = XmlLib.memberCode;
                    Record.Sequence = nCtr.ToString();
                    Record.Action = VesselMessageVesselAction.Replace;
                    Record.VesselCode = Dr["VESSEL1_CODE"].ToString();
                    Record.VesselName = Dr["VESSEL1_NAME"].ToString();

                    aList.Add(Record );
                }
            }
            Total_Records = nCtr;
            return (VesselMessageVessel[])aList.ToArray(typeof(VesselMessageVessel));
        }
    }
}