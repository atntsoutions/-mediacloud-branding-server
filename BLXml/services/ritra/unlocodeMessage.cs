using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.UC;

namespace BLXml
{
    public partial class unlocodeMessage : BLXml.models.XmlRoot
    {
        DBConnection Con_Oracle = null;
        public override void Generate()
        {
            this.MODULE_ID = "UL";
            this.FileName = "UL";
            if (XmlLib.Agent_Name == "RITRA")
            {
                this.MODULE_ID = "ULCE";
                this.FileName = "ULCE";
            }
            UnLoCodeMessage MyList = new UnLoCodeMessage();
            MyList.UnLoCodeRecord = GetRecords();
            if (Total_Records > 0)
            {
                MyList.MessageInfo = GetMessageInfo();
                if (XmlLib.SaveInTempFolder)
                {
                    XmlLib.FolderId = Guid.NewGuid().ToString().ToUpper();
                    XmlLib.File_Type = "XML";
                    XmlLib.File_Category = "ULCE";
                    XmlLib.File_Processid = "ULCE" + this.MessageNumber;
                    XmlLib.File_Display_Name = FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlLib.File_Name = Lib.GetFileName(XmlLib.report_folder, XmlLib.FolderId, XmlLib.File_Display_Name);
                    XmlSerializer serializer =
                         new XmlSerializer(typeof(UnLoCodeMessage));
                    TextWriter writer = new StreamWriter(XmlLib.File_Name);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
                else
                {
                    XmlLib.CreateSentFolder();
                    FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlSerializer serializer =
                          new XmlSerializer(typeof(UnLoCodeMessage));
                    TextWriter writer = new StreamWriter(FileName);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
            }
        }
        public UnLoCodeMessageMessageInfo GetMessageInfo()
        {
            UnLoCodeMessageMessageInfo VMInfo = new UnLoCodeMessageMessageInfo();
            this.MessageNumber = XmlLib.GetNewMessageNumber();
            VMInfo.MessageNumber = this.MessageNumber;
            VMInfo.MessageSender = XmlLib.messageSenderField;
            VMInfo.MessageRecipient = XmlLib.messageRecipientField;
            VMInfo.CreatedDateTime = XmlLib.GetCreatedDate();
            VMInfo.CreatedDateTimeSpecified = true;
            return VMInfo;
        }
        public UnLoCodeMessageUnLoCodeRecord[] GetRecords()
        {
            int nCtr = 0;
            DataTable Dt = null;
            string sql = "";
            System.Collections.ArrayList aList = new ArrayList();
            UnLoCodeMessageUnLoCodeRecord Record;
            Con_Oracle = new DBConnection();

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
                //sql += " and (sysdate - mbl.hbl_pol_etd) between 2 and  60 ";
                sql += " and (((sysdate - mbl.hbl_pol_etd) between 2 and  60) or (sysdate - mbl.hbl_prealert_date) <= 5 ) ";
            }

            Dt = new DataTable();
            Dt = Con_Oracle.ExecuteQuery(sql);
          
            foreach (DataRow Dr in Dt.Rows)
            {
                if (Dr["VESSEL1_CODE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new UnLoCodeMessageUnLoCodeRecord();
                    Record.Sequence = nCtr.ToString();
                    Record.Action = UnLoCodeMessageUnLoCodeRecordAction.Replace ;
                    Record.UnLoCode  = Dr["VESSEL1_CODE"].ToString();
                    Record.Name = Dr["VESSEL1_NAME"].ToString();
                    aList.Add(Record);
                }
            }

            /*
            // PORT
            sql = "";
            sql += " SELECT DISTINCT POL_CODE AS CODE , POL_NAME AS FNAME ";
            sql += " from TABLE_XMLEDI  "; ;
            sql += " where (POL_CODE IS NOT NULL ) ";
            sql += " UNION ";
            sql += " SELECT DISTINCT POFD_CODE,POFD_NAME ";
            sql += " from TABLE_XMLEDI  "; ;
            sql += " where (POFD_CODE IS NOT NULL ) ";
            sql += " UNION ";
            sql += " SELECT DISTINCT TRANSIT1_CODE,TRANSIT1_NAME ";
            sql += " from TABLE_XMLEDI  "; ;
            sql += " where (TRANSIT1_CODE IS NOT NULL ) ";
            sql += " UNION ";
            sql += " SELECT DISTINCT TRANSIT2_CODE,TRANSIT2_NAME ";
            sql += " from TABLE_XMLEDI  "; ;
            sql += " where (TRANSIT2_CODE IS NOT NULL ) ";
            sql += " UNION ";
            sql += " SELECT DISTINCT TRANSIT3_CODE,TRANSIT3_NAME ";
            sql += " from TABLE_XMLEDI  "; ;
            sql += " where (TRANSIT3_CODE IS NOT NULL ) ";
            sql += " UNION ";
            sql += " SELECT DISTINCT TRANSIT4_CODE,TRANSIT4_NAME ";
            sql += " from TABLE_XMLEDI "; 
            sql += " where (TRANSIT4_CODE IS NOT NULL ) ";
            */

            sql = " select distinct pol.param_code as code,pol.param_name as name from hblm mbl";
            sql += " inner join trackingm trk on mbl.hbl_pkid = trk.trk_parent_id";
            sql += " left join param pol on trk.trk_pol_id = pol.param_pkid ";
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
                //sql += " and (sysdate - mbl.hbl_pol_etd) between 2 and  60 ";
                sql += " and (((sysdate - mbl.hbl_pol_etd) between 2 and  60) or (sysdate - mbl.hbl_prealert_date) <= 5 ) ";
            }

            sql += " UNION ALL ";

            sql += " select distinct pod.param_code as code,pod.param_name as name from hblm mbl";
            sql += " inner join trackingm trk on mbl.hbl_pkid = trk.trk_parent_id";
            sql += " left join param pod on trk.trk_pod_id = pod.param_pkid ";
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

            sql += " UNION ALL ";

            sql += " select distinct pofd.param_code as code,pofd.param_name as name from hblm mbl";
            sql += " inner join trackingm trk on mbl.hbl_pkid = trk.trk_parent_id";
            sql += " left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid ";
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
            Dt = Con_Oracle.ExecuteQuery(sql);

            foreach (DataRow Dr in Dt.Rows)
            {
                if (Dr["Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new UnLoCodeMessageUnLoCodeRecord();
                    Record.Sequence = nCtr.ToString();
                    Record.Action = UnLoCodeMessageUnLoCodeRecordAction.Replace;
                    Record.UnLoCode = Dr["Code"].ToString();
                    Record.Name = Dr["Name"].ToString();
                    aList.Add(Record);
                }
            }

            /*
            // REC PLACE, ISSUE PLACE
            sql = "";
            sql += " SELECT DISTINCT PLACE_CODE AS CODE ,PLACE_NAME AS FNAME ";
            sql += " from TABLE_XMLEDI "; ;
            sql += " where (PLACE_CODE IS NOT NULL ) ";
            */

            sql = " select distinct  rcpt.param_code as code ,rcpt.param_name as name from hblm mbl";
            sql += " inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += " inner join jobm job on hbl.hbl_pkid = job.jobs_hbl_id";
            sql += " left join param rcpt on job.job_place_receipt_id = rcpt.param_pkid";
            sql += " where mbl.rec_company_code = '" + XmlLib.Company_Code + "' ";
            sql += " and mbl.rec_branch_code  = '" + XmlLib.Branch_Code + "' ";
            sql += " and job.rec_category  = 'SEA EXPORT' ";
            sql += " and mbl.hbl_agent_id = '" + XmlLib.Agent_Id + "' ";
            if (XmlLib.HBL_BL_NOS.Length > 0)
            {
                sql += " and hbl.hbl_bl_no in (" + XmlLib.HBL_BL_NOS + ")";
            }
            else if (XmlLib.MBL_IDS.Length > 0)
            {
                sql += " and mbl.hbl_pkid in (" + XmlLib.MBL_IDS + ")";
            }
            else
            {
                // sql += " and (sysdate - mbl.hbl_pol_etd) between 2 and  60 ";
                sql += " and (((sysdate - mbl.hbl_pol_etd) between 2 and  60) or (sysdate - mbl.hbl_prealert_date) <= 5 ) ";
            }

            Dt = new DataTable();
            Dt = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow Dr in Dt.Rows)
            {
                if (Dr["Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new UnLoCodeMessageUnLoCodeRecord();
                    Record.Sequence = nCtr.ToString();
                    Record.Action = UnLoCodeMessageUnLoCodeRecordAction.Replace;
                    Record.UnLoCode = Dr["Code"].ToString();
                    Record.Name = Dr["Name"].ToString();
                    aList.Add(Record);
                }
            }

            Con_Oracle.CloseConnection();

            Total_Records = nCtr;
            return (UnLoCodeMessageUnLoCodeRecord[])aList.ToArray(typeof(UnLoCodeMessageUnLoCodeRecord));
        }
    }
}