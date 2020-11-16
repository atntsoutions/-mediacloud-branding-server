using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.VS;

namespace BLXml
{
    public partial class VesselSchedule : BLXml.models.XmlRoot
    {
        DBConnection Con_Oracle = null;
        public override void Generate()
        {
            this.MODULE_ID = "VS";
            this.FileName = "VS";
            if (XmlLib.Agent_Name == "RITRA")
            {
                this.MODULE_ID = "VSCE";
                this.FileName = "VSCE";
            }
            VesselScheduleMessage MyList = new VesselScheduleMessage();
            
            MyList.VesselScheduleRecord = GetRecords();
            if (Total_Records > 0)
            {
                MyList.MessageInfo = GetMessageInfo();
                if (XmlLib.SaveInTempFolder)
                {
                    XmlLib.FolderId = Guid.NewGuid().ToString().ToUpper();
                    XmlLib.File_Type = "XML";
                    XmlLib.File_Category = "VSCE";
                    XmlLib.File_Processid = "VSCE" + this.MessageNumber;
                    XmlLib.File_Display_Name = FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlLib.File_Name = Lib.GetFileName(XmlLib.report_folder, XmlLib.FolderId, XmlLib.File_Display_Name);
                    XmlSerializer serializer =
                         new XmlSerializer(typeof(VesselScheduleMessage));
                    TextWriter writer = new StreamWriter(XmlLib.File_Name);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
                else
                {
                    XmlLib.CreateSentFolder();
                    FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlSerializer serializer =
                          new XmlSerializer(typeof(VesselScheduleMessage));
                    TextWriter writer = new StreamWriter(FileName);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
            }

        }
        public VesselScheduleMessageMessageInfo GetMessageInfo()
        {
            VesselScheduleMessageMessageInfo VMInfo = new VesselScheduleMessageMessageInfo();
            this.MessageNumber = XmlLib.GetNewMessageNumber();
            VMInfo.MessageNumber = this.MessageNumber;
            VMInfo.MessageSender = XmlLib.messageSenderField;
            VMInfo.MessageRecipient = XmlLib.messageRecipientField;
            VMInfo.CreatedDateTime = XmlLib.GetCreatedDate();
            VMInfo.CreatedDateTimeSpecified = true;
            return VMInfo;
        }
        public VesselScheduleMessageVesselScheduleRecord[] GetRecords()
        {
            int nCtr = 0;
            DataTable Dt_New = null;
            string sql = "";
            string sRemarks = "";

            ArrayList aList = new ArrayList();
            VesselScheduleMessageVesselScheduleRecord Record;
            VesselScheduleMessageVesselScheduleRecordPOD Pod;
            VesselScheduleMessageVesselScheduleRecordPOL Pol;
            VesselScheduleMessageVesselScheduleRecordLiner Liner;

            sql = "";
           
            /*
            sql += " select   ";
            sql += " VESSEL1_CODE, VESSEL2_CODE, VESSEL3_CODE,VESSEL4_CODE, ";
            sql += " VESSEL1_NAME,VESSEL2_NAME,VESSEL3_NAME,VESSEL4_NAME, ";
            sql += " Vessel1_voyage,Vessel2_voyage,Vessel3_voyage,Vessel4_voyage, ";
            sql += " TRANSIT1_CODE,TRANSIT2_CODE,TRANSIT3_CODE,TRANSIT4_CODE ";
            sql += " TRANSIT1_NAME,TRANSIT2_NAME,TRANSIT3_NAME,TRANSIT4_NAME, ";
            sql += " Vessel1_Etd,Vessel1_Eta, ";
            sql += " Vessel2_Etd,Vessel2_Eta, ";
            sql += " Vessel3_Etd,Vessel3_Eta, ";
            sql += " Vessel4_Etd,Vessel4_Eta, ";
            sql += " POFD_CODE,POFD_NAME,POD_CODE,POD_NAME,POL_CODE,POL_NAME, ";
            sql += " Liner_Code, ";
            sql += " Liner_Name, ";
            sql += " Agent_Code, ";
            sql += " Agent_Name ";
            sql += " from TABLE_XMLEDI ";
            //sql += " where (AGENT_ID = '" + Lib.Agent_Id + "') ";
            */

            sql = " select mbl.hbl_no as MBL_BOOKNO,mbl.rec_branch_code, vsl.param_code as vessel_code,vsl.param_name as vessel_name ";
            sql += " ,trk.trk_voyage as vessel_voyage,trk.trk_pol_etd as vessel_etd,trk.trk_pod_eta as vessel_eta";
            sql += " ,pol.param_code as pol_code,pol.param_name as pol_name";
            sql += " ,pod.param_code as pod_code,pod.param_name as pod_name";
            sql += " ,pofd.param_code as pofd_code,pofd.param_name as pofd_name";
            sql += " ,lnr.param_code  as liner_code,lnr.param_name as liner_name ";
            sql += " from hblm mbl";
            sql += " inner join trackingm trk on mbl.hbl_pkid = trk.trk_parent_id";
            sql += " left join param vsl on trk_vsl_id= vsl.param_pkid ";
            sql += " left join param pol on trk.trk_pol_id = pol.param_pkid";
            sql += " left join param pod on trk.trk_pod_id = pod.param_pkid";
            sql += " left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid";
            sql += " left join param lnr on mbl.hbl_carrier_id = lnr.param_pkid";
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
            sql += " order by trk_parent_id,trk_order";

            Dt_New = new DataTable();
            Con_Oracle = new DBConnection();
            Dt_New = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Dr in Dt_New.Rows)
            {
                if (Dr["VESSEL_CODE"].ToString().Trim().Length > 0)
                {
                    sRemarks = "";
                    if (Dr["VESSEL_ETD"].Equals(DBNull.Value))
                        sRemarks +="| Error:ETD ";
                    if (Dr["POL_CODE"].Equals(DBNull.Value))
                        sRemarks += "| Error:POL ";
                    if (Dr["VESSEL_ETA"].Equals(DBNull.Value))
                        sRemarks += "| Error:ETA ";
                    if (Dr["POD_CODE"].Equals(DBNull.Value))
                        sRemarks += "| Error:POD ";
                    if (sRemarks != "")
                    {
                        sRemarks = String.Concat("MBLSL#:", Dr["MBL_BOOKNO"].ToString(), "| BR:", Dr["REC_BRANCH_CODE"].ToString(), sRemarks);
                        XmlLib.AddToErrorList("TRACKING", sRemarks);
                    }

                    nCtr++;
                    Record = new VesselScheduleMessageVesselScheduleRecord();
                    Record.MemberCode = XmlLib.memberCode;
                    Record.Sequence = nCtr.ToString();
                    Record.Action = VesselScheduleMessageVesselScheduleRecordAction.Replace;
                   // Record.VesselCode = Dr["VESSEL_CODE"].ToString();
                    Record.VesselCode = Dr["VESSEL_NAME"].ToString();
                    Record.VoyageNo = Dr["VESSEL_VOYAGE"].ToString();

                    Pod = new VesselScheduleMessageVesselScheduleRecordPOD();
                    Pol = new VesselScheduleMessageVesselScheduleRecordPOL();

                    if (!Dr["VESSEL_ETD"].Equals(DBNull.Value))
                    {
                        Pol.ETD = (DateTime)Dr["VESSEL_ETD"];
                        Pol.Value = XmlLib.GetPortCode(Dr["POL_CODE"].ToString());
                    }
                    
                    if (!Dr["VESSEL_ETA"].Equals(DBNull.Value))
                    {
                        Pod.ETA = (DateTime)Dr["VESSEL_ETA"];
                        Pod.ETASpecified = true;
                        if (Dr["POD_CODE"].ToString().Length > 0)
                            Pod.Value = Dr["POD_CODE"].ToString();
                        else
                            Pod.Value = XmlLib.GetPortCode(Dr["POD_CODE"].ToString());
                    }

                    Record.POD = Pod;
                    Record.POL = Pol;
                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Liner = new VesselScheduleMessageVesselScheduleRecordLiner();
                        Liner.LinerName = Dr["LINER_NAME"].ToString();
                        Liner.Value = Dr["LINER_CODE"].ToString();
                        Liner.SCAC = "";
                        Record.Liner = Liner;
                    }
                    aList.Add(Record);
                }


                /*
                if (Dr["VESSEL1_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new VesselScheduleMessageVesselScheduleRecord();
                    Record.MemberCode = XmlLib.memberCode;
                    Record.Sequence = nCtr.ToString();
                    Record.Action = VesselScheduleMessageVesselScheduleRecordAction.Replace;
                    Record.VesselCode = Dr["VESSEL1_CODE"].ToString();
                    Record.VoyageNo = Dr["VESSEL1_VOYAGE"].ToString();

                    Pod = new VesselScheduleMessageVesselScheduleRecordPOD();
                    Pol = new VesselScheduleMessageVesselScheduleRecordPOL();

                    if (!Dr["VESSEL1_ETD"].Equals(DBNull.Value))
                    {
                        Pol.ETD = (DateTime)Dr["VESSEL1_ETD"];
                        Pol.Value = XmlLib.GetPortCode( Dr["POL_CODE"].ToString());
                    }
                    if (!Dr["VESSEL1_ETA"].Equals(DBNull.Value))
                    {
                        Pod.ETA = (DateTime)Dr["VESSEL1_ETA"];
                        Pod.ETASpecified = true;
                        if (Dr["TRANSIT1_CODE"].ToString().Length > 0)
                            Pod.Value = Dr["TRANSIT1_CODE"].ToString();
                        else if (Dr["TRANSIT2_CODE"].ToString().Length > 0)
                            Pod.Value = "";
                        else if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                            Pod.Value = "";
                        else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                            Pod.Value = "";
                        else
                            Pod.Value =XmlLib.GetPortCode(  Dr["POD_CODE"].ToString());
                    }

                    Record.POD = Pod;
                    Record.POL = Pol;
                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Liner = new VesselScheduleMessageVesselScheduleRecordLiner();
                        Liner.LinerName = Dr["LINER_NAME"].ToString();
                        Liner.Value = Dr["LINER_CODE"].ToString();
                        Liner.SCAC = "";
                        Record.Liner = Liner;
                    }
                    aList.Add(Record);
                }
                if (Dr["VESSEL2_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new VesselScheduleMessageVesselScheduleRecord();
                    Record.MemberCode = Lib.memberCode;
                    Record.Sequence = nCtr.ToString();
                    Record.Action = VesselScheduleMessageVesselScheduleRecordAction.Replace;
                    Record.VesselCode = Dr["VESSEL2_CODE"].ToString();
                    Record.VoyageNo = Dr["VESSEL2_VOYAGE"].ToString();

                    Pod = new VesselScheduleMessageVesselScheduleRecordPOD();
                    Pol = new VesselScheduleMessageVesselScheduleRecordPOL();
                    if (!Dr["VESSEL2_ETD"].Equals(DBNull.Value))
                    {
                        Pol.ETD = (DateTime)Dr["VESSEL2_ETD"];
                        Pol.Value = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    }


                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                        Pol.Value = Dr["TRANSIT1_CODE"].ToString();
                    if (!Dr["VESSEL2_ETA"].Equals(DBNull.Value))
                    {
                        Pod.ETA = (DateTime)Dr["VESSEL2_ETA"];
                        Pod.ETASpecified = true;
                        if (Dr["TRANSIT2_CODE"].ToString().Length > 0)
                            Pod.Value = Dr["TRANSIT2_CODE"].ToString();
                        else if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                            Pod.Value = "";
                        else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                            Pod.Value = "";
                        else
                            Pod.Value = Lib.GetPortCode( Dr["POD_CODE"].ToString());
                    }

                    Record.POD = Pod;
                    Record.POL = Pol;

                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Liner = new VesselScheduleMessageVesselScheduleRecordLiner();
                        Liner.LinerName = Dr["LINER_NAME"].ToString();
                        Liner.Value = Dr["LINER_CODE"].ToString();
                        Liner.SCAC = "";
                        Record.Liner = Liner;
                    }
                    aList.Add(Record);
                }
                if (Dr["vessel3_Code"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new VesselScheduleMessageVesselScheduleRecord();
                    Record.MemberCode = Lib.memberCode;
                    Record.Sequence = nCtr.ToString();
                    Record.Action = VesselScheduleMessageVesselScheduleRecordAction.Replace;

                    Record.VesselCode = Dr["VESSEL3_CODE"].ToString();
                    Record.VoyageNo = Dr["VESSEL3_VOYAGE"].ToString();

                    Pod = new VesselScheduleMessageVesselScheduleRecordPOD();
                    Pol = new VesselScheduleMessageVesselScheduleRecordPOL();
                    if (!Dr["VESSEL3_ETD"].Equals(DBNull.Value))
                    {
                        Pol.ETD = (DateTime)Dr["VESSEL3_ETD"];
                        Pol.Value = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    }
                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                        Pol.Value = Dr["TRANSIT1_CODE"].ToString();
                    if (Dr["TRANSIT2_CODE"].ToString().Trim().Length > 0)
                        Pol.Value = Dr["TRANSIT2_CODE"].ToString();
                    if (!Dr["VESSEL3_ETA"].Equals(DBNull.Value))
                    {
                        Pod.ETA = (DateTime)Dr["VESSEL3_ETA"];
                        Pod.ETASpecified = true;
                        if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                            Pod.Value = Dr["TRANSIT3_CODE"].ToString();
                        else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                            Pod.Value = "";
                        else
                            Pod.Value = Lib.GetPortCode( Dr["POD_CODE"].ToString());
                    }

                    Record.POD = Pod;
                    Record.POL = Pol;

                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Liner = new VesselScheduleMessageVesselScheduleRecordLiner();
                        Liner.LinerName = Dr["LINER_NAME"].ToString();
                        Liner.Value = Dr["LINER_CODE"].ToString();
                        Liner.SCAC = "";
                        Record.Liner = Liner;
                    }
                    aList.Add(Record);
                }
                if (Dr["VESSEL4_CODE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new VesselScheduleMessageVesselScheduleRecord();
                    Record.MemberCode = Lib.memberCode;
                    Record.Sequence = nCtr.ToString();
                    Record.Action = VesselScheduleMessageVesselScheduleRecordAction.Replace;
                    Record.VesselCode = Dr["VESSEL4_CODE"].ToString();
                    Record.VoyageNo = Dr["VESSEL4_VOYAGE"].ToString();

                    Pod = new VesselScheduleMessageVesselScheduleRecordPOD();
                    Pol = new VesselScheduleMessageVesselScheduleRecordPOL();
                    if (!Dr["VESSEL4_ETD"].Equals(DBNull.Value))
                    {
                        Pol.ETD = (DateTime)Dr["VESSEL4_ETD"];
                        Pol.Value = Lib.GetPortCode( Dr["JOB_POL"].ToString());
                    }
                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                        Pol.Value = Dr["TRANSIT1_CODE"].ToString();
                    if (Dr["TRANSIT2_CODE"].ToString().Trim().Length > 0)
                        Pol.Value = Dr["TRANSIT2_CODE"].ToString();
                    if (Dr["TRANSIT3_CODE"].ToString().Trim().Length > 0)
                        Pol.Value = Dr["TRANSIT3_CODE"].ToString();

                    if (!Dr["VESSEL4_ETA"].Equals(DBNull.Value))
                    {
                        Pod.ETA = (DateTime)Dr["VESSEL4_ETA"];
                        Pod.Value =Lib.GetPortCode(  Dr["POD_CODE"].ToString());
                        Pod.ETASpecified = true;
                    }

                    Record.POD = Pod;
                    Record.POL = Pol;

                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Liner = new VesselScheduleMessageVesselScheduleRecordLiner();
                        Liner.LinerName = Dr["LINER_NAME"].ToString();
                        Liner.Value = Dr["LINER_CODE"].ToString();
                        Liner.SCAC = "";
                        Record.Liner = Liner;
                    }
                    aList.Add(Record);
                }*/
            }
            Total_Records = nCtr;
            return (VesselScheduleMessageVesselScheduleRecord[])aList.ToArray(typeof(VesselScheduleMessageVesselScheduleRecord));
        }

    }
}
