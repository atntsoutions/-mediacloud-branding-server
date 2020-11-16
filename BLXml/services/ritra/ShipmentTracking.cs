using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.ST;

namespace BLXml
{
    public partial class shipmentTracking : BLXml.models.XmlRoot
    {
        DBConnection Con_Oracle = null;
        public override void Generate()
        {
            this.MODULE_ID = "ST";
            this.FileName = "ST";
            if (XmlLib.Agent_Name == "RITRA")
            {
                this.MODULE_ID = "STCE";
                this.FileName = "STCE";
            }
            ShipmentTrackingMessage MyList = new ShipmentTrackingMessage();

            MyList.ShipmentEvents = GetRecords();
            if (Total_Records > 0)
            {
                MyList.MessageInfo = GetMessageInfo();
                if (XmlLib.SaveInTempFolder)
                {
                    XmlLib.FolderId = Guid.NewGuid().ToString().ToUpper();
                    XmlLib.File_Type = "XML";
                    XmlLib.File_Category = "STCE";
                    XmlLib.File_Processid = "STCE" + this.MessageNumber;
                    XmlLib.File_Display_Name = FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlLib.File_Name = Lib.GetFileName(XmlLib.report_folder, XmlLib.FolderId, XmlLib.File_Display_Name);
                    XmlSerializer serializer =
                          new XmlSerializer(typeof(ShipmentTrackingMessage));
                    TextWriter writer = new StreamWriter(XmlLib.File_Name);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
                else
                {

                    XmlLib.CreateSentFolder();
                    FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlSerializer serializer =
                          new XmlSerializer(typeof(ShipmentTrackingMessage));
                    TextWriter writer = new StreamWriter(FileName);
                    serializer.Serialize(writer, MyList);
                    writer.Close();
                }
            }
        }
        public ShipmentTrackingMessageMessageInfo GetMessageInfo()
        {
            ShipmentTrackingMessageMessageInfo VMInfo = new ShipmentTrackingMessageMessageInfo();
            this.MessageNumber = XmlLib.GetNewMessageNumber();
            VMInfo.MessageNumber = this.MessageNumber;
            VMInfo.MessageSender = XmlLib.messageSenderField;
            VMInfo.MessageRecipient = XmlLib.messageRecipientField;
            VMInfo.CreatedDateTime = XmlLib.GetCreatedDate();
            VMInfo.CreatedDateTimeSpecified = true;
            return VMInfo;
        }
        public ShipmentTrackingMessageShipmentEvents[] GetRecords()
        {
            string mID = "";
            int nCtr = 0;
            DataTable Dt = null;
            DataRow Dr_Job = null;
            string sql = "";
            string str = "";
            System.Collections.ArrayList aList = new ArrayList();
            ShipmentTrackingMessageShipmentEvents Record;

            /*
            sql += " select   ";
            sql += " HBLS_HBL_ID,JOB_ID,HBL_NO,HBLS_BL_NO,HBL_YEAR,";
            sql += " OPR_SBILL_NO, OPR_SBILL_DATE, " ;
            sql += " OPR_GR_NUMBER,OPR_GR_DATE, ";
            sql += " JOB_BOOKING_DATE,OPR_CARGO_RECEIVED_ON,OPR_CLEARED_DATE, ";
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
            sql += " HBLS_POFD_ETA_CONF, HBLS_POFD_ETA,";
            sql += " Liner_Code, ";
            sql += " Liner_Name, ";
            sql += " Agent_Code, ";
            sql += " Agent_Name, ";
            sql += " PLACE_CODE,PLACE_NAME ";
            sql += " from TABLE_XMLEDI  "; ;
            sql += " order by hbl_no";
            */

            sql = " select hbl.hbl_pkid,hbl.hbl_year,hbl.hbl_bl_no ";
            sql += " ,vsl.param_code as vessel_code,vsl.param_name as vessel_name,trk.trk_parent_id,trk.trk_order ";
            sql += " ,trk.trk_voyage as vessel_voyage,trk.trk_pol_etd as vessel_etd,trk.trk_pod_eta as vessel_eta";
            sql += " ,pol.param_code as pol_code,pol.param_name as pol_name";
            sql += " ,pod.param_code as pod_code,pod.param_name as pod_name";
            sql += " ,pofd.param_code as pofd_code,pofd.param_name as pofd_name,mbl.hbl_pofd_eta as pofd_eta, mbl.hbl_pofd_eta_confirm as pofd_eta_confirm";
            sql += " ,lnr.param_code  as liner_code,lnr.param_name as liner_name ";
            sql += " from hblm mbl";
            sql += " inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
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
                //sql += " and (sysdate - mbl.hbl_pol_etd) between 2 and  60 ";
                sql += " and (((sysdate - mbl.hbl_pol_etd) between 2 and  60) or (sysdate - mbl.hbl_prealert_date) <= 5 ) ";
            }
            sql += " order by hbl.hbl_no";

            Dt = new DataTable();
            Con_Oracle = new DBConnection();
            Dt = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow Dr in Dt.Rows)
            {
                if (mID != Dr["hbl_pkid"].ToString())
                {
                    mID = Dr["hbl_pkid"].ToString();
                   
                    Dr_Job = GetJobDetais(Dr["hbl_pkid"].ToString());

                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEvents();
                    Record.MemberCode = XmlLib.memberCode;
                    Record.Action = ShipmentTrackingMessageShipmentEventsAction.Replace;
                    Record.STSeq = nCtr.ToString();
                    str = Dr["HBL_YEAR"].ToString();
                    str += (Dr_Job != null ? Dr_Job["job_docno"].ToString() : "");
                    Record.BookingNumber = str;
                   // Record.BookingNumber = Dr["HBL_YEAR"].ToString() + Dr["JOB_ID"].ToString();
                    Record.HouseBLNumber = Dr["HBL_BL_NO"].ToString();
                    Record.ContainerNumber = "*ALL";
                    Record.ShipmentEvent = getShipmentEvents(Dr, Dt, Dr_Job);
                    aList.Add(Record);
                }
            }
            Total_Records = nCtr;
            return (ShipmentTrackingMessageShipmentEvents[])aList.ToArray(typeof(ShipmentTrackingMessageShipmentEvents));
        }

        private DataRow GetJobDetais(string hbl_id)
        {
            DataRow Dr_target = null;
            string sql = "";
            DataTable Dt = null;

            sql = " select job_docno,job_date, rcpt.param_code  as place_code ";
            sql += "   ,opr_cargo_received_on,opr_cleared_date";
            sql += "   from jobm job";
            sql += "   left join param rcpt on job.job_place_receipt_id = rcpt.param_pkid";
            sql += "   left join joboperationsm opr on job.job_pkid = opr.opr_job_id";
            sql += "  where jobs_hbl_id = '" + hbl_id + "'";
            sql += "  order by job.rec_created_date ";
            Dt = new DataTable();
            Con_Oracle = new DBConnection();
            Dt = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            if (Dt.Rows.Count > 0)
                Dr_target = Dt.Rows[0];
            return Dr_target;
        }

        public ShipmentTrackingMessageShipmentEventsShipmentEvent[] getShipmentEvents(DataRow dRow, DataTable Dt, DataRow jobRow)
        {
            string sql = "";
            int nCtr = 0;
            System.Collections.ArrayList aList = new ArrayList();
            ShipmentTrackingMessageShipmentEventsShipmentEvent Record;
            ShipmentTrackingMessageShipmentEventsShipmentEventVesselVoyage Vessel;
            DataRow Dr = dRow;

            if (jobRow != null)
            {
                if (jobRow["job_date"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.BKG;
                    Record.EventDateTime = (DateTime)jobRow["job_date"];
                    Record.Location = jobRow["PLACE_CODE"].ToString();
                    aList.Add(Record);
                }
                if (jobRow["OPR_CARGO_RECEIVED_ON"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.RCV;
                    Record.EventDateTime = (DateTime)jobRow["OPR_CARGO_RECEIVED_ON"];
                    Record.Location = jobRow["PLACE_CODE"].ToString();
                    aList.Add(Record);
                }
                if (jobRow["OPR_CLEARED_DATE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.ECC;
                    Record.EventDateTime = (DateTime)jobRow["OPR_CLEARED_DATE"];
                    Record.Location = jobRow["PLACE_CODE"].ToString();
                    aList.Add(Record);
                }
            }
            string ParentId = "111";
            foreach (DataRow dr in Dt.Select("hbl_pkid='" + Dr["hbl_pkid"].ToString() + "'", "trk_parent_id,trk_order"))
            {
                //if (ParentId != dr["trk_parent_id"].ToString())
                //{
                //ParentId = dr["trk_parent_id"].ToString();
                if (dr["VESSEL_CODE"].ToString().Trim().Length > 0)
                {
                    if (ParentId == "111")
                    {
                        nCtr++;
                        Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Record.EventCode = eventCode.ONB;
                        if (!dr["VESSEL_ETD"].Equals(DBNull.Value))
                            Record.EventDateTime = (DateTime)dr["VESSEL_ETD"];
                        Record.Location = XmlLib.GetPortCode(dr["POL_CODE"].ToString());
                        aList.Add(Record);
                        ParentId = "";
                    }

                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.TSP;
                    if (!dr["VESSEL_ETD"].Equals(DBNull.Value))
                        Record.EventDateTime = (DateTime)dr["VESSEL_ETD"];

                    Vessel = new ShipmentTrackingMessageShipmentEventsShipmentEventVesselVoyage();
                  //  Vessel.VesselCode = dr["VESSEL_CODE"].ToString();
                    Vessel.VesselCode = dr["VESSEL_NAME"].ToString();
                    Vessel.VoyageNumber = dr["VESSEL_VOYAGE"].ToString();
                    Vessel.FromPort = XmlLib.GetPortCode(dr["POL_CODE"].ToString());
                    Vessel.FromPort_Name = dr["POL_NAME"].ToString();
                    Record.Location = XmlLib.GetPortCode(dr["POL_CODE"].ToString());
                    Vessel.ToPort = XmlLib.GetPortCode(dr["POD_CODE"].ToString());
                    Vessel.ToPort_Name = dr["POD_NAME"].ToString();

                    if (!dr["VESSEL_ETD"].Equals(DBNull.Value))
                    {
                        Vessel.ETD = (DateTime)dr["VESSEL_ETD"];
                        Vessel.ETDSpecified = true;
                    }
                    if (!dr["VESSEL_ETA"].Equals(DBNull.Value))
                    {
                        Vessel.ETA = (DateTime)dr["VESSEL_ETA"];
                        Vessel.ETASpecified = true;
                    }
                    Vessel.MemberCode = XmlLib.memberCode;
                    if (dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Vessel.LinerCode = dr["LINER_CODE"].ToString();
                        Vessel.LinerCode_Name = dr["LINER_NAME"].ToString();
                    }
                    Record.VesselVoyage = Vessel;
                    aList.Add(Record);
                }
                //}
            }
            if (Dr["pofd_eta_confirm"].ToString().Trim() == "Y")
            {
                nCtr++;
                Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                Record.EventCode = eventCode.ACW;
                if (!Dr["POFD_ETA"].Equals(DBNull.Value))
                    Record.EventDateTime = (DateTime)Dr["POFD_ETA"];
                Record.Location = XmlLib.GetPortCode(Dr["POFD_CODE"].ToString());
                aList.Add(Record);
            }

            Total_Records = nCtr;



            /*
            if (true)
            {
                //DataRow Dr = Dt.Rows[0];
                DataRow Dr = dRow;
                if (Dr["JOB_BOOKING_DATE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.BKG;
                    Record.EventDateTime = (DateTime)Dr["JOB_BOOKING_DATE"];
                    Record.Location = Dr["PLACE_CODE"].ToString();
                    aList.Add(Record);
                }
                if (Dr["OPR_CARGO_RECEIVED_ON"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.RCV;
                    Record.EventDateTime = (DateTime)Dr["OPR_CARGO_RECEIVED_ON"];
                    Record.Location = Dr["PLACE_CODE"].ToString();
                    aList.Add(Record);
                }
                if (Dr["OPR_CLEARED_DATE"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.ECC;
                    Record.EventDateTime = (DateTime)Dr["OPR_CLEARED_DATE"];
                    Record.Location = Dr["PLACE_CODE"].ToString();
                    aList.Add(Record);
                }
                if (Dr["VESSEL1_CODE"].ToString().Trim().Length > 0) //&& Dr["JOB_FEEDER_VESS_CONF"].ToString().Trim().Length > 0
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.ONB;
                    if (!Dr["VESSEL1_ETD"].Equals(DBNull.Value))
                        Record.EventDateTime = (DateTime)Dr["VESSEL1_ETD"];
                    Record.Location =XmlLib.GetPortCode(  Dr["POL_CODE"].ToString());
                    
                    aList.Add(Record);

                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.TSP;
                    if (!Dr["VESSEL1_ETD"].Equals(DBNull.Value))
                        Record.EventDateTime = (DateTime)Dr["VESSEL1_ETD"];
                    Vessel = new ShipmentTrackingMessageShipmentEventsShipmentEventVesselVoyage();
                    Vessel.VesselCode = Dr["VESSEL1_CODE"].ToString();
                    Vessel.VoyageNumber = Dr["VESSEL1_VOYAGE"].ToString();
                    Vessel.FromPort = XmlLib.GetPortCode( Dr["POL_CODE"].ToString());
                    Vessel.FromPort_Name = Dr["POL_NAME"].ToString(); 
                    Record.Location = XmlLib.GetPortCode( Dr["POL_CODE"].ToString());

                    if (Dr["TRANSIT1_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = Dr["TRANSIT1_CODE"].ToString();
                        Vessel.ToPort_Name = Dr["TRANSIT1_NAME"].ToString();
                    }
                    else if (Dr["TRANSIT2_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = "";
                        Vessel.ToPort_Name = "";
                    }
                    else if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = "";
                        Vessel.ToPort_Name = "";
                    }
                    else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = "";
                        Vessel.ToPort_Name = "";
                    }
                    else
                    {
                        Vessel.ToPort = XmlLib.GetPortCode( Dr["POD_CODE"].ToString());
                        Vessel.ToPort_Name = Dr["POD_NAME"].ToString();
                    }


                    if (!Dr["VESSEL1_ETD"].Equals(DBNull.Value))
                    {
                        Vessel.ETD = (DateTime)Dr["VESSEL1_ETD"];
                        Vessel.ETDSpecified = true;
                    }
                    if (!Dr["VESSEL1_ETA"].Equals(DBNull.Value))
                    {
                        Vessel.ETA = (DateTime)Dr["VESSEL1_ETA"];
                        Vessel.ETASpecified = true;
                    }
                    Vessel.MemberCode = Lib.memberCode;
                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Vessel.LinerCode = Dr["LINER_CODE"].ToString();
                        Vessel.LinerCode_Name = Dr["LINER_NAME"].ToString();
                    }
                    Record.VesselVoyage = Vessel; 
                    aList.Add(Record);
                }
                if (Dr["Vessel2_Code"].ToString().Trim().Length > 0) //&& Dr["JOB_FEEDER2_VESS_CONF"].ToString().Trim().Length > 0
                {
                    if (Dr["Vessel1_Code"].ToString().Trim().Length <= 0)
                    {
                        nCtr++;
                        Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Record.EventCode = eventCode.ONB;
                        if (!Dr["VESSEL2_ETD"].Equals(DBNull.Value))
                            Record.EventDateTime = (DateTime)Dr["VESSEL12_ETD"];
                        Record.Location =Lib.GetPortCode( Dr["POL_CODE"].ToString());
                        aList.Add(Record);
                    }
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.TSP;
                    Record.Location = Dr["POL_CODE"].ToString();
                    if (!Dr["VESSEL2_ETD"].Equals(DBNull.Value))
                        Record.EventDateTime = (DateTime)Dr["VESSEL2_ETD"];
                    Vessel = new ShipmentTrackingMessageShipmentEventsShipmentEventVesselVoyage();
                    Vessel.VesselCode = Dr["VESSEL2_CODE"].ToString();
                    Vessel.VoyageNumber = Dr["VESSEL2_VOYAGE"].ToString();
                    Vessel.FromPort =Lib.GetPortCode(  Dr["POL_CODE"].ToString());
                    Vessel.FromPort_Name = Dr["POL_NAME"].ToString();
                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                    {
                        Vessel.FromPort = Dr["TRANSIT1_CODE"].ToString();
                        Vessel.FromPort_Name = Dr["TRANSIT1_NAME"].ToString();
                    }


                    Record.Location = Vessel.FromPort;

                    if (Dr["TRANSIT2_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = Dr["TRANSIT2_CODE"].ToString();
                        Vessel.ToPort_Name = Dr["TRANSIT2_NAME"].ToString();
                    }
                    else if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = "";
                        Vessel.ToPort_Name = "";
                    }
                    else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = "";
                        Vessel.ToPort_Name = "";
                    }
                    else
                    {
                        Vessel.ToPort =Lib.GetPortCode(  Dr["POD_CODE"].ToString());
                        Vessel.ToPort_Name = Dr["POD_NAME"].ToString();
                    }




                    if (!Dr["VESSEL2_ETD"].Equals(DBNull.Value))
                    {
                        Vessel.ETD = (DateTime)Dr["VESSEL2_ETD"];
                        Vessel.ETDSpecified = true;
                    }
                    if (!Dr["VESSEL2_ETA"].Equals(DBNull.Value))
                    {
                        Vessel.ETA = (DateTime)Dr["VESSEL2_ETA"];
                        Vessel.ETASpecified = true;
                    }
                    Vessel.MemberCode = Lib.memberCode;
                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Vessel.LinerCode = Dr["LINER_CODE"].ToString();
                        Vessel.LinerCode_Name = Dr["LINER_NAME"].ToString();
                    }
                    Record.VesselVoyage = Vessel;
                    aList.Add(Record);
                }
                if (Dr["VESSEL3_Code"].ToString().Trim().Length > 0) //&& Dr["JOB_FEEDER3_VESS_CONF"].ToString().Trim().Length > 0
                {
                    if (Dr["VESSEL1_Code"].ToString().Trim().Length <= 0 && Dr["VESSEL2_Code"].ToString().Trim().Length <= 0)
                    {
                        nCtr++;
                        Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Record.EventCode = eventCode.ONB;
                        if (!Dr["VESSEL3_ETD"].Equals(DBNull.Value))
                            Record.EventDateTime = (DateTime)Dr["VESSEL3_ETD"];
                        Record.Location = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                        aList.Add(Record);
                    }

                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.TSP;
                    Record.Location = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    if (!Dr["VESSEL3_ETD"].Equals(DBNull.Value))
                        Record.EventDateTime = (DateTime)Dr["VESSEL3_ETD"];
                    Vessel = new ShipmentTrackingMessageShipmentEventsShipmentEventVesselVoyage();
                    Vessel.VesselCode = Dr["VESSEL3_CODE"].ToString();
                    Vessel.VoyageNumber = Dr["VESSEL3_VOYAGE"].ToString();
                    Vessel.FromPort = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    Vessel.FromPort_Name = Dr["POL_NAME"].ToString();
                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                    {
                        Vessel.FromPort = Dr["TRANSIT1_CODE"].ToString();
                        Vessel.FromPort_Name = Dr["TRANSIT1_NAME"].ToString();
                    }
                    if (Dr["TRANSIT2_CODE"].ToString().Trim().Length > 0)
                    {
                        Vessel.FromPort = Dr["TRANSIT2_CODE"].ToString();
                        Vessel.FromPort_Name = Dr["TRANSIT2_NAME"].ToString();
                    }


                    Record.Location = Vessel.FromPort;

                    if (Dr["TRANSIT3_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = Dr["TRANSIT3_CODE"].ToString();
                        Vessel.ToPort_Name = Dr["TRANSIT3_NAME"].ToString();
                    }
                    else if (Dr["VESSEL4_CODE"].ToString().Length > 0)
                    {
                        Vessel.ToPort = "";
                        Vessel.ToPort_Name = "";
                    }
                    else
                    {
                        Vessel.ToPort = Lib.GetPortCode( Dr["POD_CODE"].ToString());
                        Vessel.ToPort_Name = Dr["POD_NAME"].ToString();
                    }



                    if (!Dr["VESSEL3_ETD"].Equals(DBNull.Value))
                    {
                        Vessel.ETD = (DateTime)Dr["VESSEL3_ETD"];
                        Vessel.ETDSpecified = true;
                    }
                    if (!Dr["VESSEL3_ETA"].Equals(DBNull.Value))
                    {
                        Vessel.ETA = (DateTime)Dr["VESSEL3_ETA"];
                        Vessel.ETASpecified = true;
                    }
                    Vessel.MemberCode = Lib.memberCode;
                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Vessel.LinerCode = Dr["LINER_CODE"].ToString();
                        Vessel.LinerCode_Name = Dr["LINER_NAME"].ToString();
                    }
                    Record.VesselVoyage = Vessel;
                    aList.Add(Record);
                }
                if (Dr["VESSEL4_Code"].ToString().Trim().Length > 0) //&& Dr["JOB_MOT_VESS_CONF"].ToString().Trim().Length > 0
                {
                    if (Dr["VESSEL1_Code"].ToString().Trim().Length <= 0 && Dr["VESSEL2_Code"].ToString().Trim().Length <= 0 && Dr["VESSEL3_Code"].ToString().Trim().Length <= 0)
                    {
                        nCtr++;
                        Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Record.EventCode = eventCode.ONB;
                        if (!Dr["VESSEL4_ETD"].Equals(DBNull.Value))
                            Record.EventDateTime = (DateTime)Dr["VESSEL4_ETD"];
                        Record.Location =Lib.GetPortCode(  Dr["POL_CODE"].ToString());
                        aList.Add(Record);
                    }
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.TSP;
                    Record.Location =Lib.GetPortCode(  Dr["POL_CODE"].ToString());
                    if (!Dr["VESSEL4_ETD"].Equals(DBNull.Value))
                        Record.EventDateTime = (DateTime)Dr["VESSEL4_ETD"];
                    Vessel = new ShipmentTrackingMessageShipmentEventsShipmentEventVesselVoyage();
                    Vessel.VesselCode = Dr["VESSEL4_CODE"].ToString();
                    Vessel.VoyageNumber = Dr["VESSEL4_VOYAGE"].ToString();
                    Vessel.FromPort = Lib.GetPortCode( Dr["POL_CODE"].ToString());
                    Vessel.FromPort_Name = Dr["POL_NAME"].ToString();
                    if (Dr["TRANSIT1_CODE"].ToString().Trim().Length > 0)
                    {
                        Vessel.FromPort = Dr["TRANSIT1_CODE"].ToString();
                        Vessel.FromPort_Name = Dr["TRANSIT1_NAME"].ToString();
                    }
                    if (Dr["TRANSIT2_CODE"].ToString().Trim().Length > 0)
                    {
                        Vessel.FromPort = Dr["TRANSIT2_CODE"].ToString();
                        Vessel.FromPort_Name = Dr["TRANSIT2_NAME"].ToString();
                    }
                    if (Dr["TRANSIT3_CODE"].ToString().Trim().Length > 0)
                    {
                        Vessel.FromPort = Dr["TRANSIT3_CODE"].ToString();
                        Vessel.FromPort_Name = Dr["TRANSIT3_NAME"].ToString();
                    }

                    Record.Location = Vessel.FromPort;

                    Vessel.ToPort = Lib.GetPortCode( Dr["POD_CODE"].ToString());
                    Vessel.ToPort_Name = Dr["POD_NAME"].ToString();

                    
                    if (!Dr["VESSEL4_ETD"].Equals(DBNull.Value))
                    {
                        Vessel.ETD = (DateTime)Dr["VESSEL4_ETD"];
                        Vessel.ETDSpecified = true;
                    }
                    if (!Dr["VESSEL4_ETA"].Equals(DBNull.Value))
                    {
                        Vessel.ETA = (DateTime)Dr["VESSEL4_ETA"];
                        Vessel.ETASpecified = true;
                    }
                    Vessel.MemberCode = Lib.memberCode;

                    if (Dr["LINER_NAME"].ToString().Trim().Length > 0)
                    {
                        Vessel.LinerCode = Dr["LINER_CODE"].ToString();
                        Vessel.LinerCode_Name = Dr["LINER_NAME"].ToString();
                    }
                    Record.VesselVoyage = Vessel;
                    aList.Add(Record);
                }
                if (Dr["HBLS_POFD_ETA_CONF"].ToString().Trim().Length > 0)
                {
                    nCtr++;
                    Record = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                    Record.EventCode = eventCode.ACW;
                    if (!Dr["HBLS_POFD_ETA"].Equals(DBNull.Value))
                        Record.EventDateTime = (DateTime)Dr["HBLS_POFD_ETA"];
                    Record.Location = XmlLib.GetPortCode( Dr["POFD_CODE"].ToString());
                    
                    aList.Add(Record);
                }

                Total_Records = nCtr;
            }*/
            return (ShipmentTrackingMessageShipmentEventsShipmentEvent[])aList.ToArray(typeof(ShipmentTrackingMessageShipmentEventsShipmentEvent));
        }
    }
}