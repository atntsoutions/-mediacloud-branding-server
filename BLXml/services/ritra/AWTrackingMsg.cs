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
using BLXml.models.AirST;

namespace BLXml
{
    public partial class AWTrackingMsg
    {
        private DataTable DT_TRACKING = new DataTable();
        public Boolean IsError = false;
        private ShipmentTrackingMessage TrackMessage = null;
        Dictionary<int, string> ErrorDic = new Dictionary<int, string>();
        private string ErrorValues = "";
        private string sql = "";
        private string MessageNumber = "";
        private int TrackSeq = 0;
        DBConnection Con_Oracle = null;

        public void Generate()
        {
            try
            {
                IsError = false;
                ReadData();
                if (DT_TRACKING.Rows.Count <= 0)
                {
                    IsError = true;
                    return;
                }
                GenerateXmlFiles();
                WriteXmlFiles();
            }
            catch (Exception ex)
            {
                IsError = true;
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

            sql = " select h.hbl_pkid,h.hbl_job_nos as job_id, h.hbl_year as job_year, ";
            sql += " m.hbl_date as mbl_date, trk.trk_voyage as track_flight_no,";
            sql += " trk_pol_etd as track_etd,trk_pod_eta as track_eta,trk_order as track_slno,";
            sql += " lnr.param_code  as track_carrier_code,lnr.param_name as track_carrier_name, ";
            sql += " pol.param_code as track_pol_code,";
            sql += " pol.param_name track_pol_name,";
            sql += " pod.param_code as track_pod_code, ";
            sql += " pod.param_code as track_pod_name ";
            sql += " from hblm m ";
            sql += " inner join hblm h on  (m.hbl_pkid = h.hbl_mbl_id )";
            sql += " inner join trackingm trk on h.hbl_mbl_id = trk.trk_parent_id";
            sql += " left join param lnr on h.hbl_carrier_id = lnr.param_pkid";
            sql += " left join param pol on trk.trk_pol_id = pol.param_pkid";
            sql += " left join param pod on trk.trk_pod_id = pod.param_pkid";
            sql += " " + sWhere;
            sql += " order by trk_order,m.hbl_date ";

            DT_TRACKING = new DataTable();
            Con_Oracle = new DBConnection();
            DT_TRACKING = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
        }
        private void GenerateXmlFiles()
        {
            TrackMessage = new ShipmentTrackingMessage();
            TrackMessage.Items = Generate_TrackingMessage();
        }
        private object[] Generate_TrackingMessage()
        {
            object[] Items = null;
            int iTotRows = 0;
            int ArrIndex = 0;
            try
            {
                TrackSeq = 0;
                DataTable DistinctTrack = DT_TRACKING.DefaultView.ToTable(true, "hbl_pkid", "mbl_date");
                iTotRows = DistinctTrack.Rows.Count;
                Items = new object[iTotRows + 1];
                Items[ArrIndex++] = Generate_MessageInfo();
                foreach (DataRow dr in DistinctTrack.Select("1=1", "mbl_date"))
                {
                    TrackSeq++;
                    Items[ArrIndex++] = Generate_ShipmentEvents(dr["hbl_pkid"].ToString());
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
        private ShipmentTrackingMessageMessageInfo Generate_MessageInfo()
        {
            ShipmentTrackingMessageMessageInfo Rec = null;
            try
            {
                this.MessageNumber = XmlLib.GetNewMessageNumber();
                Rec = new ShipmentTrackingMessageMessageInfo();
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


        private ShipmentTrackingMessageShipmentEvents Generate_ShipmentEvents(string HBL_ID)
        {
            ShipmentTrackingMessageShipmentEvents Rec = null;
            string bkNo = "";
            try
            {
                foreach (DataRow Dr in DT_TRACKING.Select("hbl_pkid ='" + HBL_ID + "'", "track_slno"))
                {
                    bkNo = Dr["JOB_YEAR"].ToString() + GetSingleJobNo(Dr["JOB_ID"].ToString());
                    break;
                }
                Rec = new ShipmentTrackingMessageShipmentEvents();
                Rec.Action = "Replace";
                Rec.STSeq = TrackSeq.ToString();
                Rec.BookingNo = bkNo;
                Rec.MemberCode = XmlLib.messageSenderField;
                Rec.ShipmentEvent = Generate_ShipmentEvent(HBL_ID);
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return Rec;
        }
        private ShipmentTrackingMessageShipmentEventsShipmentEvent[] Generate_ShipmentEvent(string HBL_ID)
        {
            ShipmentTrackingMessageShipmentEventsShipmentEvent Rec = null;
            ShipmentTrackingMessageShipmentEventsShipmentEvent[] mList = null;
            DataRow[] DDrTracks = null;
            int TotTrks = 0;
            int ArrIndex = 0;
            string POLCODE = "";
            string POLNAME = "";
            try
            {
                DDrTracks = DT_TRACKING.Select("hbl_pkid ='" + HBL_ID + "'", "track_slno");
                TotTrks = DDrTracks.Length;

                mList = new ShipmentTrackingMessageShipmentEventsShipmentEvent[TotTrks];
                foreach (DataRow Dr in DDrTracks)
                {
                    if (Dr["track_slno"].ToString() == "1")
                    {
                        POLCODE = Dr["track_pol_code"].ToString();
                        POLNAME = Dr["track_pol_name"].ToString();

                        Rec = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Rec.EventDateTime = ConvertYMDDate(DateTime.Now.ToString());
                        Rec.EventSeq = Dr["track_slno"].ToString();
                        Rec.EventCode = "";
                        Rec.FlightInformation = Generate_FlightInformation(Dr, POLCODE, POLNAME);
                        Rec.Party = Generate_Party();
                        mList[ArrIndex++] = Rec;

                        POLCODE = Dr["track_pod_code"].ToString();
                        POLNAME = Dr["track_pod_name"].ToString();
                    }
                    if (Dr["track_slno"].ToString() == "2")
                    {
                        Rec = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Rec.EventDateTime = ConvertYMDDate(DateTime.Now.ToString());
                        Rec.EventSeq = Dr["track_slno"].ToString();
                        Rec.EventCode = "";
                        Rec.FlightInformation = Generate_FlightInformation(Dr, POLCODE, POLNAME);
                        Rec.Party = Generate_Party();
                        mList[ArrIndex++] = Rec;

                        POLCODE = Dr["track_pod_code"].ToString();
                        POLNAME = Dr["track_pod_name"].ToString();
                    }
                    if (Dr["track_slno"].ToString() == "3")
                    {
                        Rec = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Rec.EventDateTime = ConvertYMDDate(DateTime.Now.ToString());
                        Rec.EventSeq = Dr["track_slno"].ToString();
                        Rec.EventCode = "";
                        Rec.FlightInformation = Generate_FlightInformation(Dr, POLCODE, POLNAME);
                        Rec.Party = Generate_Party();
                        mList[ArrIndex++] = Rec;

                        POLCODE = Dr["track_pod_code"].ToString();
                        POLNAME = Dr["track_pod_name"].ToString();
                    }
                    if (Dr["track_slno"].ToString() == "4")
                    {
                        Rec = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Rec.EventDateTime = ConvertYMDDate(DateTime.Now.ToString());
                        Rec.EventSeq = Dr["track_slno"].ToString();
                        Rec.EventCode = "";
                        Rec.FlightInformation = Generate_FlightInformation(Dr, POLCODE, POLNAME);
                        Rec.Party = Generate_Party();
                        mList[ArrIndex++] = Rec;

                        POLCODE = Dr["track_pod_code"].ToString();
                        POLNAME = Dr["track_pod_name"].ToString();
                    }
                    if (Dr["track_slno"].ToString() == "5")
                    {
                        Rec = new ShipmentTrackingMessageShipmentEventsShipmentEvent();
                        Rec.EventDateTime = ConvertYMDDate(DateTime.Now.ToString());
                        Rec.EventSeq = Dr["track_slno"].ToString();
                        Rec.EventCode = "";
                        Rec.FlightInformation = Generate_FlightInformation(Dr, POLCODE, POLNAME);
                        Rec.Party = Generate_Party();
                        mList[ArrIndex++] = Rec;

                        POLCODE = Dr["track_pod_code"].ToString();
                        POLNAME = Dr["track_pod_name"].ToString();
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private ShipmentTrackingMessageShipmentEventsShipmentEventFlightInformation[] Generate_FlightInformation(DataRow dr, string polcode, string polname)
        {
            ShipmentTrackingMessageShipmentEventsShipmentEventFlightInformation Rec = null;
            ShipmentTrackingMessageShipmentEventsShipmentEventFlightInformation[] mList = null;
            try
            {
                mList = new ShipmentTrackingMessageShipmentEventsShipmentEventFlightInformation[1];
                Rec = new ShipmentTrackingMessageShipmentEventsShipmentEventFlightInformation();
                Rec.FlightNo = dr["track_flight_no"].ToString();
                Rec.FlightCarrier = dr["track_carrier_name"].ToString();
                Rec.IATAFrom = polcode;
                Rec.IATAFromName = polname;
                Rec.ETD = ConvertYMDDate(dr["track_etd"].ToString());
                Rec.IATATo = dr["track_pod_code"].ToString();
                Rec.IATAToName = dr["track_pod_name"].ToString();
                Rec.ETA = ConvertYMDDate(dr["track_eta"].ToString());
                mList[0] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private ShipmentTrackingMessageShipmentEventsShipmentEventParty[] Generate_Party()
        {
            ShipmentTrackingMessageShipmentEventsShipmentEventParty Rec = null;
            ShipmentTrackingMessageShipmentEventsShipmentEventParty[] mList = null;
            try
            {
                mList = new ShipmentTrackingMessageShipmentEventsShipmentEventParty[1];
                Rec = new ShipmentTrackingMessageShipmentEventsShipmentEventParty();
                Rec.PartyAction = "";
                Rec.CompanyCode = "";
                Rec.CompanyName = "";
                mList[0] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private void WriteXmlFiles()
        {
            try
            {
                if (TrackMessage == null || IsError)
                    return;

                string FileName = "AWT";
                FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(ShipmentTrackingMessage));
                StreamWriter writer = new StreamWriter(FileName);
                mySerializer.Serialize(writer, TrackMessage, ns);
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
        private string GetSingleJobNo(string sJobNos)
        {
            if (sJobNos.ToString().Contains(","))
            {
                string[] sJob = sJobNos.Split(',');
                sJobNos = sJob[0];
            }
            return sJobNos;
        }
    }
}
