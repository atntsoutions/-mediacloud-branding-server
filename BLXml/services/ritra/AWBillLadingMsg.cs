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
using BLXml.models.AirBL;

namespace BLXml
{
    public partial class AWBillLadingMsg
    {
        private DataTable DT_SHIPMENT = new DataTable();
        public Boolean IsError = false;
        private AirwayBillMessage AwbMessage = null;
      //  Dictionary<int, string> ErrorDic = new Dictionary<int, string>();
        private string ErrorValues = "";
        private string sql = "";
        private string MessageNumber = "";
        private int AwbSeq = 0;
         
        DBConnection Con_Oracle = null;
        Dictionary<int, string> HblDesc = new Dictionary<int, string>();

        public void Generate()
        {
            try
            {
                IsError = false;
                ReadData();
                if(DT_SHIPMENT.Rows.Count<=0)
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
            
            sql = " select h.hbl_pkid,m.hbl_pkid as mbl_pkid,m.hbl_bl_no as mbl_no,h.hbl_bl_no as hbls_bl_no,";
            sql += "  h.hbl_year,h.hbl_job_nos as job_id,";
            sql += "  pol.param_code as POL_CODE,";
            sql += "  pol.param_name POL_NAME,";
            sql += "  pod.param_code as POD_CODE,";
            sql += "  pod.param_name POD_NAME,";
            sql += "  pofd.param_code as POFD_CODE,";
            sql += "  pofd.param_name as POFD_NAME,";
            sql += "  rcpt.param_code as RECEIPTPLACE_CODE,";
            sql += "  rcpt.param_name as RECEIPTPLACE_NAME,";
            sql += "  pofdcntry.param_code as POFD_COUNTRY_CODE,";
            sql += "  pofdcntry.param_name as POFD_COUNTRY_NAME,";
            sql += "  Agent.cust_Code as Agent_Code,";
            sql += "  Agent.cust_Name as Agent_Name,";
            sql += "  shipper.cust_code as shipper_code,";
            sql += "  shipper.cust_name as Shipper_Name,";
            sql += "  Consignee.cust_Code as Consignee_Code,";
            sql += "  Consignee.cust_Name as Consignee_Name,";
            sql += "  notify.cust_Code as notify_code,";
            sql += "  bl_notify_name as notify_name,";
            sql += "  h.hbl_terms as hblstatus_name,";
            sql += "  m.hbl_terms as mblstatus_name,";
            sql += "  pkgunit.param_code as hbls_pack_type_code,";
            sql += "  h.hbl_cbm as hbls_total_cbm,h.hbl_chwt as hbls_total_chargeable_weight,";
            sql += "  h.hbl_grwt as hbls_total_gross_weight,h.hbl_ntwt as hbls_total_net_weight,";
            sql += "  h.hbl_pkg as hbls_total_cartons,";
            sql += "  comm.param_code as commodity_code,comm.param_name as commodity_name ";
            sql += "  from hblm m";
            sql += "  inner join hblm h on  (m.hbl_pkid = h.hbl_mbl_id )";
            sql += "  left join bl b on ( h.hbl_pkid = b.bl_pkid)  ";
            sql += "  left join param pol on(h.hbl_pol_id=pol.param_pkid)";
            sql += "  left join param pod on (h.hbl_pod_id = pod.param_pkid)";
            sql += "  left join param pofd on (h.hbl_pofd_id = pofd.param_pkid)";
            sql += "  left join param	rcpt	on (h.hbl_place_receipt_id = rcpt.param_pkid)";
            sql += "  left join param pofdcntry on (h.hbl_pofd_country_id = pofdcntry.param_pkid)";
            sql += "  left join customerm  Agent on (h.hbl_Agent_id = Agent.cust_pkid)";
            sql += "  left join customerm  shipper on (h.hbl_exp_id = shipper.cust_pkid)";
            sql += "  left join customerm  Consignee on (h.hbl_imp_id = consignee.cust_pkid)";
            sql += "  left join customerm  notify on (b.bl_notify_id = notify.cust_pkid)";
            sql += "  left join param comm on (h.hbl_commodity_id = comm.param_pkid)";
            sql += "  left join param pkgunit on (h.hbl_pkg_unit_id = pkgunit.param_pkid)";
            sql += " where m.rec_company_code = '" + XmlLib.Company_Code + "' ";
            sql += " and m.rec_branch_code  = '" + XmlLib.Branch_Code + "' ";
            sql += " and m.rec_category  = 'AIR EXPORT' ";
            sql += " and m.hbl_agent_id = '" + XmlLib.Agent_Id + "' ";
            if (XmlLib.HBL_BL_NOS.Length > 0)
            {
                sql += " and h.hbl_bl_no in (" + XmlLib.HBL_BL_NOS + ")";
            }
            else
            {
                sql += " and ((sysdate - m.hbl_pol_etd) between 2 and  60) ";
            }
            sql += " order by h.hbl_no";

            DT_SHIPMENT = new DataTable();
            Con_Oracle = new DBConnection();
            DT_SHIPMENT = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
        }

        private void GenerateXmlFiles()
        {
            AwbMessage = new AirwayBillMessage();
            AwbMessage.Items = Generate_AirWayBillMsg();
        }

        private object[] Generate_AirWayBillMsg()
        {
            object[] Items = null;
            int iTotRows = 0;
            int ArrIndex = 0;
            try
            {

                AwbSeq = 0;
                DataTable DistinctHBL = DT_SHIPMENT.DefaultView.ToTable(true, "hbl_pkid", "hbls_bl_no");
                iTotRows = DistinctHBL.Rows.Count;

                Items = new object[iTotRows + 1];
                Items[ArrIndex++] = Generate_MessageInfo();
                foreach (DataRow dr in DistinctHBL.Select("1=1", "hbls_bl_no"))
                {
                    AwbSeq++;
                    Items[ArrIndex++] = Generate_AirWayBill(dr["hbl_pkid"].ToString());
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
        private AirwayBillMessageMessageInfo Generate_MessageInfo()
        {
            AirwayBillMessageMessageInfo Rec = null;
            try
            {
                this.MessageNumber = XmlLib.GetNewMessageNumber();

                Rec = new AirwayBillMessageMessageInfo();
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

        private AirwayBillMessageAirwayBill Generate_AirWayBill(string HBL_ID)
        {
            AirwayBillMessageAirwayBill Rec = null;
            string PreData = "1";
            string EtaDischarge = "";
            DataTable Dt_Track;
            try
            { 

                sql = "  select trk.trk_voyage as track_flight_no,trk_pol_etd as track_etd,trk_pod_eta as track_eta";
                sql += "  ,pol.param_code as  track_pol_code,pol.param_name as  track_pol_name";
                sql += "  ,pod.param_code as  track_pod_code,pod.param_name as  track_pod_name";
                sql += "  ,lnr.param_code  as track_carrier_code,lnr.param_name as track_carrier_name ";
                sql += "  from hblm h";
                sql += "  inner join trackingm trk on h.hbl_mbl_id = trk.trk_parent_id";
                sql += "  left join param pol on trk.trk_pol_id = pol.param_pkid";
                sql += "  left join param pod on trk.trk_pod_id = pod.param_pkid";
                sql += "  left join param lnr on h.hbl_carrier_id = lnr.param_pkid";
                sql += "  where  h.hbl_pkid = '" + HBL_ID + "'";
                sql += "  order by  trk_order";


                Dt_Track = new DataTable();
                Con_Oracle = new DBConnection();
                Dt_Track = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in DT_SHIPMENT.Select("hbl_pkid ='" + HBL_ID + "'", "hbl_pkid"))
                {
                    if (PreData != Dr["hbl_pkid"].ToString())
                    {
                        PreData = Dr["hbl_pkid"].ToString();
                        EtaDischarge = "";
                        Rec = new AirwayBillMessageAirwayBill();
                        Rec.Movement = "";
                        Rec.BookingNo = Dr["HBL_YEAR"].ToString() + GetSingleJobNo(Dr["JOB_ID"].ToString());
                        Rec.AWBSeq = AwbSeq.ToString();
                        Rec.Action = "Replace";
                        Rec.MemberCode = XmlLib.messageSenderField;
                        Rec.HAWBnumber = Dr["hbls_bl_no"].ToString();
                        Rec.MAWBnumber = Dr["mbl_no"].ToString();
                        Rec.AgentIATACode = "";
                        Rec.AirportOfLoading = Dr["pol_code"].ToString();
                        Rec.AirportOfLoadingName = Dr["pol_name"].ToString();
                        Rec.FlightNoLoading = "";
                        Rec.FlightCarrier = "";
                        Rec.ETDLoading = "";
                        Rec.AirportVia1 = "";
                        Rec.AirportVia1Name = "";
                        Rec.FlightNoVia1 = "";
                        Rec.Via1Carrier = "";
                        Rec.ETDVia1 = "";
                        Rec.AirportVia2 = "";
                        Rec.AirportVia2Name = "";
                        Rec.FlightNoVia2 = "";
                        Rec.Via2Carrier = "";
                        Rec.ETDVia2 = "";
                        Rec.AirportVia3 = "";
                        Rec.AirportVia3Name = "";
                        Rec.FlightNoVia3 = "";
                        Rec.Via3Carrier = "";
                        Rec.ETDVia3 = "";
                        if (Dt_Track.Rows.Count > 0)
                        {
                            Rec.FlightNoLoading = Dt_Track.Rows[0]["track_flight_no"].ToString();
                            Rec.FlightCarrier = Dt_Track.Rows[0]["track_carrier_name"].ToString();
                            Rec.ETDLoading = ConvertYMDDate(Dt_Track.Rows[0]["track_etd"].ToString());
                            Rec.AirportVia1 = Dt_Track.Rows[0]["track_pod_code"].ToString();
                            Rec.AirportVia1Name = Dt_Track.Rows[0]["track_pod_name"].ToString();

                            EtaDischarge = ConvertYMDDate(Dt_Track.Rows[0]["track_eta"].ToString());
                        }
                        if (Dt_Track.Rows.Count > 1)
                        {
                            Rec.FlightNoVia1 = Dt_Track.Rows[1]["track_flight_no"].ToString();
                            Rec.Via1Carrier = Dt_Track.Rows[1]["track_carrier_name"].ToString();
                            Rec.ETDVia1 = ConvertYMDDate(Dt_Track.Rows[1]["track_etd"].ToString());
                            Rec.AirportVia2 = Dt_Track.Rows[1]["track_pod_code"].ToString();
                            Rec.AirportVia2Name = Dt_Track.Rows[1]["track_pod_name"].ToString();

                            EtaDischarge = ConvertYMDDate(Dt_Track.Rows[1]["track_eta"].ToString());
                        }
                        if (Dt_Track.Rows.Count > 2)
                        {
                            Rec.FlightNoVia2 = Dt_Track.Rows[2]["track_flight_no"].ToString();
                            Rec.Via2Carrier = Dt_Track.Rows[2]["track_carrier_name"].ToString();
                            Rec.ETDVia2 = ConvertYMDDate(Dt_Track.Rows[2]["track_etd"].ToString());
                            Rec.AirportVia3 = Dt_Track.Rows[2]["track_pod_code"].ToString();
                            Rec.AirportVia3Name = Dt_Track.Rows[2]["track_pod_name"].ToString();

                            EtaDischarge = ConvertYMDDate(Dt_Track.Rows[2]["track_eta"].ToString());
                        }
                        if (Dt_Track.Rows.Count > 3)
                        {
                            Rec.FlightNoVia3 = Dt_Track.Rows[3]["track_flight_no"].ToString();
                            Rec.Via3Carrier = Dt_Track.Rows[3]["track_carrier_name"].ToString();
                            Rec.ETDVia3 = ConvertYMDDate(Dt_Track.Rows[3]["track_etd"].ToString());

                            EtaDischarge = ConvertYMDDate(Dt_Track.Rows[3]["track_eta"].ToString());
                        }

                        Rec.AirportOfDischarge = Dr["pod_code"].ToString();
                        Rec.AirportOfDischargeName = Dr["pod_name"].ToString();
                        Rec.ETADischarge = EtaDischarge;
                        Rec.PlaceOfReceipt = Dr["Receiptplace_name"].ToString();
                        Rec.PlaceOfDelivery = Dr["pofd_code"].ToString();
                        Rec.PlaceOfDeliveryCountry = Dr["pofd_country_code"].ToString();

                        if (Dr["mblstatus_name"].ToString() == "FREIGHT PREPAID")
                            Rec.PaymentTerm_Master = "P";
                        else if (Dr["mblstatus_name"].ToString() == "FREIGHT COLLECT")
                            Rec.PaymentTerm_Master = "C";
                        else
                            Rec.PaymentTerm_Master = "";

                        if (Dr["hblstatus_name"].ToString() == "FREIGHT PREPAID")
                            Rec.PaymentTerm_House = "P";
                        else if (Dr["hblstatus_name"].ToString() == "FREIGHT COLLECT")
                            Rec.PaymentTerm_House = "C";
                        else
                            Rec.PaymentTerm_House = "";
                        Rec.IncoTerms = "";
                        Rec.BookingReferences = "";
                        Rec.Parties = Generate_AirWayBillParties(HBL_ID);
                        Rec.Commodities = Generate_AirWayBillCommodities(HBL_ID);

                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return Rec;
        }

        private AirwayBillMessageAirwayBillPartiesParty[] Generate_AirWayBillParties(string HBL_ID)
        {
            AirwayBillMessageAirwayBillPartiesParty Rec = null;
            AirwayBillMessageAirwayBillPartiesParty[] mPartyList = null;
            int ArrIndex = 0;
            string PreData = "1";
            try
            {
                mPartyList = new AirwayBillMessageAirwayBillPartiesParty[4];
                foreach (DataRow Dr in DT_SHIPMENT.Select("hbl_pkid ='" + HBL_ID + "'", "hbl_pkid"))
                {
                    if (PreData != Dr["hbl_pkid"].ToString())
                    {
                        PreData = Dr["hbl_pkid"].ToString();

                        Rec = new AirwayBillMessageAirwayBillPartiesParty();
                        Rec.PartyType = "SHIPPER";
                        Rec.CompanyCode = Dr["shipper_code"].ToString();
                        Rec.CompanyName = Dr["shipper_Name"].ToString();
                        mPartyList[ArrIndex++] = Rec;

                        Rec = new AirwayBillMessageAirwayBillPartiesParty();
                        Rec.PartyType = "CONSIGNEE";
                        Rec.CompanyCode = Dr["Consignee_Code"].ToString();
                        Rec.CompanyName = Dr["Consignee_Name"].ToString();
                        mPartyList[ArrIndex++] = Rec;

                        Rec = new AirwayBillMessageAirwayBillPartiesParty();
                        Rec.PartyType = "NOTIFY";
                        Rec.CompanyCode = Dr["notify_code"].ToString();
                        if (Rec.CompanyCode == "")
                            Rec.CompanyCode = Dr["notify_name"].ToString();
                        Rec.CompanyName = Dr["notify_name"].ToString();
                        mPartyList[ArrIndex++] = Rec;

                        Rec = new AirwayBillMessageAirwayBillPartiesParty();

                        Rec.PartyType = "AGENT";
                        Rec.CompanyCode = Dr["Agent_Code"].ToString();
                        Rec.CompanyName = Dr["Agent_Name"].ToString();
                        mPartyList[ArrIndex++] = Rec;

                    }
                }

            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mPartyList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodity[] Generate_AirWayBillCommodities(string HBL_ID)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodity Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodity[] mList = null;
            int ArrIndex = 0;
            string carton_unit = "";
            string total_cartons = "";
            string total_cbm = "";
            string chargeable_weight = "";
            string gross_weight = "";
            string commodity = "";
            try
            {
                foreach (DataRow Dr in DT_SHIPMENT.Select("hbl_pkid ='" + HBL_ID + "'", "hbl_pkid"))
                {
                    carton_unit = Dr["hbls_pack_type_code"].ToString();
                    total_cartons = Dr["hbls_total_cartons"].ToString();
                    total_cbm = Dr["hbls_total_cbm"].ToString();
                    chargeable_weight = Dr["hbls_total_chargeable_weight"].ToString();
                    gross_weight = Dr["hbls_total_gross_weight"].ToString();
                    commodity = Dr["commodity_code"].ToString();
                    break;
                }
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodity[1];
                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodity();
                Rec.ItemNumber = "";
                Rec.PurchaseOrder = "";
                Rec.CommodityCode = commodity;
                Rec.ItemSeq = (ArrIndex + 1).ToString();
                Rec.MarksNumbers = Generate_MarksNumbers(HBL_ID);
                Rec.CargoDescriptions = Generate_Descriptions(HBL_ID);
                Rec.CommColli = Generate_Cartons(carton_unit, total_cartons);
                Rec.CommMeasurementActual = Generate_ActualCbm("CBM", total_cbm);
                Rec.CommMeasurementCargeable = Generate_ChargeableCbm("CBM", "");
                Rec.CommWeightActual = Generate_ActualWeight("KGS", gross_weight);
                Rec.CommWeightCargeable = Generate_ChargeableWeight("KGS", chargeable_weight);
                Rec.CommHandling = Generate_CommHandling(HBL_ID);
                mList[ArrIndex++] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }

        private AirwayBillMessageAirwayBillCommoditiesCommodityMarksNumbersMarksNumber[] Generate_MarksNumbers(string HBL_ID)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityMarksNumbersMarksNumber Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityMarksNumbersMarksNumber[] mList = null;
            string styleno = "";
            string Orderno = "";
            string SbNo = "";
            string SbDate = "";
            int ArrIndex = 0;
            int ArrLen = 3;
            bool differentSbDate = false;
            string OrdDesc = "";
            string StrDesc = "";
            HblDesc = new Dictionary<int, string>();
            try
            {
                Con_Oracle = new DBConnection();
                sql = " select distinct ord_po, ord_style,ord_desc from joborderm a  ";
                sql += " inner join jobm b on a.ord_parent_id = b.job_pkid";
                sql += " where  b.jobs_hbl_id ='" + HBL_ID + "'";
                sql += " order by ord_po ";
                DataTable Dt_Temp = new DataTable();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Temp.Rows.Count <= 0)
                {
                    sql = " select distinct itm_orderno as ord_po,itm_styleno as ord_style,itm_desc as ord_desc from itemm a";
                    sql += " inner join jobm b on a.itm_job_id = b.job_pkid ";
                    sql += " where b.jobs_hbl_id ='" + HBL_ID + "'";
                    sql += " order by itm_orderno ";
                    Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                }
                foreach (DataRow dr in Dt_Temp.Rows)
                {
                    if (Orderno.Trim() != "")
                        Orderno += ",";
                    Orderno += dr["ord_po"].ToString();

                    if (styleno.Trim() != "")
                        styleno += ",";
                    styleno += dr["ord_style"].ToString();

                    StrDesc = dr["ord_desc"].ToString().Replace("\n", "");
                    if (!HblDesc.ContainsValue(StrDesc)) // to get distinct description use dictonary
                        HblDesc.Add(HblDesc.Count, StrDesc);
                }

                sql = " select opr_Sbill_Date,opr_sbill_no from joboperationsm a";
                sql += "  inner join jobm b on a.opr_job_id = b.job_pkid ";
                sql += "  where  b.jobs_hbl_id ='" + HBL_ID + "'";
                sql += " order by opr_sbill_no";
                Dt_Temp = new DataTable();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);

                Con_Oracle.CloseConnection();

                if (Dt_Temp.Rows.Count > 0)
                {
                    DataTable DistinctSB = Dt_Temp.DefaultView.ToTable(true, "opr_Sbill_Date");
                    if (DistinctSB.Rows.Count > 1)
                        differentSbDate = true;

                    foreach (DataRow dr in Dt_Temp.Rows)
                    {
                        if (!dr["opr_Sbill_Date"].Equals(DBNull.Value))
                            SbDate = ((DateTime)dr["opr_Sbill_Date"]).ToString("dd.MM.yyyy");

                        if (SbNo.Trim() != "")
                            SbNo += ",";
                        SbNo += dr["opr_sbill_no"].ToString();
                        if (differentSbDate)
                            SbNo += " DT:" + SbDate;
                    }
                }
                if (!differentSbDate && SbNo.Trim() != "" && SbDate.Trim() != "")
                    SbNo += " DT:" + SbDate;


                if (styleno.Trim() != "")
                    styleno = "STYLE#:" + styleno;
                if (Orderno.Trim() != "")
                    Orderno = "ORDER#:" + Orderno;
                if (SbNo.Trim() != "")
                    SbNo = "SB.NO:" + SbNo;

                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityMarksNumbersMarksNumber[ArrLen];

                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityMarksNumbersMarksNumber();
                Rec.Value = styleno;
                mList[ArrIndex++] = Rec;

                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityMarksNumbersMarksNumber();
                Rec.Value = Orderno;
                mList[ArrIndex++] = Rec;

                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityMarksNumbersMarksNumber();
                Rec.Value = SbNo;
                mList[ArrIndex++] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodityCargoDescriptionsCargoDescription[] Generate_Descriptions(string HBL_ID)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityCargoDescriptionsCargoDescription Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityCargoDescriptionsCargoDescription[] mList = null;
            try
            {
                 
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityCargoDescriptionsCargoDescription[HblDesc.Count];
                for (int i = 0; i < HblDesc.Count; i++)
                {
                    Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityCargoDescriptionsCargoDescription();
                    Rec.Value = HblDesc[i];
                    mList[i] = Rec;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodityCommColli[] Generate_Cartons(string sUnit, string sValue)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityCommColli Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityCommColli[] mList = null;
            try
            {
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityCommColli[1];
                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityCommColli();
                Rec.PackageType = sUnit;
                Rec.Value = sValue;
                mList[0] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementActual[] Generate_ActualCbm(string sUnit, string sValue)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementActual Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementActual[] mList = null;
            try
            {
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementActual[1];
                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementActual();
                Rec.UOM = sUnit;
                Rec.Value = sValue;
                mList[0] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementCargeable[] Generate_ChargeableCbm(string sUnit, string sValue)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementCargeable Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementCargeable[] mList = null;
            try
            {
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementCargeable[1];
                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityCommMeasurementCargeable();
                Rec.UOM = sUnit;
                Rec.Value = sValue;
                mList[0] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightActual[] Generate_ActualWeight(string sUnit, string sValue)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightActual Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightActual[] mList = null;
            try
            {
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightActual[1];
                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightActual();
                Rec.UOM = sUnit;
                Rec.Value = sValue;
                mList[0] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightCargeable[] Generate_ChargeableWeight(string sUnit, string sValue)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightCargeable Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightCargeable[] mList = null;
            try
            {
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightCargeable[1];
                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityCommWeightCargeable();
                Rec.UOM = sUnit;
                Rec.Value = sValue;
                mList[0] = Rec;
            }
            catch (Exception Ex)
            {
                IsError = true;
                throw Ex;
            }
            return mList;
        }
        private AirwayBillMessageAirwayBillCommoditiesCommodityCommHandlingHandlingInformation[] Generate_CommHandling(string HBL_ID)
        {
            AirwayBillMessageAirwayBillCommoditiesCommodityCommHandlingHandlingInformation Rec = null;
            AirwayBillMessageAirwayBillCommoditiesCommodityCommHandlingHandlingInformation[] mList = null;
            try
            {
                mList = new AirwayBillMessageAirwayBillCommoditiesCommodityCommHandlingHandlingInformation[1];
                Rec = new AirwayBillMessageAirwayBillCommoditiesCommodityCommHandlingHandlingInformation();
                Rec.Value = "";
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
                if (AwbMessage == null || IsError)
                    return;

                string FileName = "AWB";
                FileName = XmlLib.sentFolder + FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(AirwayBillMessage));
                StreamWriter writer = new StreamWriter(FileName);
                mySerializer.Serialize(writer, AwbMessage, ns);
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
