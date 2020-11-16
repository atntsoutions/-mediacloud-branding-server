using System;
using System.Data;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.MexicoBL;
using System.Text;

namespace BLXml 
{
    public class MexicoBLRpt : BaseReport
    {
        private DataTable DT_BLINFO = new DataTable();
        private DataTable DT_BLORD = new DataTable();
        public Boolean IsError = false;
        public string ErrorMessage = "";
        private MessageBLInfo BLMessage = null;
        private string sql = "";
        public string PKID = "";
        public string File_Name = "";
        public string File_Subject = "";
        DBConnection Con_Oracle = null;
        public string RowType = "";
        StringBuilder sb = new StringBuilder();
        public void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;

                ErrorMessage = AllValid();
                if (ErrorMessage.Length > 0)
                {
                    IsError = true;
                    return;
                }

                ReadData();
                if (DT_BLINFO.Rows.Count <= 0||DT_BLORD.Rows.Count<=0)
                {
                    IsError = true;
                    ErrorMessage = "Details not Found";
                    return;
                }
                if (RowType == "CHECK-LIST")
                {
                    GenerateCsvFiles();
                    WriteCsvFiles();
                }
                else
                {
                    GenerateXmlFiles();
                    WriteXmlFiles();
                }
                DT_BLINFO.Rows.Clear();
            }
            catch (Exception ex)
            {
                IsError = true;
                ErrorMessage += " |" + ex.Message.ToString();
            }
        }

        private string AllValid()
        {
            string sError = "";
            string str = "";
            DataTable Dt_Temp;
            Con_Oracle = new DBConnection();
            try
            {

                sql = " select distinct carr.param_name as source_carrier ,lkcarr.targetid as target_carrier";
                sql += "   from hblm mbl";
                sql += "   left join param carr on mbl.hbl_carrier_id = carr.param_pkid";
                sql += "   left join linkm2 lkcarr on mbl.hbl_carrier_id = lkcarr.sourceid and lkcarr.sourcetable='MEXICO-TMM' and lkcarr.sourcetype='SEA CARRIER'";
                sql += "   where mbl.hbl_pkid='" + PKID + "' and lkcarr.targetid is null";
                Dt_Temp = new DataTable();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                str = "";
                foreach (DataRow dr in Dt_Temp.Rows)
                {
                    if (str != "")
                        str += ", ";
                    str += dr["source_carrier"].ToString();
                }
                if (str != "")
                    Lib.AddError(ref sError, " Carrier "+ str +" not linked ");


                sql = " select distinct cntrtype.param_code as source_cntr_type, lkcntrtype.targetid as target_cntr_type";
                sql += "   from hblm hbl";
                sql += "   inner join jobm job on hbl.hbl_pkid = job.jobs_hbl_id";
                sql += "   left join packingm pkg on job.job_pkid = pkg.pack_job_id";
                sql += "   left join containerm cntr on pkg.pack_cntr_id = cntr.cntr_pkid";
                sql += "   left join param cntrtype on cntr.cntr_type_id = cntrtype.param_pkid";
                sql += "   left join linkm2 lkcntrtype on cntr.cntr_type_id = lkcntrtype.sourceid and lkcntrtype.sourcetable='MEXICO-TMM' and lkcntrtype.sourcetype='CONTAINER'";
                sql += "   where hbl.hbl_mbl_id='" + PKID + "' and lkcntrtype.targetid is null";
                sql += "   order by cntrtype.param_code";
                Dt_Temp = new DataTable();
                Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                str = "";
                foreach (DataRow dr in Dt_Temp.Rows)
                {
                    if (str != "")
                        str += ", ";
                    str += dr["source_cntr_type"].ToString();
                }
                if (str != "")
                    Lib.AddError(ref sError, " | Container Types "+ str+ " not linked " );

            }
            catch (Exception Ex)
            {
                Lib.AddError(ref sError, Ex.Message.ToString());
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
            }
            Con_Oracle.CloseConnection();
            return sError;
        }
            private void ReadData()
        {
            string str = "";
            string mblbk = "";
            Con_Oracle = new DBConnection();
            /*
            sql = " select distinct mbl.hbl_no as mblbk,carr.param_code as carrier";
            sql += "  ,vsl.param_name as vslname";
            sql += "  ,mbl.hbl_vessel_no as voyage";
            sql += "  ,orgcntry.param_code as oricountry";
            sql += "  ,pol.param_code as pol";
            sql += "  ,pod.param_code as pod";
            sql += "  ,pofd.param_code as finaldestination";
            sql += "  ,'' as tsport";
            sql += "  ,'' as tsportdate";
            sql += "  ,mbl.hbl_pol_cutoff as cutoff";
            sql += "  ,mbl.hbl_pol_etd as etd";
            sql += "  ,mbl.hbl_pod_eta as eta";
            sql += "  ,'' as Traffic";
            sql += "  ,ord_uid as poid";
            sql += "  ,ord_po as po";
            sql += "  ,ord_style as styleno";
            sql += "  ,ord_po as posupplier";
            sql += "  ,'' as oldpo";
            sql += "  ,hbl.hbl_bl_no as hblnumber";
            sql += "  ,'' as documenttype";
            sql += "  ,shpr.cust_name as shipper";
            sql += "  ,shprcntry.param_code as shippercountry";
            sql += "  ,shpraddr.add_line1 as shpr_add1,shpraddr.add_line2 as shpr_add2,shpraddr.add_line3 as shpr_add3,shpraddr.add_line4 as shpr_add4";
            sql += "  ,shpraddr.add_tel as shipperphonenumber";
            sql += "  ,nfy.cust_name as  notifyparty";
            sql += "  ,mbl.hbl_bl_no as mblnumber";
            sql += "  ,'' as servicecontract";
            sql += "  ,cntr.cntr_no as containernumber";
            sql += "  ,cntrtype.param_code as containertype";
            sql += "  ,cntr.cntr_csealno as carrierseal";
            sql += "  ,cntr.cntr_asealno as shipperseal";
            sql += "  ,mbl.hbl_shipment_type  as loadtype";
            sql += "  ,cntr_trafinsp as IntertraficInspection";
            sql += "  ,'' as InspectionIn";
            sql += "  ,cntr_inspsup as InspectionSupplier";
            sql += "   from hblm mbl";
            sql += "  inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += "  inner join jobm job on hbl.hbl_pkid = job.jobs_hbl_id";
            sql += "  inner join joborderm ord on job.job_pkid = ord.ord_parent_id";
            sql += "  left join param carr on mbl.hbl_carrier_id = carr.param_pkid";
            sql += "  left join param vsl on mbl.hbl_vessel_id = vsl.param_pkid";
            sql += "  left join param orgcntry on job.job_origin_country_id = orgcntry.param_pkid";
            sql += "  left join param pol on job.job_pol_id = pol.param_pkid";
            sql += "  left join param pod on job.job_pod_id = pod.param_pkid";
            sql += "  left join param pofd on job.job_pofd_id = pofd.param_pkid ";
            sql += "  left join customerm shpr on job.job_exp_id = shpr.cust_pkid";
            sql += "  left join addressm shpraddr on job.job_exp_br_id = shpraddr.add_pkid";
            sql += "  left join param shprcntry on shpraddr.add_country_id = shprcntry.param_pkid";
            sql += "  left join bl on hbl.hbl_pkid = bl.bl_pkid";
            sql += "  left join customerm nfy on bl.bl_notify_id = nfy.cust_pkid";
            sql += "  left join packingm pkg on job.job_pkid = pkg.pack_job_id";
            sql += "  inner join packingcntrorder pkgcntr on pkg.pack_pkid = pkgcntr.pcntr_pack_id";
            sql += "  left join containerm cntr on pkg.pack_cntr_id = cntr.cntr_pkid";
            sql += "  left join param cntrtype on cntr.cntr_type_id = cntrtype.param_pkid";
            sql += "  where mbl.hbl_pkid = '" + PKID + "'";
            sql += "  order by shpr.cust_name,ord_po";
            */

            sql = "  select mbl.hbl_no as mblbk,mbl.hbl_bl_no as mblnumber";
            sql += "   ,case when mbl.hbl_shipment_type='FCL' or mbl.hbl_shipment_type='LCL' then mbl.hbl_shipment_type else cast('FCLCONSOL' as nvarchar2(15)) end  as loadtype";
            sql += "   ,carr.targetid as carrier";
            sql += "   ,vsl.param_name as vslname";
            sql += "   ,mbl.hbl_vessel_no as voyage";
            sql += "   ,case when nvl(length(pol.param_code),0)>=5 then substr(pol.param_code,0,2) else null end as oricountry";
            sql += "   ,pol.param_code as pol";
            sql += "   ,pod.param_code as pod";
            sql += "   ,pofd.param_code as finaldestination";
            sql += "   ,'' as tsport";
            sql += "   ,'' as tsportdate";
            sql += "   ,mbl.hbl_pol_cutoff as cutoff";
            sql += "   ,mbl.hbl_pol_etd as etd";
            sql += "   ,mbl.hbl_pod_eta as eta";
            sql += "   ,'' as Traffic";
            sql += "   from hblm mbl";
            sql += "   left join linkm2 carr on mbl.hbl_carrier_id = carr.sourceid and carr.sourcetable='MEXICO-TMM' and carr.sourcetype='SEA CARRIER'";
            sql += "   left join param vsl on mbl.hbl_vessel_id = vsl.param_pkid";
            sql += "   left join param pol on mbl.hbl_pol_id = pol.param_pkid";
            sql += "   left join param pod on mbl.hbl_pod_id = pod.param_pkid";
            sql += "   left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid ";
            sql += "  where mbl.hbl_pkid = '" + PKID + "'";

            DT_BLINFO = new DataTable();
            DT_BLINFO = Con_Oracle.ExecuteQuery(sql);

            DT_BLORD = new DataTable();
            if (DT_BLINFO.Rows.Count > 0)
            {
                mblbk = DT_BLINFO.Rows[0]["mblbk"].ToString();

                sql = " select ";
                sql += "    ord_uid as poid";
                sql += "   ,ord_po as po";
                sql += "   ,ord_style as styleno";
                sql += "   ,ord_po as posupplier";
                sql += "   ,'' as oldpo";
                sql += "   ,nvl(hbl.hbl_bl_no,hbl.hbl_fcr_no) as hblnumber";
                sql += "   ,case when hbl.hbl_bl_no is null then 'FCR' else 'ORIGINAL' end  as documenttype";
                sql += "   ,nvl(ord.ord_exp_name,shpr.cust_name) as shipper";
                sql += "   ,'' as shippercountry";
                sql += "  ,'' as shpr_add1,'' as shpr_add2,'' as shpr_add3,'' as shpr_add4";
                sql += "  ,'' as shipperphonenumber";
                sql += "   ,nfy.cust_name as  notifyparty";
                sql += "    ,'" + DT_BLINFO.Rows[0]["mblnumber"].ToString() + "' as mblnumber";
                sql += "   ,sc.param_name as servicecontract";
                sql += "   ,cntr.cntr_no as containernumber";
                sql += "   ,cntrtype.targetid as containertype";
                sql += "   ,cntr.cntr_csealno as shipperseal";//Custom seal as shipper seal
                sql += "   ,cntr.cntr_asealno as carrierseal";
                sql += "   ,'" + DT_BLINFO.Rows[0]["loadtype"].ToString() + "'  as loadtype";
                sql += "   ,case when nvl(cntr_trafinsp,'N')='N' then 'NOT' else 'YES' end as IntertraficInspection";
                sql += "   ,cntr_inspin as InspectionIn";
                sql += "   ,cntr_inspsup as InspectionSupplier";
                sql += "   from hblm hbl";
                sql += "   inner join jobm job on hbl.hbl_pkid = job.jobs_hbl_id";
                sql += "   inner join joborderm ord on job.job_pkid = ord.ord_parent_id";
                sql += "   left join customerm shpr on ord.ord_exp_id = shpr.cust_pkid";
                sql += "   left join bl on hbl.hbl_pkid = bl.bl_pkid";
                sql += "   left join customerm nfy on bl.bl_notify_id = nfy.cust_pkid";
                sql += "   left join packingm pkg on job.job_pkid = pkg.pack_job_id";
                sql += "   inner join packingcntrorder pkgcntr on ord.ord_pkid = pkgcntr.pcntr_ord_id";
                sql += "   left join containerm cntr on pkg.pack_cntr_id = cntr.cntr_pkid";
                sql += "   left join linkm2 cntrtype on cntr.cntr_type_id = cntrtype.sourceid and cntrtype.sourcetable='MEXICO-TMM' and cntrtype.sourcetype='CONTAINER'";
                sql += "   left join param  sc on cntr.cntr_service_contract_id = sc.param_pkid ";
                sql += "   where hbl.hbl_mbl_id = '" + PKID + "'";
                sql += "   order by shpr.cust_name,ord_po";
                DT_BLORD = Con_Oracle.ExecuteQuery(sql);

                str = "";
                foreach (DataRow dr in DT_BLORD.Rows)
                {
                    dr["containernumber"] = dr["containernumber"].ToString().Replace("-", "").Replace(" ", "");
                    dr["InspectionIn"] = dr["InspectionIn"].ToString().Replace(" ", "").Replace("WHAREHOUSE", "'S WHAREHOUSE");

                    if (str != "")
                        str += ",";
                    str += dr["PO"].ToString() + "/" + dr["styleno"].ToString();

                    if (Lib.Conv2Decimal(dr["poid"].ToString()) <= 0)
                    {
                        IsError = true;
                        ErrorMessage += " |  PO ID Cannot be Blank for PO " + dr["PO"].ToString();
                    }

                    if (dr["po"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | PO Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }
                    if (dr["styleno"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | Style No Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }
                    if (dr["hblnumber"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | HBL Number Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }
                    if (dr["shipper"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | Shipper Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }
                    if (dr["notifyparty"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | Notify Party Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }

                    if (dr["servicecontract"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | Service Contract Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }

                    if (dr["containertype"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | Container Type Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }
                    if (dr["carrierseal"].ToString() == "")
                    {
                        IsError = true;
                        ErrorMessage += " | Carrier Seal Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                    }

                    if (dr["IntertraficInspection"].ToString()=="YES")
                    {
                        if(dr["InspectionIn"].ToString()=="")
                        {
                            IsError = true;
                            ErrorMessage += " | Inspection In Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                        }

                        if (dr["InspectionSupplier"].ToString() == "")
                        {
                            IsError = true;
                            ErrorMessage += " | Inspection Supplier Cannot be Blank for Cntr " + dr["containernumber"].ToString();
                        }
                    }
                }

                if (DT_BLINFO.Rows[0]["mblnumber"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | MBL Number Cannot be Blank for mblbk " + mblbk.ToString();
                }

                if (DT_BLINFO.Rows[0]["loadtype"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | Shipment Type Cannot be Blank for mblbk " + mblbk.ToString();
                }

                if (DT_BLINFO.Rows[0]["carrier"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | Carrier Cannot be Blank for mblbk " + mblbk.ToString();
                }

                if (DT_BLINFO.Rows[0]["vslname"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | Vessel Name Cannot be Blank for mblbk " + mblbk.ToString();
                }
                if (DT_BLINFO.Rows[0]["voyage"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | Voyage Cannot be Blank for mblbk " + mblbk.ToString();
                }
                if (DT_BLINFO.Rows[0]["pol"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | POL Cannot be Blank for mblbk " + mblbk.ToString();
                }
                if (DT_BLINFO.Rows[0]["pod"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | POD Cannot be Blank for mblbk " + mblbk.ToString();
                }
                if (DT_BLINFO.Rows[0]["oricountry"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | Origin Country Cannot be Blank for mblbk " + mblbk.ToString();
                }
                if (DT_BLINFO.Rows[0]["cutoff"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | Cuttoff Cannot be Blank for mblbk " + mblbk.ToString();
                }
                if (DT_BLINFO.Rows[0]["etd"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | ETD Cannot be Blank for mblbk " + mblbk.ToString();
                }
                if (DT_BLINFO.Rows[0]["eta"].ToString() == "")
                {
                    IsError = true;
                    ErrorMessage += " | ETA Cannot be Blank for mblbk " + mblbk.ToString();
                }

                DT_BLORD.AcceptChanges();
            }
            Con_Oracle.CloseConnection();
            File_Subject = "MBLBK-" + mblbk + ", PO/STYLE-" + str;
        }

        private void GenerateXmlFiles()
        {

            BLMessage = new MessageBLInfo();
            BLMessage.ProcessID = XmlLib.PROCESSID;
            BLMessage.VslVoy = Generate_BLinfoVslVoy();
            BLMessage.Orders = Generate_BLinfoOrders();
        }

        private MessageBLInfoVslVoy[] Generate_BLinfoVslVoy()
        {
            MessageBLInfoVslVoy Rec = null;
            MessageBLInfoVslVoy[] mVslList = null;
            int ArrIndex = 0;
            try
            {
                mVslList = new MessageBLInfoVslVoy[1];
                foreach (DataRow Dr in DT_BLINFO.Rows)
                {
                    Rec = new MessageBLInfoVslVoy();
                    Rec.Carrier = Lib.GetTruncated(Dr["Carrier"].ToString(), 30);
                    Rec.VslName = Lib.GetTruncated(Dr["VslName"].ToString(), 30);
                    Rec.Voygae = Lib.GetTruncated(Dr["voyage"].ToString(), 30);
                    Rec.OriCountry = Dr["OriCountry"].ToString();
                    Rec.POL = Lib.GetPortCode(Dr["POL"].ToString());
                    Rec.POD = Lib.GetPortCode(Dr["POD"].ToString());
                    Rec.FinalDestination = Lib.GetPortCode(Dr["FinalDestination"].ToString());
                    Rec.TSPort = Dr["TSPort"].ToString();
                    Rec.TSPortDate = Lib.DatetoStringDisplayformat(Dr["TSPortDate"]);
                    Rec.CutOff = Lib.DatetoStringDisplayformat(Dr["CutOff"]);
                    Rec.ETD = Lib.DatetoStringDisplayformat(Dr["ETD"]);
                    Rec.ETA = Lib.DatetoStringDisplayformat(Dr["ETA"]);
                    Rec.Traffic = Dr["Traffic"].ToString();
                    mVslList[ArrIndex++] = Rec;
                    break;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mVslList;
        }

        private MessageBLInfoOrdersOrder[] Generate_BLinfoOrders()
        {
            MessageBLInfoOrdersOrder Rec = null;
            MessageBLInfoOrdersOrder[] mOrdList = null;
            int ArrIndex = 0;

            string str = "";
            try
            {
                mOrdList = new MessageBLInfoOrdersOrder[DT_BLORD.Rows.Count];
                foreach (DataRow Dr in DT_BLORD.Rows)
                {
                    Rec = new MessageBLInfoOrdersOrder();
                    Rec.POID = Dr["poid"].ToString();
                    Rec.PO = Lib.GetTruncated(Dr["po"].ToString(),20);
                    Rec.StyleNo = Lib.GetTruncated(Dr["styleno"].ToString(), 20);
                    Rec.POSupplier = Lib.GetTruncated(Dr["posupplier"].ToString(), 20);
                    Rec.OLDPO = Lib.GetTruncated(Dr["oldpo"].ToString(), 20);
                    Rec.HBLNumber = Lib.GetTruncated(Dr["HBLNumber"].ToString(), 20);
                    Rec.DocumentType = Lib.GetTruncated(Dr["DocumentType"].ToString(), 20);
                    Rec.Shipper = Lib.GetTruncated(Dr["Shipper"].ToString(),50);
                    Rec.ShipperCountry = Dr["ShipperCountry"].ToString();

                    str = Dr["shpr_add1"].ToString();
                    AddAddress(ref str, Dr["shpr_add2"].ToString());
                    AddAddress(ref str, Dr["shpr_add3"].ToString());
                    AddAddress(ref str, Dr["shpr_add4"].ToString());
                    //if (Dr["shpr_add2"].ToString() != "")
                    //{
                    //    if (str != "" && !str.Trim().EndsWith(","))
                    //        str += ",";
                    //    str += Dr["shpr_add2"].ToString();
                    //}
                    //if (Dr["shpr_add3"].ToString() != "")
                    //{
                    //    if (str != "" && !str.Trim().EndsWith(","))
                    //        str += ",";
                    //    str += Dr["shpr_add3"].ToString();
                    //}
                    //if (Dr["shpr_add4"].ToString() != "")
                    //{
                    //    if (str != "" && !str.Trim().EndsWith(","))
                    //        str += ",";
                    //    str += Dr["shpr_add4"].ToString();
                    //}
                    Rec.ShipperAddress = str;
                    Rec.ShipperPhoneNumber = Dr["ShipperPhoneNumber"].ToString();
                    Rec.NotifyParty = Lib.GetTruncated(Dr["NotifyParty"].ToString(),50);
                    Rec.MBLNumber = Lib.GetTruncated(Dr["MBLNumber"].ToString(),20);
                    Rec.ServiceContract = Lib.GetTruncated(Dr["ServiceContract"].ToString(), 20);
                    Rec.ContainerNumber = Lib.GetCntrno(Dr["ContainerNumber"].ToString());
                    Rec.ContainerType = Dr["ContainerType"].ToString();
                    Rec.CarrierSeal = Dr["CarrierSeal"].ToString();
                    Rec.ShipperSeal = Dr["ShipperSeal"].ToString();
                    Rec.LoadType = Dr["LoadType"].ToString();
                    Rec.IntertraficInspection = Dr["IntertraficInspection"].ToString();
                    Rec.InspectionIn = Dr["InspectionIn"].ToString();
                    Rec.InspectionSupplier = Dr["InspectionSupplier"].ToString();
                    mOrdList[ArrIndex++] = Rec;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mOrdList;
        }

        private void WriteXmlFiles()
        {
            try
            {
                if (BLMessage == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | BL Vessel Voyage Not Generated.";
                    return;
                }

                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", "");
                XmlSerializer mySerializer = new XmlSerializer(typeof(MessageBLInfo));
                StreamWriter writer = new StreamWriter(File_Name);
                mySerializer.Serialize(writer, BLMessage, ns);
                writer.Close();
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }

        private void GenerateCsvFiles()
        {
            string str = "";
           
            DataRow DR = null;
            if (DT_BLINFO.Rows.Count > 0)
                DR = DT_BLINFO.Rows[0];

            if (DR == null)
                return;

            sb = new StringBuilder();
            sb.Append("Carrier"); sb.Append(",");
            sb.Append("VslName"); sb.Append(",");
            sb.Append("Voygae"); sb.Append(",");
            sb.Append("OriCountry"); sb.Append(",");
            sb.Append("POL"); sb.Append(",");
            sb.Append("POD"); sb.Append(",");
            sb.Append("FinalDestination"); sb.Append(",");
            sb.Append("TSPort"); sb.Append(",");
            sb.Append("TSPortDate"); sb.Append(",");
            sb.Append("CutOff"); sb.Append(",");
            sb.Append("ETD"); sb.Append(",");
            sb.Append("ETA"); sb.Append(",");
            sb.Append("TRAFFIC"); sb.Append(",");
            sb.Append("POID"); sb.Append(",");
            sb.Append("PO"); sb.Append(",");
            sb.Append("StyleNo"); sb.Append(",");
            sb.Append("POSupplier"); sb.Append(",");
            sb.Append("OLDPO"); sb.Append(",");
            sb.Append("HBLNumber"); sb.Append(",");
            sb.Append("DocumentType"); sb.Append(",");
            sb.Append("Shipper"); sb.Append(",");
            sb.Append("ShipperCountry"); sb.Append(",");
            sb.Append("ShipperAddress"); sb.Append(",");
            sb.Append("ShipperPhoneNumber"); sb.Append(",");
            sb.Append("NotifyParty"); sb.Append(",");
            sb.Append("MBLNumber"); sb.Append(",");
            sb.Append("ServiceContract"); sb.Append(",");
            sb.Append("ContainerNumber"); sb.Append(",");
            sb.Append("ContainerType"); sb.Append(",");
            sb.Append("CarrierSeal"); sb.Append(",");
            sb.Append("ShipperSeal"); sb.Append(",");
            sb.Append("LoadType"); sb.Append(",");
            sb.Append("IntertraficInspection"); sb.Append(",");
            sb.Append("InspectionIn"); sb.Append(",");
            sb.Append("InspectionSupplier");
            foreach (DataRow Dr in DT_BLORD.Rows)
            {
                sb.AppendLine();
                sb.Append(Lib.GetTruncated(DR["Carrier"].ToString(), 30)); sb.Append(",");
                sb.Append(Lib.GetTruncated(DR["VslName"].ToString(), 30)); sb.Append(",");
                sb.Append(Lib.GetTruncated(DR["voyage"].ToString(), 30)); sb.Append(",");
                sb.Append(DR["OriCountry"].ToString()); sb.Append(",");
                sb.Append(Lib.GetPortCode(DR["POL"].ToString())); sb.Append(",");
                sb.Append(Lib.GetPortCode(DR["POD"].ToString())); sb.Append(",");
                sb.Append(Lib.GetPortCode(DR["FinalDestination"].ToString())); sb.Append(",");
                sb.Append(DR["TSPort"].ToString()); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(DR["TSPortDate"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(DR["CutOff"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(DR["ETD"])); sb.Append(",");
                sb.Append(Lib.DatetoStringDisplayformat(DR["ETA"])); sb.Append(",");
                sb.Append(DR["Traffic"].ToString()); sb.Append(",");

                sb.Append(Dr["poid"].ToString()); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["po"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["styleno"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["posupplier"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["oldpo"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["HBLNumber"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["DocumentType"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["Shipper"].ToString(), 50)); sb.Append(",");
                sb.Append(Dr["ShipperCountry"].ToString()); sb.Append(",");

                str = Dr["shpr_add1"].ToString();
                AddAddress(ref str, Dr["shpr_add2"].ToString());
                AddAddress(ref str, Dr["shpr_add3"].ToString());
                AddAddress(ref str, Dr["shpr_add4"].ToString());

                sb.Append(str.Replace(",","")); sb.Append(",");
                sb.Append(Dr["ShipperPhoneNumber"].ToString()); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["NotifyParty"].ToString(), 50)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["MBLNumber"].ToString(), 20)); sb.Append(",");
                sb.Append(Lib.GetTruncated(Dr["ServiceContract"].ToString(),20).Replace(",","")); sb.Append(",");
                sb.Append(Lib.GetCntrno(Dr["ContainerNumber"].ToString())); sb.Append(",");
                sb.Append(Dr["ContainerType"].ToString()); sb.Append(",");
                sb.Append(Dr["CarrierSeal"].ToString()); sb.Append(",");
                sb.Append(Dr["ShipperSeal"].ToString()); sb.Append(",");
                sb.Append(Dr["LoadType"].ToString()); sb.Append(",");
                sb.Append(Dr["IntertraficInspection"].ToString()); sb.Append(",");
                sb.Append(Dr["InspectionIn"].ToString()); sb.Append(",");
                sb.Append(Dr["InspectionSupplier"].ToString());
            }
        }

        private void WriteCsvFiles()
        {
            try
            {
                if (sb == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Cargo process Not Generated.";
                    return;
                }

                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                System.IO.File.AppendAllText(File_Name, sb.ToString());
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }

        private void AddAddress(ref string sAddress, string str)
        {
            if (str != "")
            {
                if (sAddress != "" && !sAddress.Trim().EndsWith(","))
                    sAddress += ",";
                sAddress += str;
            }
        }
    }
}
