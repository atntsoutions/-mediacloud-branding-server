using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using BLXml.models;
using BLXml.models.CmarBL;

namespace BLXml
{
    public partial class CmarBillLading : BLXml.models.XmlRoot
    {
        private DataTable DT_SHIPMENT = new DataTable();
        private DataTable DT_MBLCNTR = new DataTable();
        private DataTable DT_HBLCNTR = new DataTable();
        private DataTable DT_HBLMARK = new DataTable();
        private DataTable DT_HBLDESC = new DataTable();

        public Boolean IsError = false;
        public string ErrorMessage = "";
        public string MessageBranchName = "";
        private BillLadingMessage BLMessage = null;


        private string ErrorValues = "";
        private string sql = "";
        private string MessageNumber = "";
        private int MBL_REFNO = 0;

        DBConnection Con_Oracle = null;
        //string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "", comp_add4 = "";
        public override void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;
                ReadData();
                if (DT_SHIPMENT.Rows.Count <= 0)
                {
                    IsError = true;
                    ErrorMessage = "Details not Found";
                    return;
                }
                //   this.MessageNumber = XmlLib.GetNewMessageNumber();
                this.MessageNumber = XmlLib.MessageNumberSeq;
                GenerateXmlFiles();
                WriteXmlFiles();

                DT_SHIPMENT.Rows.Clear();
                DT_MBLCNTR.Rows.Clear();
                DT_HBLCNTR.Rows.Clear();
                DT_HBLDESC.Rows.Clear();
                DT_HBLMARK.Rows.Clear();
            }
            catch (Exception ex)
            {
                IsError = true;
                ErrorMessage += " |" + ex.Message.ToString();
            }
        }

        private void ReadData()
        {
            Con_Oracle = new DBConnection();

            sql = " select mbl.hbl_pkid as MBL_PKID,mbl.hbl_bl_no as MBL_NO,mbl.hbl_no as  MBL_BKNO,";
            sql += " mbl.hbl_date as MBL_DATE,";
            sql += " magnt.cust_code as M_AGENT_CODE,";
            sql += " magnt.cust_name as M_AGENT_NAME,";
            sql += " mlnr.param_code as M_CARRIER_CODE,";
            sql += " mlnr.param_name as M_CARRIER_NAME,";
            sql += " pol.param_code as POL_CODE,";
            sql += " pol.param_name as POL_NAME,";
            sql += " mbl.hbl_pol_etd as POL_ETD,";
            sql += " pod.param_code as POD_CODE,";
            sql += " pod.param_name as POD_NAME,";
            sql += " mbl.hbl_pod_eta as POD_ETA,";
            sql += " pofd.param_name as PLACE_DELIVERY,";
            sql += " mbl.hbl_pofd_eta as DELIVERY_DATE,";//hs.hbls_delivery_date
            sql += " vsl.param_name as VESSEL_NAME,";
            sql += " mbl.hbl_vessel_no as VESSEL_VOYAGE,";
            sql += " mbl.hbl_terms as  M_FREIGHT_STATUS,";
            sql += " cntry.param_name as  ORIGIN_COUNTRY_NAME,";
            sql += " hbl.hbl_pkid as HBL_PKID,hbl.hbl_mbl_id as HBLS_MBL_ID,hbl.hbl_no as  HBL_NO,hbl.hbl_bl_no as HBLS_BL_NO,";
            sql += " shpr.cust_code as SHIPPER_CODE,";
            sql += " shpr.cust_name as SHIPPER_NAME,";
            sql += " shpradd.add_line1 as SHIPPER_ADD1,";
            sql += " shpradd.add_line2 as SHIPPER_ADD2,";
            sql += " shpradd.add_line3 as SHIPPER_ADD3,";
            sql += " shpradd.add_line4 AS SHIPPER_ADD4,";
            sql += " cnge.cust_code as CONSIGNEE_CODE,";
            sql += " cnge.cust_name as CONSIGNEE_NAME,";
            sql += " cngeadd.add_line1 as CONSIGNEE_ADD1,";
            sql += " cngeadd.add_line2 as CONSIGNEE_ADD2,";
            sql += " cngeadd.add_line3 as CONSIGNEE_ADD3,";
            sql += " cngeadd.add_line4 as CONSIGNEE_ADD4,";
            sql += " nfy.cust_code as NOTIFY_CODE,";
            sql += " nfy.cust_name as NOTIFY_NAME,";
            sql += " nfyadd.add_line1 as NOTIFY_ADD1,";
            sql += " nfyadd.add_line2 as NOTIFY_ADD2,";
            sql += " nfyadd.add_line3 as NOTIFY_ADD3,";
            sql += " nfyadd.add_line4 as NOTIFY_ADD4,";
            sql += " agnt.cust_code as AGENT_CODE,";
            sql += " agnt.cust_name as AGENT_NAME,";
            sql += " pofd.param_name as DESTINATION_PLACE,";
            sql += " mbl.hbl_pofd_eta as DESTINATION_ETA,";
            sql += " hbl.hbl_pkg as PACKAGES,";
            sql += " pkgunt.param_code as UOM,";
            sql += " com.param_name as COMMODITY_NAME,";
            sql += " hbl.hbl_grwt as WEIGHT,";
            sql += " hbl.hbl_cbm as CBM,";
            sql += " hbl.hbl_pcs as PCS,";
            sql += " hbl.hbl_terms as  H_FREIGHT_STATUS,";
            sql += " hbl.hbl_nature as SHIPMENT_TERM,";
            sql += " nvl(hbl.hbl_nomination,cnge.cust_nomination) as SHIPMENT_TYPE,";
            sql += " '' as CARGODESCRIPTIONS,";
            sql += " '' as MARKSNUMBERS";

            sql += " from hblm mbl";
            sql += " inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += " inner join jobm job on hbl.hbl_pkid = job.jobs_hbl_id";
            sql += " left join customerm magnt on mbl.hbl_agent_id = magnt.cust_pkid";
            sql += " left join param mlnr on mbl.hbl_carrier_id = mlnr.param_pkid ";
            sql += " left join param pol on job.job_pol_id = pol.param_pkid";
            sql += " left join param pod on job.job_pod_id  = pod.param_pkid";
            sql += " left join param pofd on job.job_pofd_id = pofd.param_pkid";
            sql += " left join param por on job.job_place_receipt_id = por.param_pkid";
            sql += " left join param vsl	on mbl.hbl_vessel_id = vsl.param_pkid";
            sql += " left join param cntry on job.job_origin_country_id = cntry.param_pkid";
            sql += " left join customerm shpr on hbl.hbl_exp_id = shpr.cust_pkid";
            sql += " left join addressm shpradd on hbl.hbl_exp_br_id = shpradd.add_pkid";
            sql += " left join customerm cnge on hbl.hbl_imp_id = cnge.cust_pkid";
            sql += " left join addressm cngeadd on hbl.hbl_imp_br_id = cngeadd.add_pkid";
            sql += " left join bl on hbl.hbl_pkid = bl.bl_pkid";
            sql += " left join customerm nfy on bl.bl_notify_id = nfy.cust_pkid";
            sql += " left join addressm nfyadd on bl.bl_notify_br_id = nfyadd.add_pkid";
            sql += " left join customerm agnt on hbl.hbl_agent_id = agnt.cust_pkid";
            sql += " left join param pkgunt on hbl.hbl_pkg_unit_id = pkgunt.param_pkid";
            sql += " left join param com on job.job_commodity_id = com.param_pkid";

            sql += " where mbl.rec_company_code = '" + XmlLib.Company_Code + "' ";
            sql += " and mbl.rec_branch_code  = '" + XmlLib.Branch_Code + "' ";
            sql += " and mbl.rec_category  = 'SEA EXPORT' ";
            sql += " and mbl.hbl_agent_id = '" + XmlLib.Agent_Id + "' ";
            sql += " and mbl.hbl_pkid in (" + XmlLib.MBL_IDS + ")";


            DT_SHIPMENT = new DataTable();
            DT_SHIPMENT = Con_Oracle.ExecuteQuery(sql);


            sql = " select pack_mbl_id,pack_container_id,";
            sql += " max(cntr_no) as cntr_no,";
            sql += " max(cntr_type) as cntr_type,";
            sql += " max(cntr_asealno) as cntr_sealno, ";
            sql += " sum(pack_pkg) as cntr_pcs, max(pkg_uom) as cntr_uom,";
            sql += " sum(pack_grwt) as cntr_weight,";
            sql += " sum(pack_cbm) as cntr_cbm,";
            sql += " max(cntr_shipment_term) as cntr_shipment_term,";
            sql += " max(cntr_shipment_type) as cntr_shipment_type ";
            sql += " from (";
            sql += " select mbl.hbl_pkid as pack_mbl_id,pack_cntr_id as pack_container_id, cntr_no, d.param_code as cntr_type ,cntr_csealno ,cntr_asealno ";
            sql += " ,pack_pkg,pkgunt.param_code as pkg_uom, pack_pcs ,pack_cbm,pack_ntwt,pack_grwt ";
            sql += " ,mbl.hbl_nature as cntr_shipment_term,mbl.hbl_shipment_type as  cntr_shipment_type ";
            sql += " from hblm mbl";
            sql += " inner join hblm hbl on mbl.hbl_pkid = hbl.hbl_mbl_id";
            sql += " inner join jobm a on hbl.hbl_pkid = a.jobs_hbl_id ";
            sql += " inner join packingm b on a.job_pkid = b.pack_job_id ";
            sql += " inner join containerm c on b.pack_cntr_id = c.cntr_pkid ";
            sql += " left join param d on c.cntr_type_id = d.param_pkid";
            sql += " left join param pkgunt on b.pack_pkg_unit_id = pkgunt.param_pkid";
            sql += " where mbl.hbl_pkid in (" + XmlLib.MBL_IDS + ")";
            sql += " )a group by pack_mbl_id,pack_container_id";

            DT_MBLCNTR = new DataTable();
            DT_MBLCNTR = Con_Oracle.ExecuteQuery(sql);


            sql = " select pack_hbl_id,pack_container_id,";
            sql += " max(cntr_no) as cntr_no,";
            sql += " max(cntr_type) as cntr_type,";
            sql += " max(cntr_asealno) as cntr_sealno, ";
            sql += " sum(pack_pkg) as cntr_pcs, max(pkg_uom) as cntr_uom,";
            sql += " sum(pack_grwt) as cntr_weight,";
            sql += " sum(pack_cbm) as cntr_cbm";
            sql += " from (";
            sql += " select hbl.hbl_pkid as pack_hbl_id,pack_cntr_id as pack_container_id, cntr_no, d.param_code as cntr_type ,cntr_csealno ,cntr_asealno ";
            sql += " ,pack_pkg,pkgunt.param_code as pkg_uom, pack_pcs ,pack_cbm,pack_ntwt,pack_grwt ";
            sql += " from hblm hbl";
            sql += " inner join jobm a on hbl.hbl_pkid = a.jobs_hbl_id ";
            sql += " inner join packingm b on a.job_pkid = b.pack_job_id ";
            sql += " inner join containerm c on b.pack_cntr_id = c.cntr_pkid ";
            sql += " left join param d on c.cntr_type_id = d.param_pkid";
            sql += " left join param pkgunt on b.pack_pkg_unit_id = pkgunt.param_pkid";
            sql += " where hbl.hbl_mbl_id in (" + XmlLib.MBL_IDS + ")";
            sql += " )a group by pack_hbl_id,pack_container_id";


            DT_HBLCNTR = new DataTable();
            DT_HBLCNTR = Con_Oracle.ExecuteQuery(sql);

            sql = "select hbl.hbl_pkid,bl_marks,bl_desc_ctr from bldesc a";
            sql += " inner join hblm hbl on a.bl_parent_id = hbl.hbl_pkid";
            sql += " where hbl.hbl_mbl_id in (" + XmlLib.MBL_IDS + ")";
            sql += " order by bl_desc_ctr";

            DT_HBLMARK = new DataTable();
            DT_HBLMARK = Con_Oracle.ExecuteQuery(sql);


            sql = "select hbl.hbl_pkid,bl_desc,bl_desc_ctr from bldesc a";
            sql += " inner join hblm hbl on a.bl_parent_id = hbl.hbl_pkid";
            sql += " where hbl.hbl_mbl_id in (" + XmlLib.MBL_IDS + ")";
            sql += " and bl_desc_ctr > 2 order by bl_desc_ctr";

            DT_HBLDESC = new DataTable();
            DT_HBLDESC = Con_Oracle.ExecuteQuery(sql);

            Con_Oracle.CloseConnection();

            if (DT_SHIPMENT.Rows.Count > 0)
            {
                DataRow dr = DT_SHIPMENT.Rows[0];
                if (!dr["POL_ETD"].Equals(DBNull.Value) && (dr["POL_ETD"].ToString() == dr["POD_ETA"].ToString()))
                {
                    IsError = true;
                    ErrorMessage += " | ETD And ETA Cannot be Same [MBLBK# " + dr["MBL_BKNO"].ToString() + "]";
                }
            }

            foreach (DataRow dr in DT_SHIPMENT.Rows)
            {
                dr["HBLS_BL_NO"] = dr["HBLS_BL_NO"].ToString().Replace("/", "");
                if (Lib.Conv2Decimal(dr["PACKAGES"].ToString()) <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | No. of Packages in JOB Cannot be Blank [SI# " + dr["HBL_NO"].ToString() + "]";
                }
                if (Lib.Conv2Decimal(dr["WEIGHT"].ToString()) <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Gr. Wt in JOB Cannot be Blank [SI# " + dr["HBL_NO"].ToString() + "]";
                }

                if (Lib.Conv2Decimal(dr["CBM"].ToString()) <= 0 && GetCntrType(dr["MBL_PKID"].ToString()) != "FCL")
                {
                    IsError = true;
                    ErrorMessage += " | Cbm in JOB Cannot be Blank [SI# " + dr["HBL_NO"].ToString() + "]";
                }
            }
            DT_SHIPMENT.AcceptChanges();
        }

        private string GetCntrType(string MBL_ID)
        {
            string mbl_cntr_type = "FCL";
            foreach (DataRow Dr in DT_MBLCNTR.Select("pack_mbl_id = '" + MBL_ID + "'", "cntr_no"))
            {
                mbl_cntr_type = Dr["cntr_shipment_type"].ToString();
                mbl_cntr_type = mbl_cntr_type.Replace("BUYERS CONSOLE", "FCL");
                break;
            }
            return mbl_cntr_type;
        }

        private void GenerateXmlFiles()
        {
            BLMessage = new BillLadingMessage();
            BLMessage.Items = Generate_MasterList_();
        }

        private object[] Generate_MasterList_()
        {
            object[] Items = null;
            int ArrIndex = 0;
            int iTotRows = 0;
            try
            {
                DataTable DistinctMBL = DT_SHIPMENT.DefaultView.ToTable(true, "mbl_pkid", "mbl_no");
                iTotRows = DistinctMBL.Rows.Count;

                Items = new object[iTotRows + 1];
                Items[ArrIndex++] = Generate_MessageInfo_();
                MBL_REFNO = 0;
                foreach (DataRow dr in DistinctMBL.Select("1=1", "mbl_no"))
                {
                    MBL_REFNO++;
                    Items[ArrIndex++] = Generate_Master_(dr["mbl_pkid"].ToString());
                    if (IsError)
                        break;
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return Items;
        }
        private BillLadingMessageMessageInfo Generate_MessageInfo_()
        {
            BillLadingMessageMessageInfo Rec = null;
            try
            {
                Rec = new BillLadingMessageMessageInfo();
                Rec.messagesender = XmlLib.messageSenderField;
                // Rec.messagenumber = XmlLib.GetNewMessageNumber();
                Rec.messagenumber = XmlLib.MessageNumberSeq.PadLeft(11, '0');
                Rec.mesagedates = ConvertYMDDate(DateTime.Now.ToString());
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return Rec;
        }

        private BillLadingMessageMaster Generate_Master_(string MBL_ID)
        {
            BillLadingMessageMaster Rec = null;
            string mbl_cntr_type = "";
            string mbl_shipment_term = "";
            string PreData = "1";
            try
            {
                foreach (DataRow Dr in DT_MBLCNTR.Select("pack_mbl_id = '" + MBL_ID + "'", "cntr_no"))
                {
                    mbl_cntr_type = Dr["cntr_shipment_type"].ToString();
                    mbl_cntr_type = mbl_cntr_type.Replace("BUYERS CONSOLE", "FCL");
                    mbl_shipment_term = Dr["cntr_shipment_term"].ToString();
                    break;
                }

                foreach (DataRow Dr in DT_SHIPMENT.Select("mbl_pkid ='" + MBL_ID + "'", "mbl_pkid"))
                {
                    if (PreData != Dr["mbl_pkid"].ToString())
                    {
                        PreData = Dr["mbl_pkid"].ToString();

                        Rec = new BillLadingMessageMaster();
                        Rec.branch_name = MessageBranchName;
                        Rec.refno = MBL_REFNO.ToString();
                        Rec.ref_date = ConvertYMDDate(DateTime.Now.ToString());
                        Rec.mbl_no = Dr["mbl_no"].ToString();
                        Rec.mbl_date = ConvertYMDDate(Dr["mbl_date"].ToString());
                        Rec.agent_code = Dr["m_agent_code"].ToString();
                        Rec.agent_name = Dr["m_agent_name"].ToString();
                        Rec.carrier_code = Dr["m_carrier_code"].ToString();
                        Rec.carrier_name = Dr["m_carrier_name"].ToString();
                        Rec.pol_code = Dr["pol_code"].ToString();
                        Rec.pol_name = Dr["pol_name"].ToString();
                        Rec.pol_etd = ConvertYMDDate(Dr["pol_etd"].ToString());
                        Rec.pod_code = Dr["pod_code"].ToString();
                        Rec.pod_name = Dr["pod_name"].ToString();
                        Rec.pod_eta = ConvertYMDDate(Dr["pod_eta"].ToString());
                        Rec.place_delivery = Dr["place_delivery"].ToString();
                        Rec.delivery_date = ConvertYMDDate(Dr["delivery_date"].ToString());
                        Rec.vessel = Dr["vessel_name"].ToString();
                        Rec.voyage = Dr["vessel_voyage"].ToString();
                        Rec.freight_status = Dr["m_freight_status"].ToString();
                        Rec.shipment_term = mbl_shipment_term;
                        Rec.cntr_type = mbl_cntr_type;
                        Rec.origin_country_name = Dr["origin_country_name"].ToString();
                        Rec.Container = Generate_MasterContainerList_(MBL_ID);
                        Rec.HouseBillLading = Generate_HouseList_(MBL_ID);
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return Rec;
        }

        private ContainerCntr[] Generate_MasterContainerList_(string MBL_ID)
        {
            ContainerCntr Rec = null;
            ContainerCntr[] mCntrList = null;
            DataRow[] DrMCntrs = null;
            string PreData = "1";
            int ArrIndex = 0;
            try
            {

                DrMCntrs = DT_MBLCNTR.Select("pack_mbl_id = '" + MBL_ID + "'", "cntr_no");

                mCntrList = new ContainerCntr[DrMCntrs.Length];
                foreach (DataRow Dr in DrMCntrs)
                {
                    if (PreData != Dr["cntr_no"].ToString())
                    {
                        PreData = Dr["cntr_no"].ToString();

                        Rec = new ContainerCntr();
                        Rec.slno = (ArrIndex + 1).ToString();
                        Rec.cntr_no = Lib.GetCntrno(Dr["cntr_no"].ToString()).Replace("-", "").Replace(" ", "");
                        Rec.cntr_type = Dr["cntr_type"].ToString();
                        Rec.cntr_sealno = Dr["cntr_sealno"].ToString();
                        Rec.cntr_pcs = Dr["cntr_pcs"].ToString();
                        Rec.cntr_uom = Dr["cntr_uom"].ToString();
                        Rec.cntr_weight = Dr["cntr_weight"].ToString();
                        Rec.cntr_cbm = Dr["cntr_cbm"].ToString();

                        mCntrList[ArrIndex++] = Rec;
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return mCntrList;
        }

        private BillLadingMessageMasterHouseBillLadingHouse[] Generate_HouseList_(string MBL_ID)
        {
            BillLadingMessageMasterHouseBillLadingHouse Rec = null;
            BillLadingMessageMasterHouseBillLadingHouse[] hList = null;
            DataRow[] DrHBLs = null;
            string PreData = "1";
            int ArrIndex = 0;
            try
            {

                DrHBLs = DT_SHIPMENT.Select("hbls_mbl_id = '" + MBL_ID + "'", "hbl_pkid");

                hList = new BillLadingMessageMasterHouseBillLadingHouse[DrHBLs.Length];
                foreach (DataRow Dr in DrHBLs)
                {
                    if (PreData != Dr["hbl_pkid"].ToString())
                    {
                        PreData = Dr["hbl_pkid"].ToString();
                        Rec = new BillLadingMessageMasterHouseBillLadingHouse();

                        Rec.slno = (ArrIndex + 1).ToString();
                        Rec.house_no = Dr["hbls_bl_no"].ToString();
                        Rec.shipper_code = Dr["shipper_code"].ToString();
                        Rec.shipper_name = Dr["shipper_name"].ToString();
                        Rec.shipper_add1 = Dr["shipper_add1"].ToString();
                        Rec.shipper_add2 = Dr["shipper_add2"].ToString();
                        Rec.shipper_add3 = Dr["shipper_add3"].ToString();
                        Rec.shipper_add4 = Dr["shipper_add4"].ToString();
                        Rec.consignee_code = Dr["consignee_code"].ToString();
                        Rec.consignee_name = Dr["consignee_name"].ToString();
                        Rec.consignee_add1 = Dr["consignee_add1"].ToString();
                        Rec.consignee_add2 = Dr["consignee_add2"].ToString();
                        Rec.consignee_add3 = Dr["consignee_add3"].ToString();
                        Rec.consignee_add4 = Dr["consignee_add4"].ToString();

                        Rec.notify_name = Dr["notify_name"].ToString();
                        Rec.notify_add1 = Dr["notify_add1"].ToString();
                        Rec.notify_add2 = Dr["notify_add2"].ToString();
                        Rec.notify_add3 = Dr["notify_add3"].ToString();
                        Rec.notify_add4 = Dr["notify_add4"].ToString();
                        Rec.agent_code = Dr["agent_code"].ToString();
                        Rec.agent_name = Dr["agent_name"].ToString();
                        Rec.place_delivery = Dr["place_delivery"].ToString();
                        Rec.delivery_date = ConvertYMDDate(Dr["delivery_date"].ToString());
                        Rec.destination_place = Dr["destination_place"].ToString();
                        Rec.destination_eta = ConvertYMDDate(Dr["destination_eta"].ToString());

                        Rec.packages = Dr["packages"].ToString();
                        Rec.uom = Dr["uom"].ToString();
                        Rec.commodity = Dr["commodity_name"].ToString();
                        Rec.weight = Dr["weight"].ToString();
                        Rec.lbs = Convert_Weight("KG2LBS", Dr["weight"].ToString(), 3);
                        Rec.cbm = Dr["cbm"].ToString();
                        Rec.cft = Convert_Weight("CBM2CFT", Dr["cbm"].ToString(), 3);
                        Rec.pcs = Dr["pcs"].ToString();
                        Rec.freight_status = Dr["h_freight_status"].ToString();
                        Rec.shipment_term = Dr["shipment_term"].ToString();
                        Rec.shipment_type = Dr["shipment_type"].ToString();
                        Rec.pono = "";
                        Rec.ams_fileno = "";
                        Rec.sub_house_no = "";
                        Rec.isf_no = "";
                        Rec.bl_req = "";
                        Rec.remark1 = "";
                        Rec.remark2 = "";
                        Rec.remark3 = "";

                        Rec.Container = Generate_HouseContainerList_(Dr["hbl_pkid"].ToString());
                        Rec.MarksNumbers = Generate_MarksNumbers_(Dr["hbl_pkid"].ToString());
                        Rec.CargoDescriptions = Generate_CargoDescriptions_(Dr["hbl_pkid"].ToString());

                        hList[ArrIndex++] = Rec;
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return hList;
        }
        private ContainerCntr[] Generate_HouseContainerList_(string HBL_ID)
        {
            ContainerCntr Rec = null;
            ContainerCntr[] hCntrList = null;
            DataRow[] DrHCntrs = null;
            string PreData = "1";
            int ArrIndex = 0;
            try
            {

                DrHCntrs = DT_HBLCNTR.Select("pack_hbl_id = '" + HBL_ID + "'", "cntr_no");

                hCntrList = new ContainerCntr[DrHCntrs.Length];
                foreach (DataRow Dr in DrHCntrs)
                {
                    if (PreData != Dr["cntr_no"].ToString())
                    {
                        PreData = Dr["cntr_no"].ToString();

                        Rec = new ContainerCntr();

                        Rec.slno = (ArrIndex + 1).ToString();
                        Rec.cntr_no = Lib.GetCntrno(Dr["cntr_no"].ToString()).Replace("-", "").Replace(" ", "");
                        Rec.cntr_type = Dr["cntr_type"].ToString();
                        Rec.cntr_sealno = Dr["cntr_sealno"].ToString();
                        Rec.cntr_pcs = Dr["cntr_pcs"].ToString();
                        Rec.cntr_uom = Dr["cntr_uom"].ToString();
                        Rec.cntr_weight = Dr["cntr_weight"].ToString();
                        Rec.cntr_cbm = Dr["cntr_cbm"].ToString();

                        hCntrList[ArrIndex++] = Rec;
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return hCntrList;
        }

        private BillLadingMessageMasterHouseBillLadingHouseMarksNumbersMarksNumber[] Generate_MarksNumbers_(string HBL_ID)
        {
            BillLadingMessageMasterHouseBillLadingHouseMarksNumbersMarksNumber Rec = null;
            BillLadingMessageMasterHouseBillLadingHouseMarksNumbersMarksNumber[] hMarks = null;
            string[] HBL_MARKS = null;
            DataRow[] DrMark = null;
            int ArrIndex = 0;
            try
            {
                //MarksNumbers = MarksNumbers.Replace("\r", "");
                //HBL_MARKS = MarksNumbers.Split('\n');

                //hMarks = new BillLadingMessageMasterHouseBillLadingHouseMarksNumbersMarksNumber[HBL_MARKS.Length];
                //for (int i = 0; i < HBL_MARKS.Length; i++)
                //{
                //    Rec = new BillLadingMessageMasterHouseBillLadingHouseMarksNumbersMarksNumber();
                //    Rec.rownum = (i + 1).ToString();
                //    Rec.cargo_marks = HBL_MARKS[i].ToString();

                //    hMarks[i] = Rec;
                //}

                DrMark = DT_HBLMARK.Select("hbl_pkid = '" + HBL_ID + "'", "bl_desc_ctr");
                hMarks = new BillLadingMessageMasterHouseBillLadingHouseMarksNumbersMarksNumber[DrMark.Length];
                foreach (DataRow dr in DrMark)
                {
                    if (!dr["bl_marks"].Equals(DBNull.Value))
                    {
                        if (dr["bl_marks"].ToString().Contains("QUANTITY/QUALITY AS PER SHIPPER") || dr["bl_marks"].ToString().Contains("CARRIER NOT RESPONSIBLE FOR"))
                            continue;
                        Rec = new BillLadingMessageMasterHouseBillLadingHouseMarksNumbersMarksNumber();
                        Rec.rownum = (ArrIndex + 1).ToString();
                        Rec.cargo_marks = dr["bl_marks"].ToString();
                        hMarks[ArrIndex++] = Rec;
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return hMarks;
        }

        private BillLadingMessageMasterHouseBillLadingHouseCargoDescriptionsCargoDescription[] Generate_CargoDescriptions_(string HBL_ID)
        {
            BillLadingMessageMasterHouseBillLadingHouseCargoDescriptionsCargoDescription Rec = null;
            BillLadingMessageMasterHouseBillLadingHouseCargoDescriptionsCargoDescription[] hDescs = null;
            string[] HBL_DESCS = null;
            DataRow[] DrDesc = null;
            int ArrIndex = 0;

            try
            {
                //CargoDescriptions = CargoDescriptions.Replace("\r", "");
                //HBL_DESCS = CargoDescriptions.Split('\n');

                //hDescs = new BillLadingMessageMasterHouseBillLadingHouseCargoDescriptionsCargoDescription[HBL_DESCS.Length];
                //for (int i = 0; i < HBL_DESCS.Length; i++)
                //{
                //    Rec = new BillLadingMessageMasterHouseBillLadingHouseCargoDescriptionsCargoDescription();
                //    Rec.rownum = (i + 1).ToString();
                //    Rec.cargo_description = HBL_DESCS[i].ToString();

                //    hDescs[i] = Rec;
                //}
                DrDesc = DT_HBLDESC.Select("hbl_pkid = '" + HBL_ID + "'", "bl_desc_ctr");
                hDescs = new BillLadingMessageMasterHouseBillLadingHouseCargoDescriptionsCargoDescription[DrDesc.Length];
                foreach (DataRow dr in DrDesc)
                {
                    if (!dr["bl_desc"].Equals(DBNull.Value))
                    {
                        Rec = new BillLadingMessageMasterHouseBillLadingHouseCargoDescriptionsCargoDescription();
                        Rec.rownum = (ArrIndex + 1).ToString();
                        Rec.cargo_description = dr["bl_desc"].ToString();
                        hDescs[ArrIndex++] = Rec;
                    }
                }
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
            return hDescs;
        }

        private void WriteXmlFiles()
        {
            try
            {
                if (BLMessage == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | BillLading Not Generated.";
                    return;
                }


                if (XmlLib.SaveInTempFolder)
                {
                    XmlLib.FolderId = Guid.NewGuid().ToString().ToUpper();
                    XmlLib.File_Type = "XML";
                    XmlLib.File_Category = "BL";
                    XmlLib.File_Processid = "BL" + this.MessageNumber;
                    // FileName = XmlLib.Agent_Name.Replace(",", "").Replace(" ", "");
                    FileName = "CMAR";
                    XmlLib.File_Display_Name = FileName + this.MessageNumber.PadLeft(11, '0') + ".XML";
                    XmlLib.File_Name = Lib.GetFileName(XmlLib.report_folder, XmlLib.FolderId, XmlLib.File_Display_Name);
                    XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                    ns.Add("", "");
                    XmlSerializer serializer =
                        new XmlSerializer(typeof(BillLadingMessage));
                    TextWriter writer = new StreamWriter(XmlLib.File_Name);
                    serializer.Serialize(writer, BLMessage, ns);
                    writer.Close();
                }

            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
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

        private string Convert_Weight(string sType, string data, int iDec)
        {
            decimal iData = 0;
            try
            {
                if (sType == "KG2LBS")
                    iData = Lib.Convert2Decimal(data) * (decimal)2.2046;
                if (sType == "CBM2CFT")
                    iData = Lib.Convert2Decimal(data) * (decimal)35.314;

                //if (sType == "LBS2KG")
                //    iData = Convert2Decimal(data) / (decimal)2.2046;
                //if (sType == "CFT2CBM")
                //    iData = Convert2Decimal(data) / (decimal)35.314;
            }
            catch (Exception Ex)
            {
                iData = 0;
                throw Ex;
            }
            return Lib.NumericFormat(iData.ToString(), iDec);
        }

    }
}
