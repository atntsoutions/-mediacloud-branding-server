using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLAccounts
{
    public class InvoiceService : BL_Base
    {
        List<InvoiceReport> DetailList = new List<InvoiceReport>();
        InvoiceReport HeaderRow;
        private DataTable Dt_MblInv = null;
        private DataTable Dt_MblOs = null;
        private DataTable Dt_InvFooter = null;
        private string report_folder = "";
        private string folderid = "";
        private string branch_code = "";
        private string company_code = "";
        private string pkid = "";
        private string report_format = "";
        private string header_format = "";
        private string detail_format = "";
        private string report_caption = "";
        private string menu_admin = "N";
        private decimal jvNetTot = 0;
        private bool IsIGST = false;
        private bool IsCSGST = false;
        private int ParticularColWidth = 0;
        Dictionary<string, string> BankInfo = new Dictionary<string, string>();
        public Dictionary<string, object> GenerateInvoice(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            report_folder = "";
            folderid = "";
            branch_code = "";
            pkid = SearchData["pkid"].ToString();
            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();

            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();

            if (SearchData.ContainsKey("branch_code"))
                branch_code = SearchData["branch_code"].ToString();

            if (SearchData.ContainsKey("report_format"))
                report_format = SearchData["report_format"].ToString();

            if (SearchData.ContainsKey("report_caption"))
                report_caption = SearchData["report_caption"].ToString();

            if (SearchData.ContainsKey("company_code"))
                company_code = SearchData["company_code"].ToString();

            if (SearchData.ContainsKey("menuadmin"))
                menu_admin = SearchData["menuadmin"].ToString();

            ReadInvoiceData();
            if (jvNetTot != Lib.Conv2Decimal(HeaderRow.jvh_net_amt.ToString()))
            {
                throw new Exception("Amount Mismatch");
            }

            InvoiceReportService InvRpt = new InvoiceReportService();
            InvRpt.hRow = HeaderRow;
            InvRpt.DT_MBLINV = Dt_MblInv;
            InvRpt.DT_MBLOS = Dt_MblOs;
            InvRpt.DT_INVFOOTER = Dt_InvFooter;
            InvRpt.DetList = DetailList;
            InvRpt.folderid = folderid;
            InvRpt.report_folder = report_folder;
            InvRpt.Report_format = report_format;
            InvRpt.HeaderFormat = header_format;
            InvRpt.DetailFormat = detail_format;
            InvRpt.ReportCaption = report_caption;
            InvRpt.company_code = company_code;
            InvRpt.branch_code = branch_code;
            InvRpt.BankInfoDic = BankInfo;
            InvRpt.Process();
            RetData.Add("filename", InvRpt.File_Name);
            RetData.Add("filetype", InvRpt.File_Type);
            RetData.Add("filedisplayname", InvRpt.File_Display_Name);

            return RetData;
        }

        private void ReadInvoiceData()
        {
            HeaderRow = new InvoiceReport();
            HeaderRow = InitRecord();
            ReadCompanyDetails();
            ReadInvoiceHeader();
            SetPrintFormat();
            ReadInvoiceDetails();
            if (HeaderRow.jvh_type == "IN")
                ReadInvoiceFooter();
        }
        private InvoiceReport InitRecord()
        {
            InvoiceReport Rec = new InvoiceReport();
            Rec.inv_comp_name = "";
            Rec.inv_comp_add1 = "";
            Rec.inv_comp_add2 = "";
            Rec.inv_comp_add3 = "";
            Rec.inv_comp_tel = "";
            Rec.inv_comp_fax = "";
            Rec.inv_comp_web = "";
            Rec.inv_comp_email = "";
            Rec.inv_comp_panno = "";
            Rec.inv_comp_uamno = "";
            Rec.inv_comp_cinno = "";
            Rec.inv_comp_gstin = "";
            Rec.inv_comp_reg_address = "";
            Rec.jvh_docno = "";
            Rec.jvh_date = "";
            Rec.jvh_gstin = "";
            Rec.jvh_state_code = "";
            Rec.jvh_state_name = "";
            Rec.jvh_cc_category = "";
            Rec.jvh_party_name = "";
            Rec.jvh_party_addr1 = "";
            Rec.jvh_party_addr2 = "";
            Rec.jvh_party_addr3 = "";
            Rec.jvh_tot_amt = 0;
            Rec.jvh_cgst_amt = 0;
            Rec.jvh_sgst_amt = 0;
            Rec.jvh_igst_amt = 0;
            Rec.jvh_net_amt = 0;
            Rec.jvh_curr_code = "";
            Rec.jvh_tot_famt = 0;
            Rec.jvh_cgst_famt = 0;
            Rec.jvh_sgst_famt = 0;
            Rec.jvh_igst_famt = 0;
            Rec.jvh_net_famt = 0;
            Rec.jvh_narration = "";
            Rec.jvh_type = "";
            Rec.jvh_reference = "";
            Rec.jvh_org_invno = "";
            Rec.jvh_acc_id = "";
            Rec.hbl_exp_id = "";
            Rec.hbl_exp_name = "";
            Rec.hbl_bl_no = "";
            Rec.hbl_no = "";
            Rec.hbl_date = "";
            Rec.mbl_bkno = "";
            Rec.mbl_no = "";
            Rec.mbl_date = "";
            Rec.mbl_vessel_name = "";
            Rec.mbl_vessel_voyage = "";
            Rec.mbl_pol_etd = "";
            Rec.hbl_pol_name = "";
            Rec.hbl_pod_name = "";
            Rec.hbl_pofd_name = "";
            Rec.hbl_freight_status = "";
            Rec.hbl_pkg = 0;
            Rec.hbl_pkg_unit = "";
            Rec.hbl_cbm = 0;
            Rec.hbl_consignee_name = "";
            Rec.hbl_imp_id = "";
            Rec.mbl_cha_name = "";
            Rec.hbl_containers = "";
            Rec.jv_ctr = 0;
            Rec.jv_sac_code = "";
            Rec.jv_acc_name = "";
            Rec.jv_total = 0;
            Rec.jv_cgst_rate = 0;
            Rec.jv_cgst_amt = 0;
            Rec.jv_sgst_rate = 0;
            Rec.jv_sgst_amt = 0;
            Rec.jv_igst_rate = 0;
            Rec.jv_igst_amt = 0;
            Rec.jv_net_total = 0;
            Rec.job_sbnos = "";
            Rec.job_invnos = "";
            Rec.job_commodity = "";
            Rec.mbl_igmno = "";
            Rec.mbl_igmdate = "";
            Rec.hbl_beno = "";
            Rec.hbl_bedate = "";
            Rec.hbl_ntwt = 0;
            Rec.hbl_grwt = 0;
            Rec.hbl_chwt = 0;
            Rec.hbl_carton_nos = "";
            Rec.hbl_wt_unit = "";
            Rec.hbl_invoice_nos = "";
            Rec.hbl_genjobtype_code = "";
            Rec.hbl_invnos_prncount = 0;
            Rec.jv_qty = 0;
            Rec.jv_rate = 0;
            Rec.jv_ftotal = 0;
            Rec.jv_exrate = 0;
            Rec.jv_curr_code = "";
            Rec.jv_total_fc = 0;
            Rec.jv_cgst_famt = 0;
            Rec.jv_sgst_famt = 0;
            Rec.jv_igst_famt = 0;
            Rec.jv_net_ftotal = 0;
            Rec.gj_seal_no = "";
            Rec.gj_loaded_on = "";
            Rec.gj_unloaded_on = "";
            Rec.gj_cfs = "";
            Rec.gj_from = "";
            Rec.gj_to1 = "";
            Rec.gj_to2 = "";
            Rec.hbl_genigm_no = "";
            Rec.gj_shipper_inv_no = "";
            Rec.job_comm_invnos = "";
            Rec.gj_mbl_no = "";
            Rec.gj_hbl_no = "";
            Rec.gj_frt_status = "";
            Rec.gj_cha_name = "";
            Rec.gj_consignee_name = "";
            Rec.gj_sb_no = "";
            Rec.gj_vessel = "";
            Rec.gj_lr_no = "";
            Rec.gj_pack_list_no = "";
            return Rec;
        }
        private void ReadCompanyDetails()
        {
            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "ADDRESS");
            mSearchData.Add("branch_code", branch_code); 

            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    if (branch_code == "KOLAF")
                        HeaderRow.inv_comp_name = Dr["BR_NAME"].ToString();
                    else
                        HeaderRow.inv_comp_name = Dr["COMP_NAME"].ToString();
                    HeaderRow.inv_comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    HeaderRow.inv_comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    HeaderRow.inv_comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    HeaderRow.inv_comp_tel = Dr["COMP_TEL"].ToString();
                    HeaderRow.inv_comp_fax = Dr["COMP_FAX"].ToString();
                    HeaderRow.inv_comp_web = Dr["COMP_WEB"].ToString();
                    HeaderRow.inv_comp_email = Dr["COMP_EMAIL"].ToString();
                    HeaderRow.inv_comp_cinno = Dr["COMP_CINNO"].ToString();
                    HeaderRow.inv_comp_panno = Dr["COMP_PANNO"].ToString();
                    HeaderRow.inv_comp_gstin = Dr["COMP_GSTIN"].ToString();
                    HeaderRow.inv_comp_reg_address = Dr["COMP_REG_ADDRESS"].ToString();
                    HeaderRow.inv_comp_uamno = Dr["COMP_UAMNO"].ToString();
                    break;
                }
            }


        }

        private void ReadInvoiceHeader()
        {
            Con_Oracle = new DBConnection();
            string str = "";
            DateTime Dt_Date;

            sql = " select jvh_docno,jvh_date,jvh_gstin,jvh_cc_category,jvh_cc_id";
            sql += "  ,st.param_code as jvh_state_code,st.param_name as jvh_state_name";
            sql += "  ,party.cust_name as jvh_party_name,partyaddr.add_line1 as jvh_party_addr1";
            sql += "  ,partyaddr.add_line2 as jvh_party_addr2,partyaddr.add_line3 as jvh_party_addr3";
            sql += "  ,jvh_tot_amt,jvh_cgst_amt,jvh_sgst_amt,jvh_igst_amt,jvh_net_amt ";
            sql += "  ,curr.param_code as jvh_curr_code";
            sql += "  ,jvh_tot_famt,jvh_cgst_famt,jvh_sgst_famt,jvh_igst_famt,jvh_net_famt,jvh_narration ";
            sql += "  ,jvh_type,jvh_reference,jvh_reference_date,jvh_acc_id,jvh_banktype,jvh_org_invno,jvh_org_invdt ";
            sql += "  from ledgerh a";
            sql += "  left join customerm party on a.jvh_acc_id = party.cust_pkid";
            sql += "  left join addressm partyaddr on a.jvh_acc_br_id = partyaddr.add_pkid";
            sql += "  left join param st on a.jvh_state_id = st.param_pkid";
            sql += "  left join param curr on a.jvh_curr_id = curr.param_pkid";
            sql += "  where jvh_pkid ='" + pkid + "'";
            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow dr in Dt_Rec.Rows)
            {
                HeaderRow.jvh_docno = dr["jvh_docno"].ToString();
                HeaderRow.jvh_date = Lib.DatetoStringDisplayformat(dr["jvh_date"]);
                HeaderRow.jvh_gstin = dr["jvh_gstin"].ToString();
                HeaderRow.jvh_state_code = dr["jvh_state_code"].ToString();
                HeaderRow.jvh_state_name = dr["jvh_state_name"].ToString();
                HeaderRow.jvh_cc_category = dr["jvh_cc_category"].ToString();
                HeaderRow.jvh_party_name = dr["jvh_party_name"].ToString();
                HeaderRow.jvh_party_addr1 = dr["jvh_party_addr1"].ToString();
                HeaderRow.jvh_party_addr2 = dr["jvh_party_addr2"].ToString();
                HeaderRow.jvh_party_addr3 = dr["jvh_party_addr3"].ToString();
                HeaderRow.jvh_tot_amt = Lib.Conv2Decimal(dr["jvh_tot_amt"].ToString());
                HeaderRow.jvh_cgst_amt = Lib.Conv2Decimal(dr["jvh_cgst_amt"].ToString());
                HeaderRow.jvh_sgst_amt = Lib.Conv2Decimal(dr["jvh_sgst_amt"].ToString());
                HeaderRow.jvh_igst_amt = Lib.Conv2Decimal(dr["jvh_igst_amt"].ToString());
                HeaderRow.jvh_net_amt = Lib.Conv2Decimal(dr["jvh_net_amt"].ToString());
                HeaderRow.jvh_curr_code = dr["jvh_curr_code"].ToString();
                HeaderRow.jvh_tot_famt = Lib.Conv2Decimal(dr["jvh_tot_famt"].ToString());
                HeaderRow.jvh_cgst_famt = Lib.Conv2Decimal(dr["jvh_cgst_famt"].ToString());
                HeaderRow.jvh_sgst_famt = Lib.Conv2Decimal(dr["jvh_sgst_famt"].ToString());
                HeaderRow.jvh_igst_famt = Lib.Conv2Decimal(dr["jvh_igst_famt"].ToString());
                HeaderRow.jvh_net_famt = Lib.Conv2Decimal(dr["jvh_net_famt"].ToString());
                HeaderRow.jvh_narration = dr["jvh_narration"].ToString();
                HeaderRow.jvh_type = dr["jvh_type"].ToString();
                str = dr["jvh_reference"].ToString();
                if (!dr["jvh_reference_date"].Equals(DBNull.Value))
                {
                    Dt_Date = (DateTime)dr["jvh_reference_date"];
                    str += "  " + Dt_Date.ToString("dd.MM.yy");
                }
                HeaderRow.jvh_reference = str;
                str = dr["jvh_org_invno"].ToString();
                if (!dr["jvh_org_invdt"].Equals(DBNull.Value))
                {
                    Dt_Date = (DateTime)dr["jvh_org_invdt"];
                    str += "  " + Dt_Date.ToString("dd.MM.yy");
                }
                HeaderRow.jvh_org_invno = str;
                HeaderRow.jvh_acc_id = dr["jvh_acc_id"].ToString();
                ReadCCDetails(dr["jvh_cc_category"].ToString(), dr["jvh_cc_id"].ToString());

                if (report_format == "FC")
                {
                    if (menu_admin == "Y")
                        ReadBankDetails("FC");
                }
                else
                {
                    if (dr["jvh_banktype"].ToString() == "SE")
                        ReadBankDetails("SEA EXPORT");
                    else if (dr["jvh_banktype"].ToString() == "SI")
                        ReadBankDetails("SEA IMPORT");
                    else if (dr["jvh_banktype"].ToString() == "AE")
                        ReadBankDetails("AIR EXPORT");
                    else if (dr["jvh_banktype"].ToString() == "AI")
                        ReadBankDetails("AIR IMPORT");
                    else if (dr["jvh_banktype"].ToString() == "BR")
                        ReadBankDetails("BROKERAGE");
                }

                break;
            }
        }
        private void ReadInvoiceFooter()
        {
            sql = " select text_value,ctr from captions where comp_code ='" + company_code + "' and branch_code is null";
            sql += " union all";
            sql += " select text_value,ctr from captions where comp_code ='" + company_code + "' and branch_code ='" + branch_code + "'";
            sql += " order by ctr ";

            Con_Oracle = new DBConnection();
            Dt_InvFooter = new DataTable();
            Dt_InvFooter = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
        }
        private void ReadCCDetails(string CC_Category, string CC_ID)
        {

            if (CC_Category == "NA")
                return;
            DateTime Dt_Date;
            Con_Oracle = new DBConnection();

            sql = " select nvl(hbl.hbl_bl_no,hbl.hbl_fcr_no) as hbl_bl_no,hbl.hbl_date,hbl.hbl_no,hbl.hbl_exp_id,shpr.cust_name as hbl_exp_name ";
            sql += "  ,mbl.hbl_bl_no as mbl_no,mbl.hbl_date as mbl_date";
            sql += "  ,vsl.param_name as mbl_vessel_name,mbl.hbl_vessel_no as mbl_vessel_voyage,mbl.hbl_pol_etd";
            sql += "  ,pol.param_name as hbl_pol_name,pod.param_name as hbl_pod_name,pofd.param_name as hbl_pofd_name,hbl.hbl_terms as hbl_freight_status";
            sql += "  ,hbl.hbl_pkg ,pkgunit.param_code as hbl_pkg_unit,hbl.hbl_cbm";
            sql += "  ,hbl.hbl_imp_id,cnge.cust_name as hbl_consignee_name,cha.cust_name as hbl_cha_name";
            sql += "  ,hbl.hbl_book_cntr as hbl_containers";
            sql += "  ,mbl.hbl_igmno as mbl_igmno, mbl.hbl_igmdate as mbl_igmdate,hbl.hbl_beno,hbl.hbl_bedate ";
            sql += "  ,hbl.hbl_ntwt,hbl.hbl_grwt,hbl.hbl_chwt ,wtunit.param_code as hbl_wt_unit,hbl.hbl_prefix as hbl_genjob_no,hbl.hbl_igmno as hbl_genigm_no ";
            sql += "  ,hbl.hbl_commodity,hbl.hbl_carton_nos,hbl.hbl_invoice_nos,type.param_code as hbl_genjobtype_code,nvl(hbl.hbl_invnos_prncount,0) as hbl_invnos_prncount ";
            sql += "  from hblm hbl";
            sql += "  left join hblm mbl on hbl.hbl_mbl_id= mbl.hbl_pkid";
            sql += "  left join param vsl on mbl.hbl_vessel_id = vsl.param_pkid";
            sql += "  left join param pol on hbl.hbl_pol_id = pol.param_pkid";
            sql += "  left join param pod on hbl.hbl_pod_id = pod.param_pkid";
            sql += "  left join param pofd on hbl.hbl_pofd_id = pofd.param_pkid";
            sql += "  left join param pkgunit on hbl.hbl_pkg_unit_id = pkgunit.param_pkid ";
            sql += "  left join customerm cnge on hbl.hbl_imp_id = cnge.cust_pkid";
            sql += "  left join customerm shpr on hbl.hbl_exp_id = shpr.cust_pkid";
            sql += "  left join customerm cha on hbl.hbl_cha_id = cha.cust_pkid";
            sql += "  left join param wtunit on hbl.hbl_grwt_unit_id = wtunit.param_pkid ";
            sql += "  left join param type on hbl.hbl_genjob_type_id = type.param_pkid ";
            sql += "  where hbl.hbl_pkid ='" + CC_ID + "'";
            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow dr in Dt_Rec.Rows)
            {
                HeaderRow.hbl_bl_no = dr["hbl_bl_no"].ToString();
                HeaderRow.hbl_no = dr["hbl_no"].ToString();
                HeaderRow.hbl_date = "";
                if (!dr["hbl_date"].Equals(DBNull.Value))
                {
                    Dt_Date = (DateTime)dr["hbl_date"];
                    HeaderRow.hbl_date = Dt_Date.ToString("dd.MM.yy");
                }

                HeaderRow.mbl_no = dr["mbl_no"].ToString();
                HeaderRow.mbl_date = "";
                if (!dr["mbl_date"].Equals(DBNull.Value))
                {
                    Dt_Date = (DateTime)dr["mbl_date"];
                    HeaderRow.mbl_date = Dt_Date.ToString("dd.MM.yy");
                }
                HeaderRow.mbl_vessel_name = dr["mbl_vessel_name"].ToString();
                HeaderRow.mbl_vessel_voyage = dr["mbl_vessel_voyage"].ToString();
                HeaderRow.mbl_pol_etd = "";
                if (!dr["hbl_pol_etd"].Equals(DBNull.Value))
                {
                    Dt_Date = (DateTime)dr["hbl_pol_etd"];
                    HeaderRow.mbl_pol_etd = Dt_Date.ToString("dd.MM.yy");
                }
                HeaderRow.mbl_cha_name = dr["hbl_cha_name"].ToString();
                HeaderRow.hbl_containers = dr["hbl_containers"].ToString();
                HeaderRow.mbl_igmno = dr["mbl_igmno"].ToString();
                HeaderRow.mbl_igmdate = "";
                if (!dr["mbl_igmdate"].Equals(DBNull.Value))
                {
                    Dt_Date = (DateTime)dr["mbl_igmdate"];
                    HeaderRow.mbl_igmdate = Dt_Date.ToString("dd.MM.yy");
                }

                HeaderRow.hbl_pol_name = dr["hbl_pol_name"].ToString();
                HeaderRow.hbl_pod_name = dr["hbl_pod_name"].ToString();
                HeaderRow.hbl_pofd_name = dr["hbl_pofd_name"].ToString();
                HeaderRow.hbl_freight_status = dr["hbl_freight_status"].ToString();
                HeaderRow.hbl_pkg = Lib.Conv2Decimal(dr["hbl_pkg"].ToString());
                HeaderRow.hbl_pkg_unit = dr["hbl_pkg_unit"].ToString();
                HeaderRow.hbl_cbm = Lib.Conv2Decimal(dr["hbl_cbm"].ToString());
                HeaderRow.hbl_consignee_name = dr["hbl_consignee_name"].ToString();
                HeaderRow.hbl_imp_id = dr["hbl_imp_id"].ToString();

                HeaderRow.hbl_beno = dr["hbl_beno"].ToString();
                HeaderRow.hbl_bedate = "";
                if (!dr["hbl_bedate"].Equals(DBNull.Value))
                {
                    Dt_Date = (DateTime)dr["hbl_bedate"];
                    HeaderRow.hbl_bedate = Dt_Date.ToString("dd.MM.yy");
                }
                HeaderRow.hbl_ntwt = Lib.Conv2Decimal(dr["hbl_ntwt"].ToString());
                HeaderRow.hbl_grwt = Lib.Conv2Decimal(dr["hbl_grwt"].ToString());
                HeaderRow.hbl_chwt = Lib.Conv2Decimal(dr["hbl_chwt"].ToString());
                HeaderRow.hbl_wt_unit = dr["hbl_wt_unit"].ToString();
                HeaderRow.hbl_genjob_no = dr["hbl_genjob_no"].ToString();
                HeaderRow.hbl_genigm_no = dr["hbl_genigm_no"].ToString();
                HeaderRow.hbl_exp_id = dr["hbl_exp_id"].ToString();
                HeaderRow.hbl_exp_name = dr["hbl_exp_name"].ToString();
                HeaderRow.hbl_carton_nos = dr["hbl_carton_nos"].ToString();
                HeaderRow.job_commodity = dr["hbl_commodity"].ToString();
                HeaderRow.hbl_invoice_nos = dr["hbl_invoice_nos"].ToString();
                HeaderRow.hbl_genjobtype_code = dr["hbl_genjobtype_code"].ToString();
                HeaderRow.hbl_invnos_prncount = Lib.Conv2Integer(dr["hbl_invnos_prncount"].ToString());
            }
            if (HeaderRow.job_commodity.Trim() == "")
                HeaderRow.job_commodity = GetCommodity(CC_ID);
            HeaderRow.job_invnos = GetInvoiceNos(CC_ID);
            HeaderRow.job_comm_invnos = GetCommInvoiceNos(CC_ID);
            HeaderRow.job_sbnos = GetsBillNos(CC_ID);

            if (HeaderRow.mbl_no.Trim() == "") //to display mblno entered in job for clearance
                GetEdiMblno(CC_ID);

            if (CC_Category.StartsWith("M"))
            {
                sql = "";
                sql += "  select mbl.hbl_bl_no as mbl_no,mbl.hbl_date as mbl_date, mbl.hbl_no as mbl_bkno ";
                sql += "  ,vsl.param_name as mbl_vessel_name,mbl.hbl_vessel_no as mbl_vessel_voyage,mbl.hbl_pol_etd";
                sql += "  ,cha.cust_name as mbl_cha_name";
                sql += "  ,mbl.hbl_igmno as mbl_igmno, mbl.hbl_igmdate as mbl_igmdate ";
                sql += "  from hblm mbl";
                sql += "  left join param vsl on mbl.hbl_vessel_id = vsl.param_pkid";
                sql += "  left join customerm cha on mbl.hbl_cha_id = cha.cust_pkid";
                sql += "  where mbl.hbl_pkid ='" + CC_ID + "'";
                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow dr in Dt_Rec.Rows)
                {
                    HeaderRow.mbl_bkno = dr["mbl_bkno"].ToString();
                    HeaderRow.mbl_no = dr["mbl_no"].ToString();
                    HeaderRow.mbl_date = "";
                    if (!dr["mbl_date"].Equals(DBNull.Value))
                    {
                        Dt_Date = (DateTime)dr["mbl_date"];
                        HeaderRow.mbl_date = Dt_Date.ToString("dd.MM.yy");
                    }
                    HeaderRow.mbl_vessel_name = dr["mbl_vessel_name"].ToString();
                    HeaderRow.mbl_vessel_voyage = dr["mbl_vessel_voyage"].ToString();
                    HeaderRow.mbl_pol_etd = "";
                    if (!dr["hbl_pol_etd"].Equals(DBNull.Value))
                    {
                        Dt_Date = (DateTime)dr["hbl_pol_etd"];
                        HeaderRow.mbl_pol_etd = Dt_Date.ToString("dd.MM.yy");
                    }
                    HeaderRow.mbl_cha_name = dr["mbl_cha_name"].ToString();
                    HeaderRow.mbl_igmno = dr["mbl_igmno"].ToString();
                    HeaderRow.mbl_igmdate = "";
                    if (!dr["mbl_igmdate"].Equals(DBNull.Value))
                    {
                        Dt_Date = (DateTime)dr["mbl_igmdate"];
                        HeaderRow.mbl_igmdate = Dt_Date.ToString("dd.MM.yy");
                    }
                }

                //to print mbl invoice footer datatable pass to inv print class
                sql = " select jvh_pkid,";
                sql += " max(hbl_no) as si, max(hbl_type) as hbl_type, max(hbl_bl_no) as hbl_no,max(hbl_terms) as terms, ";
                sql += " max(exp.cust_name) as exp_name,";
                sql += " max(imp.cust_name) as imp_name,";
                sql += " jvh_docno,jvh_vrno, jvh_date,  ";
                sql += " sum(case when acc_main_code='1105' then jv_credit else 0 end ) as inv_frt, ";
                sql += " sum(case when acc_main_code='1106' then jv_credit else 0 end ) as inv_thc, ";
                sql += " sum(case when acc_main_code='1107' then jv_credit else 0 end ) as inv_tpt, ";
                sql += " sum(jv_credit) as inv_total ";
                sql += " from hblm  ";
                sql += " left join ledgerh a on hbl_pkid = jvh_cc_id ";
                sql += " inner join ledgert b on a.jvh_pkid = jv_parent_id";
                sql += " inner join acctm acc on b.jv_acc_id =acc.acc_pkid and jv_row_type not in ('HEADER')";
                sql += " left join customerm exp on hbl_exp_id = exp.cust_pkid";
                sql += " left join customerm imp on hbl_imp_id = imp.cust_pkid";
                sql += " where hbl_mbl_id ='" + CC_ID + "'  ";
                sql += " group by jvh_pkid, jvh_docno,jvh_vrno, jvh_date";
                sql += " order by jvh_vrno";

                Dt_MblInv = new DataTable();
                Dt_MblInv = Con_Oracle.ExecuteQuery(sql);

                sql = " select jvh_docno,jvh_vrno,jvh_date,jv_debit, nvl(sum(xref_amt),0) as credit , ";
                sql += " max(xref_cr_jv_date) as cr_date,";
                sql += " jv_debit - nvl(sum(xref_amt),0) as balance ,";
                sql += " max(cust_crdays) as cr_days, ";
                sql += " trunc(sysdate - jvh_date,0)  as os_days    ";
                sql += " from ledgerh h  ";
                sql += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id) ";
                sql += " inner join  customerm a on (L.jv_acc_id = a.cust_pkid ) ";
                sql += " left  join ledgerxref X on (L.jv_pkid=X.xref_dr_jv_id ) ";
                sql += " left  join param s on ( jv_acc_id = param_pkid) ";
                sql += " where  jvh_cc_id in( select hbl_pkid from hblm where hbl_mbl_id = '" + CC_ID + "') and L.jv_debit > 0      ";
                sql += " group by  jv_pkid,jvh_docno,jvh_vrno, jvh_date, jv_debit   ";
                sql += " order by jvh_vrno";
                Dt_MblOs = new DataTable();
                Dt_MblOs = Con_Oracle.ExecuteQuery(sql);
            }

            if (CC_Category == "GENERAL JOB")
            {
                sql = " select gj_seal_no,gj_loaded_on,gj_unloaded_on,gj_cfs,gj_from,gj_to1,gj_to2,gj_shipper_inv_no,gj_cargo ";
                sql += " ,gj_consignee_name,gj_vessel, gj_mbl_no,gj_hbl_no,gj_frt_status,gj_cha_name,gj_sb_no,gj_pack_list_no,gj_lr_no  ";
                sql += "  from genjobm ";
                sql += "  where gj_parent_id ='" + CC_ID + "'";
                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow dr in Dt_Rec.Rows)
                {
                    HeaderRow.gj_seal_no = dr["gj_seal_no"].ToString();
                    HeaderRow.gj_loaded_on = Lib.DatetoStringDisplayformat(dr["gj_loaded_on"]);
                    HeaderRow.gj_unloaded_on = Lib.DatetoStringDisplayformat(dr["gj_unloaded_on"]);
                    HeaderRow.gj_cfs = dr["gj_cfs"].ToString();
                    HeaderRow.gj_from = dr["gj_from"].ToString();
                    HeaderRow.gj_to1 = dr["gj_to1"].ToString();
                    HeaderRow.gj_to2 = dr["gj_to2"].ToString();
                    HeaderRow.gj_shipper_inv_no = dr["gj_shipper_inv_no"].ToString();
                    HeaderRow.gj_vessel = dr["gj_vessel"].ToString();
                    HeaderRow.gj_mbl_no = dr["gj_mbl_no"].ToString();
                    HeaderRow.gj_hbl_no = dr["gj_hbl_no"].ToString();
                    HeaderRow.gj_frt_status = dr["gj_frt_status"].ToString();
                    HeaderRow.gj_cha_name = dr["gj_cha_name"].ToString();
                    HeaderRow.gj_sb_no = dr["gj_sb_no"].ToString();
                    HeaderRow.gj_consignee_name = dr["gj_consignee_name"].ToString();
                    HeaderRow.job_commodity = dr["gj_cargo"].ToString();
                    HeaderRow.gj_pack_list_no = dr["gj_pack_list_no"].ToString();
                    HeaderRow.gj_lr_no = dr["gj_lr_no"].ToString();
                    break;
                }
            }
            Con_Oracle.CloseConnection();
        }

        private string GetCommodity(string HBLID)
        {
            string sCommodity = "";
            sql = " select distinct  commodity.param_name as job_commodity_name from Jobm a  ";
            sql += "  left join param commodity on a.job_commodity_id = commodity.param_pkid  ";
            sql += "  where a.jobs_hbl_id ='" + HBLID + "'";
            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow dr in Dt_Rec.Rows)
            {
                if (sCommodity.Trim() != "")
                    sCommodity += ",";
                sCommodity += dr["job_commodity_name"].ToString();
            }
            return sCommodity;
        }

        private string GetEdiMblno(string HBLID)
        {
            sql= "select hbl_pkid from hblm where rec_company_code ='CPL' and rec_branch_code='BLRAF' and rec_category='AIR EXPORT' and nvl(length(hbl_mbl_id),0) <= 0 and hbl_pkid = '" + HBLID + "'";
            if (Con_Oracle.IsRowExists(sql))
            {
                sql= "select job_edi_mbl_no,job_edi_mbl_date from jobm a where nvl(length(job_edi_mbl_no),0)>0 and jobs_hbl_id ='" + HBLID + "'";
                DataTable Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                DateTime Dt_Date;
                foreach (DataRow dr in Dt_Rec.Rows)
                {
                    HeaderRow.mbl_no = dr["job_edi_mbl_no"].ToString();
                    HeaderRow.mbl_date = "";
                    if (!dr["job_edi_mbl_date"].Equals(DBNull.Value))
                    {
                        Dt_Date = (DateTime)dr["job_edi_mbl_date"];
                        HeaderRow.mbl_date = Dt_Date.ToString("dd.MM.yy");
                    }
                    break;
                }
            }

            return "";
        }

        private string GetsBillNos(string HBLID)
        {
            string sBillNos = ""; string SbDate = ""; bool differentSbDate = false;
            sql = " select distinct opr_Sbill_Date,opr_sbill_no from joboperationsm a";
            sql += "  inner join jobm b on a.opr_job_id = b.job_pkid ";
            sql += "  where  b.jobs_hbl_id ='" + HBLID + "'";
            sql += " order by opr_sbill_no";
            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            if (Dt_Rec.Rows.Count > 0)
            {
                DataTable DistinctSB = Dt_Rec.DefaultView.ToTable(true, "opr_Sbill_Date");
                if (DistinctSB.Rows.Count > 1)
                    differentSbDate = true;

                foreach (DataRow dr in Dt_Rec.Rows)
                {
                    if (!dr["opr_Sbill_Date"].Equals(DBNull.Value))
                        SbDate = ((DateTime)dr["opr_Sbill_Date"]).ToString("dd.MM.yy");


                    if (sBillNos.Trim() != "")
                        sBillNos += ",";
                    sBillNos += dr["opr_sbill_no"].ToString();
                    if (differentSbDate)
                        sBillNos += "/" + SbDate;
                }
                if (!differentSbDate && sBillNos.Trim() != "" && SbDate.Trim() != "")
                    sBillNos += "/" + SbDate;
            }
            return sBillNos;
        }
        private string GetInvoiceNos(string HBLID)
        {
            string sInvNos = ""; string InvDate = ""; bool differentInvDate = false;
            sql = " select distinct jexp_invoice_no,jexp_invoice_date from jobexpm a";
            sql += "  inner join jobm b on a.jexp_job_id = b.job_pkid ";
            sql += "  where  b.jobs_hbl_id ='" + HBLID + "'";
            sql += "  order by jexp_invoice_no";
            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            if (Dt_Rec.Rows.Count > 0)
            {
                DataTable DistinctInv = Dt_Rec.DefaultView.ToTable(true, "jexp_invoice_date");
                if (DistinctInv.Rows.Count > 1)
                    differentInvDate = true;

                foreach (DataRow dr in Dt_Rec.Rows)
                {
                    if (!dr["jexp_invoice_date"].Equals(DBNull.Value))
                        InvDate = ((DateTime)dr["jexp_invoice_date"]).ToString("dd.MM.yy");

                    if (sInvNos.Trim() != "")
                        sInvNos += ",";
                    sInvNos += dr["jexp_invoice_no"].ToString();
                    if (differentInvDate)
                        sInvNos += "/" + InvDate;
                }
                if (!differentInvDate && sInvNos.Trim() != "" && InvDate.Trim() != "")
                    sInvNos += "/" + InvDate;
            }
            return sInvNos;
        }

        private string GetCommInvoiceNos(string HBLID)
        {
            string sInvNos = ""; string InvDate = ""; bool differentInvDate = false;
            sql = " select distinct jexp_comm_invoice_no from jobexpm a";
            sql += "  inner join jobm b on a.jexp_job_id = b.job_pkid ";
            sql += "  where  b.jobs_hbl_id ='" + HBLID + "'";
            sql += "  order by jexp_comm_invoice_no";
            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow dr in Dt_Rec.Rows)
            {
                if (sInvNos.Trim() != "")
                    sInvNos += ",";
                sInvNos += dr["jexp_comm_invoice_no"].ToString();
            }
            return sInvNos;
        }

        private void SetPrintFormat()
        {
            ParticularColWidth = 30;
            IsIGST = false; IsCSGST = false;
            if (Lib.Conv2Decimal(HeaderRow.jvh_igst_amt.ToString()) > 0 || HeaderRow.jvh_type=="IN-ES")
                IsIGST = true;
            decimal scAmt = Lib.Conv2Decimal(HeaderRow.jvh_cgst_amt.ToString()) + Lib.Conv2Decimal(HeaderRow.jvh_sgst_amt.ToString());
            if (scAmt > 0)
                IsCSGST = true;

            header_format = HeaderRow.jvh_cc_category;
            if (report_format == "SUMMARY" || report_format == "FC")
            {
                ParticularColWidth = 640;
                detail_format = "SUMMARY";
                if (IsIGST)
                {
                    ParticularColWidth = 455;
                    detail_format = "SUMMARY-I-GST";
                }
                if (IsCSGST)
                {
                    ParticularColWidth = 370;
                    detail_format = "SUMMARY-CS-GST";
                }
            }
            else if (report_format == "DETAIL")
            {
                ParticularColWidth = 475;
                detail_format = "DETAIL";
                if (IsIGST)
                {
                    ParticularColWidth = 265;
                    detail_format = "DETAIL-I-GST";
                }
                if (IsCSGST)
                {
                    ParticularColWidth = 175;
                    detail_format = "DETAIL-CS-GST";
                }
            }
        }

        private void ReadInvoiceDetails()
        {
            Con_Oracle = new DBConnection();
            DetailList = new List<InvoiceReport>();
            InvoiceReport mRow = new InvoiceReport();

            sql  = " select sac.param_code as jv_sac_code, jv_acc_name,jv_qty,jv_rate,jv_ftotal";
            sql += " ,jv_exrate,curr.param_code as jv_curr_code,jv_drcr";
            sql += " ,jv_total,jv_cgst_rate,jv_cgst_amt,jv_sgst_rate,jv_sgst_amt,jv_igst_rate,jv_igst_amt,jv_net_total ";
            sql += " ,jv_total_fc,jv_cgst_famt,jv_sgst_famt,jv_igst_famt,jv_net_ftotal";
            sql += " from ledgert a";
            sql += " left join param sac on a.jv_sac_id = sac.param_pkid";
            sql += " left join acctm acc on a.jv_acc_id = acc.acc_pkid";
            sql += " left join param curr on a.jv_curr_id = curr.param_pkid";
            sql += " where jv_parent_id ='" + pkid + "' and nvl(jv_row_type,'JV') not in('HEADER','GST') ";
            sql += " order by jv_ctr";

            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            int iCtr = 0; jvNetTot = 0;
            string[] AccNameArry;
            foreach (DataRow dr in Dt_Rec.Rows)
            {
                iCtr++;
                AccNameArry = Lib.ConvertString2Lines(dr["jv_acc_name"].ToString(), ParticularColWidth, "WORD", "Calibri", 9);
                mRow = new InvoiceReport();
                mRow.jv_ctr = iCtr;
                mRow.jv_sac_code = dr["jv_sac_code"].ToString();
                mRow.jv_acc_name = AccNameArry.Length > 0 ? AccNameArry[0].Trim() : "";
                mRow.jv_total = Lib.Conv2Decimal(dr["jv_total"].ToString());
                mRow.jv_cgst_rate = Lib.Conv2Decimal(dr["jv_cgst_rate"].ToString());
                mRow.jv_cgst_amt = Lib.Conv2Decimal(dr["jv_cgst_amt"].ToString());
                mRow.jv_sgst_rate = Lib.Conv2Decimal(dr["jv_sgst_rate"].ToString());
                mRow.jv_sgst_amt = Lib.Conv2Decimal(dr["jv_sgst_amt"].ToString());
                mRow.jv_igst_rate = Lib.Conv2Decimal(dr["jv_igst_rate"].ToString());
                mRow.jv_igst_amt = Lib.Conv2Decimal(dr["jv_igst_amt"].ToString());
                mRow.jv_net_total = Lib.Conv2Decimal(dr["jv_net_total"].ToString());
                mRow.jv_qty = Lib.Conv2Decimal(dr["jv_qty"].ToString());
                mRow.jv_rate = Lib.Conv2Decimal(dr["jv_rate"].ToString());
                mRow.jv_ftotal = Lib.Conv2Decimal(dr["jv_ftotal"].ToString());
                mRow.jv_exrate = Lib.Conv2Decimal(dr["jv_exrate"].ToString());
                mRow.jv_total_fc = Lib.Conv2Decimal(dr["jv_total_fc"].ToString());
                mRow.jv_cgst_famt = Lib.Conv2Decimal(dr["jv_cgst_famt"].ToString());
                mRow.jv_sgst_famt = Lib.Conv2Decimal(dr["jv_sgst_famt"].ToString());
                mRow.jv_igst_famt = Lib.Conv2Decimal(dr["jv_igst_famt"].ToString());
                mRow.jv_net_ftotal = Lib.Conv2Decimal(dr["jv_net_ftotal"].ToString());
                mRow.jv_curr_code = dr["jv_curr_code"].ToString();
                DetailList.Add(mRow);

                for (int i = 1; i < AccNameArry.Length; i++)//Adding Wrapped lines
                {
                    mRow = new InvoiceReport();
                    mRow.jv_ctr = 0;
                    mRow.jv_sac_code = "";
                    mRow.jv_acc_name = AccNameArry[i].Trim();
                    mRow.jv_total = 0;
                    mRow.jv_cgst_rate = 0;
                    mRow.jv_cgst_amt = 0;
                    mRow.jv_sgst_rate = 0;
                    mRow.jv_sgst_amt = 0;
                    mRow.jv_igst_rate = 0;
                    mRow.jv_igst_amt = 0;
                    mRow.jv_net_total = 0;
                    mRow.jv_qty = 0;
                    mRow.jv_rate = 0;
                    mRow.jv_ftotal = 0;
                    mRow.jv_exrate = 0;
                    mRow.jv_total_fc = 0;
                    mRow.jv_cgst_famt = 0;
                    mRow.jv_sgst_famt = 0;
                    mRow.jv_igst_famt = 0;
                    mRow.jv_net_ftotal = 0;
                    mRow.jv_curr_code = "";
                    DetailList.Add(mRow);
                }

                if ( dr["jv_drcr"].ToString () == "DR" )
                    jvNetTot += Lib.Conv2Decimal(dr["jv_net_total"].ToString());
                else
                    jvNetTot -= Lib.Conv2Decimal(dr["jv_net_total"].ToString());
            }
            jvNetTot = Math.Abs(jvNetTot);
            Con_Oracle.CloseConnection();
        }

        private void ReadBankDetails(string BankType)
        {
            DataTable Dt_Rec = new DataTable();
            BankInfo = new Dictionary<string, string>();
            sql = "select caption, name from settings where parentid ='" + branch_code + "' and tabletype ='" + BankType + "'";
            Con_Oracle = new DBConnection();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow dr in Dt_Rec.Rows)
            {
                if (!BankInfo.ContainsKey(dr["caption"].ToString()))
                {
                    if (dr["name"].ToString().Trim() != "")
                        BankInfo.Add(dr["caption"].ToString(), dr["name"].ToString());
                }
            }
        }
    }
}


