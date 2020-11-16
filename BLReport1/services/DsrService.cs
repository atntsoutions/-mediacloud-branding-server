using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

using XL.XSheet;

namespace BLReport1
{
    public class DsrService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<DsrReport> mList = new List<DsrReport>();
        DsrReport mrow;
        int iRow = 0;
        int iCol = 0;
        string type = "";

        string report_folder = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string searchtype = "";
        string searchstring = "";
        string rec_category = "";
        string searchexpid = "";
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string ErrorMessage = "";
        string job_type = "";
        string shipper_id = "";
        string consignee_id = "";
        string agent_id = "";
        string carrier_id = "";
        string pol_id = "";
        string pod_id = "";
        string format_type = "";
        string job_liner_agent = "";
        string job_agent = "";
        string job_liner = "";
        Boolean all = false;

       
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<DsrReport>();
            ErrorMessage = "";
            try
            {

                type = SearchData["type"].ToString();
                rec_category = SearchData["rec_category"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
                type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();
                job_type = SearchData["job_type"].ToString();

                all = (Boolean)SearchData["all"];

                if (SearchData.ContainsKey("shipper_id"))
                    shipper_id = SearchData["shipper_id"].ToString();

                if (SearchData.ContainsKey("consignee_id"))
                    consignee_id = SearchData["consignee_id"].ToString();

                if (SearchData.ContainsKey("agent_id"))
                    agent_id = SearchData["agent_id"].ToString();

                if (SearchData.ContainsKey("carrier_id"))
                    carrier_id = SearchData["carrier_id"].ToString();

                if (SearchData.ContainsKey("pol_id"))
                    pol_id = SearchData["pol_id"].ToString();

                if (SearchData.ContainsKey("pod_id"))
                    pod_id = SearchData["pod_id"].ToString();
                if (SearchData.ContainsKey("format_type"))
                    format_type = SearchData["format_type"].ToString();

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);

                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");


                //if (type == "SCREEN" && from_date != "NULL" && to_date != "NULL")
                //{
                //    DateTime dt_frm = DateTime.Parse(from_date);
                //    DateTime dt_to = DateTime.Parse(to_date);
                //    int days = (dt_to - dt_frm).Days;

                //    if (days > 31)
                //        Lib.AddError(ref ErrorMessage, " | Only one month data range can be used,use excel to download");
                //}
                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }


                if (rec_category == "SEA EXPORT")
                {

                    sql = " select job_pkid,job_date,job_docno,job_type,job_invoice_nos,job_prefix,a.rec_branch_code as branch,";

                    sql += " a.job_nature,mbl.hbl_book_no as booking_no,mbl.hbl_date as booking_date,mbl.hbl_prealert_date as prealert_send_on,hbl.hbl_ar_invnos as our_invoice,hbl.hbl_ar_invamt,hbl.hbl_ar_gstamt, ";

                    sql += "  shpr.cust_name as shipper_name ,cons.cust_name as consignee_name,agent.cust_name as job_agent_name,";
                    sql += "  hbl.hbl_bl_no as hbl_bl_no,hbl.hbl_no,mbl.hbl_bl_no as mbl_bl_no,mbl.hbl_no as mbl_no,a.job_cbm as job_cbm,a.job_remarks as job_remarks,";
                    sql += "  a.job_pkg as job_pkg,a.job_pcs as job_pcs,a.job_ntwt as job_ntwt,a.job_grwt as job_grwt,";
                    sql += "  pol.param_name as pol_name,pod.param_name as pod_name,liner.param_name as liner_name,";
                    sql += "  nvl(sman2.param_name, sman.param_name) as sman_name,";
                    sql += "  comdty.param_name as commodity,cons.cust_nomination as nomination,a.job_terms as job_terms,a.job_status as job_status,";
                    sql += "  opr.opr_sbill_no as sbill_no, opr.opr_sbill_date as sbill_date, opr.opr_cargo_received_on as cargo_received_on,";
                    sql += "  fwd.cust_name as forwarder_name,vsl.param_name as mbl_vessel_name,mbl.hbl_vessel_no as mbl_vessel_no,";
                    sql += "  opr.opr_stuffed_at as stuffed_at,opr.opr_stuffed_on as stuffed_on,hbl.hbl_book_cntr as book_cntr,";
                    sql += "  mbl.hbl_pol_etd as sob,mbl.hbl_pofd_eta as destination_eta,jopr.OPR_EP_REC_DATE as ep_received_on,";
                    sql += "  a.job_cntr_type as cntr_type,a.job_cntr_teu as cntr_teu,scheme.param_name as scheme,";
                    sql += "  hbl.hbl_released_date as hbl_released_date,mbl.hbl_released_date as mbl_released_date,cha.cust_name as job_cha_name, ";
                    sql += "  a.job_cntr,jobliner.param_name as job_liner_name,jobagent.cust_name as job_agent,vsl2.param_name as mbl_vessel2 ";
                    sql += "  from jobm a ";

                    sql += "  left join customerm shpr on a.job_exp_id = shpr.cust_pkid";
                    sql += "  left join custdet  cd on a.rec_branch_code = cd.det_branch_code and a.job_exp_id = cd.det_cust_id ";

                    sql += " left join param sman on shpr.cust_sman_id = sman.param_pkid";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += "  left join customerm cons on a.job_imp_id = cons.cust_pkid  ";
                    sql += "  left join hblm hbl on a.jobs_hbl_id  = hbl.hbl_pkid";
                    sql += "  left join customerm agent on hbl.hbl_agent_id = agent.cust_pkid";
                    sql += "  left join param liner on hbl.hbl_carrier_id = liner.param_pkid";
                    sql += "  left join param pol on a.job_pol_id = pol.param_pkid ";
                    sql += "  left join param comdty on a.job_commodity_id = comdty.param_pkid";
                    sql += "  left join param pod on a.job_pod_id = pod.param_pkid ";
                    sql += "  left join joboperationsm opr on a.job_pkid = opr.opr_job_id   ";
                    sql += "  left join customerm fwd on a.job_forwarder_id = fwd.cust_pkid ";
                    sql += "  left join hblm mbl on  hbl.hbl_mbl_id  = mbl.hbl_pkid";
                    sql += "  left join param vsl on mbl.hbl_vessel_id = vsl.param_pkid";
                    sql += "  left join joboperationsm jopr on a.job_pkid = jopr.opr_job_id ";
                    sql += "  left join customerm cha on a.job_cha_id = cha.cust_pkid ";
                    sql += "  left join param scheme on a.job_billtype_id = scheme.param_pkid ";
                    sql += "  left join param jobliner on a.job_carrier_id = jobliner.param_pkid";
                    sql += "  left join customerm jobagent on a.job_agent_id = jobagent.cust_pkid";
                    sql += "  left join param vsl2 on mbl.hbl_vessel2_id = vsl2.param_pkid";
                    sql += "  where a.rec_category = 'SEA EXPORT' ";
                    if (!all)
                    {
                        sql += " and a.rec_branch_code = '{BRCODE}' ";
                    }
                    
                     sql += "  and a.job_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    
                    if (job_type != "ALL")
                        sql += " and job_type ='{JOBTYPE}'";
                    if (shipper_id.Length > 0)
                        sql += " and a.job_exp_id = '" + shipper_id + "' ";
                    if (consignee_id.Length > 0)
                        sql += " and a.job_imp_id = '" + consignee_id + "' ";
                    if (agent_id.Length > 0)
                        sql += " and ( hbl.hbl_agent_id = '" + agent_id + "' or a.job_agent_id = '" + agent_id + "' ) ";
                    if (carrier_id.Length > 0)
                        sql += " and ( hbl.hbl_carrier_id = '" + carrier_id + "' or a.job_carrier_id = '" + carrier_id + "' )";
                    if (pol_id.Length > 0)
                        sql += " and a.job_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and a.job_pod_id = '" + pod_id + "'";
                  
                    if (type_date == "DATE")
                        sql += " order by a.rec_branch_code,job_date";
                    else if (type_date == "JOB-TYPE")
                        sql += " order by a.rec_branch_code,job_type";
                    else
                        sql += " order by a.rec_branch_code,job_docno";


                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{JOBTYPE}", job_type);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                   
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        job_agent = "";
                        job_liner = "";
                        job_liner_agent = "";
                        mrow = new DsrReport();
                        mrow.job_pkid = Dr["job_pkid"].ToString();
                        mrow.job_date = Lib.DatetoStringDisplayformat(Dr["job_date"]);
                        mrow.job_docno = Dr["job_docno"].ToString();
                        mrow.job_invoice_nos = Dr["job_invoice_nos"].ToString();
                        mrow.job_prefix = Dr["job_prefix"].ToString();
                        mrow.job_shipper = Dr["shipper_name"].ToString();
                        mrow.job_consignee = Dr["consignee_name"].ToString();
                        mrow.hbl_bl_no = Dr["hbl_bl_no"].ToString();
                        mrow.mbl_bl_no = Dr["mbl_bl_no"].ToString();
                        mrow.job_cbm = Lib.Conv2Decimal(Dr["job_cbm"].ToString());
                        mrow.job_pkg = Lib.Conv2Decimal(Dr["job_pkg"].ToString());
                        mrow.job_pcs = Lib.Conv2Decimal(Dr["job_pcs"].ToString());
                        mrow.job_ntwt = Lib.Conv2Decimal(Dr["job_ntwt"].ToString());
                        mrow.job_grwt = Lib.Conv2Decimal(Dr["job_grwt"].ToString());
                        mrow.job_pol = Dr["pol_name"].ToString();
                        mrow.job_pod = Dr["pod_name"].ToString();
                        mrow.opr_sbill_no = Dr["sbill_no"].ToString();
                        mrow.opr_sbill_date = Lib.DatetoStringDisplayformat(Dr["sbill_date"]);
                        mrow.opr_cargo_received_on = Lib.DatetoStringDisplayformat(Dr["cargo_received_on"]);
                        mrow.forwarder_name = Dr["forwarder_name"].ToString();
                        mrow.mbl_vessel_name = Dr["mbl_vessel_name"].ToString();
                        mrow.mbl_vessel_no = Dr["mbl_vessel_no"].ToString();
                        mrow.opr_stuffed_at = Dr["stuffed_at"].ToString();
                        mrow.opr_stuffed_on = Lib.DatetoStringDisplayformat(Dr["stuffed_on"]);
                        mrow.hbl_book_cntr = Dr["book_cntr"].ToString();
                        mrow.mbl_pol_etd = Lib.DatetoStringDisplayformat(Dr["sob"]);
                        mrow.mbl_pofd_eta = Lib.DatetoStringDisplayformat(Dr["destination_eta"]);
                        mrow.job_type = Dr["job_type"].ToString();
                        mrow.liner_name = Dr["liner_name"].ToString();
                        mrow.job_agent_name = Dr["job_agent_name"].ToString();
                        mrow.job_remarks = Dr["job_remarks"].ToString();
                        mrow.job_nomination = Dr["nomination"].ToString();
                        mrow.job_commodity = Dr["commodity"].ToString();
                        mrow.job_terms = Dr["job_terms"].ToString();
                        mrow.job_status = Dr["job_status"].ToString();
                        mrow.salesman = Dr["sman_name"].ToString();
                        mrow.branch = Dr["branch"].ToString();
                        mrow.opr_ep_rec_date = Lib.DatetoStringDisplayformat(Dr["ep_received_on"]);
                        mrow.mbl_book_no = Dr["booking_no"].ToString();
                        mrow.mbl_book_date = Lib.DatetoStringDisplayformat(Dr["booking_date"]);
                        mrow.mbl_prealert_date = Lib.DatetoStringDisplayformat(Dr["prealert_send_on"]);
                        mrow.hbl_ar_invnos = Dr["our_invoice"].ToString();
                        mrow.job_nature = Dr["job_nature"].ToString();
                        mrow.job_cntr_type = Dr["cntr_type"].ToString();
                        mrow.job_cntr_teu = Lib.Conv2Decimal(Dr["cntr_teu"].ToString());
                        mrow.job_billtype_id = Dr["scheme"].ToString();
                        mrow.hbl_released_date = Lib.DatetoStringDisplayformat(Dr["hbl_released_date"]);
                        mrow.mbl_released_date = Lib.DatetoStringDisplayformat(Dr["mbl_released_date"]);
                        mrow.job_cha_name = Dr["job_cha_name"].ToString();
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.hbl_ar_invamt = Lib.Conv2Decimal(Dr["hbl_ar_invamt"].ToString());
                        mrow.hbl_ar_gstamt = Lib.Conv2Decimal(Dr["hbl_ar_gstamt"].ToString());
                        mrow.job_cntr = Dr["job_cntr"].ToString();                     
                        job_agent = Dr["job_agent"].ToString();
                        job_liner = Dr["job_liner_name"].ToString();

                        if (job_agent != "")
                            job_liner_agent = job_agent;
                        if (job_liner_agent != "" && job_liner != "")
                            job_liner_agent += " / ";
                        job_liner_agent += job_liner;

                        mrow.job_liner_agent = job_liner_agent;

                        mrow.mbl_vessel2_name = Dr["mbl_vessel2"].ToString();
                        mList.Add(mrow);

                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                        {
                            if (format_type == "GENERAL")
                                PrintDsrReport();
                            if(format_type == "STATUS")
                                PrintStatusDsrReport();
                            if(format_type == "SHIPPER")
                                PrintShipperDsrReport();
                        }
                    }
                    Dt_List.Rows.Clear();
                }

                if (rec_category == "AIR EXPORT")
                {

                    sql = " select a.rec_created_date as created_on,a.rec_branch_code as branch, ";
                    sql += "  job_docno, job_date,job_type,job_prefix,shpr.cust_name as shipper_name ,cons.cust_name as consignee_name,";
                    sql += "  cha.cust_name as job_cha_name,agent.cust_name as job_agent_name,";
                    sql += "  pol.param_name as job_pol_name,pod.param_name as job_pod_name,pofd.param_name as job_pofd_name,";
                    sql += "  comdty.param_name as commodity,a.job_nomination as nomination,a.job_terms as job_terms,a.job_status as job_status,";
                    sql += "  nvl(sman2.param_name, sman.param_name) as salesman,";
                    sql += "  opr_sbill_no as sbno, opr_sbill_date as sb_date, opr_cargo_received_on as cargo_date, opr_cleared_date as cleard_date,";
                    sql += "  job_terms as freight_status,cons.cust_nomination as job_nomination ,job_chwt,job_grwt,job_ntwt,job_remarks,";
                    sql += "  hbl.hbl_bl_no as house,hbl.hbl_no as hbl_no, hbl.hbl_date as house_date,hbl.hbl_invoice_nos as invoiceno,";
                    sql += "  mbl.hbl_bl_no as mbl_bl_no,mbl.hbl_no as mbl_no,mbl.hbl_date as mbl_date,liner.param_name as liner_name,fwdr.cust_name as forwarder_name,";
                    sql += "  mbl.hbl_grwt as mbl_grwt,mbl.hbl_chwt as mbl_chwt,mbl.hbl_folder_no as folderno,  mbl.hbl_folder_sent_date as folder_sent_date,";
                    sql += "  job_pkg as job_total_cartons,job_invoice_nos,opr_drawback_slno,opr_drawback_date,opr_drawback_amt,hbl.hbl_ar_invnos as our_invoice,hbl.hbl_ar_invamt,hbl.hbl_ar_gstamt,";
                    sql += "  jobliner.param_name as job_liner_name,jobagent.cust_name as job_agent";
                    sql += "  from jobm a ";
                    
                    sql += "  left join customerm shpr on a.job_exp_id = shpr.cust_pkid";
                    sql += "  left join custdet  cd on a.rec_branch_code = cd.det_branch_code and a.job_exp_id = cd.det_cust_id ";

                    sql += "  left join param sman on shpr.cust_sman_id = sman.param_pkid";
                    sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += "  left join customerm cons on a.job_imp_id = cons.cust_pkid  ";
                    sql += "  left join customerm cha on a.job_cha_id = cha.cust_pkid  ";
                    sql += "  left join param pol on a.job_pol_id = pol.param_pkid ";
                    sql += "  left join param comdty on a.job_commodity_id = comdty.param_pkid";
                    sql += "  left join param pod on a.job_pod_id = pod.param_pkid ";
                    sql += "  left join param pofd on a.job_pofd_id = pofd.param_pkid ";
                    sql += "  left join joboperationsm opr on a.job_pkid = opr_job_id";
                   
                    sql += "  left join hblm hbl on a.jobs_hbl_id = hbl.hbl_pkid";
                    sql += "  left join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                    sql += "  left join customerm agent on hbl.hbl_agent_id = agent.cust_pkid";
                    sql += "  left join param liner on hbl.hbl_carrier_id = liner.param_pkid";
                    sql += "  left join customerm fwdr on a.job_forwarder_id = fwdr.cust_pkid";

                    sql += "  left join param jobliner on a.job_carrier_id = jobliner.param_pkid";
                    sql += "  left join customerm jobagent on a.job_agent_id = jobagent.cust_pkid";

                    sql += "  where  a.rec_category = 'AIR EXPORT' ";
                    if(!all)
                    {
                        sql += " and a.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += "  and a.job_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";


                    if (job_type != "ALL")
                        sql += " and job_type ='{JOBTYPE}'";

                    if (shipper_id.Length > 0)
                        sql += " and a.job_exp_id = '" + shipper_id + "' ";

                    if (consignee_id.Length > 0)
                        sql += " and a.job_imp_id = '" + consignee_id + "' ";
                    if (agent_id.Length > 0)
                        sql += " and ( hbl.hbl_agent_id = '" + agent_id + "' or a.job_agent_id = '" + agent_id + "' ) ";
                    if (carrier_id.Length > 0)
                        sql += " and ( hbl.hbl_carrier_id = '" + carrier_id + "' or a.job_carrier_id = '" + carrier_id + "' )";
                    if (pol_id.Length > 0)
                        sql += " and a.job_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and a.job_pod_id = '" + pod_id + "'";

                    if (type_date == "DATE")
                        sql += " order by a.rec_branch_code,job_date";
                    else
                        sql += " order by a.rec_branch_code,job_docno";


                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{JOBTYPE}", job_type);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();


                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        job_agent = "";
                        job_liner = "";
                        job_liner_agent = "";

                        mrow = new DsrReport();
                        mrow.job_date = Lib.DatetoStringDisplayformat(Dr["job_date"]);
                        mrow.job_docno = Dr["job_docno"].ToString();
                        mrow.job_type = Dr["job_type"].ToString();
                        //  mrow.job_invoice_nos = Dr["job_invoice_nos"].ToString();
                        mrow.job_prefix = Dr["job_prefix"].ToString();
                        mrow.job_shipper = Dr["shipper_name"].ToString();
                        mrow.job_consignee = Dr["consignee_name"].ToString();

                        mrow.job_cha_name = Dr["job_cha_name"].ToString();
                        mrow.job_agent_name = Dr["job_agent_name"].ToString();
                        mrow.job_pol = Dr["job_pol_name"].ToString();
                        mrow.job_pod = Dr["job_pod_name"].ToString();
                        mrow.job_pofd_name = Dr["job_pofd_name"].ToString();
                        mrow.salesman = Dr["salesman"].ToString();
                        mrow.opr_sbill_no = Dr["sbno"].ToString();
                        mrow.opr_sbill_date = Lib.DatetoStringDisplayformat(Dr["sb_date"]);
                        mrow.opr_cargo_received_on = Lib.DatetoStringDisplayformat(Dr["cargo_date"]);
                        mrow.opr_cleared_date = Lib.DatetoStringDisplayformat(Dr["cleard_date"]);
                        mrow.job_terms = Dr["freight_status"].ToString();
                        mrow.job_nomination = Dr["job_nomination"].ToString();
                        mrow.job_chwt = Lib.Conv2Decimal(Dr["job_chwt"].ToString());
                        mrow.job_grwt = Lib.Conv2Decimal(Dr["job_grwt"].ToString());
                        mrow.job_ntwt = Lib.Conv2Decimal(Dr["job_ntwt"].ToString());
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.hbl_bl_no = Dr["house"].ToString();
                        mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["house_date"]);
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.mbl_bl_no = Dr["mbl_bl_no"].ToString();
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.liner_name = Dr["liner_name"].ToString();
                        mrow.forwarder_name = Dr["forwarder_name"].ToString();
                        mrow.mbl_grwt = Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                        mrow.mbl_chwt = Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                        mrow.mbl_folder_no = Dr["folderno"].ToString();
                        mrow.mbl_folder_sent_date = Lib.DatetoStringDisplayformat(Dr["folder_sent_date"]);
                        mrow.job_pkg = Lib.Conv2Decimal(Dr["job_total_cartons"].ToString());
                        mrow.job_invoice_nos = Dr["job_invoice_nos"].ToString();
                        mrow.opr_drawback_slno = Dr["opr_drawback_slno"].ToString();
                        mrow.opr_drawback_date = Lib.DatetoStringDisplayformat(Dr["opr_drawback_date"]);
                        mrow.opr_drawback_amt = Lib.Conv2Decimal(Dr["opr_drawback_amt"].ToString());
                        mrow.job_remarks = Dr["job_remarks"].ToString();
                        mrow.job_nomination = Dr["nomination"].ToString();
                        mrow.job_commodity = Dr["commodity"].ToString();
                        mrow.job_terms = Dr["job_terms"].ToString();
                        mrow.job_status = Dr["job_status"].ToString();
                        mrow.branch = Dr["branch"].ToString();

                        mrow.hbl_ar_invnos = Dr["our_invoice"].ToString();
                        mrow.hbl_ar_invamt = Lib.Conv2Decimal(Dr["hbl_ar_invamt"].ToString());
                        mrow.hbl_ar_gstamt = Lib.Conv2Decimal(Dr["hbl_ar_gstamt"].ToString());

                        job_agent = Dr["job_agent"].ToString();
                        job_liner = Dr["job_liner_name"].ToString();

                        if (job_agent != "")
                            job_liner_agent = job_agent;
                        if (job_liner_agent != "" && job_liner != "")
                            job_liner_agent += " / ";
                        job_liner_agent += job_liner;

                        mrow.job_liner_agent = job_liner_agent;

                        mList.Add(mrow);

                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintDsrAirReport();
                    }
                    Dt_List.Rows.Clear();
                }


                if (rec_category == "SEA IMPORT" || rec_category == "AIR IMPORT")
                {

                    sql = "  select h.hbl_no as hbl_no,h.hbl_bl_no as hbl_bl_no,h.hbl_date as hbl_date,m.hbl_no as mbl_no,m.hbl_bl_no as  mbl_bl_no,m.hbl_date as mbl_date ";
                    sql += "  ,shpr.cust_name as exporter_name,m.rec_branch_code as branch";
                    sql += "  ,cnge.cust_name as importer_name,nvl(sman2.param_name, sman.param_name) as sman_name";
                    sql += "  ,agent.cust_name as agent_name,impj_edi_no as job_edi_no,impj_sbno,impj_sbdate";
                    sql += "  ,h.hbl_invoice_nos as job_invno,pol.param_name as mbl_pol_name,pod.param_name as mbl_pod_name,h.hbl_pkg as hbl_cartons,h.hbl_cbm as hbl_cbm,h.hbl_grwt as hbl_grwt,h.hbl_chwt as hbl_chwt";
                    sql += "  ,h.hbl_bl_no,cha.cust_name as cha_name,m.hbl_pod_eta as mbl_eta,impj_remarks as remarks";
                    sql += "  ,m.hbl_jobtype as hbl_be_type,impj_docs_required as hbl_docs_required,impj_edichklst_sent_on as hbl_edichklst_sent_on";
                    sql += "  ,h.hbl_beno as hbl_beno,h.hbl_bedate as hbl_debate,impj_status as hbl_status,impj_status_date as hbl_status_date ";
                    sql += "  ,impj_cleared_on as hbl_cleared_on,h.hbl_invoice_nos,h.hbl_remarks as hbl_remarks ";
                    sql += "  ,lnr.param_name as liner_name,impj_doc_recvd_date as hbl_doc_recvd_date,impj_doc_send_date as hbl_doc_send_date,impj_waybill_no as hbl_waybill_no,impj_waybill_date as hbl_waybill_date ";
                    sql += "  ,fwd.cust_name as forwarder_name,h.hbl_ar_invnos as our_invoice,h.hbl_ar_invamt,h.hbl_ar_gstamt ";
                    sql += "  ,jobsts.param_name as job_status ";
                    sql += "  from hblm m";
                    sql += "  left join hblm h on m.hbl_pkid = h.hbl_mbl_id";
                    sql += "  left join impjobm j on h.hbl_pkid = j.impj_parent_id";
                    sql += "  left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";

                    sql += "  left join customerm cnge on h.hbl_imp_id = cnge.cust_pkid ";
                    sql += "  left join custdet  cd on h.rec_branch_code = cd.det_branch_code and h.hbl_imp_id = cd.det_cust_id ";

                    sql += " left join param sman on cnge.cust_sman_id = sman.param_pkid ";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += "  left join customerm agent on m.hbl_agent_id = agent.cust_pkid ";
                    sql += "  left join param pol on  m.hbl_pol_id = pol.param_pkid";
                    sql += "  left join param pod on  m.hbl_pod_id = pod.param_pkid";
                    sql += "  left join customerm cha on m.hbl_cha_id = cha.cust_pkid";
                    sql += "  left join param lnr  on m.hbl_carrier_id = lnr.param_pkid";
                    sql += "  left join customerm fwd on m.hbl_forwarder_id = fwd.cust_pkid";
                    sql += "  left join param jobsts  on m.hbl_status_id = jobsts.param_pkid";
                    if (rec_category == "SEA IMPORT")
                    {
                        sql += "  where  m.rec_category = 'SEA IMPORT'  and m.hbl_type = 'MBL-SI' ";

                        if(!all)
                        {
                            sql += " and m.rec_branch_code = '{BRCODE}' ";
                        }

                    }
                    else
                    {
                        sql += "  where m.rec_category = 'AIR IMPORT'  and m.hbl_type = 'MBL-AI' ";
                        if (!all)
                        {
                            sql += " and m.rec_branch_code = '{BRCODE}' ";
                        }
                    }
                    sql += "  and m.hbl_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    if (job_type != "ALL")
                        sql += " and m.hbl_jobtype ='{JOBTYPE}'";

                    if (shipper_id.Length > 0)
                        sql += " and h.hbl_imp_id = '" + shipper_id + "' ";

                    if (consignee_id.Length > 0)
                        sql += " and  h.hbl_exp_id = '" + consignee_id + "' ";
                    if (agent_id.Length > 0)
                        sql += " and m.hbl_agent_id = '" + agent_id + "' ";
                    if (carrier_id.Length > 0)
                        sql += " and m.hbl_carrier_id = '" + carrier_id + "'";
                    if (pol_id.Length > 0)
                        sql += " and m.hbl_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and m.hbl_pod_id = '" + pod_id + "'";

                    if (type_date == "DATE")
                        sql += " order by m.rec_branch_code,m.hbl_date";
                    else
                        sql += " order by m.rec_branch_code,m.hbl_no";

                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{JOBTYPE}", job_type);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new DsrReport();
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.hbl_bl_no = Dr["hbl_bl_no"].ToString();
                        mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hbl_date"]);
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.mbl_bl_no = Dr["mbl_bl_no"].ToString();
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.hbl_exporter_name = Dr["exporter_name"].ToString();
                        mrow.hbl_importer_name = Dr["importer_name"].ToString();
                        mrow.hbl_agent_name = Dr["agent_name"].ToString();
                        mrow.impj_edi_no = Dr["job_edi_no"].ToString();
                        mrow.hbl_invoice_nos = Dr["job_invno"].ToString();
                        mrow.mbl_pol = Dr["mbl_pol_name"].ToString();
                        mrow.mbl_pod = Dr["mbl_pod_name"].ToString();
                        mrow.hbl_pkg = Lib.Conv2Decimal(Dr["hbl_cartons"].ToString());
                        mrow.hbl_cbm = Lib.Conv2Decimal(Dr["hbl_cbm"].ToString());
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        mrow.hbl_chwt = Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());
                        mrow.cha_name = Dr["cha_name"].ToString();
                        mrow.hbl_pod_eta = Lib.DatetoStringDisplayformat(Dr["mbl_eta"]);
                        mrow.impj_be_type = Dr["hbl_be_type"].ToString();
                        mrow.impj_docs_required = Dr["hbl_docs_required"].ToString();
                        mrow.impj_edichklst_sent_on = Lib.DatetoStringDisplayformat(Dr["hbl_edichklst_sent_on"]);
                        mrow.hbl_beno = Dr["hbl_beno"].ToString();
                        mrow.hbl_bedate = Lib.DatetoStringDisplayformat(Dr["hbl_debate"]);
                        mrow.impj_status = Dr["hbl_status"].ToString();
                        mrow.impj_status_date = Lib.DatetoStringDisplayformat(Dr["hbl_status_date"]);
                        mrow.impj_cleared_on = Lib.DatetoStringDisplayformat(Dr["hbl_cleared_on"]);
                        mrow.hbl_remarks = Dr["remarks"].ToString();
                        mrow.liner_name = Dr["liner_name"].ToString();
                        mrow.impj_doc_recvd_date = Lib.DatetoStringDisplayformat(Dr["hbl_doc_recvd_date"]);
                        mrow.impj_doc_send_date = Lib.DatetoStringDisplayformat(Dr["hbl_doc_send_date"]);
                        mrow.impj_waybill_no = Dr["hbl_waybill_no"].ToString();
                        mrow.impj_waybill_date = Lib.DatetoStringDisplayformat(Dr["hbl_waybill_date"]);
                        mrow.forwarder_name = Dr["forwarder_name"].ToString();
                        mrow.impj_sbno = Dr["impj_sbno"].ToString();
                        mrow.impj_sbdate = Lib.DatetoStringDisplayformat(Dr["impj_sbdate"]);
                        mrow.salesman = Dr["sman_name"].ToString();
                        mrow.branch = Dr["branch"].ToString();

                        mrow.hbl_ar_invnos = Dr["our_invoice"].ToString();
                        mrow.hbl_ar_invamt = Lib.Conv2Decimal(Dr["hbl_ar_invamt"].ToString());
                        mrow.hbl_ar_gstamt = Lib.Conv2Decimal(Dr["hbl_ar_gstamt"].ToString());
                        mrow.job_status = Dr["job_status"].ToString();

                        mList.Add(mrow);

                       
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintImportDsrReport();
                    }
                    Dt_List.Rows.Clear();


                }

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", type);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);
            return RetData;
        }

        public IDictionary<string, object> UpdateDsrRemarks(Dictionary<string, object> SearchData)
        {
            string pkid = SearchData["pkid"].ToString();           
            string remarks = SearchData["remarks"].ToString();
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            try
            {

                //if (remarks.Length > 60)
                //{
                //    remarks = remarks.Substring(0, 60);
                //}

                Con_Oracle = new DBConnection();

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("jobm", "EDIT", "job_pkid", pkid);              
                Rec.InsertString("job_remarks", remarks);


                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                    throw Ex;
                }
            }
            RetData.Add("status", "OK");
            return RetData;
        }


        private void PrintDsrReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(all)
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                else
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "DSRReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
               
                if (!all)
                {

                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 10;
                    WS.Columns[2].Width = 256 * 13;
                    WS.Columns[3].Width = 256 * 12;
                    WS.Columns[4].Width = 256 * 20;
                    WS.Columns[5].Width = 256 * 20;
                    WS.Columns[6].Width = 256 * 50;
                    WS.Columns[7].Width = 256 * 15;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 20;
                    WS.Columns[10].Width = 256 * 20;
                    WS.Columns[11].Width = 256 * 18;
                    WS.Columns[12].Width = 256 * 16;
                    WS.Columns[13].Width = 256 * 13;
                    WS.Columns[14].Width = 256 * 13;
                    WS.Columns[15].Width = 256 * 13;
                    WS.Columns[16].Width = 256 * 14;
                    WS.Columns[17].Width = 256 * 15;
                    WS.Columns[18].Width = 256 * 17;
                    WS.Columns[19].Width = 256 * 15;
                    WS.Columns[20].Width = 256 * 15;
                    WS.Columns[21].Width = 256 * 15;
                    WS.Columns[22].Width = 256 * 20;
                    WS.Columns[23].Width = 256 * 20;
                    WS.Columns[24].Width = 256 * 17;
                    WS.Columns[25].Width = 256 * 17;
                    WS.Columns[26].Width = 256 * 15;
                    WS.Columns[27].Width = 256 * 18;
                    WS.Columns[28].Width = 256 * 15;
                    WS.Columns[29].Width = 256 * 10;
                    WS.Columns[30].Width = 256 * 15;
                    WS.Columns[31].Width = 256 * 20;
                    WS.Columns[32].Width = 256 * 15;
                    WS.Columns[33].Width = 256 * 15;
                    WS.Columns[34].Width = 256 * 15;
                    WS.Columns[35].Width = 256 * 7;
                    WS.Columns[36].Width = 256 * 15;
                    WS.Columns[37].Width = 256 * 20;
                    WS.Columns[38].Width = 256 * 18;
                    WS.Columns[39].Width = 256 * 18;
                    WS.Columns[40].Width = 256 * 15;
                    WS.Columns[41].Width = 256 * 7;
                    WS.Columns[42].Width = 256 * 15;
                    WS.Columns[43].Width = 256 * 15;
                    WS.Columns[44].Width = 256 * 15;
                    WS.Columns[45].Width = 256 * 15;
                    WS.Columns[46].Width = 256 * 15;
                    WS.Columns[47].Width = 256 * 16;
                    WS.Columns[48].Width = 256 * 16;
                    WS.Columns[49].Width = 256 * 16;
                    WS.Columns[50].Width = 256 * 15;
                    WS.Columns[51].Width = 256 * 18;
                    WS.Columns[52].Width = 256 * 35;
                    WS.Columns[53].Width = 256 * 15;
                    WS.Columns[54].Width = 256 * 25;
                    WS.Columns[55].Width = 256 * 15;
                   
                    
                }
                else
                {

                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 14;
                    WS.Columns[2].Width = 256 * 10;
                    WS.Columns[3].Width = 256 * 13;
                    WS.Columns[4].Width = 256 * 12;
                    WS.Columns[5].Width = 256 * 20;
                    WS.Columns[6].Width = 256 * 20;
                    WS.Columns[7].Width = 256 * 50;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 15;
                    WS.Columns[10].Width = 256 * 20;
                    WS.Columns[11].Width = 256 * 20;
                    WS.Columns[12].Width = 256 * 18;
                    WS.Columns[13].Width = 256 * 16;
                    WS.Columns[14].Width = 256 * 13;
                    WS.Columns[15].Width = 256 * 13;
                    WS.Columns[16].Width = 256 * 13;
                    WS.Columns[17].Width = 256 * 14;
                    WS.Columns[18].Width = 256 * 15;
                    WS.Columns[19].Width = 256 * 17;
                    WS.Columns[20].Width = 256 * 15;
                    WS.Columns[21].Width = 256 * 15;
                    WS.Columns[22].Width = 256 * 15;
                    WS.Columns[23].Width = 256 * 20;
                    WS.Columns[24].Width = 256 * 20;
                    WS.Columns[25].Width = 256 * 17;
                    WS.Columns[26].Width = 256 * 17;
                    WS.Columns[27].Width = 256 * 15;
                    WS.Columns[28].Width = 256 * 18;
                    WS.Columns[29].Width = 256 * 15;
                    WS.Columns[30].Width = 256 * 10;
                    WS.Columns[31].Width = 256 * 15;
                    WS.Columns[32].Width = 256 * 20;
                    WS.Columns[33].Width = 256 * 15;
                    WS.Columns[34].Width = 256 * 15;
                    WS.Columns[35].Width = 256 * 15;
                    WS.Columns[36].Width = 256 * 7;
                    WS.Columns[37].Width = 256 * 15;
                    WS.Columns[38].Width = 256 * 20;
                    WS.Columns[39].Width = 256 * 18;
                    WS.Columns[40].Width = 256 * 18;
                    WS.Columns[41].Width = 256 * 15;
                    WS.Columns[42].Width = 256 * 7;
                    WS.Columns[43].Width = 256 * 15;
                    WS.Columns[44].Width = 256 * 15;
                    WS.Columns[45].Width = 256 * 15;
                    WS.Columns[46].Width = 256 * 15;
                    WS.Columns[47].Width = 256 * 15;
                    WS.Columns[48].Width = 256 * 16;
                    WS.Columns[49].Width = 256 * 16;
                    WS.Columns[50].Width = 256 * 16;
                    WS.Columns[51].Width = 256 * 15;
                    WS.Columns[52].Width = 256 * 18;
                    WS.Columns[53].Width = 256 * 35;
                    WS.Columns[54].Width = 256 * 15;
                    WS.Columns[55].Width = 256 * 25;
                    WS.Columns[56].Width = 256 * 15;

                }




                iRow = 0; iCol = 1;
               
                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "DSR-SE ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }

                Lib.WriteData(WS, iRow, iCol++, "JOB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CLR/REF#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "JOB/AGENT/CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JOB TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TERMS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXP INV#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NO OF PKGS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NT WT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GR WT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VOLUME / M3", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPMENT MODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINER SIZE/TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "S.BILLL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "S/B DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHA", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARGO RECEIVED ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BOOKING NUMBER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BOOKING DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
              
                Lib.WriteData(WS, iRow, iCol++, "CARTED / STUFFED AT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINER STUFFED ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VESSEL NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VOYAGE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ETD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "JOB/CONTAINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "SI FILED ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SOB", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ETA DESTINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINER#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "HBL RELEASED ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL RELEASED ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PRE-ALERT SEND ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OUR INVOCIE#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AMOUNT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST AMOUNT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                // Lib.WriteData(WS, iRow, iCol++, "Our Invoice Date", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAYMENT RECEIVED ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL NUMBER OF TUE'S", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALES PERSON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "E/P RECEIVED ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SCHEME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);


                foreach (DsrReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if(all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.job_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_prefix, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_shipper, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_liner_agent, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_invoice_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pkg, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_commodity, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_nature, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_cntr_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_sbill_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.opr_sbill_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_cha_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_cargo_received_on, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_book_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                   
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_stuffed_at, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.opr_stuffed_on, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_vessel_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_vessel_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_book_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pofd_eta, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.liner_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_released_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_released_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_prealert_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_gstamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    //  Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_cntr_teu, _Color, false, "", "L", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.salesman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_remarks, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.opr_ep_rec_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_billtype_id, _Color, false, "", "L", "", _Size, false, 325, "", true);

                }
                // iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
        private void PrintStatusDsrReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if (all)
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                else
                {
                    mSearchData.Add("branch_code", branch_code);
                }


                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "DSRReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 6;
                WS.Columns[2].Width = 256 * 11;
                WS.Columns[3].Width = 256 * 34;
                WS.Columns[4].Width = 256 * 38;
                WS.Columns[5].Width = 256 * 33;
                WS.Columns[6].Width = 256 * 29;
                WS.Columns[7].Width = 256 * 16;
                WS.Columns[8].Width = 256 * 11;
                WS.Columns[9].Width = 256 * 19;
                WS.Columns[10].Width = 256 * 23;
                WS.Columns[11].Width = 256 * 13;
                WS.Columns[12].Width = 256 * 9;
                WS.Columns[13].Width = 256 * 9;
                WS.Columns[14].Width = 256 * 13;
                WS.Columns[15].Width = 256 * 9;
                WS.Columns[16].Width = 256 * 12;
                WS.Columns[17].Width = 256 * 12;
                WS.Columns[18].Width = 256 * 15;


                iRow = 0; iCol = 1;
  
                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "DSR-SE ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "JOB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SBILL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARGO-RECEIVED-ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CNTR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VESSEL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VESSEL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SOB", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DESTINATION-ETA", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                foreach (DsrReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.job_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_shipper, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_invoice_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_sbill_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.opr_sbill_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_cargo_received_on, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_vessel_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_vessel_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pofd_eta, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                }
                // iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }


        private void PrintShipperDsrReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            iRow = 0;
            iCol = 0;
            int i = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if (all)
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                else
                {
                    mSearchData.Add("branch_code", branch_code);
                }


                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "DSRReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                if(all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 6;
                    WS.Columns[2].Width = 256 * 10;
                    WS.Columns[3].Width = 256 * 7;
                    WS.Columns[4].Width = 256 * 7;
                    WS.Columns[5].Width = 256 * 33;
                    WS.Columns[6].Width = 256 * 33;
                    WS.Columns[7].Width = 256 * 24;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 11;
                    WS.Columns[10].Width = 256 * 9;
                    WS.Columns[11].Width = 256 * 9;
                    WS.Columns[12].Width = 256 * 18;
                    WS.Columns[13].Width = 256 * 16;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 10;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 19;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 13;
                    WS.Columns[20].Width = 256 * 21;
                    WS.Columns[21].Width = 256 * 22;
                    WS.Columns[22].Width = 256 * 25;
                    WS.Columns[23].Width = 256 * 20;
                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 6;
                    WS.Columns[2].Width = 256 * 7;
                    WS.Columns[3].Width = 256 * 7;
                    WS.Columns[4].Width = 256 * 33;
                    WS.Columns[5].Width = 256 * 33;
                    WS.Columns[6].Width = 256 * 24;
                    WS.Columns[7].Width = 256 * 12;
                    WS.Columns[8].Width = 256 * 11;
                    WS.Columns[9].Width = 256 * 9;
                    WS.Columns[10].Width = 256 * 9;
                    WS.Columns[11].Width = 256 * 18;
                    WS.Columns[12].Width = 256 * 16;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 10;
                    WS.Columns[15].Width = 256 * 12;
                    WS.Columns[16].Width = 256 * 19;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 13;
                    WS.Columns[19].Width = 256 * 21;
                    WS.Columns[20].Width = 256 * 22;
                    WS.Columns[21].Width = 256 * 25;
                    WS.Columns[22].Width = 256 * 15;
                }
               


                iRow = 0; iCol = 1;

                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "DSR-SE ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "SL#", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "JOB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NO.OF PKGS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GROSS WT.", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINER#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VESSEL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ETD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ETA", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOAD PORT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DISCHARGE PORT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VESSEL 2", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                foreach (DsrReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    i++;
                    Lib.WriteData(WS, iRow, iCol++, i, _Color, false, "", "R", "", _Size, false, 325, "", true);
                    if (all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_shipper, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_invoice_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pkg, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_nature, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_sbill_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.opr_sbill_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_vessel_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pofd_eta, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_remarks, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_vessel2_name, _Color, false, "", "L", "", _Size, false, 325, "", true);

                }
                // iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void PrintImportDsrReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(all)
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                else
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "DSRReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                if(!all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 15;
                    WS.Columns[2].Width = 256 * 6;
                    WS.Columns[3].Width = 256 * 15;
                    WS.Columns[4].Width = 256 * 6;
                    WS.Columns[5].Width = 256 * 15;
                    WS.Columns[6].Width = 256 * 15;
                    WS.Columns[7].Width = 256 * 20;
                    WS.Columns[8].Width = 256 * 20;
                    WS.Columns[9].Width = 256 * 20;
                    WS.Columns[10].Width = 256 * 15;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 15;
                    WS.Columns[16].Width = 256 * 15;
                    WS.Columns[17].Width = 256 * 15;
                    WS.Columns[18].Width = 256 * 17;
                    WS.Columns[19].Width = 256 * 15;
                    WS.Columns[20].Width = 256 * 15;
                    WS.Columns[21].Width = 256 * 15;
                    WS.Columns[22].Width = 256 * 15;
                    WS.Columns[23].Width = 256 * 15;
                    WS.Columns[24].Width = 256 * 17;
                    WS.Columns[25].Width = 256 * 15;
                    WS.Columns[26].Width = 256 * 15;
                    WS.Columns[27].Width = 256 * 15;
                    WS.Columns[28].Width = 256 * 20;
                    WS.Columns[29].Width = 256 * 20;
                    WS.Columns[30].Width = 256 * 15;
                    WS.Columns[31].Width = 256 * 15;
                    WS.Columns[32].Width = 256 * 15;
                    WS.Columns[33].Width = 256 * 15;
                    WS.Columns[34].Width = 256 * 20;
                    WS.Columns[35].Width = 256 * 15;

                    WS.Columns[36].Width = 256 * 20;
                    WS.Columns[37].Width = 256 * 20;

                   
                    WS.Columns[38].Width = 256 * 15;
                    WS.Columns[39].Width = 256 * 15;
                    WS.Columns[40].Width = 256 * 15;

                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 15;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 6;
                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 6;
                    WS.Columns[6].Width = 256 * 15;
                    WS.Columns[7].Width = 256 * 15;
                    WS.Columns[8].Width = 256 * 20;
                    WS.Columns[9].Width = 256 * 20;
                    WS.Columns[10].Width = 256 * 20;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 15;
                    WS.Columns[14].Width = 256 * 15;
                    WS.Columns[15].Width = 256 * 15;
                    WS.Columns[16].Width = 256 * 15;
                    WS.Columns[17].Width = 256 * 15;
                    WS.Columns[18].Width = 256 * 15;
                    WS.Columns[19].Width = 256 * 17;
                    WS.Columns[20].Width = 256 * 15;
                    WS.Columns[21].Width = 256 * 15;
                    WS.Columns[22].Width = 256 * 15;
                    WS.Columns[23].Width = 256 * 15;
                    WS.Columns[24].Width = 256 * 15;
                    WS.Columns[25].Width = 256 * 17;
                    WS.Columns[26].Width = 256 * 15;
                    WS.Columns[27].Width = 256 * 15;
                    WS.Columns[28].Width = 256 * 15;
                    WS.Columns[29].Width = 256 * 20;
                    WS.Columns[30].Width = 256 * 20;
                    WS.Columns[31].Width = 256 * 15;
                    WS.Columns[32].Width = 256 * 15;
                    WS.Columns[33].Width = 256 * 15;
                    WS.Columns[34].Width = 256 * 15;
                    WS.Columns[35].Width = 256 * 20;
                    WS.Columns[36].Width = 256 * 15;

                    WS.Columns[37].Width = 256 * 15;
                    WS.Columns[38].Width = 256 * 20;

                    WS.Columns[39].Width = 256 * 15;
                    WS.Columns[40].Width = 256 * 15;
                    WS.Columns[41].Width = 256 * 15;
                    WS.Columns[42].Width = 256 * 15;
                    WS.Columns[43].Width = 256 * 15;

                }



                iRow = 0; iCol = 1;
                
               
                //WS.Columns[17].Style.NumberFormat = "#0.000";

                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                if (rec_category == "SEA IMPORT")
                {
                    Lib.WriteData(WS, iRow, 1, "DSR-SI ", _Color, true, "", "L", "", 15, false, 325, "", true);
                }
                else
                {
                    Lib.WriteData(WS, iRow, 1, "DSR-AI ", _Color, true, "", "L", "", 15, false, 325, "", true);
                }
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MASTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOUSE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IMPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
               // Lib.WriteData(WS, iRow, iCol++, "EDI-JOB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "INV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                // Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "PKG", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                //Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
              //  Lib.WriteData(WS, iRow, iCol++, "SBILL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
              //  Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD-ETA", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DOCS-REQUIRED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EDI-CHECKLIST-SENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BE.NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BE.DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CLEARED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DOC-RECEIVED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DOC-SENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                if (rec_category == "AIR IMPORT")
                {
                    Lib.WriteData(WS, iRow, iCol++, "WAYBILL.NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "WAYBILL.DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "OUR INVOICE#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AMOUNT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST AMOUNT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHA", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JOB-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                foreach (DsrReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if(all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }

                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_exporter_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_importer_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.salesman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                   // Lib.WriteData(WS, iRow, iCol++, Rec.impj_edi_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_invoice_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pkg, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                   // Lib.WriteData(WS, iRow, iCol++, Rec.impj_sbno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                  //  Lib.WriteData(WS, iRow, iCol++, Rec.impj_sbdate, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_pod_eta, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.impj_be_type, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.impj_docs_required, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.impj_edichklst_sent_on, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_beno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_bedate, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.impj_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.impj_status_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.impj_cleared_on, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_remarks, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.liner_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.impj_doc_recvd_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.impj_doc_send_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    if (rec_category == "AIR IMPORT")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.impj_waybill_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.impj_waybill_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    }
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_gstamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cha_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                }
                // iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void PrintDsrAirReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if (all)
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                else
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "AIR-DSRReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                if (!all)
                {

                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 8;
                    WS.Columns[2].Width = 256 * 11;
                    WS.Columns[3].Width = 256 * 9;
                    WS.Columns[4].Width = 256 * 22;
                    WS.Columns[5].Width = 256 * 22;
                    WS.Columns[6].Width = 256 * 21;
                    WS.Columns[7].Width = 256 * 12;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 41;
                    WS.Columns[10].Width = 256 * 14;
                    WS.Columns[11].Width = 256 * 11;
                    WS.Columns[12].Width = 256 * 12;
                    WS.Columns[13].Width = 256 * 14;
                    WS.Columns[14].Width = 256 * 11;
                    WS.Columns[15].Width = 256 * 10;
                    WS.Columns[16].Width = 256 * 12;
                    WS.Columns[17].Width = 256 * 15;
                    WS.Columns[18].Width = 256 * 12;
                    WS.Columns[19].Width = 256 * 16;
                    WS.Columns[20].Width = 256 * 16;
                    WS.Columns[21].Width = 256 * 14;
                    WS.Columns[22].Width = 256 * 14;
                    WS.Columns[23].Width = 256 * 14;
                    WS.Columns[24].Width = 256 * 14;
                    WS.Columns[25].Width = 256 * 7;
                    WS.Columns[26].Width = 256 * 9;
                    WS.Columns[27].Width = 256 * 14;
                    WS.Columns[28].Width = 256 * 23;
                    WS.Columns[29].Width = 256 * 7;
                    WS.Columns[30].Width = 256 * 14;
                    WS.Columns[31].Width = 256 * 14;
                    WS.Columns[32].Width = 256 * 19;
                    WS.Columns[33].Width = 256 * 14;
                    WS.Columns[34].Width = 256 * 14;
                    WS.Columns[35].Width = 256 * 14;
                    WS.Columns[36].Width = 256 * 14;
                    WS.Columns[37].Width = 256 * 20;
                    WS.Columns[38].Width = 256 * 29;
                    WS.Columns[39].Width = 256 * 14;
                    WS.Columns[40].Width = 256 * 14;
                    WS.Columns[41].Width = 256 * 14;
                    WS.Columns[42].Width = 256 * 14;
                }
                else
                {

                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 8;
                    WS.Columns[3].Width = 256 * 11;
                    WS.Columns[4].Width = 256 * 9;
                    WS.Columns[5].Width = 256 * 22;
                    WS.Columns[6].Width = 256 * 22;
                    WS.Columns[7].Width = 256 * 21;
                    WS.Columns[8].Width = 256 * 12;
                    WS.Columns[9].Width = 256 * 12;
                    WS.Columns[10].Width = 256 * 41;
                    WS.Columns[11].Width = 256 * 14;
                    WS.Columns[12].Width = 256 * 11;
                    WS.Columns[13].Width = 256 * 12;
                    WS.Columns[14].Width = 256 * 14;
                    WS.Columns[15].Width = 256 * 11;
                    WS.Columns[16].Width = 256 * 10;
                    WS.Columns[17].Width = 256 * 12;
                    WS.Columns[18].Width = 256 * 15;
                    WS.Columns[19].Width = 256 * 12;
                    WS.Columns[20].Width = 256 * 16;
                    WS.Columns[21].Width = 256 * 16;
                    WS.Columns[22].Width = 256 * 14;
                    WS.Columns[23].Width = 256 * 14;
                    WS.Columns[24].Width = 256 * 14;
                    WS.Columns[25].Width = 256 * 14;
                    WS.Columns[26].Width = 256 * 7;
                    WS.Columns[27].Width = 256 * 9;
                    WS.Columns[28].Width = 256 * 14;
                    WS.Columns[29].Width = 256 * 23;
                    WS.Columns[30].Width = 256 * 7;
                    WS.Columns[31].Width = 256 * 14;
                    WS.Columns[32].Width = 256 * 14;
                    WS.Columns[33].Width = 256 * 19;
                    WS.Columns[34].Width = 256 * 14;
                    WS.Columns[35].Width = 256 * 14;
                    WS.Columns[36].Width = 256 * 14;
                    WS.Columns[37].Width = 256 * 14;
                    WS.Columns[38].Width = 256 * 20;
                    WS.Columns[39].Width = 256 * 29;
                    WS.Columns[40].Width = 256 * 14;
                    WS.Columns[41].Width = 256 * 14;
                    WS.Columns[42].Width = 256 * 14;
                    WS.Columns[43].Width = 256 * 14;
                }
                // WS.ViewOptions.ShowGridLines = false;
             


                iRow = 0; iCol = 1;
               
               
                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "DSR-AE ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "JOB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REF#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "JOB/AGENT/CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TERMS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SBILL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHA", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARGO-RECEIVED-ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CLEARED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOUSE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "MSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MASTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                //  Lib.WriteData(WS, iRow, iCol++, "FORWARDER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL-CARTONS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OUR INVOICE#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AMOUNT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST AMOUNT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                foreach (DsrReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if(all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.job_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_prefix, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_shipper, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_invoice_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pod, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_liner_agent, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_commodity, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_sbill_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.opr_sbill_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_cha_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pofd_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.salesman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.opr_cargo_received_on, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.opr_cleared_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.job_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.liner_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    // Lib.WriteData(WS, iRow, iCol++, Rec.forwarder_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_folder_sent_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_pkg, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.job_remarks, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_gstamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                }
                // iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }







    }
}

