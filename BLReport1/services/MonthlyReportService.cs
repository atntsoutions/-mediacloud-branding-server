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
    public class MonthlyReportService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<MonthlyReport> mList = new List<MonthlyReport>();
        MonthlyReport mrow;
        int iRow = 0;
        int iCol = 0;
        string type = "";
        string types = "";
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
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string agent_id = "";
        string shipper_id = "";
        string consignee_id = "";
        string carrier_id = "";
        string pol_id = "";
        string pod_id = "";
        string ErrorMessage = "";
        decimal tot_hbl_ntwt = 0;
        decimal tot_hbl_grwt = 0;
        decimal tot_hbl_chwt = 0;
        decimal tot_mbl_grwt = 0;
        decimal tot_mbl_chwt = 0;
        decimal tot_publish_rate = 0;
        decimal tot_informed_rate = 0;
        decimal tot_sell_informed = 0;
        decimal tot_rebate = 0;
        decimal tot_exwork = 0;
        decimal tot_cbm = 0;
        decimal tot_ntwt = 0;
        decimal tot_teu = 0;
        Boolean all = false;


        decimal tot_rebate_house = 0;

        Boolean bAdmin = false;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<MonthlyReport>();
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

                if (SearchData.ContainsKey("agent_id"))
                    agent_id = SearchData["agent_id"].ToString();

                if (SearchData.ContainsKey("shipper_id"))
                    shipper_id = SearchData["shipper_id"].ToString();

                if (SearchData.ContainsKey("consignee_id"))
                    consignee_id = SearchData["consignee_id"].ToString();

                if (SearchData.ContainsKey("carrier_id"))
                    carrier_id = SearchData["carrier_id"].ToString();

                if (SearchData.ContainsKey("pol_id"))
                    pol_id = SearchData["pol_id"].ToString();

                if (SearchData.ContainsKey("pod_id"))
                    pod_id = SearchData["pod_id"].ToString();

                all = (Boolean)SearchData["all"];


                bAdmin = (Boolean)SearchData["badmin"];

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

                if (rec_category == "HBL-AE")
                {
                    sql = " SELECT h.hbl_pkid,h.hbl_type,h.hbl_no as SINO,h.rec_created_date as SI_created_date,";
                    sql += " m.hbl_folder_no as folder_no ,m.HBL_FOLDER_SENT_DATE as folder_sent,";
                    sql += " m.hbl_bl_no as mbl_no, m.hbl_date as mbl_date,";
                    sql += " m.hbl_terms as mbl_status,h.hbl_bl_no as hbl_no, ";
                    sql += " h.hbl_date as hbl_date ,h.hbl_terms as hbl_status,h.rec_branch_code as branch,";

                    sql += " h.hbl_ar_invnos as inv_no,h.hbl_ar_invamt as inv_amt,h.hbl_ar_gstamt as gst_amt,";
                    sql += " shpr.cust_name as shipper_name,cons.cust_name as consignee_name,";
                    sql += " agent.cust_name as agent_name,agent.rec_created_date as created_on,";

                   // sql += " cons.cust_nomination as hbl_nomination,";
                    sql += " nvl(h.hbl_nomination,cons.cust_nomination) as hbl_nomination,";

                    sql += " carrier.param_name as carrier_name,pol.param_name as pol_name,";
                    sql += " pod.param_name as pod_name,pofd.param_name as pofd_name,";
                    sql += " m.hbl_pol_etd  as pol_etd,";

                  // sql += " nvl(sman2.param_name,sman.param_name) as sman_name ,";
                    sql += " nvl(sman1.param_name,nvl(sman2.param_name,sman.param_name)) as sman_name,";
                    sql += " nvl(sman1.param_pkid,nvl(sman2.param_pkid,sman.param_pkid)) as sman_id,";

                    sql += " h.hbl_grwt as hbl_grwt, h.hbl_chwt as hbl_chwt,m.hbl_grwt as mbl_grwt, m.hbl_chwt as mbl_chwt,";
                    sql += " air_netnet as netnet,air_publish_rate  as publish_rate, ";
                    sql += " air_counter_informed as informed_rate,air_sell_informed as sell_informed,";
                    sql += " air_rebate as rebate, h.hbl_rebate_amt_inr as rebate_house, air_exworks as exworks,commodity.param_name as commodty_name";
                    sql += " from hblm h ";
                    sql += " left join hblm m on h.hbl_mbl_id = m.hbl_pkid";

                    sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";
                    sql += " left join custdet  cd on h.rec_branch_code = cd.det_branch_code and h.hbl_exp_id = cd.det_cust_id ";

                    sql += " left join param sman on shpr.cust_sman_id =sman.param_pkid ";
                    sql += " left join param sman1 on h.hbl_salesman_id = sman1.param_pkid";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
                    sql += " left join customerm agent on h.hbl_agent_id = agent.cust_pkid";
                    sql += " left join param carrier on h.hbl_carrier_id = carrier.param_pkid";
                    sql += " left join param pol on h.hbl_pol_id = pol.param_pkid";
                    sql += " left join param pod on h.hbl_pod_id = pod.param_pkid";
                    sql += " left join param pofd on h.hbl_pofd_id = pofd.param_pkid";
                    sql += " left join customerm cha on h.hbl_cha_id = cha.cust_pkid";
                    sql += " left join param commodity on m.hbl_commodity_id = commodity.param_pkid";
                    sql += " left join param vessel on m.hbl_vessel_id = vessel.param_pkid";
                    sql += " left join aircostm acostm on m.hbl_pkid = acostm.air_mblid";
                    sql += " where h.rec_category = 'AIR EXPORT' ";
                    if (!all)
                    {
                        sql += " and h.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += " and h.hbl_type ='HBL-AE'";

                    if (type_date == "SOB")
                        sql += "  and m.hbl_pol_etd between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    else
                        sql += "  and to_char(h.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    if (agent_id.Length > 0)
                        sql += " and h.hbl_agent_id = '" + agent_id + "' ";
                    if (shipper_id.Length > 0)
                        sql += " and h.hbl_exp_id = '" + shipper_id + "' ";
                    if (consignee_id.Length > 0)
                        sql += " and h.hbl_imp_id = '" + consignee_id + "'";
                    if (carrier_id.Length > 0)
                        sql += " and h.hbl_carrier_id = '" + carrier_id + "'";
                    if (pol_id.Length > 0)
                        sql += " and h.hbl_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and h.hbl_pod_id = '" + pod_id + "'";


                    if (type_date == "SOB")
                        sql += " order by h.rec_branch_code,m.hbl_date";
                    else
                        sql += " order by h.rec_branch_code,h.rec_created_date";


                    sql = sql.Replace("{TYPES}", types);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{CATEGORY}", rec_category);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();


                    tot_hbl_grwt = 0;
                    tot_hbl_chwt = 0;
                    tot_mbl_grwt = 0;
                    tot_mbl_chwt = 0;
                    tot_publish_rate = 0;
                    tot_informed_rate = 0;
                    tot_sell_informed = 0;
                    tot_rebate = 0;
                    tot_exwork = 0;

                    string pre_data = "";
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new MonthlyReport();
                        mrow.displayed = false;
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        if (pre_data != Dr["mbl_no"].ToString())
                        {
                            pre_data = Dr["mbl_no"].ToString();
                            mrow.mbl_grwt = Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                            mrow.mbl_chwt = Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                            mrow.netnet = Dr["netnet"].ToString();
                            mrow.publish_rate = Lib.Conv2Decimal(Dr["publish_rate"].ToString());
                            mrow.informed_rate = Lib.Conv2Decimal(Dr["informed_rate"].ToString());
                            mrow.sell_informed = Lib.Conv2Decimal(Dr["sell_informed"].ToString());
                            mrow.rebate = Lib.Conv2Decimal(Dr["rebate"].ToString());
                            mrow.exworks = Lib.Conv2Decimal(Dr["exworks"].ToString());
                        }
                        else
                        {
                            mrow.mbl_grwt = 0;
                            mrow.mbl_chwt = 0;
                            mrow.publish_rate = 0;
                            mrow.informed_rate = 0;
                            mrow.sell_informed = 0;
                            mrow.rebate = 0;
                            mrow.exworks = 0;
                            mrow.netnet = "";

                        }
                        mrow.hbl_pkid = Dr["hbl_pkid"].ToString();
                        mrow.hbl_type = Dr["hbl_type"].ToString();
                        mrow.sino = Dr["SINO"].ToString();
                        mrow.folder_no = Dr["folder_no"].ToString();
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.folder_sent = Lib.DatetoStringDisplayformat(Dr["folder_sent"]);
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.mbl_status = Dr["mbl_status"].ToString();
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hbl_date"]);
                        mrow.hbl_status = Dr["hbl_status"].ToString();
                        mrow.shipper_name = Dr["shipper_name"].ToString();
                        mrow.consignee_name = Dr["consignee_name"].ToString();
                        mrow.agent_name = Dr["agent_name"].ToString();
                        mrow.hbl_nomination = Dr["hbl_nomination"].ToString();
                        mrow.carrier_name = Dr["carrier_name"].ToString();
                        mrow.pol_name = Dr["pol_name"].ToString();
                        mrow.pod_name = Dr["pod_name"].ToString();
                        mrow.pofd_name = Dr["pofd_name"].ToString();
                        mrow.pol_etd = Lib.DatetoStringDisplayformat(Dr["pol_etd"]);
                        mrow.sman_name = Dr["sman_name"].ToString();
                        mrow.sman_id = Dr["sman_id"].ToString();
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        mrow.hbl_chwt = Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());
                        mrow.commodty_name = Dr["commodty_name"].ToString();
                        mrow.hbl_ar_invnos = Dr["inv_no"].ToString();
                        mrow.hbl_ar_invamt = Lib.Conv2Decimal(Dr["inv_amt"].ToString());
                        mrow.hbl_ar_gstamt = Lib.Conv2Decimal(Dr["gst_amt"].ToString());
                        mrow.branch = Dr["branch"].ToString();
                        mrow.agent_created_date = Lib.DatetoStringDisplayformat(Dr["created_on"]);
                        mrow.created_date = Lib.DatetoStringDisplayformat(Dr["SI_created_date"]);
                        mList.Add(mrow);




                        tot_hbl_grwt += Lib.Conv2Decimal(mrow.hbl_grwt.ToString());
                        tot_hbl_chwt += Lib.Conv2Decimal(mrow.hbl_chwt.ToString());
                        tot_mbl_chwt += Lib.Conv2Decimal(mrow.mbl_chwt.ToString());
                        tot_mbl_grwt += Lib.Conv2Decimal(mrow.mbl_grwt.ToString());
                        tot_publish_rate += Lib.Conv2Decimal(mrow.publish_rate.ToString());
                        tot_informed_rate += Lib.Conv2Decimal(mrow.informed_rate.ToString());
                        tot_sell_informed += Lib.Conv2Decimal(mrow.sell_informed.ToString());
                        tot_rebate += Lib.Conv2Decimal(mrow.rebate.ToString());
                        tot_exwork += Lib.Conv2Decimal(mrow.exworks.ToString());

                    }
                    if (mList.Count > 1)
                    {
                        mrow = new MonthlyReport();
                        mrow.displayed = false;
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.sino = "TOTAL";
                        mrow.hbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_chwt.ToString(), 3));
                        mrow.hbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_grwt.ToString(), 3));
                        mrow.mbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_chwt.ToString(), 3));
                        mrow.mbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_grwt.ToString(), 3));
                        mrow.rebate = Lib.Conv2Decimal(Lib.NumericFormat(tot_rebate.ToString(), 2));
                        mrow.exworks = Lib.Conv2Decimal(Lib.NumericFormat(tot_exwork.ToString(), 2));
                        mList.Add(mrow);
                    }



                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintAirExportMonthlyReport();
                    }
                    Dt_List.Rows.Clear();
                }
                if (rec_category == "HBL-SE")
                {
                    sql = " SELECT h.hbl_pkid,h.hbl_type,h.hbl_no as SINO,h.rec_created_date as SI_created_date,";
                    sql += " m.hbl_folder_no as folder_no ,m.HBL_FOLDER_SENT_DATE as folder_sent,";
                    sql += " m.hbl_bl_no as mbl_no, m.hbl_date as mbl_date,";
                    sql += " m.hbl_terms as mbl_status,h.hbl_bl_no as hbl_no, ";
                    sql += " h.hbl_date as hbl_date ,h.hbl_terms as hbl_status,h.rec_branch_code as branch,";

                    sql += " h.hbl_ar_invnos as inv_no,h.hbl_ar_invamt as inv_amt,h.hbl_ar_gstamt as gst_amt,";
                    sql += " h.hbl_type as type,m.hbl_book_cntr_teu as teu,h.hbl_cbm as cbm,h.hbl_nature as terms,m.hbl_nature as mterms,";
                    sql += " h. hbl_book_cntr as book_cntr,h.hbl_job_nos as job_nos,h.hbl_ntwt as ntwt,";
                    sql += " m.hbl_shipment_type as shipment_type,";
                    sql += " shpr.cust_name as shipper_name,cons.cust_name as consignee_name,";
                    sql += " agent.cust_name as agent_name,agent.rec_created_date as created_on,";

                   // sql += " cons.cust_nomination as hbl_nomination,";
                    sql += " nvl(h.hbl_nomination,cons.cust_nomination) as hbl_nomination,";

                    sql += " carrier.param_name as carrier_name,pol.param_name as pol_name,";
                    sql += " pod.param_name as pod_name,pofd.param_name as pofd_name,";
                    sql += " m.hbl_pol_etd  as pol_etd,";

                   // sql += " nvl(sman2.param_name,sman.param_name) as sman_name,";//sman.param_name as sman_name
                    sql += " nvl(sman1.param_name,nvl(sman2.param_name,sman.param_name)) as sman_name,";
                    sql += " nvl(sman1.param_pkid,nvl(sman2.param_pkid,sman.param_pkid)) as sman_id,";

                    sql += " h.hbl_grwt as hbl_grwt, h.hbl_chwt as hbl_chwt,h.hbl_grwt as hbl_grwt, m.hbl_chwt as mbl_chwt,";
                    sql += " air_netnet as netnet,air_publish_rate  as publish_rate, ";
                    sql += " air_counter_informed as informed_rate,air_sell_informed as sell_informed,";
                    sql += " air_rebate as rebate, h.hbl_rebate_amt_inr as rebate_house, air_exworks as exworks,commodity.param_name as commodty_name";
                    sql += " from hblm h ";
                    sql += " left join hblm m on h.hbl_mbl_id = m.hbl_pkid";

                    sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";
                    sql += " left join custdet  cd on h.rec_branch_code = cd.det_branch_code and h.hbl_exp_id = cd.det_cust_id ";

                    sql += " left join param sman on shpr.cust_sman_id = sman.param_pkid";
                    sql += " left join param sman1 on h.hbl_salesman_id = sman1.param_pkid";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
                    sql += " left join customerm agent on h.hbl_agent_id = agent.cust_pkid";
                    sql += " left join param carrier on h.hbl_carrier_id = carrier.param_pkid";
                   
                    sql += " left join param pol on h.hbl_pol_id = pol.param_pkid";
                    sql += " left join param pod on h.hbl_pod_id = pod.param_pkid";
                    sql += " left join param pofd on h.hbl_pofd_id = pofd.param_pkid";
                    sql += " left join customerm cha on h.hbl_cha_id = cha.cust_pkid";
                    sql += " left join param commodity on m.hbl_commodity_id = commodity.param_pkid";
                    sql += " left join param vessel on m.hbl_vessel_id = vessel.param_pkid";
                    sql += " left join aircostm acostm on m.hbl_pkid = acostm.air_mblid";
                    sql += " where h.rec_category = 'SEA EXPORT' ";
                    if (!all)
                    {
                        sql += " and h.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += " and h.hbl_type ='HBL-SE'";

                    if (type_date == "SOB")
                        sql += "  and m.hbl_pol_etd between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    else
                        sql += "  and to_char(h.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    
                     if (agent_id.Length > 0)
                           sql += " and h.hbl_agent_id = '" + agent_id+ "' ";
                    if (shipper_id.Length > 0)
                        sql += " and h.hbl_exp_id = '"+shipper_id+"' ";
                    if (consignee_id.Length > 0)
                        sql += " and h.hbl_imp_id = '"+consignee_id+"'";
                    if (carrier_id.Length > 0)
                        sql += " and h.hbl_carrier_id = '"+carrier_id+"'";
                    if (pol_id.Length > 0)
                        sql += " and h.hbl_pol_id = '"+pol_id+"'";
                    if (pod_id.Length > 0)
                        sql += " and h.hbl_pod_id = '"+pod_id+"'";

                    if (type_date == "SOB")
                        sql += " order by h.rec_branch_code,m.hbl_date";
                    else
                        sql += " order by h.rec_branch_code,h.rec_created_date";


                    //  sql += "  and m.hbl_pol_etd between '{FDATE}' and '{EDATE}' ";
                    // sql += " order by   m.hbl_date";


                    sql = sql.Replace("{BRCODE}", branch_code);

                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();


                    tot_hbl_grwt = 0;
                    tot_hbl_chwt = 0;
                    tot_mbl_grwt = 0;
                    tot_mbl_chwt = 0;
                    tot_publish_rate = 0;
                    tot_informed_rate = 0;
                    tot_sell_informed = 0;
                    tot_rebate = 0;
                    tot_exwork = 0;

                    string pre_data = "";
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new MonthlyReport();
                        mrow.displayed = false;
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        if (pre_data != Dr["mbl_no"].ToString())
                        {
                            pre_data = Dr["mbl_no"].ToString();
                           
                            mrow.hbl_book_cntr_teu = Lib.Conv2Decimal(Dr["teu"].ToString());
                        }
                        else
                        {
                           // mrow.hbl_grwt = 0;
                          //  mrow.hbl_chwt = 0;
                          //  mrow.hbl_cbm = 0;
                            mrow.hbl_book_cntr_teu = 0;
                        //    mrow.hbl_ntwt = 0;

                        }

                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        mrow.hbl_ntwt = Lib.Conv2Decimal(Dr["ntwt"].ToString());
                        mrow.hbl_cbm = Lib.Conv2Decimal(Dr["cbm"].ToString());


                        if (Dr["shipment_type"].ToString() == "LCL")
                            mrow.hbl_book_cntr_teu = 0;
                        mrow.hbl_pkid = Dr["hbl_pkid"].ToString();
                        mrow.hbl_type = Dr["hbl_type"].ToString();
                        mrow.sino = Dr["SINO"].ToString();
                        mrow.folder_no = Dr["folder_no"].ToString();
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.folder_sent = Lib.DatetoStringDisplayformat(Dr["folder_sent"]);
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.mbl_status = Dr["mbl_status"].ToString();
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hbl_date"]);
                        mrow.hbl_status = Dr["hbl_status"].ToString();
                        mrow.shipper_name = Dr["shipper_name"].ToString();
                        mrow.consignee_name = Dr["consignee_name"].ToString();
                        mrow.agent_name = Dr["agent_name"].ToString();
                        mrow.hbl_nomination = Dr["hbl_nomination"].ToString();
                        mrow.carrier_name = Dr["carrier_name"].ToString();
                        mrow.pol_name = Dr["pol_name"].ToString();
                        mrow.pod_name = Dr["pod_name"].ToString();
                        mrow.pofd_name = Dr["pofd_name"].ToString();
                        mrow.pol_etd = Lib.DatetoStringDisplayformat(Dr["pol_etd"]);
                        mrow.sman_name = Dr["sman_name"].ToString();
                        mrow.sman_id = Dr["sman_id"].ToString();
                        mrow.commodty_name = Dr["commodty_name"].ToString();
                        mrow.hbl_type = Dr["type"].ToString();
                        mrow.hbl_nature = Dr["terms"].ToString();
                        mrow.mbl_nature = Dr["mterms"].ToString();
                        mrow.hbl_terms = Dr["hbl_status"].ToString();
                        mrow.mbl_terms = Dr["mbl_status"].ToString();
                        mrow.shipment_type = Dr["shipment_type"].ToString();
                        mrow.hbl_book_cntr = Dr["book_cntr"].ToString();
                        mrow.hbl_job_nos = Dr["job_nos"].ToString();
                        mrow.hbl_ar_invnos = Dr["inv_no"].ToString();
                        mrow.hbl_ar_invamt = Lib.Conv2Decimal(Dr["inv_amt"].ToString());
                        mrow.hbl_ar_gstamt = Lib.Conv2Decimal(Dr["gst_amt"].ToString());
                        mrow.branch = Dr["branch"].ToString();
                        mrow.agent_created_date = Lib.DatetoStringDisplayformat(Dr["created_on"]);
                        mrow.created_date = Lib.DatetoStringDisplayformat(Dr["SI_created_date"]);
                        mList.Add(mrow);
                        tot_mbl_grwt += Lib.Conv2Decimal(mrow.hbl_grwt.ToString());
                        tot_hbl_ntwt += Lib.Conv2Decimal(mrow.hbl_ntwt.ToString());
                        tot_cbm += Lib.Conv2Decimal(mrow.hbl_cbm.ToString());
                        tot_teu += Lib.Conv2Decimal(mrow.hbl_book_cntr_teu.ToString());


                    }
                    if (mList.Count > 1)
                    {
                        mrow = new MonthlyReport();
                        mrow.displayed = false;
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.sino = "TOTAL";
                        mrow.hbl_cbm = Lib.Conv2Decimal(Lib.NumericFormat(tot_cbm.ToString(), 3));
                        mrow.hbl_book_cntr_teu = Lib.Conv2Decimal(Lib.NumericFormat(tot_teu.ToString(), 3));
                        mList.Add(mrow);
                    }


                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintSeaExportMonthlyReport();
                    }
                    Dt_List.Rows.Clear();
                }
                if (rec_category == "HBL-AI")
                {
                    sql = " SELECT h.hbl_no as SINO,h.rec_created_date as SI_created_date,";
                    sql += " m.hbl_folder_no as folder_no ,m.HBL_FOLDER_SENT_DATE as folder_sent,";
                    sql += " m.hbl_bl_no as mbl_no, m.hbl_date as mbl_date,";
                    sql += " m.hbl_terms as mbl_status,h.hbl_bl_no as hbl_no, ";
                    sql += " h.hbl_date as hbl_date ,h.hbl_terms as hbl_status,h.rec_branch_code as branch,";

                    sql += " h.hbl_ar_invnos as inv_no,h.hbl_ar_invamt as inv_amt,h.hbl_ar_gstamt as gst_amt,";
                    sql += " shpr.cust_name as shipper_name,cons.cust_name as consignee_name,";
                    sql += " agent.cust_name as agent_name,agent.rec_created_date as created_on,";

                    sql += " shpr.cust_nomination as hbl_nomination,";
                    //sql += " nvl(h.hbl_nomination,shpr.cust_nomination) as hbl_nomination,";

                    sql += " carrier.param_name as carrier_name,pol.param_name as pol_name,";
                    sql += " pod.param_name as pod_name,pofd.param_name as pofd_name,";
                    sql += " m.hbl_pol_etd  as pol_etd,";

                    sql += " nvl(sman2.param_pkid,sman.param_pkid) as sman_id,";
                    sql += " nvl(sman2.param_name,sman.param_name) as sman_name,";// sman.param_name as sman_name,
                   // sql += " nvl(sman1.param_name,nvl(sman2.param_name,sman.param_name)) as sman_name,";

                    sql += " h.hbl_grwt as hbl_grwt, h.hbl_chwt as hbl_chwt,m.hbl_grwt as mbl_grwt, m.hbl_chwt as mbl_chwt,";
                    sql += " h.hbl_rebate_amt_inr as rebate_house, commodity.param_name as commodty_name";
                    sql += " from hblm h ";
                    sql += " left join hblm m on h.hbl_mbl_id = m.hbl_pkid";
                    sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";

                    sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
                    sql += " left join custdet cd on h.rec_branch_code = cd.det_branch_code and h.hbl_imp_id = cd.det_cust_id ";

                    sql += " left join param sman on cons.cust_sman_id = sman.param_pkid";
                  //  sql += " left join param sman1 on h.hbl_salesman_id = sman1.param_pkid";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += " left join customerm agent on m.hbl_agent_id = agent.cust_pkid";
                    sql += " left join param carrier on m.hbl_carrier_id = carrier.param_pkid";
                   
                    sql += " left join param pol on m.hbl_pol_id = pol.param_pkid";
                    sql += " left join param pod on m.hbl_pod_id = pod.param_pkid";
                    sql += " left join param pofd on m.hbl_pofd_id = pofd.param_pkid";
                    sql += " left join customerm cha on h.hbl_cha_id = cha.cust_pkid";
                    sql += " left join param commodity on m.hbl_commodity_id = commodity.param_pkid";
                    sql += " left join param vessel on m.hbl_vessel_id = vessel.param_pkid";
                    sql += " where h.rec_category = 'AIR IMPORT' ";
                    if (!all)
                    {
                        sql += " and h.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += " and h.hbl_type ='HBL-AI'";

                    if (type_date == "SOB")
                        sql += "  and m.hbl_pol_etd between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    else
                        sql += "  and to_char(h.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    if (agent_id.Length > 0)
                        sql += " and m.hbl_agent_id = '" + agent_id + "' ";
                    if (shipper_id.Length > 0)
                        sql += " and h.hbl_imp_id = '" + shipper_id + "' ";
                    if (consignee_id.Length > 0)
                        sql += " and h.hbl_exp_id = '" + consignee_id + "'";
                    if (carrier_id.Length > 0)
                        sql += " and m.hbl_carrier_id = '" + carrier_id + "'";
                    if (pol_id.Length > 0)
                        sql += " and m.hbl_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and m.hbl_pod_id = '" + pod_id + "'";


                    if (type_date == "SOB")
                        sql += " order by h.rec_branch_code,m.hbl_date";
                    else
                        sql += " order by h.rec_branch_code,h.rec_created_date";



                    sql = sql.Replace("{TYPES}", types);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{CATEGORY}", rec_category);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();


                    tot_hbl_grwt = 0;
                    tot_hbl_chwt = 0;
                    tot_mbl_grwt = 0;
                    tot_mbl_chwt = 0;
                    tot_publish_rate = 0;
                    tot_informed_rate = 0;
                    tot_sell_informed = 0;
                    tot_rebate = 0;
                    tot_exwork = 0;

                    string pre_data = "";
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new MonthlyReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        if (pre_data != Dr["mbl_no"].ToString())
                        {
                            pre_data = Dr["mbl_no"].ToString();
                            mrow.mbl_grwt = Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                            mrow.mbl_chwt = Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                        }
                        else
                        {
                            mrow.mbl_grwt = 0;
                            mrow.mbl_chwt = 0;
                        }
                        mrow.sino = Dr["SINO"].ToString();
                        mrow.folder_no = Dr["folder_no"].ToString();
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.folder_sent = Lib.DatetoStringDisplayformat(Dr["folder_sent"]);
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.mbl_status = Dr["mbl_status"].ToString();
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hbl_date"]);
                        mrow.hbl_status = Dr["hbl_status"].ToString();
                        mrow.shipper_name = Dr["shipper_name"].ToString();
                        mrow.consignee_name = Dr["consignee_name"].ToString();
                        mrow.agent_name = Dr["agent_name"].ToString();
                        mrow.hbl_nomination = Dr["hbl_nomination"].ToString();
                        mrow.carrier_name = Dr["carrier_name"].ToString();
                        mrow.pol_name = Dr["pol_name"].ToString();
                        mrow.pod_name = Dr["pod_name"].ToString();
                        mrow.pofd_name = Dr["pofd_name"].ToString();
                        mrow.pol_etd = Lib.DatetoStringDisplayformat(Dr["pol_etd"]);
                        mrow.sman_name = Dr["sman_name"].ToString();
                        mrow.sman_id = Dr["sman_id"].ToString();
                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        mrow.hbl_chwt = Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());
                        mrow.commodty_name = Dr["commodty_name"].ToString();
                        mrow.hbl_ar_invnos = Dr["inv_no"].ToString();
                        mrow.hbl_ar_invamt = Lib.Conv2Decimal(Dr["inv_amt"].ToString());
                        mrow.hbl_ar_gstamt = Lib.Conv2Decimal(Dr["gst_amt"].ToString());
                        mrow.branch = Dr["branch"].ToString();
                        mrow.agent_created_date = Lib.DatetoStringDisplayformat(Dr["created_on"]);
                        mrow.created_date = Lib.DatetoStringDisplayformat(Dr["SI_created_date"]);
                        mList.Add(mrow);




                        tot_hbl_grwt += Lib.Conv2Decimal(mrow.hbl_grwt.ToString());
                        tot_hbl_chwt += Lib.Conv2Decimal(mrow.hbl_chwt.ToString());
                        tot_mbl_chwt += Lib.Conv2Decimal(mrow.mbl_chwt.ToString());
                        tot_mbl_grwt += Lib.Conv2Decimal(mrow.mbl_grwt.ToString());


                    }
                    if (mList.Count > 1)
                    {
                        mrow = new MonthlyReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.sino = "TOTAL";
                        mrow.hbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_chwt.ToString(), 3));
                        mrow.hbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_hbl_grwt.ToString(), 3));
                        mrow.mbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_chwt.ToString(), 3));
                        mrow.mbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(tot_mbl_grwt.ToString(), 3));
                        //  mrow.rebate = Lib.Conv2Decimal(Lib.NumericFormat(tot_rebate.ToString(), 2));
                        //   mrow.exworks = Lib.Conv2Decimal(Lib.NumericFormat(tot_exwork.ToString(), 2));
                        mList.Add(mrow);
                    }



                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintAirImportMonthlyReport();
                    }
                    Dt_List.Rows.Clear();
                }
                if (rec_category == "HBL-SI")
                {
                    sql = " SELECT h.hbl_no as SINO,h.rec_created_date as SI_created_date,";
                    sql += " m.hbl_folder_no as folder_no ,m.HBL_FOLDER_SENT_DATE as folder_sent,";
                    sql += " m.hbl_bl_no as mbl_no, m.hbl_date as mbl_date,";
                    sql += " m.hbl_terms as mbl_status,h.hbl_bl_no as hbl_no, ";
                    sql += " h.hbl_date as hbl_date ,h.hbl_terms as hbl_status,h.rec_branch_code as branch,";

                    sql += " h.hbl_ar_invnos as inv_no,h.hbl_ar_invamt as inv_amt,h.hbl_ar_gstamt as gst_amt,";
                    sql += " h.hbl_type as type,m.hbl_book_cntr_teu as teu,h.hbl_cbm as cbm,h.hbl_nature as terms,m.hbl_nature as mterms,";
                    sql += " h. hbl_book_cntr as book_cntr,h.hbl_job_nos as job_nos,h.hbl_ntwt as ntwt,";
                    sql += " m.hbl_shipment_type as shipment_type,";
                    sql += " shpr.cust_name as shipper_name,cons.cust_name as consignee_name,";
                    sql += " agent.cust_name as agent_name,agent.rec_created_date as created_on,";

                    sql += " shpr.cust_nomination as hbl_nomination,";
                  //  sql += " nvl(h.hbl_nomination,shpr.cust_nomination) as hbl_nomination,";

                    sql += " carrier.param_name as carrier_name,pol.param_name as pol_name,";
                    sql += " pod.param_name as pod_name,pofd.param_name as pofd_name,";
                    sql += " m.hbl_pol_etd  as pol_etd,";

                    sql += " nvl(sman2.param_pkid,sman.param_pkid) as sman_id,";
                    sql += " nvl(sman2.param_name,sman.param_name) as sman_name,";// sman.param_name as sman_name,
                  //  sql += " nvl(sman1.param_name,nvl(sman2.param_name,sman.param_name)) as sman_name,";

                    sql += " h.hbl_grwt as hbl_grwt, h.hbl_chwt as hbl_chwt,h.hbl_grwt as hbl_grwt, m.hbl_chwt as mbl_chwt,";
                    sql += " air_netnet as netnet,air_publish_rate  as publish_rate, ";
                    sql += " air_counter_informed as informed_rate,air_sell_informed as sell_informed,";
                    sql += " air_rebate as rebate,  h.hbl_rebate_amt_inr as rebate_house,air_exworks as exworks,commodity.param_name as commodty_name";
                    sql += " from hblm h ";
                    sql += " left join hblm m on h.hbl_mbl_id = m.hbl_pkid";
                    sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";

                    sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
                    sql += " left join custdet cd on h.rec_branch_code = cd.det_branch_code and h.hbl_imp_id = cd.det_cust_id ";

                    sql += " left join param sman on cons.cust_sman_id = sman.param_pkid";
                  //  sql += " left join param sman1 on h.hbl_salesman_id = sman1.param_pkid";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += " left join customerm agent on m.hbl_agent_id = agent.cust_pkid";
                    sql += " left join param carrier on m.hbl_carrier_id = carrier.param_pkid";
                  
                    sql += " left join param pol on m.hbl_pol_id = pol.param_pkid";
                    sql += " left join param pod on m.hbl_pod_id = pod.param_pkid";
                    sql += " left join param pofd on m.hbl_pofd_id = pofd.param_pkid";
                    sql += " left join customerm cha on h.hbl_cha_id = cha.cust_pkid";
                    sql += " left join param commodity on m.hbl_commodity_id = commodity.param_pkid";
                    sql += " left join param vessel on m.hbl_vessel_id = vessel.param_pkid";
                    sql += " left join aircostm acostm on m.hbl_pkid = acostm.air_mblid";
                    sql += " where h.rec_category = 'SEA IMPORT' ";
                    if (!all)
                    {
                        sql += " and h.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += " and h.hbl_type ='HBL-SI'";


                    if (type_date == "SOB")
                        sql += "  and m.hbl_pol_etd between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                    else
                        sql += "  and to_char(h.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                    if (agent_id.Length > 0)
                        sql += " and m.hbl_agent_id = '" + agent_id + "' ";
                    if (shipper_id.Length > 0)
                        sql += " and h.hbl_imp_id = '" + shipper_id + "' ";
                    if (consignee_id.Length > 0)
                        sql += " and h.hbl_exp_id = '" + consignee_id + "'";
                    if (carrier_id.Length > 0)
                        sql += " and m.hbl_carrier_id = '" + carrier_id + "'";
                    if (pol_id.Length > 0)
                        sql += " and m.hbl_pol_id = '" + pol_id + "'";
                    if (pod_id.Length > 0)
                        sql += " and m.hbl_pod_id = '" + pod_id + "'";


                    if (type_date == "SOB")
                        sql += " order by  h.rec_branch_code,m.hbl_date";
                    else
                        sql += " order by h.rec_branch_code,h.rec_created_date";




                    sql = sql.Replace("{BRCODE}", branch_code);

                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();


                    tot_hbl_grwt = 0;
                    tot_hbl_chwt = 0;
                    tot_mbl_grwt = 0;
                    tot_mbl_chwt = 0;
                    tot_publish_rate = 0;
                    tot_informed_rate = 0;
                    tot_sell_informed = 0;
                    tot_rebate = 0;
                    tot_exwork = 0;

                    string pre_data = "";
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new MonthlyReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "BLACK";
                        if (pre_data != Dr["mbl_no"].ToString())
                        {
                            pre_data = Dr["mbl_no"].ToString();
                           
                            mrow.hbl_book_cntr_teu = Lib.Conv2Decimal(Dr["teu"].ToString());
                        }
                        else
                        {
                            //mrow.hbl_grwt = 0;
                            //mrow.hbl_chwt = 0;
                            //mrow.hbl_cbm = 0;
                            mrow.hbl_book_cntr_teu = 0;
                           // mrow.hbl_ntwt = 0;

                        }

                        mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                        mrow.hbl_ntwt = Lib.Conv2Decimal(Dr["ntwt"].ToString());
                        mrow.hbl_cbm = Lib.Conv2Decimal(Dr["cbm"].ToString());

                        if (Dr["shipment_type"].ToString() == "LCL")
                            mrow.hbl_book_cntr_teu = 0;
                        mrow.sino = Dr["SINO"].ToString();
                        mrow.folder_no = Dr["folder_no"].ToString();
                        mrow.mbl_no = Dr["mbl_no"].ToString();
                        mrow.folder_sent = Lib.DatetoStringDisplayformat(Dr["folder_sent"]);
                        mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                        mrow.mbl_status = Dr["mbl_status"].ToString();
                        mrow.hbl_no = Dr["hbl_no"].ToString();
                        mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hbl_date"]);
                        mrow.hbl_status = Dr["hbl_status"].ToString();
                        mrow.shipper_name = Dr["shipper_name"].ToString();
                        mrow.consignee_name = Dr["consignee_name"].ToString();
                        mrow.agent_name = Dr["agent_name"].ToString();
                        mrow.hbl_nomination = Dr["hbl_nomination"].ToString();
                        mrow.carrier_name = Dr["carrier_name"].ToString();
                        mrow.pol_name = Dr["pol_name"].ToString();
                        mrow.pod_name = Dr["pod_name"].ToString();
                        mrow.pofd_name = Dr["pofd_name"].ToString();
                        mrow.pol_etd = Lib.DatetoStringDisplayformat(Dr["pol_etd"]);
                        mrow.sman_name = Dr["sman_name"].ToString();
                        mrow.sman_id = Dr["sman_id"].ToString();
                        mrow.commodty_name = Dr["commodty_name"].ToString();
                        mrow.hbl_type = Dr["type"].ToString();
                        mrow.hbl_nature = Dr["terms"].ToString();
                        mrow.mbl_nature = Dr["mterms"].ToString();
                        mrow.hbl_terms = Dr["hbl_status"].ToString();
                        mrow.mbl_terms = Dr["mbl_status"].ToString();
                        mrow.shipment_type = Dr["shipment_type"].ToString();
                        mrow.hbl_book_cntr = Dr["book_cntr"].ToString();
                        mrow.hbl_job_nos = Dr["job_nos"].ToString();
                        mrow.hbl_ar_invnos = Dr["inv_no"].ToString();
                        mrow.hbl_ar_invamt = Lib.Conv2Decimal(Dr["inv_amt"].ToString());
                        mrow.hbl_ar_gstamt = Lib.Conv2Decimal(Dr["gst_amt"].ToString());
                        mrow.branch = Dr["branch"].ToString();
                        mrow.agent_created_date = Lib.DatetoStringDisplayformat(Dr["created_on"]);
                        mrow.created_date = Lib.DatetoStringDisplayformat(Dr["SI_created_date"]);
                        mList.Add(mrow);
                        tot_mbl_grwt += Lib.Conv2Decimal(mrow.hbl_grwt.ToString());
                        tot_hbl_ntwt += Lib.Conv2Decimal(mrow.hbl_ntwt.ToString());
                        tot_cbm += Lib.Conv2Decimal(mrow.hbl_cbm.ToString());
                        tot_teu += Lib.Conv2Decimal(mrow.hbl_book_cntr_teu.ToString());


                    }
                    if (mList.Count > 1)
                    {
                        mrow = new MonthlyReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "RED";
                        mrow.sino = "TOTAL";
                        mrow.hbl_cbm = Lib.Conv2Decimal(Lib.NumericFormat(tot_cbm.ToString(), 3));
                        mrow.hbl_book_cntr_teu = Lib.Conv2Decimal(Lib.NumericFormat(tot_teu.ToString(), 3));
                        mList.Add(mrow);
                    }



                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintSeaImportMonthlyReport();
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

        public IDictionary<string, object> UpdateMonReport(Dictionary<string, object> SearchData)
        {
            string pkid = SearchData["pkid"].ToString();
            string nomination = SearchData["nomination"].ToString();
            string old_nomination = SearchData["old_nomination"].ToString();
            string smanid = SearchData["smanid"].ToString();
            string smanname = SearchData["smanname"].ToString();
            string old_smanname = SearchData["old_smanname"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string user_code = SearchData["user_code"].ToString();
            string refno = SearchData["hblno"].ToString();
            string type = SearchData["type"].ToString();
            string periods ="";
            string shipper = "";
            string hblids = "";

            if (SearchData.ContainsKey("shipper"))
                shipper = SearchData["shipper"].ToString();

            if (SearchData.ContainsKey("periods"))
                periods = SearchData["periods"].ToString();

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            try
            {

                Con_Oracle = new DBConnection();

                if (type == "SALESMAN")
                    sql = "update hblm set hbl_salesman_id ='" + smanid + "' where hbl_pkid='" + pkid + "'";
                if (type == "NOMINATION")
                    sql = "update hblm set hbl_nomination ='" + nomination + "' where hbl_pkid='" + pkid + "'";

                if (type == "SALESMAN-ALL")
                {
                    hblids = pkid;
                    if (pkid.Contains(","))
                        pkid = pkid.Replace(",", "','");
                    sql = "update hblm set hbl_salesman_id ='" + smanid + "' where hbl_pkid in ('" + pkid + "')";
                }

                if (sql != "")
                {
                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                }
                Con_Oracle.CloseConnection();
                string remarks = "";
                remarks = "";
                if (type == "SALESMAN")
                    remarks += "SMAN-NEW: " + smanname + ", OLD:" + old_smanname;
                if (type == "NOMINATION")
                    remarks += "NOM-NEW: " + nomination + ", OLD:" + old_nomination;
                if (type == "SALESMAN-ALL")
                {
                    if (hblids.Contains(","))
                    {
                        string[] sdataIds = hblids.Split(',');
                        string[] sdataoldman = old_smanname.Split(',');
                        string[] sdatarefnos = refno.Split(',');
                        for (int i = 0; i < sdataIds.Length; i++)
                        {
                            remarks = "SMAN-ALL-NEW: " + smanname + ", OLD:" + sdataoldman[i].ToString() + ", SHPR:" + shipper + ", Dt: " + periods;
                            Lib.AuditLog("MONTHLY-REPORT", rowtype, "EDIT", company_code, branch_code, user_code, sdataIds[i], sdatarefnos[i], remarks);
                        }
                    }
                    else
                    {
                        remarks = "SMAN-ALL-NEW: " + smanname + ", OLD:" + old_smanname + ", SHPR:" + shipper + ", Dt: " + periods;
                        Lib.AuditLog("MONTHLY-REPORT", rowtype, "EDIT", company_code, branch_code, user_code, pkid, refno, remarks);
                    }
                }
                else
                    Lib.AuditLog("MONTHLY-REPORT", rowtype, "EDIT", company_code, branch_code, user_code, pkid, refno, remarks);
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
        private void PrintSeaImportMonthlyReport()
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
                    mSearchData.Add("branch_code","HOCPL");
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

                File_Display_Name = "MonthlyReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;
                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;
                WS.Columns[22].Width = 256 * 15;
                WS.Columns[23].Width = 256 * 15;
                WS.Columns[24].Width = 256 * 15;
                WS.Columns[25].Width = 256 * 15;
                WS.Columns[26].Width = 256 * 15;
                WS.Columns[27].Width = 256 * 20;
                WS.Columns[28].Width = 256 * 15;
                WS.Columns[29].Width = 256 * 15;
                WS.Columns[30].Width = 256 * 20;
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;

                iRow = 0; iCol = 1;
                if(all)
                {
                    WS.Columns[21].Style.NumberFormat = "#0.000";
                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[34].Style.NumberFormat = "#0.00";
                    WS.Columns[35].Style.NumberFormat = "#0.00";
                }
                else
                {
                    WS.Columns[20].Style.NumberFormat = "#0.000";
                    WS.Columns[21].Style.NumberFormat = "#0.000";
                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[33].Style.NumberFormat = "#0.00";
                    WS.Columns[34].Style.NumberFormat = "#0.00";
                }
               
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
                Lib.WriteData(WS, iRow, 1, "SEA-IMPORT MONTHLY REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER SENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT-CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL ETD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TEU", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "H-MOVEMENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "CONTAINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                foreach (MonthlyReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.folder_sent, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.shipper_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.agent_created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.carrier_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pol_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr_teu, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.commodty_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_gstamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_cbm, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr_teu, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }


        private void PrintSeaExportMonthlyReport()
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

                File_Display_Name = "MonthlyReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;
                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;
                WS.Columns[22].Width = 256 * 15;
                WS.Columns[23].Width = 256 * 15;
                WS.Columns[24].Width = 256 * 15;
                WS.Columns[25].Width = 256 * 15;
                WS.Columns[26].Width = 256 * 15;
                WS.Columns[27].Width = 256 * 15;
                WS.Columns[28].Width = 256 * 15;
                WS.Columns[29].Width = 256 * 15;
                WS.Columns[30].Width = 256 * 20;
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;

                iRow = 0; iCol = 1;
                if (all)
                {
                    WS.Columns[21].Style.NumberFormat = "#0.000";
                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[34].Style.NumberFormat = "#0.00";
                    WS.Columns[35].Style.NumberFormat = "#0.00";
                }
                else
                {
                    WS.Columns[20].Style.NumberFormat = "#0.000";
                    WS.Columns[21].Style.NumberFormat = "#0.000";
                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[33].Style.NumberFormat = "#0.00";
                    WS.Columns[34].Style.NumberFormat = "#0.00";
                }
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
                Lib.WriteData(WS, iRow, 1, "SEA-EXPORT MONTHLY REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER SENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT-CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL ETD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TEU", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-MOVEMENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-MOVEMENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JOB-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                foreach (MonthlyReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.folder_sent, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.shipper_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.agent_created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.carrier_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pol_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr_teu, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.shipment_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_job_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.commodty_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_gstamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_cbm, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr_teu, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }


        private void PrintAirImportMonthlyReport()
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



                File_Display_Name = "AIR MONTHLY REPORT.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;
                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;
                WS.Columns[22].Width = 256 * 15;
                WS.Columns[23].Width = 256 * 15;
                WS.Columns[24].Width = 256 * 15;
                WS.Columns[25].Width = 256 * 20;
                WS.Columns[26].Width = 256 * 15;
                WS.Columns[27].Width = 256 * 15;
                WS.Columns[28].Width = 256 * 15;
                WS.Columns[29].Width = 256 * 15;
                WS.Columns[30].Width = 256 * 15;
                WS.Columns[31].Width = 256 * 20;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                iRow = 0; iCol = 1;
                if(all)
                {
                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[25].Style.NumberFormat = "#0.000";
                    WS.Columns[27].Style.NumberFormat = "#0.00";
                    WS.Columns[28].Style.NumberFormat = "#0.00";
                    WS.Columns[29].Style.NumberFormat = "#0.00";
                }
                else
                {
                    WS.Columns[21].Style.NumberFormat = "#0.000";
                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[26].Style.NumberFormat = "#0.00";
                    WS.Columns[27].Style.NumberFormat = "#0.00";
                    WS.Columns[28].Style.NumberFormat = "#0.00";
                }
               



                iRow++;
                _Size = 14;
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
                Lib.WriteData(WS, iRow, 1, "AIR-IMPORT MONTHLY REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
                _Size = 10;
                iRow++;
                iRow++;

                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER SENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT-CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL ETD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                foreach (MonthlyReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }

                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.folder_sent, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.shipper_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.agent_created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.carrier_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pol_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.commodty_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_gstamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void PrintAirExportMonthlyReport()
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
                    mSearchData.Add("branch_code","HOCPL");
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



                File_Display_Name = "AIR MONTHLY REPORT.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;
                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;
                WS.Columns[22].Width = 256 * 15;
                WS.Columns[23].Width = 256 * 15;
                WS.Columns[24].Width = 256 * 15;
                WS.Columns[25].Width = 256 * 15;
                WS.Columns[26].Width = 256 * 15;
                WS.Columns[27].Width = 256 * 15;
                WS.Columns[28].Width = 256 * 15;
                WS.Columns[29].Width = 256 * 15;
                WS.Columns[30].Width = 256 * 15;
                WS.Columns[31].Width = 256 * 20;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                iRow = 0; iCol = 1;
                if (all)
                {
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[25].Style.NumberFormat = "#0.000";
                    WS.Columns[26].Style.NumberFormat = "#0.000";
                    WS.Columns[28].Style.NumberFormat = "#0.00";
                    WS.Columns[29].Style.NumberFormat = "#0.00";
                    WS.Columns[30].Style.NumberFormat = "#0.00";
                    WS.Columns[31].Style.NumberFormat = "#0.00";
                    WS.Columns[32].Style.NumberFormat = "#0.00";
                    WS.Columns[35].Style.NumberFormat = "#0.00";
                    WS.Columns[36].Style.NumberFormat = "#0.00";
                }
                else
                {

                    WS.Columns[22].Style.NumberFormat = "#0.000";
                    WS.Columns[23].Style.NumberFormat = "#0.000";
                    WS.Columns[24].Style.NumberFormat = "#0.000";
                    WS.Columns[25].Style.NumberFormat = "#0.000";
                    WS.Columns[27].Style.NumberFormat = "#0.00";
                    WS.Columns[28].Style.NumberFormat = "#0.00";
                    WS.Columns[29].Style.NumberFormat = "#0.00";
                    WS.Columns[30].Style.NumberFormat = "#0.00";
                    WS.Columns[31].Style.NumberFormat = "#0.00";
                    WS.Columns[34].Style.NumberFormat = "#0.00";
                    WS.Columns[35].Style.NumberFormat = "#0.00";
                }

                iRow++;
                _Size = 14;
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
                Lib.WriteData(WS, iRow, 1, "AIR-EXPORT MONTHLY REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
                _Size = 10;
                iRow++;
                iRow++;

                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER SENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT-CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL ETD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NETNET", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PUBLISH RATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INFORMED RATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL INFORMED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, "REBATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "EXWORK", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GST", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                foreach (MonthlyReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.folder_sent, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.shipper_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.agent_created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.carrier_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pol_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.netnet, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.publish_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.informed_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sell_informed, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        
                            Lib.WriteData(WS, iRow, iCol++, Rec.rebate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exworks, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.commodty_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_gstamt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.sino, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        
                            Lib.WriteData(WS, iRow, iCol++, Rec.rebate, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exworks, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                    }
                }
                iRow++;

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
