using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;


using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts

{
    public class ProfitService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<Profit> mList = new List<Profit>();
        Profit mrow;
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
        string branch_name = "";
        string year_code = "";
        string searchtype = "";
        string searchstring = "";
        string rec_category = "";
        string searchexpid = "";
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string code = "";
        string id = "";
        string acc_name = "";
        string MAWB = "";
        string ddp_ddu_exwork = "";
        string ErrorMessage = "";

        int finyear = 0;

        Boolean main_code = false;
        Boolean all = false;
        Boolean isnewformat = false;
        decimal tot_buy = 0;
        decimal tot_sell = 0;
        decimal tot_profit = 0;
        decimal tot_total = 0;
        decimal tot = 0;

        decimal _profit = 0;
        decimal _buy = 0;
        decimal _roi = 0;

        decimal tot_income = 0;
        decimal tot_expense = 0;

        decimal total = 0;
        Dictionary<int, string> DupDic = new Dictionary<int, string>();
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<Profit>();
            ErrorMessage = "";
            try
            {
                type = SearchData["type"].ToString();

                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
                branch_name = SearchData["branch_name"].ToString();

                type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();

                all = (Boolean)SearchData["all"];
                isnewformat = (Boolean)SearchData["isnewformat"];

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);

                finyear = Lib.Conv2Integer(SearchData["finyear"].ToString());



                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");

                //id = SearchData["acc_id"].ToString();
                //acc_name = SearchData["acc_name"].ToString();

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

                if (!isnewformat)
                {

                    if (type_date == "AIR-EXPORT-FORWARDING")
                    {
                        sql = " select a.*, sell - buy as profit from (";
                        sql += " select  mbl.hbl_date as mbl_date, mbl.hbl_no as mblslno,a.rec_branch_code as branch,";
                        sql += " mbl.hbl_bl_no as mblno, hbl.hbl_no as SINO,  hbl.hbl_bl_no as hblno,";
                        sql += " max(case when jvh_type = 'PN' then jvh_date else null end ) as buy_date,";
                        sql += " max(case when jvh_type = 'IN' then jvh_date else null end ) as sell_date,";
                        sql += " max( exp.cust_name) as exporter_name,";
                        sql += " max( imp.cust_name) as consignee_name,";
                        sql += " max( agent.cust_name) as agent_name,";
                        sql += " max(nvl(sman2.param_name,sman.param_name)) as sman_name,";
                        sql += " max( imp.cust_nomination) as nomination,";
                        sql += " max( mbl.hbl_terms) as mbl_frt_status,";
                        sql += " max( hbl.hbl_terms) as hbl_frt_status,";
                        sql += " max(mbl.hbl_chwt) as mbl_chwt,";
                        sql += " max(hbl.hbl_chwt) as hbl_chwt,";
                        sql += " max(mbl.hbl_grwt) as mbl_grwt,";
                        sql += " max(hbl.hbl_grwt) as hbl_grwt, ";
                        sql += " max(pol.param_name) as pol, ";
                        sql += " max(pod.param_name) as pod, ";
                        sql += " max(pofd.param_name) as pofd, ";

                        sql += " max(a.jvh_year) as fin_year, ";
                        sql += " max(mbl.hbl_folder_no) as mbl_folder_no, ";
                        sql += " max(hbl.hbl_ar_invnos) as inv_nos, ";
                        sql += " max(expadd.add_city) as shpr_location, ";
                        sql += " max(expstate.param_name) as shpr_state, ";
                        sql += " max(exp.rec_created_date) as exp_created,";
                        sql += " max(notify.bl_notify_name) as notify, ";
                        sql += " max(liner.param_name) as airline, ";
                        sql += " max(cmdty.param_name) as commodity,";

                        //sql += " max(orgcntry.param_name) as org_country, ";
                        sql += " max(podcntry.param_name) as pod_country, ";
                        //sql += " max(pofdcntry.param_name) as pofd_country, ";
                        sql += " max(buyer.cust_name) as buyer_name, ";

                        sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205001') and jvh_type not in('HO','IN-ES') then ABS(ct_amount) else 0 end ) as 	frt_dr,";
                        sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205002') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	fsc_dr,";
                        sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205003') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	wrs_dr,";
                        sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205017') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	mcc_dr,";
                        sql += " sum( case when jv_drcr = 'DR' and acc_code  like '1205%'  and acc_code not in ('1205001','1205002', '1205003', '1205017') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	oth_dr,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205001') and jvh_type  not in('HO','IN-ES')   then ABS(ct_amount) else 0 end ) as 	frt_cr,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205002') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	fsc_cr,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205003') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	wrs_cr,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205017') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	mcc_cr,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_code  like '1205%'  and acc_code not in ('1205001','1205002', '1205003', '1205017',  '1205010') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	oth_cr,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205010') and jvh_type  not in('HO','IN-ES')  then ct_amount else 0 end ) as 	margin_cr,";
                        sql += " sum( case when jvh_type in('HO','IN-ES') and jv_drcr = 'DR' and acc_code  in('1205001')  then ABS(ct_amount) else 0 end ) as 	frt_ho_dr,";
                        sql += " sum( case when jvh_type in('HO','IN-ES') and jv_drcr = 'CR' and acc_code  in('1205001')  then ABS(ct_amount) else 0 end ) as 	frt_ho_cr,";

                        sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                        sql += " sum( case when jv_drcr = 'DR' and acc_code  like '1205%'  then ABS(ct_amount) else 0 end ) as 	buy,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_code  like '1205%'  then ABS(ct_amount) else 0 end ) as 	sell";
                        sql += " from ledgerh a";
                        sql += " inner join ledgert b on  jvh_pkid = jv_parent_id ";
                        sql += " inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                        sql += " inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid";
                        sql += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                        sql += " inner join acctm e on jv_acc_id = acc_pkid";

                        sql += " left join customerm exp on hbl.hbl_exp_id = exp.cust_pkid";
                        sql += " left join custdet  cd on hbl.rec_branch_code = cd.det_branch_code and hbl.hbl_exp_id = cd.det_cust_id ";

                        sql += " left join param sman on exp.cust_sman_id = sman.param_pkid";
                        sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                        sql += " left join customerm imp on hbl.hbl_imp_id = imp.cust_pkid";

                        sql += " left join customerm agent on hbl.hbl_agent_id = agent.cust_pkid";

                        sql += " left join param pol on hbl.hbl_pol_id = pol.param_pkid";
                        sql += " left join param pod on hbl.hbl_pod_id = pod.param_pkid";
                        sql += " left join param pofd on hbl.hbl_pofd_id = pofd.param_pkid";

                        sql += " left join addressm expadd on hbl.hbl_exp_br_id = expadd.add_pkid";
                        sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid";
                        sql += " left join bl notify on hbl.hbl_pkid = notify.bl_pkid";
                        sql += " left join param liner on hbl.hbl_carrier_id = liner.param_pkid";
                        sql += " left join param cmdty on hbl.hbl_commodity_id = cmdty.param_pkid";

                        //sql += " left join param orgcntry on hbl.hbl_origin_country_id = orgcntry.param_pkid ";
                        sql += " left join param podcntry on hbl.hbl_pod_country_id = podcntry.param_pkid ";
                        //sql += " left join param pofdcntry on hbl.hbl_pofd_country_id = pofdcntry.param_pkid ";
                        sql += " left join customerm buyer on hbl.hbl_buyer_id = buyer.cust_pkid ";

                        sql += " where a.rec_company_code = '{COMPCODE}' ";

                        if (!all)
                        {
                            sql += " and a.rec_branch_code = '{BRCODE}' ";
                        }
                        sql += " and hbl.hbl_type = 'HBL-AE' and mbl.hbl_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')    ";
                        sql += " group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid, hbl.hbl_no, hbl.hbl_bl_no ";
                        sql += " ) a";
                        sql += " order by branch,mbl_date, mblno";

                        sql = sql.Replace("{BRCODE}", branch_code);
                        sql = sql.Replace("{COMPCODE}", company_code);
                        sql = sql.Replace("{FDATE}", from_date);
                        sql = sql.Replace("{EDATE}", to_date);

                        Con_Oracle = new DBConnection();
                        Dt_List = new DataTable();
                        Dt_List = Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CloseConnection();

                        tot_buy = 0;
                        tot_sell = 0;
                        tot_profit = 0;
                        tot_total = 0;
                        total = 0;


                        if (Dt_List.Rows.Count > 0)
                        {
                            MAWB = Dt_List.Rows[0]["mblno"].ToString();
                        }

                        foreach (DataRow Dr in Dt_List.Rows)
                        {

                            if (MAWB != Dr["mblno"].ToString())
                            {
                                if (mrow != null)
                                {
                                    mrow.total = total;
                                    tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                                    if (_buy > 0)
                                        mrow.roi = (_profit / _buy) * 100;
                                    _profit = 0; _buy = 0;


                                }
                                total = 0;
                                MAWB = Dr["mblno"].ToString();
                            }


                            mrow = new Profit();
                            mrow.rowtype = "DETAIL";
                            mrow.rowcolor = "BLACK";
                            mrow.branch = Dr["branch"].ToString();
                            mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                            mrow.mbl_no = Dr["mblslno"].ToString();
                            mrow.mbl_bl_no = Dr["mblno"].ToString();
                            mrow.hbl_no = Dr["SINO"].ToString();
                            mrow.hbl_bl_no = Dr["hblno"].ToString();
                            mrow.buy_date = Lib.DatetoStringDisplayformat(Dr["buy_date"]);
                            mrow.sell_date = Lib.DatetoStringDisplayformat(Dr["sell_date"]);
                            mrow.exporter = Dr["exporter_name"].ToString();
                            mrow.consignee = Dr["consignee_name"].ToString();
                            mrow.agent = Dr["agent_name"].ToString();
                            mrow.sman = Dr["sman_name"].ToString();
                            mrow.nomination = Dr["nomination"].ToString();
                            mrow.mbl_terms = Dr["mbl_frt_status"].ToString();
                            mrow.hbl_terms = Dr["hbl_frt_status"].ToString();

                            mrow.mbl_chwt = Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                            mrow.hbl_chwt = Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());
                            mrow.mbl_grwt = Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                            mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());
                            mrow.frt_dr = Lib.Conv2Decimal(Dr["frt_dr"].ToString());
                            mrow.fsc_dr = Lib.Conv2Decimal(Dr["fsc_dr"].ToString());
                            mrow.wrs_dr = Lib.Conv2Decimal(Dr["wrs_dr"].ToString());
                            mrow.mcc_dr = Lib.Conv2Decimal(Dr["mcc_dr"].ToString());
                            mrow.oth_dr = Lib.Conv2Decimal(Dr["oth_dr"].ToString());
                            mrow.frt_cr = Lib.Conv2Decimal(Dr["frt_cr"].ToString());
                            mrow.fsc_cr = Lib.Conv2Decimal(Dr["fsc_cr"].ToString());
                            mrow.wrs_cr = Lib.Conv2Decimal(Dr["wrs_cr"].ToString());
                            mrow.mcc_cr = Lib.Conv2Decimal(Dr["mcc_cr"].ToString());
                            mrow.oth_cr = Lib.Conv2Decimal(Dr["oth_cr"].ToString());
                            mrow.margin_cr = Lib.Conv2Decimal(Dr["margin_cr"].ToString());
                            mrow.frt_ho_dr = Lib.Conv2Decimal(Dr["frt_ho_dr"].ToString());
                            mrow.frt_ho_cr = Lib.Conv2Decimal(Dr["frt_ho_cr"].ToString());
                            mrow.rebate_dr = Lib.Conv2Decimal(Dr["rebate_dr"].ToString());

                            mrow.buy = Lib.Conv2Decimal(Dr["buy"].ToString());
                            mrow.sell = Lib.Conv2Decimal(Dr["sell"].ToString());
                            mrow.profit = Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.pol = Dr["pol"].ToString();
                            mrow.pod = Dr["pod"].ToString();
                            mrow.pofd = Dr["pofd"].ToString();

                            mrow.jvh_year = Dr["fin_year"].ToString();
                            mrow.mbl_folder_no = Dr["mbl_folder_no"].ToString();
                            mrow.hbl_ar_invnos = Dr["inv_nos"].ToString();
                            mrow.exp_city = Dr["shpr_location"].ToString();
                            mrow.exp_state = Dr["shpr_state"].ToString();
                            mrow.exp_created = Lib.DatetoStringDisplayformat(Dr["exp_created"]);
                            mrow.bl_notify_name = Dr["notify"].ToString();
                            mrow.mbl_liner = Dr["airline"].ToString();
                            mrow.mbl_commodity = Dr["commodity"].ToString();

                            //mrow.org_country = Dr["org_country"].ToString();
                            mrow.pod_country = Dr["pod_country"].ToString();
                            //mrow.pofd_country = Dr["pofd_country"].ToString();
                            mrow.buyer_name = Dr["buyer_name"].ToString();


                            total += Lib.Conv2Decimal(Dr["profit"].ToString());

                            _buy += Lib.Conv2Decimal(Dr["buy"].ToString());
                            _profit += Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.roi = 0;

                            mList.Add(mrow);

                            tot_buy += Lib.Conv2Decimal(mrow.buy.ToString());
                            tot_sell += Lib.Conv2Decimal(mrow.sell.ToString());
                            tot_profit += Lib.Conv2Decimal(mrow.profit.ToString());



                        }
                        if (mList.Count > 1)
                        {


                            mrow.total = total;
                            tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                            mrow = new Profit();
                            mrow.rowtype = "TOTAL";
                            mrow.rowcolor = "RED";
                            mrow.mbl_date = "TOTAL";
                            mrow.buy = Lib.Conv2Decimal(Lib.NumericFormat(tot_buy.ToString(), 2));
                            mrow.sell = Lib.Conv2Decimal(Lib.NumericFormat(tot_sell.ToString(), 2));
                            mrow.profit = Lib.Conv2Decimal(Lib.NumericFormat(tot_profit.ToString(), 2));
                            mrow.total = Lib.Conv2Decimal(Lib.NumericFormat(tot_total.ToString(), 2));

                            mList.Add(mrow);


                        }


                        if (type == "EXCEL")
                        {
                            if (mList != null)
                                PrintAirExportForwardingReport();
                        }
                        Dt_List.Rows.Clear();
                    }

                    if (type_date == "AIR-IMPORT")
                    {



                        sql = "  select a.*, sell - buy as profit from (";
                        sql += "  select  mbl.hbl_date as mawb_date, mbl.hbl_no as mblslno,hbl.hbl_date as hawb_date,a.rec_branch_code as branch,";
                        sql += "  a.jvh_year as fin_year,hbl.hbl_no as SINO, hbl.rec_created_date as si_date, carr.param_name as liner,";
                        sql += "  hbl.hbl_bl_no as hawb_no,mbl.hbl_bl_no as mawb_no,hbl.hbl_remarks as discription,";
                        sql += "  max(case when jvh_type = 'PN' then jvh_date else null end ) as buy_date,";
                        sql += "  max(case when jvh_type = 'IN' then jvh_date else null end ) as sell_date,";
                        sql += "  max( exp.cust_name) as exporter_name,";
                        sql += "  max( imp.cust_name) as consignee_name,max(impaddr.add_city) as imp_city,max(impstate.param_name) as imp_state,";
                        sql += "  max( agnt.cust_name) as agent,";
                        sql += "  max(nvl(sman2.param_name,sman.param_name)) as sman_name,";
                        sql += "  max( exp.cust_nomination) as nomination,";
                        sql += "  max( mbl.hbl_terms) as mbl_frt_status,";
                        sql += "  max( hbl.hbl_terms) as hbl_frt_status,";
                        sql += "  max(mbl.hbl_chwt) as mbl_chwt,";
                        sql += "  max(hbl.hbl_chwt) as hbl_chwt,";
                        sql += "  max(mbl.hbl_grwt) as mbl_grwt,";
                        sql += "  max(hbl.hbl_grwt) as hbl_grwt, ";
                        sql += " max(pol.param_name) as pol, ";
                        sql += " max(pod.param_name) as pod, ";
                        sql += " max(pofd.param_name) as pofd, ";


                        sql += " max(mbl.hbl_folder_no) as mbl_folder_no, ";
                        sql += " max(imp.rec_created_date) as imp_created, ";
                        sql += " max(notify.bl_notify_name) as notify, ";
                        sql += " max(hbl.hbl_ar_invnos) as inv_nos, ";
                        sql += " max(cntry.param_name) as orgin_country, ";
                        sql += " max(mbl.hbl_jobtype) as job_type, ";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1401') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1401,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1402') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1402,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1403') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1403,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1404') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1404,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1405') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1405,";

                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1401') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1401,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1402') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1402,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1403') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1403,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1404') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1404,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1405') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1405,";

                        sql += "  sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'DR' and acc_code  in('1405001')  then ABS(ct_amount) else 0 end ) as  cost_dr,";
                        sql += "  sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'CR' and acc_code  in('1405001')  then ABS(ct_amount) else 0 end ) as 	cost_cr,";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_code  in('1401100','1402100','1403100','1404100','1405100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code in('1401','1402','1403','1404','1405') then ABS(ct_amount) else 0 end ) as buy,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code in('1401','1402','1403','1404','1405')    then ABS(ct_amount) else 0 end ) as sell ";

                        sql += "  from ledgerh a";
                        sql += "  inner join ledgert b on  jvh_pkid = jv_parent_id ";
                        sql += "  inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                        sql += "  inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid";
                        sql += "  left join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                        sql += "  inner join acctm e on jv_acc_id = acc_pkid";
                        sql += "  left join customerm exp on hbl.hbl_exp_id = exp.cust_pkid";

                        sql += "  left join customerm imp on hbl.hbl_imp_id = imp.cust_pkid";
                        sql += "  left join custdet  cd on hbl.rec_branch_code = cd.det_branch_code and hbl.hbl_imp_id = cd.det_cust_id ";
                        sql += "  left join param sman on imp.cust_sman_id = sman.param_pkid";
                        sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                        sql += "  left join addressm impaddr on hbl.hbl_imp_br_id=impaddr.add_pkid";
                        sql += "  left join param impstate on impaddr.add_state_id=impstate.param_pkid";
                        sql += "  left join param carr on (hbl.hbl_carrier_id= carr.param_pkid)";

                        sql += "  left join customerm agnt on hbl.hbl_agent_id=agnt.cust_pkid";

                        sql += " left join param pol on mbl.hbl_pol_id = pol.param_pkid";
                        sql += " left join param pod on mbl.hbl_pod_id = pod.param_pkid";
                        sql += " left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid";

                        sql += " left join bl notify on hbl.hbl_pkid = notify.bl_pkid";
                        sql += " left join param cntry on hbl.hbl_origin_country_id = cntry.param_pkid";

                        sql += "  where a.rec_company_code = '{COMPCODE}' ";
                        if (!all)
                        {
                            sql += " and a.rec_branch_code = '{BRCODE}' ";
                        }
                        sql += " and hbl.hbl_type = 'HBL-AI' and to_char(hbl.rec_created_date,'DD-MON-YYYY') between  to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')   ";
                        sql += "  group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid, hbl.hbl_no, hbl.hbl_bl_no, hbl.rec_created_date,hbl.hbl_date,";
                        sql += "  hbl.hbl_remarks,carr.param_name,a.jvh_year ";
                        sql += "  ) a";
                        sql += "  order by branch,mawb_no,mawb_date ";



                        sql = sql.Replace("{BRCODE}", branch_code);
                        sql = sql.Replace("{COMPCODE}", company_code);
                        sql = sql.Replace("{FDATE}", from_date);
                        sql = sql.Replace("{EDATE}", to_date);

                        Con_Oracle = new DBConnection();
                        Dt_List = new DataTable();
                        Dt_List = Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CloseConnection();


                        tot_profit = 0;
                        tot_total = 0;
                        total = 0;
                        tot_income = 0;
                        tot_expense = 0;

                        if (Dt_List.Rows.Count > 0)
                        {
                            MAWB = Dt_List.Rows[0]["mawb_no"].ToString();
                        }

                        foreach (DataRow Dr in Dt_List.Rows)
                        {

                            if (MAWB != Dr["mawb_no"].ToString())
                            {
                                if (mrow != null)
                                {
                                    mrow.total = total;
                                    tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                                    if (_buy > 0)
                                        mrow.roi = (_profit / _buy) * 100;
                                    _profit = 0; _buy = 0;

                                }
                                total = 0;
                                MAWB = Dr["mawb_no"].ToString();
                            }


                            mrow = new Profit();
                            mrow.rowtype = "DETAIL";
                            mrow.rowcolor = "BLACK";

                            mrow.branch = Dr["branch"].ToString();
                            mrow.mbl_bl_no = Dr["mawb_no"].ToString();
                            mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mawb_date"]);
                            mrow.buy_date = Lib.DatetoStringDisplayformat(Dr["buy_date"]);
                            mrow.hbl_no = Dr["SINO"].ToString();
                            mrow.hbl_rec_creared_date = Lib.DatetoStringDisplayformat(Dr["si_date"]);
                            mrow.sell_date = Lib.DatetoStringDisplayformat(Dr["sell_date"]);
                            mrow.hbl_bl_no = Dr["hawb_no"].ToString();
                            mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hawb_date"]);
                            mrow.discription = Dr["discription"].ToString();
                            mrow.exporter = Dr["exporter_name"].ToString();
                            mrow.consignee = Dr["consignee_name"].ToString();
                            mrow.consignee_city = Dr["imp_city"].ToString();
                            mrow.consignee_state = Dr["imp_state"].ToString();
                            mrow.agent = Dr["agent"].ToString();
                            mrow.liner = Dr["liner"].ToString();
                            mrow.nomination = Dr["nomination"].ToString();
                            mrow.sman = Dr["sman_name"].ToString();
                            mrow.mbl_terms = Dr["mbl_frt_status"].ToString();
                            mrow.hbl_terms = Dr["hbl_frt_status"].ToString();

                            mrow.mbl_chwt = Lib.Conv2Decimal(Dr["mbl_chwt"].ToString());
                            mrow.mbl_grwt = Lib.Conv2Decimal(Dr["mbl_grwt"].ToString());
                            mrow.hbl_chwt = Lib.Conv2Decimal(Dr["hbl_chwt"].ToString());
                            mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());

                            mrow.ex_1401 = Lib.Conv2Decimal(Dr["ex_1401"].ToString());
                            mrow.ex_1402 = Lib.Conv2Decimal(Dr["ex_1402"].ToString());
                            mrow.ex_1403 = Lib.Conv2Decimal(Dr["ex_1403"].ToString());
                            mrow.ex_1404 = Lib.Conv2Decimal(Dr["ex_1404"].ToString());
                            mrow.ex_1405 = Lib.Conv2Decimal(Dr["ex_1405"].ToString());
                            mrow.in_1401 = Lib.Conv2Decimal(Dr["in_1401"].ToString());
                            mrow.in_1402 = Lib.Conv2Decimal(Dr["in_1402"].ToString());
                            mrow.in_1403 = Lib.Conv2Decimal(Dr["in_1403"].ToString());
                            mrow.in_1404 = Lib.Conv2Decimal(Dr["in_1404"].ToString());
                            mrow.in_1405 = Lib.Conv2Decimal(Dr["in_1405"].ToString());

                            mrow.cost_dr = Lib.Conv2Decimal(Dr["cost_dr"].ToString());
                            mrow.cost_cr = Lib.Conv2Decimal(Dr["cost_cr"].ToString());

                            mrow.rebate_dr = Lib.Conv2Decimal(Dr["rebate_dr"].ToString());

                            mrow.income = Lib.Conv2Decimal(Dr["sell"].ToString());
                            mrow.expense = Lib.Conv2Decimal(Dr["buy"].ToString());
                            mrow.profit = Lib.Conv2Decimal(Dr["profit"].ToString());

                            mrow.pol = Dr["pol"].ToString();
                            mrow.pod = Dr["pod"].ToString();
                            mrow.pofd = Dr["pofd"].ToString();

                            mrow.mbl_folder_no = Dr["mbl_folder_no"].ToString();
                            mrow.jvh_year = Dr["fin_year"].ToString();
                            mrow.exp_created = Lib.DatetoStringDisplayformat(Dr["imp_created"]);
                            mrow.bl_notify_name = Dr["notify"].ToString();
                            mrow.hbl_ar_invnos = Dr["inv_nos"].ToString();
                            mrow.hbl_orgin_country = Dr["orgin_country"].ToString();
                            mrow.mbl_jobtype = Dr["job_type"].ToString();

                            total += Lib.Conv2Decimal(Dr["profit"].ToString());

                            _buy += Lib.Conv2Decimal(Dr["buy"].ToString());
                            _profit += Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.roi = 0;


                            mList.Add(mrow);

                            tot_income += Lib.Conv2Decimal(mrow.income.ToString());
                            tot_expense += Lib.Conv2Decimal(mrow.expense.ToString());
                            tot_profit += Lib.Conv2Decimal(mrow.profit.ToString());



                        }
                        if (mList.Count > 1)
                        {


                            mrow.total = total;
                            tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                            mrow = new Profit();
                            mrow.rowtype = "TOTAL";
                            mrow.rowcolor = "RED";
                            mrow.mbl_bl_no = "TOTAL";
                            mrow.income = Lib.Conv2Decimal(Lib.NumericFormat(tot_income.ToString(), 2));
                            mrow.expense = Lib.Conv2Decimal(Lib.NumericFormat(tot_expense.ToString(), 2));
                            mrow.profit = Lib.Conv2Decimal(Lib.NumericFormat(tot_profit.ToString(), 2));
                            mrow.total = Lib.Conv2Decimal(Lib.NumericFormat(tot_total.ToString(), 2));

                            mList.Add(mrow);


                        }


                        if (type == "EXCEL")
                        {
                            if (mList != null)
                                PrintAirImportReport();
                        }
                        Dt_List.Rows.Clear();
                    }

                    if (type_date == "SEA-EXPORT-FORWARDING")
                    {

                        sql = " select a.*,  ";
                        sql += " exp.cust_name as exporter_name,";
                        sql += " imp.cust_name as consignee_name, ";
                        sql += " agent.cust_name as agent_name,";
                        sql += " nvl(sman2.param_name,sman.param_name) as sman_name, ";
                        sql += " imp.cust_nomination as nomination,";
                        sql += " pol.param_name as pol,";
                        sql += " pod.param_name as pod,";
                        sql += " pofd.param_name as pofd,";
                        sql += " podcntry.param_name as pod_country,";
                        sql += " buyer.cust_name as buyer_name,  ";
                        sql += " status.param_name as mbl_status,";
                        sql += " expadd.add_city as shpr_location,";
                        sql += " expstate.param_name as shpr_state,  ";
                        sql += " exp.rec_created_date as shpr_created, ";
                        sql += " notify.bl_notify_name as notify, ";
                        sql += " liner.param_name as liner, ";
                        sql += " cmdty.param_name as commodity,  ";
                        sql += " sell - buy as profit from ( ";

                        sql += " select  mbl.hbl_date as mbl_date, mbl.hbl_no as mblslno,a.rec_branch_code as branch, ";
                        sql += " mbl.hbl_bl_no as mblno, hbl.hbl_no as SINO,  hbl.hbl_bl_no as hblno,";
                        sql += " hbl.hbl_pkid, ";
                        sql += " max(case when jvh_type = 'PN' then jvh_date else null end ) as buy_date,";
                        sql += " max(case when jvh_type = 'IN' then jvh_date else null end ) as sell_date,";
                        sql += " max(hbl.hbl_exp_id) as hbl_exp_id, ";
                        sql += " max(hbl.hbl_exp_br_id) as hbl_exp_br_id, ";
                        sql += " max(hbl.hbl_imp_id) as hbl_imp_id, ";
                        sql += " max(hbl.hbl_agent_id) as hbl_agent_id,";
                        sql += " max(hbl.rec_branch_code) as hbl_rec_branch_code, ";
                        sql += " max(mbl.hbl_terms) as mbl_frt_status,";
                        sql += " max(hbl.hbl_terms) as hbl_frt_status,";
                        sql += " max(hbl.hbl_cbm) as hbl_cbm,";
                        sql += " max(hbl.hbl_grwt) as hbl_grwt,";
                        sql += " max(hbl.hbl_pol_id) as hbl_pol_id,";
                        sql += " max(hbl.hbl_pod_id) as hbl_pod_id,";
                        sql += " max(hbl.hbl_pofd_id) as hbl_pofd_id,";
                        sql += " max(hbl.hbl_pod_country_id) as hbl_pod_country_id,";
                        sql += " max(hbl.hbl_buyer_id) as hbl_buyer_id,  ";
                        sql += " max(a.jvh_year) as fin_year, ";
                        sql += " max(mbl.hbl_folder_no) as mbl_folder_no,";
                        sql += " max(mbl.hbl_pol_etd) as etd,";
                        sql += " max(hbl.hbl_book_cntr) as cntr,";
                        sql += " max(mbl.hbl_status_id) as mbl_status_id,";
                        sql += " max(hbl.hbl_book_cntr_teu) as teu,  ";
                        sql += " max(hbl.hbl_date) as hbl_date, ";
                        sql += " max(hbl.hbl_ar_invnos) as inv_nos,  ";
                        sql += " max(hbl.hbl_carrier_id) as hbl_carrier_id, ";
                        sql += " max(hbl.hbl_commodity_id) as hbl_commodity_id,  ";
                        sql += " max(hbl.hbl_ddp) as ddp,  ";
                        sql += " max(hbl.hbl_ddu) as ddu,";
                        sql += " max(hbl.hbl_ex_works) as ex_works,";
                        sql += " max(mbl.hbl_shipment_type) as shipment_type, ";
                        sql += " sum( case when jv_drcr = 'DR' and acc_main_code  in('1105') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as ex_1105,";
                        sql += " sum( case when jv_drcr = 'DR' and acc_main_code  in('1106') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as ex_1106, ";
                        sql += " sum( case when jv_drcr = 'DR' and acc_main_code  in('1107') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as ex_1107, ";
                        sql += " sum( case when jv_drcr = 'CR' and acc_main_code  in('1105') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as in_1105,";
                        sql += " sum( case when jv_drcr = 'CR' and acc_main_code  in('1106') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as in_1106, ";
                        sql += " sum( case when jv_drcr = 'CR' and acc_main_code  in('1107') and jvh_type  not in('HO','IN-ES')   then ABS(ct_amount) else 0 end ) as in_1107,";
                        sql += " sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'DR' and acc_main_code  in('1105')  then ABS(ct_amount) else 0 end ) as  cost_dr, ";
                        sql += " sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'CR' and acc_main_code  in('1105')  then ABS(ct_amount) else 0 end ) as cost_cr,";
                        sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1105100','1106100','1107100')  then ABS(ct_amount) else 0 end ) as rebate_dr,";
                        sql += " sum( case when jv_drcr = 'DR' and acc_main_code in('1105','1106','1107') then ABS(ct_amount) else 0 end ) as buy, ";
                        sql += " sum( case when jv_drcr = 'CR' and acc_main_code in('1105','1106','1107')    then ABS(ct_amount) else 0 end ) as sell     ";
                        sql += " from ledgerh a ";
                        sql += " inner join ledgert b on  jvh_pkid = jv_parent_id   ";
                        sql += " inner join costcentert c on b.jv_pkid = c.ct_jv_id  ";
                        sql += " inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid  ";
                        sql += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid  ";
                        sql += " inner join acctm e on jv_acc_id = acc_pkid  ";
                        sql += "  where a.rec_company_code = '{COMPCODE}' ";
                        if (!all)
                        {
                            sql += " and a.rec_branch_code = '{BRCODE}' ";
                        }
                        sql += " and hbl.hbl_type = 'HBL-SE' and mbl.hbl_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')   ";
                        sql += " group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid, hbl.hbl_no, hbl.hbl_bl_no  ";

                        sql += " ) a ";
                        sql += " left join customerm exp on a.hbl_exp_id = exp.cust_pkid  ";
                        sql += " left join custdet  cd on a.hbl_rec_branch_code = cd.det_branch_code and a.hbl_exp_id = cd.det_cust_id   ";
                        sql += " left join param sman on exp.cust_sman_id = sman.param_pkid  ";
                        sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid  ";
                        sql += " left join customerm imp on a.hbl_imp_id = imp.cust_pkid  ";
                        sql += " left join customerm agent on a.hbl_agent_id = agent.cust_pkid ";
                        sql += " left join param pol on a.hbl_pol_id = pol.param_pkid ";
                        sql += " left join param pod on a.hbl_pod_id = pod.param_pkid ";
                        sql += " left join param pofd on a.hbl_pofd_id = pofd.param_pkid ";
                        sql += " left join param podcntry on a.hbl_pod_country_id = podcntry.param_pkid  ";
                        sql += " left join customerm buyer on a.hbl_buyer_id = buyer.cust_pkid  ";
                        sql += " left join param status on a.mbl_status_id = status.param_pkid ";
                        sql += " left join addressm expadd on a.hbl_exp_br_id = expadd.add_pkid ";
                        sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid ";
                        sql += " left join bl notify on a.hbl_pkid = notify.bl_pkid ";
                        sql += " left join param liner on a.hbl_carrier_id = liner.param_pkid ";
                        sql += " left join param cmdty on a.hbl_commodity_id = cmdty.param_pkid ";
                        sql += " order by branch,mbl_date, mblno";


                        sql = sql.Replace("{BRCODE}", branch_code);
                        sql = sql.Replace("{COMPCODE}", company_code);
                        sql = sql.Replace("{FDATE}", from_date);
                        sql = sql.Replace("{EDATE}", to_date);

                        Con_Oracle = new DBConnection();
                        Dt_List = new DataTable();
                        Dt_List = Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CloseConnection();

                        tot_buy = 0;
                        tot_sell = 0;
                        tot_profit = 0;
                        tot_total = 0;
                        total = 0;

                        DupDic = new Dictionary<int, string>();

                        if (Dt_List.Rows.Count > 0)
                        {
                            MAWB = Dt_List.Rows[0]["mblno"].ToString();
                        }

                        foreach (DataRow Dr in Dt_List.Rows)
                        {

                            if (MAWB != Dr["mblno"].ToString())
                            {
                                if (mrow != null)
                                {
                                    mrow.total = total;

                                    tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                                    if (_buy > 0)
                                        mrow.roi = (_profit / _buy) * 100;
                                    _profit = 0; _buy = 0;


                                }
                                total = 0;
                                MAWB = Dr["mblno"].ToString();
                            }


                            mrow = new Profit();
                            ddp_ddu_exwork = "";
                            mrow.rowtype = "DETAIL";
                            mrow.rowcolor = "BLACK";
                            mrow.branch = Dr["branch"].ToString();
                            mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                            mrow.mbl_no = Dr["mblslno"].ToString();
                            mrow.mbl_bl_no = Dr["mblno"].ToString();
                            mrow.hbl_no = Dr["SINO"].ToString();
                            mrow.hbl_bl_no = Dr["hblno"].ToString();
                            mrow.buy_date = Lib.DatetoStringDisplayformat(Dr["buy_date"]);
                            mrow.sell_date = Lib.DatetoStringDisplayformat(Dr["sell_date"]);
                            mrow.exporter = Dr["exporter_name"].ToString();
                            mrow.consignee = Dr["consignee_name"].ToString();
                            mrow.agent = Dr["agent_name"].ToString();

                            mrow.sman = Dr["sman_name"].ToString();
                            mrow.nomination = Dr["nomination"].ToString();
                            mrow.mbl_terms = Dr["mbl_frt_status"].ToString();
                            mrow.hbl_terms = Dr["hbl_frt_status"].ToString();


                            mrow.hbl_cbm = Lib.Conv2Decimal(Dr["hbl_cbm"].ToString());

                            mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());

                            mrow.ex_1105 = Lib.Conv2Decimal(Dr["ex_1105"].ToString());
                            mrow.ex_1106 = Lib.Conv2Decimal(Dr["ex_1106"].ToString());
                            mrow.ex_1107 = Lib.Conv2Decimal(Dr["ex_1107"].ToString());

                            mrow.in_1105 = Lib.Conv2Decimal(Dr["in_1105"].ToString());
                            mrow.in_1106 = Lib.Conv2Decimal(Dr["in_1106"].ToString());
                            mrow.in_1107 = Lib.Conv2Decimal(Dr["in_1107"].ToString());

                            mrow.cost_dr = Lib.Conv2Decimal(Dr["cost_dr"].ToString());
                            mrow.cost_cr = Lib.Conv2Decimal(Dr["cost_cr"].ToString());

                            mrow.rebate_dr = Lib.Conv2Decimal(Dr["rebate_dr"].ToString());

                            mrow.buy = Lib.Conv2Decimal(Dr["buy"].ToString());
                            mrow.sell = Lib.Conv2Decimal(Dr["sell"].ToString());
                            mrow.profit = Lib.Conv2Decimal(Dr["profit"].ToString());

                            mrow.pol = Dr["pol"].ToString();
                            mrow.pod = Dr["pod"].ToString();
                            mrow.pofd = Dr["pofd"].ToString();

                            mrow.jvh_year = Dr["fin_year"].ToString();
                            mrow.mbl_folder_no = Dr["mbl_folder_no"].ToString();
                            mrow.mbl_pol_etd = Lib.DatetoStringDisplayformat(Dr["etd"]);
                            mrow.hbl_book_cntr = Dr["cntr"].ToString();
                            mrow.mbl_status = Dr["mbl_status"].ToString();

                            if (Dr["shipment_type"].ToString() == "CONSOLE" || Dr["shipment_type"].ToString() == "BUYERS CONSOLE")
                            {
                                if (DupDic.ContainsValue(Dr["mblslno"].ToString()))
                                {
                                    mrow.hbl_book_cntr_teu = 0;
                                }
                                else
                                {
                                    DupDic.Add(DupDic.Count, Dr["mblslno"].ToString());
                                    mrow.hbl_book_cntr_teu = Lib.Conv2Decimal(Dr["teu"].ToString());
                                }
                            }
                            else if (Dr["shipment_type"].ToString() == "LCL")
                                mrow.hbl_book_cntr_teu = 0;
                            else
                                mrow.hbl_book_cntr_teu = Lib.Conv2Decimal(Dr["teu"].ToString());

                            mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hbl_date"]);
                            mrow.hbl_ar_invnos = Dr["inv_nos"].ToString();
                            mrow.exp_city = Dr["shpr_location"].ToString();
                            mrow.exp_state = Dr["shpr_state"].ToString();
                            mrow.exp_created = Lib.DatetoStringDisplayformat(Dr["shpr_created"]);
                            mrow.bl_notify_name = Dr["notify"].ToString();
                            mrow.mbl_liner = Dr["liner"].ToString();
                            mrow.mbl_commodity = Dr["commodity"].ToString();
                            mrow.pod_country = Dr["pod_country"].ToString();
                            mrow.buyer_name = Dr["buyer_name"].ToString();
                            if (Dr["ddp"].ToString() != "")
                                ddp_ddu_exwork = Dr["ddp"].ToString();
                            if (Dr["ddu"].ToString() != "")
                                ddp_ddu_exwork += " / ";
                            ddp_ddu_exwork += Dr["ddu"].ToString();
                            if (Dr["ex_works"].ToString() != "")
                                ddp_ddu_exwork += " / ";
                            ddp_ddu_exwork += Dr["ex_works"].ToString();

                            mrow.hbl_ddp_ddu_exwork = ddp_ddu_exwork;
                            mrow.mbl_shipment_type = Dr["shipment_type"].ToString();

                            total += Lib.Conv2Decimal(Dr["profit"].ToString());

                            _buy += Lib.Conv2Decimal(Dr["buy"].ToString());
                            _profit += Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.roi = 0;


                            mList.Add(mrow);

                            tot_buy += Lib.Conv2Decimal(mrow.buy.ToString());
                            tot_sell += Lib.Conv2Decimal(mrow.sell.ToString());
                            tot_profit += Lib.Conv2Decimal(mrow.profit.ToString());



                        }
                        if (mList.Count > 1)
                        {


                            mrow.total = total;
                            tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                            mrow = new Profit();
                            mrow.rowtype = "TOTAL";
                            mrow.rowcolor = "RED";
                            mrow.mbl_date = "TOTAL";
                            mrow.buy = Lib.Conv2Decimal(Lib.NumericFormat(tot_buy.ToString(), 2));
                            mrow.sell = Lib.Conv2Decimal(Lib.NumericFormat(tot_sell.ToString(), 2));
                            mrow.profit = Lib.Conv2Decimal(Lib.NumericFormat(tot_profit.ToString(), 2));
                            mrow.total = Lib.Conv2Decimal(Lib.NumericFormat(tot_total.ToString(), 2));

                            mList.Add(mrow);


                        }


                        if (type == "EXCEL")
                        {
                            if (mList != null)
                                PrintSeaExportForwardingReport();
                        }
                        Dt_List.Rows.Clear();
                    }

                    if (type_date == "SEA-IMPORT")
                    {

                        sql = " select a.*, sell - buy as profit from (";
                        sql += "   select  mbl.hbl_date as mbl_date, mbl.hbl_no as mblslno,hbl.hbl_date as hbl_date,a.rec_branch_code as branch,";
                        sql += "   a.jvh_year as fin_year,hbl.hbl_no as SINO, hbl.rec_created_date as si_date, carr.param_name as liner,";
                        sql += "   hbl.hbl_bl_no as hbl_no,mbl.hbl_bl_no as mbl_no,hbl.hbl_remarks as discription,";
                        sql += "   max(case when jvh_type = 'PN' then jvh_date else null end ) as buy_date,";
                        sql += "   max(case when jvh_type = 'IN' then jvh_date else null end ) as sell_date,";
                        sql += "   max( exp.cust_name) as exporter_name,";
                        sql += "   max( imp.cust_name) as consignee_name,max(impaddr.add_city) as imp_city,max(impstate.param_name) as imp_state,";
                        sql += "   max( agnt.cust_name) as agent,";
                        sql += "   max(nvl(sman2.param_name,sman.param_name)) as sman_name,";
                        sql += "   max( exp.cust_nomination) as nomination,";
                        sql += "   max( mbl.hbl_terms) as mbl_frt_status,";
                        sql += "   max( hbl.hbl_terms) as hbl_frt_status,";
                        sql += "   max(hbl.hbl_cbm) as hbl_cbm,";
                        sql += "   max(hbl.hbl_ntwt) as hbl_ntwt,";
                        sql += "   max(hbl.hbl_grwt) as hbl_grwt, ";
                        sql += " max(pol.param_name) as pol, ";
                        sql += " max(pod.param_name) as pod, ";
                        sql += " max(pofd.param_name) as pofd, ";

                        sql += " max(mbl.hbl_folder_no) as mbl_folder_no, ";
                        sql += " max(hbl.hbl_bl_no) as bl_no, ";
                        sql += " max(notify.bl_notify_name) as notify, ";
                        sql += " max(status.param_name) as mbl_status, ";
                        sql += " max(hbl.hbl_ar_invnos) as inv_nos, ";
                        sql += " max(cntry.param_name) as orgin_country, ";
                        sql += " max(hbl.hbl_book_cntr) as cntr, ";
                        sql += " max(hbl.hbl_book_cntr_teu) as teu, ";
                        sql += " max(hbl.hbl_nature) as nature, ";
                        sql += " max(mbl.hbl_jobtype) as job_type, ";
                        sql += " max(imp.rec_created_date) as imp_created, ";

                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1301') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1301,";
                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1302') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1302,";
                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1303') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1303,";
                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1304') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1304,";
                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1305') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1305,";
                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1306') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1306,";
                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1307') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1307,";

                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1301') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1301,";
                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1302') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1302,";
                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1303') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1303,";
                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1304') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1304,";
                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1305') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1305,";
                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1306') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1306,";
                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1307') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1307,";

                        sql += "   sum( case when jvh_type  in('HO','IN-ES') ' and jv_drcr = 'DR' and acc_main_code  in('1305')  then ABS(ct_amount) else 0 end ) as  cost_dr,";
                        sql += "   sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'CR' and acc_main_code  in('1305')  then ABS(ct_amount) else 0 end ) as 	cost_cr,";

                        sql += "   sum( case when jv_drcr = 'DR' and acc_code  in('1301100','1302100','1303100','1304100','1305100','1306100','1307100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                        sql += "   sum( case when jv_drcr = 'DR' and acc_main_code in('1301','1302','1303','1304','1305','1306','1307') then ABS(ct_amount) else 0 end ) as buy,";
                        sql += "   sum( case when jv_drcr = 'CR' and acc_main_code in('1301','1302','1303','1304','1305','1306','1307')    then ABS(ct_amount) else 0 end ) as sell";



                        sql += "   from ledgerh a";
                        sql += "   inner join ledgert b on  jvh_pkid = jv_parent_id ";
                        sql += "   inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                        sql += "   inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid";
                        sql += "   left join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                        sql += "   inner join acctm e on jv_acc_id = acc_pkid";
                        sql += "   left join customerm exp on hbl.hbl_exp_id = exp.cust_pkid";
                        sql += "   left join customerm imp on hbl.hbl_imp_id = imp.cust_pkid";
                        sql += "   left join custdet  cd on hbl.rec_branch_code = cd.det_branch_code and hbl.hbl_imp_id = cd.det_cust_id ";
                        sql += "   left join param sman on imp.cust_sman_id = sman.param_pkid";
                        sql += "   left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                        sql += "   left join addressm impaddr on hbl.hbl_imp_br_id=impaddr.add_pkid";
                        sql += "   left join param impstate on impaddr.add_state_id=impstate.param_pkid";
                        sql += "   left join param carr on hbl.hbl_carrier_id= carr.param_pkid";
                        sql += "   left join customerm agnt on hbl.hbl_agent_id=agnt.cust_pkid";

                        sql += " left join param pol on mbl.hbl_pol_id = pol.param_pkid";
                        sql += " left join param pod on mbl.hbl_pod_id = pod.param_pkid";
                        sql += " left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid";

                        sql += " left join bl notify on hbl.hbl_pkid = notify.bl_pkid";
                        sql += " left join param status on mbl.hbl_status_id = status.param_pkid";
                        sql += " left join param cntry on hbl.hbl_origin_country_id = cntry.param_pkid";

                        sql += "   where a.rec_company_code = '{COMPCODE}' ";

                        if (!all)
                        {
                            sql += "   and a.rec_branch_code = '{BRCODE}' ";
                        }

                        sql += "   and hbl.hbl_type = 'HBL-SI' and to_char(hbl.rec_created_date,'DD-MON-YYYY') between  to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";
                        sql += "   group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid,hbl.rec_created_date, hbl.hbl_no,hbl.hbl_date,";
                        sql += "  hbl.hbl_bl_no,hbl.hbl_remarks,carr.param_name,a.jvh_year";
                        sql += "   ) a";
                        sql += "   order by branch,mbl_no,mbl_date";


                        sql = sql.Replace("{BRCODE}", branch_code);
                        sql = sql.Replace("{COMPCODE}", company_code);
                        sql = sql.Replace("{FDATE}", from_date);
                        sql = sql.Replace("{EDATE}", to_date);

                        Con_Oracle = new DBConnection();
                        Dt_List = new DataTable();
                        Dt_List = Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CloseConnection();


                        tot_profit = 0;
                        tot_total = 0;
                        total = 0;
                        tot_income = 0;
                        tot_expense = 0;

                        if (Dt_List.Rows.Count > 0)
                        {
                            MAWB = Dt_List.Rows[0]["mbl_no"].ToString();
                        }

                        foreach (DataRow Dr in Dt_List.Rows)
                        {

                            if (MAWB != Dr["mbl_no"].ToString())
                            {
                                if (mrow != null)
                                {
                                    mrow.total = total;
                                    tot_total += Lib.Conv2Decimal(mrow.total.ToString());



                                    if (_buy > 0)
                                        mrow.roi = (_profit / _buy) * 100;
                                    _profit = 0; _buy = 0;


                                }
                                total = 0;
                                MAWB = Dr["mbl_no"].ToString();
                            }


                            mrow = new Profit();
                            mrow.rowtype = "DETAIL";
                            mrow.rowcolor = "BLACK";

                            mrow.mbl_bl_no = Dr["mbl_no"].ToString();
                            mrow.branch = Dr["branch"].ToString();
                            mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                            mrow.buy_date = Lib.DatetoStringDisplayformat(Dr["buy_date"]);
                            mrow.hbl_no = Dr["SINO"].ToString();
                            mrow.hbl_rec_creared_date = Lib.DatetoStringDisplayformat(Dr["si_date"]);
                            mrow.sell_date = Lib.DatetoStringDisplayformat(Dr["sell_date"]);
                            mrow.hbl_bl_no = Dr["hbl_no"].ToString();
                            mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["hbl_date"]);
                            mrow.discription = Dr["discription"].ToString();
                            mrow.exporter = Dr["exporter_name"].ToString();
                            mrow.consignee = Dr["consignee_name"].ToString();
                            mrow.consignee_city = Dr["imp_city"].ToString();
                            mrow.consignee_state = Dr["imp_state"].ToString();
                            mrow.agent = Dr["agent"].ToString();
                            mrow.liner = Dr["liner"].ToString();
                            mrow.nomination = Dr["nomination"].ToString();
                            mrow.sman = Dr["sman_name"].ToString();
                            mrow.mbl_terms = Dr["mbl_frt_status"].ToString();
                            mrow.hbl_terms = Dr["hbl_frt_status"].ToString();

                            mrow.hbl_cbm = Lib.Conv2Decimal(Dr["hbl_cbm"].ToString());
                            mrow.hbl_ntwt = Lib.Conv2Decimal(Dr["hbl_ntwt"].ToString());
                            mrow.hbl_grwt = Lib.Conv2Decimal(Dr["hbl_grwt"].ToString());

                            mrow.ex_1301 = Lib.Conv2Decimal(Dr["ex_1301"].ToString());
                            mrow.ex_1302 = Lib.Conv2Decimal(Dr["ex_1302"].ToString());
                            mrow.ex_1303 = Lib.Conv2Decimal(Dr["ex_1303"].ToString());
                            mrow.ex_1304 = Lib.Conv2Decimal(Dr["ex_1304"].ToString());
                            mrow.ex_1305 = Lib.Conv2Decimal(Dr["ex_1305"].ToString());
                            mrow.ex_1306 = Lib.Conv2Decimal(Dr["ex_1306"].ToString());
                            mrow.ex_1307 = Lib.Conv2Decimal(Dr["ex_1307"].ToString());

                            mrow.in_1301 = Lib.Conv2Decimal(Dr["in_1301"].ToString());
                            mrow.in_1302 = Lib.Conv2Decimal(Dr["in_1302"].ToString());
                            mrow.in_1303 = Lib.Conv2Decimal(Dr["in_1303"].ToString());
                            mrow.in_1304 = Lib.Conv2Decimal(Dr["in_1304"].ToString());
                            mrow.in_1305 = Lib.Conv2Decimal(Dr["in_1305"].ToString());
                            mrow.in_1306 = Lib.Conv2Decimal(Dr["in_1306"].ToString());
                            mrow.in_1307 = Lib.Conv2Decimal(Dr["in_1307"].ToString());


                            mrow.cost_dr = Lib.Conv2Decimal(Dr["cost_dr"].ToString());
                            mrow.cost_cr = Lib.Conv2Decimal(Dr["cost_cr"].ToString());

                            mrow.rebate_dr = Lib.Conv2Decimal(Dr["rebate_dr"].ToString());

                            mrow.expense = Lib.Conv2Decimal(Dr["buy"].ToString());
                            mrow.income = Lib.Conv2Decimal(Dr["sell"].ToString());

                            mrow.profit = Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.pol = Dr["pol"].ToString();
                            mrow.pod = Dr["pod"].ToString();
                            mrow.pofd = Dr["pofd"].ToString();

                            mrow.mbl_folder_no = Dr["mbl_folder_no"].ToString();
                            mrow.jvh_year = Dr["fin_year"].ToString();
                            mrow.hbl_bl_no = Dr["bl_no"].ToString();
                            mrow.exp_created = Lib.DatetoStringDisplayformat(Dr["imp_created"]);
                            mrow.bl_notify_name = Dr["notify"].ToString();
                            mrow.mbl_status = Dr["mbl_status"].ToString();
                            mrow.hbl_ar_invnos = Dr["inv_nos"].ToString();
                            mrow.hbl_orgin_country = Dr["orgin_country"].ToString();
                            mrow.hbl_book_cntr = Dr["cntr"].ToString();
                            mrow.hbl_book_cntr_teu = Lib.Conv2Decimal(Dr["teu"].ToString());
                            mrow.hbl_nature = Dr["nature"].ToString();
                            mrow.mbl_jobtype = Dr["job_type"].ToString();


                            total += Lib.Conv2Decimal(Dr["profit"].ToString());

                            _buy += Lib.Conv2Decimal(Dr["buy"].ToString());
                            _profit += Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.roi = 0;


                            mList.Add(mrow);

                            tot_income += Lib.Conv2Decimal(mrow.income.ToString());
                            tot_expense += Lib.Conv2Decimal(mrow.expense.ToString());
                            tot_profit += Lib.Conv2Decimal(mrow.profit.ToString());



                        }
                        if (mList.Count > 1)
                        {


                            mrow.total = total;
                            tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                            mrow = new Profit();
                            mrow.rowtype = "TOTAL";
                            mrow.rowcolor = "RED";
                            mrow.mbl_bl_no = "TOTAL";
                            mrow.income = Lib.Conv2Decimal(Lib.NumericFormat(tot_income.ToString(), 2));
                            mrow.expense = Lib.Conv2Decimal(Lib.NumericFormat(tot_expense.ToString(), 2));
                            mrow.profit = Lib.Conv2Decimal(Lib.NumericFormat(tot_profit.ToString(), 2));
                            mrow.total = Lib.Conv2Decimal(Lib.NumericFormat(tot_total.ToString(), 2));

                            mList.Add(mrow);


                        }


                        if (type == "EXCEL")
                        {
                            if (mList != null)
                                PrintSeaImportReport();
                        }
                        Dt_List.Rows.Clear();
                    }

                    if (type_date == "SEA-EXPORT-CLEARING")
                    {


                        sql = "  select a.*, sell - buy as profit from (";
                        sql += "  select j.job_docno as job_no,j.job_date as job_date,a.rec_branch_code as branch,";
                        sql += "  max(case when jv_drcr = 'DR' then jvh_date else null end ) as buy_date,";
                        sql += "  max(case when jv_drcr = 'CR' then jvh_date else null end ) as sell_date,";
                        sql += "  max( exp.cust_name) as exporter_name,";
                        sql += "  max( imp.cust_name) as consignee_name,";
                        sql += "  max(nvl(sman2.param_name,sman.param_name)) as sman_name,";
                        sql += "  max( imp.cust_nomination) as nomination,";
                        sql += "  max( j.job_terms) as job_frt_status,";
                        sql += "  max(j.job_ntwt) as job_ntwt,";
                        sql += "  max(j.job_grwt) as job_grwt,";
                        sql += "  max(j.job_cbm) as job_cbm,";
                        sql += "  max(j.job_chwt) as job_chwt, ";
                        sql += "  max(j.job_type) as job_type, ";
                        sql += " max(pol.param_name) as pol, ";
                        sql += " max(pod.param_name) as pod, ";
                        sql += " max(pofd.param_name) as pofd, ";

                        sql += " max(podcntry.param_name) as pod_country, ";
                        sql += " max(buyer.cust_name) as buyer_name, ";

                        sql += "  max(agent.cust_name) as agent, ";

                        sql += "  max(expadd.add_city) as shpr_location, ";
                        sql += "  max(notify.bl_notify_name) as notify, ";
                        sql += "  max(expstate.param_name) as shpr_state, ";
                        sql += "  max(exp.rec_created_date) as shpr_created, ";
                        sql += "  max(cmdty.param_name) as commodity, ";
                        sql += "  max(h.hbl_ar_invnos) as job_invoice_nos, ";
                        sql += "  max(j.job_cntr_type) as job_cntr_type, ";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1101')   then ABS(ct_amount) else 0 end ) as 	ex_1101,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1102')   then ABS(ct_amount) else 0 end ) as 	ex_1102,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1103')   then ABS(ct_amount) else 0 end ) as 	ex_1103,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1101')   then ABS(ct_amount) else 0 end ) as 	in_1101,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1102')   then ABS(ct_amount) else 0 end ) as 	in_1102,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1103')   then ABS(ct_amount) else 0 end ) as 	in_1103,";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_code  in('1101100','1102100','1103100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code in('1101','1102','1103')   then ABS(ct_amount) else 0 end ) as 	buy,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code in('1101','1102','1103')  then ABS(ct_amount) else 0 end ) as 	sell";



                        sql += "  from ledgerh a";
                        sql += "  inner join ledgert b on  jvh_pkid = jv_parent_id ";
                        sql += "  inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                        sql += "  inner join jobm j on c.ct_cost_id = j.job_pkid";
                        sql += "  inner join acctm e on jv_acc_id = acc_pkid";
                        sql += "  left join customerm exp on j.job_exp_id = exp.cust_pkid";
                        sql += "  left join custdet  cd on j.rec_branch_code = cd.det_branch_code and j.job_exp_id = cd.det_cust_id ";
                        sql += "  left join param sman on exp.cust_sman_id = sman.param_pkid";
                        sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                        sql += "  left join customerm imp on j.job_imp_id = imp.cust_pkid";

                        sql += " left join param pol on j.job_pol_id = pol.param_pkid";
                        sql += " left join param pod on j.job_pod_id = pod.param_pkid";
                        sql += " left join param pofd on j.job_pofd_id = pofd.param_pkid";

                        sql += " left  join hblm h on j.jobs_hbl_id = h.hbl_pkid";
                        sql += " left join customerm agent on h.hbl_agent_id = agent.cust_pkid";
                        sql += " left join addressm expadd on j.job_exp_br_id = expadd.add_pkid";
                        sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid";
                        sql += " left join bl notify on h.hbl_pkid = notify.bl_pkid";
                        sql += " left join param cmdty on j.job_commodity_id = cmdty.param_pkid";
                        sql += " left join param podcntry on h.hbl_pod_country_id = podcntry.param_pkid ";
                        sql += " left join customerm buyer on h.hbl_buyer_id = buyer.cust_pkid ";

                        sql += "  where a.rec_company_code = '{COMPCODE}' ";
                        if (!all)
                        {
                            sql += "  and a.rec_branch_code = '{BRCODE}'";
                        }

                        sql += "  and j.rec_category = 'SEA EXPORT' and j.job_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')   ";
                        sql += "  group by a.rec_branch_code,j.job_date,j.job_docno ";
                        sql += "  ) a";
                        sql += "  order by branch,job_no,job_date";



                        sql = sql.Replace("{BRCODE}", branch_code);
                        sql = sql.Replace("{COMPCODE}", company_code);
                        sql = sql.Replace("{FDATE}", from_date);
                        sql = sql.Replace("{EDATE}", to_date);

                        Con_Oracle = new DBConnection();
                        Dt_List = new DataTable();
                        Dt_List = Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CloseConnection();


                        tot_profit = 0;
                        tot_total = 0;
                        total = 0;
                        tot_income = 0;
                        tot_expense = 0;

                        if (Dt_List.Rows.Count > 0)
                        {
                            MAWB = Dt_List.Rows[0]["job_no"].ToString();
                        }

                        foreach (DataRow Dr in Dt_List.Rows)
                        {

                            if (MAWB != Dr["job_no"].ToString())
                            {
                                if (mrow != null)
                                {
                                    mrow.total = total;
                                    tot_total += Lib.Conv2Decimal(mrow.total.ToString());



                                    if (_buy > 0)
                                        mrow.roi = (_profit / _buy) * 100;
                                    _profit = 0; _buy = 0;


                                }
                                total = 0;
                                MAWB = Dr["job_no"].ToString();
                            }


                            mrow = new Profit();
                            mrow.rowtype = "DETAIL";
                            mrow.rowcolor = "BLACK";

                            mrow.job_no = Dr["job_no"].ToString();
                            mrow.job_date = Lib.DatetoStringDisplayformat(Dr["job_date"]);
                            mrow.branch = Dr["branch"].ToString();
                            mrow.buy_date = Lib.DatetoStringDisplayformat(Dr["buy_date"]);
                            mrow.sell_date = Lib.DatetoStringDisplayformat(Dr["sell_date"]);
                            mrow.exporter = Dr["exporter_name"].ToString();
                            mrow.consignee = Dr["consignee_name"].ToString();
                            mrow.sman = Dr["sman_name"].ToString();
                            mrow.nomination = Dr["nomination"].ToString();
                            mrow.job_terms = Dr["job_frt_status"].ToString();
                            mrow.job_chwt = Lib.Conv2Decimal(Dr["job_chwt"].ToString());
                            mrow.job_ntwt = Lib.Conv2Decimal(Dr["job_ntwt"].ToString());
                            mrow.job_cbm = Lib.Conv2Decimal(Dr["job_cbm"].ToString());
                            mrow.job_grwt = Lib.Conv2Decimal(Dr["job_grwt"].ToString());
                            mrow.ex_1101 = Lib.Conv2Decimal(Dr["ex_1101"].ToString());
                            mrow.ex_1102 = Lib.Conv2Decimal(Dr["ex_1102"].ToString());
                            mrow.ex_1103 = Lib.Conv2Decimal(Dr["ex_1103"].ToString());
                            mrow.in_1101 = Lib.Conv2Decimal(Dr["in_1101"].ToString());
                            mrow.in_1102 = Lib.Conv2Decimal(Dr["in_1102"].ToString());
                            mrow.in_1103 = Lib.Conv2Decimal(Dr["in_1103"].ToString());

                            mrow.rebate_dr = Lib.Conv2Decimal(Dr["rebate_dr"].ToString());

                            mrow.expense = Lib.Conv2Decimal(Dr["buy"].ToString());
                            mrow.income = Lib.Conv2Decimal(Dr["sell"].ToString());
                            mrow.profit = Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.pol = Dr["pol"].ToString();
                            mrow.pod = Dr["pod"].ToString();
                            mrow.pofd = Dr["pofd"].ToString();
                            mrow.agent = Dr["agent"].ToString();
                            mrow.job_type = Dr["job_type"].ToString();

                            mrow.bl_notify_name = Dr["notify"].ToString();
                            mrow.exp_city = Dr["shpr_location"].ToString();
                            mrow.exp_state = Dr["shpr_state"].ToString();
                            mrow.exp_created = Lib.DatetoStringDisplayformat(Dr["shpr_created"]);
                            mrow.job_commodity = Dr["commodity"].ToString();
                            mrow.job_invoice_nos = Dr["job_invoice_nos"].ToString();
                            mrow.job_cntr_type = Dr["job_cntr_type"].ToString();
                            mrow.pod_country = Dr["pod_country"].ToString();
                            mrow.buyer_name = Dr["buyer_name"].ToString();

                            total += Lib.Conv2Decimal(Dr["profit"].ToString());


                            _buy += Lib.Conv2Decimal(Dr["buy"].ToString());
                            _profit += Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.roi = 0;



                            mList.Add(mrow);

                            tot_income += Lib.Conv2Decimal(mrow.income.ToString());
                            tot_expense += Lib.Conv2Decimal(mrow.expense.ToString());
                            tot_profit += Lib.Conv2Decimal(mrow.profit.ToString());



                        }
                        if (mList.Count > 1)
                        {


                            mrow.total = total;
                            tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                            mrow = new Profit();
                            mrow.rowtype = "TOTAL";
                            mrow.rowcolor = "RED";
                            mrow.mbl_bl_no = "TOTAL";
                            mrow.income = Lib.Conv2Decimal(Lib.NumericFormat(tot_income.ToString(), 2));
                            mrow.expense = Lib.Conv2Decimal(Lib.NumericFormat(tot_expense.ToString(), 2));
                            mrow.profit = Lib.Conv2Decimal(Lib.NumericFormat(tot_profit.ToString(), 2));
                            mrow.total = Lib.Conv2Decimal(Lib.NumericFormat(tot_total.ToString(), 2));

                            mList.Add(mrow);


                        }


                        if (type == "EXCEL")
                        {
                            if (mList != null)
                                PrintSeaExportClearingReport();
                        }
                        Dt_List.Rows.Clear();
                    }

                    if (type_date == "AIR-EXPORT-CLEARING")
                    {
                        sql = "  select a.*, sell - buy as profit from (";
                        sql += "  select j.job_docno as job_no,j.job_date as job_date,a.rec_branch_code as branch,";
                        sql += "  max(case when jv_drcr = 'DR' then jvh_date else null end ) as buy_date,";
                        sql += "  max(case when jv_drcr = 'CR' then jvh_date else null end ) as sell_date,";
                        sql += "  max( exp.cust_name) as exporter_name,";
                        sql += "  max( imp.cust_name) as consignee_name,";
                        sql += "  max(nvl(sman2.param_name,sman.param_name)) as sman_name,";
                        sql += "  max( imp.cust_nomination) as nomination,";
                        sql += "  max( j.job_terms) as job_frt_status,";
                        sql += "  max(j.job_ntwt) as job_ntwt,";
                        sql += "  max(j.job_grwt) as job_grwt,";
                        sql += "  max(j.job_cbm) as job_cbm,";
                        sql += "  max(j.job_chwt) as job_chwt, ";
                        sql += " max(pol.param_name) as pol, ";
                        sql += " max(pod.param_name) as pod, ";
                        sql += " max(pofd.param_name) as pofd, ";
                        sql += " max(podcntry.param_name) as pod_country, ";
                        sql += " max(buyer.cust_name) as buyer_name, ";

                        sql += " max(agent.cust_name) as agent, ";
                        sql += " max(j.job_type) as job_type, ";
                        sql += " max(notify.bl_notify_name) as notify, ";
                        sql += " max(expadd.add_city) as shpr_location, ";
                        sql += " max(expstate.param_name) as shpr_state, ";
                        sql += " max(exp.rec_created_date) as shpr_created, ";
                        sql += " max(cmdty.param_name) as commodity, ";
                        sql += " max(h.hbl_ar_invnos) as job_invoice_nos,";


                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1201')   then ABS(ct_amount) else 0 end ) as 	ex_1201,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1202')   then ABS(ct_amount) else 0 end ) as 	ex_1202,";
                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1203')   then ABS(ct_amount) else 0 end ) as 	ex_1203,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1201')   then ABS(ct_amount) else 0 end ) as 	in_1201,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1202')   then ABS(ct_amount) else 0 end ) as 	in_1202,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1203')   then ABS(ct_amount) else 0 end ) as 	in_1203,";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_code  in('1201100','1202100','1203100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                        sql += "  sum( case when jv_drcr = 'DR' and acc_main_code in('1201','1202','1203')   then ABS(ct_amount) else 0 end ) as 	buy,";
                        sql += "  sum( case when jv_drcr = 'CR' and acc_main_code in('1201','1202','1203')  then ABS(ct_amount) else 0 end ) as 	sell";


                        sql += "  from ledgerh a";
                        sql += "  inner join ledgert b on  jvh_pkid = jv_parent_id ";
                        sql += "  inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                        sql += "  inner join jobm j on c.ct_cost_id = j.job_pkid";
                        sql += "  inner join acctm e on jv_acc_id = acc_pkid";
                        sql += "  left join customerm exp on j.job_exp_id = exp.cust_pkid";
                        sql += "  left join custdet  cd on j.rec_branch_code = cd.det_branch_code and j.job_exp_id = cd.det_cust_id ";
                        sql += "  left join param sman on exp.cust_sman_id = sman.param_pkid";
                        sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                        sql += "  left join customerm imp on j.job_imp_id = imp.cust_pkid";

                        sql += " left join param pol on j.job_pol_id = pol.param_pkid";
                        sql += " left join param pod on j.job_pod_id = pod.param_pkid";
                        sql += " left join param pofd on j.job_pofd_id = pofd.param_pkid";

                        sql += " left  join hblm h on j.jobs_hbl_id = h.hbl_pkid";
                        sql += " left join customerm agent on h.hbl_agent_id = agent.cust_pkid";
                        sql += " left join bl notify on h.hbl_pkid = notify.bl_pkid";
                        sql += " left join addressm expadd on j.job_exp_br_id = expadd.add_pkid";
                        sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid ";
                        sql += " left join param cmdty on j.job_commodity_id = cmdty.param_pkid";
                        sql += " left join param podcntry on h.hbl_pod_country_id = podcntry.param_pkid ";
                        sql += " left join customerm buyer on h.hbl_buyer_id = buyer.cust_pkid ";

                        sql += "  where a.rec_company_code = '{COMPCODE}' ";
                        if (!all)
                        {
                            sql += "  and a.rec_branch_code = '{BRCODE}'";
                        }

                        sql += "  and j.rec_category = 'AIR EXPORT' and j.job_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')   ";
                        sql += "  group by a.rec_branch_code,j.job_date,j.job_docno";

                        sql += "  ) a";
                        sql += "  order by branch,job_no,job_date";



                        sql = sql.Replace("{BRCODE}", branch_code);
                        sql = sql.Replace("{COMPCODE}", company_code);
                        sql = sql.Replace("{FDATE}", from_date);
                        sql = sql.Replace("{EDATE}", to_date);

                        Con_Oracle = new DBConnection();
                        Dt_List = new DataTable();
                        Dt_List = Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CloseConnection();


                        tot_profit = 0;
                        tot_total = 0;
                        total = 0;
                        tot_income = 0;
                        tot_expense = 0;

                        if (Dt_List.Rows.Count > 0)
                        {
                            MAWB = Dt_List.Rows[0]["job_no"].ToString();
                        }

                        foreach (DataRow Dr in Dt_List.Rows)
                        {

                            if (MAWB != Dr["job_no"].ToString())
                            {
                                if (mrow != null)
                                {
                                    mrow.total = total;
                                    tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                                    if (_buy > 0)
                                        mrow.roi = (_profit / _buy) * 100;
                                    _profit = 0; _buy = 0;



                                }
                                total = 0;
                                MAWB = Dr["job_no"].ToString();
                            }


                            mrow = new Profit();
                            mrow.rowtype = "DETAIL";
                            mrow.rowcolor = "BLACK";

                            mrow.job_no = Dr["job_no"].ToString();
                            mrow.job_date = Lib.DatetoStringDisplayformat(Dr["job_date"]);
                            mrow.branch = Dr["branch"].ToString();
                            mrow.buy_date = Lib.DatetoStringDisplayformat(Dr["buy_date"]);
                            mrow.sell_date = Lib.DatetoStringDisplayformat(Dr["sell_date"]);
                            mrow.exporter = Dr["exporter_name"].ToString();
                            mrow.consignee = Dr["consignee_name"].ToString();
                            mrow.sman = Dr["sman_name"].ToString();
                            mrow.nomination = Dr["nomination"].ToString();
                            mrow.job_terms = Dr["job_frt_status"].ToString();
                            mrow.job_chwt = Lib.Conv2Decimal(Dr["job_chwt"].ToString());
                            mrow.job_ntwt = Lib.Conv2Decimal(Dr["job_ntwt"].ToString());
                            mrow.job_cbm = Lib.Conv2Decimal(Dr["job_cbm"].ToString());
                            mrow.job_grwt = Lib.Conv2Decimal(Dr["job_grwt"].ToString());
                            mrow.ex_1201 = Lib.Conv2Decimal(Dr["ex_1201"].ToString());
                            mrow.ex_1202 = Lib.Conv2Decimal(Dr["ex_1202"].ToString());
                            mrow.ex_1203 = Lib.Conv2Decimal(Dr["ex_1203"].ToString());
                            mrow.in_1201 = Lib.Conv2Decimal(Dr["in_1201"].ToString());
                            mrow.in_1202 = Lib.Conv2Decimal(Dr["in_1202"].ToString());
                            mrow.in_1203 = Lib.Conv2Decimal(Dr["in_1203"].ToString());

                            mrow.rebate_dr = Lib.Conv2Decimal(Dr["rebate_dr"].ToString());

                            mrow.expense = Lib.Conv2Decimal(Dr["buy"].ToString());
                            mrow.income = Lib.Conv2Decimal(Dr["sell"].ToString());
                            mrow.profit = Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.pol = Dr["pol"].ToString();
                            mrow.pod = Dr["pod"].ToString();
                            mrow.pofd = Dr["pofd"].ToString();
                            mrow.agent = Dr["agent"].ToString();
                            mrow.job_type = Dr["job_type"].ToString();
                            mrow.bl_notify_name = Dr["notify"].ToString();
                            mrow.exp_city = Dr["shpr_location"].ToString();
                            mrow.exp_state = Dr["shpr_state"].ToString();
                            mrow.exp_created = Lib.DatetoStringDisplayformat(Dr["shpr_created"]);
                            mrow.job_commodity = Dr["commodity"].ToString();
                            mrow.job_invoice_nos = Dr["job_invoice_nos"].ToString();
                            mrow.pod_country = Dr["pod_country"].ToString();
                            mrow.buyer_name = Dr["buyer_name"].ToString();

                            total += Lib.Conv2Decimal(Dr["profit"].ToString());

                            _buy += Lib.Conv2Decimal(Dr["buy"].ToString());
                            _profit += Lib.Conv2Decimal(Dr["profit"].ToString());
                            mrow.roi = 0;



                            mList.Add(mrow);

                            tot_income += Lib.Conv2Decimal(mrow.income.ToString());
                            tot_expense += Lib.Conv2Decimal(mrow.expense.ToString());
                            tot_profit += Lib.Conv2Decimal(mrow.profit.ToString());



                        }
                        if (mList.Count > 1)
                        {


                            mrow.total = total;
                            tot_total += Lib.Conv2Decimal(mrow.total.ToString());

                            mrow = new Profit();
                            mrow.rowtype = "TOTAL";
                            mrow.rowcolor = "RED";
                            mrow.mbl_bl_no = "TOTAL";
                            mrow.income = Lib.Conv2Decimal(Lib.NumericFormat(tot_income.ToString(), 2));
                            mrow.expense = Lib.Conv2Decimal(Lib.NumericFormat(tot_expense.ToString(), 2));
                            mrow.profit = Lib.Conv2Decimal(Lib.NumericFormat(tot_profit.ToString(), 2));
                            mrow.total = Lib.Conv2Decimal(Lib.NumericFormat(tot_total.ToString(), 2));

                            mList.Add(mrow);


                        }


                        if (type == "EXCEL")
                        {
                            if (mList != null)
                                PrintAirExportClearingReport();
                        }
                        Dt_List.Rows.Clear();
                    }
                }

                if (isnewformat)
                {
                    PrintProfitReport();
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

        private void PrintAirExportForwardingReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";

            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
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

                File_Display_Name = "ProfitReport.xls";
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
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                WS.Columns[37].Width = 256 * 15;
                WS.Columns[38].Width = 256 * 15;
                WS.Columns[39].Width = 256 * 15;
                WS.Columns[40].Width = 256 * 15;
                WS.Columns[41].Width = 256 * 15;
                WS.Columns[42].Width = 256 * 15;
                WS.Columns[43].Width = 256 * 15;
                WS.Columns[44].Width = 256 * 15;
                WS.Columns[45].Width = 256 * 15;
                WS.Columns[46].Width = 256 * 15;
                WS.Columns[47].Width = 256 * 15;
                WS.Columns[48].Width = 256 * 15;
                WS.Columns[49].Width = 256 * 15;
                WS.Columns[50].Width = 256 * 15;
                WS.Columns[51].Width = 256 * 15;
                WS.Columns[52].Width = 256 * 15;
                WS.Columns[53].Width = 256 * 15;
                WS.Columns[54].Width = 256 * 15;
                WS.Columns[55].Width = 256 * 15;


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
                Lib.WriteData(WS, iRow, 1, "PROFIT REPORT : " + branch_name, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBLSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAWB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HAWB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUY-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "FIN-YEAR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE-NOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUYER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOTIFY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
               
                Lib.WriteData(WS, iRow, iCol++, "AIRLINE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
               
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD-COUNTRY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "M-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FRT-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FSC-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "WRS-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MCC-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OTH-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FRT+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FSC+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "WRS+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MCC+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OTH+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MARGIN+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COSTING-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COSTING+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REBATE-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUY", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "PROFIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ROI%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);



                decimal val = 0;
                foreach (Profit Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.buy_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.sell_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exporter, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_city, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_state, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.exp_created, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_year, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.buyer_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.bl_notify_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                       
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_liner, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_commodity, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_country, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.frt_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.fsc_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.wrs_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mcc_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.oth_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.frt_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.fsc_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.wrs_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mcc_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.oth_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.margin_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.frt_ho_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.frt_ho_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rebate_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sell, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.buy, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.roi, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);



                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_date, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
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
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sell, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.buy, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

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

        private void PrintAirImportReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
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

                File_Display_Name = "ProfitReport.xls";
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
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                WS.Columns[37].Width = 256 * 15;
                WS.Columns[38].Width = 256 * 15;
                WS.Columns[39].Width = 256 * 15;
                WS.Columns[40].Width = 256 * 15;
                WS.Columns[41].Width = 256 * 15;
                WS.Columns[42].Width = 256 * 15;
                WS.Columns[43].Width = 256 * 15;
                WS.Columns[44].Width = 256 * 15;
                WS.Columns[45].Width = 256 * 15;
                WS.Columns[46].Width = 256 * 15;
                WS.Columns[47].Width = 256 * 15;
                WS.Columns[48].Width = 256 * 15;
                WS.Columns[49].Width = 256 * 15;
                WS.Columns[50].Width = 256 * 15;
                WS.Columns[51].Width = 256 * 15;



                iRow = 0; iCol = 1;
               
                //WS.Columns[23].Style.NumberFormat = "#0.000";
                //WS.Columns[24].Style.NumberFormat = "#0.000";
                //WS.Columns[25].Style.NumberFormat = "#0.000";
                //WS.Columns[26].Style.NumberFormat = "#0.000";
                //WS.Columns[27].Style.NumberFormat = "#0.00";
                //WS.Columns[28].Style.NumberFormat = "#0.00";
                //WS.Columns[29].Style.NumberFormat = "#0.00";
                //WS.Columns[30].Style.NumberFormat = "#0.00";
                //WS.Columns[31].Style.NumberFormat = "#0.00";
                //WS.Columns[32].Style.NumberFormat = "#0.00";
                //WS.Columns[33].Style.NumberFormat = "#0.00";
                //WS.Columns[34].Style.NumberFormat = "#0.00";
                //WS.Columns[35].Style.NumberFormat = "#0.00";
                //WS.Columns[36].Style.NumberFormat = "#0.00";
                //WS.Columns[37].Style.NumberFormat = "#0.00";
                //WS.Columns[38].Style.NumberFormat = "#0.00";
                //WS.Columns[39].Style.NumberFormat = "#0.00";
                //WS.Columns[40].Style.NumberFormat = "#0.00";
                //WS.Columns[41].Style.NumberFormat = "#0.00";
                //WS.Columns[42].Style.NumberFormat = "#0.00";
                //WS.Columns[43].Style.NumberFormat = "#0.00";
                //WS.Columns[44].Style.NumberFormat = "#0.00";

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
                Lib.WriteData(WS, iRow, 1, "PROFIT REPORT : " + branch_name, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "MAWB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUY-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HAWB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DESCRIPTION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IMPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FIN-YEAR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
               
                Lib.WriteData(WS, iRow, iCol++, "NOTIFY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE-NOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COUNTRY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "LINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "M-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "H-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "IN-1401", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1402", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1403", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1404", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1405", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1401", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1402", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1403", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1404", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1405", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REBATE-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INCOME", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPENSE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "PROFIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ROI%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                decimal val = 0;
                foreach (Profit Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.buy_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_rec_creared_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.sell_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.discription, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exporter, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_city, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_state, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.exp_created, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_year, _Color, false, "", "L", "", _Size, false, 325, "", true);
                       
                        Lib.WriteData(WS, iRow, iCol++, Rec.bl_notify_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_orgin_country, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_jobtype, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.liner, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);




                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1401, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1402, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1403, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1404, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1405, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1401, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1402, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1403, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1404, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1405, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cost_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cost_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rebate_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.roi, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);



                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
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

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);




                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

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

        private void PrintSeaExportForwardingReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
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

                File_Display_Name = "ProfitReport.xls";
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
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                WS.Columns[37].Width = 256 * 15;
                WS.Columns[38].Width = 256 * 15;
                WS.Columns[39].Width = 256 * 15;
                WS.Columns[40].Width = 256 * 15;
                WS.Columns[41].Width = 256 * 15;
                WS.Columns[42].Width = 256 * 15;
                WS.Columns[43].Width = 256 * 15;
                WS.Columns[44].Width = 256 * 15;
                WS.Columns[45].Width = 256 * 15;
                WS.Columns[46].Width = 256 * 15;
                WS.Columns[47].Width = 256 * 15;
                WS.Columns[48].Width = 256 * 15;
                WS.Columns[49].Width = 256 * 15;
                WS.Columns[50].Width = 256 * 15;
                WS.Columns[51].Width = 256 * 15;
                WS.Columns[52].Width = 256 * 15;
                WS.Columns[53].Width = 256 * 15;
                WS.Columns[54].Width = 256 * 15;



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
                Lib.WriteData(WS, iRow, 1, "PROFIT REPORT : " + branch_name, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBLSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUY-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "FIN-YEAR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ETD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CNTR-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TEU", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE-NOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "BUYER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOTIFY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DDP/DDU/EX-WORK", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                

                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD-COUNTRY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "H-CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1105", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1106", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1107", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1105", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1106", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1107", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "REBATE-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "BUY", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PROFIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ROI%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                decimal val = 0;
                foreach (Profit Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.buy_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.sell_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exporter, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_city, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_state, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.exp_created, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_year, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_shipment_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr_teu, _Color, false, "", "L", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.buyer_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.bl_notify_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_liner, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_commodity, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ddp_ddu_exwork, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_country, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1105, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1106, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1107, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1105, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1106, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1107, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cost_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cost_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rebate_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.buy, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sell, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.roi, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_date, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
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

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.buy, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sell, _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, true, "BT", "R", "", _Size, false, 325, "", true);



                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

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

        private void PrintSeaImportReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
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

                File_Display_Name = "ProfitReport.xls";
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
                WS.Columns[8].Width = 256 * 20;
                WS.Columns[9].Width = 256 * 20;
                WS.Columns[10].Width = 256 * 20;
                WS.Columns[11].Width = 256 * 20;
                WS.Columns[12].Width = 256 * 20;
                WS.Columns[13].Width = 256 * 20;
                WS.Columns[14].Width = 256 * 20;
                WS.Columns[15].Width = 256 * 20;
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
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                WS.Columns[37].Width = 256 * 15;
                WS.Columns[38].Width = 256 * 15;
                WS.Columns[39].Width = 256 * 15;
                WS.Columns[40].Width = 256 * 15;
                WS.Columns[41].Width = 256 * 15;
                WS.Columns[42].Width = 256 * 15;
                WS.Columns[43].Width = 256 * 15;
                WS.Columns[44].Width = 256 * 15;
                WS.Columns[45].Width = 256 * 15;
                WS.Columns[46].Width = 256 * 15;
                WS.Columns[47].Width = 256 * 15;
                WS.Columns[48].Width = 256 * 15;
                WS.Columns[49].Width = 256 * 15;
                WS.Columns[50].Width = 256 * 15;
                WS.Columns[51].Width = 256 * 15;
                WS.Columns[52].Width = 256 * 15;
                WS.Columns[53].Width = 256 * 15;
                WS.Columns[54].Width = 256 * 15;
                WS.Columns[55].Width = 256 * 15;
                WS.Columns[56].Width = 256 * 15;
                WS.Columns[57].Width = 256 * 15;
                WS.Columns[58].Width = 256 * 15;
                WS.Columns[59].Width = 256 * 15;
                WS.Columns[60].Width = 256 * 15;
                WS.Columns[61].Width = 256 * 15;
                WS.Columns[62].Width = 256 * 15;
                WS.Columns[63].Width = 256 * 15;
                WS.Columns[64].Width = 256 * 15;



                iRow = 0; iCol = 1;
               
                //WS.Columns[23].Style.NumberFormat = "#0.000";
                //WS.Columns[24].Style.NumberFormat = "#0.000";
                //WS.Columns[25].Style.NumberFormat = "#0.000";
                //WS.Columns[26].Style.NumberFormat = "#0.00";
                //WS.Columns[27].Style.NumberFormat = "#0.00";
                //WS.Columns[28].Style.NumberFormat = "#0.00";
                //WS.Columns[29].Style.NumberFormat = "#0.00";
                //WS.Columns[30].Style.NumberFormat = "#0.00";
                //WS.Columns[31].Style.NumberFormat = "#0.00";
                //WS.Columns[32].Style.NumberFormat = "#0.00";
                //WS.Columns[33].Style.NumberFormat = "#0.00";
                //WS.Columns[34].Style.NumberFormat = "#0.00";
                //WS.Columns[35].Style.NumberFormat = "#0.00";
                //WS.Columns[36].Style.NumberFormat = "#0.00";
                //WS.Columns[37].Style.NumberFormat = "#0.00";
                //WS.Columns[38].Style.NumberFormat = "#0.00";
                //WS.Columns[39].Style.NumberFormat = "#0.00";
                //WS.Columns[40].Style.NumberFormat = "#0.00";
                //WS.Columns[41].Style.NumberFormat = "#0.00";
                //WS.Columns[42].Style.NumberFormat = "#0.00";
                //WS.Columns[43].Style.NumberFormat = "#0.00";
                //WS.Columns[44].Style.NumberFormat = "#0.00";
                //WS.Columns[45].Style.NumberFormat = "#0.00";
                //WS.Columns[46].Style.NumberFormat = "#0.00";
               

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
                Lib.WriteData(WS, iRow, 1, "PROFIT REPORT : " + branch_name, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUY-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DESCRIPTION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IMPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FIN-YEAR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                //Lib.WriteData(WS, iRow, iCol++, "HBL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
               
                Lib.WriteData(WS, iRow, iCol++, "NOTIFY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE-NOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COUNTRY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CNTR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TEU", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NATURE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "LINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "H-CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "EX-1301", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1302", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1303", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1304", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1305", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1306", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1307", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "IN-1301", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1302", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1303", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1304", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1305", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1306", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1307", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "COST-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REBATE-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPENSE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INCOME", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "PROFIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ROI%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                decimal val = 0;
                foreach (Profit Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.buy_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_rec_creared_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.sell_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.discription, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exporter, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_city, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_state, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.exp_created, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_year, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        //Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                       
                        Lib.WriteData(WS, iRow, iCol++, Rec.bl_notify_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_status, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ar_invnos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_orgin_country, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_book_cntr_teu, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nature, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_jobtype, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.liner, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);


                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1301, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1302, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1303, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1304, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1305, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1306, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1307, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1301, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1302, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1303, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1304, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1305, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1306, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1307, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.cost_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cost_cr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rebate_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.roi, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
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
                        //Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
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

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);


                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

                    }

                }


                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void PrintSeaExportClearingReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
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

                File_Display_Name = "ProfitReport.xls";
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
                WS.Columns[5].Width = 256 * 20;
                WS.Columns[6].Width = 256 * 20;
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
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                WS.Columns[37].Width = 256 * 15;
                WS.Columns[38].Width = 256 * 15;
                WS.Columns[39].Width = 256 * 15;
                WS.Columns[40].Width = 256 * 15;
                WS.Columns[41].Width = 256 * 15;
                WS.Columns[42].Width = 256 * 15;
                WS.Columns[43].Width = 256 * 15;
                WS.Columns[44].Width = 256 * 15;
                WS.Columns[45].Width = 256 * 15;
                WS.Columns[46].Width = 256 * 15;
                WS.Columns[47].Width = 256 * 15;
                WS.Columns[48].Width = 256 * 15;
                WS.Columns[49].Width = 256 * 15;
                WS.Columns[50].Width = 256 * 15;

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
                Lib.WriteData(WS, iRow, 1, "PROFIT REPORT : " + branch_name, _Color, true, "", "L", "", 15, false, 325, "", true);
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
                Lib.WriteData(WS, iRow, iCol++, "BUY-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IMPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "BUYER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOTIFY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE-NOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CNTR-TYPES", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JOB-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD-COUNTRY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
               // Lib.WriteData(WS, iRow, iCol++, "CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1101", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1102", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1103", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1101", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1102", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1103", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REBATE-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPENSE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INCOME", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PROFIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ROI%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                //STOP


                decimal val = 0;
                foreach (Profit Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.job_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.buy_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.sell_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exporter, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_city, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_state, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.exp_created, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.buyer_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.bl_notify_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        

                        Lib.WriteData(WS, iRow, iCol++, Rec.job_commodity, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_invoice_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_cntr_type, _Color, false, "", "L", "", _Size, false, 325, "", true);


                        Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_country, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.job_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                      //  Lib.WriteData(WS, iRow, iCol++, Rec.job_chwt, _Color, false, "", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_cbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1101, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1102, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1103, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1101, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1102, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1103, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rebate_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.roi, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);



                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
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
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
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
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);

                    }

                }


                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void PrintAirExportClearingReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
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

                File_Display_Name = "ProfitReport.xls";
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
                WS.Columns[6].Width = 256 * 20;
                WS.Columns[7].Width = 256 * 20;
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
                WS.Columns[31].Width = 256 * 15;
                WS.Columns[32].Width = 256 * 15;
                WS.Columns[33].Width = 256 * 15;
                WS.Columns[34].Width = 256 * 15;
                WS.Columns[35].Width = 256 * 15;
                WS.Columns[36].Width = 256 * 15;
                WS.Columns[37].Width = 256 * 15;
                WS.Columns[38].Width = 256 * 15;
                WS.Columns[39].Width = 256 * 15;
                WS.Columns[40].Width = 256 * 15;
                WS.Columns[41].Width = 256 * 15;
                WS.Columns[42].Width = 256 * 15;
                WS.Columns[43].Width = 256 * 15;
                WS.Columns[44].Width = 256 * 15;
                WS.Columns[45].Width = 256 * 15;
                WS.Columns[46].Width = 256 * 15;
                WS.Columns[47].Width = 256 * 15;
                WS.Columns[48].Width = 256 * 15;
                WS.Columns[49].Width = 256 * 15;

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
                Lib.WriteData(WS, iRow, 1, "PROFIT REPORT : " + branch_name, _Color, true, "", "L", "", 15, false, 325, "", true);
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
                Lib.WriteData(WS, iRow, iCol++, "BUY-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IMPORTER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUYER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOTIFY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                

                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE-NOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JOB-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD-COUNTRY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                //Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1201", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1202", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-1203", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1201", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1202", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-1203", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REBATE-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXPENSE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INCOME", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PROFIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ROI%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);




                decimal val = 0;
                foreach (Profit Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.job_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.buy_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.sell_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exporter, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_city, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.exp_state, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.exp_created, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.agent, _Color, false, "", "L", "", _Size, false, 325, "", false);

                        Lib.WriteData(WS, iRow, iCol++, Rec.job_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.buyer_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.bl_notify_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_commodity, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_invoice_nos, _Color, false, "", "L", "", _Size, false, 325, "", true);


                        Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pod_country, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pofd, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.job_ntwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_chwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.job_grwt, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        //Lib.WriteData(WS, iRow, iCol++, Rec.job_cbm, _Color, false, "", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1201, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1202, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ex_1203, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1201, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1202, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.in_1203, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rebate_dr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.roi, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);





                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if (all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
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
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        //  Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.expense, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.income, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.profit, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        val = Rec.total != null ? Lib.Conv2Decimal(Rec.total.ToString()) : 0;
                        Lib.WriteData(WS, iRow, iCol++, val, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "", true);


                    }

                }


                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }



        private void PrintProfitReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";

            string REPORT_CAPTION = "";
            Color _Color = Color.Black;
            int _Size = 10;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if (!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
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


                sql = " select MBLID, TYPE, REPORTDATE, BRANCH, MBLNO, MBL_DATE, MBLSLNO, MBL_STATUS, FOLDER_NO, JOB_TYPE,SINO, SI_DATE, HBLNO, HBL_DATE, BUY_DATE, SELL_DATE, ";
                sql += " EXPORTER_NAME, CONSIGNEE_NAME, AGENT_NAME, SMAN_NAME, NOMINATION, MBL_FRT_STATUS, HBL_FRT_STATUS, PKG,PCS,  MBL_CHWT, HBL_CHWT, MBL_GRWT, HBL_GRWT, ";
                sql += " HBL_NTWT, HBL_CBM, POL, POD, POFD, ORGIN_COUNTRY, ETD, NATURE, DISCRIPTION, CNTR, TEU, FIN_YEAR, INV_NOS, SHPR_LOCATION, ";
                sql += " SHPR_STATE, SHPR_CREATED, NOTIFY, EX_WORKS, DDP, DDU, LINER, COMMODITY, POD_COUNTRY, BUYER_NAME, ";
                sql += " EX_01, EX_02, EX_03, EX_04, EX_05, EX_06, EX_07, EX_10, EX_17, EX_OT, ";
                sql += " IN_01, IN_02, IN_03, IN_04, IN_05, IN_06, IN_07, IN_10, IN_17, IN_OT, ";
                sql += " COST_DR, COST_CR, REBATE_DR, BUY, SELL, PROFIT, ROI ";
                sql += " from profitreport ";
                sql += " where 1=1 ";

                if (finyear >0)
                    sql += "  and jv_year = " + finyear.ToString();

                if (!all)
                    sql += " and branch = '{BRCODE}' ";
                sql +=" and reportdate between to_date('{FDATE}', 'DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')   ";
                sql += " order by branch,type,reportdate ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();


                File_Display_Name = "ProfitReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                
                for ( int i=0; i<=55; i++)
                    WS.Columns[i].Width = 256 * 15;
                WS.Columns[0].Width = 256 * 2;

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
                Lib.WriteData(WS, iRow, 1, "PROFIT REPORT : " + branch_name, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "MBLSL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HAWB#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUY-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "FIN-YEAR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE-NOS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUYER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOTIFY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMMODITY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD-COUNTRY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "PKG", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PCS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "M-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-CHWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "M-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "H-GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "TEU", _Color, true, "BT", "R", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, "CNTR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JOB-TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);



                Lib.WriteData(WS, iRow, iCol++, "EX01-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX02-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX03-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX04-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX05-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX06-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX07-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX10-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX17-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXOT-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "IN01+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN02+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN03+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN04+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN05+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN06+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN07+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN10+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN17+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INOT+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "COSTING-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COSTING+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REBATE-", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SELL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUY", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "PROFIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ROI%", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                
                foreach (DataRow Dr in Dt_List.Rows )
                {
                    iRow++;
                    iCol = 1;

                    Lib.WriteData(WS, iRow, iCol++, Dr["branch"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["type"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Lib.ExcelCompatibleDate(Dr["reportdate"],Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["mblslno"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["mblno"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["sino"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["hblno"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["buy_date"]), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["sell_date"]), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["exporter_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    if (Dr["type"].ToString() == "AIR-IMPORT" || Dr["type"].ToString() =="SEA-IMPORT")
                        Lib.WriteData(WS, iRow, iCol++, Dr["consignee_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    else 
                        Lib.WriteData(WS, iRow, iCol++, Dr["exporter_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    

                    Lib.WriteData(WS, iRow, iCol++, Dr["shpr_location"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["shpr_state"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["shpr_created"]), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["consignee_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["agent_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["fin_year"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["folder_no"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["inv_nos"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["buyer_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["notify"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["liner"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["commodity"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["sman_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["nomination"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["mbl_frt_status"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["hbl_frt_status"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["pol"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["pod"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);


                    if (Dr["type"].ToString() == "AIR-IMPORT" || Dr["type"].ToString() == "SEA-IMPORT")
                        Lib.WriteData(WS, iRow, iCol++, Dr["orgin_country"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    else 
                        Lib.WriteData(WS, iRow, iCol++, Dr["pod_country"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Dr["pofd"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Integer(Dr["pkg"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["pcs"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["mbl_chwt"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["hbl_chwt"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["mbl_grwt"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["hbl_chwt"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["hbl_cbm"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["teu"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);

                    Lib.WriteData(WS, iRow, iCol++, Dr["cntr"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["job_type"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);



                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_01"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_02"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_03"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_04"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_05"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_06"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_07"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_10"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_17"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ex_ot"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_01"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_02"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_03"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_04"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_05"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_06"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_07"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_10"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_17"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["in_ot"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["cost_dr"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["cost_cr"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["rebate_dr"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["sell"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["buy"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["profit"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["roi"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
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
