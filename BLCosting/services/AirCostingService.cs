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


namespace BLCosting
{
    public class AirCostingService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;

        int iRow = 0;
        int iCol = 0;
        string report_type = "";
        string report_folder = "";
        string report_folderid = "";
        string report_comp_code = "";
        string report_pkid = "";
        string report_branch_code = "";
        string File_Name = "";
        string File_Display_Name = "myreport.xls";
        string Print_FC_Bank = "N";
        private string Bank_Company = "";
        private string Bank_Acno = "";
        private string Bank_Name = "";
        private string Bank_Ifsc_Code = "";
        private string Bank_Add1 = "";
        private string Bank_Add2 = "";
        private string Bank_Add3 = "";

        private DataTable dt_master;
        private DataTable dt_house;
        private DataTable dt_costdet;
        private DataTable dt_bank = new DataTable();
        private DataTable dt_costdet2;

        private DataRow DR_MASTER;

        LovService lov = null;
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Costingm> mList = new List<Costingm>();
            Costingm mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string from_date = "";
            if (SearchData.ContainsKey("from_date"))
                from_date = SearchData["from_date"].ToString();
            string to_date = "";
            if (SearchData.ContainsKey("to_date"))
                to_date = SearchData["to_date"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);

                sWhere = " where a.rec_company_code = '{COMPANY_CODE}' ";
                sWhere += " and a.rec_branch_code = '{BRANCH_CODE}' ";
                sWhere += " and a.cost_year =  {FYEAR} ";
                sWhere += " and a.cost_source = 'AIR EXPORT COSTING' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.cost_folderno) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(agent.cust_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.cost_refno) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }
                if (from_date != "NULL")
                    sWhere += "  and a.cost_date >= to_date('{FDATE}','DD-MON-YYYY') ";
                if (to_date != "NULL")
                    sWhere += "  and a.cost_date <= to_date('{EDATE}','DD-MON-YYYY') ";

                sWhere = sWhere.Replace("{COMPANY_CODE}", company_code);
                sWhere = sWhere.Replace("{BRANCH_CODE}", branch_code);
                sWhere = sWhere.Replace("{FDATE}", from_date);
                sWhere = sWhere.Replace("{EDATE}", to_date);
                sWhere = sWhere.Replace("{FYEAR}", year_code);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM costingm  a ";
                    sql += " left join customerm agent on a.cost_agent_id = agent.cust_pkid ";
                    sql += sWhere;
                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        page_rowcount = Lib.Conv2Integer(Dt_Temp.Rows[0]["total"].ToString());
                        page_count = Lib.Conv2Integer(Dt_Temp.Rows[0]["page_total"].ToString());
                    }
                    page_current = 1;
                }
                else
                {
                    if (type == "FIRST")
                        page_current = 1;
                    if (type == "PREV" && page_current > 1)
                        page_current--;
                    if (type == "NEXT" && page_current < page_count)
                        page_current++;
                    if (type == "LAST")
                        page_current = page_count;
                }

                startrow = (page_current - 1) * page_rows + 1;
                endrow = (startrow + page_rows) - 1;

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select * from ( ";
                sql += " select cost_pkid,cost_cfno,cost_refno,cost_date,cost_folderno,agent.cust_name as agent_name,jvagent.cust_name as jvagent_name, ";
                sql += " cost_jv_ho_vrno,cost_jv_br_vrno,cost_jv_br_invno, cost_jv_posted, ";
                sql += " curr.param_code as cost_currency_code,cost_exrate,cost_profit, cost_our_profit, cost_your_profit,";
                sql += " cost_drcr_amount_inr,cost_drcr_amount,cost_checked_on,cost_sent_on,";
                sql += " row_number() over(order by cost_date,cost_cfno) rn ";
                sql += " from costingm a ";
                sql += " left join customerm agent on a.cost_agent_id = agent.cust_pkid ";
                sql += " left join customerm jvagent on a.cost_jv_agent_id = jvagent.cust_pkid ";
                sql += " left join param curr on a.cost_currency_id = curr.param_pkid";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";

                sql += " order by cost_date,cost_cfno";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Costingm();
                    mRow.cost_pkid = Dr["cost_pkid"].ToString();
                    mRow.cost_refno = Dr["cost_refno"].ToString();
                    mRow.cost_folderno = Dr["cost_folderno"].ToString();
                    mRow.cost_agent_name = Dr["agent_name"].ToString();
                    mRow.cost_jv_agent_name = Dr["jvagent_name"].ToString();

                    mRow.cost_date = Lib.DatetoStringDisplayformat(Dr["cost_date"]);

                    mRow.cost_jv_ho_vrno = Dr["cost_jv_ho_vrno"].ToString();
                    mRow.cost_jv_br_vrno = Dr["cost_jv_br_vrno"].ToString();
                    mRow.cost_jv_br_invno = Dr["cost_jv_br_invno"].ToString();
                    mRow.cost_jv_posted = false;
                    if (Dr["cost_jv_posted"].ToString() == "Y")
                        mRow.cost_jv_posted = true;

                    mRow.cost_currency_code = Dr["cost_currency_code"].ToString();
                    mRow.cost_exrate = Lib.Conv2Decimal(Dr["cost_exrate"].ToString());
                    mRow.cost_profit = Lib.Conv2Decimal(Dr["cost_profit"].ToString());
                    mRow.cost_our_profit = Lib.Conv2Decimal(Dr["cost_our_profit"].ToString());
                    mRow.cost_your_profit = Lib.Conv2Decimal(Dr["cost_your_profit"].ToString());
                    mRow.cost_drcr_amount_inr = Lib.Conv2Decimal(Dr["cost_drcr_amount_inr"].ToString());
                    mRow.cost_drcr_amount = Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString());
                    mRow.cost_checked_on = Lib.DatetoStringDisplayformat(Dr["cost_checked_on"]);
                    mRow.cost_sent_on = Lib.DatetoStringDisplayformat(Dr["cost_sent_on"]);
                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);

            return RetData;
        }

        public Dictionary<string, object> Process(Dictionary<string, object> SearchData)
        {
            string SQL = "";
            decimal exwork = 0, rebate = 0;
            decimal buy_pp = 0, buy_cc = 0, buy_tot = 0;
            decimal sell_pp = 0, sell_cc = 0, sell_tot = 0;

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            List<Costingd> mList = new List<Costingd>();
            Costingd dRow = new Costingd();

            string costingid = SearchData["pkid"].ToString();
            decimal informrate = Lib.Conv2Decimal(SearchData["informrate"].ToString());
            string id = "";

            try
            {
                DataTable Dt_Rec = new DataTable();

                Con_Oracle = new DBConnection();



                SQL = " select cost_mblid from costingm  where cost_pkid = '" + costingid + "'";

                DataTable Dt_costing = new DataTable();
                Dt_costing = Con_Oracle.ExecuteQuery(SQL);
                foreach (DataRow Dr in Dt_costing.Rows)
                {
                    id = Dr["cost_mblid"].ToString();
                    break;
                }

                SQL = "";
                SQL += " select hbl_pkid,";
                SQL += " max(HBL_GRWT) as GRWT,";
                SQL += " max(HBL_CHWT) as CHWT, ";
                SQL += " max(HBL_NO) as HBL_NO,";
                SQL += " max(HBL_BL_NO) as HBL_BL_NO, ";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205001') then (nvl(jv_qty,0) * " + informrate.ToString() + ")  else 0 end) as FRT_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205001') then jv_total else 0 end) as FRT_CC, ";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205003') then jv_total else 0 end) as WRS_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205003') then jv_total else 0 end) as WRS_CC, ";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205002') then jv_total else 0 end) as MYC_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205002') then jv_total else 0 end) as MYC_CC, ";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205017') then jv_total else 0 end) as MCC_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205017') then jv_total else 0 end) as MCC_CC, ";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205010') then jv_total else 0 end) as SRC_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205010') then jv_total else 0 end) as SRC_CC, ";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205022','1205016','1205015') then jv_total else 0 end) as OTH_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205022','1205016','1205015') then jv_total else 0 end) as OTH_CC ";
                SQL += " from (";
                SQL += "   select mbl.hbl_pkid,mbl.hbl_no,mbl.hbl_bl_no, mbl.hbl_grwt,mbl.hbl_chwt, cast('PREPAID' as nvarchar2(10)) as status,jv_qty, jv_total,acc_code  ";
                SQL += "   from ledgerh a";
                SQL += "   inner join ledgert b on a.jvh_pkid =  b.jv_parent_id";
                SQL += "   inner join hblm mbl on a.jvh_cc_id = mbl.hbl_pkid ";
                SQL += "   inner join param curr on jv_curr_id = curr.param_pkid";
                SQL += "   left join acctm ac on b.jv_acc_id = ac.acc_pkid";
                SQL += "   where jvh_cc_id ='" + id + "' and jv_row_type = 'DR-LEDGER' and curr.param_code = 'INR'";
                SQL += "   union all   ";
                SQL += "   select mbl.hbl_pkid,mbl.hbl_no,mbl.hbl_bl_no,mbl.hbl_grwt,mbl.hbl_chwt, cast('COLLECT' as nvarchar2(10)) as status,inv_qty, inv_total,acc_code   ";
                SQL += "   from jobincome a";
                SQL += "   inner join hblm mbl on a.inv_parent_id = mbl.hbl_pkid";
                SQL += "   inner join param curr on inv_curr_id = curr.param_pkid ";
                SQL += "   left join acctm ac on a.inv_acc_id = ac.acc_pkid ";
                SQL += "   where inv_parent_id ='" + id + "' ";//and curr.param_code = 'INR'
                SQL += " ) a group by hbl_pkid";
                DataTable Dt_MblExpense = Con_Oracle.ExecuteQuery(SQL);

                // Hbl Income/Expense
                SQL = "";
                SQL += " select hbl_pkid,hbl_bl_no,";
                SQL += " max(HBL_GRWT) as GRWT,";
                SQL += " max(HBL_CHWT) as CHWT, ";
                SQL += " max(HBL_NO) as HBL_NO,";

                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205001')  then inv_total else 0 end) as FRT_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205001')  then inv_total else 0 end) as FRT_CC,";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205001')  then inv_rate else 0 end) as FRT_RATE_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205001')  then inv_rate else 0 end) as FRT_RATE_CC,";

                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205003')  then inv_total else 0 end) as WRS_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205003')  then inv_total else 0 end) as WRS_CC,";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205003')  then inv_rate else 0 end) as WRS_RATE_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205003')  then inv_rate else 0 end) as WRS_RATE_CC,";

                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205002')  then inv_total else 0 end) as MYC_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205002')  then inv_total else 0 end) as MYC_CC,";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205002')  then inv_rate else 0 end) as MYC_RATE_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205002')  then inv_rate else 0 end) as MYC_RATE_CC,";

                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205017')  then inv_total else 0 end) as MCC_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205017')  then inv_total else 0 end) as MCC_CC,";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205017')  then inv_rate else 0 end) as MCC_RATE_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205017')  then inv_rate else 0 end) as MCC_RATE_CC,";

                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205010')  then inv_total else 0 end) as SRC_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205010')  then inv_total else 0 end) as SRC_CC,";
                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205010')  then inv_rate else 0 end) as SRC_RATE_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205010')  then inv_rate else 0 end) as SRC_RATE_CC,";

                SQL += " sum(case when status = 'PREPAID' and acc_code in ('1205022','1205016','1205015', '1205019')  then inv_total else 0 end) as OTH_PP,";
                SQL += " sum(case when status = 'COLLECT' and acc_code in ('1205022','1205016','1205015', '1205019')  then inv_total else 0 end) as OTH_CC,";

                SQL += " sum(inv_rebate_amt) as rebate ";
                SQL += " from (";
                SQL += " select hbl.hbl_pkid,hbl.hbl_no,hbl.hbl_bl_no,hbl.hbl_grwt,hbl.hbl_chwt, inv_type as status,inv_rate, inv_total, inv_rebate_amt,acc_code ";
                SQL += " from jobincome a";
                SQL += " inner join hblm hbl on a.inv_parent_id = hbl.hbl_pkid";
                SQL += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid ";
                SQL += " inner join param curr on inv_curr_id = curr.param_pkid";
                SQL += " left join acctm ac on a.inv_acc_id = ac.acc_pkid ";
                SQL += " where mbl.hbl_pkid ='" + id + "' and inv_source <> 'EX-WORK'"; //and curr.param_code = 'INR'
                SQL += " ) a  group by hbl_pkid,hbl_bl_no order by hbl_bl_no";

                DataTable Dt_HblIncome = Con_Oracle.ExecuteQuery(SQL);

                // Hbl Expense // EXWORK
                SQL = "";
                SQL += " select sum(round(jv_net_total / jv_exrate,2)) as exwork ";
                SQL += " from jobincome a";
                SQL += " inner join hblm hbl on a.inv_parent_id = hbl.hbl_pkid";
                SQL += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid ";
                SQL += " inner join param curr on inv_curr_id = curr.param_pkid";
                SQL += " left join ledgert l on inv_pkid = jv_pkid ";
                SQL += " where mbl.hbl_pkid ='" + id + "' and curr.param_code = 'INR' and inv_source = 'EX-WORK'";
                SQL += " group by hbl.hbl_pkid, hbl.hbl_bl_no";
                DataTable Dt_ExWork = Con_Oracle.ExecuteQuery(SQL);
                Con_Oracle.CloseConnection();

                mList = new List<Costingd>();
                buy_pp = 0; buy_cc = 0; buy_tot = 0;
                foreach (DataRow Dr in Dt_MblExpense.Rows)
                {
                    dRow = new Costingd();
                    dRow.costd_pkid = Guid.NewGuid().ToString().ToUpper();
                    dRow.costd_parent_id = costingid;
                    dRow.costd_type = "BUY";
                    dRow.costd_acc_id = Dr["hbl_pkid"].ToString();
                    dRow.costd_sino = Dr["hbl_no"].ToString();
                    dRow.costd_blno = Dr["hbl_bl_no"].ToString();
                    dRow.costd_grwt = Lib.Conv2Decimal(Dr["grwt"].ToString());
                    dRow.costd_chwt = Lib.Conv2Decimal(Dr["chwt"].ToString());
                    dRow.costd_frt_pp = Lib.Conv2Decimal(Dr["frt_pp"].ToString());
                    dRow.costd_frt_cc = Lib.Conv2Decimal(Dr["frt_cc"].ToString());
                    dRow.costd_frt_rate_pp = 0;
                    dRow.costd_frt_rate_cc = 0;
                    dRow.costd_wrs_pp = Lib.Conv2Decimal(Dr["wrs_pp"].ToString());
                    dRow.costd_wrs_cc = Lib.Conv2Decimal(Dr["wrs_cc"].ToString());
                    dRow.costd_wrs_rate_pp = 0;
                    dRow.costd_wrs_rate_cc = 0;
                    dRow.costd_myc_pp = Lib.Conv2Decimal(Dr["myc_pp"].ToString());
                    dRow.costd_myc_cc = Lib.Conv2Decimal(Dr["myc_cc"].ToString());
                    dRow.costd_myc_rate_pp = 0;
                    dRow.costd_myc_rate_cc = 0;
                    dRow.costd_mcc_pp = Lib.Conv2Decimal(Dr["mcc_pp"].ToString());
                    dRow.costd_mcc_cc = Lib.Conv2Decimal(Dr["mcc_cc"].ToString());
                    dRow.costd_mcc_rate_pp = 0;
                    dRow.costd_mcc_rate_cc = 0;
                    dRow.costd_src_pp = Lib.Conv2Decimal(Dr["src_pp"].ToString());
                    dRow.costd_src_cc = Lib.Conv2Decimal(Dr["src_cc"].ToString());
                    dRow.costd_src_rate_pp = 0;
                    dRow.costd_src_rate_cc = 0;
                    dRow.costd_oth_pp = Lib.Conv2Decimal(Dr["oth_pp"].ToString());
                    dRow.costd_oth_cc = Lib.Conv2Decimal(Dr["oth_cc"].ToString());
                    dRow.costd_oth_rate_pp = 0;
                    dRow.costd_oth_rate_cc = 0;
                    dRow.costd_ctr = 1;




                    buy_pp = dRow.costd_frt_pp;
                    buy_pp += dRow.costd_wrs_pp;
                    buy_pp += dRow.costd_myc_pp;
                    buy_pp += dRow.costd_mcc_pp;
                    buy_pp += dRow.costd_src_pp;
                    buy_pp += dRow.costd_oth_pp;
                    buy_pp = Lib.Conv2Decimal(Lib.NumericFormat(buy_pp.ToString(), 2));

                    buy_cc = dRow.costd_frt_cc;
                    buy_cc += dRow.costd_wrs_cc;
                    buy_cc += dRow.costd_myc_cc;
                    buy_cc += dRow.costd_mcc_cc;
                    buy_cc += dRow.costd_src_cc;
                    buy_cc += dRow.costd_oth_cc;
                    buy_cc = Lib.Conv2Decimal(Lib.NumericFormat(buy_cc.ToString(), 2));

                    buy_tot = buy_pp + buy_cc;
                    buy_tot = Lib.Conv2Decimal(Lib.NumericFormat(buy_tot.ToString(), 2));

                    dRow.costd_pp = buy_pp;
                    dRow.costd_cc = buy_cc;
                    dRow.costd_tot = buy_tot;
                    mList.Add(dRow);
                    break;
                }
                rebate = 0;
                int iCtr = 0;
                sell_pp = 0; sell_cc = 0; sell_tot = 0;
                foreach (DataRow Dr in Dt_HblIncome.Rows)
                {
                    iCtr++;
                    rebate += Lib.Convert2Decimal(Dr["rebate"].ToString());

                    dRow = new Costingd();
                    dRow.costd_pkid = Guid.NewGuid().ToString().ToUpper();
                    dRow.costd_parent_id = costingid;
                    dRow.costd_type = "SELL";
                    dRow.costd_acc_id = Dr["hbl_pkid"].ToString();
                    dRow.costd_sino = Dr["hbl_no"].ToString();
                    dRow.costd_blno = Dr["hbl_bl_no"].ToString();
                    dRow.costd_grwt = Lib.Conv2Decimal(Dr["grwt"].ToString());
                    dRow.costd_chwt = Lib.Conv2Decimal(Dr["chwt"].ToString());
                    dRow.costd_frt_pp = Lib.Conv2Decimal(Dr["frt_pp"].ToString());
                    dRow.costd_frt_cc = Lib.Conv2Decimal(Dr["frt_cc"].ToString());
                    dRow.costd_frt_rate_pp = Lib.Conv2Decimal(Dr["frt_rate_pp"].ToString());
                    dRow.costd_frt_rate_cc = Lib.Conv2Decimal(Dr["frt_rate_cc"].ToString());
                    dRow.costd_wrs_pp = Lib.Conv2Decimal(Dr["wrs_pp"].ToString());
                    dRow.costd_wrs_cc = Lib.Conv2Decimal(Dr["wrs_cc"].ToString());
                    dRow.costd_wrs_rate_pp = Lib.Conv2Decimal(Dr["wrs_rate_pp"].ToString());
                    dRow.costd_wrs_rate_cc = Lib.Conv2Decimal(Dr["wrs_rate_cc"].ToString());
                    dRow.costd_myc_pp = Lib.Conv2Decimal(Dr["myc_pp"].ToString());
                    dRow.costd_myc_cc = Lib.Conv2Decimal(Dr["myc_cc"].ToString());
                    dRow.costd_myc_rate_pp = Lib.Conv2Decimal(Dr["myc_rate_pp"].ToString());
                    dRow.costd_myc_rate_cc = Lib.Conv2Decimal(Dr["myc_rate_cc"].ToString());
                    dRow.costd_mcc_pp = Lib.Conv2Decimal(Dr["mcc_pp"].ToString());
                    dRow.costd_mcc_cc = Lib.Conv2Decimal(Dr["mcc_cc"].ToString());
                    dRow.costd_mcc_rate_pp = Lib.Conv2Decimal(Dr["mcc_rate_pp"].ToString());
                    dRow.costd_mcc_rate_cc = Lib.Conv2Decimal(Dr["mcc_rate_cc"].ToString());
                    dRow.costd_src_pp = Lib.Conv2Decimal(Dr["src_pp"].ToString());
                    dRow.costd_src_cc = Lib.Conv2Decimal(Dr["src_cc"].ToString());
                    dRow.costd_src_rate_pp = Lib.Conv2Decimal(Dr["src_rate_pp"].ToString());
                    dRow.costd_src_rate_cc = Lib.Conv2Decimal(Dr["src_rate_cc"].ToString());
                    dRow.costd_oth_pp = Lib.Conv2Decimal(Dr["oth_pp"].ToString());
                    dRow.costd_oth_cc = Lib.Conv2Decimal(Dr["oth_cc"].ToString());
                    dRow.costd_oth_rate_pp = 0;
                    dRow.costd_oth_rate_cc = 0;
                    dRow.costd_ctr = iCtr;


                    sell_pp = dRow.costd_frt_pp;
                    sell_pp += dRow.costd_wrs_pp;
                    sell_pp += dRow.costd_myc_pp;
                    sell_pp += dRow.costd_mcc_pp;
                    sell_pp += dRow.costd_src_pp;
                    sell_pp += dRow.costd_oth_pp;
                    sell_pp = Lib.Conv2Decimal(Lib.NumericFormat(sell_pp.ToString(), 2));

                    sell_cc = dRow.costd_frt_cc;
                    sell_cc += dRow.costd_wrs_cc;
                    sell_cc += dRow.costd_myc_cc;
                    sell_cc += dRow.costd_mcc_cc;
                    sell_cc += dRow.costd_src_cc;
                    sell_cc += dRow.costd_oth_cc;
                    sell_cc = Lib.Conv2Decimal(Lib.NumericFormat(sell_cc.ToString(), 2));

                    sell_tot = sell_pp + sell_cc;
                    sell_tot = Lib.Conv2Decimal(Lib.NumericFormat(sell_tot.ToString(), 2));

                    dRow.costd_pp = sell_pp;
                    dRow.costd_cc = sell_cc;
                    dRow.costd_tot = sell_tot;

                    mList.Add(dRow);

                }

                foreach (DataRow Dr in Dt_ExWork.Rows)
                {
                    exwork = Lib.Convert2Decimal(Dr["exwork"].ToString());
                    break;
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("list", mList);
            RetData.Add("rebate", rebate);
            RetData.Add("exwork", exwork);
            return RetData;
        }


        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Costingm mRow = new Costingm();

            string id = SearchData["pkid"].ToString();
            bool bok = false;
            decimal Tot_InvoiceAmt = 0;
            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select cost_pkid, cost_type, cost_source, cost_cfno,cost_refno,cost_folderno,cost_mblid,mbl.hbl_bl_no as cost_mblno";
                sql += " ,mbl.hbl_date as  cost_sob_date, mbl.hbl_folder_sent_date as cost_folder_recdon,  cost_agent_id,agnt.cust_code as cost_agent_code,agnt.cust_name as cost_agent_name,cost_year,cost_date";
                sql += " ,cost_edit_code,cost_exrate,cost_currency_id,c.param_code as cost_currency_code ,cost_rebate";
                sql += " ,cost_ex_works,cost_inform_rate,cost_other_charges,cost_hand_charges ";
                sql += " ,cost_buy_pp,cost_buy_cc,cost_sell_pp,cost_sell_cc,cost_format";
                sql += " ,cost_buy_tot,cost_sell_tot";
                sql += " ,cost_profit,cost_our_profit,cost_your_profit,cost_drcr_amount,cost_income,cost_expense,cost_sell_chwt ";
                sql += " ,cost_jv_agent_id,agnt2.cust_code as cost_jv_agent_code,agnt2.cust_name as cost_jv_agent_name,cost_jv_agent_br_id";
                sql += " ,agntaddr.add_branch_slno as  cost_jv_agent_br_no,agntaddr.add_line1||'\n'||agntaddr.add_line2||'\n'||agntaddr.add_line3 as  cost_jv_agent_br_addr ,cost_jv_br_inv_id ";
                sql += " from costingm a  ";
                sql += " left join hblm mbl on a.cost_mblid = mbl.hbl_pkid ";
                sql += " left join param c on a.cost_currency_id = c.param_pkid ";
                sql += " left join customerm agnt on a.cost_agent_id = agnt.cust_pkid ";
                sql += " left join customerm agnt2 on a.cost_jv_agent_id = agnt2.cust_pkid ";
                sql += " left join addressm agntaddr on a.cost_jv_agent_br_id = agntaddr.add_pkid ";
                sql += " where  a.cost_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    bok = true;
                    mRow = new Costingm();
                    mRow.cost_pkid = Dr["cost_pkid"].ToString();
                    mRow.cost_cfno = Lib.Conv2Integer(Dr["cost_cfno"].ToString());
                    mRow.cost_type = Dr["cost_type"].ToString();
                    mRow.cost_source = Dr["cost_source"].ToString();
                    mRow.cost_refno = Dr["cost_refno"].ToString();
                    mRow.cost_folderno = Dr["cost_folderno"].ToString();
                    mRow.cost_mblid = Dr["cost_mblid"].ToString();
                    mRow.cost_mblno = Dr["cost_mblno"].ToString();
                    mRow.cost_sob_date = Lib.DatetoStringDisplayformat(Dr["cost_sob_date"]);
                    mRow.cost_agent_id = Dr["cost_agent_id"].ToString();
                    mRow.cost_agent_code = Dr["cost_agent_code"].ToString();
                    mRow.cost_agent_name = Dr["cost_agent_name"].ToString();
                    mRow.cost_year = Lib.Conv2Integer(Dr["cost_year"].ToString());
                    mRow.cost_date = Lib.DatetoString(Dr["cost_date"]);
                    mRow.cost_folder_recdon = Lib.DatetoStringDisplayformat(Dr["cost_folder_recdon"]);
                    mRow.cost_exrate = Lib.Conv2Decimal(Dr["cost_exrate"].ToString());
                    mRow.cost_currency_id = Dr["cost_currency_id"].ToString();
                    mRow.cost_currency_code = Dr["cost_currency_code"].ToString();
                    mRow.cost_buy_pp = Lib.Conv2Decimal(Dr["cost_buy_pp"].ToString());
                    mRow.cost_buy_cc = Lib.Conv2Decimal(Dr["cost_buy_cc"].ToString());
                    mRow.cost_sell_pp = Lib.Conv2Decimal(Dr["cost_sell_pp"].ToString());
                    mRow.cost_sell_cc = Lib.Conv2Decimal(Dr["cost_sell_cc"].ToString());
                    mRow.cost_buy_tot = Lib.Conv2Decimal(Dr["cost_buy_tot"].ToString());
                    mRow.cost_sell_tot = Lib.Conv2Decimal(Dr["cost_sell_tot"].ToString());
                    mRow.cost_rebate = Lib.Conv2Decimal(Dr["cost_rebate"].ToString());
                    mRow.cost_ex_works = Lib.Conv2Decimal(Dr["cost_ex_works"].ToString());
                    mRow.cost_hand_charges = Lib.Conv2Decimal(Dr["cost_hand_charges"].ToString());
                    mRow.cost_inform_rate = Lib.Conv2Decimal(Dr["cost_inform_rate"].ToString());
                    mRow.cost_other_charges = Lib.Conv2Decimal(Dr["cost_other_charges"].ToString());
                    mRow.cost_profit = Lib.Conv2Decimal(Dr["cost_profit"].ToString());
                    mRow.cost_our_profit = Lib.Conv2Decimal(Dr["cost_our_profit"].ToString());
                    mRow.cost_your_profit = Lib.Conv2Decimal(Dr["cost_your_profit"].ToString());
                    mRow.cost_drcr_amount = Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString());
                    mRow.cost_income = Lib.Conv2Decimal(Dr["cost_income"].ToString());
                    mRow.cost_expense = Lib.Conv2Decimal(Dr["cost_expense"].ToString());
                    mRow.cost_format = Dr["cost_format"].ToString();
                    mRow.cost_sell_chwt = Lib.Conv2Decimal(Dr["cost_sell_chwt"].ToString());
                    mRow.cost_edit_code = Dr["cost_edit_code"].ToString();
                    mRow.cost_jv_agent_id = Dr["cost_jv_agent_id"].ToString();
                    mRow.cost_jv_agent_code = Dr["cost_jv_agent_code"].ToString();
                    mRow.cost_jv_agent_name = Dr["cost_jv_agent_name"].ToString();
                    mRow.cost_jv_agent_br_id = Dr["cost_jv_agent_br_id"].ToString();
                    mRow.cost_jv_agent_br_no = Dr["cost_jv_agent_br_no"].ToString();
                    mRow.cost_jv_agent_br_addr = Dr["cost_jv_agent_br_addr"].ToString();
                    mRow.cost_jv_br_inv_id = Dr["cost_jv_br_inv_id"].ToString();
                    Tot_InvoiceAmt = Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString());

                    break;
                }

                if (bok)
                {
                    List<Costingd> mList = new List<Costingd>();
                    Costingd bRow;

                    sql = "select costd_pkid,  costd_parent_id,costd_acc_id ";
                    sql += " ,costd_type,costd_grwt,costd_chwt ";
                    sql += " ,costd_frt_pp,costd_frt_cc,costd_frt_rate_pp,costd_frt_rate_cc";
                    sql += " ,costd_wrs_pp,costd_wrs_cc,costd_wrs_rate_pp,costd_wrs_rate_cc";
                    sql += " ,costd_myc_pp,costd_myc_cc,costd_myc_rate_pp,costd_myc_rate_cc";
                    sql += " ,costd_mcc_pp,costd_mcc_cc,costd_mcc_rate_pp,costd_mcc_rate_cc";
                    sql += " ,costd_src_pp,costd_src_cc,costd_src_rate_pp,costd_src_rate_cc";
                    sql += " ,costd_oth_pp,costd_oth_cc,costd_cc,costd_pp,costd_oth_rate_pp,costd_oth_rate_cc,costd_tot,b.hbl_bl_no as bl_no ";
                    sql += " ,costd_category,costd_ctr ";
                    sql += " from costingd a ";
                    sql += " left join hblm b on a.costd_acc_id = b.hbl_pkid";
                    sql += " where costd_parent_id ='{ID}' ";
                    sql += " and nvl(costd_category,'COSTING') = 'COSTING' ";
                    sql += " order by costd_ctr ";
                    sql = sql.Replace("{ID}", id);

                    Dt_Rec = new DataTable(); 
                    Con_Oracle = new DBConnection();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    foreach (DataRow Dr in Dt_Rec.Rows)
                    {
                        bRow = new Costingd();
                        bRow.costd_pkid = Dr["costd_pkid"].ToString();
                        bRow.costd_parent_id = Dr["costd_parent_id"].ToString();
                        bRow.costd_acc_id = Dr["costd_acc_id"].ToString();
                        bRow.costd_type = Dr["costd_type"].ToString();
                        bRow.costd_grwt = Lib.Conv2Decimal(Dr["costd_grwt"].ToString());
                        bRow.costd_chwt = Lib.Conv2Decimal(Dr["costd_chwt"].ToString());
                        bRow.costd_frt_pp = Lib.Conv2Decimal(Dr["costd_frt_pp"].ToString());
                        bRow.costd_frt_cc = Lib.Conv2Decimal(Dr["costd_frt_cc"].ToString());
                        bRow.costd_frt_rate_pp = Lib.Conv2Decimal(Dr["costd_frt_rate_pp"].ToString());
                        bRow.costd_frt_rate_cc = Lib.Conv2Decimal(Dr["costd_frt_rate_cc"].ToString());
                        bRow.costd_wrs_pp = Lib.Conv2Decimal(Dr["costd_wrs_pp"].ToString());
                        bRow.costd_wrs_cc = Lib.Conv2Decimal(Dr["costd_wrs_cc"].ToString());
                        bRow.costd_wrs_rate_pp = Lib.Conv2Decimal(Dr["costd_wrs_rate_pp"].ToString());
                        bRow.costd_wrs_rate_cc = Lib.Conv2Decimal(Dr["costd_wrs_rate_cc"].ToString());
                        bRow.costd_myc_pp = Lib.Conv2Decimal(Dr["costd_myc_pp"].ToString());
                        bRow.costd_myc_cc = Lib.Conv2Decimal(Dr["costd_myc_cc"].ToString());
                        bRow.costd_myc_rate_pp = Lib.Conv2Decimal(Dr["costd_myc_rate_pp"].ToString());
                        bRow.costd_myc_rate_cc = Lib.Conv2Decimal(Dr["costd_myc_rate_cc"].ToString());
                        bRow.costd_mcc_pp = Lib.Conv2Decimal(Dr["costd_mcc_pp"].ToString());
                        bRow.costd_mcc_cc = Lib.Conv2Decimal(Dr["costd_mcc_cc"].ToString());
                        bRow.costd_mcc_rate_pp = Lib.Conv2Decimal(Dr["costd_mcc_rate_pp"].ToString());
                        bRow.costd_mcc_rate_cc = Lib.Conv2Decimal(Dr["costd_mcc_rate_cc"].ToString());
                        bRow.costd_src_pp = Lib.Conv2Decimal(Dr["costd_src_pp"].ToString());
                        bRow.costd_src_cc = Lib.Conv2Decimal(Dr["costd_src_cc"].ToString());
                        bRow.costd_src_rate_pp = Lib.Conv2Decimal(Dr["costd_src_rate_pp"].ToString());
                        bRow.costd_src_rate_cc = Lib.Conv2Decimal(Dr["costd_src_rate_cc"].ToString());
                        bRow.costd_oth_pp = Lib.Conv2Decimal(Dr["costd_oth_pp"].ToString());
                        bRow.costd_oth_cc = Lib.Conv2Decimal(Dr["costd_oth_cc"].ToString());
                        bRow.costd_oth_rate_pp = Lib.Conv2Decimal(Dr["costd_oth_rate_pp"].ToString());
                        bRow.costd_oth_rate_cc = Lib.Conv2Decimal(Dr["costd_oth_rate_cc"].ToString());
    
                        bRow.costd_pp = Lib.Conv2Decimal(Dr["costd_pp"].ToString());
                        bRow.costd_cc = Lib.Conv2Decimal(Dr["costd_cc"].ToString());
                        bRow.costd_tot = Lib.Conv2Decimal(Dr["costd_tot"].ToString());
                        bRow.costd_blno = Dr["bl_no"].ToString();
                        bRow.costd_ctr = Lib.Conv2Integer(Dr["costd_ctr"].ToString());
                        bRow.costd_category = Dr["costd_category"].ToString();
                        mList.Add(bRow);
                    }
                    mRow.DetailList2 = mList;

                    mList = new List<Costingd>();

                    sql = "select costd_pkid,  costd_parent_id,  costd_acc_id,  costd_acc_name , costd_category, ";
                    sql += " costd_blno,costd_acc_qty,costd_acc_rate,";
                    sql += " costd_acc_amt,costd_ctr,costd_remarks ";
                    sql += " from costingd a ";
                    sql += " where costd_parent_id ='{ID}' ";
                    sql += " and nvl(costd_category,'COSTING') = 'INVOICE' ";
                    sql += " order by costd_ctr ";
                    sql = sql.Replace("{ID}", id);

                    Dt_Rec = new DataTable();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_Rec.Rows)
                    {
                        bRow = new Costingd();
                        bRow.costd_pkid = Dr["costd_pkid"].ToString();
                        bRow.costd_parent_id = Dr["costd_parent_id"].ToString();
                        bRow.costd_acc_id = Dr["costd_acc_id"].ToString();
                        bRow.costd_acc_name = Dr["costd_acc_name"].ToString();
                        bRow.costd_blno = Dr["costd_blno"].ToString();
                        bRow.costd_acc_qty = Lib.Conv2Decimal(Dr["costd_acc_qty"].ToString());
                        bRow.costd_acc_rate = Lib.Conv2Decimal(Dr["costd_acc_rate"].ToString());
                        bRow.costd_acc_amt = Lib.Conv2Decimal(Dr["costd_acc_amt"].ToString());
                        bRow.costd_ctr = Lib.Conv2Integer(Dr["costd_ctr"].ToString());
                        bRow.costd_category = Dr["costd_category"].ToString();
                        bRow.costd_remarks = Dr["costd_remarks"].ToString();
                        mList.Add(bRow);
                    }

                    mRow.cost_tot_acc_amt = 0;
                    if (Dt_Rec.Rows.Count > 0)
                        mRow.cost_tot_acc_amt = Tot_InvoiceAmt;
                    mRow.DetailList = mList;
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }

        public string AllValid(Costingm Record)
        {
            string str = "";
            Boolean bError = false;
            try
            {
                if (Record.cost_folderno.Trim().Length <= 0)
                    Lib.AddError(ref str, " | AWB No cannot be blank");


                if (!Lib.IsInFinYear(Record.cost_date, Record._globalvariables.year_start_date, Record._globalvariables.year_end_date, true))
                {
                    bError = true;
                    Lib.AddError(ref str, " | Invalid Date (Future Date or Date not in Financial Year)");
                }

                if (Record.cost_folderno.Trim().Length > 0)
                {
                    sql = "select cost_pkid from (";
                    sql += "select cost_pkid  from costingm a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and a.cost_folderno = '{FOLDERNO}' ";
                    sql += " and a.cost_source ='AIR EXPORT COSTING'";
                    sql += ") a where cost_pkid <> '{PKID}'";

                    sql = sql.Replace("{FOLDERNO}", Record.cost_folderno);
                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                    sql = sql.Replace("{CATEGORY}", Record.rec_category);
                    sql = sql.Replace("{PKID}", Record.cost_pkid);

                    if (Con_Oracle.IsRowExists(sql))
                    {
                        bError = true;
                        Lib.AddError(ref str, " | AWB No already Exists");
                    }
                }
                if (Record.rec_mode == "EDIT")
                {
                    if (Record.cost_jv_agent_id.Trim().Length > 0 || Record.cost_jv_agent_br_id.Trim().Length > 0)
                    {
                        sql = "select add_pkid from addressm where add_pkid = '{ADD_BRID}'";
                        sql += " and  add_parent_id = '{PARENT_ID}'";
                        sql = sql.Replace("{ADD_BRID}", Record.cost_jv_agent_br_id);
                        sql = sql.Replace("{PARENT_ID}", Record.cost_jv_agent_id);
                        if (!Con_Oracle.IsRowExists(sql))
                        {
                            bError = true;
                            Lib.AddError(ref str, " Invalid Agent Address ");
                        }
                    }
                }
                if (!bError)
                {
                    if (Record.rec_mode == "ADD")
                    {
                        sql = "";
                        sql += "select cost_pkid  from costingm a ";
                        sql += " where a.rec_company_code = '{COMPCODE}'";
                        sql += " and a.rec_branch_code = '{BRCODE}'";
                        sql += " and a.rec_category = '{CATEGORY}'";
                        sql += " and to_char(cost_date,'MON') ='{MON}'";
                        sql += " and to_char(cost_date,'yyyy') = '{MONYEAR}'";
                        sql += " and cost_date > '{DATE}'";

                        sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                        sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                        sql = sql.Replace("{CATEGORY}", Record.rec_category);
                        sql = sql.Replace("{DATE}", Lib.StringToDate(Record.cost_date));
                        DateTime dtcostdate = DateTime.Parse(Record.cost_date);
                        sql = sql.Replace("{MON}", dtcostdate.ToString("MMM").ToUpper());
                        sql = sql.Replace("{MONYEAR}", dtcostdate.Year.ToString());

                        if (Con_Oracle.IsRowExists(sql))
                        {
                            bError = true;
                            Lib.AddError(ref str, " | Back Dated Entry Not Possible ");
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(Costingm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            string docrefno = "";
            string doc_prefix = "";


            decimal nDrcrUsd = 0;

            try
            {
                Con_Oracle = new DBConnection();

                if (Record.cost_folderno.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "AWB NO Cannot Be Empty");

                ErrorMessage = AllValid(Record);

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (Record.rec_mode == "ADD")
                {
                    doc_prefix = "";
                    lov = new LovService();
                    DataRow lovRow_Doc_Prefix = lov.getSettings(Record._globalvariables.branch_code, "COST-AIR-PREFIX");
                    if (lovRow_Doc_Prefix != null)
                        doc_prefix = lovRow_Doc_Prefix["name"].ToString();
                    else
                        throw new Exception("Prefix Not Found");

                    if (Record.cost_date.Trim().Length > 0)
                    {
                        DateTime dtbooking = DateTime.Parse(Record.cost_date);
                        string JOB_MON = dtbooking.ToString("MMM").ToUpper();
                        string JOB_MON_YEAR = dtbooking.Year.ToString();
                        string JOB_MON_NO = dtbooking.Month.ToString();

                        // Create REFNO based on company, branch, Fin-Year, month, cost_type 
                        // CPL / MBISF / 2018 / JAN / SEA
                        // CPL / MBISF / 2018 / JAN / AIR

                        sql = "";
                        sql += " select nvl(max(cost_cfno),0) + 1 as monno from costingm a";
                        sql += " where a.rec_company_code = '{COMPCODE}'";
                        sql += " and a.rec_branch_code = '{BRCODE}'";
                        sql += " and a.cost_year =  {FYEAR} ";
                        sql += " and to_char(cost_date,'MON') ='{COSTMON}'";
                        sql += " and to_char(cost_date,'yyyy') = '{COSTMONYEAR}'";
                        sql += " and a.cost_source in ('AIR EXPORT COSTING','DRCR ISSUE')";
                        sql += " and a.cost_type = 'AIR' ";

                        sql = sql.Replace("{COSTMON}", JOB_MON);
                        sql = sql.Replace("{COSTMONYEAR}", JOB_MON_YEAR);
                        sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                        sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                        sql = sql.Replace("{FYEAR}", Record._globalvariables.year_code);
                        //   sql = sql.Replace("{CATEGORY}", Record.rec_category);
                        // sql = sql.Replace("{COSTTYPE}", Record.cost_type);

                        DataTable DT_MON = new DataTable();
                        DT_MON = Con_Oracle.ExecuteQuery(sql);
                        if (DT_MON.Rows.Count > 0)
                        {
                            Record.cost_cfno = Lib.Conv2Integer(DT_MON.Rows[0]["monno"].ToString());
                            docrefno = String.Concat(doc_prefix, "\\", JOB_MON_YEAR, JOB_MON_NO.ToString().PadLeft(2, '0'), Record.cost_cfno.ToString().PadLeft(5, '0'));
                            Record.cost_refno = docrefno;
                        }
                    }
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("costingm", Record.rec_mode, "cost_pkid", Record.cost_pkid);
                Rec.InsertString("cost_mblid", Record.cost_mblid);
                Rec.InsertString("cost_folderno", Record.cost_folderno);
                Rec.InsertString("cost_agent_id", Record.cost_agent_id);
                Rec.InsertDate("cost_date", Record.cost_date);
                Rec.InsertString("cost_currency_id", Record.cost_currency_id);
                Rec.InsertNumeric("cost_exrate", Record.cost_exrate.ToString());
                Rec.InsertNumeric("cost_rebate", Record.cost_rebate.ToString());
                Rec.InsertNumeric("cost_ex_works", Record.cost_ex_works.ToString());
                Rec.InsertNumeric("cost_buy_pp", Record.cost_buy_pp.ToString());
                Rec.InsertNumeric("cost_buy_cc", Record.cost_buy_cc.ToString());
                Rec.InsertNumeric("cost_sell_pp", Record.cost_sell_pp.ToString());
                Rec.InsertNumeric("cost_sell_cc", Record.cost_sell_cc.ToString());
                Rec.InsertString("cost_format", Record.cost_format);
                Rec.InsertNumeric("cost_buy_tot", Record.cost_buy_tot.ToString());
                Rec.InsertNumeric("cost_sell_tot", Record.cost_sell_tot.ToString());
                Rec.InsertNumeric("cost_hand_charges", Record.cost_hand_charges.ToString());
                Rec.InsertNumeric("cost_inform_rate", Record.cost_inform_rate.ToString());
                Rec.InsertNumeric("cost_other_charges", Record.cost_other_charges.ToString());

                Rec.InsertNumeric("cost_profit", Record.cost_profit.ToString());
                Rec.InsertNumeric("cost_our_profit", Record.cost_our_profit.ToString());
                Rec.InsertNumeric("cost_your_profit", Record.cost_your_profit.ToString());

                Rec.InsertNumeric("cost_expense", Record.cost_expense.ToString());
                Rec.InsertNumeric("cost_income", Record.cost_income.ToString());

                if (Lib.Conv2Decimal(Record.cost_exrate.ToString()) > 0)
                {
                    nDrcrUsd = Lib.Conv2Decimal(Record.cost_drcr_amount.ToString()) / Lib.Conv2Decimal(Record.cost_exrate.ToString());
                    nDrcrUsd = Lib.RoundNumber_Latest(nDrcrUsd.ToString(), 2, true);
                }
                if (nDrcrUsd > 0)
                    Rec.InsertString("cost_drcr", "DR");
                else
                    Rec.InsertString("cost_drcr", "CR");
                Rec.InsertNumeric("cost_drcr_amount", nDrcrUsd.ToString());
                Rec.InsertNumeric("cost_drcr_amount_inr", Record.cost_drcr_amount.ToString());
                Rec.InsertNumeric("cost_sell_chwt", Record.cost_sell_chwt.ToString());
                Rec.InsertString("cost_jv_agent_id", Record.cost_jv_agent_id);
                Rec.InsertString("cost_jv_agent_br_id", Record.cost_jv_agent_br_id);
                Rec.InsertString("cost_cntr", Record.cost_folderno);

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("cost_cfno", Record.cost_cfno.ToString());
                    Rec.InsertString("cost_refno", Record.cost_refno);
                    Rec.InsertNumeric("cost_year", Record._globalvariables.year_code);

                    Rec.InsertString("cost_type", Record.cost_type);
                    Rec.InsertString("cost_source", Record.cost_source);
                    Rec.InsertString("cost_prefix", doc_prefix);

                    Rec.InsertString("rec_category", Record.rec_category);
                    Rec.InsertString("rec_deleted", "N");
                    Rec.InsertString("cost_edit_code", "{S}");
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "SYSDATE");
                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "SYSDATE");
                }

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                sql = "Delete from Costingd where costd_parent_id = '" + Record.cost_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                int iCtr = 0;
                foreach (Costingd Row in Record.DetailList2)
                {
                    iCtr++;
                    Rec.CreateRow("Costingd", "ADD", "costd_pkid", Row.costd_pkid);
                    Rec.InsertString("costd_parent_id", Record.cost_pkid);
                    Rec.InsertString("costd_acc_id", Row.costd_acc_id);
                    Rec.InsertString("costd_type", Row.costd_type);
                    Rec.InsertNumeric("costd_grwt", Row.costd_grwt.ToString());
                    Rec.InsertNumeric("costd_chwt", Row.costd_chwt.ToString());
                    Rec.InsertNumeric("costd_frt_pp", Row.costd_frt_pp.ToString());
                    Rec.InsertNumeric("costd_frt_cc", Row.costd_frt_cc.ToString());
                    Rec.InsertNumeric("costd_frt_rate_pp", Row.costd_frt_rate_pp.ToString());
                    Rec.InsertNumeric("costd_frt_rate_cc", Row.costd_frt_rate_cc.ToString());
                    Rec.InsertNumeric("costd_wrs_pp", Row.costd_wrs_pp.ToString());
                    Rec.InsertNumeric("costd_wrs_cc", Row.costd_wrs_cc.ToString());
                    Rec.InsertNumeric("costd_wrs_rate_pp", Row.costd_wrs_rate_pp.ToString());
                    Rec.InsertNumeric("costd_wrs_rate_cc", Row.costd_wrs_rate_cc.ToString());
                    Rec.InsertNumeric("costd_myc_pp", Row.costd_myc_pp.ToString());
                    Rec.InsertNumeric("costd_myc_cc", Row.costd_myc_cc.ToString());
                    Rec.InsertNumeric("costd_myc_rate_pp", Row.costd_myc_rate_pp.ToString());
                    Rec.InsertNumeric("costd_myc_rate_cc", Row.costd_myc_rate_cc.ToString());
                    Rec.InsertNumeric("costd_mcc_pp", Row.costd_mcc_pp.ToString());
                    Rec.InsertNumeric("costd_mcc_cc", Row.costd_mcc_cc.ToString());
                    Rec.InsertNumeric("costd_mcc_rate_pp", Row.costd_mcc_rate_pp.ToString());
                    Rec.InsertNumeric("costd_mcc_rate_cc", Row.costd_mcc_rate_cc.ToString());
                    Rec.InsertNumeric("costd_src_pp", Row.costd_src_pp.ToString());
                    Rec.InsertNumeric("costd_src_cc", Row.costd_src_cc.ToString());
                    Rec.InsertNumeric("costd_src_rate_pp", Row.costd_src_rate_pp.ToString());
                    Rec.InsertNumeric("costd_src_rate_cc", Row.costd_src_rate_cc.ToString());
                    Rec.InsertNumeric("costd_oth_pp", Row.costd_oth_pp.ToString());
                    Rec.InsertNumeric("costd_oth_cc", Row.costd_oth_cc.ToString());
                    Rec.InsertNumeric("costd_oth_rate_pp", Row.costd_oth_rate_pp.ToString());
                    Rec.InsertNumeric("costd_oth_rate_cc", Row.costd_oth_rate_cc.ToString());
                    Rec.InsertNumeric("costd_pp", Row.costd_pp.ToString());
                    Rec.InsertNumeric("costd_cc", Row.costd_cc.ToString());
                    Rec.InsertNumeric("costd_tot", Row.costd_tot.ToString());
                    Rec.InsertNumeric("costd_ctr", iCtr.ToString());
                    Rec.InsertString("costd_category", "COSTING");
                    sql = Rec.UpdateRow();
                    Con_Oracle.ExecuteNonQuery(sql);
                }

                iCtr = 0;
                foreach (Costingd Row in Record.DetailList)
                {
                    iCtr++;
                    if (Row.costd_acc_name != "" || Row.costd_remarks != "" || Lib.Conv2Decimal(Row.costd_acc_amt.ToString()) != 0)
                    {
                        Rec.CreateRow("Costingd", "ADD", "costd_pkid", Row.costd_pkid);
                        Rec.InsertString("costd_parent_id", Record.cost_pkid);
                        Rec.InsertString("costd_acc_id", Row.costd_acc_id);
                        Rec.InsertString("costd_acc_name", Row.costd_acc_name);
                        Rec.InsertString("costd_blno", Row.costd_blno);
                        Rec.InsertNumeric("costd_ctr", iCtr.ToString());
                        Rec.InsertNumeric("costd_acc_qty", Row.costd_acc_qty.ToString());
                        Rec.InsertNumeric("costd_acc_rate", Row.costd_acc_rate.ToString());
                        Rec.InsertNumeric("costd_acc_amt", Row.costd_acc_amt.ToString());
                        Rec.InsertString("costd_category", "INVOICE");
                        Rec.InsertString("costd_remarks", Row.costd_remarks);
                        sql = Rec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                }

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("docno", docrefno);
            return RetData;
        }

        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "IMPORT DATA");
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("dtlist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }

        



        public Dictionary<string, object> PrintNote(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            

            try
            {
                Bank_Company = "";
                Bank_Acno = "";
                Bank_Name = "";
                Bank_Ifsc_Code = "";
                Bank_Add1 = "";
                Bank_Add2 = "";
                Bank_Add3 = "";

                string id = SearchData["pkid"].ToString();

                //printing 
                if (SearchData.ContainsKey("type"))
                    report_type = SearchData["type"].ToString();
                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();
                if (SearchData.ContainsKey("folderid"))
                    report_folderid = SearchData["folderid"].ToString();
                if (SearchData.ContainsKey("comp_code"))
                    report_comp_code = SearchData["comp_code"].ToString();
                if (SearchData.ContainsKey("branch_code"))
                    report_branch_code = SearchData["branch_code"].ToString();
                if (SearchData.ContainsKey("printfcbank"))
                    Print_FC_Bank = SearchData["printfcbank"].ToString();

                report_pkid = SearchData["pkid"].ToString();

                report_folder = System.IO.Path.Combine(report_folder, report_pkid);
                File_Name = System.IO.Path.Combine(report_folder, report_pkid);




                DataTable Dt_Rec = new DataTable();

                Con_Oracle = new DBConnection();

                sql = "";

                sql = " select  cost_refno, cost_folderno,cost_date, b.hbl_bl_no,b.hbl_date, hbl_folder_no, b.hbl_type,b.rec_category, ";
                sql += " agent.cust_name as agent_name,  ";
                sql += " agentadd.add_line1 as agent_line1,";
                sql += " agentadd.add_line2 as agent_line2,";
                sql += " agentadd.add_line3 as agent_line3,";
                sql += " agentadd.add_line4 as agent_line4,";
                sql += " pol.param_name as pol_name,";
                sql += " pod.param_name as pod_name,";
                sql += " curr.param_code as curr_code,";
                sql += " cost_exrate,cost_buy_pp,cost_buy_cc,";
                sql += " cost_sell_pp,cost_sell_cc,cost_rebate,";
                sql += " cost_ex_works,cost_hand_charges,cost_kamai,";
                sql += " cost_other_charges,cost_asper_amount,cost_buy_tot,";
                sql += " cost_sell_tot,cost_profit ,cost_our_profit,cost_your_profit,";
                sql += " cost_drcr_amount,cost_drcr_amount_inr,cost_expense,cost_income,cost_sell_chwt,cost_format,cost_inform_rate ";
                sql += " from costingm a";
                sql += " inner join hblm b on a.cost_mblid = b.hbl_pkid";
                //sql += " left join customerm agent on b.hbl_agent_id = agent.cust_pkid";
                //sql += " left join addressm agentadd on hbl_agent_br_id = agentadd.add_pkid";
                sql += " left join customerm agent on a.cost_jv_agent_id = agent.cust_pkid";
                sql += " left join addressm agentadd on a.cost_jv_agent_br_id = agentadd.add_pkid";
                sql += " left join param pol on hbl_pol_id = pol.param_pkid";
                sql += " left join param pod on hbl_pod_id = pod.param_pkid";
                sql += " left join param curr on cost_currency_id = curr.param_pkid";
                sql += " where cost_pkid ='" + report_pkid + "'";

                dt_master = Con_Oracle.ExecuteQuery(sql);

                if (dt_master.Rows.Count > 0)
                {
                    DR_MASTER = dt_master.Rows[0];
                }

                sql = "";
                sql += " select  h.hbl_bl_no, cons.cust_name as consignee_name, shpr.cust_name as shipper_name ";
                sql += " from costingm a";
                sql += " inner join hblm m on a.cost_mblid = m.hbl_pkid";
                sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
                sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
                sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";
                sql += " where cost_pkid ='" + report_pkid + "'";

                dt_house = Con_Oracle.ExecuteQuery(sql);

                sql = "select costd_type,costd_grwt,costd_chwt ";
                sql += " ,costd_frt_pp,costd_frt_cc,costd_frt_rate_pp,costd_frt_rate_cc";
                sql += " ,costd_wrs_pp,costd_wrs_cc,costd_wrs_rate_pp,costd_wrs_rate_cc";
                sql += " ,costd_myc_pp,costd_myc_cc,costd_myc_rate_pp,costd_myc_rate_cc";
                sql += " ,costd_mcc_pp,costd_mcc_cc,costd_mcc_rate_pp,costd_mcc_rate_cc";
                sql += " ,costd_src_pp,costd_src_cc,costd_src_rate_pp,costd_src_rate_cc";
                sql += " ,costd_oth_pp,costd_oth_cc,costd_oth_rate_pp,costd_oth_rate_cc,costd_cc,costd_pp,costd_tot  ";
                sql += " ,costd_ctr ";
                sql += " ,b.hbl_bl_no ,b.hbl_terms ";
                sql += " from costingd a ";
                sql += " left join hblm b on a.costd_acc_id = b.hbl_pkid";
                sql += " where costd_parent_id ='" + report_pkid + "'";
                sql += " and nvl(costd_category,'COSTING') = 'COSTING' ";
                sql += " order by costd_ctr ";
                sql = sql.Replace("{ID}", id);

                dt_costdet = Con_Oracle.ExecuteQuery(sql);

                sql = "";
                sql += " select  costd_acc_name ,costd_acc_amt,costd_remarks ";
                sql += " from costingd ";
                sql += " where costd_parent_id ='" + report_pkid + "'";
                sql += " and nvl(costd_category,'COSTING') = 'INVOICE' ";
                sql += " order by costd_ctr";
                dt_costdet2 = Con_Oracle.ExecuteQuery(sql);

                if (Print_FC_Bank == "Y")
                {
                    dt_bank = new DataTable();
                    sql = "select caption, name from settings where parentid ='" + report_branch_code + "' and tabletype ='FC'";
                    dt_bank = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow dr in dt_bank.Rows)
                    {
                        if (dr["caption"].ToString() == "BANK_COMPANY")
                            Bank_Company = dr["name"].ToString();
                        else if (dr["caption"].ToString() == "BANK_ACNO")
                            Bank_Acno = dr["name"].ToString();
                        else if (dr["caption"].ToString() == "BANK_NAME")
                            Bank_Name = dr["name"].ToString();
                        else if (dr["caption"].ToString() == "BANK_IFSC_CODE")
                            Bank_Ifsc_Code = dr["name"].ToString();
                        else if (dr["caption"].ToString() == "BANK_ADD1")
                            Bank_Add1 = dr["name"].ToString();
                        else if (dr["caption"].ToString() == "BANK_ADD2")
                            Bank_Add2 = dr["name"].ToString();
                        else if (dr["caption"].ToString() == "BANK_ADD3")
                            Bank_Add3 = dr["name"].ToString();
                    }
                }

                Con_Oracle.CloseConnection();

                if (report_type == "EXCEL")
                {
                    if (Lib.CreateFolder(report_folder))
                        ProcessExcelFile();
                }
                if (report_type == "EXCEL2")
                {
                    if (Lib.CreateFolder(report_folder))
                        ProcessDetailFile();
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("filename", File_Name + ".xls");
            RetData.Add("filetype", report_type);
            RetData.Add("filedisplayname", File_Display_Name.Replace("\\", ""));
            return RetData;

        }


        private void ProcessExcelFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 0;

            decimal nDrCRAmt = 0;


            decimal buy_pp = 0;
            decimal buy_cc = 0;
            decimal buy_tot = 0;

            decimal sell_pp = 0;
            decimal sell_cc = 0;
            decimal sell_tot = 0;


            decimal kamai = 0;

            decimal rebate = 0;
            decimal exwork = 0;
            decimal other = 0;

            decimal income = 0;
            decimal expense = 0;

            decimal profit = 0;
            decimal our_profit = 0;
            decimal your_profit = 0;


            string sTitle = "";

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];
            WS.PrintOptions.Portrait = true;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;

            WS.Columns[0].Width = 256;
            WS.Columns[1].Width = 256 * 16;
            WS.Columns[2].Width = 256 * 10;
            WS.Columns[3].Width = 256 * 10;
            WS.Columns[4].Width = 256 * 10;
            WS.Columns[5].Width = 256 * 10;
            WS.Columns[6].Width = 256 * 12;
            WS.Columns[7].Width = 256 * 10;
            WS.Columns[8].Width = 256 * 10;
            WS.Columns[9].Width = 256 * 10;
            WS.Columns[10].Width = 256 * 12;
            WS.Columns[11].Width = 256 * 12;
            WS.Columns[12].Width = 256 * 12;

            iRow = 1; iCol = 1;

            //iRow = Lib.WriteHoAddress(WS, report_comp_code, iRow, iCol,7,1,true);


            string comp_name = "";
            string comp_add1 = "";
            string comp_add2 = "";
            string comp_add3 = "";
            string comp_add4 = "";


            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "COMP_ADDRESS");
            mSearchData.Add("comp_code", report_comp_code);
            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    comp_name = Dr["COMP_NAME"].ToString();
                    comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    // comp_add4 = "Email : " + Dr["COMP_email"].ToString() + " Web : " + Dr["COMP_WEB"].ToString();
                    comp_add4 = "Email : hodoc@cargomar.in Web : " + Dr["COMP_WEB"].ToString();
                    break;
                }
            }

            iRow = 1; iCol = 1;
            _Color = Color.Black;
            _Size = 16;
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_name, "Calibri", 14, true, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add1, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add2, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add3, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add4, "Calibri", 12, false, Color.Black, "C", "C", "", "");

            DateTime Dt;

            string sDate = ((DateTime)DR_MASTER["cost_date"]).ToString("dd/MM/yyyy");
            string sCntr = "";
            string Str = "";

            iRow++;

            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "DEBIT NOTE";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "CREDIT NOTE";

            Lib.WriteMergeCell(WS, iRow++, 1, 12, 2, sTitle, "Calibri", 18, true, Color.Black, "C", "C", "TB", "THIN");

            iRow += 2;

            _Size = 13;

            Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 9 : 5), "NUMBER", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 10 : 6), DR_MASTER["COST_REFNO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            File_Display_Name = DR_MASTER["COST_REFNO"].ToString() + "-" + DR_MASTER["COST_FOLDERNO"].ToString() + ".xls";

            iRow++;

            Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_LINE1"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 9 : 5), "DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 10 : 6), sDate, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE2"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE3"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE4"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "WE DEBIT YOUR ACCOUNT FOR THE FOLLOWING";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "WE CREDIT YOUR ACCOUNT FOR THE FOLLOWING";

            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, sTitle, "Calibri", 12, true, Color.Black, "C", "C", "TB", "THIN");

            Lib.WriteData(WS, iRow, 1, "MAWB NO", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["HBL_BL_NO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            iRow++;
            Lib.WriteData(WS, iRow, 1, "HAWB NO", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

            Str = "";
            foreach (DataRow Dr in dt_house.Rows)
            {
                if (Str != "")
                    Str += ",";
                Str += Dr["hbl_bl_no"].ToString();
            }

            Lib.WriteMergeCell(WS, iRow, 3, 5, 1, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
            iRow++;

            Lib.WriteData(WS, iRow, 1, "CONSIGNEE", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

            int iCount = 0;
            Str = "";
            String PreStr = "1";
            foreach (DataRow Dr in dt_house.Select("1=1", "consignee_name"))
            {
                if (Str != "")
                    Str += "\n";

                if (PreStr != Dr["consignee_name"].ToString())
                {
                    PreStr = Dr["consignee_name"].ToString();
                    Str += Dr["consignee_name"].ToString();
                    iCount++;
                }
            }
            if (iCount == 0)
                iCount = 1;

            Lib.WriteMergeCell(WS, iRow, 3, 5, iCount, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
            iRow += iCount;

            Lib.WriteData(WS, iRow, 1, "MAWB DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, Lib.DatetoStringDisplayformat(DR_MASTER["HBL_DATE"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow, 1, "AIRPORT OF DEPARTURE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["POL_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;

            Lib.WriteData(WS, iRow, 1, "AIRPORT OF DESTINATION", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["POD_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;

            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, "", "Calibri", 11, true, Color.Black, "C", "C", "T", "THIN");

            iCol = 1;
            _Color = Color.Black;
            _Border = "";
            _Size = 12;
            if (DR_MASTER["cost_format"].ToString() == "HANDLING")
            {
                decimal nAmt = 0;
                iRow += 4;
                Lib.WriteData(WS, iRow, 2, "PARTICULARS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 7, "AMOUNT(" + DR_MASTER["curr_code"].ToString() + ")", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);

                iRow++;

                if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                    Lib.WriteData(WS, iRow, 2, "OUR HANDLING CHARGES", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                else
                    Lib.WriteData(WS, iRow, 2, "YOUR HANDLING CHARGES", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                nAmt = GetConvertAmt(DR_MASTER["cost_hand_charges"].ToString(), DR_MASTER["cost_exrate"].ToString());
                Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow++;

                if (Lib.Conv2Decimal(DR_MASTER["cost_buy_pp"].ToString()) != 0)
                {
                    Lib.WriteData(WS, iRow, 2, "BUY PP", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    nAmt = GetConvertAmt(DR_MASTER["cost_buy_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    iRow++;
                }

                if (Lib.Conv2Decimal(DR_MASTER["cost_ex_works"].ToString()) != 0)
                {
                    Lib.WriteData(WS, iRow, 2, "EX.Work", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    nAmt = GetConvertAmt(DR_MASTER["cost_ex_works"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    iRow++;
                }
                if (Lib.Conv2Decimal(DR_MASTER["cost_other_charges"].ToString()) != 0)
                {
                    Lib.WriteData(WS, iRow, 2, "Other Charges", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    nAmt = GetConvertAmt(DR_MASTER["cost_other_charges"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    iRow++;
                }

                iRow++;

                Lib.WriteData(WS, iRow, 2, "TOTAL", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 7, DR_MASTER["cost_drcr_amount"], _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow += 6;

                nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());

            }
            else if (DR_MASTER["cost_format"].ToString() == "PP")
            {
                _Border = "LTBR";
                iRow += 6;
                decimal sellrate = 0;
                decimal informrate = 0;
                decimal irate = 0;
                decimal nAmt = 0;

                foreach (DataRow dr in dt_costdet.Select("costd_type = 'SELL'", "costd_ctr"))
                {
                    sellrate += Lib.Conv2Decimal(dr["costd_frt_rate_pp"].ToString());
                    sellrate += Lib.Conv2Decimal(dr["costd_frt_rate_cc"].ToString());
                    break;
                }

                Lib.WriteMergeCell(WS, iRow, 1, 3, 2, "YOUR HANDLING CHARGES", "Calibri", 11, true, Color.Black, "C", "C", "TBLR", "THIN");
                Lib.WriteData(WS, iRow, 4, "INR/KG", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 5, "KGS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 6, "TOTAL INR", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 7, "TOTAL", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                iRow++;

                informrate = Lib.Conv2Decimal(DR_MASTER["cost_inform_rate"].ToString());
                //irate = (sellrate - informrate) / 2;
                //irate = Lib.Conv2Decimal(Lib.NumericFormat(irate.ToString(), 2));
                if (Lib.Conv2Decimal(DR_MASTER["cost_sell_chwt"].ToString()) != 0)//Changed on as per req by ajith doc on 07/01/2019
                {
                    irate = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount_inr"].ToString()) / Lib.Conv2Decimal(DR_MASTER["cost_sell_chwt"].ToString());
                    irate = Lib.Conv2Decimal(Lib.NumericFormat(irate.ToString(), 2));
                    irate = Math.Abs(irate);
                }

                //nAmt = GetConvertAmt(irate.ToString(), DR_MASTER["cost_exrate"].ToString());
                //Lib.WriteData(WS, iRow, 4, nAmt, _Color, false, _Border, "L", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, 4, irate, _Color, false, _Border, "L", "", _Size, false, 325, "#,0.00", true);

                Lib.WriteData(WS, iRow, 5, DR_MASTER["cost_sell_chwt"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                decimal nDrCRAmt_INR = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount_inr"].ToString());
                if (nDrCRAmt_INR < 0)
                    nDrCRAmt_INR = Math.Abs(nDrCRAmt_INR);
                Lib.WriteData(WS, iRow, 6, nDrCRAmt_INR, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());
                if (nDrCRAmt < 0)
                    nDrCRAmt = Math.Abs(nDrCRAmt);
                Lib.WriteData(WS, iRow, 7, nDrCRAmt, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                iRow += 6;
            }
            else
            {

                buy_pp = Lib.Conv2Decimal(DR_MASTER["COST_BUY_PP"].ToString());
                buy_cc = Lib.Conv2Decimal(DR_MASTER["COST_BUY_CC"].ToString());
                buy_tot = buy_pp + buy_cc;

                sell_pp = Lib.Conv2Decimal(DR_MASTER["COST_SELL_PP"].ToString());
                sell_cc = Lib.Conv2Decimal(DR_MASTER["COST_SELL_CC"].ToString());
                sell_tot = sell_pp + sell_cc;


                rebate = Lib.Conv2Decimal(DR_MASTER["COST_REBATE"].ToString());
                exwork = Lib.Conv2Decimal(DR_MASTER["COST_EX_WORKS"].ToString());
                other = Lib.Conv2Decimal(DR_MASTER["COST_OTHER_CHARGES"].ToString());


                decimal income_pp = 0;
                decimal income_cc = 0;
                decimal expense_pp = 0;
                decimal expense_cc = 0;

                income = 0;
                expense = 0;

                profit = Lib.Conv2Decimal(DR_MASTER["COST_PROFIT"].ToString());
                profit = GetConvertAmt(profit.ToString(), DR_MASTER["cost_exrate"].ToString());
                our_profit = Lib.Conv2Decimal(DR_MASTER["COST_OUR_PROFIT"].ToString());
                our_profit = GetConvertAmt(our_profit.ToString(), DR_MASTER["cost_exrate"].ToString());
                your_profit = Lib.Conv2Decimal(DR_MASTER["COST_YOUR_PROFIT"].ToString());
                your_profit = GetConvertAmt(your_profit.ToString(), DR_MASTER["cost_exrate"].ToString());
               
                iRow += 2;
                Str = "A.INCOME FREIGHT COLLECTED BY " + DR_MASTER["agent_name"].ToString();
                Lib.WriteData(WS, iRow, 1, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                //Lib.WriteData(WS, iRow, 11, sell_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                //Lib.WriteData(WS, iRow, 12, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow++;
                _Border = "LTRB";
                Str = DR_MASTER["curr_code"].ToString();
                Lib.WriteMergeCell(WS, iRow, 1, 1, 3, "HAWB", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 2, 1, 3, "FRT STATUS", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 3, 1, 3, "WEIGHT", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 4, 1, 3, "RATE/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 5, 1, 3, "FSC/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 6, 1, 3, "SRC/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 7, 1, 3, "WRS/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 8, 1, 3, "MCC/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 9, 1, 3, "OTHERS (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 10, 1, 3, "TOTAL PP (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 11, 1, 3, "TOTAL CC (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 12, 1, 3, "TOTAL (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
              //  Lib.WriteMergeCell(WS, iRow, 13, 1, 3, "TOTAL (INR)", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                iRow++; iRow++;
                bool bok = false;
                decimal nAmt = 0;  
                foreach (DataRow dr in dt_costdet.Select("costd_type = 'SELL'", "costd_ctr"))
                {
                    bok = true;
                    iRow++;

                    Lib.WriteData(WS, iRow, 1, dr["hbl_bl_no"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Str = "";
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        Str = "PPD";
                    else if (dr["hbl_terms"].ToString() == "FREIGHT COLLECT")
                        Str = "FOB";
                    Lib.WriteData(WS, iRow, 2, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 3, dr["costd_chwt"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_frt_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_frt_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 4, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_myc_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_myc_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 5, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_src_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_src_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 6, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_wrs_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_wrs_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 7, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_mcc_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_mcc_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 8, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);

                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_oth_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_oth_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 9, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true); 
                    nAmt = GetConvertAmt(dr["costd_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 10, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    nAmt = GetConvertAmt(dr["costd_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 11, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    nAmt = GetConvertAmt(dr["costd_tot"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 12, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
               //     Lib.WriteData(WS, iRow, 13, dr["costd_tot"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                }
                if (bok)
                {
                    iRow++;
                    Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Str = "";
                    Lib.WriteData(WS, iRow, 2, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 3, DR_MASTER["cost_sell_chwt"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 4, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 5, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 6, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 7, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 8, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 9, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    sell_pp = GetConvertAmt(sell_pp.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 10, sell_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    sell_cc = GetConvertAmt(sell_cc.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 11, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                  //  nAmt = sell_tot;
                    sell_tot = GetConvertAmt(sell_tot.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 12, sell_tot, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                  //  Lib.WriteData(WS, iRow, 13, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                }

                iRow++;
                iRow++;
                _Border = "";

                income_pp += sell_pp;
                income_cc += sell_cc;
                income += sell_tot;

                //Str = "A.INCOME FREIGHT COLLECTED BY " + DR_MASTER["agent_name"].ToString();
                //Lib.WriteData(WS, iRow, 1, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                //Lib.WriteData(WS, iRow, 11, sell_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                //Lib.WriteData(WS, iRow, 12, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow++;
                iRow++;
                iRow++;
                 
                Lib.WriteData(WS, iRow++, 1, "B.EXPENSE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                _Border = "TBLR";
                Str = DR_MASTER["curr_code"].ToString();
                Lib.WriteMergeCell(WS, iRow, 1,2, 3, "WEIGHT", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                //Lib.WriteMergeCell(WS, iRow, 2, 1, 3, "", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 3, 1, 3, "FRT/KG", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 4, 1, 3, "FRT", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 5, 1, 3, "FSC", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 6, 1, 3, "SRC", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 7, 1, 3, "WRS", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 8, 1, 3, "MCC", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 9, 1, 3, "OTHERS", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 10, 1, 3, "TOTAL PP (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 11, 1, 3, "TOTAL CC (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                Lib.WriteMergeCell(WS, iRow, 12, 1, 3, "TOTAL (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
              //  Lib.WriteMergeCell(WS, iRow, 13, 1, 3, "TOTAL (INR)", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
                iRow++; iRow++;
                bok = false;  
                foreach (DataRow dr in dt_costdet.Select("costd_type = 'BUY'", "costd_ctr"))
                {
                    bok = true;
                    iRow++;
                    Lib.WriteMergeCell(WS, iRow, 1, 2, 1, dr["costd_chwt"].ToString(), "Calibri", _Size, false, _Color, "L", "T", _Border, "THIN", true, false);
                    nAmt = GetConvertAmt(DR_MASTER["COST_INFORM_RATE"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 3, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_frt_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_frt_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 4, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_myc_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_myc_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 5, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_src_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_src_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 6, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_wrs_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_wrs_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 7, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_mcc_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_mcc_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 8, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);

                    if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
                        nAmt = GetConvertAmt(dr["costd_oth_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    else
                        nAmt = GetConvertAmt(dr["costd_oth_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 9, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    nAmt = GetConvertAmt(dr["costd_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 10, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    nAmt = GetConvertAmt(dr["costd_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 11, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    nAmt = GetConvertAmt(dr["costd_tot"].ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 12, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                  //  Lib.WriteData(WS, iRow, 13, dr["costd_tot"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                }
                if (bok)
                {
                    iRow++;
                    Lib.WriteMergeCell(WS, iRow, 1, 2, 1, "TOTAL", "Calibri", _Size, true, _Color, "L", "T", _Border, "THIN", true, false);
                    //Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Str = "";
                   // Lib.WriteData(WS, iRow, 2, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 3, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 4, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 5, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 6, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 7, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 8, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 9, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    buy_pp = GetConvertAmt(buy_pp.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 10, buy_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    buy_cc = GetConvertAmt(buy_cc.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 11, buy_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                   // nAmt = buy_tot;
                    buy_tot = GetConvertAmt(buy_tot.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 12, buy_tot, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                   // Lib.WriteData(WS, iRow, 13, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                }

                iRow++;

                expense_pp += buy_pp;
                expense_cc += buy_cc;
                expense += buy_tot;
                _Border = "";  
                if (rebate > 0)
                {
                    Lib.WriteData(WS, iRow, 1, "REBATE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    rebate = GetConvertAmt(rebate.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 10, rebate, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                   
                    iRow++;
                    expense_pp += rebate;
                    expense += rebate;
                }

                if (exwork > 0)
                {
                    
                    Lib.WriteData(WS, iRow, 1, "Ex-WORK", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    exwork = GetConvertAmt(exwork.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 10, exwork, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    iRow++;
                }

                if (other > 0)
                {
                     
                    Lib.WriteData(WS, iRow, 1, "OTHER CHARGES", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    other = GetConvertAmt(other.ToString(), DR_MASTER["cost_exrate"].ToString());
                    Lib.WriteData(WS, iRow, 10, other, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    iRow++;
                    expense_pp += other;
                    expense += other;
                }
                 
                iRow++;
                iRow++;

                Lib.WriteData(WS, iRow, 1, "C.NET PROFIT/ LOSS(+ / -) A - B", _Color, false, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, 12, profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow++;
                iRow++;
   
                Lib.WriteData(WS, iRow, 1, "PROFIT / LOSS(+ / -) SHARE", _Color,false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 10, our_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, 11, your_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 10, buy_pp + our_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, 11, your_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow++;
                iRow++;
                Str = "FREIGHT AND OTHER CHARGES COLLECTED BY " + DR_MASTER["agent_name"].ToString();
                Lib.WriteData(WS, iRow, 1, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 11, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                iRow++;

                nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());

                _Size++;

                iRow++;
                iRow++;

                if (nDrCRAmt > 0)
                {
                    Lib.WriteData(WS, iRow, 1, "NET DUE FROM " + DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 10, Math.Abs(nDrCRAmt), _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                }
                else
                {
                    Lib.WriteData(WS, iRow, 1, "NET DUE TO " + DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 11, Math.Abs(nDrCRAmt), _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                }
                iRow += 6;
            }


            if (nDrCRAmt < 0)
                nDrCRAmt = Math.Abs(nDrCRAmt);

            string sAmt = Lib.NumericFormat(nDrCRAmt.ToString(), 2);

            string sWords = "";
            if (DR_MASTER["curr_code"].ToString() != "INR")
                sWords = Number2Word_USD.Convert(sAmt, DR_MASTER["CURR_CODE"].ToString(), "CENTS");
            if (DR_MASTER["curr_code"].ToString() == "INR")
                sWords = Number2Word_RS.Convert(sAmt, "INR", "PAISE");


            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, sWords, "Calibri", 14, true, Color.Black, "L", "C", "TB", "THIN");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, "E.&.O.E", "Calibri", 12, true, Color.Black, "L", "C", "TB", "THIN");

            if (Print_FC_Bank == "Y")
            {
                iRow++;
                iRow++;
                Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, "BANK DETAILS", "Calibri", 12, true, Color.Black, "L", "C", "B", "THIN");
                Lib.WriteData(WS, iRow, 1, "BENEFICIARY NAME", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Company, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "USD ACCOUNT NO", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Acno, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "BANK NAME", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Name, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "BANK ADDRESS", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Add1, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Add2, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "SWIFT CODE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Ifsc_Code, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "CORRESPONDENT BANK", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, Bank_Add3, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                WS.Cells.GetSubrangeRelative(iRow++, 1, 7, 1).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
            }
            if (DR_MASTER["cost_format"].ToString() == "HANDLING" || DR_MASTER["cost_format"].ToString() == "PP") 
            {
                WS.Columns[12].Delete();
                WS.Columns[11].Delete();
                WS.Columns[10].Delete();
            }
            
            WB.SaveXls(File_Name + ".xls");
        }

        private void ProcessDetailFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 0;
            bool IsRemarksExist = false;

            decimal nDrCRAmt = 0;


            decimal buy_pp = 0;
            decimal buy_cc = 0;
            decimal buy_tot = 0;

            decimal sell_pp = 0;
            decimal sell_cc = 0;
            decimal sell_tot = 0;


            decimal kamai = 0;

            decimal rebate = 0;
            decimal exwork = 0;
            decimal other = 0;

            decimal income = 0;
            decimal expense = 0;

            decimal profit = 0;
            decimal our_profit = 0;
            decimal your_profit = 0;


            string sTitle = "";

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];
            WS.PrintOptions.Portrait = true;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;

            WS.Columns[0].Width = 256;
            WS.Columns[1].Width = 256 * 16;
            WS.Columns[2].Width = 256 * 10;
            WS.Columns[3].Width = 256 * 10;
            WS.Columns[4].Width = 256 * 10;
            WS.Columns[5].Width = 256 * 10;
            WS.Columns[6].Width = 256 * 12;
            WS.Columns[7].Width = 256 * 10;
            WS.Columns[8].Width = 256 * 10;
            WS.Columns[9].Width = 256 * 10;
            WS.Columns[10].Width = 256 * 12;
            WS.Columns[11].Width = 256 * 12;
            WS.Columns[12].Width = 256 * 12;

            iRow = 1; iCol = 1;

            //iRow = Lib.WriteHoAddress(WS, report_comp_code, iRow, iCol,7,1,true);


            string comp_name = "";
            string comp_add1 = "";
            string comp_add2 = "";
            string comp_add3 = "";
            string comp_add4 = "";


            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "COMP_ADDRESS");
            mSearchData.Add("comp_code", report_comp_code);
            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    comp_name = Dr["COMP_NAME"].ToString();
                    comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    // comp_add4 = "Email : " + Dr["COMP_email"].ToString() + " Web : " + Dr["COMP_WEB"].ToString();
                    comp_add4 = "Email : hodoc@cargomar.in Web : " + Dr["COMP_WEB"].ToString();
                    break;
                }
            }

            iRow = 1; iCol = 1;
            _Color = Color.Black;
            _Size = 16;
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_name, "Calibri", 14, true, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add1, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add2, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add3, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add4, "Calibri", 12, false, Color.Black, "C", "C", "", "");

            DateTime Dt;

            string sDate = ((DateTime)DR_MASTER["cost_date"]).ToString("dd/MM/yyyy");
            string sCntr = "";
            string Str = "";

            iRow++;

            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "DEBIT NOTE";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "CREDIT NOTE";

            Lib.WriteMergeCell(WS, iRow++, 1, 12, 2, sTitle, "Calibri", 18, true, Color.Black, "C", "C", "TB", "THIN");

            iRow += 2;

            _Size = 13;

            Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 9 : 5), "NUMBER", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 10 : 6), DR_MASTER["COST_REFNO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            File_Display_Name = DR_MASTER["COST_REFNO"].ToString() + "-" + DR_MASTER["COST_FOLDERNO"].ToString() + ".xls";

            iRow++;

            Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_LINE1"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 9 : 5), "DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 10 : 6), sDate, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE2"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE3"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE4"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "WE DEBIT YOUR ACCOUNT FOR THE FOLLOWING";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "WE CREDIT YOUR ACCOUNT FOR THE FOLLOWING";

            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, sTitle, "Calibri", 12, true, Color.Black, "C", "C", "TB", "THIN");

            Lib.WriteData(WS, iRow, 1, "MAWB NO", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["HBL_BL_NO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            iRow++;
            Lib.WriteData(WS, iRow, 1, "HAWB NO", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

            Str = "";
            foreach (DataRow Dr in dt_house.Rows)
            {
                if (Str != "")
                    Str += ",";
                Str += Dr["hbl_bl_no"].ToString();
            }

            Lib.WriteMergeCell(WS, iRow, 3, 5, 1, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
            iRow++;

            Lib.WriteData(WS, iRow, 1, "CONSIGNEE", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

            int iCount = 0;
            Str = "";
            String PreStr = "1";
            foreach (DataRow Dr in dt_house.Select("1=1", "consignee_name"))
            {
                if (Str != "")
                    Str += "\n";

                if (PreStr != Dr["consignee_name"].ToString())
                {
                    PreStr = Dr["consignee_name"].ToString();
                    Str += Dr["consignee_name"].ToString();
                    iCount++;
                }
            }
            if (iCount == 0)
                iCount = 1;

            Lib.WriteMergeCell(WS, iRow, 3, 5, iCount, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
            iRow += iCount;

            Lib.WriteData(WS, iRow, 1, "MAWB DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, Lib.DatetoStringDisplayformat(DR_MASTER["HBL_DATE"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow, 1, "AIRPORT OF DEPARTURE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["POL_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;

            Lib.WriteData(WS, iRow, 1, "AIRPORT OF DESTINATION", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["POD_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;

            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, "", "Calibri", 11, true, Color.Black, "C", "C", "T", "THIN");

            iCol = 1;
            _Color = Color.Black;
            _Border = "";
            _Size = 12;

            decimal nAmt = 0;

            IsRemarksExist = false;
            foreach (DataRow Dr in dt_costdet2.Rows)
            {
                if (Dr["costd_remarks"].ToString().Trim().Length > 0)
                {
                    IsRemarksExist = true;
                    break;
                }
            }
            iRow += 4;
            Lib.WriteData(WS, iRow, 1, "PARTICULARS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            if (IsRemarksExist)
                Lib.WriteData(WS, iRow, 4, "REMARKS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 7, "AMOUNT(" + DR_MASTER["curr_code"].ToString() + ")", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            foreach (DataRow Dr in dt_costdet2.Rows)
            {
                iRow++;
                Lib.WriteData(WS, iRow, 1, Dr["costd_acc_name"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                if (IsRemarksExist)
                    Lib.WriteData(WS, iRow, 4, Dr["costd_remarks"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 7, Dr["costd_acc_amt"], _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            }

            iRow++;

            Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 7, DR_MASTER["cost_drcr_amount"], _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            iRow += 6;

            nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());

            if (nDrCRAmt < 0)
                nDrCRAmt = Math.Abs(nDrCRAmt);

            string sAmt = Lib.NumericFormat(nDrCRAmt.ToString(), 2);

            string sWords = "";
            if (DR_MASTER["curr_code"].ToString() != "INR")
                sWords = Number2Word_USD.Convert(sAmt, DR_MASTER["CURR_CODE"].ToString(), "CENTS");
            if (DR_MASTER["curr_code"].ToString() == "INR")
                sWords = Number2Word_RS.Convert(sAmt, "INR", "PAISE");


            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, sWords, "Calibri", 14, true, Color.Black, "L", "C", "TB", "THIN");
            Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, "E.&.O.E", "Calibri", 12, true, Color.Black, "L", "C", "TB", "THIN");
            if (Print_FC_Bank == "Y")
            {
                iRow++;
                iRow++;
                Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, "BANK DETAILS", "Calibri", 12, true, Color.Black, "L", "C", "B", "THIN");
                Lib.WriteData(WS, iRow, 1, "BENEFICIARY NAME", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Company, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "USD ACCOUNT NO", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Acno, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "BANK NAME", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Name, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "BANK ADDRESS", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Add1, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Add2, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "SWIFT CODE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow++, 3, Bank_Ifsc_Code, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 1, "CORRESPONDENT BANK", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, Bank_Add3, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                WS.Cells.GetSubrangeRelative(iRow++, 1, 7, 1).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
            }

            WS.Columns[12].Delete();
            WS.Columns[11].Delete();
            WS.Columns[10].Delete();
            WB.SaveXls(File_Name + ".xls");
        }
        private decimal GetConvertAmt(string StrAmt,string StrExRt)
        {
            decimal sAmt = Lib.Conv2Decimal(StrAmt);
            decimal exrt = Lib.Conv2Decimal(StrExRt);

            if (exrt>1)
            {
                sAmt = sAmt / exrt;
                sAmt = Lib.RoundNumber_Latest(sAmt.ToString(), 2, true);
            }
            return sAmt;
        }
    }
}

