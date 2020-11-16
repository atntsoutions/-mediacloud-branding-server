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
    public class ConsoleCostingService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;

        ExcelFile file = null;
        ExcelWorksheet ws = null;
        CellRange myCell;

        List<Costingd> InvList = new List<Costingd>();
        Costingd InvRow;

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
        private DataTable dt_bank = new DataTable();

        private DataTable DTP_COSTINGM;
        private DataTable DTP_COSTINGD;
        private DataTable DTP_DESTN;
        private DataTable dt_master;
        private DataTable dt_house;
        private DataTable dt_cntr;
        private DataTable dt_costDet;

        private DataRow DR_MASTER;

        private decimal FH_RATE1 = 0;
        private decimal FH_RATE2 = 0;
        private decimal FH_RATE3 = 0;
        private decimal FH_LIMIT1 = 0;
        private decimal FH_LIMIT2 = 0;
        private decimal INCENTIVE_RATE = 0;
        private decimal SEAL_EXPENSE = 0;
        private decimal MOTHER_DEST_EXPENSE = 0;
       
        private decimal OTH_CHRG_RITRA = 0;
        private string Master_POFD = "";
        private decimal HAULAGE_PER_CBM = Decimal.Parse("17.5");
        private decimal HAULAGE_MIN_RATE = Decimal.Parse("95.00");
        private decimal WEIGHT_DIVIDER = 250;
        private decimal DESTUFF_P_D = 495, HANDLING_FEE = 0;
        private decimal TRUK_COST = 0, CNTR_SHIFT_CHRGS = 0;
        private decimal VESSEl_CHANGE_CHRGS = 0;
        private decimal BUY_ACD_CC = 0;
        private decimal BUY_ACD_PP = 0;
        private decimal BUY_DDC_CC = 0;
        private decimal BUY_AMENMENT_CHRGS = 0;
        decimal TOT_B = 0, TOT_C = 0;
        private Boolean Is_haulage_Minimum = false;
        decimal TOT_CC = 0, TOT_PP = 0, TOT_HAULAGE = 0, TOT_REBATE = 0;
        decimal PROFIT = 0;


        decimal EXCHANGE_RATE = 40;
        string MBL_FRT_STATUS = "";
        decimal TOT_BUY_PREPAID = 0;
        decimal TOT_BUY_COLLECT = 0;
        decimal TOT_BUY = 0;
        public int TOTAL_FOLDERS = 0;
        private int TOTAL_CONSIGNEE = 0;
        private Dictionary<string, decimal> myDict = null;
        Dictionary<object,object> HT_SPL_INCENTIVE = null;
        Dictionary<object, object> HT_PRINT = null;
        private decimal TOT_NON_NOM_CBM_NO_INCENTVE = 0;
        private decimal TOT_NON_NOM_CBM_SPL_INCENTVE = 0;
        private decimal TOT_NON_NOM_CBM = 0;
        private decimal TOT_NOM_CBM = 0;
        private decimal TOT_MUTUAL_CBM = 0;
        private decimal TOT_CBM = 0;
        private decimal PER_CBM_RATE = 0;

        private decimal TOT_NON_NOM_PRO_CHRGS = 0;
        private decimal TOT_NOM_PRO_CHRGS = 0;
        private decimal TOT_MUTUAL_PRO_CHRGS = 0;
        private decimal BUY_FRT_PP = 0;
        private decimal EX_WORK_OTHERS = 0;
        private decimal EX_WORK_CHRGS = 0;
        decimal TOT_NETDUE = 0;

        decimal TotHandlingChrgs = 0, FhandCC = 0, TOT_A = 0, FRT_B = 0, NET_C = 0, Tot_CHRGS_CC = 0;
        decimal IncentiveLocal = 0, TOTAL_FROM_CHARGES = 0, FhandPP = 0, MutualPP = 0;
        decimal TOTAL_TO_CHARGES = 0, TOT_A2 = 0, FRT_B2 = 0, NET_C2 = 0, Tot_CHRGS_PP = 0;

        private string AGENT_FORMAT = "";
        LovService lov = null;

        //using Database Columns for air
       //WRS - BAF
       //MYC - CAF
       //MCC - DDC
       //SRC - ADC

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
                sWhere += " and a.cost_source = 'SE CONSOLE COSTING' ";
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
            //using Database Columns of air
            //WRS - BAF
            //MYC - CAF
            //MCC - DDC
            //SRC - ADC

            string SQL = "";
            string Previous_CostData = "";
            decimal exwork = 0, rebate = 0;
            decimal buy_pp = 0, buy_cc = 0, buy_tot = 0;
            decimal sell_pp = 0, sell_cc = 0, sell_tot = 0;
            decimal tot_mbl_cbm = 0, tot_mbl_grwt = 0, tot_mbl_chwt = 0;

            decimal mbuy_pp = 0, mbuy_cc = 0, mbuy_tot = 0;
            decimal msell_pp = 0, msell_cc = 0, msell_tot = 0;

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, int> CngeGrp = new Dictionary<string, int>();
            Costingm mRow = new Costingm();
            List<Costingd> mList = new List<Costingd>();
            Costingd dRow = new Costingd();

            string costingid = SearchData["pkid"].ToString();
            string agent = SearchData["agent"].ToString();
            string cntrs = SearchData["cntrs"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string id = "";
            
            try
            {
                UpdateMasterRate(costingid, branch_code, agent, cntrs); // Cost Master Rate

                if (agent.ToString().Contains("RITRA"))
                    AGENT_FORMAT = "RITRA";
                else if (agent.ToString().Contains("TRAFFIC TECH"))
                    AGENT_FORMAT = "TRAFFIC-TECH";

                Con_Oracle = new DBConnection();

                sql = "select cr_rate_code,cr_rate_value from consolerate ";
                sql += " where cr_rate_type='MINRATE' ";
                DataTable dt_temp = new DataTable();
                dt_temp = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow dr in dt_temp.Rows)
                {
                    if (dr["cr_rate_code"].ToString() == "FREEHAND_RT1")
                        FH_RATE1 = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "FREEHAND_RT2")
                        FH_RATE2 = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "FREEHAND_RT3")
                        FH_RATE3 = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "FREEHAND_LT1")
                        FH_LIMIT1 = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "FREEHAND_LT2")
                        FH_LIMIT2 = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "INCENT_RATE")
                        INCENTIVE_RATE = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "HAULG_MIN_WELTN")
                        HAULAGE_MIN_RATE = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "SEAL_EXP_ACTN")
                        SEAL_EXPENSE = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "DEST_EXP_MOTHER")
                        MOTHER_DEST_EXPENSE = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                    else if (dr["cr_rate_code"].ToString() == "HAULAGE_PER_CBM")
                        HAULAGE_PER_CBM = Lib.Convert2Decimal(dr["cr_rate_value"].ToString());
                }

                SQL = " select cost_mblid from costingm  where cost_pkid = '" + costingid + "'";
                DataTable Dt_costing = new DataTable();
                Dt_costing = Con_Oracle.ExecuteQuery(SQL);
                foreach (DataRow Dr in Dt_costing.Rows)
                {
                    id = Dr["cost_mblid"].ToString();
                    break;
                }


                SQL = "  select hbl_pkid,";
                SQL += "  0 as GRWT,";
                SQL += "  0 as CHWT, ";
                SQL += "  0 as CBM, ";
                SQL += "  max(HBL_NO) as HBL_NO,";
                SQL += "  max(HBL_BL_NO) as HBL_BL_NO, ";
                SQL += "  max(HBL_TERMS) as HBL_TERMS,";
                SQL += "  max(HBL_NOMINATION) as HBL_NOMINATION,";
                SQL += "  max(POFD_NAME) as POFD_NAME,";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105001') then jv_ftotal else 0 end) as FRT_PP,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105001') then jv_ftotal else 0 end) as FRT_CC, ";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105002') then jv_ftotal else 0 end) as WRS_PP_BAF,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105002') then jv_ftotal else 0 end) as WRS_CC_BAF, ";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105003') then jv_ftotal else 0 end) as MYC_PP_CAF,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105003') then jv_ftotal else 0 end) as MYC_CC_CAF, ";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105011') then jv_ftotal else 0 end) as MCC_PP_DDC,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105011') then jv_ftotal else 0 end) as MCC_CC_DDC, ";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1106001') then jv_ftotal else 0 end) as SRC_PP_ADC,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1106001') then jv_ftotal else 0 end) as SRC_CC_ADC, ";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1106019', '1105022','1105016','1105015') then jv_ftotal else 0 end) as OTH_PP,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1106019','1105022','1105016','1105015') then jv_ftotal else 0 end) as OTH_CC ";
                SQL += "  from (";
                SQL += "    select mbl.hbl_pkid,mbl.hbl_no,mbl.hbl_bl_no,mbl.hbl_terms as hbl_terms";
                SQL += "    ,mbl.hbl_nomination as hbl_nomination, cast('PREPAID' as nvarchar2(10)) as status";
                SQL += "    ,jv_qty,jv_ftotal,acc_code,pofd.param_name as pofd_name  ";
                SQL += "    from ledgerh a";
                SQL += "    inner join ledgert b on a.jvh_pkid =  b.jv_parent_id";
                SQL += "    inner join hblm mbl on a.jvh_cc_id = mbl.hbl_pkid ";
                SQL += "    inner join param curr on jv_curr_id = curr.param_pkid";
                SQL += "    left join acctm ac on b.jv_acc_id = ac.acc_pkid";
                SQL += "    left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid ";
                SQL += "    where jvh_cc_id ='" + id + "' and jv_row_type = 'DR-LEDGER' and curr.param_code <> 'INR'";
                SQL += "    union all   ";
                SQL += "    select mbl.hbl_pkid,mbl.hbl_no,mbl.hbl_bl_no,mbl.hbl_terms as hbl_terms";
                SQL += "    ,mbl.hbl_nomination as hbl_nomination, cast('COLLECT' as nvarchar2(10)) as status";
                SQL += "    ,inv_qty, inv_ftotal,acc_code,pofd.param_name as pofd_name   ";
                SQL += "    from jobincome a";
                SQL += "    inner join hblm mbl on a.inv_parent_id = mbl.hbl_pkid";
                SQL += "    inner join param curr on inv_curr_id = curr.param_pkid ";
                SQL += "    left join acctm ac on a.inv_acc_id = ac.acc_pkid ";
                SQL += "    left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid ";
                SQL += "   where inv_parent_id ='" + id + "'  and curr.param_code <> 'INR'";
                SQL += "  ) a group by hbl_pkid";
                DataTable Dt_MblExpense = Con_Oracle.ExecuteQuery(SQL);

                // Hbl Income/Expense
                SQL = "";
                SQL = " select hbl_pkid,hbl_bl_no,";
                SQL += "  max(HBL_GRWT) as GRWT,";
                SQL += "  max(HBL_CHWT) as CHWT, ";
                SQL += "  max(HBL_CBM) as CBM, ";
                SQL += "  max(HBL_NO) as HBL_NO,";
                SQL += "  max(SHIPPER_NAME) as SHIPPER_NAME,";
                SQL += "  max(CONSIGNEE_NAME) as CONSIGNEE_NAME,";
                SQL += "  max(HBL_TERMS) as HBL_TERMS,";
                SQL += "  max(HBL_NOMINATION) as HBL_NOMINATION,";
                SQL += "  max(POFD_NAME) as POFD_NAME,";

                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105001')  then inv_ftotal else 0 end) as FRT_PP,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105001')  then inv_ftotal else 0 end) as FRT_CC,";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105001')  then inv_rate else 0 end) as FRT_RATE_PP,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105001')  then inv_rate else 0 end) as FRT_RATE_CC,";

                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105011')  then inv_ftotal else 0 end) as MCC_PP_DDC,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105011')  then inv_ftotal else 0 end) as MCC_CC_DDC,";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1105011')  then inv_rate else 0 end) as MCC_RATE_PP_DDC,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1105011')  then inv_rate else 0 end) as MCC_RATE_CC_DDC,";

                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1106001')  then inv_ftotal else 0 end) as SRC_PP_ADC,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1106001')  then inv_ftotal else 0 end) as SRC_CC_ADC,";
                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1106001')  then inv_rate else 0 end) as SRC_RATE_PP_ADC,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1106001')  then inv_rate else 0 end) as SRC_RATE_CC_ADC,";

                SQL += "  sum(case when status = 'PREPAID' and acc_code in ('1106019','1105005','1105010','1105017','1105020','1105021','1105025')  then inv_ftotal else 0 end) as OTH_PP,";
                SQL += "  sum(case when status = 'COLLECT' and acc_code in ('1106019','1105005','1105010','1105017','1105020','1105021','1105025')  then inv_ftotal else 0 end) as OTH_CC,";
                SQL += "  sum(inv_rebate_amt) as rebate ";
                SQL += "  from (";
                SQL += " select hbl.hbl_pkid,hbl.hbl_no,hbl.hbl_bl_no,hbl.hbl_grwt,hbl.hbl_chwt,hbl.hbl_cbm";
                SQL += "  , inv_type as status,inv_rate, inv_ftotal, inv_rebate_amt,acc_code ";
                SQL += "  ,shpr.cust_name as shipper_name,cnge.cust_name as consignee_name,hbl.hbl_terms ,nvl(hbl.hbl_nomination,mbl.hbl_nomination) as hbl_nomination";
                SQL += "  ,pofd.param_name as pofd_name";
                SQL += "  from jobincome a";
                SQL += "  inner join hblm hbl on a.inv_parent_id = hbl.hbl_pkid";
                SQL += "  inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid ";
                SQL += "  left join acctm ac on a.inv_acc_id = ac.acc_pkid ";
                SQL += "  left join customerm shpr on hbl.hbl_exp_id = shpr.cust_pkid ";
                SQL += "  left join customerm cnge on hbl.hbl_imp_id = cnge.cust_pkid ";
                SQL += "  left join param pofd on hbl.hbl_pofd_id = pofd.param_pkid ";
                SQL += "  where mbl.hbl_pkid ='" + id + "' and inv_source <> 'EX-WORK'";
                SQL += "  ) a  group by hbl_pkid,hbl_bl_no order by hbl_bl_no";
                SQL += "  ";

                DataTable Dt_HblIncome = Con_Oracle.ExecuteQuery(SQL);
                foreach (DataRow Dr in Dt_HblIncome.Rows)
                {
                    if (!CngeGrp.ContainsKey(Dr["CONSIGNEE_NAME"].ToString()))
                        CngeGrp.Add(Dr["CONSIGNEE_NAME"].ToString(), CngeGrp.Count + 1);
                }

                SQL = "";
                SQL = "select sum(hbl_cbm) as cbm,sum(hbl_grwt) as grwt,sum(hbl_chwt) as chwt from hblm where hbl_mbl_id='" + id + "'";
                DataTable Dt_Mbl = Con_Oracle.ExecuteQuery(SQL);
                if (Dt_Mbl.Rows.Count > 0)
                {
                    tot_mbl_cbm = Lib.Conv2Decimal(Lib.NumericFormat(Dt_Mbl.Rows[0]["cbm"].ToString(), 3));
                    tot_mbl_grwt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_Mbl.Rows[0]["grwt"].ToString(), 3));
                    tot_mbl_chwt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_Mbl.Rows[0]["chwt"].ToString(), 3));
                }

                // Hbl Expense // EXWORK
                SQL = "";
                SQL += " select sum(round(jv_net_total / jv_exrate,2)) as exwork ";
                SQL += " from jobincome a";
                SQL += " inner join hblm hbl on a.inv_parent_id = hbl.hbl_pkid";
                SQL += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid ";
                SQL += " inner join param curr on inv_curr_id = curr.param_pkid";
                SQL += " left join ledgert l on inv_pkid = jv_pkid ";
                SQL += " where mbl.hbl_pkid ='" + id + "' and curr.param_code <> 'INR' and inv_source = 'EX-WORK'";
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
                    dRow.costd_shipper_name = "";
                    dRow.costd_consignee_name = "";
                    dRow.costd_hbl_nomination = Dr["hbl_nomination"].ToString();
                    dRow.costd_hbl_terms = Dr["hbl_terms"].ToString();
                    dRow.costd_pofd_name = Dr["pofd_name"].ToString();
                    dRow.costd_grwt = tot_mbl_grwt;
                    dRow.costd_chwt = tot_mbl_chwt;
                    dRow.costd_actual_cbm = tot_mbl_cbm;
                    dRow.costd_cbm = tot_mbl_cbm;
                    dRow.costd_frt_pp = Lib.Conv2Decimal(Dr["frt_pp"].ToString());
                    dRow.costd_frt_cc = Lib.Conv2Decimal(Dr["frt_cc"].ToString());
                    dRow.costd_frt_rate_pp = 0;
                    dRow.costd_frt_rate_cc = 0;
                    dRow.costd_baf_pp = Lib.Conv2Decimal(Dr["wrs_pp_baf"].ToString());
                    dRow.costd_baf_cc = Lib.Conv2Decimal(Dr["wrs_cc_baf"].ToString());
                    dRow.costd_baf_rate_pp = 0;
                    dRow.costd_baf_rate_cc = 0;
                    dRow.costd_caf_pp = Lib.Conv2Decimal(Dr["myc_pp_caf"].ToString());
                    dRow.costd_caf_cc = Lib.Conv2Decimal(Dr["myc_cc_caf"].ToString());
                    dRow.costd_caf_rate_pp = 0;
                    dRow.costd_caf_rate_cc = 0;
                    dRow.costd_ddc_pp = Lib.Conv2Decimal(Dr["mcc_pp_ddc"].ToString());
                    dRow.costd_ddc_cc = Lib.Conv2Decimal(Dr["mcc_cc_ddc"].ToString());
                    dRow.costd_ddc_rate_pp = 0;
                    dRow.costd_ddc_rate_cc = 0;
                    dRow.costd_acd_pp = Lib.Conv2Decimal(Dr["src_pp_adc"].ToString());
                    dRow.costd_acd_cc = Lib.Conv2Decimal(Dr["src_cc_adc"].ToString());
                    dRow.costd_acd_rate_pp = 0;
                    dRow.costd_acd_rate_cc = 0;
                    dRow.costd_oth_pp = Lib.Conv2Decimal(Dr["oth_pp"].ToString());
                    dRow.costd_oth_cc = Lib.Conv2Decimal(Dr["oth_cc"].ToString());
                    dRow.costd_oth_rate_pp = 0;
                    dRow.costd_oth_rate_cc = 0;
                    dRow.costd_ctr = 1;
                    dRow.costd_agent_format = AGENT_FORMAT;

                    dRow.costd_incentive_rate = INCENTIVE_RATE;
                    dRow.costd_fh_limit1 = FH_LIMIT1;
                    dRow.costd_fh_rate1 = FH_RATE1;
                    dRow.costd_fh_limit2 = FH_LIMIT2;
                    dRow.costd_fh_rate2 = FH_RATE2;
                    dRow.costd_fh_limit3 = FH_LIMIT2;
                    dRow.costd_fh_rate3 = FH_RATE3;
                    dRow.costd_rebate = 0;
                    dRow.costd_amenment_chrgs = 0;
                    dRow.costd_seal_chrgs = 0;
                    dRow.costd_fh_chrg_perhouse = 0;
                    dRow.costd_oth_chrgs_ritra = 0;
                    dRow.costd_ex_chrg_ritrahouse = 0;
                    dRow.costd_spl_incentive_rate = 0;
                    dRow.costd_incentive_notreceived = false;
                    dRow.costd_house_notinclude = false;


                    buy_pp = dRow.costd_frt_pp;
                    buy_pp += dRow.costd_baf_pp;
                    buy_pp += dRow.costd_caf_pp;
                    buy_pp += dRow.costd_ddc_pp;
                    buy_pp += dRow.costd_acd_pp;
                    buy_pp += dRow.costd_oth_pp;
                    buy_pp = Lib.Conv2Decimal(Lib.NumericFormat(buy_pp.ToString(), 2));
                    mbuy_pp += buy_pp;

                    buy_cc = dRow.costd_frt_cc;
                    buy_cc += dRow.costd_baf_cc;
                    buy_cc += dRow.costd_caf_cc;
                    buy_cc += dRow.costd_ddc_cc;
                    buy_cc += dRow.costd_acd_cc;
                    buy_cc += dRow.costd_oth_cc;
                    buy_cc = Lib.Conv2Decimal(Lib.NumericFormat(buy_cc.ToString(), 2));
                    mbuy_cc += buy_cc;

                    buy_tot = buy_pp + buy_cc;
                    buy_tot = Lib.Conv2Decimal(Lib.NumericFormat(buy_tot.ToString(), 2));
                    mbuy_tot += buy_tot;

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
                    dRow.costd_shipper_name = Dr["shipper_name"].ToString();
                    dRow.costd_consignee_name = Dr["consignee_name"].ToString();
                    if (CngeGrp.ContainsKey(Dr["consignee_name"].ToString()))
                        dRow.costd_consignee_group = "C" + CngeGrp[Dr["consignee_name"].ToString()].ToString();
                    else
                        dRow.costd_consignee_group = "";
                    dRow.costd_hbl_nomination = Dr["hbl_nomination"].ToString();
                    dRow.costd_hbl_terms = Dr["hbl_terms"].ToString();
                    dRow.costd_pofd_name = Dr["pofd_name"].ToString();

                    dRow.costd_grwt = Lib.Conv2Decimal(Dr["grwt"].ToString());
                    dRow.costd_chwt = Lib.Conv2Decimal(Dr["chwt"].ToString());
                    dRow.costd_cbm = Lib.Conv2Decimal(Dr["cbm"].ToString());
                    dRow.costd_actual_cbm = Lib.Conv2Decimal(Dr["cbm"].ToString());

                    dRow.costd_frt_pp = Lib.Conv2Decimal(Dr["frt_pp"].ToString());
                    dRow.costd_frt_cc = Lib.Conv2Decimal(Dr["frt_cc"].ToString());
                    dRow.costd_frt_rate_pp = Lib.Conv2Decimal(Dr["frt_rate_pp"].ToString());
                    dRow.costd_frt_rate_cc = Lib.Conv2Decimal(Dr["frt_rate_cc"].ToString());
                    dRow.costd_baf_pp = 0;
                    dRow.costd_baf_cc = 0;
                    dRow.costd_baf_rate_pp = 0;
                    dRow.costd_baf_rate_cc = 0;
                    dRow.costd_caf_pp = 0;
                    dRow.costd_caf_cc = 0;
                    dRow.costd_caf_rate_pp = 0;
                    dRow.costd_caf_rate_cc = 0;
                    dRow.costd_ddc_pp = Lib.Conv2Decimal(Dr["mcc_pp_ddc"].ToString());
                    dRow.costd_ddc_cc = Lib.Conv2Decimal(Dr["mcc_cc_ddc"].ToString());
                    dRow.costd_ddc_rate_pp = Lib.Conv2Decimal(Dr["mcc_rate_pp_ddc"].ToString());
                    dRow.costd_ddc_rate_cc = Lib.Conv2Decimal(Dr["mcc_rate_cc_ddc"].ToString());
                    dRow.costd_acd_pp = Lib.Conv2Decimal(Dr["src_pp_adc"].ToString());
                    dRow.costd_acd_cc = Lib.Conv2Decimal(Dr["src_cc_adc"].ToString());
                    dRow.costd_acd_rate_pp = Lib.Conv2Decimal(Dr["src_rate_pp_adc"].ToString());
                    dRow.costd_acd_rate_cc = Lib.Conv2Decimal(Dr["src_rate_cc_adc"].ToString());
                    dRow.costd_oth_pp = Lib.Conv2Decimal(Dr["oth_pp"].ToString());
                    dRow.costd_oth_cc = Lib.Conv2Decimal(Dr["oth_cc"].ToString());
                    dRow.costd_oth_rate_pp = 0;
                    dRow.costd_oth_rate_cc = 0;
                    dRow.costd_ctr = iCtr;
                    dRow.costd_agent_format = AGENT_FORMAT;

                    dRow.costd_incentive_rate = INCENTIVE_RATE;
                    dRow.costd_fh_limit1 = FH_LIMIT1;
                    dRow.costd_fh_rate1 = FH_RATE1;
                    dRow.costd_fh_limit2 = FH_LIMIT2;
                    dRow.costd_fh_rate2 = FH_RATE2;
                    dRow.costd_fh_limit3 = FH_LIMIT2;
                    dRow.costd_fh_rate3 = FH_RATE3;
                    dRow.costd_rebate = Lib.Convert2Decimal(Dr["rebate"].ToString());
                    dRow.costd_amenment_chrgs = 0;
                    dRow.costd_seal_chrgs = 0;
                    dRow.costd_fh_chrg_perhouse = 0;
                    dRow.costd_oth_chrgs_ritra = 0;
                    dRow.costd_ex_chrg_ritrahouse = 0;
                    dRow.costd_spl_incentive_rate = 0;
                    dRow.costd_incentive_notreceived = false;
                    dRow.costd_house_notinclude = false;


                    sell_pp = dRow.costd_frt_pp;
                    sell_pp += dRow.costd_baf_pp;
                    sell_pp += dRow.costd_caf_pp;
                    sell_pp += dRow.costd_ddc_pp;
                    sell_pp += dRow.costd_acd_pp;
                    sell_pp += dRow.costd_oth_pp;
                    sell_pp = Lib.Conv2Decimal(Lib.NumericFormat(sell_pp.ToString(), 2));
                    msell_pp += sell_pp;

                    sell_cc = dRow.costd_frt_cc;
                    sell_cc += dRow.costd_baf_cc;
                    sell_cc += dRow.costd_caf_cc;
                    sell_cc += dRow.costd_ddc_cc;
                    sell_cc += dRow.costd_acd_cc;
                    sell_cc += dRow.costd_oth_cc;
                    sell_cc = Lib.Conv2Decimal(Lib.NumericFormat(sell_cc.ToString(), 2));
                    msell_cc += sell_cc;

                    sell_tot = sell_pp + sell_cc;
                    sell_tot = Lib.Conv2Decimal(Lib.NumericFormat(sell_tot.ToString(), 2));
                    msell_tot += sell_tot;

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

                //InvList = new List<Costingd>();
                //decimal Tot_InvoiceAmt = 0;
                //foreach (DataRow Dr in Dt_Invoice.Rows)
                //{
                //    dRow = new Costingd();
                //    dRow.costd_pkid = Dr["costd_pkid"].ToString();
                //    dRow.costd_parent_id = Dr["costd_parent_id"].ToString();
                //    dRow.costd_acc_id = Dr["costd_acc_id"].ToString();
                //    dRow.costd_acc_name = Dr["costd_acc_name"].ToString();
                //    dRow.costd_blno = Dr["costd_blno"].ToString();
                //    dRow.costd_acc_qty = Lib.Conv2Decimal(Dr["costd_acc_qty"].ToString());
                //    dRow.costd_acc_rate = Lib.Conv2Decimal(Dr["costd_acc_rate"].ToString());
                //    dRow.costd_acc_amt = Lib.Conv2Decimal(Dr["costd_acc_amt"].ToString());
                //    dRow.costd_ctr = Lib.Conv2Integer(Dr["costd_ctr"].ToString());
                //    dRow.costd_category = Dr["costd_category"].ToString();
                //    mList.Add(dRow);
                //}


                Dt_costing.Rows.Clear();
                Dt_MblExpense.Rows.Clear();
                Dt_HblIncome.Rows.Clear();
                Dt_ExWork.Rows.Clear();
                //Dt_Invoice.Rows.Clear();

                mRow.cost_pkid = costingid;
                mRow.DetailList2 = mList;
                //  mRow.cost_tot_acc_amt = Tot_InvoiceAmt;
                //  mRow.DetailList = InvList;
                GenerateCostingDet(mRow);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("list", mList);
            RetData.Add("buypp", mbuy_pp);
            RetData.Add("buycc", mbuy_cc);
            RetData.Add("buytot", mbuy_tot);
            RetData.Add("sellpp", msell_pp);
            RetData.Add("sellcc", msell_cc);
            RetData.Add("selltot", msell_tot);
            RetData.Add("rebate", rebate);
            RetData.Add("exwork", exwork);
            RetData.Add("handlingcharge", TotHandlingChrgs);
            RetData.Add("expense", TOTAL_TO_CHARGES);
            RetData.Add("income", TOTAL_FROM_CHARGES);
            RetData.Add("drcramt", TOT_NETDUE);
            RetData.Add("tprofit", TOT_A2);
            RetData.Add("sprofit", PROFIT);
            RetData.Add("format", AGENT_FORMAT);
            return RetData;
        }

        private void UpdateMasterRate(string CostID,string branch_code,string AgentName,string Cntrs)
        {
            try
            {
                string SQL = "";
                Con_Oracle = new DBConnection();

                sql = "select * from  consolerate where cr_rate_type='DESTIN' ";
                sql += " and cr_branch_code='" + branch_code + "'";
                if (AgentName.Contains("ACTION"))
                sql += " and cr_agent_name ='ACTION'";
                else if (AgentName.Contains("GATE4EU"))
                    sql += " and cr_agent_name ='GATE4EU'";
                else if (AgentName.Contains("MOTHERLINES"))
                    sql += " and cr_agent_name ='MOTHERLINES'";
                else if (AgentName.Contains("TRAFFIC TECH INTERNATIONAL"))
                    sql += " and cr_agent_name ='WELLTON'";
                else
                    sql += " and 1=4 ";
                if (Cntrs.Contains("/20"))
                    sql += " and cr_cntr_type='20' ";
                else if (Cntrs.Contains("/40"))
                    sql += " and cr_cntr_type='40' ";
                else
                    sql += " and 1=2";
                DataTable Dt_RateM = new DataTable();
                Dt_RateM = Con_Oracle.ExecuteQuery(sql);
                if (Dt_RateM.Rows.Count > 0)
                {
                    SQL = " UPDATE COSTINGM SET ";
                    SQL += " EX_RATE_GBP =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_EX_RATE_GBP"].ToString()) + ",";
                    SQL += " ORG_INC_THC =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_INC_THC"].ToString()) + ",";
                    SQL += " ORG_INC_BL =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_INC_BL"].ToString()) + ",";
                    SQL += " ORG_EXP_THC =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_THC"].ToString()) + ",";
                    SQL += " ORG_EXP_EMTYPLCE =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_EMTYPLCE"].ToString()) + ",";
                    SQL += " ORG_EXP_MISC =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_MISC"].ToString()) + ",";
                    SQL += " ORG_EXP_STUFF =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_STUFF"].ToString()) + ",";
                    SQL += " ORG_EXP_TRANS =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_TRANS"].ToString()) + ",";
                    SQL += " ORG_EXP_SURREND =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_SURREND"].ToString()) + ",";
                    SQL += " DES_INC_THC =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_INC_THC"].ToString()) + ",";
                    SQL += " DES_INC_BL =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_INC_BL"].ToString()) + ",";
                    SQL += " DES_EXP_TERML =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_TERML"].ToString()) + ",";
                    SQL += " DES_EXP_BL =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_BL"].ToString()) + ",";
                    SQL += " DES_EXP_SHUNT =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_SHUNT"].ToString()) + ",";
                    SQL += " DES_EXP_UNPACK =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_UNPACK"].ToString()) + ",";
                    SQL += " DES_EXP_LOLO =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_LOLO"].ToString()) + ",";
                    SQL += " DES_EXP_SECURTY =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_SECURTY"].ToString()) + ",";
                    SQL += " DES_INC_HNDG_TON =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_INC_HNDG_TON"].ToString()) + ",";
                    SQL += " DES_INC_HNDG_CBM =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_INC_HNDG_CBM"].ToString()) + ",";
                    SQL += " ORG_EXP_CFS  =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_CFS"].ToString()) + ",";
                    SQL += " ORG_EXP_SURVEY  =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_SURVEY"].ToString()) + ",";
                    SQL += " ORG_EXP_CSEAL  =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_ORG_EXP_CSEAL"].ToString()) + ",";
                    SQL += " DES_EXP_ISPS  =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_ISPS"].ToString()) + ",";
                  
                    SQL += " DES_EXP_TPW  =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_TPW"].ToString()) + ",";
                    SQL += " DES_EXP_TDOC  =" + Lib.Convert2Decimal(Dt_RateM.Rows[0]["CR_DES_EXP_TDOC"].ToString());
                    SQL += " WHERE COST_PKID = '" + CostID + "' AND REC_CATEGORY='SEA EXPORT'";

                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    
                }
                
                Con_Oracle.CloseConnection();
                Dt_RateM.Rows.Clear();
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

        }

        public Dictionary<string, object> GenerateCostingDet(Costingm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            decimal mbuy_pp = 0, mbuy_cc = 0, mbuy_tot = 0;
            decimal msell_pp = 0, msell_cc = 0, msell_tot = 0, rebate = 0, exwork = 0;
            try
            {
                Con_Oracle = new DBConnection();
                Con_Oracle.BeginTransaction();
                SaveCostingDet(Record, "GENERATE");
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                LoadTables(Record.cost_pkid);
                
                mbuy_cc = 0;mbuy_pp = 0;mbuy_tot = 0;
                foreach(DataRow dr in DTP_COSTINGM.Rows)
                {
                    mbuy_cc += Lib.Conv2Decimal(dr["costd_cc"].ToString());
                    mbuy_pp += Lib.Conv2Decimal(dr["costd_pp"].ToString());
                    mbuy_tot += Lib.Conv2Decimal(dr["costd_tot"].ToString());
                }
                msell_cc = 0; msell_pp = 0; msell_tot = 0;rebate = 0;
                foreach (DataRow dr in DTP_COSTINGD.Rows)
                {
                    rebate += Lib.Conv2Decimal(dr["costd_rebate"].ToString());
                    msell_cc += Lib.Conv2Decimal(dr["costd_cc"].ToString());
                    msell_pp += Lib.Conv2Decimal(dr["costd_pp"].ToString());
                    msell_tot += Lib.Conv2Decimal(dr["costd_tot"].ToString());
                }
                foreach (DataRow Dr in DTP_DESTN.Rows)
                {
                    exwork = Lib.Convert2Decimal(Dr["cost_ex_works"].ToString());
                    break;
                }
                if (exwork == 0)
                    exwork = Record.cost_ex_works;
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
            
            RetData.Add("list", Record.DetailList2);
            RetData.Add("buypp", mbuy_pp);
            RetData.Add("buycc", mbuy_cc);
            RetData.Add("buytot", mbuy_tot);
            RetData.Add("sellpp", msell_pp);
            RetData.Add("sellcc", msell_cc);
            RetData.Add("selltot", msell_tot);
            RetData.Add("rebate", rebate);
            RetData.Add("exwork", exwork);
            RetData.Add("handlingcharge", TotHandlingChrgs);
            RetData.Add("expense", TOTAL_TO_CHARGES);
            RetData.Add("income", TOTAL_FROM_CHARGES);
            RetData.Add("drcramt", TOT_NETDUE);
            RetData.Add("tprofit", TOT_A2);
            RetData.Add("sprofit", PROFIT);
            RetData.Add("format", AGENT_FORMAT);
            return RetData;
        }

        private void SaveCostingDet(Costingm Record,string sType)
        {
            //using Database Columns of air
            //WRS - BAF
            //MYC - CAF
            //MCC - DDC
            //SRC - ADC

            string CostData = "";
            DBRecord Rec = new DBRecord();

            sql = "Delete from Costingd where costd_parent_id = '" + Record.cost_pkid + "'";
            if (sType == "GENERATE")
                sql += " and costd_category ='COSTING'";

            Con_Oracle.ExecuteNonQuery(sql);
            int iCtr = 0;
            foreach (Costingd Row in Record.DetailList2)
            {
                iCtr++;
                Rec.CreateRow("Costingd", "ADD", "costd_pkid", Row.costd_pkid);
                Rec.InsertString("costd_parent_id", Record.cost_pkid);
                Rec.InsertString("costd_acc_id", Row.costd_acc_id);
                Rec.InsertString("costd_acc_name", Row.costd_blno);
                Rec.InsertString("costd_type", Row.costd_type);
                Rec.InsertNumeric("costd_grwt", Row.costd_grwt.ToString());
                Rec.InsertNumeric("costd_chwt", Row.costd_chwt.ToString());
                Rec.InsertNumeric("costd_cbm", Row.costd_cbm.ToString());
                Rec.InsertNumeric("costd_frt_pp", Row.costd_frt_pp.ToString());
                Rec.InsertNumeric("costd_frt_cc", Row.costd_frt_cc.ToString());
                Rec.InsertNumeric("costd_frt_rate_pp", Row.costd_frt_rate_pp.ToString());
                Rec.InsertNumeric("costd_frt_rate_cc", Row.costd_frt_rate_cc.ToString());
                Rec.InsertNumeric("costd_wrs_pp", Row.costd_baf_pp.ToString());
                Rec.InsertNumeric("costd_wrs_cc", Row.costd_baf_cc.ToString());
                Rec.InsertNumeric("costd_wrs_rate_pp", Row.costd_baf_rate_pp.ToString());
                Rec.InsertNumeric("costd_wrs_rate_cc", Row.costd_baf_rate_cc.ToString());
                Rec.InsertNumeric("costd_myc_pp", Row.costd_caf_pp.ToString());
                Rec.InsertNumeric("costd_myc_cc", Row.costd_caf_cc.ToString());
                Rec.InsertNumeric("costd_myc_rate_pp", Row.costd_caf_rate_pp.ToString());
                Rec.InsertNumeric("costd_myc_rate_cc", Row.costd_caf_rate_cc.ToString());
                Rec.InsertNumeric("costd_mcc_pp", Row.costd_ddc_pp.ToString());
                Rec.InsertNumeric("costd_mcc_cc", Row.costd_ddc_cc.ToString());
                Rec.InsertNumeric("costd_mcc_rate_pp", Row.costd_ddc_rate_pp.ToString());
                Rec.InsertNumeric("costd_mcc_rate_cc", Row.costd_ddc_rate_cc.ToString());
                Rec.InsertNumeric("costd_src_pp", Row.costd_acd_pp.ToString());
                Rec.InsertNumeric("costd_src_cc", Row.costd_acd_cc.ToString());
                Rec.InsertNumeric("costd_src_rate_pp", Row.costd_acd_rate_pp.ToString());
                Rec.InsertNumeric("costd_src_rate_cc", Row.costd_acd_rate_cc.ToString());
                Rec.InsertNumeric("costd_oth_pp", Row.costd_oth_pp.ToString());
                Rec.InsertNumeric("costd_oth_cc", Row.costd_oth_cc.ToString());
                Rec.InsertNumeric("costd_oth_rate_pp", Row.costd_oth_rate_pp.ToString());
                Rec.InsertNumeric("costd_oth_rate_cc", Row.costd_oth_rate_cc.ToString());
                Rec.InsertNumeric("costd_pp", Row.costd_pp.ToString());
                Rec.InsertNumeric("costd_cc", Row.costd_cc.ToString());
                Rec.InsertNumeric("costd_tot", Row.costd_tot.ToString());
                Rec.InsertNumeric("costd_ctr", iCtr.ToString());
                Rec.InsertString("costd_category", "COSTING");

                CostData = Row.costd_hbl_nomination == null ? "" : Row.costd_hbl_nomination;
                CostData += ",";
                CostData += Row.costd_hbl_terms == null ? "" : Row.costd_hbl_terms;
                CostData += ",";
                CostData += Row.costd_incentive_rate.ToString();
                CostData += ",";
                CostData += Row.costd_fh_rate1.ToString();
                CostData += ",";
                CostData += Row.costd_fh_limit1.ToString();
                CostData += ",";
                CostData += Row.costd_fh_rate2.ToString();
                CostData += ",";
                CostData += Row.costd_fh_limit2.ToString();
                CostData += ",";
                CostData += Row.costd_fh_rate3.ToString();
                CostData += ",";
                CostData += Row.costd_pofd_name == null ? "" : Row.costd_pofd_name.Replace(",", "-");
                CostData += ",";
                CostData += Row.costd_oth_chrgs_ritra.ToString();
                CostData += ",";
                CostData += ((Row.costd_incentive_notreceived) ? "Y" : "N");
                CostData += ",";
                CostData += Row.costd_fh_chrg_perhouse.ToString();
                CostData += ",";
                CostData += Row.costd_spl_incentive_rate.ToString();
                CostData += ",";
                CostData += ((Row.costd_house_notinclude) ? "Y" : "N");
                CostData += ",";
                CostData += Row.costd_ex_chrg_ritrahouse.ToString();
                CostData += ",";
                CostData += Row.costd_haulage_per_cbm.ToString();
                CostData += ",";
                CostData += Row.costd_haulage_min_rate.ToString();
                CostData += ",";
                CostData += Row.costd_haulage_wt_divider.ToString();
                CostData += ",";
                CostData += Row.costd_destuff_pd.ToString();
                CostData += ",";
                CostData += Row.costd_handling_fee.ToString();
                CostData += ",";
                CostData += Row.costd_truck_cost.ToString();
                CostData += ",";
                CostData += Row.costd_cntr_shifit.ToString();
                CostData += ",";
                CostData += Row.costd_vessel_chrgs.ToString();
                CostData += ",";
                CostData += Row.costd_ex_works.ToString();
                CostData += ",";
                CostData += Lib.DatetoStringDisplayformat(DateTime.Now);

                if (CostData.Length > 150)
                    CostData = CostData.Substring(0, 150);

                Rec.InsertString("costd_cost_data", CostData);
                Rec.InsertNumeric("costd_rebate", Row.costd_rebate.ToString());
                Rec.InsertNumeric("costd_amnt_chrgs", Row.costd_amenment_chrgs.ToString());
                Rec.InsertString("costd_consgne_grp", Row.costd_consignee_group);
                Rec.InsertNumeric("costd_seal_chrgs", Row.costd_seal_chrgs.ToString());

                sql = Rec.UpdateRow();
                Con_Oracle.ExecuteNonQuery(sql);
                sql = "update hblm set hbl_terms ='" + Row.costd_hbl_terms + "' where hbl_pkid ='" + Row.costd_acc_id + "'";
                Con_Oracle.ExecuteNonQuery(sql);
            }

            if (sType != "GENERATE")//While saving
            {
                iCtr = 0;
                foreach (Costingd Row in Record.DetailList)
                {
                    iCtr++;
                    if (Row.costd_acc_name != ""|| Row.costd_remarks !="" || Lib.Conv2Decimal(Row.costd_acc_amt.ToString()) != 0)
                    {
                        Rec.CreateRow("Costingd", "ADD", "costd_pkid", Row.costd_pkid);
                        Rec.InsertString("costd_parent_id", Record.cost_pkid);
                        Rec.InsertString("costd_acc_id", Row.costd_acc_id);
                        Rec.InsertString("costd_acc_name", Row.costd_acc_name);
                        Rec.InsertString("costd_blno", Row.costd_blno);
                        Rec.InsertNumeric("costd_ctr", iCtr.ToString());
                        Rec.InsertNumeric("costd_brate", Row.costd_brate.ToString());
                        Rec.InsertNumeric("costd_srate", Row.costd_srate.ToString());
                        Rec.InsertNumeric("costd_split", Row.costd_split.ToString());
                        Rec.InsertNumeric("costd_acc_qty", Row.costd_acc_qty.ToString());
                        Rec.InsertNumeric("costd_acc_rate", Row.costd_acc_rate.ToString());
                        Rec.InsertNumeric("costd_acc_amt", Row.costd_acc_amt.ToString());
                        Rec.InsertString("costd_category", "INVOICE");
                        Rec.InsertString("costd_remarks", Row.costd_remarks);
                        sql = Rec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                }
            }
        }

        private void LoadTables(string CostingID)
        {
            AGENT_FORMAT = "";
            //sql = "select JINV_PKID from jobinvoice where jinv_source = 'MBLSEAEXP' ";
            //sql += " and jinv_parent_id='" + this.ROWUID + "' and JINV_CALC_ON='CBM' ";
            //DataTable dt_Temp = new DataTable();
            //dt_Temp = orConnection.RunSql(sql);
            //if (dt_Temp.Rows.Count > 0)
            //    Is_CalcOnCBM = true;

            sql = " select costd_frt_pp,costd_pp,costd_cc,costd_tot,costd_cost_data,";
            sql += "  costd_src_pp as acd_pp, ";
            sql += "  costd_src_cc as acd_cc, ";
            sql += "  costd_mcc_cc as ddc_cc, ";
            sql += "  costd_amnt_chrgs, ";
            sql += "  costd_seal_chrgs, ";
            sql += "  null as Frt_Status, ";
            sql += "  null as nom, ";
            sql += "  null as no_incentve, ";
            sql += "  null as spl_incent_rate, ";
            sql += "  null as handling_chrg, ";
            sql += "  'N' as House_not_Include, ";
            sql += "  null as House_ex_chrg ";
            sql += "  from costingd ";
            sql += "  where COSTD_PARENT_ID = '" + CostingID + "' and COSTD_TYPE='BUY' ";
            Con_Oracle = new DBConnection();

            DTP_COSTINGM = new DataTable();
            DTP_COSTINGM = Con_Oracle.ExecuteQuery(sql);
            string[] Sdata = null;
            foreach (DataRow dr in DTP_COSTINGM.Rows)
            {
                if (dr["costd_cost_data"].ToString() != "")
                {
                    Sdata = dr["costd_cost_data"].ToString().Split(',');
                    dr["nom"] = Sdata[0].Replace("MUTAL", "MUTUAL");
                    dr["frt_status"] = Sdata[1];
                    INCENTIVE_RATE = Lib.Convert2Decimal(Sdata[2]);
                    FH_RATE1 = Lib.Convert2Decimal(Sdata[3]);
                    FH_LIMIT1 = Lib.Convert2Decimal(Sdata[4]);
                    FH_RATE2 = Lib.Convert2Decimal(Sdata[5]);
                    FH_LIMIT2 = Lib.Convert2Decimal(Sdata[6]);
                    FH_RATE3 = Lib.Convert2Decimal(Sdata[7]);
                    Master_POFD = Sdata[8];
                    OTH_CHRG_RITRA = Lib.Convert2Decimal(Sdata[9]);

                    HAULAGE_PER_CBM = Sdata.Length > 15 ? Lib.Convert2Decimal(Sdata[15]) : 0;
                    HAULAGE_MIN_RATE = Sdata.Length > 16 ? Lib.Convert2Decimal(Sdata[16]) : 0;
                    WEIGHT_DIVIDER = Sdata.Length > 17 ? Lib.Convert2Decimal(Sdata[17]) : 0;
                    DESTUFF_P_D = Sdata.Length > 18 ? Lib.Convert2Decimal(Sdata[18]) : 0;
                    HANDLING_FEE = Sdata.Length > 19 ? Lib.Convert2Decimal(Sdata[19]) : 0;
                    TRUK_COST = Sdata.Length > 20 ? Lib.Convert2Decimal(Sdata[20]) : 0;
                    CNTR_SHIFT_CHRGS = Sdata.Length > 21 ? Lib.Convert2Decimal(Sdata[21]) : 0;
                    VESSEl_CHANGE_CHRGS = Sdata.Length > 22 ? Lib.Convert2Decimal(Sdata[22]) : 0;
                    EX_WORK_CHRGS = Sdata.Length > 23 ? Lib.Convert2Decimal(Sdata[23]) : 0;

                    //else if (CmbAgent.Text == "WELLTON")
                    //{
                    //    EX_WORK_CHRGS = Common.Convert2Decimal(Sdata[2]);
                    //    HAULAGE_PER_CBM = Common.Convert2Decimal(Sdata[3]);
                    //    HAULAGE_MIN_RATE = Common.Convert2Decimal(Sdata[4]);
                    //    WEIGHT_DIVIDER = Common.Convert2Decimal(Sdata[5]);
                    //    DESTUFF_P_D = Common.Convert2Decimal(Sdata[6]);
                    //    HANDLING_FEE = Common.Convert2Decimal(Sdata[7]);
                    //    TRUK_COST = Common.Convert2Decimal(Sdata[8]);
                    //    CNTR_SHIFT_CHRGS = Common.Convert2Decimal(Sdata[9]);
                    //    VESSEl_CHANGE_CHRGS = Common.Convert2Decimal(Sdata[10]);
                    //    Master_POFD = Sdata[12];
                    //}
                    //else if (CmbAgent.Text == "MOTHERLINES")
                    //{
                    //    Master_POFD = Sdata[3];
                    //    MOTHER_DEST_EXPENSE = Common.Convert2Decimal(Sdata[4]);

                    //}
                    //else if (CmbAgent.Text == "SEABRIDGE" || CmbAgent.Text == "ACTION" || CmbAgent.Text == "GATE4EU")
                    //{
                    //    HAULAGE_OTHERS = Common.Convert2Decimal(Sdata[2]);//Sdata[3] other destination chrg GBP
                    //    Master_POFD = Sdata[4];
                    //    OTHER_CHARGES = Common.Convert2Decimal(Sdata[5]);//other chrgs for HBL
                    //    EX_WORK_OTHERS = Common.Convert2Decimal(Sdata[6]);
                    //    SEAL_EXPENSE = Common.Convert2Decimal(Sdata[7]);
                    //}
                }
            }
            DTP_COSTINGM.AcceptChanges();// Dr_Target["CBM2"] = Common.Convert2Decimal(Dr["hbls_total_cbm"].ToString());
            // Sell Rates
            sql = " select costd_frt_pp,costd_frt_cc,costd_frt_rate_pp,costd_frt_rate_cc,costd_pp,costd_cc,costd_tot ,costd_rebate,";
            sql += "  costd_cbm as cbm,h.hbl_cbm as cbm2,costd_grwt,costd_cost_data,h.hbl_pkid,h.hbl_bl_no as blno, hbl_no,";
            sql += "  costd_mcc_pp as ddc_pp,costd_mcc_rate_pp as ddc_rate_pp, ";
            sql += "  costd_mcc_cc as ddc_cc,costd_mcc_rate_cc as ddc_rate_cc, ";
            sql += "  costd_src_pp as acd_pp, ";
            sql += "  costd_src_cc as acd_cc, ";
            sql += "  costd_oth_pp, ";
            sql += "  costd_oth_cc, ";
            sql += "  cons.cust_name as Consignee,costd_consgne_grp,";
            sql += "  pofd.param_name as pofd, ";
            sql += "  pol.param_name as pol, ";
            sql += "  vessel.param_name as vessel_name,m.hbl_vessel_no as hbl_vessel_no, ";
            sql += "  null as di_it,null as di_equ_surchrg, null as di_fsc, null as di_cfs_doc,";
            sql += "  null as di_stripping,null as di_dad,null as di_inbond,null as di_load_out,";
            sql += "  null as di_haz,null as di_do,";
            sql += "  null as Frt_Status, ";
            sql += "  null as nom, ";
            sql += "  null as no_incentve, ";
            sql += "  null as spl_incent_rate, ";
            sql += "  null as handling_chrg, ";
            sql += "  null as Haulage, ";
            sql += "  null as Oth_Des_GBP, ";
            sql += "  null as Ex_Work_Wellton, ";
            sql += "  'N' as House_not_Include, ";
            sql += "  null as House_ex_chrg ";
            sql += "  from costingd a ";
            sql += "  left join hblm h on (costd_acc_id = h.hbl_pkid )";
            sql += "  left join hblm m on (h.hbl_mbl_id = m.hbl_pkid )";
            sql += "  left join customerm cons on (h.hbl_imp_id = cons.cust_pkid) ";
            sql += "  left join param pofd on ( h.hbl_pofd_id = pofd.param_pkid) ";
            sql += "  left join param pol on ( h.hbl_pol_id = pol.param_pkid) ";
            sql += "  left join param vessel on ( m.hbl_vessel_id = vessel.param_pkid) ";
            sql += "  where COSTD_PARENT_ID = '" + CostingID + "' and COSTD_TYPE = 'SELL' order by h.hbl_bl_no ";
            DTP_COSTINGD = new DataTable();
            DTP_COSTINGD = Con_Oracle.ExecuteQuery(sql);
            DTP_COSTINGD.Columns.Add("LBS", typeof(System.Decimal));

            Sdata = null;
            decimal Lbs = 0;
            foreach (DataRow dr in DTP_COSTINGD.Rows)
            {
                if (dr["costd_cost_data"].ToString() != "")
                {
                    Sdata = dr["costd_cost_data"].ToString().Split(',');
                    dr["nom"] = Sdata[0].Replace("MUTAL", "MUTUAL");  
                    dr["frt_status"] = Sdata[1];
                    //ritra
                    dr["no_incentve"] = Sdata[10];
                    dr["handling_chrg"] = Lib.Convert2Decimal(Sdata[11]);
                    dr["spl_incent_rate"] = Lib.Convert2Decimal(Sdata[12]);
                    dr["House_not_Include"] = Sdata[13];
                    dr["House_ex_chrg"] = Lib.Convert2Decimal(Sdata[14]);
                    //Wellton
                    dr["Haulage"] = Sdata.Length > 15 ? Lib.Convert2Decimal(Sdata[15]) : 0;
                    dr["Ex_Work_Wellton"] = Sdata.Length > 23 ? Lib.Convert2Decimal(Sdata[23]) : 0;                     
                }
                Lbs = Lib.Convert2Decimal(dr["costd_grwt"].ToString()) * Lib.Convert2Decimal("2.2046");
                dr["LBS"] = Lib.Convert2Decimal(Lib.NumericFormat(Lbs.ToString(), 0));
            }
            DTP_COSTINGD.AcceptChanges();

            sql = "select a.*,agnt.cust_name as cost_agent_name from costingm a ";
            sql += " left join customerm agnt on a.cost_agent_id = agnt.cust_pkid";
            sql += " where cost_pkid='" + CostingID + "'";
            DTP_DESTN = new DataTable();
            DTP_DESTN = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            if(DTP_DESTN.Rows.Count>0)
            {
                if (DTP_DESTN.Rows[0]["cost_agent_name"].ToString().Contains("RITRA"))
                {
                    AGENT_FORMAT = "RITRA";
                    Process_RITRA();
                }
                else if (DTP_DESTN.Rows[0]["cost_agent_name"].ToString().Contains("TRAFFIC TECH"))
                {
                    AGENT_FORMAT = "TRAFFIC-TECH";
                    Process_WELLTON();
                }
            }
            
        }
        /******************RITRA************************/
        private void Process_RITRA()
        {
            MBL_FRT_STATUS = "";
            TOT_BUY_PREPAID = 0;
            TOT_BUY_COLLECT = 0;
            BUY_FRT_PP = 0;
            TOT_BUY = 0;
            TOT_NON_NOM_PRO_CHRGS = 0;
            TOT_NOM_PRO_CHRGS = 0;
            TOT_MUTUAL_PRO_CHRGS = 0;
            TOT_NETDUE = 0;
            TotHandlingChrgs = 0;
            decimal TempAmt = 0;
            TOT_NON_NOM_CBM_SPL_INCENTVE = 0;
            myDict = new Dictionary<string, decimal>();
            HT_SPL_INCENTIVE = new Dictionary<object, object>();
            EX_WORK_CHRGS = 0;
            try
            {
                int sky = 0;
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and nom='NOMINATION' and frt_Status <> 'FREIGHT PREPAID'"))
                {
                    AddToList(dr["COSTD_CONSGNE_GRP"].ToString(), Lib.Convert2Decimal(dr["cbm"].ToString()));
                }
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and nom='FREEHAND' and frt_Status = 'FREIGHT PREPAID'"))
                {
                    if (Lib.Convert2Decimal(dr["SPL_INCENT_RATE"].ToString()) > 0)
                    {
                        HT_SPL_INCENTIVE.Add(sky, dr["cbm"].ToString() + "," + dr["SPL_INCENT_RATE"].ToString());
                        sky++;
                    }
                }
                TOT_NON_NOM_CBM_NO_INCENTVE = Lib.Convert2Decimal(DTP_COSTINGD.Compute("sum(cbm)", "House_not_Include <> 'Y' and nom='FREEHAND' and no_incentve='Y'").ToString());
                TOT_NON_NOM_CBM_SPL_INCENTVE = Lib.Convert2Decimal(DTP_COSTINGD.Compute("sum(cbm)", "House_not_Include <> 'Y' and nom='FREEHAND' and spl_incent_rate > 0.00 ").ToString());
                TOT_NON_NOM_CBM = Lib.Convert2Decimal(DTP_COSTINGD.Compute("sum(cbm)", "House_not_Include <> 'Y' and nom='FREEHAND'").ToString());
                TOT_NOM_CBM = Lib.Convert2Decimal(DTP_COSTINGD.Compute("sum(cbm)", "House_not_Include <> 'Y' and nom='NOMINATION' and frt_Status <>'FREIGHT PREPAID'").ToString());
                TOT_MUTUAL_CBM = Lib.Convert2Decimal(DTP_COSTINGD.Compute("sum(cbm)", "House_not_Include <> 'Y' and (nom='MUTUAL' or (nom='NOMINATION' and frt_Status ='FREIGHT PREPAID'))").ToString());
                TOT_CBM = TOT_NON_NOM_CBM + TOT_NOM_CBM + TOT_MUTUAL_CBM;
            }
            catch (Exception)
            {
                throw;
            }

            HT_PRINT = new Dictionary<object, object>();
            foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y'", "BLNO"))
            {
                if (dr["NOM"].ToString() == "NOMINATION" && dr["FRT_STATUS"].ToString() != "FREIGHT PREPAID")
                {
                    if (Lib.Convert2Decimal(dr["HANDLING_CHRG"].ToString()) > 0) //if manually handling chrg given wll take other wise calculation per house
                    {
                        TotHandlingChrgs += Lib.Convert2Decimal(dr["HANDLING_CHRG"].ToString());
                    }
                    else
                    {
                        if (myDict[dr["COSTD_CONSGNE_GRP"].ToString()] < FH_LIMIT1)
                        {
                            TempAmt = Lib.Convert2Decimal(dr["CBM"].ToString()) * GetNominationRate(dr["COSTD_CONSGNE_GRP"].ToString());
                            TotHandlingChrgs += TempAmt;
                        }
                        else
                        {
                            if (!HT_PRINT.ContainsKey(dr["COSTD_CONSGNE_GRP"].ToString()))
                            {
                                HT_PRINT.Add(dr["COSTD_CONSGNE_GRP"].ToString(), "Y");
                                TempAmt = GetNominationRate(dr["COSTD_CONSGNE_GRP"].ToString());
                                TotHandlingChrgs += TempAmt;
                            }
                        }
                    }
                }
                EX_WORK_CHRGS += Lib.Convert2Decimal(dr["HOUSE_EX_CHRG"].ToString());
            }

            foreach (DataRow dr in DTP_COSTINGM.Rows)
            {
                MBL_FRT_STATUS = dr["frt_status"].ToString().Trim();
                BUY_FRT_PP = Lib.Convert2Decimal(dr["COSTD_FRT_PP"].ToString());
                TOT_BUY_COLLECT = Lib.Convert2Decimal(dr["COSTD_CC"].ToString());
                TOT_BUY_PREPAID = Lib.Convert2Decimal(dr["COSTD_PP"].ToString());
                
                if (MBL_FRT_STATUS.Trim() == "")
                    MBL_FRT_STATUS = "FREIGHT COLLECT";

                if (MBL_FRT_STATUS == "FREIGHT COLLECT")//modify from 30/08/2013
                    TOT_BUY = TOT_BUY_COLLECT;
                else if (MBL_FRT_STATUS == "FREIGHT PREPAID")
                    TOT_BUY = TOT_BUY_PREPAID;
                break;
            }

            if (TOT_CBM > 0)
                PER_CBM_RATE = TOT_BUY / TOT_CBM;
            else
                PER_CBM_RATE = TOT_BUY;

            FhandCC = Lib.Convert2Decimal(DTP_COSTINGD.Compute("sum(costd_frt_cc)", "House_not_Include <> 'Y' and nom='FREEHAND' and frt_Status ='FREIGHT COLLECT'").ToString());

            Tot_CHRGS_CC = 0;
            foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and nom='MUTUAL' and frt_Status <> 'FREIGHT PREPAID'", "blno"))
            {
                TOT_A = Lib.Convert2Decimal(dr["CBM"].ToString()) * Lib.Convert2Decimal(dr["COSTD_FRT_RATE_CC"].ToString());
                TOT_A = Lib.Convert2Decimal(Lib.NumericFormat(TOT_A.ToString(), 3));

                FRT_B = Lib.Convert2Decimal(dr["CBM"].ToString()) * PER_CBM_RATE;
                FRT_B = Lib.Convert2Decimal(Lib.NumericFormat(FRT_B.ToString(), 3));

                NET_C = (TOT_A - FRT_B) / 2;
                NET_C = Lib.Convert2Decimal(Lib.NumericFormat(NET_C.ToString(), 3));

                Tot_CHRGS_CC += NET_C;
            }

            IncentiveLocal = (TOT_NON_NOM_CBM - (TOT_NON_NOM_CBM_NO_INCENTVE + TOT_NON_NOM_CBM_SPL_INCENTVE)) * INCENTIVE_RATE;//Free Hand /Incentive
            string[] SplData = null;
            for (int k = 0; k < HT_SPL_INCENTIVE.Count; k++)
            {
                SplData = HT_SPL_INCENTIVE[k].ToString().Split(',');
                IncentiveLocal += Lib.Convert2Decimal(SplData[0]) * Lib.Convert2Decimal(SplData[1]);
            }

            TOT_NOM_PRO_CHRGS = TOT_NOM_CBM * PER_CBM_RATE;
            TOT_NOM_PRO_CHRGS = Lib.Convert2Decimal(Lib.NumericFormat(TOT_NOM_PRO_CHRGS.ToString(), 3));
            TOT_MUTUAL_PRO_CHRGS = TOT_MUTUAL_CBM * PER_CBM_RATE;
            TOT_MUTUAL_PRO_CHRGS = Lib.Convert2Decimal(Lib.NumericFormat(TOT_MUTUAL_PRO_CHRGS.ToString(), 3));

            // TOTAL_FROM_CHARGES = TotHandlingChrgs + FhandCC + IncentiveLocal + BUY_FRT_PP + Tot_CHRGS_CC + OTH_CHRG_RITRA + EX_WORK_CHRGS;
            TOTAL_FROM_CHARGES = TotHandlingChrgs + FhandCC + IncentiveLocal + Tot_CHRGS_CC + OTH_CHRG_RITRA + EX_WORK_CHRGS;
            if (TOT_BUY_PREPAID > 0 && MBL_FRT_STATUS == "FREIGHT PREPAID")
                TOTAL_FROM_CHARGES += TOT_NOM_PRO_CHRGS + TOT_MUTUAL_PRO_CHRGS;

            /*********************A-END***********************************************/
            /*********************B-START***********************************************/
            decimal fhandCBM = 0;
            foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and nom='FREEHAND' and ( frt_Status = 'FREIGHT PREPAID' or costd_frt_rate_cc > 0 )"))
            {
                fhandCBM += Lib.Convert2Decimal(dr["cbm"].ToString());
            }
            if (fhandCBM > 0)
            {
                FhandPP = fhandCBM * PER_CBM_RATE;
                FhandPP = Lib.Convert2Decimal(Lib.NumericFormat(FhandPP.ToString(), 3));
            }

            decimal MutualCBM = 0;
            foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and (nom='MUTUAL' or nom='NOMINATION') and frt_Status = 'FREIGHT PREPAID'"))
            {
                MutualCBM += Lib.Convert2Decimal(dr["cbm"].ToString());
            }
            if (MutualCBM > 0)
            {
                MutualPP = MutualCBM * PER_CBM_RATE;
                MutualPP = Lib.Convert2Decimal(Lib.NumericFormat(MutualPP.ToString(), 3));
            }

            Tot_CHRGS_PP = 0;
            foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and (nom='MUTUAL' or nom='NOMINATION') and frt_Status = 'FREIGHT PREPAID'", "blno"))
            {
                TOT_A2 = Lib.Convert2Decimal(dr["CBM"].ToString()) * Lib.Convert2Decimal(dr["COSTD_FRT_RATE_PP"].ToString());
                TOT_A2 = Lib.Convert2Decimal(Lib.NumericFormat(TOT_A2.ToString(), 3));

                FRT_B2 = Lib.Convert2Decimal(dr["CBM"].ToString()) * PER_CBM_RATE;
                FRT_B2 = Lib.Convert2Decimal(Lib.NumericFormat(FRT_B2.ToString(), 3));

                NET_C2 = (TOT_A2 - FRT_B2) / 2;
                NET_C2 = Lib.Convert2Decimal(Lib.NumericFormat(NET_C2.ToString(), 3));

                Tot_CHRGS_PP += NET_C2;
            }

            //TOTAL_TO_CHARGES = MutualPP + FhandPP + Tot_CHRGS_PP;
            TOTAL_TO_CHARGES = MutualPP + Tot_CHRGS_PP;
            if (TOT_BUY_COLLECT > 0 && MBL_FRT_STATUS == "FREIGHT COLLECT")
                TOTAL_TO_CHARGES += FhandPP;

            /*********************B-END***********************************************/
            // NET DUE
            TOT_NETDUE = TOTAL_FROM_CHARGES - TOTAL_TO_CHARGES;
            TOT_NETDUE = Lib.Convert2Decimal(Lib.NumericFormat(TOT_NETDUE.ToString(), 3));
        }
        private void SetupRitraColums()
        {
            ws.Columns[0].Width = 255 * 18;
            ws.Columns[1].Width = 255 * 9;
            ws.Columns[2].Width = 255 * 12;
            ws.Columns[3].Width = 255 * 13;
            ws.Columns[4].Width = 255 * 13;
            ws.Columns[5].Width = 255 * 13;
            ws.Columns[6].Width = 255 * 12;
            for (int s = 0; s < 7; s++)
            {
                ws.Columns[s].Style.Font.Name = "Arial";
                ws.Columns[s].Style.Font.Size = 8 * 20;
            }
            ws.Columns[1].Style.NumberFormat = "#0.000";
            ws.Columns[4].Style.NumberFormat = "#0.000";
            ws.Columns[5].Style.NumberFormat = "#0.000";
            ws.Columns[6].Style.NumberFormat = "#0.000";
        }
        private void Print_RITRA()
        {
            DataRow Drm = null;
            try
            {
                if (DTP_DESTN.Rows.Count <= 0)
                    return;
                Drm = DTP_DESTN.Rows[0];

                file = new ExcelFile();
                ws = file.Worksheets.Add("Report");
                ws.PrintOptions.Portrait = true;
                ws.PrintOptions.FitWorksheetWidthToPages = 1;
                SetupRitraColums();
                WriteData(0, 0, "STATEMENT NO.");
                WriteData(0, 1, Drm["COST_REFNO"]);//Txt_RefNo.Text
                WriteData(0, 5, "DATE  " + Lib.DatetoStringDisplayformat(Drm["COST_DATE"]));
                WriteData(2, 0, "VESSEL ");
                foreach (DataRow dr in DTP_COSTINGD.Rows)
                {
                    if (dr["House_not_Include"].ToString() != "Y")
                    {
                        WriteData(2, 1, dr["VESSEL_NAME"].ToString() + "  V- " + dr["HBL_VESSEL_NO"].ToString());
                        break;
                    }
                }
                WriteData(3, 0, "CONTAINER ");
                WriteData(3, 1, Drm["COST_CNTR"]);
                Merge_Cell(4, 0, 1, 2);
                WriteData(4, 0, "LEFT", "B/L NO");
                ws.Cells[4, 0].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);

                Merge_Cell(4, 1, 1, 2);
                WriteData(4, 1, "RIGHT", "CBM");
                ws.Cells[4, 1].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);

                Merge_Cell(4, 2, 1, 2);
                WriteData(4, 2, "LEFT", "RITRA/ CMAR");
                ws.Cells[4, 2].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);

                Merge_Cell(4, 3, 1, 2);
                WriteData(4, 3, "LEFT", "DESTINATION");
                ws.Cells[4, 3].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);

                Merge_Cell(4, 4, 2, 1);
                WriteData(4, 4, "FREIGHT COLLECT FOR CMAR");
                ws.Cells[4, 4].Style.Borders.SetBorders(MultipleBorders.Top, System.Drawing.Color.Black, LineStyle.Thin);
                ws.Cells[4, 4].Style.Font.Size = 8 * 20;

                Merge_Cell(5, 4, 1, 1);
                WriteData(5, 4, "RIGHT", "RATE USD");
                ws.Cells[5, 4].Style.Borders.SetBorders(MultipleBorders.Bottom, System.Drawing.Color.Black, LineStyle.Thin);

                Merge_Cell(5, 5, 1, 1);
                WriteData(5, 5, "RIGHT", "TOTAL USD");
                ws.Cells[5, 5].Style.Borders.SetBorders(MultipleBorders.Bottom, System.Drawing.Color.Black, LineStyle.Thin);

                Merge_Cell(4, 6, 1, 2);
                WriteData(4, 6, "RIGHT", "HANDLING USD");
                ws.Cells[4, 6].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);

                iRow = 7;
                decimal TempAmt = 0;
                decimal TotFrtCmar = 0;
                HT_PRINT = new Dictionary<object, object>();
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y'", "BLNO"))
                {
                    WriteData(iRow, 0, dr["BLNO"]);
                    WriteData(iRow, 1, dr["CBM"]);
                    if (dr["NOM"].ToString() == "FREEHAND")
                        WriteData(iRow, 2, "CMAR");
                    else if (dr["NOM"].ToString() == "NOMINATION" && dr["FRT_STATUS"].ToString() != "FREIGHT PREPAID")
                        WriteData(iRow, 2, "RITRA");
                    else
                        WriteData(iRow, 2, "RITRA/ CMAR");
                    WriteData(iRow, 3, dr["POFD"]);
                    if (dr["NOM"].ToString() == "FREEHAND" && Lib.Convert2Decimal(dr["COSTD_FRT_RATE_CC"].ToString()) > 0)
                    {
                        TempAmt = Lib.Convert2Decimal(dr["COSTD_FRT_RATE_CC"].ToString());
                        WriteData(iRow, 4, TempAmt);
                        TempAmt = TempAmt * Lib.Convert2Decimal(dr["CBM"].ToString());
                        TempAmt = Lib.Convert2Decimal(Lib.NumericFormat(TempAmt.ToString(), 3));
                        WriteData(iRow, 5, TempAmt);
                        TotFrtCmar += TempAmt;
                    }
                    if (dr["NOM"].ToString() == "NOMINATION" && dr["FRT_STATUS"].ToString() != "FREIGHT PREPAID")
                    {
                        if (Lib.Convert2Decimal(dr["HANDLING_CHRG"].ToString()) > 0) //if manually handling chrg given wll take other wise calculation per house
                        {
                            WriteData(iRow, 6, Lib.Convert2Decimal(dr["HANDLING_CHRG"].ToString()));
                        }
                        else
                        {
                            if (myDict[dr["COSTD_CONSGNE_GRP"].ToString()] < FH_LIMIT1)
                            {
                                TempAmt = Lib.Convert2Decimal(dr["CBM"].ToString()) * GetNominationRate(dr["COSTD_CONSGNE_GRP"].ToString());
                                WriteData(iRow, 6, TempAmt);
                            }
                            else
                            {
                                if (!HT_PRINT.ContainsKey(dr["COSTD_CONSGNE_GRP"].ToString()))
                                {
                                    HT_PRINT.Add(dr["COSTD_CONSGNE_GRP"].ToString(), "Y");
                                    TempAmt = GetNominationRate(dr["COSTD_CONSGNE_GRP"].ToString());
                                    WriteData(iRow, 6, TempAmt);
                                }
                            }
                        }
                    }
                    iRow++;
                }

                for (int c = 0; c < 7; c++)
                    ws.Cells[iRow, c].Style.Borders.SetBorders(MultipleBorders.Bottom, System.Drawing.Color.Black, LineStyle.Thin);
                iRow++;
                WriteData(iRow, 1, TOT_CBM);
                if (TotFrtCmar > 0)
                    WriteData(iRow, 5, TotFrtCmar);
                if (TotHandlingChrgs > 0)
                    WriteData(iRow, 6, TotHandlingChrgs);
                for (int c = 0; c < 7; c++)
                    ws.Cells[iRow, c].Style.Borders.SetBorders(MultipleBorders.Bottom, System.Drawing.Color.Black, LineStyle.Thin);
                iRow++;

                iRow++;
                WriteData(iRow, 4, "RIGHT", "CBM");
                WriteData(iRow++, 5, "RIGHT", "USD");
                WriteData(iRow, 0, "CARGOMAR GENERATED CARGO");

                if (TOT_NON_NOM_CBM > 0)
                {
                    WriteData(iRow, 4, TOT_NON_NOM_CBM);
                    TOT_NON_NOM_PRO_CHRGS = TOT_NON_NOM_CBM * PER_CBM_RATE;
                    TOT_NON_NOM_PRO_CHRGS = Lib.Convert2Decimal(Lib.NumericFormat(TOT_NON_NOM_PRO_CHRGS.ToString(), 3));
                    WriteData(iRow, 5, TOT_NON_NOM_PRO_CHRGS);
                }
                iRow++;
                WriteData(iRow, 0, "RITRA GENERATED CARGO");
                if (TOT_NOM_CBM > 0)
                {
                    WriteData(iRow, 4, TOT_NOM_CBM);
                    TOT_NOM_PRO_CHRGS = TOT_NOM_CBM * PER_CBM_RATE;
                    TOT_NOM_PRO_CHRGS = Lib.Convert2Decimal(Lib.NumericFormat(TOT_NOM_PRO_CHRGS.ToString(), 3));
                    WriteData(iRow, 5, TOT_NOM_PRO_CHRGS);
                }
                iRow++;
                WriteData(iRow, 0, "CARGOMAR/RITRA GENERATED CARGO");
                if (TOT_MUTUAL_CBM > 0)
                {
                    WriteData(iRow, 4, TOT_MUTUAL_CBM);
                    TOT_MUTUAL_PRO_CHRGS = TOT_MUTUAL_CBM * PER_CBM_RATE;
                    TOT_MUTUAL_PRO_CHRGS = Lib.Convert2Decimal(Lib.NumericFormat(TOT_MUTUAL_PRO_CHRGS.ToString(), 3));
                    WriteData(iRow, 5, TOT_MUTUAL_PRO_CHRGS);
                }
                iRow++;
                if (TOT_CBM > 0)
                {
                    WriteData(iRow, 4, TOT_CBM);
                    ws.Cells[iRow, 4].Style.Borders.SetBorders(MultipleBorders.Top, System.Drawing.Color.Black, LineStyle.Thin);
                }
                if (TOT_BUY > 0)
                {
                    WriteData(iRow, 5, TOT_BUY);
                    if (TOT_BUY_COLLECT > 0 && MBL_FRT_STATUS == "FREIGHT COLLECT")
                        WriteData(iRow, 6, "(COLLECT)");
                    else if (TOT_BUY_PREPAID > 0 && MBL_FRT_STATUS == "FREIGHT PREPAID")
                        WriteData(iRow, 6, "(PREPAID)");
                    ws.Cells[iRow, 5].Style.Borders.SetBorders(MultipleBorders.Top, System.Drawing.Color.Black, LineStyle.Thin);
                }
                iRow++;
                iRow++;
                /*********************A-START***********************************************/
                WriteData(iRow++, 0, "A. DUE FROM " + Drm["COST_AGENT_NAME"].ToString());
                WriteData(iRow, 0, "OCEAN FREIGHT PAYABLE");
                if (TOT_NOM_PRO_CHRGS > 0 && TOT_BUY_PREPAID > 0 && MBL_FRT_STATUS == "FREIGHT PREPAID")
                    WriteData(iRow, 5, TOT_NOM_PRO_CHRGS);
                iRow++;
                WriteData(iRow, 0, "OCEAN FREIGHT PAYABLE FOR CMAR/RITRA GENERATED CARGO");
                if (TOT_MUTUAL_PRO_CHRGS > 0 && TOT_BUY_PREPAID > 0 && MBL_FRT_STATUS == "FREIGHT PREPAID")
                    WriteData(iRow, 5, TOT_MUTUAL_PRO_CHRGS);
                iRow++;
                WriteData(iRow, 0, "OCEAN FREIGHT COLLECT AT DESTINATION");
                if (FhandCC > 0)
                    WriteData(iRow, 5, FhandCC);
                iRow++;
                if (OTH_CHRG_RITRA > 0)
                {
                    WriteData(iRow, 0, "OTHER CHARGES PAID BY CARGOMAR");
                    WriteData(iRow, 5, OTH_CHRG_RITRA);
                    iRow++;
                }
                WriteData(iRow++, 0, "PROFIT SHARE FOR CMAR/RITRA GENERATED CARGO (A-B)/2");
                WriteData(iRow, 0, "LEFT", "HBL#");
                WriteData(iRow, 1, "RIGHT", "CBM");
                WriteData(iRow, 2, "RIGHT", "RATE");
                WriteData(iRow, 3, "RIGHT", "TOTAL(A)");
                WriteData(iRow++, 4, "RIGHT", "FRT.PAYABLE(B)");

                Tot_CHRGS_CC = 0;
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and nom='MUTUAL' and frt_Status <> 'FREIGHT PREPAID'", "blno"))
                {
                    TOT_A = Lib.Convert2Decimal(dr["CBM"].ToString()) * Lib.Convert2Decimal(dr["COSTD_FRT_RATE_CC"].ToString());
                    TOT_A = Lib.Convert2Decimal(Lib.NumericFormat(TOT_A.ToString(), 3));
                    FRT_B = Lib.Convert2Decimal(dr["CBM"].ToString()) * PER_CBM_RATE;
                    FRT_B = Lib.Convert2Decimal(Lib.NumericFormat(FRT_B.ToString(), 3));
                    NET_C = (TOT_A - FRT_B) / 2;
                    NET_C = Lib.Convert2Decimal(Lib.NumericFormat(NET_C.ToString(), 3));
                    Tot_CHRGS_CC += NET_C;

                    WriteData(iRow, 0, dr["BLNO"]);
                    WriteData(iRow, 1, dr["CBM"]);
                    WriteData(iRow, 2, dr["COSTD_FRT_RATE_CC"]);
                    WriteData(iRow, 3, TOT_A);
                    WriteData(iRow, 4, FRT_B);
                    WriteData(iRow, 5, NET_C);
                    iRow++;
                }
                iRow++;
                WriteData(iRow, 0, "HANDLING CHARGES");
                if (TotHandlingChrgs > 0)
                    WriteData(iRow, 5, TotHandlingChrgs);
                iRow++;
                string str = "*INCENTIVE FOR LOCAL GENERATED CARGO @ USD " + INCENTIVE_RATE.ToString() + "/CBM x ";
                if (TOT_NON_NOM_CBM - (TOT_NON_NOM_CBM_NO_INCENTVE + TOT_NON_NOM_CBM_SPL_INCENTVE) > 0)
                    str += Lib.NumericFormat((TOT_NON_NOM_CBM - (TOT_NON_NOM_CBM_NO_INCENTVE + TOT_NON_NOM_CBM_SPL_INCENTVE)).ToString(), 3);
                WriteData(iRow, 0, str);
                if (IncentiveLocal > 0 & (TOT_NON_NOM_CBM - (TOT_NON_NOM_CBM_NO_INCENTVE + TOT_NON_NOM_CBM_SPL_INCENTVE) > 0))//Free Hand /Incentive
                    WriteData(iRow, 5, IncentiveLocal);
                string[] SplData = null;
                for (int k = 0; k < HT_SPL_INCENTIVE.Count; k++)
                {
                    SplData = HT_SPL_INCENTIVE[k].ToString().Split(',');
                    iRow++;
                    str = "*INCENTIVE FOR LOCAL GENERATED CARGO @ USD " + SplData[1] + "/CBM x ";
                    str += SplData[0];
                    WriteData(iRow, 0, str);
                    WriteData(iRow, 5, Lib.Convert2Decimal(SplData[0]) * Lib.Convert2Decimal(SplData[1]));
                }
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and house_ex_chrg > 0.000", "blno"))
                {
                    iRow++;
                    str = "EX-WORK CHARGES AGAINST HBL# " + dr["blno"].ToString();
                    WriteData(iRow, 0, str);
                    WriteData(iRow, 5, Lib.Convert2Decimal(dr["house_ex_chrg"].ToString()));
                }
                iRow++;
                iRow++;
                if (TOTAL_FROM_CHARGES > 0)
                {
                    WriteData(iRow, 5, TOTAL_FROM_CHARGES);
                    ws.Cells[iRow, 5].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);
                }
                iRow++;
                /*********************A-END***********************************************/
                /*********************B-START***********************************************/
                WriteData(iRow++, 0, "B. DUE TO " + Drm["COST_AGENT_NAME"].ToString());
                WriteData(iRow, 0, "OCEAN FREIGHT PAYABLE");
                if (FhandPP > 0 && TOT_BUY_COLLECT > 0 && MBL_FRT_STATUS == "FREIGHT COLLECT")
                    WriteData(iRow, 5, FhandPP);
                iRow++;
                WriteData(iRow, 0, "OCEAN FREIGHT PAYABLE FOR CMAR/RITRA GENERATED CARGO");
                if (MutualPP > 0)
                    WriteData(iRow, 5, MutualPP);
                iRow++;
                WriteData(iRow++, 0, "PROFIT SHARE FOR CMAR/RITRA GENERATED CARGO (A-B)/2");
                WriteData(iRow, 0, "LEFT", "HBL#");
                WriteData(iRow, 1, "RIGHT", "CBM");
                WriteData(iRow, 2, "RIGHT", "RATE");
                WriteData(iRow, 3, "RIGHT", "TOTAL(A)");
                WriteData(iRow++, 4, "RIGHT", "FRT.PAYABLE(B)");

                Tot_CHRGS_PP = 0;
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and (nom='MUTUAL' or nom='NOMINATION') and frt_Status = 'FREIGHT PREPAID'", "blno"))
                {
                    TOT_A2 = Lib.Convert2Decimal(dr["CBM"].ToString()) * Lib.Convert2Decimal(dr["COSTD_FRT_RATE_PP"].ToString());
                    TOT_A2 = Lib.Convert2Decimal(Lib.NumericFormat(TOT_A2.ToString(), 3));
                    FRT_B2 = Lib.Convert2Decimal(dr["CBM"].ToString()) * PER_CBM_RATE;
                    FRT_B2 = Lib.Convert2Decimal(Lib.NumericFormat(FRT_B2.ToString(), 3));
                    NET_C2 = (TOT_A2 - FRT_B2) / 2;
                    NET_C2 = Lib.Convert2Decimal(Lib.NumericFormat(NET_C2.ToString(), 3));
                    Tot_CHRGS_PP += NET_C2;

                    WriteData(iRow, 0, dr["BLNO"]);
                    WriteData(iRow, 1, dr["CBM"]);
                    WriteData(iRow, 2, dr["COSTD_FRT_RATE_PP"]);
                    WriteData(iRow, 3, TOT_A2);
                    WriteData(iRow, 4, FRT_B2);
                    WriteData(iRow, 5, NET_C2);
                    iRow++;
                }
                iRow++;
                if (TOTAL_TO_CHARGES > 0)
                {
                    WriteData(iRow, 5, TOTAL_TO_CHARGES);
                    ws.Cells[iRow, 5].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);
                }
                /*********************B-END***********************************************/
                iRow++;
                iRow++;
                if (TOT_NETDUE >= 0)
                {
                    WriteData(iRow, 0, "NET DUE FROM " + Drm["COST_AGENT_NAME"].ToString());
                    WriteData(iRow, 5, TOT_NETDUE);
                }
                else
                {
                    WriteData(iRow, 0, "NET DUE TO " + Drm["COST_AGENT_NAME"].ToString());
                    WriteData(iRow, 5, Math.Abs(TOT_NETDUE));
                }
                for (int c = 0; c < 7; c++)
                    ws.Cells[iRow, c].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);

                ws.Columns[0].InsertEmpty(1);
                ws.Columns[0].Width = 255 * 6;
                file.SaveXls(File_Name + ".xls");
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        private void AddToList(string sKey, decimal nTot)
        {
            if (myDict.ContainsKey(sKey))
            {
                nTot = nTot + myDict[sKey];
                myDict[sKey] = nTot;
            }
            else
            {
                myDict.Add(sKey, nTot);
            }
        }
        private decimal GetNominationRate(string Consignee)
        {
            decimal TotCbm = 0;
            decimal CbmRate = 0;
            if (myDict.ContainsKey(Consignee))
                TotCbm = myDict[Consignee];
            if (TotCbm < FH_LIMIT1)
                CbmRate = FH_RATE1;
            else if (TotCbm < FH_LIMIT2)
                CbmRate = FH_RATE2;
            else
                CbmRate = FH_RATE3;
            return CbmRate;
        }
        private void Fill_RITRAInvoice()
        {
            DataRow Drm = null;
            try
            {
                if (DTP_DESTN.Rows.Count <= 0)
                    return;
                Drm = DTP_DESTN.Rows[0];

                iRow = 7;
                decimal TempAmt = 0;
                decimal TotFrtCmar = 0;
                HT_PRINT = new Dictionary<object, object>();
                Boolean badd = false;
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y'", "BLNO"))
                {

                    if (dr["NOM"].ToString() == "FREEHAND")
                        badd = false;
                    else if (dr["NOM"].ToString() == "NOMINATION" && dr["FRT_STATUS"].ToString() != "FREIGHT PREPAID")
                        badd = true;
                    else
                        badd = false;


                    if (badd)
                    {
                        InvRow = new Costingd();
                        InvRow.costd_pkid = Guid.NewGuid().ToString().ToUpper();
                        InvRow.costd_parent_id = Drm["cost_pkid"].ToString();
                        InvRow.costd_category = "INVOICE";
                        InvRow.costd_blno = dr["BLNO"].ToString();
                        InvRow.costd_acc_name = "OUR HANDLING CHARGES";
                        InvRow.costd_remarks = "USD 5/CBM";
                        InvRow.costd_srate = 0;
                        InvRow.costd_brate = 0;
                        InvRow.costd_split = 0;
                        InvRow.costd_acc_qty = Lib.Conv2Decimal(Lib.NumericFormat(dr["CBM"].ToString(), 3));
                        InvRow.costd_acc_rate = Lib.Conv2Decimal(Lib.NumericFormat("5", 2));
  
                        if (dr["NOM"].ToString() == "NOMINATION" && dr["FRT_STATUS"].ToString() != "FREIGHT PREPAID")
                        {
                            if (Lib.Convert2Decimal(dr["HANDLING_CHRG"].ToString()) > 0) //if manually handling chrg given wll take other wise calculation per house
                            {
                                InvRow.costd_acc_amt = Lib.Conv2Decimal(Lib.NumericFormat(dr["HANDLING_CHRG"].ToString(), 2));
                            }
                            else
                            {
                                if (myDict[dr["COSTD_CONSGNE_GRP"].ToString()] < FH_LIMIT1)
                                {
                                    TempAmt = Lib.Convert2Decimal(dr["CBM"].ToString()) * GetNominationRate(dr["COSTD_CONSGNE_GRP"].ToString());
                                    InvRow.costd_acc_amt = Lib.Conv2Decimal(Lib.NumericFormat(TempAmt.ToString(), 2));
                                }
                                else
                                {
                                    if (!HT_PRINT.ContainsKey(dr["COSTD_CONSGNE_GRP"].ToString()))
                                    {
                                        HT_PRINT.Add(dr["COSTD_CONSGNE_GRP"].ToString(), "Y");
                                        TempAmt = GetNominationRate(dr["COSTD_CONSGNE_GRP"].ToString());
                                        InvRow.costd_acc_amt = Lib.Conv2Decimal(Lib.NumericFormat(TempAmt.ToString(), 2));
                                    }
                                }
                            }
                        }
                        InvList.Add(InvRow);
                    }
                }

                Tot_CHRGS_CC = 0;
                foreach (DataRow dr in DTP_COSTINGD.Select("House_not_Include <> 'Y' and nom='MUTUAL' and frt_Status <> 'FREIGHT PREPAID'", "blno"))
                {
                    TOT_A = Lib.Convert2Decimal(dr["CBM"].ToString()) * Lib.Convert2Decimal(dr["COSTD_FRT_RATE_CC"].ToString());
                    TOT_A = Lib.Convert2Decimal(Lib.NumericFormat(TOT_A.ToString(), 3));
                    FRT_B = Lib.Convert2Decimal(dr["CBM"].ToString()) * PER_CBM_RATE;
                    FRT_B = Lib.Convert2Decimal(Lib.NumericFormat(FRT_B.ToString(), 3));
                    NET_C = (TOT_A - FRT_B) / 2;
                    NET_C = Lib.Convert2Decimal(Lib.NumericFormat(NET_C.ToString(), 3));
                    Tot_CHRGS_CC += NET_C;

                    InvRow = new Costingd();
                    InvRow.costd_pkid = Guid.NewGuid().ToString().ToUpper();
                    InvRow.costd_parent_id = Drm["cost_pkid"].ToString();
                    InvRow.costd_category = "INVOICE";
                    InvRow.costd_blno = dr["BLNO"].ToString();
                    InvRow.costd_acc_name = "OUR HANDLING CHARGES";
                    InvRow.costd_remarks = "[S/R:" + TOT_A.ToString() + " B/R:" + FRT_B.ToString() + "]/2";
                    InvRow.costd_srate= Lib.Conv2Decimal(Lib.NumericFormat(TOT_A.ToString(), 3));
                    InvRow.costd_brate = Lib.Conv2Decimal(Lib.NumericFormat(FRT_B.ToString(), 3));
                    InvRow.costd_split = 2;
                    InvRow.costd_acc_qty = 1;
                    InvRow.costd_acc_rate = Lib.Conv2Decimal(Lib.NumericFormat(NET_C.ToString(), 2));
                    InvRow.costd_acc_amt = Lib.Conv2Decimal(Lib.NumericFormat(NET_C.ToString(), 2));
                    InvList.Add(InvRow);
                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        /******************RITRA************************/
        /******************WELLTON************************/
        private void SetupWelltonColums()
        {
            ws.Columns[0].Width = 255 * 2;
            ws.Columns[1].Width = 255 * 9;
            ws.Columns[2].Width = 255 * 8;
            ws.Columns[3].Width = 255 * 10;
            ws.Columns[4].Width = 250 * 8;
            ws.Columns[5].Width = 250 * 8;
            ws.Columns[6].Width = 250 * 8;
            ws.Columns[7].Width = 250 * 8;
            ws.Columns[8].Width = 255 * 9;
            ws.Columns[9].Width = 255 * 9;
            ws.Columns[10].Width = 255 * 9;
            ws.Columns[11].Width = 275 * 7;
            ws.Columns[12].Width = 250 * 7;
            for (int s = 0; s < 13; s++)
            {
                ws.Columns[s].Style.Font.Name = "Arial";
                ws.Columns[s].Style.Font.Size = 8 * 20;
                ws.Columns[s].Style.NumberFormat = "#0.000";
            }
        }
        private void Process_WELLTON()
        {
            try
            {
                Is_haulage_Minimum = false;
                TOT_CBM = 0;
                TOT_CC = 0; TOT_PP = 0;
                TOT_HAULAGE = 0; TOT_REBATE = 0; PROFIT = 0; EX_WORK_CHRGS = 0;
                foreach (DataRow dr in DTP_COSTINGD.Select("1=1", "BLNO"))
                {
                    if (dr["POFD"].ToString() == "TORONTO" && Master_POFD != "TORONTO")
                    {
                        HAULAGE_PER_CBM = Lib.Convert2Decimal(dr["Haulage"].ToString());
                        TOT_HAULAGE += GetHaulageRate(dr["CBM2"].ToString(), dr["COSTD_GRWT"].ToString());
                    }
                    TOT_CBM += Lib.Convert2Decimal(Lib.NumericFormat(dr["CBM"].ToString(), 3));
                    TOT_PP += Lib.Convert2Decimal(Lib.NumericFormat(dr["COSTD_PP"].ToString(), 3));
                    TOT_CC += Lib.Convert2Decimal(Lib.NumericFormat(dr["COSTD_CC"].ToString(), 3));
                    TOT_REBATE += Lib.Convert2Decimal(Lib.NumericFormat(dr["COSTD_REBATE"].ToString(), 3));
                    EX_WORK_CHRGS += Lib.Convert2Decimal(Lib.NumericFormat(dr["EX_WORK_WELLTON"].ToString(), 3));
                }

                TOT_A = Lib.Convert2Decimal(Lib.NumericFormat((TOT_CC + TOT_PP).ToString(), 3));

                if (TOT_HAULAGE > 0)
                    if (HAULAGE_MIN_RATE > TOT_HAULAGE)
                    {
                        Is_haulage_Minimum = true;
                        TOT_HAULAGE = HAULAGE_MIN_RATE;
                    }
                TOT_HAULAGE = Lib.Convert2Decimal(Lib.NumericFormat(TOT_HAULAGE.ToString(), 3));

                TOT_BUY_COLLECT = 0;
                TOT_BUY_PREPAID = 0;
                BUY_ACD_CC = 0;
                BUY_ACD_PP = 0;
                BUY_DDC_CC = 0;
                BUY_AMENMENT_CHRGS = 0;
                foreach (DataRow dr in DTP_COSTINGM.Rows)
                {
                    TOT_BUY_COLLECT = Lib.Convert2Decimal(dr["COSTD_CC"].ToString());
                    TOT_BUY_PREPAID = Lib.Convert2Decimal(dr["COSTD_PP"].ToString());
                    BUY_ACD_CC = Lib.Convert2Decimal(dr["ACD_CC"].ToString());
                    BUY_ACD_PP = Lib.Convert2Decimal(dr["ACD_PP"].ToString());
                    BUY_DDC_CC = Lib.Convert2Decimal(dr["DDC_CC"].ToString());
                    BUY_AMENMENT_CHRGS = Lib.Convert2Decimal(dr["COSTD_AMNT_CHRGS"].ToString());
                }
                TOT_BUY_COLLECT = TOT_BUY_COLLECT - BUY_ACD_CC;//tot collect chrgs contains ACD Collect chrgs so substracting ACD from totCollect;
                TOT_BUY_PREPAID = TOT_BUY_PREPAID - BUY_ACD_PP;// -do-
                TOT_BUY_COLLECT = TOT_BUY_COLLECT - BUY_DDC_CC;

                TOT_B = TOT_BUY_COLLECT + BUY_DDC_CC + DESTUFF_P_D + HANDLING_FEE + BUY_ACD_CC + TOT_HAULAGE;
                TOT_B = Lib.Convert2Decimal(Lib.NumericFormat(TOT_B.ToString(), 3));

                TOT_C = TOT_BUY_PREPAID + TRUK_COST + CNTR_SHIFT_CHRGS + VESSEl_CHANGE_CHRGS;
                TOT_C += EX_WORK_CHRGS + TOT_REBATE + BUY_ACD_PP + BUY_AMENMENT_CHRGS;
                TOT_C = Lib.Convert2Decimal(Lib.NumericFormat(TOT_C.ToString(), 3));

                TOT_A2 = TOT_A - (TOT_B + TOT_C);
                PROFIT = TOT_A2 / 2;
                PROFIT = Lib.Convert2Decimal(Lib.NumericFormat(PROFIT.ToString(), 3));
                TOT_NETDUE = (TOT_B + PROFIT - TOT_CC) * -1;
                TOT_NETDUE = Lib.Convert2Decimal(Lib.NumericFormat(TOT_NETDUE.ToString(), 3));

                //this to show expense/income on screen while process
                TOTAL_FROM_CHARGES = TOT_A ;
                TOTAL_TO_CHARGES = TOT_B ;
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        private void Print_WELLTON()
        {
            DataRow Drm = null;
            try
            {
                if (DTP_DESTN.Rows.Count <= 0)
                    return;
                Drm = DTP_DESTN.Rows[0];

                file = new ExcelFile();
                ws = file.Worksheets.Add("Report");
                ws.PrintOptions.Portrait = true;
                ws.PrintOptions.FitWorksheetWidthToPages = 1;
                SetupWelltonColums();
                WriteData(1, 8, "DATE  " + Lib.DatetoStringDisplayformat(Drm["COST_DATE"]));
                WriteData(1, 1, "VESSEL ");
                foreach (DataRow dr in DTP_COSTINGD.Rows)
                {
                    WriteData(1, 2, dr["VESSEL_NAME"].ToString() + "  V- " + dr["HBL_VESSEL_NO"].ToString());
                    break;
                }
                WriteData(2, 1, "CONTAINER ");
                WriteData(2, 2, Drm["COST_CNTR"]);
                WriteData(2, 8, "STATEMENT NO. " + Drm["COST_REFNO"]);

                iRow = 3;
                WriteData(iRow, 1, "BLNO");
                Merge_Cell(iRow, 2, 2, 1);
                WriteData(iRow, 2, "LEFT", "MEASUREMENT");

                WriteData(iRow + 1, 2, "RIGHT", "CBM");
                WriteData(iRow + 1, 3, "LEFT", "DESTN.");
                WriteData(iRow, 4, "RIGHT", "FRT/ CBM");
                WriteData(iRow + 1, 4, "RIGHT", "USD");

                WriteData(iRow, 5, "RIGHT", "DDC/ CBM");
                WriteData(iRow + 1, 5, "RIGHT", "USD");
                WriteData(iRow, 6, "RIGHT", "ACD");
                WriteData(iRow + 1, 6, "RIGHT", "USD");
                WriteData(iRow, 7, "RIGHT", "OTH/ LS");
                WriteData(iRow + 1, 7, "RIGHT", "USD");
                Merge_Cell(iRow, 8, 2, 1);
                WriteData(iRow, 8, "FREIGHT");
                WriteData(iRow + 1, 8, "RIGHT", "PREPAID");
                WriteData(iRow + 1, 9, "RIGHT", "FOB");

                WriteData(iRow, 10, "RIGHT", "TOTAL");
                WriteData(iRow + 1, 10, "RIGHT", "USD");
                WriteData(iRow, 11, "RIGHT", "HAULAGE");
                WriteData(iRow + 1, 11, "RIGHT", "USD");
                WriteData(iRow, 12, "RIGHT", "REBATE");
                WriteData(iRow + 1, 12, "RIGHT", "USD");
                for (int c = 1; c < 13; c++)
                {
                    ws.Cells[iRow, c].Style.Borders.SetBorders(MultipleBorders.Top, System.Drawing.Color.Black, LineStyle.Thin);
                    ws.Cells[iRow + 1, c].Style.Borders.SetBorders(MultipleBorders.Bottom, System.Drawing.Color.Black, LineStyle.Thin);
                }
                iRow += 2;
                decimal TempAmt = 0;
                bool printMinValue = false;
                foreach (DataRow dr in DTP_COSTINGD.Select("1=1", "BLNO"))
                {
                    WriteData(iRow, 1, "LEFT", dr["BLNO"].ToString());
                    WriteData(iRow, 2, "RIGHT", dr["CBM"]);
                    WriteData(iRow, 3, "LEFT", dr["POFD"].ToString());
                    if (Lib.Convert2Decimal(dr["COSTD_FRT_RATE_CC"].ToString()) > 0)
                        WriteData(iRow, 4, "RIGHT", GetCbmRate(dr["COSTD_FRT_RATE_CC"].ToString(), " C"));
                    else if (Lib.Convert2Decimal(dr["COSTD_FRT_RATE_PP"].ToString()) > 0)
                        WriteData(iRow, 4, "RIGHT", GetCbmRate(dr["COSTD_FRT_RATE_PP"].ToString(), " P"));

                    if (Lib.Convert2Decimal(dr["DDC_RATE_CC"].ToString()) > 0)
                        WriteData(iRow, 5, "RIGHT", GetCbmRate(dr["DDC_RATE_CC"].ToString(), " C"));
                    else if (Lib.Convert2Decimal(dr["DDC_RATE_PP"].ToString()) > 0)
                        WriteData(iRow, 5, "RIGHT", GetCbmRate(dr["DDC_RATE_PP"].ToString(), " P"));

                    if (Lib.Convert2Decimal(dr["ACD_CC"].ToString()) > 0)
                        WriteData(iRow, 6, "RIGHT", Lib.NumericFormat(dr["ACD_CC"].ToString(), 2) + " C");
                    else if (Lib.Convert2Decimal(dr["ACD_PP"].ToString()) > 0)
                        WriteData(iRow, 6, "RIGHT", Lib.NumericFormat(dr["ACD_PP"].ToString(), 2) + " P");

                    if (Lib.Convert2Decimal(dr["COSTD_OTH_CC"].ToString()) > 0)
                        WriteData(iRow, 7, "RIGHT", Lib.NumericFormat(dr["COSTD_OTH_CC"].ToString(), 2) + " C");
                    else if (Lib.Convert2Decimal(dr["COSTD_OTH_PP"].ToString()) > 0)
                        WriteData(iRow, 7, "RIGHT", Lib.NumericFormat(dr["COSTD_OTH_PP"].ToString(), 2) + " P");

                    if (Lib.Convert2Decimal(dr["COSTD_PP"].ToString()) > 0)
                        WriteData(iRow, 8, "RIGHT", dr["COSTD_PP"]);
                    if (Lib.Convert2Decimal(dr["COSTD_CC"].ToString()) > 0)
                        WriteData(iRow, 9, "RIGHT", dr["COSTD_CC"]);

                    TempAmt = Lib.Convert2Decimal(dr["COSTD_PP"].ToString()) + Lib.Convert2Decimal(dr["COSTD_CC"].ToString());
                    if (TempAmt > 0)
                        WriteData(iRow, 10, "RIGHT", TempAmt);

                    if (dr["POFD"].ToString() == "TORONTO" && Master_POFD != "TORONTO")
                    {
                        HAULAGE_PER_CBM = Lib.Convert2Decimal(dr["Haulage"].ToString());
                        if (Is_haulage_Minimum && !printMinValue)
                        {
                            printMinValue = true;
                            WriteData(iRow, 11, "RIGHT", TOT_HAULAGE);
                        }
                        else if (!Is_haulage_Minimum)
                            WriteData(iRow, 11, "RIGHT", GetHaulageRate(dr["CBM2"].ToString(), dr["COSTD_GRWT"].ToString()));
                    }

                    if (Lib.Convert2Decimal(dr["COSTD_REBATE"].ToString()) > 0)
                        WriteData(iRow, 12, "RIGHT", dr["COSTD_REBATE"]);
                    iRow++;
                }
                iRow++;
                if (TOT_CBM > 0)
                    WriteData(iRow, 2, "RIGHT", false, false, TOT_CBM);
                if (TOT_PP > 0)
                    WriteData(iRow, 8, "RIGHT", false, false, TOT_PP);
                if (TOT_CC > 0)
                    WriteData(iRow, 9, "RIGHT", false, false, TOT_CC);
                if (TOT_A > 0)
                    WriteData(iRow, 10, "RIGHT", false, false, TOT_A);
                if (TOT_HAULAGE > 0)
                    WriteData(iRow, 11, "RIGHT", false, false, TOT_HAULAGE);
                if (TOT_REBATE > 0)
                    WriteData(iRow, 12, "RIGHT", false, false, TOT_REBATE);
                for (int c = 1; c < 13; c++)
                {
                    ws.Cells[iRow, c].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);
                }
                iRow = iRow + 2;
                WriteData(iRow, 8, "RIGHT", false, false, "CMAR");
                WriteData(iRow, 9, "RIGHT", false, false, "WELLTON");
                WriteData(iRow, 10, "RIGHT", false, false, "TOTAL");
                iRow = iRow + 2;
                WriteData(iRow, 1, "LEFT", "A. FREIGHT COLLECT BY CMAR/ WELLTON");
                WriteData(iRow, 8, "RIGHT", false, true, TOT_PP);
                WriteData(iRow, 9, "RIGHT", false, true, TOT_CC);
                WriteData(iRow, 10, "RIGHT", false, true, TOT_A);

                iRow = iRow + 2;

                WriteData(iRow, 1, "LEFT", "B. EXPENSE INCURRED BY WELLTON");
                iRow++;
                WriteData(iRow, 1, "LEFT", "   OCEAN FREIGHT");
                if (TOT_BUY_COLLECT > 0)
                    WriteData(iRow, 9, "RIGHT", TOT_BUY_COLLECT);
                iRow++;
                if (BUY_DDC_CC > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   DDC");
                    WriteData(iRow, 9, "RIGHT", BUY_DDC_CC);
                    iRow++;
                }
                if (DESTUFF_P_D > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   DESTUFFING, P & D etc.");
                    WriteData(iRow, 9, "RIGHT", DESTUFF_P_D);
                    iRow++;
                }
                if (HANDLING_FEE > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   HANDLING FEE PAYABLE");
                    WriteData(iRow, 9, "RIGHT", HANDLING_FEE);
                    iRow++;
                }
                if (BUY_ACD_CC > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   ACD CHARGES");
                    WriteData(iRow, 9, "RIGHT", BUY_ACD_CC);
                    iRow++;
                }

                if (TOT_HAULAGE > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   ONFORWADING CHARGES");
                    WriteData(iRow, 9, "RIGHT", TOT_HAULAGE);
                    iRow++;
                }
                if (TOT_B > 0)
                {
                    WriteData(iRow, 9, "RIGHT", false, true, TOT_B);
                    ws.Cells[iRow, 9].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);
                }
                iRow = iRow + 2;

                WriteData(iRow, 1, "LEFT", "C. EXPENSE INCURRED BY CARGOMAR");
                iRow++;
                WriteData(iRow, 1, "LEFT", "   OCEAN FREIGHT");
                if (TOT_BUY_PREPAID > 0)
                    WriteData(iRow, 8, "RIGHT", TOT_BUY_PREPAID);
                iRow++;
                if (TRUK_COST > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   TRUCKING COST FOR HBL NOS.");
                    WriteData(iRow, 8, "RIGHT", TRUK_COST);
                    iRow++;
                }
                if (CNTR_SHIFT_CHRGS > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   ORIGIN CONTAINER SHIFTING CHARGES.");
                    WriteData(iRow, 8, "RIGHT", CNTR_SHIFT_CHRGS);
                    iRow++;
                }
                if (VESSEl_CHANGE_CHRGS > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   VESSEL CHANGE CHARGES");
                    WriteData(iRow, 8, "RIGHT", VESSEl_CHANGE_CHRGS);
                    iRow++;
                }
                if (EX_WORK_CHRGS > 0)
                {
                    foreach (DataRow dr in DTP_COSTINGD.Select("1=1", "BLNO"))
                        if (Lib.Convert2Decimal(dr["EX_WORK_WELLTON"].ToString()) > 0)
                        {
                            WriteData(iRow, 1, "LEFT", "   EX-WORK CHARGES HBL# " + dr["BLNO"].ToString());
                            WriteData(iRow, 8, "RIGHT", Lib.Convert2Decimal(Lib.NumericFormat(dr["EX_WORK_WELLTON"].ToString(), 3)));
                            iRow++;
                        }
                }
                if (TOT_REBATE > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   REBATE");
                    WriteData(iRow, 8, "RIGHT", TOT_REBATE);
                    iRow++;
                }

                if (BUY_ACD_PP > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   ACD CHARGES");
                    WriteData(iRow, 8, "RIGHT", BUY_ACD_PP);
                    iRow++;
                }
                if (BUY_AMENMENT_CHRGS > 0)
                {
                    WriteData(iRow, 1, "LEFT", "   AMMENDMENT CHARGES");
                    WriteData(iRow, 8, "RIGHT", BUY_AMENMENT_CHRGS);
                    iRow++;
                }
                if (TOT_C > 0)
                {
                    WriteData(iRow, 8, "RIGHT", false, true, TOT_C);
                    ws.Cells[iRow, 8].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);
                }
                iRow = iRow + 2;

                WriteData(iRow, 1, "LEFT", "D. NET PROFIT / LOSS (+/-) A-(B+C)");
                WriteData(iRow, 10, "RIGHT", false, true, TOT_A2);
                iRow++;
                WriteData(iRow, 1, "LEFT", "   PROFIT / LOSS (+/-) SHARE");
                WriteData(iRow, 8, "RIGHT", false, false, PROFIT);
                WriteData(iRow, 9, "RIGHT", false, false, PROFIT);
                ws.Cells[iRow, 8].Style.Borders.SetBorders(MultipleBorders.Bottom, System.Drawing.Color.Black, LineStyle.Thin);
                ws.Cells[iRow, 9].Style.Borders.SetBorders(MultipleBorders.Bottom, System.Drawing.Color.Black, LineStyle.Thin);
                iRow++;
                WriteData(iRow, 1, "LEFT", "E. TOTAL");
                TempAmt = TOT_C + PROFIT;
                WriteData(iRow, 8, "RIGHT", false, true, TempAmt);
                TempAmt = TOT_B + PROFIT;
                WriteData(iRow, 9, "RIGHT", false, true, TempAmt);
                iRow++;
                WriteData(iRow, 1, "LEFT", "F. FREIGHT COLLECT BY WELLTON");
                WriteData(iRow, 9, "RIGHT", false, false, TOT_CC);
                iRow++;
                iRow++;
                WriteData(iRow, 1, "LEFT", "   AMOUNT DUE TO WELLTON (E-F)");
                TempAmt = TOT_NETDUE * -1;
                WriteData(iRow, 9, "RIGHT", false, true, TempAmt);
                iRow++;
                iRow++;
                if (TempAmt < 0)
                    WriteData(iRow, 1, "LEFT", "   NET DUE FROM WELLTON ");
                else
                    WriteData(iRow, 1, "LEFT", "   NET DUE TO WELLTON ");
                TempAmt = Math.Abs(TempAmt);
                TempAmt = Lib.Convert2Decimal(Lib.NumericFormat(TempAmt.ToString(), 2));
                ws.Cells[iRow, 8].Style.NumberFormat = "#0.00";
                WriteData(iRow, 8, "RIGHT", false, true, TempAmt);
                for (int c = 1; c < 13; c++)
                {
                    ws.Cells[iRow, c].Style.Borders.SetBorders(MultipleBorders.Horizontal, System.Drawing.Color.Black, LineStyle.Thin);
                }
                for (int r = 3; r <= iRow; r++)
                {
                    ws.Cells[r, 1].Style.Borders.SetBorders(MultipleBorders.Left, System.Drawing.Color.Black, LineStyle.Thin);
                    ws.Cells[r, 12].Style.Borders.SetBorders(MultipleBorders.Right, System.Drawing.Color.Black, LineStyle.Thin);
                }
                iRow++;
                iRow++;
                WriteData(iRow, 1, "LEFT", "P=Prepaid, C=Collect");

                file.SaveXls(File_Name + ".xls");
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        private decimal GetHaulageRate(string gCbm, string gGrwt)
        {
            decimal hRt = 0, GrwtCbm = 0;
            //GrwtCbm = Common.Convert2Decimal(gGrwt) / WEIGHT_DIVIDER;
            //if (Common.Convert2Decimal(gCbm) > GrwtCbm)
            GrwtCbm = Lib.Convert2Decimal(gCbm);

            hRt = GrwtCbm * HAULAGE_PER_CBM;
            hRt = Lib.Convert2Decimal(Lib.NumericFormat(hRt.ToString(), 3));
            return hRt;
        }
        private string GetCbmRate(string sRate, string sChar)
        {
            decimal sRt = 0;
            sRt = Lib.Convert2Decimal(sRate);
            return Lib.NumericFormat(sRt.ToString(), 2) + sChar;
        }
        /******************WELLTON************************/

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

                sql = " select cost_pkid, cost_type, cost_source, cost_cfno,cost_refno,cost_folderno,cost_mblid,mbl.hbl_bl_no as cost_mblno,cost_cntr as cost_book_cntr";
                sql += " ,mbl.hbl_date as  cost_sob_date, mbl.hbl_folder_sent_date as cost_folder_recdon,  cost_agent_id,agnt.cust_code as cost_agent_code,agnt.cust_name as cost_agent_name,cost_year,cost_date";
                sql += " ,cost_edit_code,cost_exrate,cost_currency_id,c.param_code as cost_currency_code ,cost_rebate";
                sql += " ,cost_ex_works,cost_inform_rate,cost_other_charges,cost_hand_charges ";
                sql += " ,cost_buy_pp,cost_buy_cc,cost_sell_pp,cost_sell_cc,cost_format";
                sql += " ,cost_buy_tot,cost_sell_tot";
                sql += " ,cost_profit,cost_our_profit,cost_your_profit,cost_drcr_amount,cost_income,cost_expense,cost_sell_chwt ";
                sql += " ,cost_jv_agent_id,agnt2.cust_code as cost_jv_agent_code,agnt2.cust_name as cost_jv_agent_name,cost_jv_agent_br_id";
                sql += " ,agntaddr.add_branch_slno as  cost_jv_agent_br_no,agntaddr.add_line1||'\n'||agntaddr.add_line2||'\n'||agntaddr.add_line3 as  cost_jv_agent_br_addr,cost_jv_br_inv_id  ";
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
                    mRow.cost_book_cntr = Dr["cost_book_cntr"].ToString();
                    mRow.cost_jv_agent_id = Dr["cost_jv_agent_id"].ToString();
                    mRow.cost_jv_agent_code = Dr["cost_jv_agent_code"].ToString();
                    mRow.cost_jv_agent_name = Dr["cost_jv_agent_name"].ToString();
                    mRow.cost_jv_agent_br_id = Dr["cost_jv_agent_br_id"].ToString();
                    mRow.cost_jv_agent_br_no = Dr["cost_jv_agent_br_no"].ToString();
                    mRow.cost_jv_agent_br_addr = Dr["cost_jv_agent_br_addr"].ToString();
                    mRow.cost_jv_br_inv_id = Dr["cost_jv_br_inv_id"].ToString();

                    Tot_InvoiceAmt = Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString());


                    if (Dr["cost_agent_name"].ToString().Contains("RITRA"))
                        AGENT_FORMAT = "RITRA";
                    else if (Dr["cost_agent_name"].ToString().Contains("TRAFFIC TECH"))
                        AGENT_FORMAT = "TRAFFIC-TECH";
                    
                    break;
                }

                if (bok)
                {
                    List<Costingd> mList = new List<Costingd>();
                    Costingd bRow;

                    sql = "select costd_pkid,  costd_parent_id,costd_acc_id ";
                    sql += " ,costd_type,costd_grwt,costd_chwt,costd_cbm,h.hbl_cbm ";
                    sql += " ,h.hbl_no as hbl_no,costd_consgne_grp";
                    sql += " ,shpr.cust_name  as shipper_name";
                    sql += " ,cnge.cust_name as consignee_name";
                    sql += " ,h.hbl_terms as hbl_terms";
                    sql += " ,h.hbl_nomination as hbl_nomination";
                    sql += " ,pofd.param_name as pofd_name";
                    sql += " ,costd_frt_pp,costd_frt_cc,costd_frt_rate_pp,costd_frt_rate_cc";
                    sql += " ,costd_wrs_pp,costd_wrs_cc,costd_wrs_rate_pp,costd_wrs_rate_cc";
                    sql += " ,costd_myc_pp,costd_myc_cc,costd_myc_rate_pp,costd_myc_rate_cc";
                    sql += " ,costd_mcc_pp,costd_mcc_cc,costd_mcc_rate_pp,costd_mcc_rate_cc";
                    sql += " ,costd_src_pp,costd_src_cc,costd_src_rate_pp,costd_src_rate_cc";
                    sql += " ,costd_oth_pp,costd_oth_cc,costd_cc,costd_pp,costd_oth_rate_pp,costd_oth_rate_cc,costd_tot,h.hbl_bl_no as bl_no ";
                    sql += " ,costd_category,costd_ctr,costd_cost_data,costd_rebate,costd_amnt_chrgs,costd_consgne_grp,costd_seal_chrgs ";
                    sql += " from costingd a ";
                    sql += " left join hblm h on a.costd_acc_id = h.hbl_pkid";
                    sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid ";
                    sql += " left join customerm cnge on h.hbl_imp_id = cnge.cust_pkid ";
                    sql += " left join param pofd on h.hbl_pofd_id = pofd.param_pkid ";
                    sql += " where costd_parent_id ='{ID}' ";
                    sql += " and nvl(costd_category,'COSTING') = 'COSTING' ";
                    sql += " order by costd_ctr ";
                    sql = sql.Replace("{ID}", id);


                    Dt_Rec = new DataTable();
                    Con_Oracle = new DBConnection();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    decimal tot_hbl_cbm = Lib.Convert2Decimal(Dt_Rec.Compute("sum(hbl_cbm)", "1=1").ToString());
                    string[] Sdata = null;
                    foreach (DataRow Dr in Dt_Rec.Rows)
                    {
                        bRow = new Costingd();
                        bRow.costd_pkid = Dr["costd_pkid"].ToString();
                        bRow.costd_parent_id = Dr["costd_parent_id"].ToString();
                        bRow.costd_acc_id = Dr["costd_acc_id"].ToString();
                        bRow.costd_type = Dr["costd_type"].ToString();
                        bRow.costd_sino = Dr["hbl_no"].ToString();
                        bRow.costd_shipper_name = Dr["shipper_name"].ToString();
                        bRow.costd_consignee_name = Dr["consignee_name"].ToString();
                        bRow.costd_consignee_group = Dr["costd_consgne_grp"].ToString();
                        bRow.costd_hbl_nomination = Dr["hbl_nomination"].ToString();
                        bRow.costd_hbl_terms = Dr["hbl_terms"].ToString();
                        bRow.costd_pofd_name = Dr["pofd_name"].ToString();
                        bRow.costd_grwt = Lib.Conv2Decimal(Dr["costd_grwt"].ToString());
                        bRow.costd_chwt = Lib.Conv2Decimal(Dr["costd_chwt"].ToString());
                        bRow.costd_frt_pp = Lib.Conv2Decimal(Dr["costd_frt_pp"].ToString());
                        bRow.costd_frt_cc = Lib.Conv2Decimal(Dr["costd_frt_cc"].ToString());
                        bRow.costd_frt_rate_pp = Lib.Conv2Decimal(Dr["costd_frt_rate_pp"].ToString());
                        bRow.costd_frt_rate_cc = Lib.Conv2Decimal(Dr["costd_frt_rate_cc"].ToString());
                        bRow.costd_baf_pp = Lib.Conv2Decimal(Dr["costd_wrs_pp"].ToString());
                        bRow.costd_baf_cc = Lib.Conv2Decimal(Dr["costd_wrs_cc"].ToString());
                        bRow.costd_baf_rate_pp = Lib.Conv2Decimal(Dr["costd_wrs_rate_pp"].ToString());
                        bRow.costd_baf_rate_cc = Lib.Conv2Decimal(Dr["costd_wrs_rate_cc"].ToString());
                        bRow.costd_caf_pp = Lib.Conv2Decimal(Dr["costd_myc_pp"].ToString());
                        bRow.costd_caf_cc = Lib.Conv2Decimal(Dr["costd_myc_cc"].ToString());
                        bRow.costd_caf_rate_pp = Lib.Conv2Decimal(Dr["costd_myc_rate_pp"].ToString());
                        bRow.costd_caf_rate_cc = Lib.Conv2Decimal(Dr["costd_myc_rate_cc"].ToString());
                        bRow.costd_ddc_pp = Lib.Conv2Decimal(Dr["costd_mcc_pp"].ToString());
                        bRow.costd_ddc_cc = Lib.Conv2Decimal(Dr["costd_mcc_cc"].ToString());
                        bRow.costd_ddc_rate_pp = Lib.Conv2Decimal(Dr["costd_mcc_rate_pp"].ToString());
                        bRow.costd_ddc_rate_cc = Lib.Conv2Decimal(Dr["costd_mcc_rate_cc"].ToString());
                        bRow.costd_acd_pp = Lib.Conv2Decimal(Dr["costd_src_pp"].ToString());
                        bRow.costd_acd_cc = Lib.Conv2Decimal(Dr["costd_src_cc"].ToString());
                        bRow.costd_acd_rate_pp = Lib.Conv2Decimal(Dr["costd_src_rate_pp"].ToString());
                        bRow.costd_acd_rate_cc = Lib.Conv2Decimal(Dr["costd_src_rate_cc"].ToString());
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
                        bRow.costd_cbm = Lib.Conv2Decimal(Dr["costd_cbm"].ToString());
                        if (bRow.costd_type == "BUY")
                            bRow.costd_actual_cbm = tot_hbl_cbm;
                        else
                            bRow.costd_actual_cbm = Lib.Conv2Decimal(Dr["hbl_cbm"].ToString());
                        bRow.costd_rebate = Lib.Conv2Decimal(Dr["costd_rebate"].ToString());
                        bRow.costd_amenment_chrgs = Lib.Conv2Decimal(Dr["costd_amnt_chrgs"].ToString());
                        bRow.costd_seal_chrgs = Lib.Conv2Decimal(Dr["costd_seal_chrgs"].ToString());
                        bRow.costd_consignee_group = Dr["costd_consgne_grp"].ToString();
                        bRow.costd_agent_format = AGENT_FORMAT;

                        if (Dr["costd_cost_data"].ToString() != "")
                        {
                            Sdata = Dr["costd_cost_data"].ToString().Split(',');
                            bRow.costd_hbl_nomination = Sdata[0];
                            bRow.costd_hbl_terms = Sdata[1];
                            bRow.costd_incentive_rate = Lib.Convert2Decimal(Sdata[2]);
                            bRow.costd_fh_rate1 = Lib.Convert2Decimal(Sdata[3]);
                            bRow.costd_fh_limit1 = Lib.Convert2Decimal(Sdata[4]);
                            bRow.costd_fh_rate2 = Lib.Convert2Decimal(Sdata[5]);
                            bRow.costd_fh_limit2 = Lib.Convert2Decimal(Sdata[6]);
                            bRow.costd_fh_rate3 = Lib.Convert2Decimal(Sdata[7]);
                            bRow.costd_fh_limit3 = bRow.costd_fh_limit2;
                            bRow.costd_pofd_name = Sdata[8];
                            bRow.costd_oth_chrgs_ritra = Lib.Convert2Decimal(Sdata[9]);
                            bRow.costd_incentive_notreceived = Sdata[10] == "Y" ? true : false;
                            bRow.costd_fh_chrg_perhouse = Lib.Convert2Decimal(Sdata[11]);
                            bRow.costd_spl_incentive_rate = Lib.Convert2Decimal(Sdata[12]);
                            bRow.costd_house_notinclude = Sdata[13] == "Y" ? true : false;
                            bRow.costd_ex_chrg_ritrahouse = Sdata.Length > 14 ? Lib.Convert2Decimal(Sdata[14]) : 0;

                            bRow.costd_haulage_per_cbm = Sdata.Length > 15 ? Lib.Convert2Decimal(Sdata[15]) : 0;
                            bRow.costd_haulage_min_rate = Sdata.Length > 16 ? Lib.Convert2Decimal(Sdata[16]) : 0;
                            bRow.costd_haulage_wt_divider = Sdata.Length > 17 ? Lib.Convert2Decimal(Sdata[17]) : 0;
                            bRow.costd_destuff_pd = Sdata.Length > 18 ? Lib.Convert2Decimal(Sdata[18]) : 0;
                            bRow.costd_handling_fee = Sdata.Length > 19 ? Lib.Convert2Decimal(Sdata[19]) : 0;
                            bRow.costd_truck_cost = Sdata.Length > 20 ? Lib.Convert2Decimal(Sdata[20]) : 0;
                            bRow.costd_cntr_shifit = Sdata.Length > 21 ? Lib.Convert2Decimal(Sdata[21]) : 0;
                            bRow.costd_vessel_chrgs = Sdata.Length > 22 ? Lib.Convert2Decimal(Sdata[22]) : 0;
                            bRow.costd_ex_works = Sdata.Length > 23 ? Lib.Convert2Decimal(Sdata[23]) : 0;
                        }
                        mList.Add(bRow);
                    }
                    mRow.DetailList2 = mList;

                    mList = new List<Costingd>();

                    sql = "select costd_pkid,  costd_parent_id,  costd_acc_id,  costd_acc_name , costd_category, ";
                    sql += " costd_blno,costd_acc_qty,costd_acc_rate,";
                    sql += " costd_acc_amt,costd_ctr,costd_remarks,costd_brate,costd_srate,costd_split ";
                    sql += " from costingd a ";
                    sql += " where costd_parent_id ='{ID}' ";
                    sql += " and nvl(costd_category,'COSTING') = 'INVOICE' ";
                    sql += " order by costd_ctr ";
                    sql = sql.Replace("{ID}", id);

                    Con_Oracle = new DBConnection();
                    Dt_Rec = new DataTable();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
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
                        bRow.costd_brate = Lib.Conv2Decimal(Dr["costd_brate"].ToString());
                        bRow.costd_srate = Lib.Conv2Decimal(Dr["costd_srate"].ToString());
                        bRow.costd_split = Lib.Conv2Decimal(Dr["costd_split"].ToString());
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
                    Lib.AddError(ref str, " | Folder No cannot be blank");


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
                    sql += " and a.cost_source ='SE CONSOLE COSTING'";
                    sql += ") a where cost_pkid <> '{PKID}'";

                    sql = sql.Replace("{FOLDERNO}", Record.cost_folderno);
                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                    sql = sql.Replace("{CATEGORY}", Record.rec_category);
                    sql = sql.Replace("{PKID}", Record.cost_pkid);

                    if (Con_Oracle.IsRowExists(sql))
                    {
                        bError = true;
                        Lib.AddError(ref str, " | Folder No already Exists");
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


            decimal nDrcrInr = 0;

            try
            {
                Con_Oracle = new DBConnection();

                if (Record.cost_folderno.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Folder NO Cannot Be Empty");

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
                    DataRow lovRow_Doc_Prefix = lov.getSettings(Record._globalvariables.branch_code, "COST-PREFIX");
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
                        sql += " and a.cost_source in ('SEA EXPORT COSTING','SE CONSOLE COSTING','DRCR ISSUE')";
                        sql += " and a.cost_type = 'SEA' ";

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
                    nDrcrInr = Lib.Conv2Decimal(Record.cost_drcr_amount.ToString()) * Lib.Conv2Decimal(Record.cost_exrate.ToString());
                    nDrcrInr = Lib.RoundNumber_Latest(nDrcrInr.ToString(), 2, true);
                }
                if (nDrcrInr > 0)
                    Rec.InsertString("cost_drcr", "DR");
                else
                    Rec.InsertString("cost_drcr", "CR");
                Rec.InsertNumeric("cost_drcr_amount", Record.cost_drcr_amount.ToString());
                Rec.InsertNumeric("cost_drcr_amount_inr", nDrcrInr.ToString());
                Rec.InsertNumeric("cost_sell_chwt", Record.cost_sell_chwt.ToString());
                Rec.InsertString("cost_jv_agent_id", Record.cost_jv_agent_id);
                Rec.InsertString("cost_jv_agent_br_id", Record.cost_jv_agent_br_id);
                Rec.InsertString("cost_cntr", Record.cost_book_cntr);

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
                SaveCostingDet(Record,"SAVE");
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

                InvList = new List<Costingd>();
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
                File_Display_Name = "myreport.xls";
                report_folder = System.IO.Path.Combine(report_folder, report_pkid);
                File_Name = System.IO.Path.Combine(report_folder, report_pkid);
                
                if (report_type == "EXCEL") //Invoice
                {
                    Con_Oracle = new DBConnection();

                    sql = "";

                    sql = " select  cost_refno, cost_folderno,cost_date, b.hbl_bl_no, hbl_folder_no, b.hbl_type,b.rec_category, ";
                    sql += " agent.cust_name as agent_name,  ";
                    sql += " agentadd.add_line1 as agent_line1,";
                    sql += " agentadd.add_line2 as agent_line2,";
                    sql += " agentadd.add_line3 as agent_line3,";
                    sql += " agentadd.add_line4 as agent_line4,";
                    sql += " vsl.param_name as vessel_name, hbl_vessel_no,";
                    sql += " pol.param_name as pol_name,";
                    sql += " pod.param_name as pod_name,";
                    sql += " curr.param_code as curr_code,";
                    sql += " cost_exrate,cost_buy_pp,cost_buy_cc,";
                    sql += " cost_sell_pp,cost_sell_cc,cost_rebate,";
                    sql += " cost_ex_works,cost_hand_charges,cost_kamai,";
                    sql += " cost_other_charges,cost_asper_amount,cost_buy_tot,";
                    sql += " cost_sell_tot,cost_profit ,cost_our_profit,cost_your_profit,";
                    sql += " cost_drcr_amount,cost_drcr_amount_inr,cost_expense,cost_income";
                    sql += " from costingm a";
                    sql += " inner join hblm b on a.cost_mblid = b.hbl_pkid";
                    //sql += " left join customerm agent on b.hbl_agent_id = agent.cust_pkid";
                    //sql += " left join addressm agentadd on hbl_agent_br_id = agentadd.add_pkid";
                    sql += " left join customerm agent on a.cost_jv_agent_id = agent.cust_pkid";
                    sql += " left join addressm agentadd on a.cost_jv_agent_br_id = agentadd.add_pkid";
                    sql += " left join param vsl on hbl_vessel_id = vsl.param_pkid";
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
                    sql += " select nvl(h.hbl_bl_no,h.hbl_fcr_no) as hbl_bl_no, cons.cust_name as consignee_name";
                    sql += " from costingm a";
                    sql += " inner join hblm m on a.cost_mblid = m.hbl_pkid";
                    sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
                    sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
                    sql += " where cost_pkid ='" + report_pkid + "'";

                    dt_house = Con_Oracle.ExecuteQuery(sql);

                    sql = "";
                    sql += " select  cntr_no,ctype.param_code as cntr_type ";
                    sql += " from costingm a";
                    sql += " inner join containerm c on cost_mblid = cntr_booking_id";
                    sql += " left join param ctype on cntr_type_id = ctype.param_pkid ";
                    sql += " where cost_pkid ='" + report_pkid + "'";

                    dt_cntr = Con_Oracle.ExecuteQuery(sql);

                    sql = "";
                    sql += " select  costd_acc_name ,costd_acc_amt,costd_remarks ";
                    sql += " from costingd ";
                    sql += " where costd_parent_id ='" + report_pkid + "'";
                    sql += " and costd_category ='INVOICE' ";
                    sql += " order by costd_ctr";
                    dt_costDet = Con_Oracle.ExecuteQuery(sql);

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

                    if (Lib.CreateFolder(report_folder))
                        ProcessDetailFile();
                }
                if (report_type == "EXCEL2")//statement
                {

                    if (Lib.CreateFolder(report_folder))
                    {
                        LoadTables(id);
                        if (DTP_DESTN.Rows.Count > 0)
                        {
                            File_Display_Name = Lib.ProperFileName(DTP_DESTN.Rows[0]["COST_REFNO"].ToString() + "-" + DTP_DESTN.Rows[0]["COST_FOLDERNO"].ToString()) + ".xls";
                            File_Name = Lib.GetFileName(report_folder, report_pkid, File_Display_Name);
                        }
                        if (AGENT_FORMAT == "RITRA")
                            Print_RITRA();
                        else if (AGENT_FORMAT == "TRAFFIC-TECH")
                            Print_WELLTON();
                    }
                }
                if (report_type == "FILL-INVOICE") //Fill Invoice from Statement
                {
                    LoadTables(id);
                    //if (DTP_DESTN.Rows.Count > 0)
                    //{
                    //    File_Display_Name = Lib.ProperFileName(DTP_DESTN.Rows[0]["COST_REFNO"].ToString() + "-" + DTP_DESTN.Rows[0]["COST_FOLDERNO"].ToString()) + ".xls";
                    //    File_Name = Lib.GetFileName(report_folder, report_pkid, File_Display_Name);
                    //}
                    if (AGENT_FORMAT == "RITRA")
                        Fill_RITRAInvoice();
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
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", InvList);
            return RetData;
        }


        //private void ProcessExcelFile()
        //{
        //    string _Border = "";
        //    Boolean _Bold = false;
        //    Color _Color = Color.Black;
        //    int _Size = 0;

        //    decimal nDrCRAmt = 0;


        //    decimal buy_pp = 0;
        //    decimal buy_cc = 0;
        //    decimal buy_tot = 0;

        //    decimal sell_pp = 0;
        //    decimal sell_cc = 0;
        //    decimal sell_tot = 0;


        //    decimal kamai = 0;

        //    decimal rebate = 0;
        //    decimal exwork = 0;
        //    decimal other = 0;

        //    decimal income = 0;
        //    decimal expense = 0;

        //    decimal profit = 0;
        //    decimal our_profit = 0;
        //    decimal your_profit = 0;


        //    string sTitle = "";

        //    string sName = "Report";
        //    WB = new ExcelFile();
        //    WB.Worksheets.Add(sName);
        //    WS = WB.Worksheets[sName];
        //    WS.PrintOptions.Portrait = true;
        //    WS.PrintOptions.FitWorksheetWidthToPages = 1;

        //    WS.Columns[0].Width = 256;
        //    WS.Columns[1].Width = 256 * 16;
        //    WS.Columns[2].Width = 256 * 10;
        //    WS.Columns[3].Width = 256 * 10;
        //    WS.Columns[4].Width = 256 * 10;
        //    WS.Columns[5].Width = 256 * 10;
        //    WS.Columns[6].Width = 256 * 12;
        //    WS.Columns[7].Width = 256 * 10;
        //    WS.Columns[8].Width = 256 * 10;
        //    WS.Columns[9].Width = 256 * 10;
        //    WS.Columns[10].Width = 256 * 12;
        //    WS.Columns[11].Width = 256 * 12;
        //    WS.Columns[12].Width = 256 * 12;

        //    iRow = 1; iCol = 1;

        //    //iRow = Lib.WriteHoAddress(WS, report_comp_code, iRow, iCol,7,1,true);


        //    string comp_name = "";
        //    string comp_add1 = "";
        //    string comp_add2 = "";
        //    string comp_add3 = "";
        //    string comp_add4 = "";


        //    Dictionary<string, object> mSearchData = new Dictionary<string, object>();
        //    LovService mService = new LovService();
        //    mSearchData.Add("table", "COMP_ADDRESS");
        //    mSearchData.Add("comp_code", report_comp_code);
        //    DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
        //    if (Dt_CompAddress != null)
        //    {
        //        foreach (DataRow Dr in Dt_CompAddress.Rows)
        //        {
        //            comp_name = Dr["COMP_NAME"].ToString();
        //            comp_add1 = Dr["COMP_ADDRESS1"].ToString();
        //            comp_add2 = Dr["COMP_ADDRESS2"].ToString();
        //            comp_add3 = Dr["COMP_ADDRESS3"].ToString();
        //            // comp_add4 = "Email : " + Dr["COMP_email"].ToString() + " Web : " + Dr["COMP_WEB"].ToString();
        //            comp_add4 = "Email : hodoc@cargomar.in Web : " + Dr["COMP_WEB"].ToString();
        //            break;
        //        }
        //    }

        //    iRow = 1; iCol = 1;
        //    _Color = Color.Black;
        //    _Size = 16;
        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_name, "Calibri", 14, true, Color.Black, "C", "C", "", "");
        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add1, "Calibri", 12, false, Color.Black, "C", "C", "", "");
        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add2, "Calibri", 12, false, Color.Black, "C", "C", "", "");
        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add3, "Calibri", 12, false, Color.Black, "C", "C", "", "");
        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, comp_add4, "Calibri", 12, false, Color.Black, "C", "C", "", "");

        //    DateTime Dt;

        //    string sDate = ((DateTime)DR_MASTER["cost_date"]).ToString("dd/MM/yyyy");
        //    string sCntr = "";
        //    string Str = "";

        //    iRow++;

        //    if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
        //        sTitle = "DEBIT NOTE";
        //    if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
        //        sTitle = "CREDIT NOTE";

        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 2, sTitle, "Calibri", 18, true, Color.Black, "C", "C", "TB", "THIN");

        //    iRow += 2;

        //    _Size = 13;

        //    Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);

        //    Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 9 : 5), "NUMBER", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 10 : 6), DR_MASTER["COST_REFNO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

        //    File_Display_Name = DR_MASTER["COST_REFNO"].ToString() + "-" + DR_MASTER["COST_FOLDERNO"].ToString() + ".xls";

        //    iRow++;

        //    Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_LINE1"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

        //    Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 9 : 5), "DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow, (DR_MASTER["cost_format"].ToString() == "PC" ? 10 : 6), sDate, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    iRow++;
        //    Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE2"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE3"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE4"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


        //    if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
        //        sTitle = "WE DEBIT YOUR ACCOUNT FOR THE FOLLOWING";
        //    if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
        //        sTitle = "WE CREDIT YOUR ACCOUNT FOR THE FOLLOWING";

        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, sTitle, "Calibri", 12, true, Color.Black, "C", "C", "TB", "THIN");

        //    Lib.WriteData(WS, iRow, 1, "MAWB NO", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow, 3, DR_MASTER["HBL_BL_NO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

        //    iRow++;
        //    Lib.WriteData(WS, iRow, 1, "HAWB NO", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

        //    Str = "";
        //    foreach (DataRow Dr in dt_house.Rows)
        //    {
        //        if (Str != "")
        //            Str += ",";
        //        Str += Dr["hbl_bl_no"].ToString();
        //    }

        //    Lib.WriteMergeCell(WS, iRow, 3, 5, 1, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
        //    iRow++;

        //    Lib.WriteData(WS, iRow, 1, "CONSIGNEE", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

        //    int iCount = 0;
        //    Str = "";
        //    String PreStr = "1";
        //    foreach (DataRow Dr in dt_house.Select("1=1", "consignee_name"))
        //    {
        //        if (Str != "")
        //            Str += "\n";

        //        if (PreStr != Dr["consignee_name"].ToString())
        //        {
        //            PreStr = Dr["consignee_name"].ToString();
        //            Str += Dr["consignee_name"].ToString();
        //            iCount++;
        //        }
        //    }
        //    if (iCount == 0)
        //        iCount = 1;

        //    Lib.WriteMergeCell(WS, iRow, 3, 5, iCount, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
        //    iRow += iCount;

        //    Lib.WriteData(WS, iRow, 1, "MAWB DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow, 3, Lib.DatetoStringDisplayformat(DR_MASTER["HBL_DATE"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    iRow++;
        //    Lib.WriteData(WS, iRow, 1, "AIRPORT OF DEPARTURE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow, 3, DR_MASTER["POL_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    iRow++;

        //    Lib.WriteData(WS, iRow, 1, "AIRPORT OF DESTINATION", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    Lib.WriteData(WS, iRow, 3, DR_MASTER["POD_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
        //    iRow++;

        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, "", "Calibri", 11, true, Color.Black, "C", "C", "T", "THIN");

        //    iCol = 1;
        //    _Color = Color.Black;
        //    _Border = "";
        //    _Size = 12;
        //    if (DR_MASTER["cost_format"].ToString() == "HANDLING")
        //    {
        //        decimal nAmt = 0;
        //        iRow += 4;
        //        Lib.WriteData(WS, iRow, 2, "PARTICULARS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 7, "AMOUNT(" + DR_MASTER["curr_code"].ToString() + ")", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);

        //        iRow++;

        //        if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
        //            Lib.WriteData(WS, iRow, 2, "OUR HANDLING CHARGES", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        else
        //            Lib.WriteData(WS, iRow, 2, "YOUR HANDLING CHARGES", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        nAmt = GetConvertAmt(DR_MASTER["cost_hand_charges"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //        Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow++;

        //        if (Lib.Conv2Decimal(DR_MASTER["cost_buy_pp"].ToString()) != 0)
        //        {
        //            Lib.WriteData(WS, iRow, 2, "BUY PP", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            nAmt = GetConvertAmt(DR_MASTER["cost_buy_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            iRow++;
        //        }

        //        if (Lib.Conv2Decimal(DR_MASTER["cost_ex_works"].ToString()) != 0)
        //        {
        //            Lib.WriteData(WS, iRow, 2, "EX.Work", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            nAmt = GetConvertAmt(DR_MASTER["cost_ex_works"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            iRow++;
        //        }
        //        if (Lib.Conv2Decimal(DR_MASTER["cost_other_charges"].ToString()) != 0)
        //        {
        //            Lib.WriteData(WS, iRow, 2, "Other Charges", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            nAmt = GetConvertAmt(DR_MASTER["cost_other_charges"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 7, nAmt, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            iRow++;
        //        }

        //        iRow++;

        //        Lib.WriteData(WS, iRow, 2, "TOTAL", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 7, DR_MASTER["cost_drcr_amount"], _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow += 6;

        //        nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());

        //    }
        //    else if (DR_MASTER["cost_format"].ToString() == "PP")
        //    {
        //        _Border = "LTBR";
        //        iRow += 6;
        //        decimal sellrate = 0;
        //        decimal informrate = 0;
        //        decimal irate = 0;
        //        decimal nAmt = 0;
        //        foreach (DataRow dr in dt_costdet.Select("costd_type = 'SELL'", "costd_ctr"))
        //        {
        //            sellrate += Lib.Conv2Decimal(dr["costd_frt_rate_pp"].ToString());
        //            sellrate += Lib.Conv2Decimal(dr["costd_frt_rate_cc"].ToString());
        //            break;
        //        }

        //        Lib.WriteMergeCell(WS, iRow, 1, 3, 2, "YOUR HANDLING CHARGES", "Calibri", 11, true, Color.Black, "C", "C", "TBLR", "THIN");
        //        Lib.WriteData(WS, iRow, 4, "INR/KG", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 5, "KGS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 6, "TOTAL INR", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 7, "TOTAL", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //        iRow++;

        //        informrate = Lib.Conv2Decimal(DR_MASTER["cost_inform_rate"].ToString());
        //        irate = (sellrate - informrate) / 2;
        //        irate = Lib.Conv2Decimal(Lib.NumericFormat(irate.ToString(), 2));

        //        //nAmt = GetConvertAmt(irate.ToString(), DR_MASTER["cost_exrate"].ToString());
        //        //Lib.WriteData(WS, iRow, 4, nAmt, _Color, false, _Border, "L", "", _Size, false, 325, "#,0.00", true);
        //        Lib.WriteData(WS, iRow, 4, irate, _Color, false, _Border, "L", "", _Size, false, 325, "#,0.00", true);

        //        Lib.WriteData(WS, iRow, 5, DR_MASTER["cost_sell_chwt"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        decimal nDrCRAmt_INR = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount_inr"].ToString());
        //        if (nDrCRAmt_INR < 0)
        //            nDrCRAmt_INR = Math.Abs(nDrCRAmt_INR);
        //        Lib.WriteData(WS, iRow, 6, nDrCRAmt_INR, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());
        //        if (nDrCRAmt < 0)
        //            nDrCRAmt = Math.Abs(nDrCRAmt);
        //        Lib.WriteData(WS, iRow, 7, nDrCRAmt, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        iRow += 6;
        //    }
        //    else
        //    {

        //        buy_pp = Lib.Conv2Decimal(DR_MASTER["COST_BUY_PP"].ToString());
        //        buy_cc = Lib.Conv2Decimal(DR_MASTER["COST_BUY_CC"].ToString());
        //        buy_tot = buy_pp + buy_cc;

        //        sell_pp = Lib.Conv2Decimal(DR_MASTER["COST_SELL_PP"].ToString());
        //        sell_cc = Lib.Conv2Decimal(DR_MASTER["COST_SELL_CC"].ToString());
        //        sell_tot = sell_pp + sell_cc;


        //        rebate = Lib.Conv2Decimal(DR_MASTER["COST_REBATE"].ToString());
        //        exwork = Lib.Conv2Decimal(DR_MASTER["COST_EX_WORKS"].ToString());
        //        other = Lib.Conv2Decimal(DR_MASTER["COST_OTHER_CHARGES"].ToString());


        //        decimal income_pp = 0;
        //        decimal income_cc = 0;
        //        decimal expense_pp = 0;
        //        decimal expense_cc = 0;

        //        income = 0;
        //        expense = 0;

        //        profit = Lib.Conv2Decimal(DR_MASTER["COST_PROFIT"].ToString());
        //        profit = GetConvertAmt(profit.ToString(), DR_MASTER["cost_exrate"].ToString());
        //        our_profit = Lib.Conv2Decimal(DR_MASTER["COST_OUR_PROFIT"].ToString());
        //        our_profit = GetConvertAmt(our_profit.ToString(), DR_MASTER["cost_exrate"].ToString());
        //        your_profit = Lib.Conv2Decimal(DR_MASTER["COST_YOUR_PROFIT"].ToString());
        //        your_profit = GetConvertAmt(your_profit.ToString(), DR_MASTER["cost_exrate"].ToString());

        //        iRow += 2;
        //        Str = "A.INCOME FREIGHT COLLECTED BY " + DR_MASTER["agent_name"].ToString();
        //        Lib.WriteData(WS, iRow, 1, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        //Lib.WriteData(WS, iRow, 11, sell_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        //Lib.WriteData(WS, iRow, 12, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow++;
        //        _Border = "LTRB";
        //        Str = DR_MASTER["curr_code"].ToString();
        //        Lib.WriteMergeCell(WS, iRow, 1, 1, 3, "HAWB", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 2, 1, 3, "FRT STATUS", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 3, 1, 3, "WEIGHT", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 4, 1, 3, "RATE/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 5, 1, 3, "FSC/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 6, 1, 3, "SRC/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 7, 1, 3, "WRS/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 8, 1, 3, "MCC/KG (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 9, 1, 3, "OTHERS (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 10, 1, 3, "TOTAL PP (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 11, 1, 3, "TOTAL CC (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 12, 1, 3, "TOTAL (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //      //  Lib.WriteMergeCell(WS, iRow, 13, 1, 3, "TOTAL (INR)", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        iRow++; iRow++;
        //        bool bok = false;
        //        decimal nAmt = 0;  
        //        foreach (DataRow dr in dt_costdet.Select("costd_type = 'SELL'", "costd_ctr"))
        //        {
        //            bok = true;
        //            iRow++;

        //            Lib.WriteData(WS, iRow, 1, dr["hbl_bl_no"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            Str = "";
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                Str = "PPD";
        //            else if (dr["hbl_terms"].ToString() == "FREIGHT COLLECT")
        //                Str = "FOB";
        //            Lib.WriteData(WS, iRow, 2, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 3, dr["costd_chwt"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_frt_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_frt_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 4, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_myc_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_myc_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 5, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_src_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_src_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 6, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_wrs_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_wrs_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 7, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_mcc_rate_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_mcc_rate_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 8, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);

        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_oth_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_oth_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 9, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true); 
        //            nAmt = GetConvertAmt(dr["costd_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 10, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            nAmt = GetConvertAmt(dr["costd_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 11, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            nAmt = GetConvertAmt(dr["costd_tot"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 12, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //       //     Lib.WriteData(WS, iRow, 13, dr["costd_tot"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        }
        //        if (bok)
        //        {
        //            iRow++;
        //            Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            Str = "";
        //            Lib.WriteData(WS, iRow, 2, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 3, DR_MASTER["cost_sell_chwt"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 4, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 5, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 6, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 7, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 8, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 9, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            sell_pp = GetConvertAmt(sell_pp.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 10, sell_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            sell_cc = GetConvertAmt(sell_cc.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 11, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //          //  nAmt = sell_tot;
        //            sell_tot = GetConvertAmt(sell_tot.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 12, sell_tot, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //          //  Lib.WriteData(WS, iRow, 13, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        }

        //        iRow++;
        //        iRow++;
        //        _Border = "";

        //        income_pp += sell_pp;
        //        income_cc += sell_cc;
        //        income += sell_tot;

        //        //Str = "A.INCOME FREIGHT COLLECTED BY " + DR_MASTER["agent_name"].ToString();
        //        //Lib.WriteData(WS, iRow, 1, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        //Lib.WriteData(WS, iRow, 11, sell_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        //Lib.WriteData(WS, iRow, 12, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow++;
        //        iRow++;
        //        iRow++;

        //        Lib.WriteData(WS, iRow++, 1, "B.EXPENSE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        _Border = "TBLR";
        //        Str = DR_MASTER["curr_code"].ToString();
        //        Lib.WriteMergeCell(WS, iRow, 1,2, 3, "WEIGHT", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        //Lib.WriteMergeCell(WS, iRow, 2, 1, 3, "", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 3, 1, 3, "FRT/KG", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 4, 1, 3, "FRT", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 5, 1, 3, "FSC", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 6, 1, 3, "SRC", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 7, 1, 3, "WRS", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 8, 1, 3, "MCC", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 9, 1, 3, "OTHERS", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 10, 1, 3, "TOTAL PP (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 11, 1, 3, "TOTAL CC (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        Lib.WriteMergeCell(WS, iRow, 12, 1, 3, "TOTAL (" + Str + ")", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //      //  Lib.WriteMergeCell(WS, iRow, 13, 1, 3, "TOTAL (INR)", "Calibri", _Size, true, _Color, "C", "T", _Border, "THIN", true, false);
        //        iRow++; iRow++;
        //        bok = false;  
        //        foreach (DataRow dr in dt_costdet.Select("costd_type = 'BUY'", "costd_ctr"))
        //        {
        //            bok = true;
        //            iRow++;
        //            Lib.WriteMergeCell(WS, iRow, 1, 2, 1, dr["costd_chwt"].ToString(), "Calibri", _Size, false, _Color, "L", "T", _Border, "THIN", true, false);
        //            nAmt = GetConvertAmt(DR_MASTER["COST_INFORM_RATE"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 3, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_frt_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_frt_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 4, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_myc_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_myc_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 5, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_src_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_src_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 6, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_wrs_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_wrs_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 7, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_mcc_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_mcc_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 8, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);

        //            if (dr["hbl_terms"].ToString() == "FREIGHT PREPAID")
        //                nAmt = GetConvertAmt(dr["costd_oth_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            else
        //                nAmt = GetConvertAmt(dr["costd_oth_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 9, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            nAmt = GetConvertAmt(dr["costd_pp"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 10, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            nAmt = GetConvertAmt(dr["costd_cc"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 11, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            nAmt = GetConvertAmt(dr["costd_tot"].ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 12, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //          //  Lib.WriteData(WS, iRow, 13, dr["costd_tot"].ToString(), _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        }
        //        if (bok)
        //        {
        //            iRow++;
        //            Lib.WriteMergeCell(WS, iRow, 1, 2, 1, "TOTAL", "Calibri", _Size, true, _Color, "L", "T", _Border, "THIN", true, false);
        //            //Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            Str = "";
        //           // Lib.WriteData(WS, iRow, 2, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 3, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 4, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 5, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 6, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 7, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 8, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 9, Str, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
        //            buy_pp = GetConvertAmt(buy_pp.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 10, buy_pp, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            buy_cc = GetConvertAmt(buy_cc.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 11, buy_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //           // nAmt = buy_tot;
        //            buy_tot = GetConvertAmt(buy_tot.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 12, buy_tot, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //           // Lib.WriteData(WS, iRow, 13, nAmt, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        }

        //        iRow++;

        //        expense_pp += buy_pp;
        //        expense_cc += buy_cc;
        //        expense += buy_tot;
        //        _Border = "";  
        //        if (rebate > 0)
        //        {
        //            Lib.WriteData(WS, iRow, 1, "REBATE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            rebate = GetConvertAmt(rebate.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 10, rebate, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);

        //            iRow++;
        //            expense_pp += rebate;
        //            expense += rebate;
        //        }

        //        if (exwork > 0)
        //        {

        //            Lib.WriteData(WS, iRow, 1, "Ex-WORK", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            exwork = GetConvertAmt(exwork.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 10, exwork, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            iRow++;
        //        }

        //        if (other > 0)
        //        {

        //            Lib.WriteData(WS, iRow, 1, "OTHER CHARGES", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //            other = GetConvertAmt(other.ToString(), DR_MASTER["cost_exrate"].ToString());
        //            Lib.WriteData(WS, iRow, 10, other, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //            iRow++;
        //            expense_pp += other;
        //            expense += other;
        //        }

        //        iRow++;
        //        iRow++;

        //        Lib.WriteData(WS, iRow, 1, "C.NET PROFIT/ LOSS(+ / -) A - B", _Color, false, _Border, "L", "", _Size, false, 325, "", true);

        //        Lib.WriteData(WS, iRow, 12, profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow++;
        //        iRow++;

        //        Lib.WriteData(WS, iRow, 1, "PROFIT / LOSS(+ / -) SHARE", _Color,false, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 10, our_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        Lib.WriteData(WS, iRow, 11, your_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow++;
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 10, buy_pp + our_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        Lib.WriteData(WS, iRow, 11, your_profit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow++;
        //        iRow++;
        //        Str = "FREIGHT AND OTHER CHARGES COLLECTED BY " + DR_MASTER["agent_name"].ToString();
        //        Lib.WriteData(WS, iRow, 1, Str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, 11, sell_cc, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        iRow++;

        //        nDrCRAmt = Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString());

        //        _Size++;

        //        iRow++;
        //        iRow++;

        //        if (nDrCRAmt > 0)
        //        {
        //            Lib.WriteData(WS, iRow, 1, "NET DUE FROM " + DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 10, Math.Abs(nDrCRAmt), _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        }
        //        else
        //        {
        //            Lib.WriteData(WS, iRow, 1, "NET DUE TO " + DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, 11, Math.Abs(nDrCRAmt), _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
        //        }
        //        iRow += 6;
        //    }


        //    if (nDrCRAmt < 0)
        //        nDrCRAmt = Math.Abs(nDrCRAmt);

        //    string sAmt = Lib.NumericFormat(nDrCRAmt.ToString(), 2);

        //    string sWords = "";
        //    if (DR_MASTER["curr_code"].ToString() != "INR")
        //        sWords = Number2Word_USD.Convert(sAmt, DR_MASTER["CURR_CODE"].ToString(), "CENTS");
        //    if (DR_MASTER["curr_code"].ToString() == "INR")
        //        sWords = Number2Word_RS.Convert(sAmt, "INR", "PAISE");


        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, sWords, "Calibri", 14, true, Color.Black, "L", "C", "TB", "THIN");
        //    Lib.WriteMergeCell(WS, iRow++, 1, 12, 1, "E.&.O.E", "Calibri", 12, true, Color.Black, "L", "C", "TB", "THIN");

        //    if (DR_MASTER["cost_format"].ToString() == "HANDLING" || DR_MASTER["cost_format"].ToString() == "PP") 
        //    {
        //        WS.Columns[12].Delete();
        //        WS.Columns[11].Delete();
        //        WS.Columns[10].Delete();
        //    }

        //    WB.SaveXls(File_Name + ".xls");
        //} 

        private void ProcessDetailFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 0;

            decimal nDrCRAmt = 0;
            bool IsRemarksExist = false;

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
            WS.Columns[2].Width = 256 * 14;
            WS.Columns[3].Width = 256 * 14;
            WS.Columns[4].Width = 256 * 14;
            WS.Columns[5].Width = 256 * 14;
            WS.Columns[6].Width = 256 * 14;
            WS.Columns[7].Width = 256 * 14;


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
            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, comp_name, "Calibri", 14, true, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, comp_add1, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, comp_add2, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, comp_add3, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, comp_add4, "Calibri", 12, false, Color.Black, "C", "C", "", "");

            DateTime Dt;

            string sDate = ((DateTime)DR_MASTER["cost_date"]).ToString("dd/MM/yyyy");
            string sCntr = "";
            string Str = "";

            iRow++;

            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "DEBIT NOTE";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "CREDIT NOTE";

            Lib.WriteMergeCell(WS, iRow++, 1, 7, 2, sTitle, "Calibri", 18, true, Color.Black, "C", "C", "TB", "THIN");

            iRow += 2;

            _Size = 12;

            Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_NAME"].ToString(), _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, 5, "NUMBER", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 6, DR_MASTER["COST_REFNO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            File_Display_Name = DR_MASTER["COST_REFNO"].ToString() + "-" + DR_MASTER["COST_FOLDERNO"].ToString() + ".xls";

            iRow++;

            Lib.WriteData(WS, iRow, 1, DR_MASTER["AGENT_LINE1"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, 5, "DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 6, sDate, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE2"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE3"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow++, 1, DR_MASTER["AGENT_LINE4"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                sTitle = "WE DEBIT YOUR ACCOUNT FOR THE FOLLOWING";
            if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) < 0)
                sTitle = "WE CREDIT YOUR ACCOUNT FOR THE FOLLOWING";

            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, sTitle, "Calibri", 12, true, Color.Black, "C", "C", "TB", "THIN");

            Lib.WriteData(WS, iRow, 1, "FEEDER VESSEL", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["VESSEL_NAME"].ToString() + " " + DR_MASTER["HBL_VESSEL_NO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            if (DR_MASTER["HBL_BL_NO"].ToString().Trim() != "")
            {
                iRow++;
                Lib.WriteData(WS, iRow, 1, "MBL NO", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, DR_MASTER["HBL_BL_NO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            }
            iRow++;
            Lib.WriteData(WS, iRow, 1, "CONTAINER", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

            int iCount = 0;
            int ipr = 0;
            foreach (DataRow Dr in dt_cntr.Rows)
            {
                if (sCntr != "")
                    sCntr += ",";
                sCntr += Dr["cntr_no"].ToString() + "[" + Dr["cntr_type"].ToString() + "]";
                iCount++;
            }


            ipr = iCount / 3;
            if (iCount % 3 > 0)
                ipr++;
            if (ipr == 0)
                ipr = 1;

            Lib.WriteMergeCell(WS, iRow, 3, 5, ipr, sCntr, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);


            iRow += ipr;
            Lib.WriteData(WS, iRow, 1, "HBLNO/CONSIGNEE", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

            iCount = 0;
            foreach (DataRow Dr in dt_house.Rows)
            {
                if (Str != "")
                    Str += ",";
                Str += Dr["hbl_bl_no"].ToString() + " / " + Dr["consignee_name"].ToString();
                iCount++;
            }
            if (iCount == 0)
                iCount = 1;

            Lib.WriteMergeCell(WS, iRow, 3, 5, iCount, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
            iRow += iCount;



            Lib.WriteData(WS, iRow, 1, "PORT OF LOADING", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["POL_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;

            Lib.WriteData(WS, iRow, 1, "DESTINATION", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["POD_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;

            Lib.WriteData(WS, iRow, 1, "OUR REFNO", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, DR_MASTER["COST_FOLDERNO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow++;

            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, "", "Calibri", 11, true, Color.Black, "C", "C", "T", "THIN");

            iCol = 1;
            _Color = Color.Black;
            _Border = "";
            _Size = 12;

            IsRemarksExist = false;
            foreach (DataRow Dr in dt_costDet.Rows)
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

            iRow++;
            foreach (DataRow Dr in dt_costDet.Rows)
            {
                //if (Lib.Conv2Decimal(DR_MASTER["cost_drcr_amount"].ToString()) > 0)
                //    Lib.WriteData(WS, iRow, 2, "OUR HANDLING CHARGES", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                //else
                //    Lib.WriteData(WS, iRow, 2, "YOUR HANDLING CHARGES", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                //Lib.WriteData(WS, iRow, 5, DR_MASTER["cost_hand_charges"], _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
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


            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, sWords, "Calibri", 11, true, Color.Black, "L", "C", "TB", "THIN");
            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, "E.&.O.E", "Calibri", 11, true, Color.Black, "L", "C", "TB", "THIN");
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
        private void WriteData(int nRow, int nCol, int Wd, string sData)
        {
            try
            {
                ws.Cells[nRow, nCol].Value = sData;
                if (Wd > 0)
                {
                    ws.Cells.GetSubrangeRelative(nRow, nCol, Wd, 1).SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        private void WriteData(int R, int C, object sData)
        {
            try
            {
                ws.Cells[R, C].Value = sData;
            }
            catch (Exception)
            {
            }
        }
        private void WriteData(int R, int C, string sAlign, object sData)
        {
            try
            {
                ws.Cells[R, C].Value = sData;
                if (sAlign == "RIGHT")
                    ws.Cells[R, C].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
                if (sAlign == "LEFT")
                    ws.Cells[R, C].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                if (sAlign == "CENTER")
                    ws.Cells[R, C].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                // ws.Cells[R, C].SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
            }
            catch (Exception)
            {
            }
        }
        private void WriteData(int R, int C, string sAlign, Boolean bBorders, Boolean bBold, object sData)
        {
            try
            {
                ws.Cells[R, C].Value = sData;
                if (sAlign == "RIGHT")
                    ws.Cells[R, C].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
                if (bBorders)
                    ws.Cells[R, C].SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
                if (bBold)
                    ws.Cells[R, C].Style.Font.Weight = ExcelFont.BoldWeight; //XL.ExcelFont.BoldWeight;
            }
            catch (Exception)
            {
            }
        }
        private void Merge_Cell(int _Row, int _Col, int _Width, int _Height)
        {
            myCell = ws.Cells.GetSubrangeRelative(_Row, _Col, _Width, _Height);
            myCell.Merged = true;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = "Arial";
            myCell.Style.Font.Size = 8 * 20;
        }
        private void WriteData(string sCaption, string sData)
        {
            try
            {
                ws.NamedRanges[sCaption].Range.Value = sData;
            }
            catch (Exception)
            {
            }
        }
        private void WriteData(string sCaption, Object sData)
        {
            try
            {
                ws.NamedRanges[sCaption].Range.Value = sData;
            }
            catch (Exception)
            {
            }
        }
        private void WriteData(string sCaption, decimal sData)
        {
            try
            {
                if (sData != 0)
                    ws.NamedRanges[sCaption].Range.Value = sData;
            }
            catch (Exception)
            {
            }
        }
    }
}

