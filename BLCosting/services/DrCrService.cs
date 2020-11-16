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
    public class DrCrService : BL_Base
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
        private DataTable dt_master;
        private DataTable dt_house;
        private DataTable dt_cntr;
        private DataTable dt_costDet;
        private DataTable dt_bank=new DataTable();

        private string Bank_Company = "";
        private string Bank_Acno = "";
        private string Bank_Name = "";
        private string Bank_Ifsc_Code = "";
        private string Bank_Add1 = "";
        private string Bank_Add2 = "";
        private string Bank_Add3 = "";

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
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string year_code = SearchData["year_code"].ToString();
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
                sWhere += " and a.cost_source = 'DRCR ISSUE' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += " upper(a.cost_folderno) like '%" + searchstring.ToUpper() + "%'";
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
                sql += " select cost_pkid,cost_cfno,cost_refno,cost_date,cost_folderno,cost_sob_date, ";
                sql += " agent.cust_name as agent_name,jvagent.cust_name as jvagent_name,cost_type, ";
                sql += " cost_drcr,cost_drcr_amount,cost_category,cost_remarks , cost_cntr,";
                sql += " cost_jv_ho_vrno,cost_jv_br_vrno,cost_jv_br_invno, cost_jv_posted,cost_checked_on,cost_sent_on,";
                sql += " curr.param_code as cost_currency_code,cost_exrate,";
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
                    mRow.cost_sob_date = Lib.DatetoStringDisplayformat(Dr["cost_sob_date"]);
                    mRow.cost_drcr = Dr["cost_drcr"].ToString();
                    mRow.cost_drcr_amount = Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString());
                    mRow.cost_category = Dr["cost_category"].ToString();
                    mRow.cost_remarks = Dr["cost_remarks"].ToString();
                    mRow.cost_type = Dr["cost_type"].ToString();

                    mRow.cost_jv_ho_vrno = Dr["cost_jv_ho_vrno"].ToString();
                    mRow.cost_jv_br_vrno = Dr["cost_jv_br_vrno"].ToString();
                    mRow.cost_jv_br_invno = Dr["cost_jv_br_invno"].ToString();
                    mRow.cost_jv_posted = false;
                    if (Dr["cost_jv_posted"].ToString() == "Y")
                        mRow.cost_jv_posted = true;
                    mRow.cost_checked_on = Lib.DatetoStringDisplayformat(Dr["cost_checked_on"]);
                    mRow.cost_sent_on = Lib.DatetoStringDisplayformat(Dr["cost_sent_on"]);
                    mRow.cost_currency_code = Dr["cost_currency_code"].ToString();
                    mRow.cost_exrate = Lib.Conv2Decimal(Dr["cost_exrate"].ToString());
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
        

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Costingm mRow = new Costingm();
            string id = SearchData["pkid"].ToString();
            bool bok = false;
            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select cost_pkid, cost_type, cost_source, cost_cfno,cost_refno,cost_date,cost_folderno,cost_mblid,mbl.hbl_book_cntr as cost_book_cntr";
                sql += " ,cost_agent_id,agnt.cust_code as cost_agent_code,agnt.cust_name as cost_agent_name,cost_agent_br_id";
                sql += " ,agntaddr.add_branch_slno as  cost_agent_br_no,agntaddr.add_line1||'\n'||agntaddr.add_line2||'\n'||agntaddr.add_line3 as  cost_agent_br_addr  ";
                sql += " ,cost_edit_code,cost_exrate,cost_currency_id,c.param_code as cost_currency_code,cost_year ";
                sql += " ,cost_remarks,cost_category,cost_drcr, cost_drcr_amount,nvl(cost_cntr,mbl.hbl_book_cntr) as cost_cntr ";
                sql += " ,cost_jv_agent_id,agnt2.cust_code as cost_jv_agent_code,agnt2.cust_name as cost_jv_agent_name,cost_jv_agent_br_id";
                sql += " ,agntaddr2.add_branch_slno as  cost_jv_agent_br_no,agntaddr2.add_line1||'\n'||agntaddr2.add_line2||'\n'||agntaddr2.add_line3 as  cost_jv_agent_br_addr,cost_jv_br_inv_id  ";
                sql += " from costingm a  ";
                sql += " left join hblm mbl on a.cost_mblid = mbl.hbl_pkid ";
                sql += " left join param c on a.cost_currency_id = c.param_pkid ";
                sql += " left join customerm agnt on a.cost_agent_id = agnt.cust_pkid ";
                sql += " left join addressm agntaddr on a.cost_agent_br_id = agntaddr.add_pkid ";
                sql += " left join customerm agnt2 on a.cost_jv_agent_id = agnt2.cust_pkid ";
                sql += " left join addressm agntaddr2 on a.cost_jv_agent_br_id = agntaddr2.add_pkid ";
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
                    mRow.cost_date = Lib.DatetoString(Dr["cost_date"]);
                    mRow.cost_folderno = Dr["cost_folderno"].ToString();
                    if (Dr["cost_category"].ToString() == "GENERAL JOB" && mRow.cost_folderno.ToString().Contains("-"))
                    {
                        string[] sdata = mRow.cost_folderno.ToString().Split('-');
                        mRow.cost_folderno = sdata[0];
                    }
                    mRow.cost_mblid = Dr["cost_mblid"].ToString();
                    mRow.cost_agent_id = Dr["cost_agent_id"].ToString();
                    mRow.cost_agent_code = Dr["cost_agent_code"].ToString();
                    mRow.cost_agent_name = Dr["cost_agent_name"].ToString();
                    mRow.cost_year = Lib.Conv2Integer(Dr["cost_year"].ToString());
                    mRow.cost_currency_id = Dr["cost_currency_id"].ToString();
                    mRow.cost_currency_code = Dr["cost_currency_code"].ToString();
                    mRow.cost_drcr = Dr["cost_drcr"].ToString();
                    mRow.cost_drcr_amount = Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString());
                    mRow.cost_category = Dr["cost_category"].ToString();
                    mRow.cost_remarks = Dr["cost_remarks"].ToString();
                    mRow.cost_edit_code = Dr["cost_edit_code"].ToString();
                    mRow.cost_book_cntr = Dr["cost_cntr"].ToString();
                    mRow.cost_exrate = Lib.Conv2Decimal(Dr["cost_exrate"].ToString());
                    mRow.cost_agent_br_id = Dr["cost_agent_br_id"].ToString();
                    mRow.cost_agent_br_no = Dr["cost_agent_br_no"].ToString();
                    mRow.cost_agent_br_addr= Dr["cost_agent_br_addr"].ToString();

                    mRow.cost_jv_agent_id = Dr["cost_jv_agent_id"].ToString();
                    mRow.cost_jv_agent_code = Dr["cost_jv_agent_code"].ToString();
                    mRow.cost_jv_agent_name = Dr["cost_jv_agent_name"].ToString();
                    mRow.cost_jv_agent_br_id = Dr["cost_jv_agent_br_id"].ToString();
                    mRow.cost_jv_agent_br_no = Dr["cost_jv_agent_br_no"].ToString();
                    mRow.cost_jv_agent_br_addr = Dr["cost_jv_agent_br_addr"].ToString();
                    mRow.cost_jv_br_inv_id = Dr["cost_jv_br_inv_id"].ToString();
                    break;
                }

                if (bok)
                {
                    List<Costingd> mList = new List<Costingd>();
                    Costingd bRow;

                    sql = "select costd_pkid,  costd_parent_id,  costd_acc_id,  costd_acc_name , ";
                    sql += " costd_blno,costd_acc_qty,costd_acc_rate,";
                    sql += " costd_acc_amt,  costd_ctr,costd_remarks,costd_brate,costd_srate,costd_split ";
                    sql += " from costingd a ";
                    sql += " where costd_parent_id ='{ID}' ";
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
                        bRow.costd_brate = Lib.Conv2Decimal(Dr["costd_brate"].ToString());
                        bRow.costd_srate = Lib.Conv2Decimal(Dr["costd_srate"].ToString());
                        bRow.costd_split = Lib.Conv2Decimal(Dr["costd_split"].ToString());
                        bRow.costd_acc_qty = Lib.Conv2Decimal(Dr["costd_acc_qty"].ToString());
                        bRow.costd_acc_rate = Lib.Conv2Decimal(Dr["costd_acc_rate"].ToString());
                        bRow.costd_acc_amt = Lib.Conv2Decimal(Dr["costd_acc_amt"].ToString());
                        bRow.costd_ctr = Lib.Conv2Integer(Dr["costd_ctr"].ToString());
                        bRow.costd_remarks = Dr["costd_remarks"].ToString();
                        mList.Add(bRow);
                    }
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
                //if (Record.cost_folderno.Trim().Length <= 0)
                //    Lib.AddError(ref str, " | Folder No cannot be blank");

                //if (Record.cost_folderno.Trim().Length > 0)
                //{
                //    sql = "select cost_pkid from (";
                //    sql += "select cost_pkid  from costingm a ";
                //    sql += " where a.rec_company_code = '{COMPCODE}'";
                //    sql += " and a.rec_branch_code = '{BRCODE}'";
                //    sql += " and a.cost_folderno = '{FOLDERNO}' ";
                //    sql += " and a.cost_source ='DRCR ISSUE'";
                //    sql += ") a where cost_pkid <> '{PKID}'";

                //    sql = sql.Replace("{FOLDERNO}", Record.cost_folderno);
                //    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                //    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                //    sql = sql.Replace("{PKID}", Record.cost_pkid);

                //    if (Con_Oracle.IsRowExists(sql))
                //    {
                //        bError = true;
                //        Lib.AddError(ref str, " | This No already Exists");
                //    }
                //}

                if (!Lib.IsInFinYear(Record.cost_date, Record._globalvariables.year_start_date, Record._globalvariables.year_end_date, true))
                {
                    bError = true;
                    Lib.AddError(ref str, " | Invalid Date (Future Date or Date not in Financial Year)");
                }

                if (Record.cost_category.Trim() != "OTHERS")
                {
                    sql = " select hbl_pkid from hblm a";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code =  '{BRCODE}'";
                    sql += " and a.hbl_year =  {YEARCODE}";
                    if (Record.cost_category.ToString() == "AIR EXPORT MAWBNO")
                    {
                        sql += " and a.hbl_type = 'MBL-AE' ";
                        sql += " and (a.hbl_bl_no  = '" + Record.cost_folderno + "')";
                    }
                    else if (Record.cost_category.ToString() == "AIR IMPORT MAWBNO")
                    {
                        sql += " and a.hbl_type = 'MBL-AI' ";
                        sql += " and (a.hbl_bl_no  = '" + Record.cost_folderno + "')";
                    }
                    else if (Record.cost_category.ToString() == "SEA EXPORT FOLDER NO")
                    {
                        sql += " and a.hbl_type = 'MBL-SE' ";
                        sql += " and (a.hbl_folder_no  = '" + Record.cost_folderno + "')";
                    }
                    else if (Record.cost_category.ToString() == "SEA IMPORT FOLDER NO")
                    {
                        sql += " and a.hbl_type = 'MBL-SI' ";
                        sql += " and (a.hbl_folder_no  = '" + Record.cost_folderno + "')";
                    }
                    else if (Record.cost_category.ToString() == "GENERAL JOB")
                    {
                        sql += " and a.hbl_type = 'JOB-GN' ";
                        sql += " and (a.hbl_no  = " + Record.cost_folderno + ")";
                    }
                    else
                        sql += " and 1 = 2 ";

                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                    sql = sql.Replace("{YEARCODE}", Record._globalvariables.year_code);
                    if (!Con_Oracle.IsRowExists(sql))
                    {
                        bError = true;
                        Lib.AddError(ref str, " Invalid data Please Find again.... ");
                    }
                }

                if (Record.cost_agent_br_id.Trim().Length > 0 || Record.cost_agent_id.Trim().Length > 0)
                {
                    sql = "select add_pkid from addressm where add_pkid = '{ADD_BRID}'";
                    sql += " and  add_parent_id = '{PARENT_ID}'";
                    sql = sql.Replace("{ADD_BRID}", Record.cost_agent_br_id);
                    sql = sql.Replace("{PARENT_ID}", Record.cost_agent_id);
                    if (!Con_Oracle.IsRowExists(sql))
                    {
                        bError = true;
                        Lib.AddError(ref str, " Invalid Agent Address ");
                    }
                }
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

            decimal nDrcrInr = 0;
            string doc_prefix = "";

            try
            {
                Con_Oracle = new DBConnection();

                if (Record.cost_folderno.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "No Cannot Be Empty");

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
                    string sCaption = "COST-PREFIX";
                    if (Record.cost_type == "SEA")
                    {
                        sCaption = "COST-PREFIX";
                        Record.rec_category = "SEA EXPORT";
                    }
                    if (Record.cost_type == "AIR")
                    {
                        sCaption = "COST-AIR-PREFIX";
                        Record.rec_category = "AIR EXPORT";
                    }

                    lov = new LovService();
                    DataRow lovRow_Doc_Prefix = lov.getSettings(Record._globalvariables.branch_code, sCaption);
                    if (lovRow_Doc_Prefix != null)
                        doc_prefix = lovRow_Doc_Prefix["name"].ToString();
                    else
                    {
                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception("Prefix Not Found");
                    }
                    

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
                        sql += " and a.cost_source in ('SEA EXPORT COSTING','AIR EXPORT COSTING','DRCR ISSUE','SE CONSOLE COSTING')";
                        sql += " and a.cost_type = '{COSTTYPE}' ";

                        sql = sql.Replace("{COSTMON}", JOB_MON);
                        sql = sql.Replace("{COSTMONYEAR}", JOB_MON_YEAR);
                        sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                        sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                        sql = sql.Replace("{FYEAR}", Record._globalvariables.year_code);
                        sql = sql.Replace("{COSTTYPE}", Record.cost_type);

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

                if (Record.cost_category.ToString() == "GENERAL JOB")
                {
                    sql = "select hbl_folder_no from hblm where hbl_folder_no is not null and hbl_pkid ='" + Record.cost_mblid + "'";
                    DataTable Dt_temp = new DataTable();
                    Dt_temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_temp.Rows.Count > 0)
                    {
                        Record.cost_folderno = Record.cost_folderno + "-" + Dt_temp.Rows[0]["hbl_folder_no"].ToString();
                    }
                    Dt_temp.Rows.Clear();
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("costingm", Record.rec_mode, "cost_pkid", Record.cost_pkid);
                Rec.InsertDate("cost_date", Record.cost_date);
                Rec.InsertString("cost_folderno", Record.cost_folderno);
                Rec.InsertString("cost_mblid", Record.cost_mblid);
                Rec.InsertString("cost_category", Record.cost_category);
                Rec.InsertString("cost_agent_id", Record.cost_agent_id);
                Rec.InsertString("cost_agent_br_id", Record.cost_agent_br_id);
                Rec.InsertString("cost_currency_id", Record.cost_currency_id);
                Rec.InsertNumeric("cost_exrate", Record.cost_exrate.ToString());
                Rec.InsertString("cost_drcr", Record.cost_drcr);
                nDrcrInr = Lib.Conv2Decimal(Record.cost_drcr_amount.ToString()) * Lib.Conv2Decimal(Record.cost_exrate.ToString());
                nDrcrInr = Lib.RoundNumber_Latest(nDrcrInr.ToString(), 2, true);

                Rec.InsertNumeric("cost_drcr_amount", Record.cost_drcr_amount.ToString());
                Rec.InsertNumeric("cost_drcr_amount_inr", nDrcrInr.ToString());
                Rec.InsertString("cost_remarks", Record.cost_remarks);
                Rec.InsertString("cost_cntr", Record.cost_book_cntr);
                Rec.InsertString("cost_jv_agent_id", Record.cost_jv_agent_id);
                Rec.InsertString("cost_jv_agent_br_id", Record.cost_jv_agent_br_id);
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("cost_cfno", Record.cost_cfno.ToString());
                    Rec.InsertString("cost_refno", Record.cost_refno);
                    Rec.InsertNumeric("cost_year", Record._globalvariables.year_code);
                    Rec.InsertString("cost_type", Record.cost_type);
                    Rec.InsertString("cost_source", Record.cost_source);
                    Rec.InsertString("rec_category", Record.rec_category);
                    Rec.InsertString("cost_prefix", doc_prefix);
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
                sql += " vsl.param_name as vessel_name, hbl_vessel_no,";
                sql += " pol.param_name as pol_name,";
                sql += " pod.param_name as pod_name,";
                sql += " curr.param_code as curr_code,cost_type,";
                sql += " cost_exrate,cost_buy_pp,cost_buy_cc,";
                sql += " cost_sell_pp,cost_sell_cc,cost_rebate,";
                sql += " cost_ex_works,cost_hand_charges,cost_kamai,";
                sql += " cost_other_charges,cost_asper_amount,cost_buy_tot,";
                sql += " cost_sell_tot,cost_profit ,cost_our_profit,cost_your_profit,";
                sql += " cost_drcr_amount,cost_drcr_amount_inr,cost_expense,cost_income,cost_drcr,cost_remarks,cost_category ";
                sql += " from costingm a";
                sql += " left join hblm b on a.cost_mblid = b.hbl_pkid";
                //sql += " left join customerm agent on a.cost_agent_id = agent.cust_pkid";
                //sql += " left join addressm agentadd on a.cost_agent_br_id = agentadd.add_pkid";
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
                sql += " select  h.hbl_bl_no, cons.cust_name as consignee_name, shpr.cust_name as shipper_name ";
                sql += " from costingm a";
                sql += " inner join hblm m on a.cost_mblid = m.hbl_pkid";
                sql += " inner join hblm h on m.hbl_pkid = h.hbl_mbl_id";
                sql += " left join customerm cons on h.hbl_imp_id = cons.cust_pkid";
                sql += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid";
                sql += " where cost_pkid ='" + report_pkid + "'";

                dt_house = Con_Oracle.ExecuteQuery(sql);

                if (DR_MASTER["cost_category"].ToString() == "SEA IMPORT FOLDER NO")
                {
                    sql = "";
                    sql += " select  cntr_no,ctype.param_code as cntr_type ";
                    sql += " from costingm a";
                    sql += " inner join impcontainerm c on cost_mblid = cntr_parent_id";
                    sql += " left join param ctype on cntr_type_id = ctype.param_pkid ";
                    sql += " where cost_pkid ='" + report_pkid + "'";
                }
                else
                {
                    sql = "";
                    sql += " select  cntr_no,ctype.param_code as cntr_type ";
                    sql += " from costingm a";
                    sql += " inner join containerm c on cost_mblid = cntr_booking_id";
                    sql += " left join param ctype on cntr_type_id = ctype.param_pkid ";
                    sql += " where cost_pkid ='" + report_pkid + "'";
                }
                dt_cntr = Con_Oracle.ExecuteQuery(sql);


                sql = "";
                sql += " select  costd_acc_name ,costd_acc_amt,costd_remarks ";
                sql += " from costingd ";
                sql += " where costd_parent_id ='" + report_pkid + "'";
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

                if (report_type == "EXCEL")
                {
                    if (Lib.CreateFolder(report_folder))
                        ProcessExcelFile();
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
            bool IsRemarksExist = false;

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
                    //comp_add4 = "Email : " + Dr["COMP_email"].ToString() + " Web : " + Dr["COMP_WEB"].ToString();
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

            if (DR_MASTER["COST_DRCR"].ToString() == "DR")
                sTitle = "DEBIT NOTE";
            if (DR_MASTER["COST_DRCR"].ToString() == "CR")
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


            if (DR_MASTER["COST_DRCR"].ToString() == "DR")
                sTitle = "WE DEBIT YOUR ACCOUNT FOR THE FOLLOWING";
            if (DR_MASTER["COST_DRCR"].ToString() == "CR")
                sTitle = "WE CREDIT YOUR ACCOUNT FOR THE FOLLOWING";

            Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, sTitle, "Calibri", 12, true, Color.Black, "C", "C", "TB", "THIN");


            if (DR_MASTER["COST_TYPE"].ToString() == "SEA")
            {
                Lib.WriteData(WS, iRow, 1, "FEEDER VESSEL", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, DR_MASTER["VESSEL_NAME"].ToString() + " " + DR_MASTER["HBL_VESSEL_NO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
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
            }
            else
            {
                Lib.WriteData(WS, iRow, 1, "MAWB", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, DR_MASTER["HBL_BL_NO"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, "MAWB DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, Lib.DatetoStringDisplayformat(DR_MASTER["HBL_DATE"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                iRow++;

                Lib.WriteData(WS, iRow, 1, "HAWB/SHIPPER/CONSIGNEE", _Color, _Bold, _Border, "LT", "", _Size, false, 325, "", true);

                int iCount = 0;
                Str = "";
                foreach (DataRow Dr in dt_house.Rows)
                {
                    if (Str != "")
                        Str += ",";
                    Str += Dr["hbl_bl_no"].ToString() + " / " + Dr["shipper_name"].ToString() + " / " + Dr["consignee_name"].ToString();
                    iCount++;
                }
                if (iCount == 0)
                    iCount = 1;

                Lib.WriteMergeCell(WS, iRow, 3, 5, iCount, Str, "Calibri", _Size, false, Color.Black, "L", "T", "", "", true);
                iRow += iCount;

                Lib.WriteData(WS, iRow, 1, "AIRPORT OF DEPARTURE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, DR_MASTER["POL_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                iRow++;

                Lib.WriteData(WS, iRow, 1, "AIRPORT OF DESTINATION", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, DR_MASTER["POD_NAME"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                iRow++;

                Lib.WriteMergeCell(WS, iRow++, 1, 7, 1, "", "Calibri", 11, true, Color.Black, "C", "C", "T", "THIN");
            }
            iCol = 1;
            _Color = Color.Black;
            _Border = "";
            _Size = 12;

            iRow += 2;
            Lib.WriteData(WS, iRow, 2, DR_MASTER["COST_REMARKS"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            iRow += 2;

            IsRemarksExist = false;
            foreach (DataRow Dr in dt_costDet.Rows)
            {
                if (Dr["costd_remarks"].ToString().Trim().Length > 0)
                {
                    IsRemarksExist = true;
                    break;
                }
            }
            Lib.WriteData(WS, iRow, 1, "PARTICULARS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            if (IsRemarksExist)
                Lib.WriteData(WS, iRow, 4, "REMARKS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 6, "AMOUNT(" + DR_MASTER["curr_code"].ToString() + ")", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            foreach (DataRow dr in dt_costDet.Rows)
            {
                iRow++;
                Lib.WriteData(WS, iRow, 1, dr["costd_acc_name"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                if (IsRemarksExist)
                    Lib.WriteData(WS, iRow, 4, dr["costd_remarks"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 6, dr["costd_acc_amt"], _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            }
            iRow++;
            iRow++;
            Lib.WriteData(WS, iRow, 1, "TOTAL", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 6, DR_MASTER["cost_drcr_amount"], _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
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



    }
}

