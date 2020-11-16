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
    public class AgentInvoiceService : BL_Base
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

        private DataTable dt_master;
        private DataTable dt_house;
        private DataTable dt_cntr;
        private DataTable dt_costDet;

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
                sWhere += " and a.cost_source = 'AGENT INVOICE' ";
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
                sql += " select cost_pkid,cost_cfno,cost_refno,cost_date,cost_folderno,cost_sob_date,";
                sql += " agent.cust_name as agent_name,jvagent.cust_name as jvagent_name,cost_type, ";
                sql += " cost_drcr,cost_drcr_amount,cost_category,cost_remarks ,";
                sql += " cost_jv_ho_vrno,cost_jv_br_vrno,cost_jv_br_invno, cost_jv_posted,cost_checked_on,cost_sent_on,";
                sql += " curr.param_code as cost_currency_code,cost_exrate,nvl(cost_ddp,'N') as cost_ddp,";
                sql += " row_number() over(order by cost_date,cost_cfno) rn ";
                sql += " from costingm a ";
                sql += " left join customerm agent on a.cost_agent_id = agent.cust_pkid ";
                sql += " left join customerm jvagent on a.cost_jv_agent_id = jvagent.cust_pkid ";
                sql += " left join param curr on a.cost_currency_id = curr.param_pkid";
                sql += sWhere;
                sql += ") a ";

                if (type != "EXCEL")
                    sql += " where rn between {startrow} and {endrow}";

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
                    mRow.cost_ddp = Dr["cost_ddp"].ToString() == "Y" ? true : false;
                    mList.Add(mRow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintList(mList, branch_code);
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
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", report_type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);

            return RetData;
        }
        private void PrintList(List<Costingm> mList, string branch_code)
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";

            Color _Color = Color.Black;
            int _Size = 10;

            iRow = 0;
            iCol = 0;
            try
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
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                report_type = "EXCEL";
                File_Display_Name = "Report.xls";
                report_folderid = Guid.NewGuid().ToString().ToUpper();
                File_Name = Lib.GetFileName(report_folder, report_folderid, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 25;
                WS.Columns[2].Width = 256 * 12;
                WS.Columns[3].Width = 256 * 5;
                WS.Columns[4].Width = 256 * 35;
                WS.Columns[5].Width = 256 * 25;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 6;
                WS.Columns[8].Width = 256 * 6;
                WS.Columns[9].Width = 256 * 10;
                WS.Columns[10].Width = 256 * 12;
                WS.Columns[11].Width = 256 * 35;
                WS.Columns[12].Width = 256 * 10;
                WS.Columns[13].Width = 256 * 10;
                WS.Columns[14].Width = 256 * 10;
                WS.Columns[15].Width = 256 * 5;

                iRow = 0; iCol = 1;

                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", 12, false, 325, "", true);
                _Size = 10;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;

                Lib.WriteData(WS, iRow, 1, str, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, true, "", "L", "", _Size, false, 325, "", true);

                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "AGENT INVOICE", _Color, true, "", "L", "", 12, false, 325, "", true);

                iRow++;
                iRow++;

                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "REF#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DRCR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CURR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EX-RATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HO-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HO-VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BR-VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POSTED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DDP", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                foreach (Costingm Rec in mList)
                {
                    iRow++;
                    iCol = 1;

                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_refno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_date, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_type, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_category, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_folderno, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_currency_code, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_drcr, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_exrate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_drcr_amount, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_jv_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_jv_ho_vrno, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_jv_br_vrno, _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, (Rec.cost_jv_posted ? 'Y' : 'N'), _Color, false, "", "L", "", _Size, false, 325, "", false);
                    Lib.WriteData(WS, iRow, iCol++, (Rec.cost_ddp ? 'Y' : 'N'), _Color, false, "", "L", "", _Size, false, 325, "", false);
                }
                iRow++;
                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
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
                sql += " ,cost_ddp ";
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
                    mRow.cost_ddp = Dr["cost_ddp"].ToString() == "Y" ? true : false;
                    mRow.cost_jv_br_inv_id = Dr["cost_jv_br_inv_id"].ToString();
                    break;
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

                // This is disabled to allow multiple folder no
                /*
                if (Record.cost_folderno.Trim().Length > 0)
                {
                    sql = "select cost_pkid from (";
                    sql += "select cost_pkid  from costingm a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and a.cost_folderno = '{FOLDERNO}' ";
                    sql += " and a.cost_source ='AGENT INVOICE'";
                    sql += ") a where cost_pkid <> '{PKID}'";

                    sql = sql.Replace("{FOLDERNO}", Record.cost_folderno);
                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                    sql = sql.Replace("{PKID}", Record.cost_pkid);

                    if (Con_Oracle.IsRowExists(sql))
                    {
                        bError = true;
                        Lib.AddError(ref str, " | This No already Exists");
                    }
                }
                */


                if (!Lib.IsInFinYear(Record.cost_date, Record._globalvariables.year_start_date, Record._globalvariables.year_end_date, true))
                {
                    bError = true;
                    Lib.AddError(ref str," | Invalid Date (Future Date or Date not in Financial Year)");
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

            string doc_prefix = "";

            decimal nDrcrInr = 0;


            DataRow lovRow_Doc_Prefix ;
            DataRow lovRow_Doc_Air_Prefix ;

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

                    lov = new LovService();
                    if (Record.cost_type == "SEA")
                    {
                        Record.rec_category = "SEA EXPORT";
                        lovRow_Doc_Prefix = lov.getSettings(Record._globalvariables.branch_code, "COST-PREFIX");
                        if (lovRow_Doc_Prefix != null)
                            doc_prefix = lovRow_Doc_Prefix["name"].ToString();
                        else
                            throw new Exception("Prefix Not Found");
                    }
                    if (Record.cost_type == "AIR")
                    {
                        Record.rec_category = "AIR EXPORT";
                        lovRow_Doc_Air_Prefix = lov.getSettings(Record._globalvariables.branch_code, "COST-AIR-PREFIX");
                        if (lovRow_Doc_Air_Prefix != null)
                            doc_prefix = lovRow_Doc_Air_Prefix["name"].ToString();
                        else
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
                        sql += " and a.cost_source =  'AGENT INVOICE' ";

                       
                        sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                        sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);
                        sql = sql.Replace("{FYEAR}", Record._globalvariables.year_code);
                       

                        DataTable DT_MON = new DataTable();
                        DT_MON = Con_Oracle.ExecuteQuery(sql);
                        if (DT_MON.Rows.Count > 0)
                        {
                            Record.cost_cfno = Lib.Conv2Integer(DT_MON.Rows[0]["monno"].ToString());
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
                Rec.InsertString("cost_refno", Record.cost_refno);
                Rec.InsertString("cost_ddp", Record.cost_ddp == true ? "Y" : "N");

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("cost_cfno", Record.cost_cfno.ToString());
                    
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

        public Dictionary<string, object> DeleteRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            DataTable Dt_Test = new DataTable();

            try
            {
                string id = SearchData["pkid"].ToString();
                string ErrorMessage = "";
                Con_Oracle = new DBConnection();

                sql = "select cost_pkid, to_char( cost_date, 'DD-MON-YYYY') as cost_date, nvl(cost_jv_posted,'N') = 'Y', cost_jv_ho_id, cost_jv_br_id from costingm where cost_pkid ='" + id + "'";
                Dt_Test = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Test.Rows.Count <= 0)
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    ErrorMessage = " Record not exists ";
                    throw new Exception(ErrorMessage);
                }

                if (Dt_Test.Rows[0]["cost_jv_posted"].ToString() == "Y")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    ErrorMessage = " Need to Release ";
                    throw new Exception(ErrorMessage);
                }
                string hoid = Dt_Test.Rows[0]["cost_jv_ho_id"].ToString();
                string brid = Dt_Test.Rows[0]["cost_jv_br_id"].ToString();

                if (ErrorMessage == "")
                {
                    Con_Oracle.BeginTransaction();

                    sql  = " delete from costcentert where ct_jvh_id = '" + brid + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " delete from ledgert where jv_parent_id = '"+ brid + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " delete from ledgerh where jvh_pkid = '" + brid + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    if (hoid != "")
                    {
                        sql = " delete from ledgert where jv_parent_id = '" + hoid + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                        sql = " delete from ledgerh where jvh_pkid = '" + hoid + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                    sql = " Delete from costingd where costd_parent_id ='" + id + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " Delete from costingm where cost_pkid ='" + id + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                }

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
            return RetData;
        }

        public Dictionary<string, object> PrintNote(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            try
            {
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
                sql += " select  costd_acc_name ,costd_acc_amt ";
                sql += " from costingd ";
                sql += " where costd_parent_id ='" + report_pkid + "'";
                sql += " order by costd_ctr";
                dt_costDet = Con_Oracle.ExecuteQuery(sql);



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
            RetData.Add("filedisplayname", File_Display_Name.Replace("\\","") );
            return RetData;

        }


        private void ProcessExcelFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 0;

            decimal nDrCRAmt = 0;

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

            if (DR_MASTER["COST_DRCR"].ToString()=="DR")
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
            iRow += 3;

            Lib.WriteData(WS, iRow, 2, "PARTICULARS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 5, "AMOUNT(" + DR_MASTER["curr_code"].ToString() + ")", _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            foreach (DataRow dr in dt_costDet.Rows)
            {
                iRow++;
                Lib.WriteData(WS, iRow, 2, dr["costd_acc_name"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 5, dr["costd_acc_amt"], _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            }
            iRow++;

            Lib.WriteData(WS, iRow, 2, "TOTAL", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 5, DR_MASTER["cost_drcr_amount"], _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00", true);
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

            WB.SaveXls(File_Name + ".xls");
        }



    }
}

