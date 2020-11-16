using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;
using System.Drawing;

namespace BLOperations
{
    public class WeekPlanningService : BL_Base
    {
        int iRow = 0;
        int iCol = 0;
        string type = "";
        ExcelFile WB;
        ExcelWorksheet WS = null;
        string report_folder = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string PKID = "";
        string company_code = "";
        string branch_code = "";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<AgentBookingm> mList = new List<AgentBookingm>();
            AgentBookingm mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
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

                sWhere = " where a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                sWhere += " and a.ab_type = 'PLANNING' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  a.ab_book_no  like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(shpr.cust_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(cnge.cust_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(agent.cust_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }
                /*
                if (from_date != "NULL")
                    sWhere += "  and ab_book_date >= '{FDATE}' ";
                if (to_date != "NULL")
                    sWhere += "  and ab_book_date <= '{EDATE}' ";
                    */

                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
                sWhere = sWhere.Replace("{FDATE}", from_date);
                sWhere = sWhere.Replace("{EDATE}", to_date);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM agentbookingm  a ";
                    sql += " left join customerm shpr on a.ab_exp_id = shpr.cust_pkid ";
                    sql += " left join customerm cnge on a.ab_imp_id = cnge.cust_pkid ";
                    sql += " left join customerm agent on a.ab_agent_id = agent.cust_pkid ";
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
                sql += " select ab_pkid,ab_book_no,ab_book_date";
                sql += " ,shpr.cust_name as ab_exp_name";
                sql += " ,cnge.cust_name as ab_imp_name";
                sql += " ,agent.cust_name as ab_agent_name ";
                sql += " ,ab_week_no,ab_week_status ";
                sql += " ,row_number() over(order by a.ab_book_no) rn ";
                sql += " from agentbookingm a ";
                sql += " left join customerm shpr on a.ab_exp_id = shpr.cust_pkid ";
                sql += " left join customerm cnge on a.ab_imp_id = cnge.cust_pkid ";
                sql += " left join customerm agent on a.ab_agent_id = agent.cust_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by ab_book_no";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new AgentBookingm();
                    mRow.ab_pkid = Dr["ab_pkid"].ToString();
                    mRow.ab_book_no = Lib.Conv2Integer(Dr["ab_book_no"].ToString());
                    mRow.ab_book_date = Lib.DatetoStringDisplayformat(Dr["ab_book_date"]);
                    mRow.ab_exp_name = Dr["ab_exp_name"].ToString();
                    mRow.ab_imp_name = Dr["ab_imp_name"].ToString();
                    mRow.ab_agent_name = Dr["ab_agent_name"].ToString();
                    mRow.ab_week_no = Lib.Conv2Integer(Dr["ab_week_no"].ToString());
                    mRow.ab_week_status = Dr["ab_week_status"].ToString();

                    mList.Add(mRow);
                }

                Dt_List.Rows.Clear();
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
            AgentBookingm mRow = new AgentBookingm();
            mRow = new AgentBookingm();
            string type = "";
            string id = "";


            if (SearchData.ContainsKey("pkid"))
                id = SearchData["pkid"].ToString();

            if (SearchData.ContainsKey("Type"))
                type = SearchData["Type"].ToString();

            report_folder = SearchData["report_folder"].ToString();

            if (SearchData.ContainsKey("file_pkid"))
                PKID = SearchData["file_pkid"].ToString();


            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select ab_pkid,ab_book_no,ab_book_date";
                sql += " ,ab_exp_id,shpr.cust_code as ab_exp_code,shpr.cust_name as ab_exp_name";
                sql += " ,ab_imp_id,cnge.cust_code as ab_imp_code,cnge.cust_name as ab_imp_name";
                sql += " ,ab_agent_id,agent.cust_code as ab_agent_code, agent.cust_name as ab_agent_name ";
                sql += " ,ab_approved,ab_remarks,ab_week_no,ab_week_status ";
                sql += " from agentbookingm a ";
                sql += " left join customerm shpr on a.ab_exp_id = shpr.cust_pkid ";
                sql += " left join customerm cnge on a.ab_imp_id = cnge.cust_pkid ";
                sql += " left join customerm agent on a.ab_agent_id = agent.cust_pkid ";
                sql += " where  a.ab_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new AgentBookingm();
                    mRow.ab_pkid = Dr["ab_pkid"].ToString();
                    mRow.ab_book_no = Lib.Conv2Integer(Dr["ab_book_no"].ToString());
                    mRow.ab_book_date = Lib.DatetoString(Dr["ab_book_date"]);
                    mRow.ab_exp_id = Dr["ab_exp_id"].ToString();
                    mRow.ab_exp_code = Dr["ab_exp_code"].ToString();
                    mRow.ab_exp_name = Dr["ab_exp_name"].ToString();
                    mRow.ab_imp_id = Dr["ab_imp_id"].ToString();
                    mRow.ab_imp_code = Dr["ab_imp_code"].ToString();
                    mRow.ab_imp_name = Dr["ab_imp_name"].ToString();
                    mRow.ab_agent_id = Dr["ab_agent_id"].ToString();
                    mRow.ab_agent_code = Dr["ab_agent_code"].ToString();
                    mRow.ab_agent_name = Dr["ab_agent_name"].ToString();
                    mRow.ab_week_no = Lib.Conv2Integer(Dr["ab_week_no"].ToString());
                    mRow.ab_week_status = Dr["ab_week_status"].ToString();
                    mRow.ab_book_displaydate = Lib.DatetoStringDisplayformat(Dr["ab_book_date"]);
                    mRow.ab_remarks = Dr["ab_remarks"].ToString();
                    break;
                }

                List<Joborderm> mList = new List<Joborderm>();
                mList = new List<Joborderm>();
                Joborderm jRow;

                string sWhere = "";
                sWhere += " where ord_week_id = '{ID}'";
                sWhere = sWhere.Replace("{ID}", id);

                sql = GetOrderListSQL(sWhere);
                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    jRow = new Joborderm();
                    jRow.ord_pkid = Dr["ord_pkid"].ToString();
                    jRow.ord_po = Dr["ord_po"].ToString();
                    jRow.ord_style = Dr["ord_style"].ToString();
                    jRow.ord_desc = Dr["ord_desc"].ToString();
                    jRow.ord_color = Dr["ord_color"].ToString();
                    jRow.ord_exp_name = Dr["ord_exp_name"].ToString();
                    jRow.ord_imp_name = Dr["ord_imp_name"].ToString();
                    jRow.ord_stylename = Dr["ord_stylename"].ToString();
                    jRow.ord_uneco = Dr["ord_uneco"].ToString();
                    jRow.ord_approved = Dr["ord_approved"].ToString() == "Y" ? true : false;
                    jRow.ord_booking_no = Dr["ord_booking_no"].ToString();
                    jRow.ord_booking_date = Lib.DatetoStringDisplayformat(Dr["ord_booking_date"]);
                    jRow.ord_rnd_insp_date = Lib.DatetoStringDisplayformat(Dr["ord_rnd_insp_date"]);
                    jRow.ord_po_rel_date = Lib.DatetoStringDisplayformat(Dr["ord_po_rel_date"]);
                    jRow.ord_cargo_ready_date = Lib.DatetoStringDisplayformat(Dr["ord_cargo_ready_date"]);
                    jRow.ord_fcr_date = Lib.DatetoStringDisplayformat(Dr["ord_fcr_date"]);
                    jRow.ord_insp_date = Lib.DatetoStringDisplayformat(Dr["ord_insp_date"]);
                    jRow.ord_stuf_date = Lib.DatetoStringDisplayformat(Dr["ord_stuf_date"]);
                    jRow.ord_whd_date = Lib.DatetoStringDisplayformat(Dr["ord_whd_date"]);
                    jRow.ord_agent_name = Dr["ord_agent_name"].ToString();
                    jRow.ord_week_no = Lib.Conv2Integer(Dr["ord_week_no"].ToString());
                    jRow.ord_ourbooking_no = Dr["ord_book_no"].ToString();
                    jRow.ord_selected = true;
                    mList.Add(jRow);
                }
                mRow.OrderList = mList;



                if (type == "EXCEL")
                {
                    if (mList != null)
                    {
                        PrintAgentBookReport(mRow);

                    }
                }
                Dt_Rec.Rows.Clear();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            RetData.Add("type", type);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }

        public string AllValid(AgentBookingm Record)
        {
            string str = "";

            return str;
        }

        public Dictionary<string, object> Save(AgentBookingm Record)
        {
            string sql = "";
            string DocNo = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            DataTable Dt_Temp;
            try
            {
                Con_Oracle = new DBConnection();

                if ((ErrorMessage = AllValid(Record)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (Record.rec_mode == "ADD")
                {
                    sql = "select nvl(max(ab_book_no) + 1,1001) as bookno from agentbookingm a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                    sql += " and a.ab_type = 'PLANNING'";

                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);

                    Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        DocNo = Dt_Temp.Rows[0]["bookno"].ToString();
                        Record.ab_book_no = Lib.Conv2Integer(Dt_Temp.Rows[0]["bookno"].ToString());
                    }
                    else
                    {
                        ErrorMessage = "Booking Number Not Found Try again";

                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception(ErrorMessage);
                    }
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("agentbookingm", Record.rec_mode, "ab_pkid", Record.ab_pkid);

                Rec.InsertDate("ab_book_date", Record.ab_book_date);
                Rec.InsertString("ab_agent_id ", Record.ab_agent_id);
                Rec.InsertString("ab_exp_id ", Record.ab_exp_id);
                Rec.InsertString("ab_imp_id ", Record.ab_imp_id);
                Rec.InsertString("ab_remarks", Record.ab_remarks);
                Rec.InsertString("rec_handled_by", Record._globalvariables.user_code);
                Rec.InsertNumeric("ab_week_no", Record.ab_week_no.ToString());
                Rec.InsertString("ab_week_status", Record.ab_week_status);
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("ab_type", "PLANNING");
                    Rec.InsertNumeric("ab_book_no", Record.ab_book_no.ToString());
                    Rec.InsertString("rec_deleted", "N");
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

                if (Record.rec_mode == "EDIT")
                {
                    sql = "update joborderm set ord_week_id = null, ord_week_no = null where ord_week_id ='" + Record.ab_pkid + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                }

                foreach (Joborderm Row in Record.OrderList)
                {
                    if (Row.ord_selected == true)
                    {
                        sql = "update joborderm set ord_week_id ='" + Record.ab_pkid + "'";
                        sql += " ,ord_week_no = " + Record.ab_week_no;
                        sql += "  where ord_pkid ='" + Row.ord_pkid + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                }
                Con_Oracle.CommitTransaction();
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
            RetData.Add("bookno", DocNo);
            return RetData;
        }



        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            //Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "SALES EXECUTIVE");
            //RetData.Add("smanlist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "CITY");
            //RetData.Add("citylist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "STATE");
            //RetData.Add("statelist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;
        }


        public IDictionary<string, object> OrderList(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Joborderm> mList = new List<Joborderm>();
            Joborderm mRow;
            string ordpo = "";

            string rowtype = SearchData["rowtype"].ToString();
            string weekid = SearchData["bookid"].ToString();
            string agentid = SearchData["agentid"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            if (SearchData.ContainsKey("ordpo"))
            {
                ordpo = SearchData["ordpo"].ToString();
                ordpo = ordpo.Replace(" ", "");
                ordpo = ordpo.Replace(",", "','");
            }

            try
            {

                sWhere = " where  a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                sWhere += " and nvl(a.ord_approved,'N') = 'Y' ";
                sWhere += " and ord_agent_id = '{AGENT_ID}' ";
                sWhere += " and ((ord_week_id is null) or ord_week_id = '{WEEK_ID}')";

                if (ordpo.Length > 0)
                    sWhere += " and ord_po in ('{ORD_PO}') ";

                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
                sWhere = sWhere.Replace("{ORD_PO}", ordpo);
                sWhere = sWhere.Replace("{AGENT_ID}", agentid);
                sWhere = sWhere.Replace("{WEEK_ID}", weekid);


                DataTable Dt_List = new DataTable();

                sql = GetOrderListSQL(sWhere);
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Joborderm();
                    mRow.ord_pkid = Dr["ord_pkid"].ToString();
                    mRow.ord_po = Dr["ord_po"].ToString();
                    mRow.ord_style = Dr["ord_style"].ToString();
                    mRow.ord_desc = Dr["ord_desc"].ToString();
                    mRow.ord_color = Dr["ord_color"].ToString();
                    mRow.ord_exp_name = Dr["ord_exp_name"].ToString();
                    mRow.ord_imp_name = Dr["ord_imp_name"].ToString();
                    mRow.ord_stylename = Dr["ord_stylename"].ToString();
                    mRow.ord_uneco = Dr["ord_uneco"].ToString();
                    mRow.ord_selected = false;
                    if (Dr["ord_week_id"].ToString() == weekid)
                        mRow.ord_selected = true;

                    mRow.ord_approved = Dr["ord_approved"].ToString() == "Y" ? true : false;
                    mRow.ord_booking_no = Dr["ord_booking_no"].ToString();
                    mRow.ord_booking_date = Lib.DatetoStringDisplayformat(Dr["ord_booking_date"]);
                    mRow.ord_rnd_insp_date = Lib.DatetoStringDisplayformat(Dr["ord_rnd_insp_date"]);
                    mRow.ord_po_rel_date = Lib.DatetoStringDisplayformat(Dr["ord_po_rel_date"]);
                    mRow.ord_cargo_ready_date = Lib.DatetoStringDisplayformat(Dr["ord_cargo_ready_date"]);
                    mRow.ord_fcr_date = Lib.DatetoStringDisplayformat(Dr["ord_fcr_date"]);
                    mRow.ord_insp_date = Lib.DatetoStringDisplayformat(Dr["ord_insp_date"]);
                    mRow.ord_stuf_date = Lib.DatetoStringDisplayformat(Dr["ord_stuf_date"]);
                    mRow.ord_whd_date = Lib.DatetoStringDisplayformat(Dr["ord_whd_date"]);
                    mRow.ord_week_no = Lib.Conv2Integer(Dr["ord_week_no"].ToString());
                    mRow.ord_ourbooking_no = Dr["ord_book_no"].ToString();
                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("list", mList);
            return RetData;
        }

        private string GetOrderListSQL(string sWhere)
        {

            sql = " select a.ord_pkid,a.ord_week_id,a.ord_po,a.ord_style";
            sql += " ,a.ord_desc,a.ord_color,a.ord_stylename ";
            sql += " ,exp.cust_name as ord_exp_name";
            sql += " ,imp.cust_name as ord_imp_name";
            sql += " ,a.rec_created_date,a.ord_uneco,a.ord_approved,a.ord_booking_no ";
            sql += " ,ord_booking_date,ord_rnd_insp_date,ord_po_rel_date,ord_cargo_ready_date";
            sql += " ,ord_fcr_date, ord_insp_date, ord_stuf_date, ord_whd_date,agent.cust_name as ord_agent_name ";
            sql += " ,ord_week_no,agentbk.ab_book_no as ord_book_no  ";
            sql += " from joborderm a ";
            sql += " left join customerm exp on a.ord_exp_id = exp.cust_pkid  ";
            sql += " left join customerm imp on a.ord_imp_id = imp.cust_pkid  ";
            sql += " left join customerm agent on a.ord_agent_id = agent.cust_pkid  ";
            sql += " left join agentbookingm agentbk on a.ord_booking_id = agentbk.ab_pkid ";
            sql += sWhere;
            sql += " order by case when ord_week_id is null then 'B' else 'A' end,a.rec_created_date";
            return sql;
        }


        private void PrintAgentBookReport(AgentBookingm mRow)
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";
            string booking_date = "";
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            iRow = 0;
            iCol = 0;
            int i = 0;
            try
            {
                // REPORT_CAPTION = searchtype;

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

                File_Display_Name = "WeeklyReport.xls";
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
                WS.Columns[5].Width = 256 * 23;
                WS.Columns[6].Width = 256 * 23;
                WS.Columns[7].Width = 256 * 26;
                WS.Columns[8].Width = 256 * 10;
                WS.Columns[9].Width = 256 * 32;
                WS.Columns[10].Width = 256 * 15;



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
                Lib.WriteData(WS, iRow, 1, "WEEKLY PLANNING ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;


                Lib.WriteData(WS, iRow, 1, "PLANNING#", _Color, true, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 2, mRow.ab_book_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 3, "DATE", _Color, true, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 4, Lib.nvlDate(mRow.ab_book_displaydate, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, "WEEK#", _Color, true, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 2, mRow.ab_week_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, "AGENT", _Color, true, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 2, mRow.ab_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, "SHIPPER", _Color, true, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 2, mRow.ab_exp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, "CONSIGNEE", _Color, true, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, 2, mRow.ab_imp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);

                iRow++;
                iRow++;

                Lib.WriteData(WS, iRow, iCol++, "PO#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STYLE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COLOR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "UNECO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DESCRIPTION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT/BOOKING#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                string ord_agent_booking_no = "";
                foreach (Joborderm Rec in mRow.OrderList)
                {
                    iRow++;
                    iCol = 1;
                    i++;
                    ord_agent_booking_no = "";
                    if (Rec.ord_agent_name != null)
                        ord_agent_booking_no = Rec.ord_agent_name;
                    if (Rec.ord_booking_no != null && ord_agent_booking_no != "")
                        ord_agent_booking_no += " / ";
                    ord_agent_booking_no += Rec.ord_booking_no;


                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_po, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_style, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_color, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_uneco, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_exp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_imp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_desc, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, ord_agent_booking_no, _Color, false, "", "L", "", _Size, false, 325, "", true);


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
