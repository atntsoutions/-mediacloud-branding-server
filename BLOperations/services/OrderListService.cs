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
    public class OrderListService : BL_Base
    {
        int iRow = 0;
        int iCol = 0;
        ExcelFile WB;
        ExcelWorksheet WS = null;
        string report_folder = "";
        string file_pkid = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Joborderm> mList = new List<Joborderm>();
            Joborderm mRow;


            string list_exp_id = "";
            string list_imp_id = "";
            string list_agent_id = "";
            string job_docno = "";
            string ord_po = "";
            string ord_invoice = "";
            string from_date = "";
            string to_date = "";
            string ord_showpending = "N";
            string ord_status = "REPORTED";
            string sort_colname = "a.rec_created_date desc";

            string sWhere = "";
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();

            string branch_code = "";
            //string branch_code = SearchData["branch_code"].ToString();

            string type = SearchData["type"].ToString();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;


            report_folder = SearchData["report_folder"].ToString();

            if (SearchData.ContainsKey("file_pkid"))
                file_pkid = SearchData["file_pkid"].ToString();


            if (SearchData.ContainsKey("list_exp_id"))
                list_exp_id = SearchData["list_exp_id"].ToString();

            if (SearchData.ContainsKey("list_imp_id"))
                list_imp_id = SearchData["list_imp_id"].ToString();

            if (SearchData.ContainsKey("list_agent_id"))
                list_agent_id = SearchData["list_agent_id"].ToString();

            if (SearchData.ContainsKey("job_docno"))
                job_docno = SearchData["job_docno"].ToString();

            if (SearchData.ContainsKey("ord_po"))
            {
                ord_po = SearchData["ord_po"].ToString();
                ord_po = ord_po.Replace(" ", "");
                ord_po = ord_po.Replace(",", "','");
            }

            if (SearchData.ContainsKey("ord_showpending"))
                ord_showpending = SearchData["ord_showpending"].ToString();

            if (SearchData.ContainsKey("ord_invoice"))
                ord_invoice = SearchData["ord_invoice"].ToString();

            if (SearchData.ContainsKey("from_date"))
                from_date = SearchData["from_date"].ToString();

            if (SearchData.ContainsKey("to_date"))
                to_date = SearchData["to_date"].ToString();

            if (SearchData.ContainsKey("ord_status"))
                ord_status = SearchData["ord_status"].ToString();

            if (SearchData.ContainsKey("sort_colname"))
                sort_colname = SearchData["sort_colname"].ToString();

            if (from_date.Length > 0)
            {
                from_date = Lib.StringToDate(from_date);
            }
            if (to_date.Length > 0)
            {
                to_date = Lib.StringToDate(to_date);
            }

            if (sort_colname == "")
                sort_colname = "a.rec_created_date desc";

            try
            {
                sWhere = "";
                sWhere = "where a.rec_company_code = '{COMPCODE}'  ";
                //  sWhere += " and a.ord_source = 'ORDER' ";


                if (list_exp_id.Length > 0)
                    sWhere += " and a.ord_exp_id = '" + list_exp_id + "' ";

                if (list_imp_id.Length > 0)
                    sWhere += " and a.ord_imp_id = '" + list_imp_id + "' ";

                if (list_agent_id.Length > 0)
                    sWhere += " and a.ord_agent_id = '" + list_agent_id + "' ";



                if (ord_po.Length > 0)
                    sWhere += " and a.ord_po in ('" + ord_po + "')";

                if (ord_invoice.Length > 0)
                    sWhere += " and a.ord_invno like '%" + ord_invoice + "%'";

                if (ord_status.Length > 0 && ord_status != "ALL")
                    sWhere += " and nvl(a.ord_status,'REPORTED') = '" + ord_status + "' ";

                if (from_date.Length > 0)
                {
                    sWhere += " and to_char(a.rec_created_date,'DD-MON-YYYY') >= to_date('{FDATE}','DD-MON-YYYY') ";
                }
                if (to_date.Length > 0)
                {
                    sWhere += " and to_char(a.rec_created_date,'DD-MON-YYYY') <= to_date('{EDATE}','DD-MON-YYYY')";
                }

                if (ord_showpending == "Y")
                    sWhere += " and a.ord_parent_id is null ";

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM joborderm a ";
                    sql += sWhere;

                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

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
                sql += " select a.ord_pkid,a.ord_parent_id,a.ord_invno,a.ord_po,a.ord_style,a.ord_cargo_status,a.ord_desc,a.ord_color,a.ord_contractno, ";
                sql += " a.ord_stylename,a.ord_cbm,a.ord_pcs,a.ord_pkg,a.ord_grwt,a.ord_ntwt,a.ord_hs_code,agent.cust_name as ord_agent_name,agent.cust_code as ord_agent_code,a.ord_agent_id,";
                sql += " a.ord_exp_id,exp.cust_name as ord_exp_name,a.ord_imp_id,a.rec_created_date,imp.cust_name as ord_imp_name,a.ord_uneco ";
                sql += " ,a.ord_track_status,agentbk.ab_book_no as ourbk_no,a.ord_booking_no as agentbk_no";
                sql += " ,ord_booking_date,ord_rnd_insp_date,ord_po_rel_date,ord_cargo_ready_date";
                sql += " ,ord_fcr_date, ord_insp_date, ord_stuf_date, ord_whd_date,ord_dlv_pol_date,ord_dlv_pod_date ";
                sql += " ,a.ord_approved as agent_approved,nvl(ord_pol,pol.param_code) as ord_pol,nvl(ord_pod,pod.param_code) as ord_pod,ord_uid,nvl(ord_status,'REPORTED') as ord_status  ";
                sql += " ,wkpln.ab_book_no as plan_no ,ord_week_no as week_no,ord_ftp_status ";
                sql += " ,ord_boarding1 ,ord_boarding2 ,ord_instock1 ,ord_instock2 ,ord_agentref_id";
                sql += " ,row_number() over(order by " + sort_colname + ") rn";
                sql += " from joborderm a ";
                sql += " left join param pol on a.ord_pol_id = pol.param_pkid";
                sql += " left join param pod on a.ord_pod_id = pod.param_pkid";
                sql += " left join customerm exp on a.ord_exp_id = exp.cust_pkid ";
                sql += " left join customerm imp on a.ord_imp_id = imp.cust_pkid ";
                sql += " left join customerm agent on a.ord_agent_id = agent.cust_pkid ";
                sql += " left join agentbookingm agentbk on a.ord_booking_id = agentbk.ab_pkid ";
                sql += " left join agentbookingm wkpln on a.ord_week_id = wkpln.ab_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";

               // sql += " order by ord_agent_name,ord_exp_name,ord_po "; agent.cust_name,exp.cust_name,ord_po

                 sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);


                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Joborderm();
                    mRow.ord_pkid = Dr["ord_pkid"].ToString();
                    mRow.ord_invno = Dr["ord_invno"].ToString();
                    mRow.ord_po = Dr["ord_po"].ToString();
                    mRow.ord_style = Dr["ord_style"].ToString();
                    mRow.ord_cargo_status = Dr["ord_cargo_status"].ToString();
                    mRow.ord_desc = Dr["ord_desc"].ToString();
                    mRow.ord_color = Dr["ord_color"].ToString();
                    mRow.ord_contractno = Dr["ord_contractno"].ToString();
                    mRow.ord_exp_name = Dr["ord_exp_name"].ToString();
                    mRow.ord_imp_name = Dr["ord_imp_name"].ToString();
                    //mRow.job_docno = Dr["job_docno"].ToString();
                    mRow.ord_stylename = Dr["ord_stylename"].ToString();
                    mRow.ord_cbm = Lib.Conv2Decimal(Dr["ord_cbm"].ToString());
                    mRow.ord_pcs = Lib.Conv2Decimal(Dr["ord_pcs"].ToString());
                    mRow.ord_pkg = Lib.Conv2Decimal(Dr["ord_pkg"].ToString());
                    mRow.ord_grwt = Lib.Conv2Decimal(Dr["ord_grwt"].ToString());
                    mRow.ord_ntwt = Lib.Conv2Decimal(Dr["ord_ntwt"].ToString());
                    mRow.ord_hs_code = Dr["ord_hs_code"].ToString();
                    mRow.ord_agent_name = Dr["ord_agent_name"].ToString();
                    mRow.ord_agent_code = Dr["ord_agent_code"].ToString();
                    mRow.rec_created_dte = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mRow.ord_uneco = Dr["ord_uneco"].ToString();
                    mRow.ord_track_status = Dr["ord_track_status"].ToString();
                    mRow.ord_booking_no = Dr["agentbk_no"].ToString();
                    mRow.ord_approved = (Dr["agent_approved"].ToString() == "Y" ? true : false);
                    mRow.ord_ourbooking_no = Dr["ourbk_no"].ToString();
                    mRow.ord_booking_date = Lib.DatetoString(Dr["ord_booking_date"]);
                    mRow.ord_rnd_insp_date = Lib.DatetoString(Dr["ord_rnd_insp_date"]);
                    mRow.ord_po_rel_date = Lib.DatetoString(Dr["ord_po_rel_date"]);
                    mRow.ord_cargo_ready_date = Lib.DatetoString(Dr["ord_cargo_ready_date"]);
                    mRow.ord_fcr_date = Lib.DatetoString(Dr["ord_fcr_date"]);
                    mRow.ord_insp_date = Lib.DatetoString(Dr["ord_insp_date"]);
                    mRow.ord_stuf_date = Lib.DatetoString(Dr["ord_stuf_date"]);
                    mRow.ord_whd_date = Lib.DatetoString(Dr["ord_whd_date"]);
                    mRow.ord_dlv_pol_date = Lib.DatetoString(Dr["ord_dlv_pol_date"]);
                    mRow.ord_dlv_pod_date = Lib.DatetoString(Dr["ord_dlv_pod_date"]);
                    mRow.ord_pol = Dr["ord_pol"].ToString();
                    mRow.ord_pod = Dr["ord_pod"].ToString();
                    mRow.ord_status = Dr["ord_status"].ToString();
                    mRow.ord_uid = Lib.Conv2Integer(Dr["ord_uid"].ToString());
                    mRow.ord_plan_no = Lib.Conv2Integer(Dr["plan_no"].ToString());
                    mRow.ord_week_no = Lib.Conv2Integer(Dr["week_no"].ToString());
                    mRow.ord_ftp_status = Dr["ord_ftp_status"].ToString();
                    mRow.ord_boarding1 = Lib.DatetoString(Dr["ord_boarding1"]);
                    mRow.ord_boarding2 = Lib.DatetoString(Dr["ord_boarding2"]);
                    mRow.ord_instock1 = Lib.DatetoString(Dr["ord_instock1"]);
                    mRow.ord_instock2 = Lib.DatetoString(Dr["ord_instock2"]);
                    mRow.ord_agentref_id = Dr["ord_agentref_id"].ToString();

                    if (mRow.ord_status == "REPORTED")
                        mRow.ord_status_color = "BLUE";
                    if (mRow.ord_status == "APPROVED")
                        mRow.ord_status_color = "GREEN";
                    if (mRow.ord_status == "CANCELLED")
                        mRow.ord_status_color = "RED";
                    if (mRow.ord_status == "ON HOLD")
                        mRow.ord_status_color = "PURPLE";

                    mList.Add(mRow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                    {
                        PrintOrderList(mList,branch_code, file_pkid);

                    }
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("list", mList);
            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rows", page_rows);
            RetData.Add("page_rowcount", page_rowcount);

            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Joborderm mRow = new Joborderm();

            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select ord_pkid,ord_exp_id,exp.cust_name as ord_exp_name,ord_imp_id,imp.cust_name as ord_imp_name,ord_invno, ";
                sql += " ord_uneco,ord_po,ord_style,ord_cbm,ord_pcs,ord_pkg,ord_grwt,ord_ntwt, ord_status, ";
                sql += " ord_hs_code,ord_cargo_status,ord_desc,ord_color,ord_stylename,ord_contractno,exp.cust_code as exp_code,imp.cust_code as imp_code, ";
                sql += " ord_agent_id,agent.cust_name as ord_agent_name,agent.cust_code as agent_code, ";
                sql += " a.ord_boarding1, a.ord_boarding2, a.ord_instock1, a.ord_instock2, ";
                sql += " a.ord_booking_date,a.ord_rnd_insp_date,a.ord_po_rel_date,a.ord_cargo_ready_date, ";
                sql += " a.ord_fcr_date,a.ord_insp_date,a.ord_stuf_date,a.ord_whd_date, a.ord_dlv_pol_date, a.ord_dlv_pod_date,";
                sql += " ord_pol_id,pol.param_code as ord_pol_code,ord_pol,ord_pod_id,pod.param_code as ord_pod_code,ord_pod,ord_uid,a.rec_category ";
                sql += " from joborderm a  ";
                sql += " left join customerm exp on a.ord_exp_id = exp.cust_pkid  ";
                sql += " left join customerm imp on a.ord_imp_id = imp.cust_pkid  ";
                sql += " left join customerm agent on a.ord_agent_id = agent.cust_pkid  ";
                sql += " left join param pol on a.ord_pol_id = pol.param_pkid ";
                sql += " left join param pod on a.ord_pod_id = pod.param_pkid ";
                sql += " where  a.ord_pkid ='" + id + "'";


                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Joborderm();
                    mRow.ord_pkid = Dr["ord_pkid"].ToString();
                    mRow.ord_status = Dr["ord_status"].ToString();
                    mRow.ord_exp_id = Dr["ord_exp_id"].ToString();
                    mRow.ord_exp_name = Dr["ord_exp_name"].ToString();
                    mRow.ord_imp_id = Dr["ord_imp_id"].ToString();
                    mRow.ord_imp_name = Dr["ord_imp_name"].ToString();
                    mRow.ord_invno = Dr["ord_invno"].ToString();
                    mRow.ord_uneco = Dr["ord_uneco"].ToString();
                    mRow.ord_po = Dr["ord_po"].ToString();
                    mRow.ord_style = Dr["ord_style"].ToString();
                    //mRow.job_docno = Dr["job_docno"].ToString();
                    mRow.ord_cbm = Lib.Conv2Decimal(Dr["ord_cbm"].ToString());
                    mRow.ord_pcs = Lib.Conv2Decimal(Dr["ord_pcs"].ToString());
                    mRow.ord_pkg = Lib.Conv2Decimal(Dr["ord_pkg"].ToString());
                    mRow.ord_grwt = Lib.Conv2Decimal(Dr["ord_grwt"].ToString());
                    mRow.ord_ntwt = Lib.Conv2Decimal(Dr["ord_ntwt"].ToString());
                    mRow.ord_hs_code = Dr["ord_hs_code"].ToString();
                    mRow.ord_cargo_status = Dr["ord_cargo_status"].ToString();
                    mRow.ord_desc = Dr["ord_desc"].ToString();
                    mRow.ord_color = Dr["ord_color"].ToString();
                    mRow.ord_stylename = Dr["ord_stylename"].ToString();
                    mRow.ord_contractno = Dr["ord_contractno"].ToString();
                    mRow.ord_exp_code = Dr["exp_code"].ToString();
                    mRow.ord_imp_code = Dr["imp_code"].ToString();
                    mRow.ord_agent_id = Dr["ord_agent_id"].ToString();
                    mRow.ord_agent_name = Dr["ord_agent_name"].ToString();
                    mRow.ord_agent_code = Dr["agent_code"].ToString();
                    mRow.ord_pol = Dr["ord_pol"].ToString();
                    mRow.ord_pod = Dr["ord_pod"].ToString();
                    mRow.ord_pol_id = Dr["ord_pol_id"].ToString();
                    mRow.ord_pod_id = Dr["ord_pod_id"].ToString();
                    mRow.ord_pol_code = Dr["ord_pol_code"].ToString();
                    mRow.ord_pod_code = Dr["ord_pod_code"].ToString();
                    mRow.ord_uid = Lib.Conv2Integer(Dr["ord_uid"].ToString());
                    mRow.rec_category = Dr["rec_category"].ToString();

                    mRow.ord_boarding1 = Lib.DatetoString(Dr["ord_boarding1"]);
                    mRow.ord_boarding2 = Lib.DatetoString(Dr["ord_boarding2"]);
                    mRow.ord_instock1 = Lib.DatetoString(Dr["ord_instock1"]);
                    mRow.ord_instock2 = Lib.DatetoString(Dr["ord_instock2"]);

                    mRow.ord_booking_date = Lib.DatetoString(Dr["ord_booking_date"]);
                    mRow.ord_rnd_insp_date = Lib.DatetoString(Dr["ord_rnd_insp_date"]);
                    mRow.ord_po_rel_date = Lib.DatetoString(Dr["ord_po_rel_date"]);
                    mRow.ord_cargo_ready_date = Lib.DatetoString(Dr["ord_cargo_ready_date"]);
                    mRow.ord_fcr_date = Lib.DatetoString(Dr["ord_fcr_date"]);
                    mRow.ord_insp_date = Lib.DatetoString(Dr["ord_insp_date"]);
                    mRow.ord_stuf_date = Lib.DatetoString(Dr["ord_stuf_date"]);
                    mRow.ord_whd_date = Lib.DatetoString(Dr["ord_whd_date"]);
                    mRow.ord_dlv_pol_date = Lib.DatetoString(Dr["ord_dlv_pol_date"]);
                    mRow.ord_dlv_pod_date = Lib.DatetoString(Dr["ord_dlv_pod_date"]);


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


        public string AllValid(Joborderm Record)
        {
            string str = "";

            try
            {
                if (Record.ord_po.Length <= 0)
                    Lib.AddError(ref str, " PO No Cannot Be Blank");

                sql = "";
                sql = "select ord_pkid from (";
                sql += " select ord_pkid from joborderm ";
                sql += " where rec_company_code = '{COMPANY_CODE}' ";
                sql += " and rec_branch_code = '{BRANCH_CODE}'";
                sql += " and ord_exp_id = '{EXP_ID}' ";
                sql += " and ord_po = '{PO}' ";

                if (Record.ord_style != "")
                    sql += " and ord_style = '{STYLE}' ";
                else
                    sql += " and ord_style  is null ";

                if (Record.ord_color != "")
                    sql += " and ord_color = '{COLOR}' ";
                else
                    sql += " and ord_color is null";

                sql += ") a where ord_pkid <> '{PKID}'";

                sql = sql.Replace("{COMPANY_CODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{BRANCH_CODE}", Record._globalvariables.branch_code);
                sql = sql.Replace("{EXP_ID}", Record.ord_exp_id);
                sql = sql.Replace("{PO}", Record.ord_po);
                sql = sql.Replace("{STYLE}", Record.ord_style);
                sql = sql.Replace("{COLOR}", Record.ord_color);
                sql = sql.Replace("{PKID}", Record.ord_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " | This PO No Already Exists");

                if (Record.rec_mode == "EDIT" && Lib.Conv2Integer(Record.ord_uid.ToString()) > 0)
                {
                    sql = "select ord_pkid from (";
                    sql += "select ord_pkid from joborderm a where a.ord_uid = {UID}  ";
                    sql += " and a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.ord_agent_id = '{AGENTID}'";
                    sql += ") a where ord_pkid <> '{PKID}'";

                    sql = sql.Replace("{UID}", Record.ord_uid.ToString());
                    sql = sql.Replace("{PKID}", Record.ord_pkid);
                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{AGENTID}", Record.ord_agent_id);

                    if (Con_Oracle.IsRowExists(sql))
                        Lib.AddError(ref str, " | This PO ID Already Exists");
                }

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Joborderm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            string uidno = "";
            bool bUidChanged = false;
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
                    Record.ord_agentref_id = "CM" + Lib.getProcessNumber(Record._globalvariables.comp_code, "JOB-ORDER", "JOB-ORDER");
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("joborderm", Record.rec_mode, "ord_pkid", Record.ord_pkid);
                Rec.InsertString("ord_exp_id", Record.ord_exp_id);
                Rec.InsertString("ord_exp_name", Record.ord_exp_name);
                Rec.InsertString("ord_imp_id", Record.ord_imp_id);
                Rec.InsertString("ord_imp_name", Record.ord_imp_name);
                Rec.InsertString("ord_invno", Record.ord_invno);
                Rec.InsertString("ord_uneco", Record.ord_uneco);
                Rec.InsertString("ord_po", Record.ord_po);
                Rec.InsertString("ord_style", Record.ord_style);
                Rec.InsertNumeric("ord_cbm", Lib.Conv2Decimal(Record.ord_cbm.ToString()).ToString());
                Rec.InsertNumeric("ord_pcs", Lib.Conv2Decimal(Record.ord_pcs.ToString()).ToString());
                Rec.InsertNumeric("ord_pkg", Lib.Conv2Decimal(Record.ord_pkg.ToString()).ToString());
                Rec.InsertNumeric("ord_grwt", Lib.Conv2Decimal(Record.ord_grwt.ToString()).ToString());
                Rec.InsertNumeric("ord_ntwt", Lib.Conv2Decimal(Record.ord_ntwt.ToString()).ToString());
                Rec.InsertString("ord_hs_code", Record.ord_hs_code);
                Rec.InsertString("ord_cargo_status", Record.ord_cargo_status);
                Rec.InsertString("ord_desc", Record.ord_desc);
                Rec.InsertString("ord_color", Record.ord_color);
                Rec.InsertString("ord_stylename", Record.ord_stylename);
                Rec.InsertString("ord_contractno", Record.ord_contractno);
                Rec.InsertString("ord_agent_id", Record.ord_agent_id);
                Rec.InsertString("ord_agent_name", Record.ord_agent_name);
                Rec.InsertString("ord_pol_id", Record.ord_pol_id);
                Rec.InsertString("ord_pol", Record.ord_pol);
                Rec.InsertString("ord_pod_id", Record.ord_pod_id);
                Rec.InsertString("ord_pod", Record.ord_pod);

                Rec.InsertDate("ord_boarding1", Record.ord_boarding1);
                Rec.InsertDate("ord_boarding2", Record.ord_boarding2);
                Rec.InsertDate("ord_instock1", Record.ord_instock1);
                Rec.InsertDate("ord_instock2", Record.ord_instock2);


                Rec.InsertString("rec_category", Record.rec_category);
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("ord_status",Record.ord_status);
                    Rec.InsertString("ord_agentref_id", Record.ord_agentref_id);
                    Rec.InsertString("ord_parent_id", Record.ord_parent_id);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "SYSDATE");
                    Rec.InsertString("ord_source", "ORDER");
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

                string srem = (Record.rec_mode == "ADD") ? "ADDED" : "EDITED";
                Lib.AuditLog("PO", "PO", Record.rec_mode, Record._globalvariables.comp_code, Record._globalvariables.branch_code, Record._globalvariables.user_code, Record.ord_pkid,  Record.ord_po,  srem );
                
                
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

            RetData.Add("uidno", uidno);
            return RetData;
        }


        public Dictionary<string, object> Upload(JobOrder_VM VM)
        {

            string SHPR_ID = "";
            string CNSG_ID = "";
            string PO = "";
            string STYLE = "";
            string BRANCH = "";
            string COMPANY = "";
            string AGENT_ID = "";
            string PARNT_ID = "";
            string COLOR = "";
            int i = 0;
            List<Joborderm> mList = new List<Joborderm>();
            Joborderm mRow = new Joborderm();
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            DataTable Dt_Uid;
            DataTable Dt_Rec = new DataTable();
            DBRecord mRec;
            bool bTrans = false;
            Con_Oracle = new DBConnection();

            string DateFormat = "DMY";

            try
            {
                GlobalVariables mGbl = VM.globalVariables;
                BRANCH = mGbl.branch_code.ToString();
                COMPANY = mGbl.comp_code.ToString();
                if (VM.ord_source == "ORDER")
                {
                    SHPR_ID = VM.ord_exp_id.ToString();
                    CNSG_ID = VM.ord_imp_id.ToString();
                    AGENT_ID = VM.ord_agent_id.ToString();
                }
                if (VM.ord_source == "JOB")
                {

                    PARNT_ID = VM.ord_parent_id.ToString();

                    sql = "";
                    sql = "select job_exp_id,job_imp_id,job_agent_id from jobm ";
                    sql += " where job_pkid  = '{PKID}' ";
                    sql = sql.Replace("{PKID}", PARNT_ID);

                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);


                    if (Dt_Temp.Rows.Count > 0)
                    {
                        SHPR_ID = Dt_Temp.Rows[0]["job_exp_id"].ToString();
                        CNSG_ID = Dt_Temp.Rows[0]["job_imp_id"].ToString();
                        AGENT_ID = Dt_Temp.Rows[0]["job_agent_id"].ToString();
                    }

                    if (AGENT_ID.Trim() == "")
                    {
                        throw new Exception("Job Agent Cannot Be Blank");
                    }
                }

                foreach (var Rec in VM.JobOrder)
                {

                    PO = "";
                    STYLE = "";
                    COLOR = "";

                    PO = Rec.ord_po.ToString();
                    STYLE = Rec.ord_style.ToString();
                    COLOR = Rec.ord_color.ToString();

                    sql = "";
                    sql += " select count(ord_pkid) from joborderm where ";
                    sql += " rec_company_code = '{COMP}' and rec_branch_code = '{BRANCH}'";
                    sql += " and ord_exp_id = '{EXP_ID}' and ord_imp_id = '{IMP_ID}' ";
                    sql += " and ord_po = '{PO}' ";

                    if (STYLE != "")
                        sql += " and ord_style = '{STYLE}' ";
                    else
                        sql += " and ord_style  is null ";

                    if (COLOR != "")
                        sql += " and ord_color = '{COLOR}' ";
                    else
                        sql += " and ord_color  is null ";

                    sql = sql.Replace("{COMP}", COMPANY);
                    sql = sql.Replace("{BRANCH}", BRANCH);
                    sql = sql.Replace("{EXP_ID}", SHPR_ID);
                    sql = sql.Replace("{IMP_ID}", CNSG_ID);
                    sql = sql.Replace("{PO}", PO);
                    sql = sql.Replace("{STYLE}", STYLE);
                    sql = sql.Replace("{COLOR}", COLOR);

                    i = Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());

                    if (i > 0)
                    {
                        mRow = new Joborderm();
                        mRow.ord_pkid = Rec.ord_pkid.ToString();
                        mList.Add(mRow);
                    }
                    else
                    {
                        Rec.ord_agentref_id = "CM" + Lib.getProcessNumber(mGbl.comp_code, "JOB-ORDER", "JOB-ORDER");

                        mRec = new DBRecord();
                        mRec.CreateRow("joborderm", "ADD", "ord_pkid", Rec.ord_pkid.ToString());
                        mRec.InsertString("ord_exp_id", SHPR_ID);
                        mRec.InsertString("ord_imp_id", CNSG_ID);
                        mRec.InsertString("ord_agent_id", AGENT_ID);
                        mRec.InsertString("ord_invno", Rec.ord_invno.ToString());
                        mRec.InsertString("ord_uneco", Rec.ord_uneco);
                        mRec.InsertString("ord_desc", Rec.ord_desc.ToString());
                        mRec.InsertString("ord_po", Rec.ord_po.ToString());
                        mRec.InsertString("ord_style", Rec.ord_style.ToString());
                        mRec.InsertString("ord_color", Rec.ord_color.ToString());
                        mRec.InsertNumeric("ord_pkg", Lib.Conv2Decimal(Rec.ord_pkg.ToString()).ToString());
                        mRec.InsertNumeric("ord_pcs", Lib.Conv2Decimal(Rec.ord_pcs.ToString()).ToString());
                        mRec.InsertNumeric("ord_ntwt", Lib.Conv2Decimal(Rec.ord_ntwt.ToString()).ToString());
                        mRec.InsertNumeric("ord_grwt", Lib.Conv2Decimal(Rec.ord_grwt.ToString()).ToString());
                        mRec.InsertNumeric("ord_cbm", Lib.Conv2Decimal(Rec.ord_cbm.ToString()).ToString());
                        mRec.InsertString("ord_hs_code", Rec.ord_hs_code.ToString());
                        mRec.InsertString("ord_contractno", Rec.ord_contractno.ToString());
                        mRec.InsertString("rec_category", Rec.rec_category.ToString());
                        mRec.InsertDate("ord_booking_date", GetFormatDate(Rec.ord_booking_date,DateFormat));
                        mRec.InsertDate("ord_rnd_insp_date", GetFormatDate(Rec.ord_rnd_insp_date, DateFormat));
                        mRec.InsertDate("ord_po_rel_date", GetFormatDate(Rec.ord_po_rel_date, DateFormat));
                        mRec.InsertDate("ord_cargo_ready_date", GetFormatDate(Rec.ord_cargo_ready_date, DateFormat));
                        mRec.InsertDate("ord_fcr_date", GetFormatDate(Rec.ord_fcr_date, DateFormat));
                        mRec.InsertDate("ord_insp_date", GetFormatDate(Rec.ord_insp_date, DateFormat));
                        mRec.InsertDate("ord_stuf_date", GetFormatDate(Rec.ord_stuf_date, DateFormat));
                        mRec.InsertDate("ord_whd_date", GetFormatDate(Rec.ord_whd_date, DateFormat));
                        mRec.InsertDate("ord_dlv_pol_date", GetFormatDate(Rec.ord_dlv_pol_date, DateFormat));
                        mRec.InsertDate("ord_dlv_pod_date", GetFormatDate(Rec.ord_dlv_pod_date, DateFormat));
                        mRec.InsertString("ord_pol", Rec.ord_pol.ToString());
                        mRec.InsertString("ord_pod", Rec.ord_pod.ToString());
                        mRec.InsertString("rec_company_code", mGbl.comp_code.ToString());
                        mRec.InsertString("rec_branch_code", mGbl.branch_code.ToString());
                        mRec.InsertString("rec_created_by", mGbl.user_name.ToString());
                        mRec.InsertFunction("rec_created_date", "SYSDATE");
                        mRec.InsertString("ord_source", VM.ord_source.ToString());
                        //mRec.InsertNumeric("ord_uid", Rec.ord_uid.ToString());
                        mRec.InsertString("ord_agentref_id", Rec.ord_agentref_id);

                        if (VM.ord_source == "JOB")
                            mRec.InsertString("ord_parent_id", PARNT_ID);

                        sql = mRec.UpdateRow();

                        Con_Oracle.BeginTransaction();
                        bTrans = true;
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                    }
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    if (bTrans)
                        Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }

            Con_Oracle.CloseConnection();
            RetData.Add("list", mList);
            return RetData;
        }

        private string GetFormatDate(string sDate,string sFormat)
        {
            string[] sData = null;
            int dd = 0, mm = 0, yy = 0;
            if (sDate.Contains("/"))
                sData = sDate.Split('/');
            else if (sDate.Contains("-"))
                sData = sDate.Split('-');
            else if (sDate.Contains("."))
                sData = sDate.Split('.');

            sDate = "";
            if (sData != null)
            {
                if (sData.Length == 3)
                {
                    if (sFormat == "DMY")
                    {
                        dd = Lib.Conv2Integer(sData[0]);
                        mm = Lib.Conv2Integer(sData[1]);
                        yy = Lib.Conv2Integer(sData[2]);
                    }
                    else if (sFormat == "MDY")
                    {
                        mm = Lib.Conv2Integer(sData[0]);
                        dd = Lib.Conv2Integer(sData[1]);
                        yy = Lib.Conv2Integer(sData[2]);
                    }
                    else if (sFormat == "YMD")
                    {
                        yy = Lib.Conv2Integer(sData[0]);
                        mm = Lib.Conv2Integer(sData[1]);
                        dd = Lib.Conv2Integer(sData[2]);
                    }

                    if (yy < 100)
                        yy = yy + 2000;
                    sDate = new DateTime(yy, mm, dd).ToString("yyyy-MM-dd");
                }
            }
            return sDate;
        }

        private void PrintOrderList(List<Joborderm> mList,string branch_code,string fileID)
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

                File_Display_Name = "OrderReport.xls";
                File_Name = Lib.GetFileName(report_folder, fileID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;


                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 20;
                WS.Columns[2].Width = 256 * 30;
                WS.Columns[3].Width = 256 * 25;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 45;
                WS.Columns[6].Width = 256 * 25;
                WS.Columns[7].Width = 256 * 20;
                WS.Columns[8].Width = 256 * 40;
                WS.Columns[9].Width = 256 * 10;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 10;
                WS.Columns[13].Width = 256 * 10;
                WS.Columns[14].Width = 256 * 10;
                WS.Columns[15].Width = 256 * 10;
                WS.Columns[16].Width = 256 * 10;
                WS.Columns[17].Width = 256 * 10;
                WS.Columns[18].Width = 256 * 10;
                WS.Columns[19].Width = 256 * 10;
                WS.Columns[20].Width = 256 * 10;
                WS.Columns[21].Width = 256 * 10;
                WS.Columns[22].Width = 256 * 10;
                WS.Columns[23].Width = 256 * 10;
                WS.Columns[24].Width = 256 * 10;
                WS.Columns[25].Width = 256 * 10;
                WS.Columns[26].Width = 256 * 10;

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
                Lib.WriteData(WS, iRow, 1, "ORDER LIST ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                
                iRow++;
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVOICE#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DESCRIPTION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PO#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STYLE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTRACT#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "UNECO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COLOR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HS.CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PKGS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PCS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NTWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRWT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BKD.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RND.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POR.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FCR.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INSP.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STUF.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "WHD.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DLV.POL.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DLV.POD.DT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                foreach (Joborderm Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    i++;
    
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_exp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_imp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_invno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_desc , _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_po, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_style, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_contractno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_uneco, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_color, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_hs_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_pkg, _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_pcs, _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_ntwt, _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_grwt, _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_cbm, _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_booking_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_rnd_insp_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_po_rel_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_cargo_ready_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_fcr_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_insp_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_stuf_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_whd_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_dlv_pol_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.ord_dlv_pod_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
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


        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "mappingm");
            parameter.Add("type", "ORDER");
            parameter.Add("branch_code", SearchData["branch_code"].ToString());
            RetData.Add("ordercolumns", lovservice.Lov(parameter)["mappingm"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "acgroupm");
            //RetData.Add("acgroupm", lovservice.Lov(parameter)["acgroupm"]);

            return RetData;
        }

        public Dictionary<string, object> DeleteRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string id = SearchData["pkid"].ToString();

            try
            {
                Con_Oracle = new DBConnection();

                if (id.Contains(","))
                    id = id.Replace(",", "','");

                sql = "delete from joborderm where ord_pkid in ('" + id + "')";

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
            return RetData;
        }
    }
}
