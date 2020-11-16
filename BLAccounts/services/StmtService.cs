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

namespace BLAccounts
{
    public class StmtService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;
        int iRow = 0;
        int iCol = 0;


        List<Stmtd> mList;

        DataTable Dt_List;


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Stmtm> mList = new List<Stmtm>();
            Stmtm mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();

            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            string year_code = SearchData["year_code"].ToString();

            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  1=1 ";



                sWhere += " and (";
                sWhere += " a.rec_company_code = '{COMP}'";
                sWhere += " and a.rec_branch_code = '{BRANCH}'";
                sWhere += " and a.stm_year =  {FINYEAR}";
                sWhere += " ) ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += " or acc_name like '%" + searchstring + "%'";
                    sWhere += ")";
                }




                sWhere = sWhere.Replace("{COMP}", company_code);
                sWhere = sWhere.Replace("{BRANCH}", branch_code);
                sWhere = sWhere.Replace("{FINYEAR}", year_code);


                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM stmtm a ";
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
                sql += "  select stm_pkid,stm_no,stm_date,acc_code, acc_name, c.param_code as curr_code,";
                sql += " stm_dr, stm_cr, stm_bal, stm_dr_inr, stm_cr_inr, stm_bal_inr,";
                sql += " row_number() over(order by stm_date,stm_no) rn, ";
                sql += " a.rec_created_by, a.rec_created_date ";
                sql += " from stmtm a ";
                sql += " left join acctm b on a.stm_accid = b.acc_pkid ";
                sql += " left join param c on a.stm_currencyid = c.param_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by stm_date, stm_no";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Stmtm();
                    mRow.stm_pkid = Dr["stm_pkid"].ToString();
                    mRow.stm_date = Lib.DatetoStringDisplayformat(Dr["stm_date"]);
                    mRow.stm_no = Lib.Conv2Integer(Dr["stm_no"].ToString());
                    mRow.stm_acc_code = Dr["acc_code"].ToString();
                    mRow.stm_acc_name = Dr["acc_name"].ToString();
                    mRow.stm_curr_code = Dr["curr_code"].ToString();

                    mRow.stm_dr = Lib.Conv2Decimal(Dr["stm_dr"].ToString());
                    mRow.stm_cr = Lib.Conv2Decimal(Dr["stm_cr"].ToString());
                    mRow.stm_bal = Lib.Conv2Decimal(Dr["stm_bal"].ToString());

                    mRow.stm_dr_inr = Lib.Conv2Decimal(Dr["stm_dr_inr"].ToString());
                    mRow.stm_cr_inr = Lib.Conv2Decimal(Dr["stm_cr_inr"].ToString());
                    mRow.stm_bal_inr = Lib.Conv2Decimal(Dr["stm_bal_inr"].ToString());

                    mRow.rec_created_by = Dr["rec_created_by"].ToString();
                    mRow.rec_created_date = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
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


        public Dictionary<string, object> GetPendingList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();


            Con_Oracle = new DBConnection();

            string pkid = SearchData["pkid"].ToString();
            string accid = SearchData["accid"].ToString();
            string currid = SearchData["currid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            string stm_date = SearchData["stm_date"].ToString();
            string todate = Lib.StringToDate(stm_date);

            List<Stmtd> mList = new List<Stmtd>();
            Stmtd cRow;

            StringBuilder sql = new StringBuilder();


            try
            {

                sql.Append(" select * from ");
                sql.Append(" (  ");
                sql.Append(" select jvh_pkid ,jv_pkid, jv_acc_id ,jvh_year ,jvh_vrno ,jvh_type ,jvh_date , ");
                sql.Append(" jvh_remarks , jvh_reference , ");
                // FCURR BALANCE
                sql.Append(" jv_curr_id,");
                sql.Append(" jv_ftotal as Amount,  ");
                sql.Append(" jv_exrate,  ");
                sql.Append(" case when jv_debit <>0 then jv_ftotal else 0 end as DR, ");
                sql.Append(" case when jv_credit <>0 then jv_ftotal else 0 end as CR, ");
                sql.Append(" jv_ftotal - nvl(xref_Amt,0) as balance, ");
                sql.Append(" nvl(xref_allocated_amt,0) as allocation, ");
                // LCURR BALANCE
                sql.Append(" jv_debit, jv_credit, ");
                sql.Append(" round((nvl(jv_debit,0) + nvl(jv_credit,0)) - (nvl(xref_Amt,0) * jv_exrate),2) as inrBalance, ");
                sql.Append(" round(nvl(Xref_Allocated_Amt,0) * jv_exrate,2) as inrAllocation, ");
                sql.Append(" a.rec_category ");
                sql.Append(" from ledgerh a inner join ledgert t on a.jvh_pkid = t.jv_parent_id ");
                sql.Append(" left join  (");
                sql.Append(" select std_jv_pkid,std_currencyid,");
                sql.Append(" sum(case when STD_PARENTID='{STD_PARENT_ID}' then 0        else std_amt end ) as xref_Amt,");
                sql.Append(" sum(case when STD_PARENTID='{STD_PARENT_ID}' then std_amt else 0 end ) as xref_Allocated_Amt ");
                sql.Append(" from stmtd where ");
                sql.Append(" std_accid =  '{ACCOUNT}' ");
                sql.Append(" and std_currencyid='{CURRENCY}' ");
                sql.Append(" and rec_company_code= '{COMPANY}' ");
                sql.Append(" and rec_branch_code= '{BRANCH}' ");
                sql.Append(" group by std_jv_pkid,std_currencyid ");
                sql.Append(" )  b on t.jv_pkid = b.std_jv_pkid ");
                sql.Append(" and t.jv_curr_id=b.std_currencyid ");
                sql.Append(" where  t.jv_acc_id = '{ACCOUNT}' ");
                sql.Append(" and t.jv_curr_id = '{CURRENCY}' ");
                sql.Append(" and a.jvh_type not in ('OP','OB','OI') ");
                sql.Append(" and a.rec_deleted = 'N' ");
                sql.Append(" and nvl(a.jvh_reference,'A') <> 'COSTING ADJUSTMENT' ");
                sql.Append(" and a.rec_company_code= '{COMPANY}' ");
                sql.Append(" and a.rec_branch_code= '{BRANCH}' ");
                sql.Append(" )  jv ");
                sql.Append(" where (  balance != 0 )");
                sql.Append(" order by jvh_reference,jvh_vrno ");


                sql.Replace("{STD_PARENT_ID}", pkid);
                sql.Replace("{ACCOUNT}", accid);
                sql.Replace("{CURRENCY}", currid);
                sql.Replace("{COMPANY}", comp_code);
                sql.Replace("{BRANCH}", branch_code);

                DataTable Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql.ToString());
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    cRow = new Stmtd();
                    cRow.std_parentid = pkid;
                    cRow.jv_entity_id = Dr["jvh_pkid"].ToString();
                    cRow.jv_pk_id = Dr["jv_pkid"].ToString();
                    cRow.jv_ac_rowid = Dr["jv_acc_id"].ToString();
                    cRow.jv_year = Lib.Conv2Integer( Dr["jvh_year"].ToString());
                    cRow.jv_vrno = Lib.Conv2Integer(Dr["jvh_vrno"].ToString());
                    cRow.jv_type = Dr["jvh_type"].ToString();
                    cRow.jv_date = Lib.DatetoString(Dr["jvh_date"]);
                    cRow.jv_display_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    cRow.jv_reference = Dr["jvh_reference"].ToString();
                    cRow.jv_remarks = Dr["jvh_remarks"].ToString();
                    cRow.rec_category = Dr["rec_category"].ToString();
                    cRow.jv_currency_rowid = Dr["jv_curr_id"].ToString();
                    cRow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    cRow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                    cRow.jv_exchange_rate = Lib.Conv2Decimal(Dr["jv_exrate"].ToString());
                    cRow.inrbalance = Lib.Conv2Decimal(Dr["inrbalance"].ToString());
                    cRow.inrallocation = Lib.Conv2Decimal(Dr["inrallocation"].ToString());
                    cRow.amount = Lib.Conv2Decimal(Dr["amount"].ToString());
                    cRow.dr = Lib.Conv2Decimal(Dr["DR"].ToString());
                    cRow.cr = Lib.Conv2Decimal(Dr["CR"].ToString());
                    cRow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    cRow.allocation = Lib.Conv2Decimal(Dr["allocation"].ToString());
                    cRow.jv_selected = false;
                    if (Lib.Conv2Decimal(Dr["allocation"].ToString()) != 0)
                        cRow.jv_selected = true;
                    mList.Add(cRow);
                }

                Dt_Rec.Rows.Clear();
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

        
        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Stmtm mRow = new Stmtm();

            try
            {
                string id = SearchData["pkid"].ToString();
                DataTable Dt_Rec = new DataTable();

                sql = "";
                sql += " select ";
                sql += " stm_pkid, stm_date, stm_year, stm_no,  ";
                sql += " stm_accid, b.acc_code, b.acc_name,stm_acc_br_id,";
                sql += " stm_currencyid, curr.param_code as curr_code, ";
                sql += " add_branch_slno,addr.add_line1||'\n'||addr.add_line2||'\n'||addr.add_line3 as  acc_br_addr, ";
                sql += " a.rec_locked, stm_edit_code, stm_edit_date ";
                sql += " from stmtm a ";
                sql += " left join acctm b on stm_accid = b.acc_pkid ";
                sql += " left join param curr on stm_currencyid = curr.param_pkid ";
                sql += " left join addressm addr on stm_acc_br_id = add_pkid ";

                sql += " where  a.stm_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Stmtm();
                    mRow.stm_pkid = Dr["stm_pkid"].ToString();
                    mRow.stm_date = Lib.DatetoString(Dr["stm_date"]);
                    mRow.stm_no = Lib.Conv2Integer(Dr["stm_no"].ToString());
                    mRow.stm_accid = Dr["stm_accid"].ToString();

                    mRow.stm_acc_br_no = Dr["add_branch_slno"].ToString();
                    mRow.stm_acc_br_id = Dr["stm_acc_br_id"].ToString();

                    mRow.stm_acc_code = Dr["acc_code"].ToString();
                    mRow.stm_acc_name = Dr["acc_name"].ToString();

                    mRow.stm_acc_br_addr = Dr["acc_br_addr"].ToString();

                    mRow.stm_curr_code = Dr["curr_code"].ToString();
                    mRow.stm_currencyid = Dr["stm_currencyid"].ToString();

                    mRow.stm_edit_code = Dr["stm_edit_code"].ToString();
                    mRow.stm_edit_date = Dr["stm_edit_date"].ToString();
                    if (Dr["rec_locked"].ToString() == "Y" && Dr["jvh_edit_date"].ToString() != System.DateTime.Today.ToString("yyyyMMdd"))
                    {
                        mRow.stm_edit_code = "";
                        mRow.rec_locked = true;
                    }



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


        public string AllValid(Stmtm Record)
        {
            string str = "";
            try
            {
                sql = "";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Stmtm Record)
        {
            int sm_no = 0;
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if (Record.stm_acc_name.Trim().Length <= 0 || Record.stm_accid.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "A/c Name Cannot Be Empty");

                if (Record.stm_currencyid.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Currency Cannot Be Empty");

                ErrorMessage = AllValid(Record);

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if (Record.rec_mode == "ADD")
                {
                    sql = "select nvl(max(stm_no),1000) + 1 as stmt_no  from stmtm where ";
                    sql += " rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    sql += " and rec_branch_code = '" + Record._globalvariables.branch_code + "'";
                    sm_no = Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("stmtm", Record.rec_mode, "stm_pkid", Record.stm_pkid);
                if (Record.rec_mode == "ADD")
                {
                    Record.stm_no = sm_no;
                    Rec.InsertNumeric("stm_year", Record._globalvariables.year_code);
                    Rec.InsertNumeric("stm_no", sm_no.ToString());
                    Rec.InsertString("stm_isgroup", "N");
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    Rec.InsertString("rec_locked", "N");
                    Rec.InsertString("rec_deleted", "N");
                    Rec.InsertString("stm_edit_code", "{S}{D}");
                    Rec.InsertString("stm_edit_date", System.DateTime.Today.ToString("yyyyMMdd"));

                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "SYSDATE");
                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "SYSDATE");
                }

                Rec.InsertDate("stm_date", Record.stm_date);
                Rec.InsertString("stm_accid", Record.stm_accid);
                Rec.InsertString("stm_acc_br_id", Record.stm_acc_br_id);
                Rec.InsertString("stm_currencyid", Record.stm_currencyid);

                Rec.InsertNumeric("stm_dr", Record.stm_dr.ToString());
                Rec.InsertNumeric("stm_cr", Record.stm_cr.ToString());
                Rec.InsertNumeric("stm_bal", Record.stm_bal.ToString());


                Rec.InsertNumeric("stm_dr_inr", Record.stm_dr_inr.ToString());
                Rec.InsertNumeric("stm_cr_inr", Record.stm_cr_inr.ToString());
                Rec.InsertNumeric("stm_bal_inr", Record.stm_bal_inr.ToString());


                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                sql = "delete from stmtd where std_parentid = '" + Record.stm_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                foreach (Stmtd Row in Record.PendingList)
                {
                    if ( Row.allocation > 0)
                    {
                        Rec = new DBRecord();
                        Rec.CreateRow("stmtd", "ADD", "std_pkid", System.Guid.NewGuid().ToString().ToUpper());
                        Rec.InsertString("std_parentid", Record.stm_pkid);
                        Rec.InsertString("std_jv_entityid", Row.jv_entity_id);
                        Rec.InsertString("std_jv_pkid", Row.jv_pk_id);
                        Rec.InsertString("std_accid", Row.jv_ac_rowid);
                        Rec.InsertString("std_currencyid", Row.jv_currency_rowid);
                        Rec.InsertNumeric("std_amt", Row.allocation.ToString());
                        Rec.InsertNumeric("std_amt_inr", Row.inrallocation.ToString());
                        Rec.InsertDate("std_jv_date", Row.jv_date);

                        Rec.InsertString("rec_deleted", "N");
                        Rec.InsertString("rec_locked", "N");

                        Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                        Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);

                        sql = Rec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                }

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                string str = "STMT # " + sm_no;
                Lib.AuditLog("STMT", "", Record.rec_mode, Record._globalvariables.comp_code, Record._globalvariables.branch_code, Record._globalvariables.user_code, Record.stm_pkid, Record.stm_no.ToString(), str);
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

            RetData.Add("stm_no", sm_no);

            return RetData;
        }

        public Dictionary<string, object> PrintList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string id = "";
            string type = "";


            string folderid = "";
            string branch_code = "";

            string report_folder = "";
            string File_Name = "";
            string File_Type = "";
            string File_Display_Name = "stmt.xls";

            string category = "";
            string comp_code = "";

            Color _Color = Color.Black;
            int _Size = 11;


            try
            {
                id = SearchData["pkid"].ToString();
                
                if (SearchData.ContainsKey("type"))
                    type = SearchData["type"].ToString();

                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();

                if (SearchData.ContainsKey("folderid"))
                    folderid = SearchData["folderid"].ToString();

                if (SearchData.ContainsKey("comp_code"))
                    comp_code = SearchData["comp_code"].ToString();

                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();


                string SQL = "";

                Con_Oracle = new DBConnection();


                Dictionary<string, object> datalist = GetPendingList(SearchData);

                mList = (List<Stmtd> )datalist["list"];

                sql = "";
                sql += "  select stm_pkid,stm_no,stm_date,acc_code, acc_name, c.param_code as curr_code,";
                sql += " stm_dr, stm_cr, stm_bal, stm_dr_inr, stm_cr_inr, stm_bal_inr ";
                sql += " from stmtm a ";
                sql += " left join acctm b on a.stm_accid = b.acc_pkid ";
                sql += " left join param c on a.stm_currencyid = c.param_pkid ";
                sql += " where stm_pkid = '" + id + "'";

                Dt_List = Con_Oracle.ExecuteQuery(sql);


                if (type == "EXCEL")
                {
                    File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                    File_Type = "EXCEL";
                    PrintExcelList(comp_code, branch_code, File_Name, category);
                }

                mList.Clear();
                mList = null;

                

            }


            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }

        private void PrintExcelList( string comp_code, string branch_code, string FileName, string category)
        {
            string _Border = "";
            string curr_code = "";

            

            Color _Color = Color.Black;
            int _Size = 11;

            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string sName = "Report";


            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];

            WS.Columns[0].Width = 256 * 3;


            WS.Columns[1].Width = 256 * 14;
            WS.Columns[2].Width = 256 * 25;
            WS.Columns[3].Width = 256 * 15;
            WS.Columns[4].Width = 256 * 10;
            WS.Columns[5].Width = 256 * 7;
            WS.Columns[6].Width = 256 * 8;

            WS.Columns[7].Width = 256 * 8;
            WS.Columns[8].Width = 256 * 8;
            WS.Columns[9].Width = 256 * 8;
            WS.Columns[10].Width = 256 * 8;
            WS.Columns[11].Width = 256 * 8;

            WS.Columns[12].Width = 256 * 10;
            WS.Columns[13].Width = 256 * 10;
            WS.Columns[14].Width = 256 * 10;
            WS.Columns[15].Width = 256 * 10;

            WS.Columns[7].Style.NumberFormat = "#,0.00";
            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";
            WS.Columns[10].Style.NumberFormat = "#,0.00";


            WS.Columns[11].Style.NumberFormat = "#,0.000";
            WS.Columns[12].Style.NumberFormat = "#,0.00";
            WS.Columns[13].Style.NumberFormat = "#,0.00";
            WS.Columns[14].Style.NumberFormat = "#,0.00";
            WS.Columns[15].Style.NumberFormat = "#,0.00";




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
                    break;
                }
            }

            iRow = 1; iCol = 1;
            _Color = Color.Black;
            _Size = 16;

            Lib.WriteMergeCell(WS, iRow, 1, 15, 1, COMPNAME, "Calibri", 15, true, Color.Black, "C", "C", "", "");
            iRow++;
            _Size = 15;
            Lib.WriteMergeCell(WS, iRow, 1, 15, 1, COMPADD1, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            Lib.WriteMergeCell(WS, iRow, 1, 15, 1, COMPADD2, "Calibri", 12, false, Color.Black, "C", "C", "", "");
            iRow++;
            string str = "";
            if (COMPTEL.Trim() != "")
                str = "TEL : " + COMPTEL;
            if (COMPFAX.Trim() != "")
                str += " FAX : " + COMPFAX;
            Lib.WriteMergeCell(WS, iRow, 1, 15, 1, str.Trim(), "Calibri", 12, false, Color.Black, "C", "C", "", "");


            iRow++;
            WS.Cells.GetSubrangeRelative(iRow, 1, 15, 1).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
            _Size = 14;
            iRow++;
            Lib.WriteMergeCell(WS, iRow, 1, 15, 2, "COSTING - ALLOCATION", "Calibri", 12, true, Color.Black, "C", "C", "TB", "THIN");

            iCol = 1; iRow++; _Size = 11;
            iRow++;

            DateTime Dt;
            string sDate = "";


            str = Dt_List.Rows[0]["stm_no"].ToString();
            Lib.WriteData(WS, iRow, 1, "Stmt#", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2,  str , _Color, false, _Border, "L", "", _Size, false, 325, "", true);

            iRow++;

            str = Dt_List.Rows[0]["acc_code"].ToString() ;
            Lib.WriteData(WS, iRow, 1, "AGENT", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            str = Dt_List.Rows[0]["acc_name"].ToString();
            Lib.WriteData(WS, iRow, 3, str, _Color, false, _Border, "L", "", _Size, false, 325, "", true);


            iRow++;

            curr_code = Dt_List.Rows[0]["curr_code"].ToString();
            Lib.WriteData(WS, iRow, 1, "CURRENCY :", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, curr_code, _Color, false, _Border, "L", "", _Size, false, 325, "", true);

            _Color = Color.Black;
            _Border = "TBLR";
            _Size = 11;
            iCol = 1; iRow++;

            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "REF#", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, false, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "DR-" + curr_code, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CR-" + curr_code, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BAL-" + curr_code, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "ALLOC-" + curr_code, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "EXRATE", _Color, false, _Border, "R", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "DR-INR", _Color, false, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CR-INR", _Color, false, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BAL-INR", _Color, false, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "ALLOC-INR", _Color, false, _Border, "R", "", _Size, false, 325, "", true);





            foreach (Stmtd Row  in mList)
            {
                if (Row.allocation != 0)
                {
                    iRow++; iCol = 1;
                    _Border = "TBLR";
                    _Color = Color.Black;
                    Lib.WriteData(WS, iRow, iCol++, Lib.StringToDate(Row.jv_date), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_reference, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_remarks, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_vrno, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_type, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.rec_category, _Color, false, _Border, "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Row.dr, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.cr, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.balance, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.allocation, _Color, false, _Border, "R", "", _Size, false, 325, "", true);


                    Lib.WriteData(WS, iRow, iCol++, Row.jv_exchange_rate, _Color, false, _Border, "R", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Row.jv_debit, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_credit, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.inrbalance, _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.inrallocation, _Color, false, _Border, "R", "", _Size, false, 325, "", true);


                }
            }

            iRow+=2;



            Lib.WriteData(WS, iRow, 1, "PARTICULARS", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, "DEBIT", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, "CREDIT", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 4, "BALANCE", _Color, false, _Border, "L", "", _Size, false, 325, "", true);

            iRow++;
            Lib.WriteData(WS, iRow, 1, Dt_List.Rows[0]["curr_code"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, Dt_List.Rows[0]["stm_dr"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, Dt_List.Rows[0]["stm_cr"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 4, Dt_List.Rows[0]["stm_bal"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);

            iRow++;
            Lib.WriteData(WS, iRow, 1, "INR", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, Dt_List.Rows[0]["stm_dr_inr"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 3, Dt_List.Rows[0]["stm_cr_inr"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 4, Dt_List.Rows[0]["stm_bal_inr"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);


            WB.SaveXls(FileName);

        }
      

        public Dictionary<string, object> GetCostOs(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();


            string report_folder = "";
            string File_Name = "";
            string File_Type = "";
            string File_Display_Name = "os.xls";
            string folderid = "";


            decimal fDr = 0;
            decimal fCr = 0;
            decimal fBal = 0;

            decimal lDr = 0;
            decimal lCr = 0;
            decimal lBal = 0;

            Con_Oracle = new DBConnection();

            string type = SearchData["type"].ToString();
            string accid = SearchData["accid"].ToString();
            string currid = SearchData["currid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string agent = SearchData["agent"].ToString();


            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();

            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();



            string stm_date = SearchData["stm_date"].ToString();



            string todate = Lib.StringToDate(stm_date);

            mList = new List<Stmtd>();
            Stmtd cRow;

            string sql = "";

            try
            {

                sql += " select * from  (  ";
                sql += " 	   select  jvh_pkid,jvh_year,  jvh_vrno,  jvh_type,  jvh_date,  nvl(jvh_location,'OTHERS') as jvh_location,  jvh_remarks,  jvh_reference,  hbl_folder_no, ";
                sql += " 	   jv_ftotal as amount,  ";
                sql += " 	   case when jv_debit  <>0 then jv_ftotal else 0 end as DR,  ";
                sql += " 	   case when jv_credit <>0 then jv_ftotal else 0 end as CR,  ";
                sql += " 	   jv_ftotal - nvl(xref_Amt,0) as balance,  ";
                sql += " 	   jv_exrate, (jv_ftotal - nvl(xref_Amt,0))*jv_exrate as inrbalance, a.rec_category  ";
                sql += " 	   from ledgerh a inner join  ledgert  d on a.jvh_pkid = d.jv_parent_id   ";
                sql += "       left join hblm on jvh_cc_id = hbl_pkid ";
                sql += " 	   left join  (    ";
                sql += " 	   	   select std_jv_pkid,std_currencyid,sum(std_amt) as xref_Amt    ";
                sql += " 		   from	  stmtd    ";
                sql += " 		   where  std_accid =  '{ACCID}'    ";
                sql += " 		   and std_jv_date<='{DATE}'       ";
                sql += " 		   and std_currencyid='{CURRENCY}'    ";
                sql += " 		   and rec_company_Code= '{COMPANY}'    ";
                sql += " 		   and rec_branch_code= '{BRANCH}'    ";
                sql += " 		   group by std_jv_pkid , std_currencyid   ";
                sql += " 	   )  b  on  d.jv_pkid = b.std_jv_pkid and d.jv_curr_id=b.std_currencyid  ";
                sql += " 	   where  jv_acc_id = '{ACCID}' ";
                sql += " 	   and jv_curr_id = '{CURRENCY}'  ";
                sql += " 	   and jvh_type not in ('OP','OB')  and a.rec_deleted = 'N'  and nvl(jvh_reference,'A') <> 'COSTING ADJUSTMENT'  ";
                sql += " 	   and a.rec_company_code= '{COMPANY}'  and a.rec_branch_code= '{BRANCH}' ";
                sql += " )  jv where  (Balance) != 0   and jvh_date<='{DATE}' ";
                sql += " order by  case when jvh_location is null then '1' || jvh_location   else  case when jvh_location ='OTHERS' then '1'|| jvh_location  else jvh_location  end end, ";

                if( type == "SCREEN") 
                    sql += " jvh_location, jvh_date, jvh_reference ";
                else
                    sql += " jvh_location, jvh_date, jvh_reference ";


                sql = sql.Replace("{ACCID}", accid);
                sql = sql.Replace("{CURRENCY}", currid);
                sql = sql.Replace("{COMPANY}", comp_code);
                sql = sql.Replace("{BRANCH}", branch_code);
                sql = sql.Replace("{DATE}", todate);

                DataTable Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql.ToString());
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    cRow = new Stmtd();
                    cRow.rowtype = "";
                    cRow.jv_entity_id = Dr["jvh_pkid"].ToString();
                    cRow.jv_vrno = Lib.Conv2Integer(Dr["jvh_vrno"].ToString());
                    cRow.jv_type = Dr["jvh_type"].ToString();
                    cRow.jv_date = Lib.DatetoString(Dr["jvh_date"]);
                    cRow.jv_display_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    cRow.jv_reference = Dr["jvh_reference"].ToString();
                    cRow.jv_remarks = Dr["jvh_remarks"].ToString();
                    cRow.jv_location = Dr["jvh_location"].ToString();
                    cRow.rec_category = Dr["rec_category"].ToString();

                    cRow.folderno = Dr["hbl_folder_no"].ToString();

                    cRow.amount = Lib.Conv2Decimal(Dr["amount"].ToString());
                    cRow.jv_exchange_rate = Lib.Conv2Decimal(Dr["jv_exrate"].ToString());

                    cRow.dr = Lib.Conv2Decimal(Dr["dr"].ToString());
                    cRow.cr = Lib.Conv2Decimal(Dr["cr"].ToString());


                    cRow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    if (cRow.dr > 0)
                    {
                        cRow.dr = cRow.balance;
                        fDr += cRow.dr;
                    }
                    else
                    {
                        cRow.cr = Math.Abs(cRow.balance);
                        fCr += cRow.cr;
                    }

                    cRow.jv_exchange_rate = Lib.Conv2Decimal(Dr["jv_exrate"].ToString());

                    cRow.inrbalance = Lib.Conv2Decimal(Dr["inrbalance"].ToString());
                    if (cRow.dr > 0)
                    {
                        cRow.jv_debit = cRow.inrbalance;
                        lDr += cRow.jv_debit;
                    }
                    else
                    {
                        cRow.jv_credit = Math.Abs(cRow.inrbalance);
                        lCr += cRow.jv_credit;
                    }

                    mList.Add(cRow);
                }
                fBal = fDr - fCr;
                lBal = lDr - lCr;


                cRow = new Stmtd();
                cRow.rowtype = "TOTAL";
                cRow.jv_reference = "TOTAL";

                if (fBal > 0)
                {
                    cRow.cr = fBal;
                    cRow.jv_reference = "NET DUE FROM " + agent;
                    fCr += cRow.cr;
                }
                if (fBal < 0)
                {
                    cRow.jv_reference = "NET DUE TO AGENT " + agent;
                    cRow.dr = Math.Abs(fBal);
                    fDr += cRow.dr;
                }

                if (lBal > 0)
                {
                    cRow.jv_credit = lBal;
                    lCr += cRow.jv_credit;
                }
                if (lBal < 0)
                {
                    cRow.jv_debit = Math.Abs(lBal);
                    lDr += cRow.jv_debit;
                }
                mList.Add(cRow);


                cRow = new Stmtd();
                cRow.rowtype = "TOTAL";
                cRow.jv_reference = "TOTAL";
                cRow.dr = fDr;
                cRow.cr = fCr;
                cRow.jv_debit = lDr;
                cRow.jv_credit = lCr;

                mList.Add(cRow);
                Dt_Rec.Rows.Clear();

                if (type == "EXCEL")
                {
                    File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                    File_Type = "EXCEL";
                    PrintOsExcel(comp_code, branch_code, File_Name, todate);

                    mList.Clear();
                    mList = null;
                }

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("list", mList);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);

            return RetData;
        }

        private void PrintOsExcel(string comp_code, string branch_code, string FileName, string todate)
        {
            string _Border = "";



            Color _Color = Color.Black;
            int _Size = 11;

            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string sName = "Report";


            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];

            WS.Columns[0].Width = 256 * 5;

            WS.Columns[1].Width = 256 * 25;
            WS.Columns[2].Width = 256 * 15;
            WS.Columns[3].Width = 256 * 25;
            WS.Columns[4].Width = 256 * 15;
            WS.Columns[5].Width = 256 * 15;
            WS.Columns[6].Width = 256 * 15;
            WS.Columns[7].Width = 256 * 10;
            WS.Columns[8].Width = 256 * 15;
            WS.Columns[9].Width = 256 * 15;
            WS.Columns[10].Width = 256 * 15;

            WS.Columns[4].Style.NumberFormat = "#,0.00";
            WS.Columns[5].Style.NumberFormat = "#,0.00";
            WS.Columns[7].Style.NumberFormat = "#,0.00000";
            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";


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
                    break;
                }
            }

            iRow = 1; iCol = 1;
            _Color = Color.Black;
            _Size = 12;

            Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, true, "", "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, true, "", "L", "", _Size, false, 325, "", true);
            iRow++;
            string str = "";
            if (COMPTEL.Trim() != "")
                str = "TEL : " + COMPTEL;
            if (COMPFAX.Trim() != "")
                str += " FAX : " + COMPFAX;

            Lib.WriteData(WS, iRow, 1, str, _Color, true, "", "L", "", _Size, false, 325, "", true);
            iRow++;
            str = "Email: hocosting@cargomar.in";
            Lib.WriteData(WS, iRow, 1, str, _Color, true, "", "L", "", _Size, false, 325, "", true);

            iRow++;
            WS.Cells.GetSubrangeRelative(iRow, 1, 10, 1).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
            _Size = 12;

            iCol = 1; iRow++; _Size = 11;

            Lib.WriteData(WS, iRow, 1, "STATEMENT OF ACCOUNTS", _Color, true, _Border, "L", "", 12, false, 325, "", true);

            iRow++;

            _Color = Color.Black;
            _Border = "TBLR";
            _Size = 11;
            iCol = 1; iRow++;

            
            Lib.WriteData(WS, iRow, iCol++, "CN / DN NOTE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "PARTICULARS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DEBIT" , _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE" , _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "EXRATE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "FOLDER#", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            str = "";

            foreach (Stmtd Row in mList)
            {
                    iRow++; iCol = 1;
                    _Border = "TBLR";
                _Border = "";
                _Color = Color.Black;

                if (str != Row.jv_location)
                {
                    if (Row.rowtype != "TOTAL")
                    {
                        str = Row.jv_location;
                        iRow++;
                        Lib.WriteData(WS, iRow, iCol++, Row.jv_location, _Color, false, "U", "L", "", _Size, false, 325, "", true);
                        iRow+=2;
                        iCol = 1;
                    }
                }

                if (Row.rowtype != "TOTAL")
                {
                    

                    Lib.WriteData(WS, iRow, iCol++, Row.jv_reference, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Row.jv_display_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, _Border, "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_remarks, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.dr, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Row.cr, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Row.rec_category, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_exchange_rate, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_debit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_credit, _Color, false, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Row.folderno, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                }
                else
                {
                    iRow+=2;

                    WS.Cells.GetSubrangeRelative(iRow, 1, 10, 1).SetBorders(MultipleBorders.Top, Color.Black, LineStyle.Thin);

                    Lib.WriteData(WS, iRow, iCol++, Row.jv_reference, _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                    iCol++;
                    iCol++;
                    Lib.WriteData(WS, iRow, iCol++, Row.dr, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Row.cr, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_debit, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Row.jv_credit, _Color, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


                    iRow++;
                    WS.Cells.GetSubrangeRelative(iRow, 1, 10, 1).SetBorders(MultipleBorders.Top, Color.Black, LineStyle.Thin);
                }

            }

            iRow += 2;





            WB.SaveXls(FileName);

        }




    }
}