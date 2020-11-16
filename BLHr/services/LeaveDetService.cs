using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLHr
{
    public class LeaveDetService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Leavem> mList = new List<Leavem>();
            Leavem mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            int levmonth = Lib.Conv2Integer(SearchData["levmonth"].ToString());
            int levyear = Lib.Conv2Integer(SearchData["levyear"].ToString());
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where a.rec_company_code = '" + company_code + "'";
                sWhere += " and a.rec_branch_code = '" + branch_code + "'";
                sWhere += " and a.lev_year = " + levyear.ToString();
                if (levmonth != 0)
                    sWhere += " and a.lev_month = " + levmonth.ToString();
                else
                    sWhere += " and a.lev_month > 0 ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(b.emp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  b.emp_no like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " ) ";
                }
               
                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM Leavem  a ";
                    sql += " left join empm b on a.lev_emp_id = b.emp_pkid ";
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
                sql += " select lev_pkid,lev_emp_id,lev_year,lev_month,lev_sl,lev_cl,lev_pl,lev_lp,lev_others ";//case when (a.rec_category ='CONFIRMED' or a.rec_category='TRANSFER') then  lev_lp else 0 end as 
                sql += " ,emp_no,emp_name ";
                sql += " ,row_number() over(order by emp_no,lev_year,lev_month) rn ";
                sql += " from leavem a ";
                sql += " left join empm b on a.lev_emp_id = b.emp_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by emp_no,lev_year,lev_month";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Leavem();
                    mRow.lev_pkid = Dr["lev_pkid"].ToString();
                    mRow.lev_emp_id = Dr["lev_emp_id"].ToString();
                    mRow.lev_emp_code = Dr["emp_no"].ToString();
                    mRow.lev_emp_name = Dr["emp_name"].ToString();
                    mRow.lev_year = Lib.Conv2Integer(Dr["lev_year"].ToString());
                    mRow.lev_month = Lib.Conv2Integer(Dr["lev_month"].ToString());
                    mRow.lev_sl = Lib.Conv2Decimal(Dr["lev_sl"].ToString());
                    mRow.lev_cl = Lib.Conv2Decimal(Dr["lev_cl"].ToString());
                    mRow.lev_pl = Lib.Conv2Decimal(Dr["lev_pl"].ToString());
                    mRow.lev_lp = Lib.Conv2Decimal(Dr["lev_lp"].ToString());
                    mRow.lev_others = Lib.Conv2Decimal(Dr["lev_others"].ToString());
                    mList.Add(mRow);
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);

            return RetData;
        }
        
        public Dictionary<string, object>  GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Leavem mRow =new Leavem();
            string id = SearchData["pkid"].ToString();
            try
            {
                DataTable Dt_Rec = new DataTable();
     
                sql = " select lev_pkid,lev_emp_id,lev_year,lev_month, ";
                sql += " lev_sl,lev_cl,lev_pl,lev_others,lev_lp, ";
                sql += " lev_holidays,lev_days_worked ,lev_pl_carry,lev_fin_year, ";
                sql += " b.emp_no as lev_emp_code, b.emp_name as  lev_emp_name,a.rec_category,lev_edit_code ";
                sql += " from leavem a  ";
                sql += " left join empm b on a.lev_emp_id = b.emp_pkid ";
                sql += " where  a.lev_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Leavem();
                    mRow.lev_pkid = Dr["lev_pkid"].ToString();
                    mRow.lev_emp_id = Dr["lev_emp_id"].ToString();
                    mRow.lev_emp_code = Dr["lev_emp_code"].ToString();
                    mRow.lev_emp_name = Dr["lev_emp_name"].ToString();
                    mRow.lev_year = Lib.Conv2Integer(Dr["lev_year"].ToString());
                    mRow.lev_month = Lib.Conv2Integer(Dr["lev_month"].ToString());
                    mRow.lev_sl = Lib.Conv2Decimal(Dr["lev_sl"].ToString());
                    mRow.lev_cl = Lib.Conv2Decimal(Dr["lev_cl"].ToString());
                    mRow.lev_pl = Lib.Conv2Decimal(Dr["lev_pl"].ToString());
                    mRow.lev_others = Lib.Conv2Decimal(Dr["lev_others"].ToString());
                    mRow.lev_lp = Lib.Conv2Decimal(Dr["lev_lp"].ToString());
                    mRow.lev_holidays = Lib.Conv2Decimal(Dr["lev_holidays"].ToString());
                    mRow.lev_days_worked = Lib.Conv2Decimal(Dr["lev_days_worked"].ToString());
                    mRow.lev_pl_carry = Lib.Conv2Decimal(Dr["lev_pl_carry"].ToString());
                    mRow.lev_fin_year = Lib.Conv2Integer(Dr["lev_fin_year"].ToString());
                    mRow.rec_category = Dr["rec_category"].ToString();
                    mRow.lev_edit_code = Dr["lev_edit_code"].ToString();
                    break;
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(Leavem Record)
        { 
            string str = "";
            DateTime tdate = DateTime.Now;
            decimal TotCL = 0, TotSL = 0, TotPL = 0;
            decimal TotCL_Tkn = 0, TotSL_Tkn = 0, TotPL_Tkn = 0;
            try
            {
                if (Record.lev_emp_id.Trim().Length <= 0)
                    Lib.AddError(ref str, " | Employee Cannot Be Empty");

                sql = "select lev_pkid from (";
                sql += "select lev_pkid from leavem a where a.rec_company_code = '{COMPCODE}' ";
                sql += " and a.lev_emp_id = '{EMP_ID}' ";
                sql += " and a.lev_month = {LEVMONTH} ";
                sql += ") a where lev_pkid <> '{PKID}'";
                sql = sql.Replace("{EMP_ID}", Record.lev_emp_id);
                sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{LEVMONTH}", Record.lev_month.ToString());
                sql = sql.Replace("{PKID}", Record.lev_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " | Leave Details for this month already Exists");


                sql = " select lev_pkid from leavem ";
                sql += " where lev_pkid = '" + Record.lev_pkid + "'";
                sql += " and lev_edit_code is null";
                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " Details Closed, Can't Edit ");

                if (Record.lev_lp > 0)
                {
                    sql = " select sal_pkid from salarym a ";
                    sql += " where a.rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    sql += " and sal_emp_id = '" + Record.lev_emp_id + "'";
                    sql += " and sal_month = "+ Record.lev_month.ToString();
                    sql += " and sal_year = " + Record.lev_year;
                    sql += " and sal_edit_code is null";
                    if (Con_Oracle.IsRowExists(sql))
                        Lib.AddError(ref str, " Payroll Closed, Can't Edit ");

                    sql = " select salh_pkid from salaryh ";
                    sql += " where rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    sql += " and rec_branch_code = '" + Record._globalvariables.branch_code + "'";
                    sql += " and salh_year = " + Record.lev_year;
                    sql += " and salh_month = " + Record.lev_month;
                    sql += " and nvl(salh_posted, 'N') = 'Y' ";
                    if (Con_Oracle.IsRowExists(sql))
                        Lib.AddError(ref str, " | Payroll Already Posted, Can't Edit ");

                }

                sql = " select nvl(lev_pl,0) + nvl(lev_pl_carry,0) as lev_pl ,lev_sl,lev_cl";
                sql += " ,lev_tot_pl as Tot_Pl_Taken ,lev_tot_sl as Tot_sl_Taken, lev_tot_cl as Tot_cl_Taken";
                sql += " from leavem a ";
                sql += " where a.rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and lev_emp_id = '" + Record.lev_emp_id + "'";
                sql += " and lev_month = 0 ";
                sql += " and lev_fin_year = " + Record._globalvariables.year_code;
                DataTable Dt_Lev = new DataTable();
                Dt_Lev = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Lev.Rows.Count > 0)
                {
                    TotCL = Lib.Convert2Decimal(Dt_Lev.Rows[0]["lev_cl"].ToString());
                    TotSL = Lib.Convert2Decimal(Dt_Lev.Rows[0]["lev_sl"].ToString());
                    TotPL = Lib.Convert2Decimal(Dt_Lev.Rows[0]["lev_pl"].ToString());
                    TotCL_Tkn = Lib.Convert2Decimal(Dt_Lev.Rows[0]["Tot_cl_Taken"].ToString());
                    TotSL_Tkn = Lib.Convert2Decimal(Dt_Lev.Rows[0]["Tot_sl_Taken"].ToString());
                    TotPL_Tkn = Lib.Convert2Decimal(Dt_Lev.Rows[0]["Tot_Pl_Taken"].ToString());

                    if ((Lib.Convert2Decimal(Record.lev_pl.ToString()) + TotPL_Tkn) > TotPL)
                    {
                        Lib.AddError(ref str, " | Privilege leave exceed the limlit (Tkn=" + TotPL_Tkn.ToString() + ", Bal = " + (TotPL - TotPL_Tkn).ToString() + ", Lmt=" + TotPL.ToString() + ")");
                    }
                    if ((Lib.Convert2Decimal(Record.lev_cl.ToString()) + TotCL_Tkn) > TotCL)
                    {
                        Lib.AddError(ref str, " | Casual leave exceed the limlit (Tkn=" + TotCL_Tkn.ToString() +", Bal = " + (TotCL- TotCL_Tkn).ToString() + ", Lmt=" + TotCL.ToString() + ")");
                    }
                    if ((Lib.Convert2Decimal(Record.lev_sl.ToString()) + TotSL_Tkn) > TotSL)
                    {
                        Lib.AddError(ref str, " | Sick leave exceed the limlit (Tkn=" + TotSL_Tkn.ToString() + ", Bal = " + (TotSL - TotSL_Tkn).ToString() + ", Lmt=" + TotSL.ToString() + ")");
                    }
                }

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }
        
        public Dictionary<string, object> Save(Leavem Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();
                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);

                DateTime dtime = new DateTime(Record.lev_year, Record.lev_month, 1);

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("Leavem", Record.rec_mode, "lev_pkid", Record.lev_pkid);
                Rec.InsertString("lev_emp_id", Record.lev_emp_id);
                Rec.InsertNumeric("lev_year", Record.lev_year.ToString());
                Rec.InsertNumeric("lev_month", Record.lev_month.ToString());
                Rec.InsertDate("lev_date", dtime.ToString(Lib.BACK_END_DATE_FORMAT));
                Rec.InsertNumeric("lev_sl", Record.lev_sl.ToString());
                Rec.InsertNumeric("lev_cl", Record.lev_cl.ToString());
                Rec.InsertNumeric("lev_pl", Record.lev_pl.ToString());
                Rec.InsertNumeric("lev_others", Record.lev_others.ToString());
                Rec.InsertNumeric("lev_lp", Record.lev_lp.ToString());
                Rec.InsertNumeric("lev_holidays", Record.lev_holidays.ToString());
                Rec.InsertNumeric("lev_days_worked", Record.lev_days_worked.ToString());
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("lev_edit_code", "{S}");
                    Rec.InsertNumeric("lev_fin_year", Record._globalvariables.year_code);
                    Rec.InsertString("rec_category", Record.rec_category);
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
                UpdateLeaveSummary(Record.lev_emp_id, Lib.Conv2Integer(Record._globalvariables.year_code.ToString()), Record._globalvariables.comp_code);
                if (Record.lev_lp > 0)
                   Lib.FindLoPAmount(Record.lev_emp_id, Record.lev_year, Record.lev_month, Record.lev_lp, Record.lev_days_worked);
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
            //parameter.Add("param_type", "COUNTRY");
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }
       
        private void UpdateLeaveSummary(string EMP_ID, int FinYear,string CompanyCode)
        {
            DataTable Dt_Lev;
            DataRow Dr = null;
          
            try
            {

                sql = "   select sum(lev_pl) as lev_pl,sum(lev_cl) as lev_cl";
                sql += "  ,sum(lev_sl) as lev_sl,sum(lev_others) as lev_others";
                sql += "  from leavem a ";
                sql += "   where rec_company_code = '" + CompanyCode + "'";
                sql += "   and lev_emp_id = '" + EMP_ID + "'";
                sql += "   and lev_month <> 0 ";
                sql += "   and lev_fin_year = " + FinYear.ToString();
                Dt_Lev = new DataTable();
                Dt_Lev = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Lev.Rows.Count > 0)
                {
                    Dr = Dt_Lev.Rows[0];

                    sql = " Update leavem set lev_tot_pl = " + Lib.Conv2Decimal(Dr["lev_pl"].ToString());
                    sql += " ,lev_tot_cl = " + Lib.Conv2Decimal(Dr["lev_cl"].ToString());
                    sql += " ,lev_tot_sl = " + Lib.Conv2Decimal(Dr["lev_sl"].ToString());
                    sql += " ,lev_tot_others = " + Lib.Conv2Decimal(Dr["lev_others"].ToString());
                    sql += " where rec_company_code = '" + CompanyCode + "'";
                    sql += " and lev_emp_id = '" + EMP_ID + "'";
                    sql += " and lev_month = 0 ";
                    sql += " and lev_fin_year =" + FinYear;

                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                }
                
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
    }
}
