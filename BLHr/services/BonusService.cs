using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using System.Drawing;
using XL.XSheet;

namespace BLHr
{
    public class BonusService : BL_Base
    {
        ExcelFile file;
        ExcelWorksheet ws = null;
        ExcelWorksheet ws2 = null;
        CellRange myCell;
        string report_folder = "";
        string folderid = "";
        string File_Name = "";
        string File_Type = "";
        string File_Display_Name = "myreport.pdf";
        int iCol = 0;
        int iRow = 0;
       

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
            //string Sal_Emp_ID = "";
            //string Sal_Emp_Code = "";
            //int PrevFin_Year = 0;
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            DataTable Dt_Temp;
            Con_Oracle = new DBConnection();
            List<Bonusm> mList = new List<Bonusm>();
            Bonusm mRow;
            bool brelived = false;
            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string company_name = SearchData["company_name"].ToString();
            string year_code = SearchData["year_code"].ToString();
            report_folder = "";
            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();
            folderid = "";
            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();

            brelived = (bool)SearchData["brelived"];

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;
            try
            {

                //PrevFin_Year = Lib.Conv2Integer(year_code);
                //PrevFin_Year -= 1;

                //sql = "select distinct sal_emp_id from salarym ";
                //sql += " where rec_company_code = '" + company_code + "'";
                //sql += " and rec_branch_code = '" + branch_code + "'";
                //sql += " and sal_fin_year =" + PrevFin_Year.ToString();
                //Dt_Temp = new DataTable();
                //Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                //Sal_Emp_ID = "";
                //foreach (DataRow dr in Dt_Temp.Rows)
                //{
                //    if (Sal_Emp_ID != "")
                //        Sal_Emp_ID += ",";
                //    Sal_Emp_ID +=  dr["sal_emp_id"].ToString() ;
                //}

                //if(Sal_Emp_ID=="")
                //{
                //    sql = " select distinct cc_code from costcentert a";
                //    sql += " inner join costcenterm b on a.ct_cost_id = b.cc_pkid and a.ct_category ='EMPLOYEE' ";
                //    sql += " where a.rec_company_code = '" + company_code + "'";
                //    sql += " and a.rec_branch_code = '" + branch_code + "'";
                //    sql += " and ct_cost_year = " + PrevFin_Year.ToString();
                //    Dt_Temp = new DataTable();
                //    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                //    Sal_Emp_Code = "";
                //    foreach (DataRow dr in Dt_Temp.Rows)
                //    {
                //        if (Sal_Emp_Code != "")
                //            Sal_Emp_Code += ",";
                //        Sal_Emp_Code += dr["cc_code"].ToString();
                //    }
                //    Sal_Emp_Code = Sal_Emp_Code.Replace(",", "','");
                //}

                //Sal_Emp_ID = Sal_Emp_ID.Replace(",", "','");

                sWhere = " where a.rec_company_code = '" + company_code + "'";
                sWhere += " and a.rec_branch_code = '" + branch_code + "'";
                sWhere += " and a.bon_fin_year = " +(Lib.Conv2Integer(year_code)-1).ToString();
                if (brelived)
                    sWhere += " and b.emp_in_payroll != 'Y' ";
                else
                    sWhere += " and b.emp_in_payroll = 'Y' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += " upper(b.emp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += " b.emp_no like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " ) ";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM bonusm  a ";
                    sql += " inner join empm b on (a.bon_emp_id = b.emp_pkid) ";
                    sql += sWhere;
                    Dt_Temp = new DataTable();
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
                sql += " select bon_pkid,emp_pkid,emp_no,emp_name,bon_days_worked ";
                sql += " ,bon_gross_wages ,bon_gross_bonus ,bon_puja_deduct ";
                sql += " ,bon_interim_deduct,bon_tax_deduct,bon_other_deduct";
                sql += " ,bon_tot_deduct,bon_net_amount,bon_actual_paid "; ;
                sql += " ,bon_paid_date ,bon_remarks ";
                sql += " ,row_number() over (order by emp_no) rn ";
                sql += " from bonusm a ";
                sql += " inner join empm b on (a.bon_emp_id = b.emp_pkid) ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by emp_no";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Bonusm();
                    mRow.bon_pkid = Dr["bon_pkid"].ToString();
                    mRow.bon_emp_id = Dr["emp_pkid"].ToString();
                    mRow.bon_emp_code = Dr["emp_no"].ToString();
                    mRow.bon_emp_name = Dr["emp_name"].ToString();
                    mRow.bon_days_worked = Lib.Conv2Integer(Dr["bon_days_worked"].ToString());
                    mRow.bon_gross_wages = Lib.Conv2Decimal(Dr["bon_gross_wages"].ToString());
                    mRow.bon_gross_bonus = Lib.Conv2Decimal(Dr["bon_gross_bonus"].ToString());
                    mRow.bon_puja_deduct = Lib.Conv2Decimal(Dr["bon_puja_deduct"].ToString());
                    mRow.bon_interim_deduct = Lib.Conv2Decimal(Dr["bon_interim_deduct"].ToString());
                    mRow.bon_tax_deduct = Lib.Conv2Decimal(Dr["bon_tax_deduct"].ToString());
                    mRow.bon_other_deduct = Lib.Conv2Decimal(Dr["bon_other_deduct"].ToString());
                    mRow.bon_tot_deduct = Lib.Conv2Decimal(Dr["bon_tot_deduct"].ToString());
                    mRow.bon_net_amount = Lib.Conv2Decimal(Dr["bon_net_amount"].ToString());
                    mRow.bon_actual_paid = Lib.Conv2Decimal(Dr["bon_actual_paid"].ToString());
                    mRow.bon_paid_date = Lib.DatetoStringDisplayformat(Dr["bon_paid_date"]);
                    mRow.bon_remarks = Dr["bon_remarks"].ToString();
                    mList.Add(mRow);
                }
                if (type == "EXCEL")
                {
                    PrintBonusLetter(company_name, company_code, branch_code, year_code, brelived);
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
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);

            return RetData;
        }

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Bonusm mRow = new Bonusm();
            string id = SearchData["pkid"].ToString();
            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select bon_pkid,bon_emp_id,bon_fin_year,bon_days_worked";
                sql += " ,bon_gross_wages ,bon_gross_bonus ,bon_puja_deduct,bon_interim_deduct";
                sql += " ,bon_tax_deduct,bon_other_deduct,bon_tot_deduct,bon_net_amount";
                sql += " ,bon_actual_paid,bon_paid_date,bon_remarks,bon_edit_code ";
                sql += " ,b.emp_no as bon_emp_code,b.emp_name as bon_emp_name,c.param_name as bon_emp_grade ";
                sql += " from bonusm a";
                sql += " left join empm b on a.bon_emp_id = b.emp_pkid";
                sql += " left join param c on b.emp_grade_id = c.param_pkid ";
                sql += " where  a.bon_pkid ='" + id + "'";
                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Bonusm();
                    mRow.bon_pkid = Dr["bon_pkid"].ToString();
                    mRow.bon_emp_id = Dr["bon_emp_id"].ToString();
                    mRow.bon_emp_code = Dr["bon_emp_code"].ToString();
                    mRow.bon_emp_name = Dr["bon_emp_name"].ToString();
                    mRow.bon_days_worked = Lib.Conv2Integer(Dr["bon_days_worked"].ToString());
                    mRow.bon_gross_wages = Lib.Conv2Decimal(Dr["bon_gross_wages"].ToString());
                    mRow.bon_gross_bonus = Lib.Conv2Decimal(Dr["bon_gross_bonus"].ToString());
                    mRow.bon_puja_deduct = Lib.Conv2Decimal(Dr["bon_puja_deduct"].ToString());
                    mRow.bon_interim_deduct = Lib.Conv2Decimal(Dr["bon_interim_deduct"].ToString());
                    mRow.bon_tax_deduct = Lib.Conv2Decimal(Dr["bon_tax_deduct"].ToString());
                    mRow.bon_other_deduct = Lib.Conv2Decimal(Dr["bon_other_deduct"].ToString());
                    mRow.bon_tot_deduct = Lib.Conv2Decimal(Dr["bon_tot_deduct"].ToString());
                    mRow.bon_net_amount = Lib.Conv2Decimal(Dr["bon_net_amount"].ToString());
                    mRow.bon_actual_paid = Lib.Conv2Decimal(Dr["bon_actual_paid"].ToString());
                    mRow.bon_paid_date = Lib.DatetoString(Dr["bon_paid_date"]);
                    mRow.bon_remarks = Dr["bon_remarks"].ToString();
                    mRow.bon_emp_grade = Dr["bon_emp_grade"].ToString();
                    mRow.bon_edit_code = Dr["bon_edit_code"].ToString();
                    break;
                }
                Dt_Rec.Rows.Clear();
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

        public string AllValid(Bonusm Record)
        {
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {
                sql = " select bon_pkid from bonusm ";
                sql += " where bon_pkid = '" + Record.bon_pkid + "'";
                sql += " and bon_edit_code is null";
                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " Details Closed, Can't Edit ");
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(Bonusm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                if ((ErrorMessage = AllValid(Record)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("bonusm", Record.rec_mode, "bon_pkid", Record.bon_pkid);
                Rec.InsertNumeric("bon_days_worked", Record.bon_days_worked.ToString());
                Rec.InsertNumeric("bon_gross_wages", Record.bon_gross_wages.ToString());
                Rec.InsertNumeric("bon_gross_bonus", Record.bon_gross_bonus.ToString());
                Rec.InsertNumeric("bon_puja_deduct", Record.bon_puja_deduct.ToString());
                Rec.InsertNumeric("bon_interim_deduct", Record.bon_interim_deduct.ToString());
                Rec.InsertNumeric("bon_tax_deduct", Record.bon_tax_deduct.ToString());
                Rec.InsertNumeric("bon_other_deduct", Record.bon_other_deduct.ToString());
                Rec.InsertNumeric("bon_tot_deduct", Record.bon_tot_deduct.ToString());
                Rec.InsertNumeric("bon_net_amount", Record.bon_net_amount.ToString());
                Rec.InsertNumeric("bon_actual_paid", Record.bon_actual_paid.ToString());
                Rec.InsertDate("bon_paid_date", Record.bon_paid_date.ToString());
                Rec.InsertString("bon_remarks", Record.bon_remarks.ToString());
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("bon_emp_id", Record.bon_emp_id.ToString());
                    Rec.InsertNumeric("bon_fin_year", (Lib.Conv2Integer(Record._globalvariables.year_code) - 1).ToString());
                    Rec.InsertString("bon_edit_code", "{S}");
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

        public IDictionary<string, object> Generate(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            List<Bonusm> mList = new List<Bonusm>();
            Bonusm mRow;
            string ErrorMessage = "";
            string SQL1, SQL2 = "";
            string type = SearchData["type"].ToString();
            // string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            string user_code = SearchData["user_code"].ToString();
            string bonempids = SearchData["bonempids"].ToString();
            //int salmonth = Lib.Conv2Integer(SearchData["salmonth"].ToString());
            //int salyear = Lib.Conv2Integer(SearchData["salyear"].ToString());
            int DaysInMonth = 0;
            string Emp_Ids = "";
            DataRow Dr_PS = null;
            bool bTrans = false;
            string pf_excluded_Cols = "";
            decimal pf_excluded_Amt = 0;
            int Fin_Year = 0;
            int PrevFin_Year = 0;
            string Sal_Emp_ID = "";
            string Sal_Emp_Code = "";
            string Bon_Emp_ID = "";
            string Bon_Emp_Code = "";
            try
            {
                //if (salmonth > 0 && salyear > 0)
                //    DaysInMonth = DateTime.DaysInMonth(salyear, salmonth);

                //if ((ErrorMessage = GenerateValid(salyear, salmonth, branch_code, company_code)) != "")
                //{
                //    if (Con_Oracle != null)
                //        Con_Oracle.CloseConnection();
                //    throw new Exception(ErrorMessage);
                //}
                Fin_Year = Lib.Conv2Integer(year_code);
                PrevFin_Year = Fin_Year - 1;

                DataTable Dt_List = new DataTable();

                sql = "select bon_emp_id,emp_no from bonusm a";
                sql += " inner join empm b on a.bon_emp_id = b.emp_pkid";
                sql += " where a.rec_company_code = '" + company_code + "'";
                sql += " and a.rec_branch_code = '" + branch_code + "'";
                sql += " and bon_fin_year =" + PrevFin_Year.ToString();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Bon_Emp_ID = "";
                Bon_Emp_Code = "";
                foreach (DataRow dr in Dt_List.Rows)
                {
                    if (Bon_Emp_ID != "")
                        Bon_Emp_ID += ",";
                    Bon_Emp_ID += dr["bon_emp_id"].ToString();
                    if (Bon_Emp_Code != "")
                        Bon_Emp_Code += ",";
                    Bon_Emp_Code += dr["emp_no"].ToString();
                }
                Bon_Emp_ID = Bon_Emp_ID.Replace(",", "','");
                Bon_Emp_Code = Bon_Emp_Code.Replace(",", "','");

                sql = "select distinct sal_emp_id from salarym ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and sal_fin_year =" + PrevFin_Year.ToString();
                if (Bon_Emp_ID != "")
                    sql += " and sal_emp_id not in ('" + Bon_Emp_ID + "')";
    
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Sal_Emp_ID = "";
                foreach (DataRow dr in Dt_List.Rows)
                {
                    if (Sal_Emp_ID != "")
                        Sal_Emp_ID += ",";
                    Sal_Emp_ID += dr["sal_emp_id"].ToString();
                }

                if (Sal_Emp_ID == "")
                {
                    sql = " select distinct cc_code from costcentert a";
                    sql += " inner join costcenterm b on a.ct_cost_id = b.cc_pkid and a.ct_category ='EMPLOYEE' ";
                    sql += " where a.rec_company_code = '" + company_code + "'";
                    sql += " and a.rec_branch_code = '" + branch_code + "'";
                    sql += " and ct_cost_year = " + PrevFin_Year.ToString();
                    if (Bon_Emp_Code != "")
                        sql += " and cc_code not in ('" + Bon_Emp_Code + "')";
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Sal_Emp_Code = "";
                    foreach (DataRow dr in Dt_List.Rows)
                    {
                        if (Sal_Emp_Code != "")
                            Sal_Emp_Code += ",";
                        Sal_Emp_Code += dr["cc_code"].ToString();
                    }
                    Sal_Emp_Code = Sal_Emp_Code.Replace(",", "','");
                }

                Dt_List = new DataTable();


                if (Sal_Emp_Code != "")//this is for first time creation of bouns otherwise pick from previous year payroll
                {
                    sql = "   select emp_pkid, emp_no,emp_name ,ROUND( ";  //ROUND

                    sql += "   case when mnth_wrkd > 0 and balday_wrkd > 0 then (7000/365)*days_worked  else ";
                    sql += "   (mnth_wrkd*583.3333333) end  ";

                    sql += "   ) as bonus,days_worked, gross_wage from (";
                    sql += "   select b.emp_pkid,  b.Emp_no,b.Emp_name         ";
                    sql += "  ,max(c.DAYS_WORKED) as  DAYS_WORKED ";
                    sql += "  ,max(c.GROSS_WAGE) as GROSS_WAGE ";
                    sql += "  ,max(c.MONTH_WORKED) as mnth_wrkd ";
                    sql += "  ,max(c.BALDAYS_WORKED) as balday_wrkd ";
                    sql += "  from empm b";
                    sql += "  left join sal2018 c on b.emp_no = c.emp_no";
                    sql += " where b.rec_company_code = '" + company_code + "'";
                    sql += " and b.rec_branch_code = '" + branch_code + "'";
                    sql += " and b.emp_no in ('" + Sal_Emp_Code + "')";
                    sql += "  group by b.emp_pkid,b.emp_no,b.emp_name order by b.emp_no";
                    sql += "  ) b ";
                    sql += " order by emp_no";

                }
                else
                {
                    sql = "   select emp_pkid, emp_no,emp_name ,ROUND( ";//CEIL
                    sql += "   (7000/365)*days_worked ";
                    sql += "   ) as bonus,days_worked, gross_wage from (";
                    sql += "   select emp_pkid,  Emp_no,Emp_name         ";
                    sql += "  ,sum(case when to_char(sal_date ,'MMYYYY') = to_char(emp_do_joining ,'MMYYYY') then";
                    sql += "   cast(to_char(last_day(emp_do_joining),'dd') as int)-cast(to_char(emp_do_joining,'dd') as int) + 1 ";
                    sql += "   else ";
                    sql += "   case when to_char(sal_date ,'MMYYYY') = to_char(emp_do_relieve ,'MMYYYY') then";
                    sql += "   cast(to_char(emp_do_relieve,'dd') as int)  ";
                    sql += "   else cast(to_char(last_day(sal_date),'dd') as int) end end ) as DAYS_WORKED  ";
                    sql += "  ,sum(sal_gross_earn) as GROSS_WAGE ";
                    sql += "  from salarym a";
                    sql += "  inner join empm b on a.sal_emp_id = b.emp_pkid  ";
                    sql += "  where a.rec_company_code = '" + company_code + "'";
                    sql += "  and a.rec_branch_code = '" + branch_code + "'";
                    sql += "  and sal_fin_year =" + PrevFin_Year.ToString();
                    sql += "  and emp_pkid in ('" + Sal_Emp_ID + "')";
                    sql += "  group by emp_pkid,emp_no,emp_name order by emp_no";
                    sql += "  ) b ";
                    sql += " order by emp_no";
                }

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                if (type == "LIST")
                {
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mRow = new Bonusm();
                        mRow.bon_emp_id = Dr["emp_pkid"].ToString();
                        mRow.bon_emp_code = Dr["emp_no"].ToString();
                        mRow.bon_emp_name = Dr["emp_name"].ToString();
                        mRow.bon_days_worked = Lib.Conv2Integer(Dr["days_worked"].ToString());
                        mRow.bon_gross_wages = Lib.Conv2Decimal(Dr["gross_wage"].ToString());
                        mRow.bon_gross_bonus = Lib.Conv2Decimal(Dr["bonus"].ToString());

                        mList.Add(mRow);
                    }
                }
                else
                {
                    bonempids = bonempids.Replace(",", "','");
                    foreach (DataRow dr in Dt_List.Select("emp_pkid in('" + bonempids + "')"))
                    {
                        bTrans = false;
                    
                        SQL1 = "  Insert into BONUSM ";
                        SQL1 += "    (BON_PKID";
                        SQL2 = " Values ('" + Guid.NewGuid().ToString().ToUpper() + "'";
                        SQL1 += "    ,BON_EMP_ID";
                        SQL2 += " ,'" + dr["EMP_PKID"].ToString() + "'";
                        SQL1 += "    ,BON_DAYS_WORKED";
                        SQL2 += " ," + Lib.Convert2Decimal(dr["days_worked"].ToString());
                        SQL1 += "    ,BON_FIN_YEAR";
                        SQL2 += " ," + PrevFin_Year.ToString();
                        SQL1 += "    ,BON_GROSS_WAGES";
                        SQL2 += "," + Lib.Convert2Decimal(dr["gross_wage"].ToString());
                        SQL1 += "    ,BON_GROSS_BONUS";
                        SQL2 += "," + Lib.Convert2Decimal(dr["bonus"].ToString());
                        SQL1 += "    ,BON_NET_AMOUNT";
                        SQL2 += "," + Lib.Convert2Decimal(dr["bonus"].ToString());
                        SQL1 += "    ,BON_ACTUAL_PAID";
                        SQL2 += "," + Lib.Convert2Decimal(dr["bonus"].ToString());
                        SQL1 += "    ,BON_EDIT_CODE";
                        SQL2 += " ,'{S}'";
                        SQL1 += "    ,REC_COMPANY_CODE";
                        SQL2 += " ,'" + company_code + "'";
                        SQL1 += "    ,REC_BRANCH_CODE ";
                        SQL2 += " ,'" + branch_code + "'";
                        SQL1 += "    ,REC_CREATED_BY ";
                        SQL2 += " ,'" + user_code + "'";
                        SQL1 += "    ,REC_CREATED_DATE )";
                        SQL2 += ",(SYSDATE))";

                        Con_Oracle.BeginTransaction();
                        sql = SQL1 + SQL2; bTrans = true;
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                    }
                }
                Con_Oracle.CloseConnection();
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

            RetData.Add("list", mList);
            return RetData;
        }

        private void PrintBonusLetter(string Comp_name, string comp_code,string br_code,string yr_code,bool bRelive)
        {
            string sql = "";
            DataTable Dt_daysWrkd = null;
            DataTable Dt_Parent = null;
            Con_Oracle = new DBConnection();
            try
            {
                sql = " select EMP_NO,EMP_NAME,EMP_FATHER_NAME,EMP_DO_BIRTH";
                sql += " ,c.param_name as EMP_DESIGNATION ";
                sql += " ,b.*";
                sql += " from empm a";
                sql += " inner join bonusm b on (a.emp_pkid=b.bon_emp_id and b.bon_fin_year = " + (Lib.Conv2Integer(yr_code) - 1).ToString() + " ) ";
                sql += " left join param c on a.emp_designation_id  = c.param_pkid";
                sql += " where bon_fin_year = " + (Lib.Conv2Integer(yr_code) - 1).ToString();
                sql += "  and b.rec_company_code = '" + comp_code + "'";
                sql += "  and b.rec_branch_code = '" + br_code + "'";
                if (bRelive)
                    sql += " and emp_in_payroll != 'Y' ";
                else
                    sql += " and emp_in_payroll = 'Y' ";
                sql += " order by emp_no";

                Dt_Parent = new DataTable();
                Dt_Parent = Con_Oracle.ExecuteQuery(sql);

                sql = "select sal_pkid from salarym where rownum=1 and  sal_fin_year = " + (Lib.Conv2Integer(yr_code) - 1).ToString();
                sql += "  and rec_company_code = '" + comp_code + "'";
                sql += "  and rec_branch_code = '" + br_code + "'";
                if (Con_Oracle.IsRowExists(sql))
                {

                    sql = " select emp_no, sum(case when to_char(sal_date ,'MMYYYY') = to_char(emp_do_joining ,'MMYYYY') then ";
                    sql += "    case when cast(to_char(last_day(emp_do_joining),'dd') as int)-cast(to_char(emp_do_joining,'dd') as int) + 1 < 27 then ";
                    sql += "    0  else 1 end  ";
                    sql += "    else case when to_char(sal_date ,'MMYYYY') = to_char(emp_do_relieve ,'MMYYYY') then ";
                    sql += "    case when cast(to_char(emp_do_relieve,'dd') as int) < 27 then 0  ";
                    sql += "    else 1 end ";
                    sql += "    else 1 end end ) as MONTH_WORKED   ";
                    sql += "    ,sum(case when to_char(sal_date ,'MMYYYY') = to_char(emp_do_joining ,'MMYYYY') then ";
                    sql += "    case when cast(to_char(last_day(emp_do_joining),'dd') as int)-cast(to_char(emp_do_joining,'dd') as int) + 1 < 27 then ";
                    sql += "    (cast(to_char(last_day(emp_do_joining),'dd') as int)-cast(to_char(emp_do_joining,'dd') as int)) + 1  else 0 end  ";
                    sql += "    else  case when to_char(sal_date ,'MMYYYY') = to_char(emp_do_relieve ,'MMYYYY') then ";
                    sql += "    case when cast(to_char(emp_do_relieve,'dd') as int) < 27 then cast(to_char(emp_do_relieve,'dd') as int)  ";
                    sql += "    else 0 end ";
                    sql += "    else 0 end end ) as DAYS_WORKED   ";
                    sql += "   from salarym a";
                    sql += "   inner join empm e on a.sal_emp_id = e.emp_pkid";
                    sql += "   inner join bonusm b on (a.sal_emp_id=b.bon_emp_id and b.bon_fin_year = " + (Lib.Conv2Integer(yr_code) - 1).ToString() + ") ";
                    sql += "   where  sal_date between '01-APR-" + (Lib.Conv2Integer(yr_code) - 1).ToString() + "' and '31-MAR-" + yr_code + "' ";
                    sql += "   and b.rec_company_code = '" + comp_code + "'";
                    sql += "   and b.rec_branch_code = '" + br_code + "'";
                    sql += "   group by emp_no,emp_name order by emp_no ";
                }
                else
                {
                    sql = "select e.emp_no,c.MONTH_WORKED, c.BALDAYS_WORKED as DAYS_WORKED ";
                    sql += "   from empm e";
                    sql += "   inner join bonusm b on (e.emp_pkid=b.bon_emp_id and b.bon_fin_year = " + (Lib.Conv2Integer(yr_code) - 1).ToString() + ") ";
                    sql += "   left join sal2018 c on e.emp_no = c.emp_no";
                    sql += "  where b.rec_company_code = '" + comp_code + "'";
                    sql += "  and b.rec_branch_code = '" + br_code + "'";
                }

                Dt_daysWrkd = new DataTable();
                Dt_daysWrkd = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                int iRow = 1;

                string fname = "myreport";
                fname = "BONUS-" + br_code + "-" + DateTime.Now.ToString("MMMM").ToUpper();
                if (fname.Length > 30)
                    fname = fname.Substring(0, 30);
                File_Display_Name = Lib.ProperFileName(fname) + ".xls";
                File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                File_Type = "xls";

                file = new ExcelFile();
                file.Worksheets.Add("Report");
                ws = file.Worksheets["Report"];
                ws.PrintOptions.Portrait = false;
                ws.PrintOptions.FitWorksheetWidthToPages = 1;

                ws.Columns[0].Width = 5 * 256;
                ws.Columns[1].Width = 25 * 256;
                ws.Columns[2].Width = 20 * 256;
                ws.Columns[3].Width = 9 * 256;
                ws.Columns[4].Width = 17 * 256;
                ws.Columns[5].Width = 7 * 256;
                ws.Columns[6].Width = 9 * 256;
                ws.Columns[7].Width = 8 * 256;
                ws.Columns[8].Width = 8 * 256;
                ws.Columns[9].Width = 6 * 256;
                ws.Columns[10].Width = 7 * 256;
                ws.Columns[11].Width = 10 * 256;
                ws.Columns[12].Width = 8 * 256;
                ws.Columns[13].Width = 8 * 256;
                ws.Columns[14].Width = 7 * 256;
                ws.Columns[15].Width = 10 * 256;
                ws.Columns[16].Width = 8 * 256;

                for (int c = 0; c < 17; c++)
                    ws.Columns[c].Style.Font.Size = 9 * 20;

                WriteHeading(iRow, 0, "FORM C", Color.Black);
                iRow++;
                WriteHeading(iRow, 0, "[See rule 4(c)]", Color.Black);
                iRow++;
                WriteHeading(iRow, 0, "BONUS PAID TO EMPLOYEES FOR THE ACCOUNTING YEAR ENDING ON THE 31-MAR-" + yr_code, Color.Black);
                iRow++;
                iRow++;
                WriteHeading(iRow, 0, "Name of the establishment : " + Comp_name, Color.Black);
                // WriteHeading(iRow, 2, "CARGOMAR PVT LTD", Color.Black);
                // ws.Cells[iRow, 2].Style.Font.UnderlineStyle = UnderlineStyle.Single;
                iRow++;
                iRow++;
                WriteHeading(iRow, 0, "No. of working days in the year : 300", Color.Black);
                iRow++;
                iRow++;
                iRow++;

                for (int c = 0; c < 17; c++)
                    ws.Cells.GetSubrangeRelative(iRow, c, 1, 1).SetBorders(MultipleBorders.Outside, Color.Gray, LineStyle.Thin);
                WriteHeading(iRow, 0, "SrNo", Color.Black);
                WriteHeading(iRow, 1, "Name of the employee", Color.Black);
                WriteHeading(iRow, 2, "Father’s Name", Color.Black);
                WriteHeading(iRow, 3, "Whether he has completed 15 years of age at the beginning of the accounting year", Color.Black);
                WriteHeading(iRow, 4, "Designation ", Color.Black);
                WriteHeading(iRow, 5, "No. of month / days worked in the year", Color.Black);
                WriteHeading(iRow, 6, "Total Salary or wage in respect of the accounting year ", Color.Black);
                WriteHeading(iRow, 7, "Amount of bonus payable under section 10 or section 11 as the case may be", Color.Black);
                WriteHeading(iRow, 8, "Puja bonus or other customary during the accounting year advance", Color.Black);
                WriteHeading(iRow, 9, "Interim bonus or bonus paid", Color.Black);
                WriteHeading(iRow, 10, "Amount of Income-tax deducted", Color.Black);
                WriteHeading(iRow, 11, "Deduction on account of financial loss, if any caused by misconduct of the employees", Color.Black);
                WriteHeading(iRow, 12, "[Total sum deducted under Columns, 9, 10, 10A and 11]", Color.Black);
                WriteHeading(iRow, 13, "Net amount payable (Column 8 minus Columns 12)", Color.Black);
                WriteHeading(iRow, 14, "Amount actually paid ", Color.Black);
                WriteHeading(iRow, 15, "Date on which paid ", Color.Black);
                WriteHeading(iRow, 16, "Signature / Thumb impression of the employees", Color.Black);
                ws.Rows[iRow].Height = ws.Rows[iRow].Height * 10;
                ws.Rows[iRow].Style.WrapText = true;
                ws.Rows[iRow].Style.VerticalAlignment = VerticalAlignmentStyle.Top;

                iRow++;
                for (int c = 0; c < 17; c++)
                {
                    ws.Cells.GetSubrangeRelative(iRow, c, 1, 1).SetBorders(MultipleBorders.Outside, Color.Gray, LineStyle.Thin);
                    ws.Cells[iRow, c].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                }
                ws.Cells[iRow, 0].Value = 1;
                ws.Cells[iRow, 1].Value = 2;
                ws.Cells[iRow, 2].Value = 3;
                ws.Cells[iRow, 3].Value = 4;
                ws.Cells[iRow, 4].Value = 5;
                ws.Cells[iRow, 5].Value = 6;
                ws.Cells[iRow, 6].Value = 7;
                ws.Cells[iRow, 7].Value = 8;
                ws.Cells[iRow, 8].Value = 9;
                ws.Cells[iRow, 9].Value = 10;
                ws.Cells[iRow, 10].Value = "10A";
                ws.Cells[iRow, 11].Value = 11;
                ws.Cells[iRow, 12].Value = 12;
                ws.Cells[iRow, 13].Value = 13;
                ws.Cells[iRow, 14].Value = 14;
                ws.Cells[iRow, 15].Value = 15;
                ws.Cells[iRow, 16].Value = 16;

                //DateTime dob = Convert.ToDateTime("24/11/1982");
                //DateTime CurrentDate = Convert.ToDateTime("01/04/" + (Lib.Conv2Integer(yr_code) - 1).ToString());

                DateTime dob = new DateTime(1982, 11, 24);
                DateTime CurrentDate = new DateTime(Lib.Conv2Integer(yr_code) - 1, 04, 01);

                TimeSpan ts = CurrentDate - dob;
                int Age = ts.Days / 365;
                DataRow[] DrDays = null;
                foreach (DataRow row in Dt_Parent.Rows)
                {
                    DrDays = Dt_daysWrkd.Select("emp_no = '" + row["EMP_NO"].ToString().Trim() + "'");
                    iRow++;
                    ws.Cells[iRow, 0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    ws.Cells[iRow, 1].Style.Font.Size = 8 * 20;
                    ws.Cells[iRow, 2].Style.Font.Size = 8 * 20;
                    ws.Cells[iRow, 4].Style.Font.Size = 8 * 20;
                    for (int c = 0; c < 17; c++)
                        ws.Cells.GetSubrangeRelative(iRow, c, 1, 1).SetBorders(MultipleBorders.Outside, Color.Gray, LineStyle.Thin);

                    ws.Cells[iRow, 0].Value = (iRow - 11);
                    ws.Cells[iRow, 1].Value = row["EMP_NAME"].ToString().Trim();
                    ws.Cells[iRow, 2].Value = row["EMP_FATHER_NAME"].ToString().Trim();
                    if (row["EMP_DO_BIRTH"].Equals(DBNull.Value))
                        dob = DateTime.Now;
                    else
                        dob = (DateTime)row["EMP_DO_BIRTH"];
                    ts = CurrentDate - dob;
                    Age = ts.Days / 365;
                    if (Age > 15)
                        ws.Cells[iRow, 3].Value = "Y";
                    else
                        ws.Cells[iRow, 3].Value = "N";

                    ws.Cells[iRow, 4].Value = row["EMP_DESIGNATION"];

                    //int days = Common.Convert2Integer(row["BON_DAYS_WORKED"].ToString());
                    //double dMon = days / 30;//4368499
                    //double dBaldays = days % 30;//4368499

                    //double months = Math.Floor(dMon);
                    //double balDays = Math.Floor(dBaldays);

                    int months = 0;
                    int balDays = 0;
                    if (DrDays != null && DrDays.Length>0)
                    {
                        months = Lib.Conv2Integer(DrDays[0]["MONTH_WORKED"].ToString());
                        balDays = Lib.Conv2Integer(DrDays[0]["DAYS_WORKED"].ToString());
                    }
                    if (balDays > 0)
                        ws.Cells[iRow, 5].Value = months.ToString() + " / " + balDays.ToString() + "days";
                    else
                        ws.Cells[iRow, 5].Value = months;

                    ws.Cells[iRow, 6].Value = row["BON_GROSS_WAGES"];
                    ws.Cells[iRow, 7].Value = row["BON_GROSS_BONUS"];
                    if (Lib.Convert2Decimal(row["BON_PUJA_DEDUCT"].ToString()) != 0)
                        ws.Cells[iRow, 8].Value = row["BON_PUJA_DEDUCT"];
                    if (Lib.Convert2Decimal(row["BON_INTERIM_DEDUCT"].ToString()) != 0)
                        ws.Cells[iRow, 9].Value = row["BON_INTERIM_DEDUCT"];
                    if (Lib.Convert2Decimal(row["BON_TAX_DEDUCT"].ToString()) != 0)
                        ws.Cells[iRow, 10].Value = row["BON_TAX_DEDUCT"];
                    if (Lib.Convert2Decimal(row["BON_OTHER_DEDUCT"].ToString()) != 0)
                        ws.Cells[iRow, 11].Value = row["BON_OTHER_DEDUCT"];
                    if (Lib.Convert2Decimal(row["BON_TOT_DEDUCT"].ToString()) != 0)
                        ws.Cells[iRow, 12].Value = row["BON_TOT_DEDUCT"];
                    if (Lib.Convert2Decimal(row["BON_NET_AMOUNT"].ToString()) != 0)
                        ws.Cells[iRow, 13].Value = row["BON_NET_AMOUNT"];
                    ws.Cells[iRow, 14].Value = row["BON_ACTUAL_PAID"];
                    ws.Cells[iRow, 15].Value = Lib.DatetoStringDisplayformat(row["BON_PAID_DATE"]).Trim();
                }

                file.SaveXls(File_Name);
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw ex;
            }
        }

        private void WriteHeading(int cRow, int cCol, object cData, Color cFontColor)
        {
            ws.Cells[cRow, cCol].Value = cData;
            ws.Cells[cRow, cCol].Style.Font.Weight = ExcelFont.BoldWeight;
            ws.Cells[cRow, cCol].Style.Font.Color = cFontColor;
        }
    }
}
