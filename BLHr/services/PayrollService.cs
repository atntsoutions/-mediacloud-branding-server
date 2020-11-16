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
    public class PayrollService : BL_Base
    {
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Salarym> mList = new List<Salarym>();
            Salarym mRow;
            SalDet dRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string empstatus = SearchData["empstatus"].ToString();
            int salmonth = 0;
            if (SearchData.ContainsKey("salmonth"))
                salmonth = Lib.Conv2Integer(SearchData["salmonth"].ToString());
            int salyear = Lib.Conv2Integer(SearchData["salyear"].ToString());
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;
            int saljvno = 0;
            DataTable Dt_Temp = null;
            try
            {

                if (salmonth > 0)
                {
                    sql = "select salh_pkid,salh_jvno from salaryh a ";
                    sql += " where a.rec_company_code = '" + company_code + "'";
                    sql += " and a.rec_branch_code = '" + branch_code + "'";
                    sql += " and a.salh_month = " + salmonth.ToString();
                    sql += " and a.salh_year = " + salyear.ToString();
                    Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        saljvno = Lib.Conv2Integer(Dt_Temp.Rows[0]["salh_jvno"].ToString());
                    }
                }

                dRow = getListColumns();

                sWhere = " where a.rec_company_code = '" + company_code + "'";
                sWhere += " and a.rec_branch_code = '" + branch_code + "'";
                if (salmonth > 0)
                    sWhere += " and a.sal_month = " + salmonth.ToString();
                sWhere += " and a.sal_year = " + salyear.ToString();
                if (empstatus != "BOTH")
                    sWhere += " and a.rec_category = '" + empstatus.ToString() + "'";

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
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total FROM salarym a ";
                    sql += " left join empm b on a.sal_emp_id = b.emp_pkid ";
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
                sql += " select sal_pkid,sal_month,sal_date,sal_emp_id,sal_gross_earn,sal_gross_deduct,sal_net,sal_lop_amt ";
                sql += " ,a01,a02,a03,a04,a05";
                sql += " ,a06,a07,a08,a09,a10";
                sql += " ,a11,a12,a13,a14,a15";
                sql += " ,a16,a17,a18,a19,a20";
                sql += " ,a21,a22,a23,a24,a25";
                sql += " ,d01,d02,d03,d04,d05";
                sql += " ,d06,d07,d08,d09,d10";
                sql += " ,d11,d12,d13,d14,d15";
                sql += " ,d16,d17,d18,d19,d20";
                sql += " ,d21,d22,d23,d24,d25";
                sql += " ,emp_no,emp_name,emp_do_joining, c.param_name as emp_grade,a.rec_printed,sal_mail_sent,sal_emp_branch_group as emp_branch_group ";
                sql += " ,row_number() over (order by emp_no) rn ";
                sql += " from salarym a ";
                sql += " inner join empm b on a.sal_emp_id = b.emp_pkid ";
                sql += " left join param c on b.emp_grade_id = c.param_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by emp_no";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Salarym();
                    mRow.sal_pkid = Dr["sal_pkid"].ToString();
                    mRow.sal_emp_id = Dr["sal_emp_id"].ToString();
                    mRow.sal_emp_code = Dr["emp_no"].ToString();
                    mRow.sal_emp_name = Dr["emp_name"].ToString();
                    mRow.sal_emp_grade = Dr["emp_grade"].ToString();
                    mRow.sal_month = Lib.Conv2Integer(Dr["sal_month"].ToString());
                    mRow.sal_date = Lib.DatetoStringDisplayformat(Dr["sal_date"]);
                    mRow.sal_emp_do_joining = Lib.DatetoStringDisplayformat(Dr["emp_do_joining"]);
                    mRow.a01 = Lib.Conv2Decimal(Dr["a01"].ToString());
                    mRow.a02 = Lib.Conv2Decimal(Dr["a02"].ToString());
                    mRow.a03 = Lib.Conv2Decimal(Dr["a03"].ToString());
                    mRow.a04 = Lib.Conv2Decimal(Dr["a04"].ToString());
                    mRow.a05 = Lib.Conv2Decimal(Dr["a05"].ToString());
                    mRow.a06 = Lib.Conv2Decimal(Dr["a06"].ToString());
                    mRow.a07 = Lib.Conv2Decimal(Dr["a07"].ToString());
                    mRow.a08 = Lib.Conv2Decimal(Dr["a08"].ToString());
                    mRow.a09 = Lib.Conv2Decimal(Dr["a09"].ToString());
                    mRow.a10 = Lib.Conv2Decimal(Dr["a10"].ToString());
                    mRow.a11 = Lib.Conv2Decimal(Dr["a11"].ToString());
                    mRow.a12 = Lib.Conv2Decimal(Dr["a12"].ToString());
                    mRow.a13 = Lib.Conv2Decimal(Dr["a13"].ToString());
                    mRow.a14 = Lib.Conv2Decimal(Dr["a14"].ToString());
                    mRow.a15 = Lib.Conv2Decimal(Dr["a15"].ToString());
                    mRow.a16 = Lib.Conv2Decimal(Dr["a16"].ToString());
                    mRow.a17 = Lib.Conv2Decimal(Dr["a17"].ToString());
                    mRow.a18 = Lib.Conv2Decimal(Dr["a18"].ToString());
                    mRow.a19 = Lib.Conv2Decimal(Dr["a19"].ToString());
                    mRow.a20 = Lib.Conv2Decimal(Dr["a20"].ToString());
                    mRow.a21 = Lib.Conv2Decimal(Dr["a21"].ToString());
                    mRow.a22 = Lib.Conv2Decimal(Dr["a22"].ToString());
                    mRow.a23 = Lib.Conv2Decimal(Dr["a23"].ToString());
                    mRow.a24 = Lib.Conv2Decimal(Dr["a24"].ToString());
                    mRow.a25 = Lib.Conv2Decimal(Dr["a25"].ToString());
                    mRow.d01 = Lib.Conv2Decimal(Dr["d01"].ToString());
                    mRow.d02 = Lib.Conv2Decimal(Dr["d02"].ToString());
                    mRow.d03 = Lib.Conv2Decimal(Dr["d03"].ToString());
                    mRow.d04 = Lib.Conv2Decimal(Dr["d04"].ToString());
                    mRow.d05 = Lib.Conv2Decimal(Dr["d05"].ToString());
                    mRow.d06 = Lib.Conv2Decimal(Dr["d06"].ToString());
                    mRow.d07 = Lib.Conv2Decimal(Dr["d07"].ToString());
                    mRow.d08 = Lib.Conv2Decimal(Dr["d08"].ToString());
                    mRow.d09 = Lib.Conv2Decimal(Dr["d09"].ToString());
                    mRow.d10 = Lib.Conv2Decimal(Dr["d10"].ToString());
                    mRow.d11 = Lib.Conv2Decimal(Dr["d11"].ToString());
                    mRow.d12 = Lib.Conv2Decimal(Dr["d12"].ToString());
                    mRow.d13 = Lib.Conv2Decimal(Dr["d13"].ToString());
                    mRow.d14 = Lib.Conv2Decimal(Dr["d14"].ToString());
                    mRow.d15 = Lib.Conv2Decimal(Dr["d15"].ToString());
                    mRow.d16 = Lib.Conv2Decimal(Dr["d16"].ToString());
                    mRow.d17 = Lib.Conv2Decimal(Dr["d17"].ToString());
                    mRow.d18 = Lib.Conv2Decimal(Dr["d18"].ToString());
                    mRow.d19 = Lib.Conv2Decimal(Dr["d19"].ToString());
                    mRow.d20 = Lib.Conv2Decimal(Dr["d20"].ToString());
                    mRow.d21 = Lib.Conv2Decimal(Dr["d21"].ToString());
                    mRow.d22 = Lib.Conv2Decimal(Dr["d22"].ToString());
                    mRow.d23 = Lib.Conv2Decimal(Dr["d23"].ToString());
                    mRow.d24 = Lib.Conv2Decimal(Dr["d24"].ToString());
                    mRow.d25 = Lib.Conv2Decimal(Dr["d25"].ToString());
                    mRow.sal_gross_earn = Lib.Conv2Decimal(Dr["sal_gross_earn"].ToString());
                    mRow.sal_gross_deduct = Lib.Conv2Decimal(Dr["sal_gross_deduct"].ToString());
                    mRow.sal_net = Lib.Conv2Decimal(Dr["sal_net"].ToString());
                    mRow.sal_lop_amt = Lib.Conv2Decimal(Dr["sal_lop_amt"].ToString());
                    mRow.rec_printed = Dr["rec_printed"].ToString() == "Y" ? true : false;
                    mRow.sal_mail_sent = Dr["sal_mail_sent"].ToString() == "Y" ? true : false;
                    mRow.sal_emp_branch_group = Lib.Conv2Integer(Dr["emp_branch_group"].ToString());
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
            RetData.Add("record", dRow);
            RetData.Add("saljvno", saljvno);
            
            return RetData;
        }

        private SalDet getListColumns()
        {
            SalDet drow = new SalDet();
            drow.a01_desc = "";
            drow.a01_visible = false;
            drow.a02_desc = "";
            drow.a02_visible = false;
            drow.a03_desc = "";
            drow.a03_visible = false;
            drow.a04_desc = "";
            drow.a04_visible = false;
            drow.a05_desc = "";
            drow.a05_visible = false;
            drow.a06_desc = "";
            drow.a06_visible = false;
            drow.a07_desc = "";
            drow.a07_visible = false;
            drow.a08_desc = "";
            drow.a08_visible = false;
            drow.a09_desc = "";
            drow.a09_visible = false;
            drow.a10_desc = "";
            drow.a10_visible = false;
            drow.a11_desc = "";
            drow.a11_visible = false;
            drow.a12_desc = "";
            drow.a12_visible = false;
            drow.a13_desc = "";
            drow.a13_visible = false;
            drow.a14_desc = "";
            drow.a14_visible = false;
            drow.a15_desc = "";
            drow.a15_visible = false;
            drow.a16_desc = "";
            drow.a16_visible = false;
            drow.a17_desc = "";
            drow.a17_visible = false;
            drow.a18_desc = "";
            drow.a18_visible = false;
            drow.a19_desc = "";
            drow.a19_visible = false;
            drow.a20_desc = "";
            drow.a20_visible = false;
            drow.a21_desc = "";
            drow.a21_visible = false;
            drow.a22_desc = "";
            drow.a22_visible = false;
            drow.a23_desc = "";
            drow.a23_visible = false;
            drow.a24_desc = "";
            drow.a24_visible = false;
            drow.a25_desc = "";
            drow.a25_visible = false;
            drow.d01_desc = "";
            drow.d01_visible = false;
            drow.d02_desc = "";
            drow.d02_visible = false;
            drow.d03_desc = "";
            drow.d03_visible = false;
            drow.d04_desc = "";
            drow.d04_visible = false;
            drow.d05_desc = "";
            drow.d05_visible = false;
            drow.d06_desc = "";
            drow.d06_visible = false;
            drow.d07_desc = "";
            drow.d07_visible = false;
            drow.d08_desc = "";
            drow.d08_visible = false;
            drow.d09_desc = "";
            drow.d09_visible = false;
            drow.d10_desc = "";
            drow.d10_visible = false;
            drow.d11_desc = "";
            drow.d11_visible = false;
            drow.d12_desc = "";
            drow.d12_visible = false;
            drow.d13_desc = "";
            drow.d13_visible = false;
            drow.d14_desc = "";
            drow.d14_visible = false;
            drow.d15_desc = "";
            drow.d15_visible = false;
            drow.d16_desc = "";
            drow.d16_visible = false;
            drow.d17_desc = "";
            drow.d17_visible = false;
            drow.d18_desc = "";
            drow.d18_visible = false;
            drow.d19_desc = "";
            drow.d19_visible = false;
            drow.d20_desc = "";
            drow.d20_visible = false;
            drow.d21_desc = "";
            drow.d21_visible = false;
            drow.d22_desc = "";
            drow.d22_visible = false;
            drow.d23_desc = "";
            drow.d23_visible = false;
            drow.d24_desc = "";
            drow.d24_visible = false;
            drow.d25_desc = "";
            drow.d25_visible = false;

            DataTable Dt_Head = new DataTable();
            sql = "select sal_code,sal_desc from salaryheadm where sal_head is not null order by sal_code";
            Dt_Head = Con_Oracle.ExecuteQuery(sql);

            foreach (DataRow dr in Dt_Head.Rows)
            {
                switch (dr["SAL_CODE"].ToString().Trim())
                {
                    case "A01":
                        drow.a01_desc = dr["SAL_DESC"].ToString(); drow.a01_visible = true;
                        break;
                    case "A02":
                        drow.a02_desc = dr["SAL_DESC"].ToString(); drow.a02_visible = true;
                        break;
                    case "A03":
                        drow.a03_desc = dr["SAL_DESC"].ToString(); drow.a03_visible = true;
                        break;
                    case "A04":
                        drow.a04_desc = dr["SAL_DESC"].ToString(); drow.a04_visible = true;
                        break;
                    case "A05":
                        drow.a05_desc = dr["SAL_DESC"].ToString(); drow.a05_visible = true;
                        break;
                    case "A06":
                        drow.a06_desc = dr["SAL_DESC"].ToString(); drow.a06_visible = true;
                        break;
                    case "A07":
                        drow.a07_desc = dr["SAL_DESC"].ToString(); drow.a07_visible = true;
                        break;
                    case "A08":
                        drow.a08_desc = dr["SAL_DESC"].ToString(); drow.a08_visible = true;
                        break;
                    case "A09":
                        drow.a09_desc = dr["SAL_DESC"].ToString(); drow.a09_visible = true;
                        break;
                    case "A10":
                        drow.a10_desc = dr["SAL_DESC"].ToString(); drow.a10_visible = true;
                        break;
                    case "A11":
                        drow.a11_desc = dr["SAL_DESC"].ToString(); drow.a11_visible = true;
                        break;
                    case "A12":
                        drow.a12_desc = dr["SAL_DESC"].ToString(); drow.a12_visible = true;
                        break;
                    case "A13":
                        drow.a13_desc = dr["SAL_DESC"].ToString(); drow.a13_visible = true;
                        break;
                    case "A14":
                        drow.a14_desc = dr["SAL_DESC"].ToString(); drow.a14_visible = true;
                        break;
                    case "A15":
                        drow.a15_desc = dr["SAL_DESC"].ToString(); drow.a15_visible = true;
                        break;
                    case "A16":
                        drow.a16_desc = dr["SAL_DESC"].ToString(); drow.a16_visible = true;
                        break;
                    case "A17":
                        drow.a17_desc = dr["SAL_DESC"].ToString(); drow.a17_visible = true;
                        break;
                    case "A18":
                        drow.a18_desc = dr["SAL_DESC"].ToString(); drow.a18_visible = true;
                        break;
                    case "A19":
                        drow.a19_desc = dr["SAL_DESC"].ToString(); drow.a19_visible = true;
                        break;
                    case "A20":
                        drow.a20_desc = dr["SAL_DESC"].ToString(); drow.a20_visible = true;
                        break;
                    case "A21":
                        drow.a21_desc = dr["SAL_DESC"].ToString(); drow.a21_visible = true;
                        break;
                    case "A22":
                        drow.a22_desc = dr["SAL_DESC"].ToString(); drow.a22_visible = true;
                        break;
                    case "A23":
                        drow.a23_desc = dr["SAL_DESC"].ToString(); drow.a23_visible = true;
                        break;
                    case "A24":
                        drow.a24_desc = dr["SAL_DESC"].ToString(); drow.a24_visible = true;
                        break;

                    case "D01":
                        drow.d01_desc = dr["SAL_DESC"].ToString(); drow.d01_visible = true;
                        break;
                    case "D02":
                        drow.d02_desc = dr["SAL_DESC"].ToString(); drow.d02_visible = true;
                        break;
                    case "D03":
                        drow.d03_desc = dr["SAL_DESC"].ToString(); drow.d03_visible = true;
                        break;
                    case "D04":
                        drow.d04_desc = dr["SAL_DESC"].ToString(); drow.d04_visible = true;
                        break;
                    case "D05":
                        drow.d05_desc = dr["SAL_DESC"].ToString(); drow.d05_visible = true;
                        break;
                    case "D06":
                        drow.d06_desc = dr["SAL_DESC"].ToString(); drow.d06_visible = true;
                        break;
                    case "D07":
                        drow.d07_desc = dr["SAL_DESC"].ToString(); drow.d07_visible = true;
                        break;
                    case "D08":
                        drow.d08_desc = dr["SAL_DESC"].ToString(); drow.d08_visible = true;
                        break;
                    case "D09":
                        drow.d09_desc = dr["SAL_DESC"].ToString(); drow.d09_visible = true;
                        break;
                    case "D10":
                        drow.d10_desc = dr["SAL_DESC"].ToString(); drow.d10_visible = true;
                        break;
                    case "D11":
                        drow.d11_desc = dr["SAL_DESC"].ToString(); drow.d11_visible = true;
                        break;
                    case "D12":
                        drow.d12_desc = dr["SAL_DESC"].ToString(); drow.d12_visible = true;
                        break;
                    case "D13":
                        drow.d13_desc = dr["SAL_DESC"].ToString(); drow.d13_visible = true;
                        break;
                    case "D14":
                        drow.d14_desc = dr["SAL_DESC"].ToString(); drow.d14_visible = true;
                        break;
                    case "D15":
                        drow.d15_desc = dr["SAL_DESC"].ToString(); drow.d15_visible = true;
                        break;
                    case "D16":
                        drow.d16_desc = dr["SAL_DESC"].ToString(); drow.d16_visible = true;
                        break;
                    case "D17":
                        drow.d17_desc = dr["SAL_DESC"].ToString(); drow.d17_visible = true;
                        break;
                    case "D18":
                        drow.d18_desc = dr["SAL_DESC"].ToString(); drow.d18_visible = true;
                        break;
                    case "D19":
                        drow.d19_desc = dr["SAL_DESC"].ToString(); drow.d19_visible = true;
                        break;
                    case "D20":
                        drow.d20_desc = dr["SAL_DESC"].ToString(); drow.d20_visible = true;
                        break;
                    case "D21":
                        drow.d21_desc = dr["SAL_DESC"].ToString(); drow.d21_visible = true;
                        break;
                    case "D22":
                        drow.d22_desc = dr["SAL_DESC"].ToString(); drow.d22_visible = true;
                        break;
                    case "D23":
                        drow.d23_desc = dr["SAL_DESC"].ToString(); drow.d23_visible = true;
                        break;
                    case "D24":
                        drow.d24_desc = dr["SAL_DESC"].ToString(); drow.d24_visible = true;
                        break;
                }
            }
            Dt_Head.Rows.Clear();
            return drow;
        }
        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Salarym mRow = new Salarym();
            string id = SearchData["pkid"].ToString();
            string smode = "ADD";
            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select sal_pkid,sal_emp_id,b.emp_no as sal_emp_code,b.emp_name as sal_emp_name";
                sql += " ,sal_date,sal_month,sal_year,sal_fin_year,sal_days_worked";
                sql += " ,a01,a02,a03,a04,a05";
                sql += " ,a06,a07,a08,a09,a10";
                sql += " ,a11,a12,a13,a14,a15";
                sql += " ,a16,a17,a18,a19,a20";
                sql += " ,a21,a22,a23,a24,a25";
                sql += " ,d01,d02,d03,d04,d05";
                sql += " ,d06,d07,d08,d09,d10";
                sql += " ,d11,d12,d13,d14,d15";
                sql += " ,d16,d17,d18,d19,d20";
                sql += " ,d21,d22,d23,d24,d25";
                sql += " ,sal_lop_amt,sal_gross_earn,sal_gross_deduct,sal_net ";
                sql += " ,sal_basic_rt,sal_da_rt,sal_pf_mon_year,sal_pf_limit";
                sql += " ,sal_pf_cel_limit,sal_pf_cel_limit_amt,sal_pf_bal,sal_pf_wage_bal ";
                sql += " ,sal_pf_base ,sal_pf_emplr,sal_pf_emplr_share,sal_pf_emplr_pension";
                sql += " ,sal_pf_emplr_pension_per ,sal_pf_eps_amt,sal_admin_per,sal_admin_amt ";
                sql += " ,sal_admin_based_on ,sal_edli_per,sal_edli_amt,sal_edli_based_on ";
                sql += " ,nvl(sal_is_esi,'N') as sal_is_esi,sal_esi_base,sal_esi_emplr_per,sal_esi_limit,sal_esi_gov_share ";
                sql += " ,sal_pay_date ,sal_work_days,sal_mail_sent,c.param_name as emp_grade,a.rec_company_code,a.rec_branch_code,a.rec_category,sal_edit_code ";
                sql += " from salarym a  ";
                sql += " inner join empm b on a.sal_emp_id = b.emp_pkid ";
                sql += " left join param c on b.emp_grade_id = c.param_pkid ";
                sql += " where a.sal_pkid ='" + id + "'";


                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    smode = "EDIT";
                    mRow.sal_pkid = Dr["sal_pkid"].ToString();
                    mRow.sal_emp_id = Dr["sal_emp_id"].ToString();
                    mRow.sal_emp_code = Dr["sal_emp_code"].ToString();
                    mRow.sal_emp_name = Dr["sal_emp_name"].ToString();
                    mRow.sal_emp_grade = Dr["emp_grade"].ToString();
                    mRow.sal_date = Lib.DatetoString(Dr["sal_date"]);
                    mRow.sal_month = Lib.Conv2Decimal(Dr["sal_month"].ToString());
                    mRow.sal_year = Lib.Conv2Decimal(Dr["sal_year"].ToString());
                    mRow.sal_fin_year = Lib.Conv2Decimal(Dr["sal_fin_year"].ToString());
                    mRow.sal_days_worked = Lib.Conv2Decimal(Dr["sal_days_worked"].ToString());
                    mRow.a01 = Lib.Conv2Decimal(Dr["a01"].ToString());
                    mRow.a02 = Lib.Conv2Decimal(Dr["a02"].ToString());
                    mRow.a03 = Lib.Conv2Decimal(Dr["a03"].ToString());
                    mRow.a04 = Lib.Conv2Decimal(Dr["a04"].ToString());
                    mRow.a05 = Lib.Conv2Decimal(Dr["a05"].ToString());
                    mRow.a06 = Lib.Conv2Decimal(Dr["a06"].ToString());
                    mRow.a07 = Lib.Conv2Decimal(Dr["a07"].ToString());
                    mRow.a08 = Lib.Conv2Decimal(Dr["a08"].ToString());
                    mRow.a09 = Lib.Conv2Decimal(Dr["a09"].ToString());
                    mRow.a10 = Lib.Conv2Decimal(Dr["a10"].ToString());
                    mRow.a11 = Lib.Conv2Decimal(Dr["a11"].ToString());
                    mRow.a12 = Lib.Conv2Decimal(Dr["a12"].ToString());
                    mRow.a13 = Lib.Conv2Decimal(Dr["a13"].ToString());
                    mRow.a14 = Lib.Conv2Decimal(Dr["a14"].ToString());
                    mRow.a15 = Lib.Conv2Decimal(Dr["a15"].ToString());
                    mRow.a16 = Lib.Conv2Decimal(Dr["a16"].ToString());
                    mRow.a17 = Lib.Conv2Decimal(Dr["a17"].ToString());
                    mRow.a18 = Lib.Conv2Decimal(Dr["a18"].ToString());
                    mRow.a19 = Lib.Conv2Decimal(Dr["a19"].ToString());
                    mRow.a20 = Lib.Conv2Decimal(Dr["a20"].ToString());
                    mRow.a21 = Lib.Conv2Decimal(Dr["a21"].ToString());
                    mRow.a22 = Lib.Conv2Decimal(Dr["a22"].ToString());
                    mRow.a23 = Lib.Conv2Decimal(Dr["a23"].ToString());
                    mRow.a24 = Lib.Conv2Decimal(Dr["a24"].ToString());
                    mRow.a25 = Lib.Conv2Decimal(Dr["a25"].ToString());
                    mRow.d01 = Lib.Conv2Decimal(Dr["d01"].ToString());
                    mRow.d02 = Lib.Conv2Decimal(Dr["d02"].ToString());
                    mRow.d03 = Lib.Conv2Decimal(Dr["d03"].ToString());
                    mRow.d04 = Lib.Conv2Decimal(Dr["d04"].ToString());
                    mRow.d05 = Lib.Conv2Decimal(Dr["d05"].ToString());
                    mRow.d06 = Lib.Conv2Decimal(Dr["d06"].ToString());
                    mRow.d07 = Lib.Conv2Decimal(Dr["d07"].ToString());
                    mRow.d08 = Lib.Conv2Decimal(Dr["d08"].ToString());
                    mRow.d09 = Lib.Conv2Decimal(Dr["d09"].ToString());
                    mRow.d10 = Lib.Conv2Decimal(Dr["d10"].ToString());
                    mRow.d11 = Lib.Conv2Decimal(Dr["d11"].ToString());
                    mRow.d12 = Lib.Conv2Decimal(Dr["d12"].ToString());
                    mRow.d13 = Lib.Conv2Decimal(Dr["d13"].ToString());
                    mRow.d14 = Lib.Conv2Decimal(Dr["d14"].ToString());
                    mRow.d15 = Lib.Conv2Decimal(Dr["d15"].ToString());
                    mRow.d16 = Lib.Conv2Decimal(Dr["d16"].ToString());
                    mRow.d17 = Lib.Conv2Decimal(Dr["d17"].ToString());
                    mRow.d18 = Lib.Conv2Decimal(Dr["d18"].ToString());
                    mRow.d19 = Lib.Conv2Decimal(Dr["d19"].ToString());
                    mRow.d20 = Lib.Conv2Decimal(Dr["d20"].ToString());
                    mRow.d21 = Lib.Conv2Decimal(Dr["d21"].ToString());
                    mRow.d22 = Lib.Conv2Decimal(Dr["d22"].ToString());
                    mRow.d23 = Lib.Conv2Decimal(Dr["d23"].ToString());
                    mRow.d24 = Lib.Conv2Decimal(Dr["d24"].ToString());
                    mRow.d25 = Lib.Conv2Decimal(Dr["d25"].ToString());
                    mRow.sal_lop_amt = Lib.Conv2Decimal(Dr["sal_lop_amt"].ToString());
                    mRow.sal_gross_earn = Lib.Conv2Decimal(Dr["sal_gross_earn"].ToString());
                    mRow.sal_gross_deduct = Lib.Conv2Decimal(Dr["sal_gross_deduct"].ToString());
                    mRow.sal_net = Lib.Conv2Decimal(Dr["sal_net"].ToString());
                    mRow.sal_basic_rt = Lib.Conv2Decimal(Dr["sal_basic_rt"].ToString());
                    mRow.sal_da_rt = Lib.Conv2Decimal(Dr["sal_da_rt"].ToString());
                    mRow.sal_pf_mon_year = Dr["sal_pf_mon_year"].ToString();
                    mRow.sal_pf_limit = Lib.Conv2Decimal(Dr["sal_pf_limit"].ToString());
                    mRow.sal_pf_cel_limit = Lib.Conv2Decimal(Dr["sal_pf_cel_limit"].ToString());
                    mRow.sal_pf_cel_limit_amt = Lib.Conv2Decimal(Dr["sal_pf_cel_limit_amt"].ToString());
                    mRow.sal_pf_bal = Lib.Conv2Decimal(Dr["sal_pf_bal"].ToString());
                    mRow.sal_pf_wage_bal = Lib.Conv2Decimal(Dr["sal_pf_wage_bal"].ToString());
                    mRow.sal_pf_base = Lib.Conv2Decimal(Dr["sal_pf_base"].ToString());
                    mRow.sal_pf_emplr = Lib.Conv2Decimal(Dr["sal_pf_emplr"].ToString());
                    mRow.sal_pf_emplr_share = Lib.Conv2Decimal(Dr["sal_pf_emplr_share"].ToString());
                    mRow.sal_pf_emplr_pension = Lib.Conv2Decimal(Dr["sal_pf_emplr_pension"].ToString());
                    mRow.sal_pf_emplr_pension_per = Lib.Conv2Decimal(Dr["sal_pf_emplr_pension_per"].ToString());
                    mRow.sal_pf_eps_amt = Lib.Conv2Decimal(Dr["sal_pf_eps_amt"].ToString());
                    mRow.sal_admin_per = Lib.Conv2Decimal(Dr["sal_admin_per"].ToString());
                    mRow.sal_admin_amt = Lib.Conv2Decimal(Dr["sal_admin_amt"].ToString());
                    mRow.sal_admin_based_on = Dr["sal_admin_based_on"].ToString();
                    mRow.sal_edli_per = Lib.Conv2Decimal(Dr["sal_edli_per"].ToString());
                    mRow.sal_edli_amt = Lib.Conv2Decimal(Dr["sal_edli_amt"].ToString());
                    mRow.sal_edli_based_on = Dr["sal_edli_based_on"].ToString();
                    mRow.sal_is_esi = Dr["sal_is_esi"].ToString() == "Y" ? true : false;
                    mRow.sal_esi_base = Lib.Conv2Decimal(Dr["sal_esi_base"].ToString());
                    mRow.sal_esi_emplr_per = Lib.Conv2Decimal(Dr["sal_esi_emplr_per"].ToString());
                    mRow.sal_esi_limit = Lib.Conv2Decimal(Dr["sal_esi_limit"].ToString());
                    mRow.sal_esi_gov_share = Lib.Conv2Decimal(Dr["sal_esi_gov_share"].ToString());
                    mRow.sal_pay_date = Lib.DatetoString(Dr["sal_pay_date"]);
                    mRow.sal_work_days = Lib.Conv2Decimal(Dr["sal_work_days"].ToString());
                    mRow.sal_mail_sent = Dr["sal_mail_sent"].ToString() == "Y" ? true : false;
                    mRow.sal_selected = false;
                    mRow.rec_company_code = Dr["rec_company_code"].ToString();
                    mRow.rec_branch_code = Dr["rec_branch_code"].ToString();
                    mRow.rec_category = Dr["rec_category"].ToString();
                    mRow.sal_edit_code = Dr["sal_edit_code"].ToString();
                    
                    break;
                }

                if (smode == "EDIT")
                {
                    sql = "select lev_pl, lev_cl, lev_sl, lev_others, lev_lp from leavem ";
                    sql += " where rec_company_code = '" + mRow.rec_company_code + "'";
                    sql += " and rec_branch_code = '" + mRow.rec_branch_code + "'";
                    sql += " and lev_emp_id = '" + mRow.sal_emp_id + "'";
                    sql += " and lev_year = " + mRow.sal_year;
                    sql += " and lev_month = " + mRow.sal_month;

                    Con_Oracle = new DBConnection();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    if (Dt_Rec.Rows.Count > 0)
                    {
                        mRow.sal_pl = Lib.Conv2Decimal(Dt_Rec.Rows[0]["lev_pl"].ToString());
                        mRow.sal_sl = Lib.Conv2Decimal(Dt_Rec.Rows[0]["lev_sl"].ToString());
                        mRow.sal_cl = Lib.Conv2Decimal(Dt_Rec.Rows[0]["lev_cl"].ToString());
                        mRow.sal_ot = Lib.Conv2Decimal(Dt_Rec.Rows[0]["lev_others"].ToString());
                        mRow.sal_lp = Lib.Conv2Decimal(Dt_Rec.Rows[0]["lev_lp"].ToString());
                        //if (mRow.rec_category == "CONFIRMED" || mRow.rec_category == "TRANSFER")
                        //    mRow.sal_lp = Lib.Conv2Decimal(Dt_Rec.Rows[0]["lev_lp"].ToString());
                        //else
                        //    mRow.sal_lp = 0;
                    }
                }

                List<SalDet> mList = new List<SalDet>();
                SalDet dRow;
                for (int i = 0; i < 12; i++)
                {
                    dRow = new SalDet();
                    dRow.e_caption1 = "";
                    dRow.e_amt1 = 0;
                    dRow.e_visible1 = false;
                    dRow.e_caption2 = "";
                    dRow.e_amt2 = 0;
                    dRow.e_visible2 = false;
                    dRow.d_caption1 = "";
                    dRow.d_amt1 = 0;
                    dRow.d_visible1 = false;
                    dRow.d_caption2 = "";
                    dRow.d_amt2 = 0;
                    dRow.d_visible2 = false;
                    mList.Add(dRow);
                }
                mRow.DetList = GetDetList(mList, mRow);
                Dt_Rec.Rows.Clear();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("mode", smode);
            RetData.Add("record", mRow);
            return RetData;
        }
        
        private List<SalDet> GetDetList(List<SalDet> dList, Salarym mRow)
        {
            DataTable Dt_Head = new DataTable();
            sql = "select sal_code,sal_head from salaryheadm where sal_head is not null order by sal_code";
            Con_Oracle = new DBConnection();
            Dt_Head = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
            foreach (DataRow dr in Dt_Head.Rows)
            {
                switch (dr["SAL_CODE"].ToString().Trim())
                {
                    case "A01":
                        dList[0].e_code1 = dr["SAL_CODE"].ToString(); dList[0].e_caption1 = dr["SAL_HEAD"].ToString(); dList[0].e_amt1 = mRow.a01; dList[0].e_visible1 = true;
                        break;
                    case "A02":
                        dList[1].e_code1 = dr["SAL_CODE"].ToString(); dList[1].e_caption1 = dr["SAL_HEAD"].ToString(); dList[1].e_amt1 = mRow.a02; dList[1].e_visible1 = true;
                        break;
                    case "A03":
                        dList[2].e_code1 = dr["SAL_CODE"].ToString(); dList[2].e_caption1 = dr["SAL_HEAD"].ToString(); dList[2].e_amt1 = mRow.a03; dList[2].e_visible1 = true;
                        break;
                    case "A04":
                        dList[3].e_code1 = dr["SAL_CODE"].ToString(); dList[3].e_caption1 = dr["SAL_HEAD"].ToString(); dList[3].e_amt1 = mRow.a04; dList[3].e_visible1 = true;
                        break;
                    case "A05":
                        dList[4].e_code1 = dr["SAL_CODE"].ToString(); dList[4].e_caption1 = dr["SAL_HEAD"].ToString(); dList[4].e_amt1 = mRow.a05; dList[4].e_visible1 = true;
                        break;
                    case "A06":
                        dList[5].e_code1 = dr["SAL_CODE"].ToString(); dList[5].e_caption1 = dr["SAL_HEAD"].ToString(); dList[5].e_amt1 = mRow.a06; dList[5].e_visible1 = true;
                        break;
                    case "A07":
                        dList[6].e_code1 = dr["SAL_CODE"].ToString(); dList[6].e_caption1 = dr["SAL_HEAD"].ToString(); dList[6].e_amt1 = mRow.a07; dList[6].e_visible1 = true;
                        break;
                    case "A08":
                        dList[7].e_code1 = dr["SAL_CODE"].ToString(); dList[7].e_caption1 = dr["SAL_HEAD"].ToString(); dList[7].e_amt1 = mRow.a08; dList[7].e_visible1 = true;
                        break;
                    case "A09":
                        dList[8].e_code1 = dr["SAL_CODE"].ToString(); dList[8].e_caption1 = dr["SAL_HEAD"].ToString(); dList[8].e_amt1 = mRow.a09; dList[8].e_visible1 = true;
                        break;
                    case "A10":
                        dList[9].e_code1 = dr["SAL_CODE"].ToString(); dList[9].e_caption1 = dr["SAL_HEAD"].ToString(); dList[9].e_amt1 = mRow.a10; dList[9].e_visible1 = true;
                        break;
                    case "A11":
                        dList[10].e_code1 = dr["SAL_CODE"].ToString(); dList[10].e_caption1 = dr["SAL_HEAD"].ToString(); dList[10].e_amt1 = mRow.a11; dList[10].e_visible1 = true;
                        break;
                    case "A12":
                        dList[11].e_code1 = dr["SAL_CODE"].ToString(); dList[11].e_caption1 = dr["SAL_HEAD"].ToString(); dList[11].e_amt1 = mRow.a12; dList[11].e_visible1 = true;
                        break;
                    case "A13":
                        dList[0].e_code2 = dr["SAL_CODE"].ToString(); dList[0].e_caption2 = dr["SAL_HEAD"].ToString(); dList[0].e_amt2 = mRow.a13; dList[0].e_visible2 = true;
                        break;
                    case "A14":
                        dList[1].e_code2 = dr["SAL_CODE"].ToString(); dList[1].e_caption2 = dr["SAL_HEAD"].ToString(); dList[1].e_amt2 = mRow.a14; dList[1].e_visible2 = true;
                        break;
                    case "A15":
                        dList[2].e_code2 = dr["SAL_CODE"].ToString(); dList[2].e_caption2 = dr["SAL_HEAD"].ToString(); dList[2].e_amt2 = mRow.a15; dList[2].e_visible2 = true;
                        break;
                    case "A16":
                        dList[3].e_code2 = dr["SAL_CODE"].ToString(); dList[3].e_caption2 = dr["SAL_HEAD"].ToString(); dList[3].e_amt2 = mRow.a16; dList[3].e_visible2 = true;
                        break;
                    case "A17":
                        dList[4].e_code2 = dr["SAL_CODE"].ToString(); dList[4].e_caption2 = dr["SAL_HEAD"].ToString(); dList[4].e_amt2 = mRow.a17; dList[4].e_visible2 = true;
                        break;
                    case "A18":
                        dList[5].e_code2 = dr["SAL_CODE"].ToString(); dList[5].e_caption2 = dr["SAL_HEAD"].ToString(); dList[5].e_amt2 = mRow.a18; dList[5].e_visible2 = true;
                        break;
                    case "A19":
                        dList[6].e_code2 = dr["SAL_CODE"].ToString(); dList[6].e_caption2 = dr["SAL_HEAD"].ToString(); dList[6].e_amt2 = mRow.a19; dList[6].e_visible2 = true;
                        break;
                    case "A20":
                        dList[7].e_code2 = dr["SAL_CODE"].ToString(); dList[7].e_caption2 = dr["SAL_HEAD"].ToString(); dList[7].e_amt2 = mRow.a20; dList[7].e_visible2 = true;
                        break;
                    case "A21":
                        dList[8].e_code2 = dr["SAL_CODE"].ToString(); dList[8].e_caption2 = dr["SAL_HEAD"].ToString(); dList[8].e_amt2 = mRow.a21; dList[8].e_visible2 = true;
                        break;
                    case "A22":
                        dList[9].e_code2 = dr["SAL_CODE"].ToString(); dList[9].e_caption2 = dr["SAL_HEAD"].ToString(); dList[9].e_amt2 = mRow.a22; dList[9].e_visible2 = true;
                        break;
                    case "A23":
                        dList[10].e_code2 = dr["SAL_CODE"].ToString(); dList[10].e_caption2 = dr["SAL_HEAD"].ToString(); dList[10].e_amt2 = mRow.a23; dList[10].e_visible2 = true;
                        break;
                    case "A24":
                        dList[11].e_code2 = dr["SAL_CODE"].ToString(); dList[11].e_caption2 = dr["SAL_HEAD"].ToString(); dList[11].e_amt2 = mRow.a24; dList[11].e_visible2 = true;
                        break;
                    //case "A25":
                    //    dList[9].e_caption2 = dr["SAL_HEAD"].ToString(); dList[9].e_amt2 = mRow.a25; dList[9].e_visible2 = true;
                    //    break;

                    case "D01":
                        dList[0].d_code1 = dr["SAL_CODE"].ToString(); dList[0].d_caption1 = dr["SAL_HEAD"].ToString(); dList[0].d_amt1 = mRow.d01; dList[0].d_visible1 = true;
                        break;
                    case "D02":
                        dList[1].d_code1 = dr["SAL_CODE"].ToString(); dList[1].d_caption1 = dr["SAL_HEAD"].ToString(); dList[1].d_amt1 = mRow.d02; dList[1].d_visible1 = true;
                        break;
                    case "D03":
                        dList[2].d_code1 = dr["SAL_CODE"].ToString(); dList[2].d_caption1 = dr["SAL_HEAD"].ToString(); dList[2].d_amt1 = mRow.d03; dList[2].d_visible1 = true;
                        break;
                    case "D04":
                        dList[3].d_code1 = dr["SAL_CODE"].ToString(); dList[3].d_caption1 = dr["SAL_HEAD"].ToString(); dList[3].d_amt1 = mRow.d04; dList[3].d_visible1 = true;
                        break;
                    case "D05":
                        dList[4].d_code1 = dr["SAL_CODE"].ToString(); dList[4].d_caption1 = dr["SAL_HEAD"].ToString(); dList[4].d_amt1 = mRow.d05; dList[4].d_visible1 = true;
                        break;
                    case "D06":
                        dList[5].d_code1 = dr["SAL_CODE"].ToString(); dList[5].d_caption1 = dr["SAL_HEAD"].ToString(); dList[5].d_amt1 = mRow.d06; dList[5].d_visible1 = true;
                        break;
                    case "D07":
                        dList[6].d_code1 = dr["SAL_CODE"].ToString(); dList[6].d_caption1 = dr["SAL_HEAD"].ToString(); dList[6].d_amt1 = mRow.d07; dList[6].d_visible1 = true;
                        break;
                    case "D08":
                        dList[7].d_code1 = dr["SAL_CODE"].ToString(); dList[7].d_caption1 = dr["SAL_HEAD"].ToString(); dList[7].d_amt1 = mRow.d08; dList[7].d_visible1 = true;
                        break;
                    case "D09":
                        dList[8].d_code1 = dr["SAL_CODE"].ToString(); dList[8].d_caption1 = dr["SAL_HEAD"].ToString(); dList[8].d_amt1 = mRow.d09; dList[8].d_visible1 = true;
                        break;
                    case "D10":
                        dList[9].d_code1 = dr["SAL_CODE"].ToString(); dList[9].d_caption1 = dr["SAL_HEAD"].ToString(); dList[9].d_amt1 = mRow.d10; dList[9].d_visible1 = true;
                        break;
                    case "D11":
                        dList[10].d_code1 = dr["SAL_CODE"].ToString(); dList[10].d_caption1 = dr["SAL_HEAD"].ToString(); dList[10].d_amt1 = mRow.d11; dList[10].d_visible1 = true;
                        break;
                    case "D12":
                        dList[11].d_code1 = dr["SAL_CODE"].ToString(); dList[11].d_caption1 = dr["SAL_HEAD"].ToString(); dList[11].d_amt1 = mRow.d12; dList[11].d_visible1 = true;
                        break;
                    case "D13":
                        dList[0].d_code2 = dr["SAL_CODE"].ToString(); dList[0].d_caption2 = dr["SAL_HEAD"].ToString(); dList[0].d_amt2 = mRow.d13; dList[0].d_visible2 = true;
                        break;
                    case "D14":
                        dList[1].d_code2 = dr["SAL_CODE"].ToString(); dList[1].d_caption2 = dr["SAL_HEAD"].ToString(); dList[1].d_amt2 = mRow.d14; dList[1].d_visible2 = true;
                        break;
                    case "D15":
                        dList[2].d_code2 = dr["SAL_CODE"].ToString(); dList[2].d_caption2 = dr["SAL_HEAD"].ToString(); dList[2].d_amt2 = mRow.d15; dList[2].d_visible2 = true;
                        break;
                    case "D16":
                        dList[3].d_code2 = dr["SAL_CODE"].ToString(); dList[3].d_caption2 = dr["SAL_HEAD"].ToString(); dList[3].d_amt2 = mRow.d16; dList[3].d_visible2 = true;
                        break;
                    case "D17":
                        dList[4].d_code2 = dr["SAL_CODE"].ToString(); dList[4].d_caption2 = dr["SAL_HEAD"].ToString(); dList[4].d_amt2 = mRow.d17; dList[4].d_visible2 = true;
                        break;
                    case "D18":
                        dList[5].d_code2 = dr["SAL_CODE"].ToString(); dList[5].d_caption2 = dr["SAL_HEAD"].ToString(); dList[5].d_amt2 = mRow.d18; dList[5].d_visible2 = true;
                        break;
                    case "D19":
                        dList[6].d_code2 = dr["SAL_CODE"].ToString(); dList[6].d_caption2 = dr["SAL_HEAD"].ToString(); dList[6].d_amt2 = mRow.d19; dList[6].d_visible2 = true;
                        break;
                    case "D20":
                        dList[7].d_code2 = dr["SAL_CODE"].ToString(); dList[7].d_caption2 = dr["SAL_HEAD"].ToString(); dList[7].d_amt2 = mRow.d20; dList[7].d_visible2 = true;
                        break;
                    case "D21":
                        dList[8].d_code2 = dr["SAL_CODE"].ToString(); dList[8].d_caption2 = dr["SAL_HEAD"].ToString(); dList[8].d_amt2 = mRow.d21; dList[8].d_visible2 = true;
                        break;
                    case "D22":
                        dList[9].d_code2 = dr["SAL_CODE"].ToString(); dList[9].d_caption2 = dr["SAL_HEAD"].ToString(); dList[9].d_amt2 = mRow.d22; dList[9].d_visible2 = true;
                        break;
                    case "D23":
                        dList[10].d_code2 = dr["SAL_CODE"].ToString(); dList[10].d_caption2 = dr["SAL_HEAD"].ToString(); dList[10].d_amt2 = mRow.d23; dList[10].d_visible2 = true;
                        break;
                    case "D24":
                        dList[11].d_code2 = dr["SAL_CODE"].ToString(); dList[11].d_caption2 = dr["SAL_HEAD"].ToString(); dList[11].d_amt2 = mRow.d24; dList[11].d_visible2 = true;
                        break;
                        //case "D25":
                        //    dList[9].d_caption2 = dr["SAL_HEAD"].ToString(); dList[9].d_amt2 = mRow.d25; dList[9].d_visible2 = true;
                        //    break;
                }

            }

            return dList;
        }
        public string AllValid(Salarym Record)
        {
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {

                sql = " select sal_pkid from salarym ";
                sql += " where sal_pkid = '" + Record.sal_pkid + "'";
                sql += " and sal_edit_code is null";
                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " | Details Closed, Can't Edit ");

                sql = " select salh_pkid from salaryh ";
                sql += " where rec_company_code = '" + Record.rec_company_code + "'";
                sql += " and rec_branch_code = '" + Record.rec_branch_code + "'";
                sql += " and salh_year = " + Record.sal_year;
                sql += " and salh_month = " + Record.sal_month;
                sql += " and nvl(salh_posted, 'N') = 'Y' ";
                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " | Already Posted, Can't Edit ");


                //Boolean bRet = true;
                //DataTable Dt_locked = new DataTable();
                //string sql = "";
                //sql = " select rec_locked from salarym where ";
                //sql += " sal_pkid='" + drow["SAL_PKID"].ToString() + "'";
                //Dt_locked = orConnection.RunSql(sql);
                //if (Dt_locked.Rows.Count > 0)
                //    if (Dt_locked.Rows[0]["REC_LOCKED"].ToString().Trim() == "Y")
                //    {
                //        bRet = false;
                //        MessageBox.Show("Details Closed, Can't Edit", "Payroll");
                //        return bRet;
                //    }
                //return bRet;



                                 //if (Record.sal_code.Trim().Length <= 0)
                                 //    Lib.AddError(ref str, " | Code Cannot Be Empty");

                                 //if (Record.sal_code.Trim().Length > 0)
                                 //{

                                 //    sql = "select sal_pkid from (";
                                 //    sql += "select sal_pkid  from salaryheadm a where (a.sal_code = '{CODE}')  ";
                                 //    sql += ") a where sal_pkid <> '{PKID}'";

                                 //    sql = sql.Replace("{CODE}", Record.sal_code);
                                 //    sql = sql.Replace("{PKID}", Record.sal_pkid);

                                 //    if (Con_Oracle.IsRowExists(sql))
                                 //        Lib.AddError(ref str, " | Code Exists");
                                 //}

                                 //if (Record.sal_desc.Trim().Length > 0)
                                 //{

                                 //    sql = "select sal_pkid from (";
                                 //    sql += "select sal_pkid  from salaryheadm a where (a.sal_desc = '{NAME}')  ";
                                 //    sql += ") a where sal_pkid <> '{PKID}'";

                                 //    sql = sql.Replace("{NAME}", Record.sal_desc);
                                 //    sql = sql.Replace("{PKID}", Record.sal_pkid);


                                 //    if (Con_Oracle.IsRowExists(sql))

                                 //        Lib.AddError(ref str, " | Description Exists");
                                 //}

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(Salarym Record)
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
                Rec.CreateRow("salarym", Record.rec_mode, "sal_pkid", Record.sal_pkid);
                Rec.InsertNumeric("a01", Record.a01.ToString());
                Rec.InsertNumeric("a02", Record.a02.ToString());
                Rec.InsertNumeric("a03", Record.a03.ToString());
                Rec.InsertNumeric("a04", Record.a04.ToString());
                Rec.InsertNumeric("a05", Record.a05.ToString());
                Rec.InsertNumeric("a06", Record.a06.ToString());
                Rec.InsertNumeric("a07", Record.a07.ToString());
                Rec.InsertNumeric("a08", Record.a08.ToString());
                Rec.InsertNumeric("a09", Record.a09.ToString());
                Rec.InsertNumeric("a10", Record.a10.ToString());
                Rec.InsertNumeric("a11", Record.a11.ToString());
                Rec.InsertNumeric("a12", Record.a12.ToString());
                Rec.InsertNumeric("a13", Record.a13.ToString());
                Rec.InsertNumeric("a14", Record.a14.ToString());
                Rec.InsertNumeric("a15", Record.a15.ToString());
                Rec.InsertNumeric("a16", Record.a16.ToString());
                Rec.InsertNumeric("a17", Record.a17.ToString());
                Rec.InsertNumeric("a18", Record.a18.ToString());
                Rec.InsertNumeric("a19", Record.a19.ToString());
                Rec.InsertNumeric("a20", Record.a20.ToString());
                Rec.InsertNumeric("a21", Record.a21.ToString());
                Rec.InsertNumeric("a22", Record.a22.ToString());
                Rec.InsertNumeric("a23", Record.a23.ToString());
                Rec.InsertNumeric("a24", Record.a24.ToString());
                Rec.InsertNumeric("a25", Record.a25.ToString());
                Rec.InsertNumeric("d01", Record.d01.ToString());
                Rec.InsertNumeric("d02", Record.d02.ToString());
                Rec.InsertNumeric("d03", Record.d03.ToString());
                Rec.InsertNumeric("d04", Record.d04.ToString());
                Rec.InsertNumeric("d05", Record.d05.ToString());
                Rec.InsertNumeric("d06", Record.d06.ToString());
                Rec.InsertNumeric("d07", Record.d07.ToString());
                Rec.InsertNumeric("d08", Record.d08.ToString());
                Rec.InsertNumeric("d09", Record.d09.ToString());
                Rec.InsertNumeric("d10", Record.d10.ToString());
                Rec.InsertNumeric("d11", Record.d11.ToString());
                Rec.InsertNumeric("d12", Record.d12.ToString());
                Rec.InsertNumeric("d13", Record.d13.ToString());
                Rec.InsertNumeric("d14", Record.d14.ToString());
                Rec.InsertNumeric("d15", Record.d15.ToString());
                Rec.InsertNumeric("d16", Record.d16.ToString());
                Rec.InsertNumeric("d17", Record.d17.ToString());
                Rec.InsertNumeric("d18", Record.d18.ToString());
                Rec.InsertNumeric("d19", Record.d19.ToString());
                Rec.InsertNumeric("d20", Record.d20.ToString());
                Rec.InsertNumeric("d21", Record.d21.ToString());
                Rec.InsertNumeric("d22", Record.d22.ToString());
                Rec.InsertNumeric("d23", Record.d23.ToString());
                Rec.InsertNumeric("d24", Record.d24.ToString());
                Rec.InsertNumeric("d25", Record.d25.ToString());
                Rec.InsertNumeric("sal_lop_amt", Record.sal_lop_amt.ToString());
                Rec.InsertNumeric("sal_gross_earn", Record.sal_gross_earn.ToString());
                Rec.InsertNumeric("sal_gross_deduct", Record.sal_gross_deduct.ToString());
                Rec.InsertNumeric("sal_net", Record.sal_net.ToString());

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("sal_emp_id", Record.sal_emp_id);
                    Rec.InsertDate("sal_date", Record.sal_date);
                    Rec.InsertNumeric("sal_month", Record.sal_month.ToString());
                    Rec.InsertNumeric("sal_year", Record.sal_year.ToString());
                    Rec.InsertNumeric("sal_fin_year", Record.sal_fin_year.ToString());
                    Rec.InsertNumeric("sal_days_worked", Record.sal_days_worked.ToString());
                    Rec.InsertNumeric("sal_basic_rt", Record.sal_basic_rt.ToString());
                    Rec.InsertNumeric("sal_da_rt", Record.sal_da_rt.ToString());
                    Rec.InsertString("sal_pf_mon_year", Record.sal_pf_mon_year);
                    Rec.InsertNumeric("sal_pf_limit", Record.sal_pf_limit.ToString());
                    Rec.InsertNumeric("sal_pf_cel_limit", Record.sal_pf_cel_limit.ToString());
                    Rec.InsertNumeric("sal_pf_cel_limit_amt", Record.sal_pf_cel_limit_amt.ToString());
                    Rec.InsertNumeric("sal_pf_bal", Record.sal_pf_bal.ToString());
                    Rec.InsertNumeric("sal_pf_wage_bal", Record.sal_pf_wage_bal.ToString());
                    Rec.InsertNumeric("sal_pf_base", Record.sal_pf_base.ToString());
                    Rec.InsertNumeric("sal_pf_emplr", Record.sal_pf_emplr.ToString());
                    Rec.InsertNumeric("sal_pf_emplr_share", Record.sal_pf_emplr_share.ToString());
                    Rec.InsertNumeric("sal_pf_emplr_pension", Record.sal_pf_emplr_pension.ToString());
                    Rec.InsertNumeric("sal_pf_emplr_pension_per", Record.sal_pf_emplr_pension_per.ToString());
                    Rec.InsertNumeric("sal_pf_eps_amt", Record.sal_pf_eps_amt.ToString());
                    Rec.InsertNumeric("sal_admin_per", Record.sal_admin_per.ToString());
                    Rec.InsertNumeric("sal_admin_amt", Record.sal_admin_amt.ToString());
                    Rec.InsertString("sal_admin_based_on", Record.sal_admin_based_on);
                    Rec.InsertNumeric("sal_edli_per", Record.sal_edli_per.ToString());
                    Rec.InsertNumeric("sal_edli_amt", Record.sal_edli_amt.ToString());
                    Rec.InsertString("sal_edli_based_on", Record.sal_edli_based_on);
                    Rec.InsertString("sal_is_esi", Record.sal_is_esi == true ? "Y" : "N");
                    Rec.InsertNumeric("sal_esi_base", Record.sal_esi_base.ToString());
                    Rec.InsertNumeric("sal_esi_emplr_per", Record.sal_esi_emplr_per.ToString());
                    Rec.InsertNumeric("sal_esi_limit", Record.sal_esi_limit.ToString());
                    Rec.InsertNumeric("sal_esi_gov_share", Record.sal_esi_gov_share.ToString());
                    Rec.InsertDate("sal_pay_date", Record.sal_pay_date);
                    Rec.InsertNumeric("sal_work_days", Record.sal_work_days.ToString());
                    Rec.InsertString("sal_mail_sent", "N");
                    Rec.InsertString("sal_edit_code", "{S}");
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

                Lib.FindLoPAmount(Record.sal_emp_id, Lib.Conv2Integer(Record.sal_year.ToString()), Lib.Conv2Integer(Record.sal_month.ToString()), Lib.Conv2Decimal(Record.sal_lp.ToString()), Lib.Conv2Decimal(Record.sal_days_worked.ToString()));
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
        public string GenerateValid(int syear, int smonth, string brcode, string compcode, string year_start_date, string year_end_date,string salpkids)
        {
            string str = "";
           DateTime tdate = DateTime.Now;
            try
            {
                string MsgMnth = "";
                int cMonth = 0, cYear = 0;
                int tempMonth = 0, tempYear = 0;

                tdate = new DateTime(syear, smonth, 1);
                if (!Lib.IsInFinYear(tdate.ToString("yyyy-MM-dd"), year_start_date,year_end_date, false))
                {
                    Lib.AddError(ref str, " | Unable to Generate (Salary Month or Year not in Financial Year)");
                }

                DataTable Dt_TEMP;
                cMonth = smonth;
                cYear = syear;

                sql = " select distinct sal_month ,sal_year ";
                sql += " from salarym a ";
                sql += " where a.rec_company_code = '" + compcode + "'";
                sql += " and a.rec_branch_code = '" + brcode + "'";
                sql += " and sal_month<>0 and sal_year<>0";
                sql += " order by sal_year desc ,sal_month desc ";
                Dt_TEMP = new DataTable();
                Dt_TEMP = Con_Oracle.ExecuteQuery(sql);
                if (Dt_TEMP.Rows.Count > 0)
                {
                    tempMonth = Lib.Conv2Integer(Dt_TEMP.Rows[0]["SAL_MONTH"].ToString());
                    tempYear = Lib.Conv2Integer(Dt_TEMP.Rows[0]["SAL_YEAR"].ToString());
                    if (tempMonth == 12)
                    {
                        tempMonth = 1;
                        tempYear = tempYear + 1;
                    }
                    else
                        tempMonth++;
                    if (cYear > 0 && tempYear > 0)
                    {
                        DateTime dtFeed = new DateTime(cYear, cMonth, 01);
                        DateTime dtNextGenert = new DateTime(tempYear, tempMonth, 1);
                        tempMonth = Lib.Conv2Integer(Dt_TEMP.Rows[Dt_TEMP.Rows.Count - 1]["SAL_MONTH"].ToString());
                        tempYear = Lib.Conv2Integer(Dt_TEMP.Rows[Dt_TEMP.Rows.Count - 1]["SAL_YEAR"].ToString());
                        DateTime dtFirstGenert = new DateTime(tempYear, tempMonth, 1);

                        if (dtFeed < dtFirstGenert)
                        {
                            Lib.AddError(ref str, " | Invalid Generate");
                        }
                        if (dtFeed > dtNextGenert)
                        {
                            Lib.AddError(ref str, " | Previous not Generated");
                        }
                    }
                }

                sql = " select sal_pkid from salarym a ";
                sql += " where a.rec_company_code = '" + compcode + "'";
                sql += " and a.rec_branch_code = '" + brcode + "'";
                sql += " and sal_month=" + cMonth;
                sql += " and sal_year=" + cYear;
                sql += " and sal_edit_code is null";
                if (salpkids != "")//during re-generation
                {
                    salpkids = salpkids.Replace(",", "','");
                    sql += " and sal_pkid in ('" + salpkids + "')";
                }
                Dt_TEMP = new DataTable();
                Dt_TEMP = Con_Oracle.ExecuteQuery(sql);
                if (Dt_TEMP.Rows.Count > 0)
                {
                    if (cMonth.ToString().Length == 1)
                        MsgMnth = "0" + cMonth.ToString();
                    else
                        MsgMnth = cMonth.ToString();
                    Lib.AddError(ref str, " | Month(" + MsgMnth + ", " + cYear.ToString() + ") Closed");
                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }
        public IDictionary<string, object> Generate(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            List<Salarym> mList = new List<Salarym>();
            Salarym mRow;
            string ErrorMessage = "";
            string SQL1, SQL2 = "";
            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            string user_code = SearchData["user_code"].ToString();
            string salpkids = SearchData["salpkids"].ToString();
            int salmonth = Lib.Conv2Integer(SearchData["salmonth"].ToString());
            int salyear = Lib.Conv2Integer(SearchData["salyear"].ToString());
            int DaysInMonth = 0;
            string Emp_Ids = "";
            string year_start_date = SearchData["year_start_date"].ToString(); 
            string year_end_date = SearchData["year_end_date"].ToString();
            DataRow Dr_PS = null;
            bool bTrans = false;
            string pf_excluded_Cols = "";
            decimal pf_excluded_Amt = 0;
            DateTime dtime;
            int levlpdays = 0;
            int levdaysWrkd = 0;
            bool bSalaryH = false;
            string SalDate = "";
            try
            {
                if (salmonth > 0 && salyear > 0)
                    DaysInMonth = DateTime.DaysInMonth(salyear, salmonth);

                if ((ErrorMessage = GenerateValid(salyear, salmonth, branch_code, company_code, year_start_date, year_end_date,salpkids)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }


                DataTable Dt_List = new DataTable();

                sql = "select sal_emp_id from salarym ";
                sql += " where rec_company_code = '" + company_code + "'";
                sql += " and rec_branch_code = '" + branch_code + "'";
                sql += " and sal_month = " + salmonth.ToString();
                sql += " and sal_year = " + salyear.ToString();
                if (salpkids != "")//during re-generation
                {
                    salpkids = salpkids.Replace(",", "','");
                    sql += " and sal_pkid not in ('" + salpkids + "')";
                }

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow dr in Dt_List.Rows)
                {
                    if (Emp_Ids != "")
                        Emp_Ids += ",";
                    Emp_Ids += dr["sal_emp_id"].ToString();
                }
                Emp_Ids = Emp_Ids.Replace(",", "','");

                sql = "";
                sql += " select emp_pkid,emp_no,emp_name,emp_do_joining,emp_do_relieve,emp_is_retired, b.param_name as emp_grade,emp_branch_group ";
                sql += " ,js.param_name as emp_job_status";
                sql += " ,sm.* ";
                sql += " from empm a ";
                sql += " left join salarym sm on (a.emp_pkid=sm.sal_emp_id and sm.sal_month=0 and sm.sal_year=0 ) ";
                sql += " left join param b on a.emp_grade_id = b.param_pkid ";
                sql += " left join param js on (a.emp_status_id = js.param_pkid)";
                sql += " where a.rec_company_code = '" + company_code + "'";
                sql += " and a.rec_branch_code = '" + branch_code + "'";
                sql += " and sm.sal_gross_earn > 0 ";
                sql += " and nvl(emp_in_payroll,'N')='Y'";
                if (Emp_Ids != "")
                    sql += " and a.emp_pkid not in ('" + Emp_Ids + "')";
                sql += " order by emp_no";

                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);

                if (type == "LIST")
                {
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mRow = new Salarym();
                        mRow.sal_emp_id = Dr["emp_pkid"].ToString();
                        mRow.sal_emp_code = Dr["emp_no"].ToString();
                        mRow.sal_emp_name = Dr["emp_name"].ToString();
                        mRow.sal_emp_grade = Dr["emp_grade"].ToString();
                        mRow.sal_emp_do_joining = Lib.DatetoStringDisplayformat(Dr["emp_do_joining"]);
                        mRow.sal_gross_earn = Lib.Conv2Decimal(Dr["sal_gross_earn"].ToString());
                        mRow.sal_gross_deduct = Lib.Conv2Decimal(Dr["sal_gross_deduct"].ToString());
                        mRow.sal_net = Lib.Conv2Decimal(Dr["sal_net"].ToString());
                        mList.Add(mRow);
                    }
                }
                else
                {
                    if (salpkids != "")//During Regeneration
                    {
                        sql = "delete from salarym where sal_pkid in ('" + salpkids + "')";
                        Con_Oracle.BeginTransaction();
                        bTrans = true;
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                        bTrans = false;
                    }

                    sql = " select ps_pkid,ps_admin_per,ps_admin_amt,ps_admin_based_on";
                    sql += " ,ps_edli_per,ps_edli_amt,ps_edli_based_on,ps_esi_emplr_per";
                    sql += " ,ps_esi_limit,ps_pf_emplr_pension_per,ps_pf_cel_limit";
                    sql += " ,ps_pf_cel_limit_amt,ps_esi_emply_per,ps_pf_per,ps_pf_col_excluded ";
                    sql += " from payroll_setting a ";
                    sql += " where a.rec_company_code = '" + company_code + "'";
                    sql += " and a.rec_branch_code = '" + branch_code + "'";
                    DataTable Dt_PS = new DataTable();
                    Dt_PS = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_PS.Rows.Count > 0)
                        Dr_PS = Dt_PS.Rows[0];
                    //Dt_PS.Rows.Clear();

                    if (salmonth > 0 && salyear > 0)
                        SalDate = Lib.StringToDate(new DateTime(salyear, salmonth, 1));
                    else
                        SalDate = "NULL";

                    bSalaryH = false;
                    foreach (DataRow dr in Dt_List.Rows)
                    {
                        bTrans = false;
                        if (!bSalaryH)
                        {
                            SaveSalaryH(dr["REC_COMPANY_CODE"].ToString(), dr["REC_BRANCH_CODE"].ToString(), SalDate, salmonth, salyear, year_code);
                            bSalaryH = true;
                        }

                        if (dr["EMP_DO_JOINING"].ToString().Trim() != "")
                        {
                            dtime = (DateTime)dr["EMP_DO_JOINING"];
                            if (dtime.Month == salmonth && dtime.Year == salyear && dtime.Day != 1)
                            {
                                levlpdays = dtime.Day - 1;
                                levdaysWrkd = DaysInMonth - levlpdays;
                                SaveLeave(dr, salmonth, salyear, levlpdays, levdaysWrkd, year_code, user_code);
                            }
                        }

                        if (dr["EMP_DO_RELIEVE"].ToString().Trim() != "")
                        {
                            dtime = (DateTime)dr["EMP_DO_RELIEVE"];
                            if (dtime.Month == salmonth && dtime.Year == salyear && dtime.Day != DaysInMonth)
                            {
                                levlpdays = DaysInMonth - dtime.Day ;
                                levdaysWrkd = DaysInMonth - levlpdays;
                                SaveLeave(dr, salmonth, salyear, levlpdays, levdaysWrkd, year_code, user_code);
                            }
                        }

                        SQL1 = "  Insert into SALARYM ";
                        SQL1 += "    (SAL_PKID";
                        SQL2 = " Values ('" + Guid.NewGuid().ToString().ToUpper() + "'";
                        SQL1 += "    ,SAL_EMP_ID";
                        SQL2 += " ,'" + dr["SAL_EMP_ID"].ToString() + "'";
                        SQL1 += "    ,SAL_DATE";
                        if (salmonth > 0 && salyear > 0)
                            SQL2 += " ,'" + SalDate + "'";
                        else
                            SQL2 += " ,Null ";
                        SQL1 += "    ,SAL_MONTH";
                        SQL2 += " ," + salmonth.ToString();
                        SQL1 += "    ,SAL_YEAR";
                        SQL2 += " ," + salyear.ToString();
                        SQL1 += "    ,SAL_DAYS_WORKED";
                        SQL2 += " ," + DaysInMonth;
                        SQL1 += "    ,SAL_WORK_DAYS";
                        SQL2 += " ," + DaysInMonth;
                        SQL1 += "    ,SAL_FIN_YEAR";
                        SQL2 += " ," + year_code;
                        SQL1 += " ,SAL_ESI_GOV_SHARE";
                        SQL2 += " , 0 ";
                        SQL1 += "    ,SAL_BASIC_RT";
                        SQL2 += "," + (Lib.Convert2Decimal(dr["A01"].ToString()) + Lib.Convert2Decimal(dr["A11"].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString()));
                        SQL1 += "    ,SAL_DA_RT";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A02"].ToString());
                        SQL1 += "    ,A01";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A01"].ToString());
                        SQL1 += "    ,A02";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A02"].ToString());
                        SQL1 += "    ,A03";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A03"].ToString());
                        SQL1 += "    ,A04";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A04"].ToString());
                        SQL1 += "    ,A05";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A05"].ToString());
                        SQL1 += "    ,A06";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A06"].ToString());
                        SQL1 += "    ,A07";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A07"].ToString());
                        SQL1 += "    ,A08";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A08"].ToString());
                        SQL1 += "    ,A09";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A09"].ToString());
                        SQL1 += "    ,A10";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A10"].ToString());
                        SQL1 += "    ,A11";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A11"].ToString());
                        SQL1 += "    ,A12";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A12"].ToString());
                        SQL1 += "    ,A13";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A13"].ToString());
                        SQL1 += "    ,A14";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A14"].ToString());
                        SQL1 += "    ,A15";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A15"].ToString());
                        SQL1 += "    ,A16";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A16"].ToString());
                        SQL1 += "    ,A17";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A17"].ToString());
                        SQL1 += "    ,A18";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A18"].ToString());
                        SQL1 += "    ,A19";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A19"].ToString());
                        SQL1 += "    ,A20";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A20"].ToString());
                        SQL1 += "    ,A21";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A21"].ToString());
                        SQL1 += "    ,A22";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A22"].ToString());
                        SQL1 += "    ,A23";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A23"].ToString());
                        SQL1 += "    ,A24";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A24"].ToString());
                        SQL1 += "    ,A25";
                        SQL2 += "," + Lib.Convert2Decimal(dr["A25"].ToString());

                        SQL1 += "    ,D01";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D01"].ToString());
                        SQL1 += "    ,D02";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D02"].ToString());
                        SQL1 += "    ,D03";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D03"].ToString());
                        SQL1 += "    ,D04";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D04"].ToString());
                        SQL1 += "    ,D05";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D05"].ToString());
                        SQL1 += "    ,D06";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D06"].ToString());
                        SQL1 += "    ,D07";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D07"].ToString());
                        SQL1 += "    ,D08";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D08"].ToString());
                        SQL1 += "    ,D09";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D09"].ToString());
                        SQL1 += "    ,D10";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D10"].ToString());
                        SQL1 += "    ,D11";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D11"].ToString());
                        SQL1 += "    ,D12";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D12"].ToString());
                        SQL1 += "    ,D13";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D13"].ToString());
                        SQL1 += "    ,D14";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D14"].ToString());
                        SQL1 += "    ,D15";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D15"].ToString());
                        SQL1 += "    ,D16";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D16"].ToString());
                        SQL1 += "    ,D17";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D17"].ToString());
                        SQL1 += "    ,D18";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D18"].ToString());
                        SQL1 += "    ,D19";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D19"].ToString());
                        SQL1 += "    ,D20";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D20"].ToString());
                        SQL1 += "    ,D21";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D21"].ToString());
                        SQL1 += "    ,D22";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D22"].ToString());
                        SQL1 += "    ,D23";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D23"].ToString());
                        SQL1 += "    ,D24";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D24"].ToString());
                        SQL1 += "    ,D25";
                        SQL2 += "," + Lib.Convert2Decimal(dr["D25"].ToString());

                        SQL1 += "    ,SAL_PF_LIMIT";
                        SQL2 += "," + Lib.Convert2Decimal(dr["SAL_PF_LIMIT"].ToString());
                        SQL1 += "    ,SAL_GROSS_EARN";
                        SQL2 += "," + Lib.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString());
                        SQL1 += "    ,SAL_GROSS_DEDUCT";
                        SQL2 += "," + Lib.Convert2Decimal(dr["SAL_GROSS_DEDUCT"].ToString());
                        SQL1 += "    ,SAL_NET";
                        SQL2 += "," + Lib.Convert2Decimal(dr["SAL_NET"].ToString());

                        SQL1 += " ,SAL_LOP_AMT";
                        SQL2 += " ,0 ";

                        SQL1 += "    ,SAL_EMP_BRANCH_GROUP";
                        SQL2 += "," + Lib.Conv2Integer(dr["EMP_BRANCH_GROUP"].ToString());

                        SQL1 += "    ,SAL_IS_RETIRED";
                        SQL2 += " ,'" + dr["EMP_IS_RETIRED"].ToString() + "'";

                        //SQL1 += "    ,SAL_PF_BASE";
                        //SQL2 += "," + Lib.Convert2Decimal(dr["SAL_PF_BASE"].ToString());
                        //SQL1 += "    ,SAL_PF_EMPLR";
                        //SQL2 += "," + Lib.Convert2Decimal(dr["SAL_PF_EMPLR"].ToString());
                        //SQL1 += "    ,SAL_PF_EMPLR_SHARE";
                        //SQL2 += "," + Lib.Convert2Decimal(dr["SAL_PF_EMPLR_SHARE"].ToString());
                        //SQL1 += "    ,SAL_PF_EMPLR_PENSION";
                        //SQL2 += "," + Lib.Convert2Decimal(dr["SAL_PF_EMPLR_PENSION"].ToString());
                        //SQL1 += "    ,SAL_PF_EPS_AMT";
                        //SQL2 += "," + Lib.Convert2Decimal(dr["SAL_PF_EPS_AMT"].ToString());

                        pf_excluded_Cols = "";
                        if (Dr_PS != null)
                        {
                            pf_excluded_Cols = Dr_PS["ps_pf_col_excluded"].ToString();
                            SQL1 += "    ,SAL_ADMIN_BASED_ON";
                            SQL2 += ",'" + Dr_PS["PS_ADMIN_BASED_ON"].ToString() + "'";

                            SQL1 += "    ,SAL_ADMIN_PER";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_ADMIN_PER"].ToString());

                            SQL1 += "    ,SAL_ADMIN_AMT";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_ADMIN_AMT"].ToString());

                            SQL1 += "    ,SAL_EDLI_BASED_ON";
                            SQL2 += ",'" + Dr_PS["PS_EDLI_BASED_ON"].ToString() + "'";

                            SQL1 += "    ,SAL_EDLI_PER";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_EDLI_PER"].ToString());

                            SQL1 += "    ,SAL_EDLI_AMT";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_EDLI_AMT"].ToString());

                            SQL1 += "    ,SAL_ESI_EMPLY_PER";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_ESI_EMPLY_PER"].ToString());

                            SQL1 += "    ,SAL_ESI_EMPLR_PER";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_ESI_EMPLR_PER"].ToString());

                            SQL1 += "    ,SAL_PF_PER";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_PF_PER"].ToString());

                            SQL1 += "    ,SAL_PF_EMPLR_PENSION_PER";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_PF_EMPLR_PENSION_PER"].ToString());

                            SQL1 += "    ,SAL_PF_CEL_LIMIT";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_PF_CEL_LIMIT"].ToString());

                            SQL1 += "    ,SAL_PF_CEL_LIMIT_AMT";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_PF_CEL_LIMIT_AMT"].ToString());

                            SQL1 += "    ,SAL_ESI_LIMIT";
                            SQL2 += "," + Lib.Convert2Decimal(Dr_PS["PS_ESI_LIMIT"].ToString());
                        }

                        SQL1 += "   ,REC_CATEGORY ";
                        if (dr["emp_job_status"].ToString() == "CONFIRMED" || dr["emp_job_status"].ToString() == "TRANSFER" || dr["emp_job_status"].ToString() == "SAL PROCESS AT HO")
                            SQL2 += ",'CONFIRMED' ";
                        else
                            SQL2 += ",'UNCONFIRM' ";
                        SQL1 += "    ,SAL_IS_ESI";
                        if (Lib.Convert2Decimal(dr["D02"].ToString()) > 0)
                            SQL2 += ",'Y'";
                        else
                            SQL2 += ",'N'";

                        SQL1 += "    ,SAL_EDIT_CODE";
                        SQL2 += " ,'{S}'";

                        SQL1 += "    ,REC_COMPANY_CODE";
                        SQL2 += " ,'" + dr["REC_COMPANY_CODE"].ToString() + "'";
                        SQL1 += "    ,REC_BRANCH_CODE ";
                        SQL2 += " ,'" + dr["REC_BRANCH_CODE"].ToString() + "'";
                        SQL1 += "    ,REC_CREATED_BY ";
                        SQL2 += " ,'" + user_code + "'";
                        SQL1 += "    ,REC_CREATED_DATE )";
                        SQL2 += ",(SYSDATE))";

                        Con_Oracle.BeginTransaction();
                        sql = SQL1 + SQL2; bTrans = true;
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();

                        pf_excluded_Amt = 0;
                        if (pf_excluded_Cols.Trim().Length > 0)
                        {
                            string[] sdata = pf_excluded_Cols.Split(',');
                            foreach (string scol in sdata)
                                if (scol.Trim().Length > 0)
                                {
                                    pf_excluded_Amt += Lib.Conv2Decimal(dr[scol.Trim()].ToString());
                                }
                        }

                        UpdateMonthEPF_ESI(dr["SAL_EMP_ID"].ToString(), salyear, salmonth, pf_excluded_Amt, company_code);
                        CalculateLop(dr["SAL_EMP_ID"].ToString(), salyear, salmonth,company_code );
                        bool bESIGovShare = IsESIGovShare(Lib.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString()), Lib.Convert2Decimal(dr["D02"].ToString()), salmonth, salyear);
                        if (bESIGovShare)
                        {
                            sql = " update salarym set ";
                            sql += " sal_gross_deduct = " + (Lib.Convert2Decimal(dr["SAL_GROSS_DEDUCT"].ToString()) - Lib.Convert2Decimal(dr["D02"].ToString())).ToString();
                            sql += " ,sal_net = " + (Lib.Convert2Decimal(dr["SAL_NET"].ToString()) + Lib.Convert2Decimal(dr["D02"].ToString())).ToString();
                            sql += " ,sal_esi_gov_share = " + Lib.Convert2Decimal(dr["D02"].ToString()).ToString();
                            sql += " ,d02 = 0 ";
                            sql += " where rec_company_code ='" + company_code + "'";
                            sql += " and sal_emp_id ='" + dr["sal_emp_id"].ToString() + "'";
                            sql += " and sal_month =" + salmonth;
                            sql += " and sal_year =" + salyear;
                            Con_Oracle.BeginTransaction();
                            bTrans = true;
                            Con_Oracle.ExecuteNonQuery(sql);
                            Con_Oracle.CommitTransaction();
                        }
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
        public Dictionary<string, object> DeleteRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            DataTable Dt_Test = new DataTable();
            try
            {
                string id = SearchData["pkid"].ToString();
                string comp_code = SearchData["comp_code"].ToString();
                string branch_code = SearchData["branch_code"].ToString();
                string user_code = SearchData["user_code"].ToString();
                string type = "";
                if (SearchData.ContainsKey("type"))
                    type = SearchData["type"].ToString();

                Con_Oracle = new DBConnection();

                if (type == "PAYROLL")
                    sql = " delete from salarym where sal_pkid ='" + id + "'";
                else
                    sql = " update empm set emp_in_payroll='N' where emp_pkid ='" + id + "'";

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
        private void UpdateMonthEPF_ESI(string EMP_ID, int Year, int Month, decimal PFExculdedAmt, string CompCode)
        {
            DataTable Dt_SalMonth;
            DataRow Dr = null;
            int DaysInMonth = 30;
            try
            {

                if (Month > 0 && Year > 0)
                {

                    DaysInMonth = DateTime.DaysInMonth(Year, Month);

                    //sal_pf_limit - PF Special Limit
                    //sal_pf_cel_limit - PF General Limit

                    sql = "    select sal_pkid,sal_days_worked,sal_gross_earn,sal_gross_deduct,sal_pf_cel_limit,sal_pf_cel_limit_amt,sal_pf_emplr_pension_per,";
                    sql += "   Round(case when nvl(sal_pf_limit,0) > 0 then sal_pf_limit else case when (sal_gross_earn - {A04}) > sal_pf_cel_limit then sal_pf_cel_limit else (sal_gross_earn - {A04}) end end ) as PF_Base,";
                    sql += "   Round(case when nvl(sal_is_retired,'N') ='Y' then 0 else case when (sal_gross_earn - {A04}) > sal_pf_cel_limit then sal_pf_cel_limit_amt else (sal_gross_earn - {A04})*(sal_pf_emplr_pension_per/100) end end) as PENSION, ";// Pension=if(Basic+DA)>6500 then 541 (Max Limit) else (Basic+DA)* 8.33/100
                    sql += "   case when (sal_gross_earn - {A04})>sal_pf_cel_limit then sal_pf_cel_limit else (sal_gross_earn - {A04}) end  as EPS_AMT, ";
                    sql += "   D01 as SAL_PF_EMPLR, "; //Actually D01 is employee PF deduction which is same as Employer Contribution
                    sql += "   sal_admin_based_on,sal_admin_per,sal_admin_amt,sal_edli_based_on,sal_edli_per,sal_edli_amt,sal_esi_emplr_per,D02,";
                    sql += "   Round(case when ((nvl(sal_gross_earn,0) + nvl(sal_lop_amt,0)) > sal_esi_limit or nvl(sal_is_esi,'N') = 'Y') then 0 else sal_gross_earn end) as sal_esi_base  ";
                    sql += "   from salarym ";
                    sql += "   where rec_company_code ='" + CompCode + "'";
                    sql += "   and sal_emp_id ='" + EMP_ID + "'";
                    sql += "   and sal_month =" + Month;
                    sql += "   and sal_year =" + Year;

                    sql = sql.Replace("{A04}", PFExculdedAmt.ToString());

                    Dt_SalMonth = new DataTable();
                    Dt_SalMonth = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_SalMonth.Rows.Count > 0)
                    {
                        Dr = Dt_SalMonth.Rows[0];
                        decimal daysWorkd = Lib.Convert2Decimal(Dr["SAL_DAYS_WORKED"].ToString());
                        if (daysWorkd > DaysInMonth)
                            daysWorkd = DaysInMonth;
                        if (daysWorkd > 0)
                        {
                            decimal Emplr_ESI = 0;
                            decimal Emply_EDLI = 0;
                            decimal Emply_ADMIN = 0;
                            decimal Emplr_PF = Lib.Convert2Decimal(Dr["SAL_PF_EMPLR"].ToString());//Same as Employee (D01)
                            decimal Emplr_PF_Pension = 0;
                            decimal pf_cel_limit = Lib.Convert2Decimal(Dr["SAL_PF_CEL_LIMIT"].ToString());
                            decimal pf_cel_limit_amt = Lib.Convert2Decimal(Dr["SAL_PF_CEL_LIMIT_AMT"].ToString());
                            decimal pf_emplr_pension_per = Lib.Convert2Decimal(Dr["SAL_PF_EMPLR_PENSION_PER"].ToString());

                            Emplr_PF_Pension = Lib.Convert2Decimal(Dr["PENSION"].ToString());

                            decimal Emplr_PF_Share = Emplr_PF - Emplr_PF_Pension;
                            decimal nTot = 0, AdminPercent = 0, EdliPercent = 0;
                            decimal Sal_PF_Base = 0;
                            Sal_PF_Base = Lib.Convert2Decimal(Dr["PF_Base"].ToString());
                            decimal EPS_Amt = 0;
                            EPS_Amt = Lib.Convert2Decimal(Dr["EPS_AMT"].ToString());

                            if (Lib.Convert2Decimal(Dr["D02"].ToString())>0)
                            {
                                Emplr_ESI = Lib.Convert2Decimal(Dr["sal_gross_earn"].ToString()) * (Lib.Convert2Decimal(Dr["SAL_ESI_EMPLR_PER"].ToString()) / 100);
                                Emplr_ESI = Lib.Conv2Decimal(Lib.NumericFormat(Emplr_ESI.ToString(), 2));
                            }

                            sql = " Update Salarym set SAL_PF_BASE = " + Sal_PF_Base.ToString();
                            sql += " ,SAL_PF_EMPLR = " + Emplr_PF.ToString();
                            sql += " ,SAL_PF_EMPLR_PENSION = " + Emplr_PF_Pension.ToString();
                            sql += " ,SAL_PF_EMPLR_SHARE  = " + Emplr_PF_Share.ToString();
                            sql += " ,SAL_PF_EPS_AMT = " + Lib.Convert2Decimal(Dr["EPS_AMT"].ToString()).ToString();
                            sql += " ,SAL_ESI_EMPLR = " + Emplr_ESI.ToString();
                            AdminPercent = Lib.Convert2Decimal(Dr["SAL_ADMIN_PER"].ToString());
                            nTot = Sal_PF_Base * (AdminPercent / 100); 
                            sql += " ,SAL_ADMIN_EMPLY = " + Lib.NumericFormat(nTot.ToString(), 2);
                            if (Dr["SAL_ADMIN_BASED_ON"].ToString() == "E")
                                sql += " ,SAL_ADMIN_AMT = " + Lib.NumericFormat(nTot.ToString(), 2);

                            EdliPercent = Lib.Convert2Decimal(Dr["SAL_EDLI_PER"].ToString());
                            // nTot = Sal_PF_Base * (EdliPercent / 100);//Changed on 05/11/2019 due emply wise edli based on pf limit not on pf base amt
                            nTot = EPS_Amt * (EdliPercent / 100);
                            sql += " ,SAL_EDLI_EMPLY = " + Lib.NumericFormat(nTot.ToString(), 2);
                            if (Dr["SAL_EDLI_BASED_ON"].ToString() == "E")
                                sql += " ,SAL_EDLI_AMT = " + Lib.NumericFormat(nTot.ToString(), 2);

                            sql += " ,SAL_ESI_BASE = " + Lib.Convert2Decimal(Dr["SAL_ESI_BASE"].ToString()).ToString();

                            sql += " where REC_COMPANY_CODE ='" + CompCode + "'";
                            sql += " and SAL_EMP_ID='" + EMP_ID + "'";
                            sql += " and SAL_MONTH=" + Month;
                            sql += " and SAL_YEAR=" + Year;

                            Con_Oracle.BeginTransaction();
                            Con_Oracle.ExecuteNonQuery(sql);
                            Con_Oracle.CommitTransaction();
                        }
                    }


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

      
        private void CalculateLop(string EMP_ID, int Year, int Month, string CompCode)
        {
            DataTable Dt_LevMonth;
            try
            {
                if (Month > 0 && Year > 0)
                {
                    sql = "   select lev_lp,lev_days_worked from leavem ";
                    sql += "   where rec_company_code ='" + CompCode + "'";
                    sql += "   and lev_emp_id ='" + EMP_ID + "'";
                    sql += "   and lev_month =" + Month;
                    sql += "   and lev_year =" + Year;
                    sql += "   and nvl(lev_lp,0) > 0 ";
                    Dt_LevMonth = new DataTable();
                    Dt_LevMonth = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_LevMonth.Rows.Count > 0)
                    {
                        Lib.FindLoPAmount(EMP_ID, Year, Month, Lib.Conv2Decimal(Dt_LevMonth.Rows[0]["lev_lp"].ToString()), Lib.Conv2Decimal(Dt_LevMonth.Rows[0]["lev_days_worked"].ToString())); 
                    }
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

        private bool IsESIGovShare(decimal grossearnings, decimal ESIAmt, int mm, int yyyy)
        {
            bool bOk = false;
            if (ESIAmt > 0)
            {
                int DaysInMonth = DateTime.DaysInMonth(yyyy, mm);
                decimal PerDaySalary = grossearnings / DaysInMonth;
                if (PerDaySalary < 140)
                    bOk = true;
            }
            return bOk;
        }


        public Dictionary<string, object> PrintSalarySheet(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            try
            {
                string filename = "";
                string filetype = "";
                string filedisplayname = "";

                string filename2 = "";
                string filetype2 = "";
                string filedisplayname2 = "";


                string type = "";
                if (SearchData.ContainsKey("type"))
                    type = SearchData["type"].ToString();

                string pkid = "";
                if (SearchData.ContainsKey("pkid"))
                    pkid = SearchData["pkid"].ToString();

                string report_folder = "";
                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();

                string folderid = "";
                if (SearchData.ContainsKey("folderid"))
                    folderid = SearchData["folderid"].ToString();

                string branch_code = "";
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();

                string company_code = "";
                if (SearchData.ContainsKey("company_code"))
                    company_code = SearchData["company_code"].ToString();

                string year_code = "";
                if (SearchData.ContainsKey("year_code"))
                    year_code = SearchData["year_code"].ToString();

                int salmonth = 15;
                if (SearchData.ContainsKey("salmonth"))
                    salmonth = Lib.Conv2Integer(SearchData["salmonth"].ToString());

                int salyear = 1900;
                if (SearchData.ContainsKey("salyear"))
                    salyear = Lib.Conv2Integer(SearchData["salyear"].ToString());

                string empstatus = "";
                if (SearchData.ContainsKey("empstatus"))
                    empstatus = SearchData["empstatus"].ToString();

                string isadmin = "N";
                if (SearchData.ContainsKey("isadmin"))
                    isadmin = SearchData["isadmin"].ToString();

                string psadmin = "N";
                if (SearchData.ContainsKey("psadmin"))
                    psadmin = SearchData["psadmin"].ToString();

                string ssadmin = "N";
                if (SearchData.ContainsKey("ssadmin"))
                    ssadmin = SearchData["ssadmin"].ToString();

                string fileformat = "PDF";
                if (SearchData.ContainsKey("filetype"))
                    fileformat = SearchData["filetype"].ToString();

                int empbrgroup = 1;
                if (SearchData.ContainsKey("empbrgroup"))
                    empbrgroup = Lib.Conv2Integer(SearchData["empbrgroup"].ToString());

                if (empbrgroup <= 0)
                    empbrgroup = 1;

                if (type == "SALSHEET")
                {
                    SalarySheetService SsRpt = new SalarySheetService();

                    SsRpt.PKID = pkid;
                    SsRpt.folderid = folderid;
                    SsRpt.report_folder = report_folder;
                    SsRpt.company_code = company_code;
                    SsRpt.branch_code = branch_code;
                    SsRpt.year_code = year_code;
                    SsRpt.Emp_Status = empstatus;
                    SsRpt.cYear = salyear;
                    SsRpt.cMonth = salmonth;
                    SsRpt.IsAdmin = ssadmin;
                    SsRpt.IsConsol = false;
                    SsRpt.fileformat = fileformat;
                    SsRpt.emp_br_grp = empbrgroup;
                    SsRpt.ProcessData();
                    filename = SsRpt.File_Name;
                    filetype = SsRpt.File_Type;
                    filedisplayname = SsRpt.File_Display_Name;

                    if (branch_code == "TUTSF" && pkid == "")
                    {
                        folderid = Guid.NewGuid().ToString().ToUpper();
                        SalarySheetService SsRpt2 = new SalarySheetService();
                        SsRpt2.PKID = pkid;
                        SsRpt2.folderid = folderid;
                        SsRpt2.report_folder = report_folder;
                        SsRpt2.company_code = company_code;
                        SsRpt2.branch_code = branch_code;
                        SsRpt2.year_code = year_code;
                        SsRpt2.Emp_Status = empstatus;
                        SsRpt2.cYear = salyear;
                        SsRpt2.cMonth = salmonth;
                        SsRpt2.IsAdmin = ssadmin;
                        SsRpt.IsConsol = false;
                        SsRpt2.fileformat = fileformat;
                        SsRpt2.emp_br_grp = 2;
                        SsRpt2.ProcessData();
                        filename2 = SsRpt2.File_Name;
                        filetype2 = SsRpt2.File_Type;
                        filedisplayname2 = SsRpt2.File_Display_Name;
                    }
                }
                else
                {

                    PaySlipService PsRpt = new PaySlipService();

                    PsRpt.PKID = pkid;
                    PsRpt.folderid = folderid;
                    PsRpt.report_folder = report_folder;
                    PsRpt.company_code = company_code;
                    PsRpt.branch_code = branch_code;
                    PsRpt.year_code = year_code;
                    PsRpt.Emp_Status = empstatus;
                    PsRpt.cYear = salyear;
                    PsRpt.cMonth = salmonth;
                    PsRpt.IsAdmin = psadmin;
                    PsRpt.ProcessData();

                    filename = PsRpt.File_Name;
                    filetype = PsRpt.File_Type;
                    filedisplayname = PsRpt.File_Display_Name;
                }

                RetData.Add("filename", filename);
                RetData.Add("filetype", filetype);
                RetData.Add("filedisplayname", filedisplayname);
                RetData.Add("filename2", filename2);
                RetData.Add("filetype2", filetype2);
                RetData.Add("filedisplayname2", filedisplayname2);
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
            return RetData;
        }


        private void SaveLeave(DataRow dr, int levmonth, int levyear, int levlp, int levDaysWorked, string sFinyear, string suser)
        {
            DateTime dtime = DateTime.Now;

            sql = "select lev_pkid from leavem ";
            sql += "   where rec_company_code ='" + dr["REC_COMPANY_CODE"].ToString() + "'";
            sql += "   and lev_emp_id ='" + dr["SAL_EMP_ID"].ToString() + "'";
            sql += "   and lev_month =" + levmonth.ToString();
            sql += "   and lev_year =" + levyear.ToString();
            if (!Con_Oracle.IsRowExists(sql))
            {
                dtime = new DateTime(levyear, levmonth, 1);

                sql = " Insert into leavem";
                sql += "  (LEV_PKID, LEV_EMP_ID, LEV_YEAR, LEV_MONTH, ";
                sql += "  LEV_LP, LEV_DAYS_WORKED, LEV_EDIT_CODE, ";
                sql += "  LEV_FIN_YEAR,LEV_DATE, REC_CATEGORY, ";
                sql += "  REC_COMPANY_CODE, REC_BRANCH_CODE, ";
                sql += "  REC_CREATED_BY, REC_CREATED_DATE )";
                sql += "  Values";
                sql += "    (";
                sql += "  [LEV_PKID], [LEV_EMP_ID], [LEV_YEAR], [LEV_MONTH], ";
                sql += "  [LEV_LP], [LEV_DAYS_WORKED], [LEV_EDIT_CODE], ";
                sql += "  [LEV_FIN_YEAR],[LEV_DATE], [REC_CATEGORY], ";
                sql += "  [REC_COMPANY_CODE], [REC_BRANCH_CODE], ";
                sql += "  [REC_CREATED_BY], [REC_CREATED_DATE])";

                sql = sql.Replace("[LEV_PKID]", "'" + Guid.NewGuid().ToString().ToUpper() + "'");
                sql = sql.Replace("[LEV_EMP_ID]", "'" + dr["SAL_EMP_ID"].ToString() + "'");
                sql = sql.Replace("[LEV_YEAR]", levyear.ToString());
                sql = sql.Replace("[LEV_MONTH]", levmonth.ToString());
                sql = sql.Replace("[LEV_LP]", levlp.ToString());
                sql = sql.Replace("[LEV_DAYS_WORKED]", levDaysWorked.ToString());
                sql = sql.Replace("[LEV_EDIT_CODE]", "'{S}'");
                sql = sql.Replace("[LEV_FIN_YEAR]", sFinyear);
                sql = sql.Replace("[REC_CATEGORY]", "'" + dr["emp_job_status"].ToString() + "'");
                sql = sql.Replace("[REC_COMPANY_CODE]", "'" + dr["REC_COMPANY_CODE"].ToString() + "'");
                sql = sql.Replace("[REC_BRANCH_CODE]", "'" + dr["REC_BRANCH_CODE"].ToString() + "'");
                sql = sql.Replace("[REC_CREATED_BY]", "'" + suser + "'");
                sql = sql.Replace("[REC_CREATED_DATE]", "sysdate");
                sql = sql.Replace("[LEV_DATE]", "'" +  dtime.ToString(Lib.BACK_END_DATE_FORMAT) + "'");

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
            }
        }
        private void SaveSalaryH(string company_code,string branch_code,string saldate, int salmonth, int salyear,string sFinyear)
        {
            
            sql = "select salh_pkid from salaryh  ";
            sql += "   where rec_company_code ='" + company_code + "'";
            sql += "   and rec_branch_code ='" + branch_code + "'";
            sql += "   and salh_month =" + salmonth.ToString();
            sql += "   and salh_year =" + salyear.ToString();
            if (!Con_Oracle.IsRowExists(sql))
            {
                sql = " Insert into salaryh ";
                sql += "  (SALH_PKID, SALH_DATE, SALH_YEAR, SALH_MONTH, ";
                sql += "  SALH_FIN_YEAR, SALH_POSTED, ";
                sql += "  REC_COMPANY_CODE, REC_BRANCH_CODE) ";
                sql += "  Values";
                sql += "    (";
                sql += "  [SALH_PKID], [SALH_DATE], [SALH_YEAR], [SALH_MONTH], ";
                sql += "  [SALH_FIN_YEAR], [SALH_POSTED],";
                sql += "  [REC_COMPANY_CODE], [REC_BRANCH_CODE]) ";
             
                sql = sql.Replace("[SALH_PKID]", "'" + Guid.NewGuid().ToString().ToUpper() + "'");
                if (saldate == "NULL")
                    sql = sql.Replace("[SALH_DATE]", "NULL");
                else
                    sql = sql.Replace("[SALH_DATE]", "'" + saldate + "'");
                sql = sql.Replace("[SALH_YEAR]", salyear.ToString());
                sql = sql.Replace("[SALH_MONTH]", salmonth.ToString());
                sql = sql.Replace("[SALH_FIN_YEAR]", sFinyear);
                sql = sql.Replace("[SALH_POSTED]", "'N'");
                sql = sql.Replace("[REC_COMPANY_CODE]", "'" + company_code + "'");
                sql = sql.Replace("[REC_BRANCH_CODE]", "'" + branch_code+ "'");
               
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
            }
        }
    }
}

  