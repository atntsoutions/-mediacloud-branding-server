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
    public class SalaryMasterService : BL_Base
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
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                dRow = getListColumns();

                sWhere = " where a.rec_company_code = '" + company_code + "'";
                sWhere += " and a.rec_branch_code = '" + branch_code + "'";
                sWhere += " and a.emp_in_payroll = 'Y' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.emp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  a.emp_no like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " ) ";
                }
               
                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM empm  a ";
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
                sql += " select sal_pkid, emp_pkid,emp_no,emp_name,sal_gross_earn,sal_gross_deduct,sal_net ";
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
                sql += " ,row_number() over (order by emp_no) rn ";
                sql += " from empm a ";
                sql += " left join salarym sm on (a.emp_pkid=sm.sal_emp_id and sm.sal_month=0 and sm.sal_year=0) ";
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
                    mRow.sal_emp_id = Dr["emp_pkid"].ToString();
                    mRow.sal_emp_code = Dr["emp_no"].ToString();
                    mRow.sal_emp_name = Dr["emp_name"].ToString();
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
            RetData.Add("record", dRow);

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
            string id = SearchData["empid"].ToString();
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
                sql += " ,sal_pay_date ,sal_work_days,sal_mail_sent,c.param_name as sal_emp_grade,sal_edit_code ";
                sql += " from salarym a  ";
                sql += " inner join empm b on a.sal_emp_id = b.emp_pkid";
                sql += " left join param c on b.emp_grade_id = c.param_pkid ";
                sql += " where a.sal_emp_id ='" + id + "'";
                sql += " and a.sal_month = 0 ";
                sql += " and a.sal_year = 0 " ;

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
                    mRow.sal_emp_grade = Dr["sal_emp_grade"].ToString();
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
                    mRow.sal_edit_code = Dr["sal_edit_code"].ToString();
                    break;
                }

                if (smode == "ADD")
                {
                    sql = "select emp_pkid,emp_no,emp_name from empm where emp_pkid ='" + id + "'";
                    Con_Oracle = new DBConnection();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    string eid = "", ecode = "", ename = "";
                    if (Dt_Rec.Rows.Count > 0)
                    {
                        eid = Dt_Rec.Rows[0]["emp_pkid"].ToString();
                        ecode = Dt_Rec.Rows[0]["emp_no"].ToString();
                        ename = Dt_Rec.Rows[0]["emp_name"].ToString();
                    }
                    mRow = NewRecord(eid, ecode, ename);
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

        private Salarym NewRecord(string _Emp_id,string _Emp_code, string _Emp_name)
        {
            Salarym Rec = new Salarym();
            Rec.sal_pkid = Guid.NewGuid().ToString().ToUpper();
            Rec.sal_emp_id = _Emp_id;
            Rec.sal_emp_code = _Emp_code;
            Rec.sal_emp_name = _Emp_name;
            Rec.sal_date = "";
            Rec.sal_month = 0;
            Rec.sal_year = 0;
            Rec.sal_fin_year = 0;
            Rec.sal_days_worked = 0;
            Rec.a01 = 0;
            Rec.a02 = 0;
            Rec.a03 = 0;
            Rec.a04 = 0;
            Rec.a05 = 0;
            Rec.a06 = 0;
            Rec.a07 = 0;
            Rec.a08 = 0;
            Rec.a09 = 0;
            Rec.a10 = 0;
            Rec.a11 = 0;
            Rec.a12 = 0;
            Rec.a13 = 0;
            Rec.a14 = 0;
            Rec.a15 = 0;
            Rec.a16 = 0;
            Rec.a17 = 0;
            Rec.a18 = 0;
            Rec.a19 = 0;
            Rec.a20 = 0;
            Rec.a21 = 0;
            Rec.a22 = 0;
            Rec.a23 = 0;
            Rec.a24 = 0;
            Rec.a25 = 0;
            Rec.d01 = 0;
            Rec.d02 = 0;
            Rec.d03 = 0;
            Rec.d04 = 0;
            Rec.d05 = 0;
            Rec.d06 = 0;
            Rec.d07 = 0;
            Rec.d08 = 0;
            Rec.d09 = 0;
            Rec.d10 = 0;
            Rec.d11 = 0;
            Rec.d12 = 0;
            Rec.d13 = 0;
            Rec.d14 = 0;
            Rec.d15 = 0;
            Rec.d16 = 0;
            Rec.d17 = 0;
            Rec.d18 = 0;
            Rec.d19 = 0;
            Rec.d20 = 0;
            Rec.d21 = 0;
            Rec.d22 = 0;
            Rec.d23 = 0;
            Rec.d24 = 0;
            Rec.d25 = 0;
            Rec.sal_lop_amt = 0;
            Rec.sal_gross_earn = 0;
            Rec.sal_gross_deduct = 0;
            Rec.sal_net = 0;
            Rec.sal_basic_rt = 0;
            Rec.sal_da_rt = 0;
            Rec.sal_pf_mon_year = "";
            Rec.sal_pf_limit = 0;
            Rec.sal_pf_cel_limit = 0;
            Rec.sal_pf_cel_limit_amt = 0;
            Rec.sal_pf_bal = 0;
            Rec.sal_pf_wage_bal = 0;
            Rec.sal_pf_base = 0;
            Rec.sal_pf_emplr = 0;
            Rec.sal_pf_emplr_share = 0;
            Rec.sal_pf_emplr_pension = 0;
            Rec.sal_pf_emplr_pension_per = 0;
            Rec.sal_pf_eps_amt = 0;
            Rec.sal_admin_per = 0;
            Rec.sal_admin_amt = 0;
            Rec.sal_admin_based_on = "";
            Rec.sal_edli_per = 0;
            Rec.sal_edli_amt = 0;
            Rec.sal_edli_based_on = "";
            Rec.sal_is_esi = false;
            Rec.sal_esi_base = 0;
            Rec.sal_esi_emplr_per = 0;
            Rec.sal_esi_limit = 0;
            Rec.sal_esi_gov_share = 0;
            Rec.sal_pay_date = "";
            Rec.sal_work_days = 0;
            Rec.sal_mail_sent = false;
            Rec.sal_esi_emply_per = 0;
            Rec.sal_pf_per = 0;
            Rec.sal_edit_code = "{S}";
            Rec.rec_mode = "ADD";

            return Rec;
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
                    Lib.AddError(ref str, " Details Closed, Can't Edit ");
                
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

                sql = "select sal_pkid from salarym a ";
                sql += " where a.sal_emp_id = '" + Record.sal_emp_id + "'";
                sql += " and a.sal_month = 0 ";
                sql += " and a.sal_year = 0 ";
                DataTable Dt_temp = Con_Oracle.ExecuteQuery(sql);
                Record.rec_mode = "ADD";
                Record.sal_pkid = Guid.NewGuid().ToString().ToUpper();
                if (Dt_temp.Rows.Count > 0)
                {
                    Record.sal_pkid = Dt_temp.Rows[0]["sal_pkid"].ToString();
                    Record.rec_mode = "EDIT";
                }

                //if (ErrorMessage != "")
                //    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                Record.sal_basic_rt = Record.a01 + Record.a11 + Record.a20;
                Record.sal_da_rt = Record.a02;

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("salarym", Record.rec_mode, "sal_pkid", Record.sal_pkid);
                Rec.InsertNumeric("sal_days_worked", Record.sal_days_worked.ToString());
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
                Rec.InsertNumeric("sal_basic_rt", Record.sal_basic_rt.ToString());
                Rec.InsertNumeric("sal_da_rt", Record.sal_da_rt.ToString());
                Rec.InsertString("sal_pf_mon_year", Record.sal_pf_mon_year);
                Rec.InsertNumeric("sal_pf_limit", Record.sal_pf_limit.ToString());
                Rec.InsertNumeric("sal_pf_per", Record.sal_pf_per.ToString());
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
                Rec.InsertNumeric("sal_esi_emply_per", Record.sal_esi_emply_per.ToString());
                Rec.InsertNumeric("sal_esi_emplr_per", Record.sal_esi_emplr_per.ToString());
                Rec.InsertNumeric("sal_esi_limit", Record.sal_esi_limit.ToString());
                Rec.InsertNumeric("sal_esi_gov_share", Record.sal_esi_gov_share.ToString());
                Rec.InsertDate("sal_pay_date", Record.sal_pay_date);
                Rec.InsertNumeric("sal_work_days", Record.sal_work_days.ToString());
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("sal_mail_sent", "N");
                    Rec.InsertDate("sal_date", DateTime.Now);
                    Rec.InsertString("sal_emp_id", Record.sal_emp_id);
                    Rec.InsertNumeric("sal_month", Record.sal_month.ToString());
                    Rec.InsertNumeric("sal_year", Record.sal_year.ToString());
                    Rec.InsertNumeric("sal_fin_year", Record._globalvariables.year_code);
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

    }
}
