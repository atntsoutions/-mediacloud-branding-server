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
    public class ArrearsService : BL_Base
    {

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
          
            List<Arrearsm> mList = new List<Arrearsm>();
            Arrearsm mRow;
            ArrDet dRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            //string from_date = SearchData["from_date"].ToString();
            //string to_date = SearchData["to_date"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            string Category = SearchData["category"].ToString();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                //DateTime dt_from = GetThisdate(from_date);
                //DateTime dt_to = GetThisdate(to_date);

                // Lib.StringToDate(from_date)

                dRow = getListColumns();

                sWhere = " where a.rec_company_code = '" + company_code + "'";
                sWhere += " and a.rec_branch_code = '" + branch_code + "'";
                sWhere += " and a.arr_fin_year =" + year_code;
                sWhere += " and nvl(a.rec_category,'MASTER')='" + Category + "' ";
                //// if (!chKAll.Checked)
                //{
                //    //sWhere += " and a.arr_from_date >= to_date('{FDATE}','DD-MON-YYYY') ";
                //    //sWhere += " and a.arr_to_date <= to_date('{EDATE}','DD-MON-YYYY') ";

                //    sWhere += " and to_char(ARR_FROM_DATE ,'MM') =" + dt_from.Month;
                //    sWhere += " and to_char(ARR_FROM_DATE ,'YYYY') =" + dt_from.Year;
                //    sWhere += " and to_char(ARR_TO_DATE ,'MM') =" + dt_to.Month;
                //    sWhere += " and to_char(ARR_TO_DATE ,'YYYY') =" + dt_to.Year;
                //}
                //sWhere = sWhere.Replace("{FDATE}", startrow.ToString());
                //sWhere = sWhere.Replace("{EDATE}", startrow.ToString());

                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(b.emp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += "  or ";
                    sWhere += "  b.emp_no like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " ) ";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  from arrearsm a ";
                    sql += " inner join empm b on a.arr_emp_id = b.emp_pkid ";
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
                //sql += " select sal_pkid, emp_pkid,emp_no,emp_name,sal_gross_earn,sal_gross_deduct,sal_net ";
                //sql += " ,a01,a02,a03,a04,a05";
                //sql += " ,a06,a07,a08,a09,a10";
                //sql += " ,a11,a12,a13,a14,a15";
                //sql += " ,a16,a17,a18,a19,a20";
                //sql += " ,a21,a22,a23,a24,a25";
                //sql += " ,d01,d02,d03,d04,d05";
                //sql += " ,d06,d07,d08,d09,d10";
                //sql += " ,d11,d12,d13,d14,d15";
                //sql += " ,d16,d17,d18,d19,d20";
                //sql += " ,d21,d22,d23,d24,d25";
                //sql += " ,row_number() over (order by emp_no) rn ";
                //sql += " from empm a ";
                //sql += " left join salarym sm on (a.emp_pkid=sm.sal_emp_id and sm.sal_month=0 and sm.sal_year=0) ";
                sql += "  select arr_pkid,arr_month,arr_fin_year,arr_emp_id,arr_from_date,arr_to_date ";//to_char(to_date(arr_month, 'MM'), 'MONTH') as
                sql += "  ,A01,A02,A03,A04,A05";
                sql += "  ,A06,A07,A08,A09,A10";
                sql += "  ,A11,A12,A13,A14,A15";
                sql += "  ,A16,A17,A18,A19,A20";
                sql += "  ,A21,A22,A23,A24,A25";
                sql += "  ,D01,D02,D03,D04,D05";
                sql += "  ,D06,D07,D08,D09,D10";
                sql += "  ,D11,D12,D13,D14,D15";
                sql += "  ,D16,D17,D18,D19,D20";
                sql += "  ,D21,D22,D23,D24,D25";
                sql += "  ,arr_net,arr_gross_earn";
                sql += "  ,arr_gross_deduct";
                sql += "  ,emp_pkid,emp_name,emp_no,c.param_name as emp_grade";
                sql += "  ,emp_do_joining,arr_lop_days,arr_lop_amt,emp_bank_acno ";
                sql += "  ,row_number() over (order by emp_no) rn ";
                sql += "  from arrearsm a";
                sql += "  inner join empm b on a.arr_emp_id = b.emp_pkid";
                sql += "  left join param c on b.emp_grade_id = c.param_pkid";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by emp_no";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Arrearsm();
                    mRow.arr_pkid = Dr["arr_pkid"].ToString();
                    mRow.arr_emp_id = Dr["emp_pkid"].ToString();
                    mRow.arr_emp_code = Dr["emp_no"].ToString();
                    mRow.arr_emp_name = Dr["emp_name"].ToString();
                    mRow.arr_from_date = Lib.DatetoString(Dr["arr_from_date"]);
                    mRow.arr_to_date = Lib.DatetoString(Dr["arr_to_date"]);
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
                    mRow.arr_gross_earn = Lib.Conv2Decimal(Dr["arr_gross_earn"].ToString());
                    mRow.arr_gross_deduct = Lib.Conv2Decimal(Dr["arr_gross_deduct"].ToString());
                    mRow.arr_net = Lib.Conv2Decimal(Dr["arr_net"].ToString());
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
            return RetData;
        }

        private DateTime GetThisdate(string sdate)
        {
            DateTime dt = DateTime.Now;
            int dd = 0, mm = 0, yy = 0;
            string[] ThisDate = null;

            if (sdate.ToString().Contains("/"))
                ThisDate = sdate.ToString().Split('/');
            else if (sdate.ToString().Contains("-"))
                ThisDate = sdate.ToString().Split('-');
            else if (sdate.ToString().Contains("."))
                ThisDate = sdate.ToString().Split('.');
            if (ThisDate != null)
            {
                if (ThisDate.Length == 3)
                {
                    yy = Lib.Conv2Integer(ThisDate[0]);
                    mm = Lib.Conv2Integer(ThisDate[1]);
                    dd = Lib.Conv2Integer(ThisDate[2]);

                }
            }
            if (mm > 0 && dd > 0)
            {
                if (yy < 100)
                    yy = yy + 2000;
                dt = new DateTime(yy, mm, dd);
            }
            else
                dt = new DateTime(1900, 1, 1);

            return dt;
        }
        private ArrDet getListColumns()
        {
            ArrDet drow = new ArrDet();
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


        public Dictionary<string, object> NewRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Arrearsm mRow = new Arrearsm();
            try
            {
                mRow = NewRow("", "", "");
                List<ArrDet> nList = new List<ArrDet>();
                ArrDet nRow;
                for (int i = 0; i < 12; i++)
                {
                    nRow = new ArrDet();
                    nRow.e_caption1 = "";
                    nRow.e_amt1 = 0;
                    nRow.e_visible1 = false;
                    nRow.e_caption2 = "";
                    nRow.e_amt2 = 0;
                    nRow.e_visible2 = false;
                    nRow.d_caption1 = "";
                    nRow.d_amt1 = 0;
                    nRow.d_visible1 = false;
                    nRow.d_caption2 = "";
                    nRow.d_amt2 = 0;
                    nRow.d_visible2 = false;
                    nList.Add(nRow);
                }
                mRow.DetList = GetDetList(nList, mRow);

                //DataTable Dt_Rec = new DataTable();

                //sql += " select sal_pkid,sal_code,sal_desc,sal_head,sal_head_order ";
                //sql += " from salaryheadm a  ";
                //sql += " where  a.sal_pkid ='" + id + "'";

                //Con_Oracle = new DBConnection();
                //Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                //Con_Oracle.CloseConnection();
                //foreach (DataRow Dr in Dt_Rec.Rows)
                //{
                //    mRow = new salaryheadm();
                //    mRow.sal_pkid = Dr["sal_pkid"].ToString();
                //    mRow.sal_code = Dr["sal_code"].ToString();
                //    mRow.sal_desc = Dr["sal_desc"].ToString();
                //    mRow.sal_head = Dr["sal_head"].ToString();
                //    mRow.sal_head_order = Lib.Conv2Integer(Dr["sal_head_order"].ToString());


                //    break;
                //}
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

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Arrearsm mRow = new  Arrearsm();
            string id = SearchData["pkid"].ToString();
            string category = SearchData["category"].ToString();
            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select arr_pkid, arr_month,arr_fin_year,arr_emp_id,arr_from_date,arr_to_date ";
                sql += "  ,A01,A02,A03,A04,A05";
                sql += "  ,A06,A07,A08,A09,A10";
                sql += "  ,A11,A12,A13,A14,A15";
                sql += "  ,A16,A17,A18,A19,A20";
                sql += "  ,A21,A22,A23,A24,A25";
                sql += "  ,D01,D02,D03,D04,D05";
                sql += "  ,D06,D07,D08,D09,D10";
                sql += "  ,D11,D12,D13,D14,D15";
                sql += "  ,D16,D17,D18,D19,D20";
                sql += "  ,D21,D22,D23,D24,D25";
                sql += "  ,arr_net,arr_gross_earn";
                sql += "  ,arr_gross_deduct";
                sql += "  ,emp_pkid,emp_name,emp_no,c.param_name as emp_grade";
                sql += "  ,emp_do_joining,arr_lop_days,arr_lop_amt,emp_bank_acno ";
                sql += "  from arrearsm a";
                sql += "  inner join empm b on a.arr_emp_id = b.emp_pkid";
                sql += "  left join param c on b.emp_grade_id = c.param_pkid";
                sql += "  where a.arr_pkid ='" + id + "'";
                sql += "  and nvl(a.rec_category,'MASTER')='" + category + "' ";
                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Arrearsm();
                    mRow.arr_pkid = Dr["arr_pkid"].ToString();
                    mRow.arr_emp_id = Dr["emp_pkid"].ToString();
                    mRow.arr_emp_code = Dr["emp_no"].ToString();
                    mRow.arr_emp_name = Dr["emp_name"].ToString();
                    mRow.arr_emp_grade = Dr["emp_grade"].ToString();
                    mRow.arr_emp_do_joining = Lib.DatetoString(Dr["emp_do_joining"]);
                    mRow.arr_emp_bank_acno = Dr["emp_bank_acno"].ToString();
                    mRow.arr_from_date = Lib.DatetoString(Dr["arr_from_date"]);
                    mRow.arr_to_date = Lib.DatetoString(Dr["arr_to_date"]);
                    mRow.arr_fin_year = Lib.Conv2Integer(Dr["arr_fin_year"].ToString());
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
                    mRow.arr_lop_amt = Lib.Conv2Decimal(Dr["arr_lop_amt"].ToString());
                    mRow.arr_lop_days = Lib.Conv2Decimal(Dr["arr_lop_days"].ToString());
                    mRow.arr_gross_earn = Lib.Conv2Decimal(Dr["arr_gross_earn"].ToString());
                    mRow.arr_gross_deduct = Lib.Conv2Decimal(Dr["arr_gross_deduct"].ToString());
                    mRow.arr_net = Lib.Conv2Decimal(Dr["arr_net"].ToString());
                    break;
                }

                //if (smode == "ADD")
                //{
                //    sql = "select emp_pkid,emp_no,emp_name from empm where emp_pkid ='" + id + "'";
                //    Con_Oracle = new DBConnection();
                //    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                //    Con_Oracle.CloseConnection();
                //    string eid = "", ecode = "", ename = "";
                //    if (Dt_Rec.Rows.Count > 0)
                //    {
                //        eid = Dt_Rec.Rows[0]["emp_pkid"].ToString();
                //        ecode = Dt_Rec.Rows[0]["emp_no"].ToString();
                //        ename = Dt_Rec.Rows[0]["emp_name"].ToString();
                //    }
                //    mRow = NewRecord(eid, ecode, ename);
                //}

                List<ArrDet> mList = new List<ArrDet>();
                ArrDet dRow;
                for (int i = 0; i < 12; i++)
                {
                    dRow = new ArrDet();
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
            RetData.Add("record", mRow);
            return RetData;
        }

        private List<ArrDet> GetDetList(List<ArrDet> dList, Arrearsm mRow)
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

        public string AllValid(Arrearsm Record)
        { 
            string str = "";
            //DateTime tdate = DateTime.Now;
            //try
            //{
            //    if (Record.sal_code.Trim().Length <= 0)
            //        Lib.AddError(ref str, " | Code Cannot Be Empty");

            //    if (Record.sal_code.Trim().Length > 0)
            //    {

            //        sql = "select sal_pkid from (";
            //        sql += "select sal_pkid  from salaryheadm a where (a.sal_code = '{CODE}')  ";
            //        sql += ") a where sal_pkid <> '{PKID}'";

            //        sql = sql.Replace("{CODE}", Record.sal_code);
            //        sql = sql.Replace("{PKID}", Record.sal_pkid);

            //        if (Con_Oracle.IsRowExists(sql))
            //            Lib.AddError(ref str, " | Code Exists");
            //    }

            //    if (Record.sal_desc.Trim().Length > 0)
            //    {

            //        sql = "select sal_pkid from (";
            //        sql += "select sal_pkid  from salaryheadm a where (a.sal_desc = '{NAME}')  ";
            //        sql += ") a where sal_pkid <> '{PKID}'";

            //        sql = sql.Replace("{NAME}", Record.sal_desc);
            //        sql = sql.Replace("{PKID}", Record.sal_pkid);


            //        if (Con_Oracle.IsRowExists(sql))

            //            Lib.AddError(ref str, " | Description Exists");
            //    }

            //}
            //catch (Exception Ex)
            //{
            //    str = Ex.Message.ToString();
            //}
            return str;
        }
        /*
        
        public Dictionary<string, object> Save(salaryheadm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();


                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);



                DBRecord Rec = new DBRecord();
                Rec.CreateRow("salaryheadm", Record.rec_mode, "sal_pkid", Record.sal_pkid);
                Rec.InsertString("sal_code", Record.sal_code);
                Rec.InsertString("sal_desc", Record.sal_desc);
                Rec.InsertString("sal_head", Record.sal_head);
                Rec.InsertNumeric("sal_head_order", Record.sal_head_order.ToString());
                if (Record.rec_mode == "ADD")
                {

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
        *
        */

        public Dictionary<string, object> Save(Arrearsm Record)
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
                Rec.CreateRow("arrearsm", Record.rec_mode, "arr_pkid", Record.arr_pkid);
              
                Rec.InsertDate("arr_from_date", Record.arr_from_date);
                Rec.InsertDate("arr_to_date", Record.arr_to_date);
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


                // Rec.InsertNumeric("sal_days_worked", Record.sal_days_worked.ToString());
                Rec.InsertNumeric("arr_gross_earn", Record.arr_gross_earn.ToString());
                Rec.InsertNumeric("arr_gross_deduct", Record.arr_gross_deduct.ToString());
                Rec.InsertNumeric("arr_net", Record.arr_net.ToString());
                Rec.InsertNumeric("arr_lop_days", Record.arr_lop_days.ToString());
                // Rec.InsertNumeric("arr_esi_months", Record.arr_lop_days.ToString());
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("arr_month", DateTime.Now.Month.ToString());
                    Rec.InsertNumeric("arr_fin_year", Record._globalvariables.year_code);
                    Rec.InsertString("arr_emp_id", Record.arr_emp_id);
                    Rec.InsertString("arr_edit_code", "{S}");
                    Rec.InsertString("rec_category", "MASTER");
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
        //private string GetDataSql(string Category)
        //{
        //    string sCond = "";
        //   // string sql = "";
        //    string StrFld = "";

        //    sCond = " arr_fin_year =" + yearCode;
        //    sCond += " and nvl(REC_CATEGORY,'MASTER')='" + Category + "' ";
        //    if (!chKAll.Checked)
        //    {
        //        sCond += " and to_char(ARR_FROM_DATE ,'MM') =" + DT_From.Value.Month;
        //        sCond += " and to_char(ARR_FROM_DATE ,'YYYY') =" + DT_From.Value.Year;
        //        sCond += " and to_char(ARR_TO_DATE ,'MM') =" + Dt_To.Value.Month;
        //        sCond += " and to_char(ARR_TO_DATE ,'YYYY') =" + Dt_To.Value.Year;
        //    }
        //    StrFld  = " ARR_PKID,TO_CHAR(TO_DATE(ARR_MONTH, 'MM'), 'MONTH') as ARR_MONTH,ARR_FIN_YEAR,ARR_EMP_ID,ARR_FROM_DATE,ARR_TO_DATE ";
        //    StrFld += " ,A01,A02,A03,A04,A05";
        //    StrFld += " ,A06,A07,A08,A09,A10";
        //    StrFld += " ,A11,A12,A13,A14,A15";
        //    StrFld += " ,A16,A17,A18,A19,A20";
        //    StrFld += " ,A21,A22,A23,A24,A25";
        //    StrFld += " ,D01,D02,D03,D04,D05";
        //    StrFld += " ,D06,D07,D08,D09,D10";
        //    StrFld += " ,D11,D12,D13,D14,D15";
        //    StrFld += " ,ARR_NET,ARR_GROSS_EARN";
        //    StrFld += " ,ARR_GROSS_DEDUCT";
        //    StrFld += " ,EMP_PKID,EMP_NAME,EMP_NO,EMP_GRADE";
        //    StrFld += " ,EMP_DO_JOINING,EMP_COMPANY,ARR_LOP_DAYS,ARR_LOP_AMT,EMP_BANK_ACNO ";
        //    if (IsHo && (CmbBranch.SelectedValue.ToString() == "ALL" || CmbBranch.SelectedValue.ToString() == "REGION"))
        //    {
        //        sql = "";
        //        foreach (DataRow dr in DT_BRANCH.Rows)
        //        {
        //            if (sql != "")
        //                sql += " union all ";
        //            sql += " select " + StrFld + ",'" + dr["BR_NAME"].ToString() + "' as BRANCH ";
        //            sql += " ,case when EMP_GRADE = 'MANAGING DIRECTOR' then 'A' else ";
        //            sql += " case when EMP_GRADE = 'DIRECTOR' then 'B' else 'C' end end as GRADE_ORDER ";
        //            sql += " from " + dr["BR_USER"].ToString() + ".view_arrearsm where " + sCond;
        //            sql += " and emp_company in (" + ChkedCompnies + ") ";
        //        }
        //        sql += " order by grade_order,emp_no ";
        //    }
        //    else
        //    {
        //        sql = " Select " + StrFld + " from " + sUser + "view_arrearsm ";
        //        sql += " where " + sCond;
        //        sql += " and emp_company in (" + ChkedCompnies + ") ";
        //        sql += " order by emp_no ";
        //    }
        //    return sql;
        //}

        private Arrearsm NewRow(string _Emp_id, string _Emp_code, string _Emp_name)
        {
            Arrearsm Rec = new Arrearsm();
            Rec.arr_pkid = Guid.NewGuid().ToString().ToUpper();
            Rec.arr_emp_id = _Emp_id;
            Rec.arr_emp_code = _Emp_code;
            Rec.arr_emp_name = _Emp_name;
            Rec.arr_from_date = "";
            Rec.arr_to_date = "";
            Rec.arr_month = 0;
            Rec.arr_fin_year = 0;
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
            Rec.arr_lop_amt = 0;
            Rec.arr_gross_earn = 0;
            Rec.arr_gross_deduct = 0;
            Rec.arr_net = 0;
            
            Rec.rec_mode = "ADD";
            return Rec;
        }
    }
}
