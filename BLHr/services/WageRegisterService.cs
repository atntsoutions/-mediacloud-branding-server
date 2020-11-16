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
    public class WageRegisterService : BL_Base
    {
        ExcelFile file;
        ExcelWorksheet ws = null;
        ExcelWorksheet ws2 = null;
        CellRange myCell;
        List<Salarym> mList = null;
        Salarym mRow;
        SalDet dRow;
        DataTable Dt_EMP = new DataTable();
        string type = "";
        string searchstring = "";
        string branch_code = "";
        string company_code = "";
        string branch_region = "";
        string empstatus = "";
        string reporttype = "";
        string report_folder = "";
        string folderid = "";
        string File_Name = "";
        string File_Type = "";
        string File_Display_Name = "myreport.pdf";
        int salmonth = 0;
        int salyear = 0;
        int iCol = 0;
        int iRow = 0;
        string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "", comp_location="";
        bool IsConsol = false;
        private int CopyFromRow = 0;
        private int CopyToRow = 0;
        private int CopyDelhiTotRow = 0;
        Dictionary<string, decimal> myDict = null;
        private string ImagePath = "";
        string sWhere = "";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            mList = new List<Salarym>();
            
            type = SearchData["type"].ToString();
            searchstring = SearchData["searchstring"].ToString().ToUpper();
            branch_code = SearchData["branch_code"].ToString();
            company_code = SearchData["company_code"].ToString();
            empstatus = SearchData["empstatus"].ToString();
            reporttype = SearchData["reporttype"].ToString();
            salyear = Lib.Conv2Integer(SearchData["salyear"].ToString());
            salmonth = 0;
            if (SearchData.ContainsKey("salmonth"))
                salmonth = Lib.Conv2Integer(SearchData["salmonth"].ToString());
            branch_region = "";
            if (SearchData.ContainsKey("branch_region"))
                branch_region = SearchData["branch_region"].ToString();
            report_folder = "";
            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();
            folderid = "";
            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();
            try
            {
                IsConsol = false;
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

               

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select sal_pkid,sal_month,sal_date,sal_emp_id,sal_gross_earn,sal_gross_deduct,sal_net,sal_lop_amt ";
                sql += " ,upper(trim(to_char(sal_date, 'MONTH')))||'-'||to_char(sal_date, 'YYYY') as sal_mon_yr";
                sql += " ,sal_days_worked";
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
                sql += " ,emp_no,emp_name,emp_do_joining,b.emp_father_name,emp_bank_acno, c.param_name as emp_grade,a.rec_printed,sal_mail_sent, d.param_name as emp_designation ";
                sql += " from salarym a ";
                sql += " inner join empm b on a.sal_emp_id = b.emp_pkid ";
                sql += " left join param c on b.emp_grade_id = c.param_pkid ";
                sql += " left join param d on b.emp_designation_id = d.param_pkid "; 
               sql += sWhere;
                sql += " order by emp_no";

                

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Salarym();
                    mRow.sal_pkid = Dr["sal_pkid"].ToString();
                    mRow.sal_emp_id = Dr["sal_emp_id"].ToString();
                    mRow.sal_emp_code = Dr["emp_no"].ToString();
                    mRow.sal_emp_name = Dr["emp_name"].ToString();
                    mRow.sal_emp_father_name = Dr["emp_father_name"].ToString();
                    mRow.sal_emp_grade = Dr["emp_grade"].ToString();
                    mRow.sal_emp_bank_acno = Dr["emp_bank_acno"].ToString();
                    mRow.sal_emp_designation = Dr["emp_designation"].ToString();
                    mRow.sal_month = Lib.Conv2Integer(Dr["sal_month"].ToString());
                    mRow.sal_date = Lib.DatetoStringDisplayformat(Dr["sal_date"]);
                    mRow.sal_emp_do_joining = Lib.DatetoStringDisplayformat(Dr["emp_do_joining"]);
                    mRow.sal_pf_mon_year = Dr["sal_mon_yr"].ToString();
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
                    mRow.sal_days_worked = Lib.Conv2Decimal(Dr["sal_days_worked"].ToString());
                    mList.Add(mRow);
                }
                if (type == "EXCEL")
                {
                    PrintWage();
                }

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
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

        public string AllValid(Salarym Record)
        {
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {
                
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

        private void ReadCompanyDetails()
        {
            comp_name = ""; comp_add1 = ""; comp_add2 = ""; comp_add3 = ""; comp_location = "";
            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "ADDRESS");
            mSearchData.Add("branch_code", branch_code);
            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    comp_name = Dr["COMP_NAME"].ToString();
                    comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    comp_location = Dr["COMP_LOCATION"].ToString(); 
                    //comp_tel = Dr["COMP_TEL"].ToString();
                    //comp_fax = Dr["COMP_FAX"].ToString();
                    //comp_web = Dr["COMP_WEB"].ToString();
                    //comp_email = Dr["COMP_EMAIL"].ToString();
                    //comp_cinno = Dr["COMP_CINNO"].ToString();
                    //comp_gstin = Dr["COMP_GSTIN"].ToString();
                    break;
                }
            }
        }


        private void PrintWage()
        {
            string fname = "myreport";
            fname = "WAGE-" + branch_code + "-" + new DateTime(salyear, salmonth, 01).ToString("MMMM").ToUpper() + "-" + salyear.ToString();
            if (fname.Length > 30)
                fname = fname.Substring(0, 30);
            File_Display_Name = Lib.ProperFileName(fname) + ".xls";
            File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
            File_Type = "xls";
            // ImagePath = report_folder + "\\Images";

            OpenFile();
            SetColumns();
            WriteHeading();
            FillData();
            file.SaveXls(File_Name);
        }

        private void OpenFile()
        {
            file = new ExcelFile();
            file.Worksheets.Add("Report");
            ws = file.Worksheets["Report"];
            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;
        }

        private void SetColumns()
        {
            iRow = 0;
            iCol = 2;
            ws.Columns[0].Width = 255 * 3;
            ws.Columns[1].Width = 255 * 25;
            ws.Columns[2].Width = 255 * 25;
            ws.Columns[3].Width = 255 * 16;
            ws.Columns[4].Width = 255 * 7;
            ws.Columns[5].Width = 255 * 6;
            ws.Columns[6].Width = 255 * 8;
            ws.Columns[7].Width = 255 * 6;
            ws.Columns[8].Width = 255 * 9;
            ws.Columns[9].Width = 255 * 4;
            ws.Columns[10].Width = 255 * 8;
            ws.Columns[11].Width = 255 * 5;
            ws.Columns[12].Width = 255 * 5; //4
            ws.Columns[13].Width = 255 * 8;
            ws.Columns[14].Width = 255 * 8;
            ws.Columns[15].Width = 255 * 8;
            ws.Columns[16].Width = 255 * 7;
            ws.Columns[17].Width = 255 * 8;
            ws.Columns[18].Width = 255 * 8;
            ws.Columns[19].Width = 255 * 12;
            for (int s = 0; s < 20; s++)
            {
                if (s > 0 && s < 4)
                {
                    ws.Columns[s].Style.Font.Name = "Arial";
                    ws.Columns[s].Style.Font.Size = 6 * 20;
                }
                else
                {
                    ws.Columns[s].Style.Font.Name = "Arial";
                    ws.Columns[s].Style.Font.Size = 7 * 20;
                }
            }
        }
        private void WriteHeading()
        {
            string str = "";
            ReadCompanyDetails();
            ws.Rows[iRow].Style.Font.Size = 10 * 20;

            //str = "";
            //if (Dt_WAGE != null)
            //    if (Dt_WAGE.Rows.Count > 0)
            //        str = Dt_WAGE.Rows[0]["REC_BRANCH_CODE"].ToString();
            WriteData(iRow, 6, Lib.GetFormNumber(branch_code, "WAGES"), true);
            WriteData(iRow, 10, "Name of Establishment");
            WriteData(iRow, 14, comp_name);

            iRow++;
            iRow++;
            myCell = ws.Cells.GetSubrangeRelative(iRow, iCol, 18, 1);
            myCell.Merged = true;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            myCell.Style.Font.Weight = ExcelFont.BoldWeight;
            myCell.Value = "REGISTER OF WAGES";
            myCell.Style.Font.Name = "Arial";
            myCell.Style.Font.Size = 18 * 20;
            ws.Rows[iRow].Height = 16 * 20;
            iRow++;
            iRow++;

            str = "Wages period from   ";
            if (salmonth > 0 && salyear > 0)
            {
                str += "01/" + salmonth.ToString().PadLeft(2, '0') + "/" + salyear.ToString() + "  to  " + DateTime.DaysInMonth(salyear, salmonth).ToString() + "/" + salmonth.ToString().PadLeft(2, '0') + "/" + salyear.ToString();
            }
            ws.Rows[iRow].Style.Font.Size = 10 * 20;
            WriteData(iRow, 0, str);
            WriteData(iRow, 6, "See Rule 29 (1)", true);
            WriteData(iRow, 10, "Place");
            WriteData(iRow, 14, comp_location);
            iRow++;

            ws.Rows[iRow].Style.Font.Size = 8 * 20;

            Merge_Cell(iRow, 0, "Sl.No", true, 1, 4);
            Merge_Cell(iRow, 1, "Name of Employee", true, 1, 4);
            Merge_Cell(iRow, 2, "Father's/ Hundband's Name", true, 1, 4);
            Merge_Cell(iRow, 3, "Designation", true, 1, 4);
            Merge_Cell(iRow, 4, "Minimum Rate of Wages Payable", true, 2, 2);
            //Merge_Cell(iRow, 5, "DA", true, 1, 2);
            Merge_Cell(iRow, 6, "Rates Of Wages Actually Paid", true, 2, 2);
            // Merge_Cell(iRow, 7, "DA", true, 1, 2);
            Merge_Cell(iRow, 8, "Total Attendance of work done", true, 1, 4);// 
            Merge_Cell(iRow, 9, "OT", true, 1, 4);
            Merge_Cell(iRow, 10, "Gross Wages Payable", true, 1, 4);
            Merge_Cell(iRow, 11, "DEDUCTIONS", true, 5, 2);
            //Merge_Cell(iRow, 12, "ESI", true, 1, 4);
            //Merge_Cell(iRow, 13, "Loan/ Advance", true, 1, 4);
            //Merge_Cell(iRow, 14, "Other Deductions", true, 1, 4);
            //Merge_Cell(iRow, 15, "Total Deductions", true, 1, 4);
            Merge_Cell(iRow, 16, "Wages Paid", true, 1, 4);
            Merge_Cell(iRow, 17, "Date of Payment", true, 1, 4);
            Merge_Cell(iRow, 18, "Signature of Employee", true, 1, 4);

            iRow++;
            iRow++;

            Merge_Cell(iRow, 4, "Basic", true, 1, 2);
            Merge_Cell(iRow, 5, "DA", true, 1, 2);
            Merge_Cell(iRow, 6, "Basic", true, 1, 2);
            Merge_Cell(iRow, 7, "DA", true, 1, 2);

            Merge_Cell(iRow, 11, "EPF", true, 1, 2);
            Merge_Cell(iRow, 12, "ESI", true, 1, 2);
            Merge_Cell(iRow, 13, "Loan/ Advance", true, 1, 2);
            Merge_Cell(iRow, 14, "Other Deductn", true, 1, 2);
            Merge_Cell(iRow, 15, "Total Deductn", true, 1, 2);

            iRow++;
        }
        private void FillData()
        {
            string str = "";
            int SlNo = 0;
            decimal DaysWork = 0;
            decimal tDeduct = 0;
            decimal cAmt = 0;
            foreach (Salarym Rec in mList)
            {
                iRow++;
                SlNo++;
                tDeduct = 0;
                DaysWork = Lib.Convert2Decimal(Rec.sal_days_worked.ToString());
                ws.Rows[iRow].Height = 16 * 20;
                WriteData(iRow, 0, SlNo, false, "ALL");
                WriteData(iRow, 1, Rec.sal_emp_name, false, "ALL");
                WriteData(iRow, 2, Rec.sal_emp_father_name, false, "ALL");
                WriteData(iRow, 3, Rec.sal_emp_designation, false, "ALL");
                WriteData(iRow, 4, Rec.sal_basic_rt, false, "ALL");
                WriteData(iRow, 5, Rec.sal_da_rt, false, "ALL");
                WriteData(iRow, 6, Rec.a01, false, "ALL");
                WriteData(iRow, 7, Rec.a02, false, "ALL");
                WriteData(iRow, 8, DaysWork, false, "ALL");
                WriteData(iRow, 9, 0, false, "ALL");
                WriteData(iRow, 10, Rec.sal_gross_earn, false, "ALL");

                tDeduct += Lib.Convert2Decimal(Rec.d01.ToString());
                WriteData(iRow, 11, Rec.d01, false, "ALL");

                tDeduct += Lib.Convert2Decimal(Rec.d02.ToString());
                WriteData(iRow, 12, Rec.d02, false, "ALL");

                cAmt = Lib.Convert2Decimal(Rec.d06.ToString()) + Lib.Convert2Decimal(Rec.d07.ToString());
                tDeduct += cAmt;
                WriteData(iRow, 13, cAmt, false, "ALL");//Loan and Advance dedct

                tDeduct = Lib.Convert2Decimal(Rec.sal_gross_deduct.ToString()) - tDeduct;
                WriteData(iRow, 14, tDeduct, false, "ALL");
                WriteData(iRow, 15, Rec.sal_gross_deduct, false, "ALL");
                WriteData(iRow, 16, Rec.sal_net, false, "ALL");
                str = "";
                //if (!Dr["SAL_PAY_DATE"].Equals(DBNull.Value))
                //{
                //    THISDATE = (DateTime)Dr["SAL_PAY_DATE"];
                //    str = THISDATE.ToString("dd/MM/yyyy").ToUpper();
                //}
                WriteData(iRow, 17, str, false, "ALL");
                ws.Cells[iRow, 17].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                WriteData(iRow, 18, null, false, "ALL");
            }

        }

        private void WriteData(int _Row, int _Col, Object sData)
        {
            WriteData(_Row, _Col, sData, false, System.Drawing.Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, string BORDERS)
        {
            WriteData(_Row, _Col, sData, false, System.Drawing.Color.Black, BORDERS);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold)
        {
            WriteData(_Row, _Col, sData, bBold, System.Drawing.Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, string BORDERS)
        {
            WriteData(_Row, _Col, sData, bBold, System.Drawing.Color.Black, BORDERS);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, System.Drawing.Color c, string BORDERS)
        {
            ws.Cells[_Row, _Col].Value = sData;
            if (bBold)
                ws.Cells[_Row, _Col].Style.Font.Weight = ExcelFont.BoldWeight;
            ws.Cells[_Row, _Col].Style.Font.Color = c;
            if (BORDERS == "ALL")
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
            else if (BORDERS == "NFORMAT")
                ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
            else if (BORDERS == "R_FORMAT")
            {
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, System.Drawing.Color.Black, LineStyle.Thin);
            }
            else if (BORDERS == "L_FORMAT")
            {
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Left, System.Drawing.Color.Black, LineStyle.Thin);
            }
        }
        private void Merge_Cell(int _Row, int _Col, int _Width, int _Height)
        {
            myCell = ws.Cells.GetSubrangeRelative(_Row, _Col, _Width, _Height);
            myCell.Merged = true;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = "Arial";
            myCell.Style.Font.Size = 8 * 20;
        }
        private void Merge_Cell(int _Row, int _Col, object sData, bool fBold, int _Width, int _Height, string FontName = "Arial", string cBorders = "ALL")
        {
            myCell = ws.Cells.GetSubrangeRelative(_Row, _Col, _Width, _Height);
            myCell.Merged = true;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = FontName;
            myCell.Style.Font.Size = 9 * 20;
            if (fBold)
                myCell.Style.Font.Weight = ExcelFont.BoldWeight;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            if (cBorders == "ALL")
                myCell.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
            myCell.Value = sData;
        }
    }
}
