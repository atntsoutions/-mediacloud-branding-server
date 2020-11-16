using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;


using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts

{
    public class TdsPayService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_Sal = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<TdsPay> mList = new List<TdsPay>();
        TdsPay mrow;
        int iRow = 0;
        int iCol = 0;
        string type = "";

        string report_folder = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string searchtype = "";
        string searchstring = "";
        string rec_category = "";
        string searchexpid = "";
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string code = "";
        string id = "";
        string acc_name = "";

        string ErrorMessage = "";
        Boolean final = false;
        Boolean allbranch = false;
        decimal tot_amount = 0;
        decimal tot_tds = 0;
        decimal tot_salary = 0;
        decimal tot_tds_rate = 0;
        decimal tot_intrest = 0;
        decimal tot_commision = 0;
        decimal tot_contract = 0;
        decimal tot_rent = 0;
        decimal tot_building = 0;
        decimal tot_fornpay = 0;
        decimal tot_ptax = 0;


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<TdsPay>();
            ErrorMessage = "";
            try
            {
                type = SearchData["type"].ToString();

                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();

                //  int count = from_date.TakeWhile(char.IsWhiteSpace).Count();
                //   type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);
                final = (Boolean)SearchData["final"];
                allbranch = (Boolean)SearchData["allbranch"];

                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");
                //if (type == "SCREEN" && from_date != "NULL" && to_date != "NULL")
                //{
                //    DateTime dt_frm = DateTime.Parse(from_date);
                //    DateTime dt_to = DateTime.Parse(to_date);
                //    int days = (dt_to - dt_frm).Days;

                //    if (days > 31)
                //        Lib.AddError(ref ErrorMessage, " | Only one month data range can be used,use excel to download");
                //}

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                sql = "";

                sql = "  select c.rec_branch_code,acc_code,jvh_date,jvh_type,jvh_vrno,";
                sql += " param_code as panno, param_name as Name,param_id1 as location,";
                sql += " jv_tds_gross_amt as Amount, jv_tds_rate as tdsper,";
                sql += " case when acc_code in('194A') then jv_credit else 0 end as interest, ";
                sql += " case when acc_code in('194H') then jv_credit else 0 end as comm,";
                sql += " case when acc_code in('194C') then jv_credit else 0 end as contract,";
                sql += " case when acc_code in('194I') then jv_credit else 0 end as rent, ";
                sql += " case when acc_code in('194IA') then jv_credit else 0 end as building,";
                sql += " case when acc_code in('192B') then jv_credit else 0 end as salary, ";
                sql += " case when acc_code in('194J') then jv_credit else 0 end as ptax, ";
                sql += " case when acc_code in('195') then jv_credit else 0 end as fornpay, ";
                sql += " jv_credit as tds ";
                sql += " from acctm ";
                sql += " inner join ledgert b on acc_pkid = jv_acc_id";
                sql += " inner join ledgerh c on jv_parent_id =  jvh_pkid";
                sql += " left join param pan on jv_pan_id = pan.param_pkid";


                if (final)
                {
                    sql += " inner join (  select xref_cr_jv_id from ledgert c inner join ledgerxref d on c.jv_pkid = d.xref_dr_jv_id   ";
                    sql += " where   c.jv_debit >0   ";
                    sql += " ) xref on b.jv_pkid = xref.xref_cr_jv_id";
                }

                sql += " where  jvh_type <> 'OP'  and jv_credit >0  ";
                sql += " and jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";
                sql += " and acc_code in ('194A','194C','194H','194I','194IA', '192B','194J', '195') ";

                if (!allbranch)
                {
                    sql += " and c.rec_branch_code = '{BRCODE}'";
                }
                sql += " order by jvh_date";

                sql = sql.Replace("{BRCODE}", branch_code);
                // sql = sql.Replace("{YEAR}", year_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);






                tot_amount = 0;
                tot_salary = 0;
                tot_tds = 0;
                tot_tds_rate = 0;
                tot_intrest = 0;
                tot_commision = 0;
                tot_contract = 0;
                tot_rent = 0;
                tot_building = 0;
                tot_ptax = 0;
                tot_fornpay = 0;

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new TdsPay();
                    mrow.rowtype = "DETAIL";
                    mrow.rowcolor = "BLACK";
                    mrow.rec_branch_code = Dr["rec_branch_code"].ToString();
                    mrow.acc_code = Dr["acc_code"].ToString();
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.jvh_type = Dr["jvh_type"].ToString();
                    mrow.jvh_vrno = Dr["jvh_vrno"].ToString();
                    mrow.panno = Dr["panno"].ToString();
                    mrow.party_name = Dr["Name"].ToString();
                    mrow.location = Dr["location"].ToString();
                    mrow.jv_tds_gross_amt = Lib.Conv2Decimal(Dr["Amount"].ToString());
                    mrow.jv_tds_rate = Lib.Conv2Decimal(Dr["tdsper"].ToString());
                    mrow.interest = Lib.Conv2Decimal(Dr["interest"].ToString());
                    mrow.commision = Lib.Conv2Decimal(Dr["comm"].ToString());
                    mrow.contract = Lib.Conv2Decimal(Dr["contract"].ToString());
                    mrow.rent = Lib.Conv2Decimal(Dr["rent"].ToString());
                    mrow.building = Lib.Conv2Decimal(Dr["building"].ToString());
                    mrow.salary = Lib.Conv2Decimal(Dr["salary"].ToString());
                    mrow.forgnpay = Lib.Conv2Decimal(Dr["fornpay"].ToString());

                    mrow.ptax = Lib.Conv2Decimal(Dr["ptax"].ToString());
                    mrow.jv_credit = Lib.Conv2Decimal(Dr["tds"].ToString());

                    mList.Add(mrow);

                    tot_amount += Lib.Conv2Decimal(mrow.jv_tds_gross_amt.ToString());
                    tot_salary += Lib.Conv2Decimal(mrow.salary.ToString());
                    tot_tds += Lib.Conv2Decimal(mrow.jv_credit.ToString());
                    tot_tds_rate += Lib.Conv2Decimal(mrow.jv_tds_rate.ToString());
                    tot_intrest += Lib.Conv2Decimal(mrow.interest.ToString());
                    tot_commision += Lib.Conv2Decimal(mrow.commision.ToString());
                    tot_contract += Lib.Conv2Decimal(mrow.contract.ToString());
                    tot_rent += Lib.Conv2Decimal(mrow.rent.ToString());
                    tot_building += Lib.Conv2Decimal(mrow.building.ToString());
                    tot_ptax += Lib.Conv2Decimal(mrow.ptax.ToString());
                    tot_fornpay += Lib.Conv2Decimal(mrow.forgnpay.ToString());
                }
                if (mList.Count > 1)
                {
                    mrow = new TdsPay();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.acc_code = "TOTAL";
                    mrow.rec_branch_code = "";

                    mrow.jv_tds_gross_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_amount.ToString(), 2));
                    mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(tot_tds.ToString(), 2));
                    mrow.salary = Lib.Conv2Decimal(Lib.NumericFormat(tot_salary.ToString(), 2));
                    mrow.jv_tds_rate = Lib.Conv2Decimal(Lib.NumericFormat(tot_tds_rate.ToString(), 2));
                    mrow.interest = Lib.Conv2Decimal(Lib.NumericFormat(tot_intrest.ToString(), 2));
                    mrow.commision = Lib.Conv2Decimal(Lib.NumericFormat(tot_commision.ToString(), 2));
                    mrow.contract = Lib.Conv2Decimal(Lib.NumericFormat(tot_contract.ToString(), 2));
                    mrow.rent = Lib.Conv2Decimal(Lib.NumericFormat(tot_rent.ToString(), 2));
                    mrow.building = Lib.Conv2Decimal(Lib.NumericFormat(tot_building.ToString(), 2));
                    mrow.ptax = Lib.Conv2Decimal(Lib.NumericFormat(tot_ptax.ToString(), 2));
                    mrow.forgnpay = Lib.Conv2Decimal(Lib.NumericFormat(tot_fornpay.ToString(), 2));
                    mList.Add(mrow);
                }


                sql = " select 'A' as dataorder,'{BR}' as branch, a.rec_branch_code, sal_date,emp_pan,emp_no,emp_name,emp_local_city, d03 as saltds,  ";
                sql += " 0 as ptax,sal_gross_earn from salarym a inner join empm b on sal_emp_id = b.emp_pkid ";
                sql += " where d03 > 0  and sal_date  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')  ";
                if (!allbranch)
                {
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                }
                sql += " order by a.rec_branch_code,sal_date ";

                sql = sql.Replace("{BRCODE}", branch_code);
                // sql = sql.Replace("{YEAR}", year_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);



                Dt_Sal = new DataTable();
                Dt_Sal = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                tot_amount = 0;
                tot_salary = 0;
                tot_tds = 0;
                tot_tds_rate = 0;
                tot_intrest = 0;
                tot_commision = 0;
                tot_contract = 0;
                tot_rent = 0;
                tot_building = 0;
                tot_ptax = 0;
                tot_fornpay = 0;


                if (Dt_Sal.Rows.Count > 0)
                {
                    mrow = new TdsPay();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.acc_code = "SALARY";
                    mrow.rec_branch_code = "";


                    mrow.jv_tds_gross_amt = 0;
                    mrow.jv_credit = 0;
                    mrow.salary = 0;
                    mrow.jv_tds_rate = 0;
                    mrow.interest = 0;
                    mrow.commision = 0;
                    mrow.contract = 0;
                    mrow.rent = 0;
                    mrow.building = 0;
                    mrow.ptax = 0;
                    mrow.forgnpay = 0;


                    mList.Add(mrow);
                }




                foreach (DataRow Dr in Dt_Sal.Rows)
                {
                    mrow = new TdsPay();
                    mrow.rowtype = "DETAIL";
                    mrow.rowcolor = "BLACK";


                    mrow.jvh_type = "";
                    mrow.jvh_vrno = "";
                    mrow.panno = "";
                    mrow.party_name = "";
                    mrow.location = "";


                    mrow.jv_tds_gross_amt = 0;
                    mrow.jv_credit = 0;
                    mrow.salary = 0;
                    mrow.jv_tds_rate = 0;
                    mrow.interest = 0;
                    mrow.commision = 0;
                    mrow.contract = 0;
                    mrow.rent = 0;
                    mrow.building = 0;
                    mrow.ptax = 0;
                    mrow.forgnpay = 0;



                    mrow.rec_branch_code = Dr["rec_branch_code"].ToString();
                    mrow.acc_code = Dr["emp_no"].ToString();
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["sal_date"]);
                    mrow.panno = Dr["emp_pan"].ToString();
                    mrow.location = Dr["emp_local_city"].ToString();
                    mrow.party_name = Dr["emp_name"].ToString();

                    mrow.jv_tds_gross_amt = Lib.Conv2Decimal(Dr["sal_gross_earn"].ToString());
                    mrow.jv_credit = Lib.Conv2Decimal(Dr["saltds"].ToString());
                    mrow.salary = Lib.Conv2Decimal(Dr["saltds"].ToString());
                    mrow.ptax = Lib.Conv2Decimal(Dr["ptax"].ToString());

                    
                    tot_amount += Lib.Conv2Decimal(mrow.jv_tds_gross_amt.ToString());
                    tot_tds += Lib.Conv2Decimal(mrow.jv_credit.ToString());
                    tot_salary += Lib.Conv2Decimal(mrow.salary.ToString());
                    tot_ptax += Lib.Conv2Decimal(mrow.ptax.ToString());


                    mList.Add(mrow);
                }

                if (Dt_Sal.Rows.Count > 1)
                {
                    mrow = new TdsPay();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.acc_code = "TOTAL";
                    mrow.rec_branch_code = "";

                    mrow.jv_tds_gross_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_amount.ToString(), 2));
                    mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(tot_tds.ToString(), 2));
                    mrow.salary = Lib.Conv2Decimal(Lib.NumericFormat(tot_salary.ToString(), 2));
                    mrow.jv_tds_rate = Lib.Conv2Decimal(Lib.NumericFormat(tot_tds_rate.ToString(), 2));
                    mrow.interest = Lib.Conv2Decimal(Lib.NumericFormat(tot_intrest.ToString(), 2));
                    mrow.commision = Lib.Conv2Decimal(Lib.NumericFormat(tot_commision.ToString(), 2));
                    mrow.contract = Lib.Conv2Decimal(Lib.NumericFormat(tot_contract.ToString(), 2));
                    mrow.rent = Lib.Conv2Decimal(Lib.NumericFormat(tot_rent.ToString(), 2));
                    mrow.building = Lib.Conv2Decimal(Lib.NumericFormat(tot_building.ToString(), 2));
                    mrow.ptax = Lib.Conv2Decimal(Lib.NumericFormat(tot_ptax.ToString(), 2));
                    mrow.forgnpay = Lib.Conv2Decimal(Lib.NumericFormat(tot_fornpay.ToString(), 2));
                    mList.Add(mrow);
                }





                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintTdsPayReport();
                }
                Dt_List.Rows.Clear();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", type);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);
            return RetData;
        }

        private void PrintTdsPayReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

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

                File_Display_Name = "TdsPayReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 10;
                WS.Columns[2].Width = 256 * 10;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 6;
                WS.Columns[5].Width = 256 * 6;
                WS.Columns[6].Width = 256 * 14;
                WS.Columns[7].Width = 256 * 28;
                WS.Columns[8].Width = 256 * 10;
                WS.Columns[9].Width = 256 * 13;
                WS.Columns[10].Width = 256 * 14;
                WS.Columns[11].Width = 256 * 13;
                WS.Columns[12].Width = 256 * 14;
                WS.Columns[13].Width = 256 * 14;
                WS.Columns[14].Width = 256 * 14;
                WS.Columns[15].Width = 256 * 14;
                WS.Columns[16].Width = 256 * 14;
                WS.Columns[17].Width = 256 * 14;
                WS.Columns[18].Width = 256 * 14;
                WS.Columns[19].Width = 256 * 14;

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

                if (final)
                {
                    str = "(FINAL)";
                }
                else
                {
                    str = "(TRIAL)";
                }


                Lib.WriteData(WS, iRow, 1, "TDS PAYABLE REPORT " + str, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (allbranch)
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ACC-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAID/CREDITED DATE DEDUCTION DATE", _Color, true, "BT", "L", "", _Size, true, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "AMOUNT PAID/CREDITED", _Color, true, "BT", "R", "", _Size, true, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEDUCTION RATE", _Color, true, "BT", "R", "", _Size, true, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEDUCTED AND DEPOSTIED TAX", _Color, true, "BT", "R", "", _Size, true, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INTREST-194A", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMM.BR-194H", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTRACT-194C", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUILDING-194IA", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RENT-194I", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALARY-192B", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PROF.TAX-194J", _Color, true, "BT", "R", "", _Size, false, 3 * 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FORGN.PAY-195", _Color, true, "BT", "R", "", _Size, false, 3 * 325, "", true);


                foreach (TdsPay Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        if (allbranch)
                            Lib.WriteData(WS, iRow, iCol++, Rec.rec_branch_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.acc_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_vrno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.panno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.location, _Color, false, "", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_tds_gross_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_tds_rate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.interest, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.commision, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.contract, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.building, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rent, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.salary, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ptax, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.forgnpay, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);



                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        if (allbranch)
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.acc_code, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);

                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_tds_gross_amt, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_tds_rate, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.interest, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.commision, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.contract, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.building, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.rent, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.salary, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.ptax, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.forgnpay, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                    }

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
