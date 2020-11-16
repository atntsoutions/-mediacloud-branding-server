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
    public class PAndLService : BL_Base
    {
        DataTable Dt_List = new DataTable();

        ExcelFile WB;
        ExcelWorksheet WS = null;

        int iRow = 0;
        int iCol = 0;

        string type = "";
        string subtype = "";
        string report_folder = "";
        string File_Name = "";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string from_date = "";
        string to_date = "";
        Boolean ismaincode = false;
        Boolean ismonthwise = false;
        string hide_ho_entries = "N";

        List<LedgerReport> mList = new List<LedgerReport>();

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();

            LedgerReport mrow;

            object nAmt = 0;


            type = SearchData["type"].ToString();
            subtype = SearchData["subtype"].ToString();
            report_folder = SearchData["report_folder"].ToString();
            PKID = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            from_date = SearchData["from_date"].ToString();
            to_date = SearchData["to_date"].ToString();
            ismaincode = (Boolean)SearchData["ismaincode"];
            ismonthwise = (Boolean)SearchData["ismonthwise"];
            hide_ho_entries = SearchData["hide_ho_entries"].ToString();


            from_date = Lib.StringToDate(from_date);
            to_date = Lib.StringToDate(to_date);

            decimal nBal = 0;
            decimal nDr = 0;
            decimal nCr = 0;

            decimal nTotDr = 0;
            decimal nTotCr = 0;
            decimal nTotDrBal = 0;
            decimal nTotCrBal = 0;

            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            //File_Name = System.IO.Path.Combine(report_folder, PKID);

            string File_Display_Name = "pandl.xls";
            File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

            try
            {
                if (ismonthwise == false)
                {
                    if (ismaincode)
                    {
                        sql = "";
                        sql += "  select null as rowtype, ";
                        sql += "  acgrp_order, level1, level2, acgrp_name, acc_main_id as acc_pkid, acc_main_code as acc_code, acc_main_name as acc_name, ";
                        sql += "  sum(jv_debit) as debit, ";
                        sql += "  sum(jv_credit) as credit, ";
                        sql += "  sum(jv_debit) - sum(jv_credit) as bal, ";
                        sql += "  0 as drbal, 0 as crbal";
                        sql += "  from ledgerh a ";
                        sql += "  inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += "  inner join acctm on jv_acc_id = acc_pkid ";
                        sql += "  inner join view_acgroupm on acc_group_id = acgrp_pkid ";
                        sql += "  where jvh_year = {YEAR} and jvh_date >= '{FDATE}' and jvh_date <= '{EDATE}' and jvh_posted ='Y' and jvh_type not in('OB','OC','OI')  ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in('HO','IN-ES') ";
                        sql += "  and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += "  and a.rec_deleted = 'N'  and level1  = 'P & L' ";
                        sql += "  group by acgrp_order, level1, level2, acgrp_name, acc_main_id, acc_main_code, acc_main_name";
                        sql += "  order by acgrp_order, acc_main_code ";
                    }
                    else
                    {
                        sql = "";
                        sql += "  select null as rowtype,";
                        sql += "  acgrp_order, level1, level2, acgrp_name, acc_pkid,acc_code, acc_name, ";
                        sql += "  sum(jv_debit) as debit, ";
                        sql += "  sum(jv_credit) as credit, ";
                        sql += "  sum(jv_debit) - sum(jv_credit) as bal, ";
                        sql += "  0 as drbal, 0 as crbal";
                        sql += "  from ledgerh a ";
                        sql += "  inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += "  inner join acctm on jv_acc_id = acc_pkid ";
                        sql += "  inner join view_acgroupm on acc_group_id = acgrp_pkid ";
                        sql += "  where jvh_year = {YEAR} and jvh_date >= '{FDATE}' and jvh_date <= '{EDATE}' and jvh_posted ='Y' and jvh_type not in('OB','OC','OI')  ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in ('HO','IN-ES') ";
                        sql += "  and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += "  and a.rec_deleted = 'N'  and level1  = 'P & L' ";
                        sql += "  group by acgrp_order, level1, level2, acgrp_name,  acc_pkid, acc_code, acc_name";
                        sql += "  order by acgrp_order, acc_code ";
                    }
                }

                // Month Wise
                if (ismonthwise == true)
                {
                    if (ismaincode)
                    {
                        sql = "";
                        sql += " select acgrp_order,level1,level2,acgrp_name,acc_pkid,acc_code, acc_name, apr,may,jun,jul,aug,sep,oct,nov,dec,jan,feb,mar, total from (";
                        sql += " select null as rowtype, acgrp_order, level1, level2, acgrp_name, acc_main_id as acc_pkid, acc_main_code as acc_code, acc_main_name as acc_name, ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'APR' then jv_debit - jv_credit else 0 end) as APR,   ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'MAY' then jv_debit - jv_credit else 0 end) as MAY,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'JUN' then jv_debit - jv_credit else 0 end) as JUN,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'JUL' then jv_debit - jv_credit else 0 end) as JUL,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'AUG' then jv_debit - jv_credit else 0 end) as AUG,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'SEP' then jv_debit - jv_credit else 0 end) as SEP,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'OCT' then jv_debit - jv_credit else 0 end) as OCT,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'NOV' then jv_debit - jv_credit else 0 end) as NOV,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'DEC' then jv_debit - jv_credit else 0 end) as DEC,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'JAN' then jv_debit - jv_credit else 0 end) as JAN,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'FEB' then jv_debit - jv_credit else 0 end) as FEB,  ";
                        sql += " sum(case when to_char(jvh_date,'MON') = 'MAR' then jv_debit - jv_credit else 0 end) as MAR,";
                        sql += " sum(jv_debit) - sum(jv_credit) as total ";
                        sql += " from ledgerh a ";
                        sql += " inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " inner join acctm on jv_acc_id = acc_pkid ";
                        sql += " inner join view_acgroupm on acc_group_id = acgrp_pkid ";
                        sql += " where jvh_year = {YEAR} and jvh_posted ='Y' and jvh_type not in('OB','OC','OI')  ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in ('HO','IN-ES') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N'  and level1  = 'P & L' ";
                        sql += " group by acgrp_order, level1, level2, acgrp_name, acc_main_id, acc_main_code, acc_main_name";
                        sql += ") a ";
                        sql += "where  apr <>0 or may <>0 or jun <>0 or jul <>0  or  aug <>0 or sep <>0 or oct <>0 or nov <>0  or  dec <>0 or jan <>0 or feb <>0 or mar <>0    ";
                        sql += "order by acgrp_order,acc_name ";
                    }
                    else
                    {
                        sql = "";
                        sql += " select acgrp_order,level1,level2,acgrp_name,acc_pkid,acc_code, acc_name, apr,may,jun,jul,aug,sep,oct,nov,dec,jan,feb,mar, total from (";
                        sql += "  select null as rowtype, acgrp_order, level1, level2, acgrp_name, acc_pkid,acc_code, acc_name,";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'APR' then jv_debit - jv_credit else 0 end) as APR,   ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'MAY' then jv_debit - jv_credit else 0 end) as MAY,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'JUN' then jv_debit - jv_credit else 0 end) as JUN,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'JUL' then jv_debit - jv_credit else 0 end) as JUL,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'AUG' then jv_debit - jv_credit else 0 end) as AUG,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'SEP' then jv_debit - jv_credit else 0 end) as SEP,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'OCT' then jv_debit - jv_credit else 0 end) as OCT,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'NOV' then jv_debit - jv_credit else 0 end) as NOV,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'DEC' then jv_debit - jv_credit else 0 end) as DEC,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'JAN' then jv_debit - jv_credit else 0 end) as JAN,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'FEB' then jv_debit - jv_credit else 0 end) as FEB,  ";
                        sql += "  sum(case when to_char(jvh_date,'MON') = 'MAR' then jv_debit - jv_credit else 0 end) as MAR,";
                        sql += "  sum(jv_debit) - sum(jv_credit) as total ";
                        sql += "  from ledgerh a ";
                        sql += "  inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += "  inner join acctm on jv_acc_id = acc_pkid ";
                        sql += "  inner join view_acgroupm on acc_group_id = acgrp_pkid ";
                        sql += "  where jvh_year = {YEAR} and jvh_posted ='Y' and jvh_type not in('OB','OC','OI')  ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in ('HO','IN-ES') ";

                        sql += "  and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += "  and a.rec_deleted = 'N'  and level1  = 'P & L' ";
                        sql += "  group by acgrp_order, level1, level2, acgrp_name, acc_pkid, acc_code, acc_name";
                        sql += " ) a ";
                        sql += " where  apr <>0 or may <>0 or jun <>0 or jul <>0  or  aug <>0 or sep <>0 or oct <>0 or nov <>0  or  dec <>0 or jan <>0 or feb <>0 or mar <>0    ";
                        sql += " order by acgrp_order,acc_name ";
                    }
                }

                sql = sql.Replace("{COMPANY}", company_code);
                sql = sql.Replace("{BRANCH}", branch_code);
                sql = sql.Replace("{YEAR}", year_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                if (ismonthwise == false)
                {
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new LedgerReport();
                        mrow.rowtype = "DETAIL";
                        mrow.level1 = Dr["level1"].ToString();
                        mrow.level2 = Dr["level2"].ToString();
                        mrow.grp_name = Dr["acgrp_name"].ToString();

                        mrow.acc_pkid = Dr["acc_pkid"].ToString();
                        mrow.acc_code = Dr["acc_code"].ToString();
                        mrow.acc_name = Dr["acc_name"].ToString();

                        nDr = Lib.Conv2Decimal(Dr["debit"].ToString());
                        nCr = Lib.Conv2Decimal(Dr["credit"].ToString());

                        nTotDr += nDr;
                        nTotCr += nCr;

                        nBal = Lib.Conv2Decimal(Dr["bal"].ToString());
                        if (nBal > 0)
                        {
                            Dr["drbal"] = nBal;
                            Dr["crbal"] = 0;
                            nTotDrBal += nBal;
                        }
                        else if (nBal < 0)
                        {
                            Dr["drbal"] = 0;
                            Dr["crbal"] = Math.Abs(nBal);
                            nTotCrBal += Math.Abs(nBal);
                        }

                        mrow.debit = Lib.Conv2Decimal(Dr["debit"].ToString());
                        mrow.credit = Lib.Conv2Decimal(Dr["credit"].ToString());
                        mrow.drbal = Lib.Conv2Decimal(Dr["drbal"].ToString());
                        mrow.crbal = Lib.Conv2Decimal(Dr["crbal"].ToString());

                        mList.Add(mrow);
                    }

                    nBal = nTotDrBal - nTotCrBal;

                    mrow = new LedgerReport();
                    mrow.grp_name = "";
                    mrow.acc_code = "";
                    mrow.rowtype = "TOTAL";
                    mrow.debit = 0;
                    mrow.credit = 0;

                    if (nBal > 0)
                    {
                        mrow.acc_name = "LOSS";
                        mrow.drbal = 0;
                        mrow.crbal = nBal;
                        nTotCrBal += nBal;
                    }
                    if (nBal < 0)
                    {
                        mrow.acc_name = "PROFIT";
                        mrow.drbal = Math.Abs(nBal);
                        mrow.crbal = 0;
                        nTotDrBal += Math.Abs(nBal);
                    }
                    mList.Add(mrow);

                    mrow = new LedgerReport();
                    mrow.rowtype = "TOTAL";
                    mrow.grp_name = "";
                    mrow.acc_code = "";
                    mrow.acc_name = "TOTAL";
                    mrow.debit = nTotDr;
                    mrow.credit = nTotCr;
                    mrow.drbal = nTotDrBal;
                    mrow.crbal = nTotCrBal;
                    mList.Add(mrow);
                    Dt_List.Rows.Clear();
                }

                if (ismonthwise == true)
                {
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new LedgerReport();
                        mrow.rowtype = "DETAIL";
                        mrow.level1 = Dr["level1"].ToString();
                        mrow.level2 = Dr["level2"].ToString();
                        mrow.grp_name = Dr["acgrp_name"].ToString();

                        mrow.acc_pkid = Dr["acc_pkid"].ToString();
                        mrow.acc_code = Dr["acc_code"].ToString();
                        mrow.acc_name = Dr["acc_name"].ToString();


                        mrow.apr = Lib.Conv2Decimal(Dr["apr"].ToString());
                        mrow.may = Lib.Conv2Decimal(Dr["may"].ToString());
                        mrow.jun = Lib.Conv2Decimal(Dr["jun"].ToString());
                        mrow.jul = Lib.Conv2Decimal(Dr["jul"].ToString());
                        mrow.aug = Lib.Conv2Decimal(Dr["aug"].ToString());
                        mrow.sep = Lib.Conv2Decimal(Dr["sep"].ToString());
                        mrow.oct = Lib.Conv2Decimal(Dr["oct"].ToString());
                        mrow.nov = Lib.Conv2Decimal(Dr["nov"].ToString());
                        mrow.dec = Lib.Conv2Decimal(Dr["dec"].ToString());
                        mrow.jan = Lib.Conv2Decimal(Dr["jan"].ToString());
                        mrow.feb = Lib.Conv2Decimal(Dr["feb"].ToString());
                        mrow.mar = Lib.Conv2Decimal(Dr["mar"].ToString());

                        mrow.bal = Lib.Conv2Decimal(Dr["total"].ToString());

                        mList.Add(mrow);
                    }

                    mrow = new LedgerReport();
                    mrow.grp_name = "TOTAL";
                    mrow.acc_code = "";
                    mrow.acc_name = "";
                    mrow.rowtype = "TOTAL";

                    mrow.apr = Lib.Conv2Decimal(Dt_List.Compute("sum(apr)", "1=1").ToString());
                    mrow.may = Lib.Conv2Decimal(Dt_List.Compute("sum(may)", "1=1").ToString());
                    mrow.jun = Lib.Conv2Decimal(Dt_List.Compute("sum(jun)", "1=1").ToString());
                    mrow.jul = Lib.Conv2Decimal(Dt_List.Compute("sum(jul)", "1=1").ToString());
                    mrow.aug = Lib.Conv2Decimal(Dt_List.Compute("sum(aug)", "1=1").ToString());
                    mrow.sep = Lib.Conv2Decimal(Dt_List.Compute("sum(sep)", "1=1").ToString());
                    mrow.oct = Lib.Conv2Decimal(Dt_List.Compute("sum(oct)", "1=1").ToString());
                    mrow.nov = Lib.Conv2Decimal(Dt_List.Compute("sum(nov)", "1=1").ToString());
                    mrow.dec = Lib.Conv2Decimal(Dt_List.Compute("sum(dec)", "1=1").ToString());
                    mrow.jan = Lib.Conv2Decimal(Dt_List.Compute("sum(jan)", "1=1").ToString());
                    mrow.feb = Lib.Conv2Decimal(Dt_List.Compute("sum(feb)", "1=1").ToString());
                    mrow.mar = Lib.Conv2Decimal(Dt_List.Compute("sum(mar)", "1=1").ToString());
                    mrow.bal = Lib.Conv2Decimal(Dt_List.Compute("sum(total)", "1=1").ToString());
                    mList.Add(mrow);

                    Dt_List.Rows.Clear();
                }

                if (type == "EXCEL")
                {
                    ProcessExcelFile();
                }

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", type);
            RetData.Add("page_current", page_current);
            RetData.Add("page_count", page_count);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", "EXCEL");
            RetData.Add("filedisplayname", File_Display_Name);

            RetData.Add("list", mList);

            return RetData;
        }


        private void ProcessExcelFile()
        {

            try
            {



                string _Border = "";
                Boolean _Bold = false;
                Color _Color = Color.Black;
                int _Size = 0;

                string sTitle = "";

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];
                WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.Portrait = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256;
                WS.Columns[1].Width = 256 * 17;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 35;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;

                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;
                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;


                WS.Columns[5].Style.NumberFormat = "#,0.00";
                WS.Columns[6].Style.NumberFormat = "#,0.00";
                WS.Columns[7].Style.NumberFormat = "#,0.00";
                WS.Columns[8].Style.NumberFormat = "#,0.00";
                WS.Columns[9].Style.NumberFormat = "#,0.00";

                WS.Columns[10].Style.NumberFormat = "#,0.00";
                WS.Columns[11].Style.NumberFormat = "#,0.00";
                WS.Columns[12].Style.NumberFormat = "#,0.00";
                WS.Columns[13].Style.NumberFormat = "#,0.00";
                WS.Columns[14].Style.NumberFormat = "#,0.00";
                WS.Columns[15].Style.NumberFormat = "#,0.00";
                WS.Columns[16].Style.NumberFormat = "#,0.00";
                WS.Columns[17].Style.NumberFormat = "#,0.00";
                WS.Columns[18].Style.NumberFormat = "#,0.00";
                WS.Columns[18].Style.NumberFormat = "#,0.00";
                WS.Columns[20].Style.NumberFormat = "#,0.00";
                WS.Columns[21].Style.NumberFormat = "#,0.00";



                iRow = 1; iCol = 1;

                iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

                if (ismonthwise)
                    sTitle = "PANDL A/C ";
                else
                    sTitle = "PANDL A/C  FROM " + Lib.getFrontEndDate(from_date) + " TO " + Lib.getFrontEndDate(to_date);

                if (ismaincode)
                    sTitle += " (MAIN CODE WISE)";
                else
                    sTitle += " (SUB CODE WISE)";

                Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

                iCol = 1;
                _Color = Color.DarkBlue;
                _Border = "TB";
                _Size = 10;

                Lib.WriteData(WS, iRow, iCol++, "GROUP", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

                if (ismonthwise)
                {
                    Lib.WriteData(WS, iRow, iCol++, "APR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "MAY", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "JUN", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "JUL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "AUG", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "SEP", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "OCT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "NOV", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "DEC", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "JAN", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "FEB", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "MAR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "BAL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                }
                else
                {
                    Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "DR-BAL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "CR-BAL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                }



                foreach (LedgerReport Row in mList)
                {
                    iRow++; iCol = 1;
                    _Border = "";
                    _Bold = false;
                    _Color = Color.Black;

                    if (Row.rowtype == "TOTAL")
                    {
                        _Border = "TB";
                        _Bold = true;
                        _Color = Color.DarkBlue;
                    }

                    Lib.WriteData(WS, iRow, iCol++, Row.grp_name, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.acc_code, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Row.acc_name, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                    if (ismonthwise)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Row.apr, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.may, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.jun, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.jul, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.aug, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.sep, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.oct, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.nov, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.dec, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.jan, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.feb, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.mar, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.bal, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    else
                    {
                        Lib.WriteData(WS, iRow, iCol++, Row.debit, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.credit, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.drbal, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Row.crbal, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                throw Ex;
            }

        }

    }
}

