using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using System.IO;

using XL.XSheet;

namespace BLAccounts
{
    public class CashBookService : BL_Base
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
        string acc_id = "";
        string acc_code = "";
        string acc_name = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string from_date = "";
        string to_date = "";
        Boolean ismaincode = false;

        Boolean showtotaldrcr = false;

        string hide_ho_entries = "N";

        /**All ledger Report print**/
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sFilter = "";

            decimal nAmt = 0;
            decimal nDr = 0;
            decimal nCr = 0;
            decimal nBal = 0;

            decimal DrTot = 0;
            decimal CrTot = 0;


            decimal row_bal = 0;
            decimal tot_row_debit = 0;
            decimal tot_row_credit = 0;
            decimal tot_row_bal = 0;



            int iCtr = 0;

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<LedgerReport> mList = new List<LedgerReport>();


            LedgerReport mrow;

            type = SearchData["type"].ToString();
            subtype = SearchData["subtype"].ToString();
            report_folder = SearchData["report_folder"].ToString();
            PKID = SearchData["pkid"].ToString();
            acc_id = SearchData["acc_id"].ToString();
            acc_code = SearchData["acc_code"].ToString();
            acc_name = SearchData["acc_name"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            from_date = SearchData["from_date"].ToString();
            to_date = SearchData["to_date"].ToString();
            ismaincode = (Boolean)SearchData["ismaincode"];
            showtotaldrcr = (Boolean)SearchData["showtotaldrcr"];
            hide_ho_entries = SearchData["hide_ho_entries"].ToString();


            string maincode = "";

            from_date = Lib.StringToDate(from_date);
            to_date = Lib.StringToDate(to_date);


            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine(report_folder, PKID);

            try
            {

                /*
                if (ismaincode)
                {
                    sql = " select acc_main_code,acc_main_name  from acctm where acc_pkid = '" + acc_id + "'";
                    DataTable Dt_test = new DataTable();
                    Dt_test = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_test.Rows.Count > 0)
                    {
                        maincode = Dt_test.Rows[0]["acc_main_code"].ToString();
                        acc_name = Dt_test.Rows[0]["acc_main_name"].ToString();
                    }
                }
                */

                if (type == "NEW")
                {
                    if (ismaincode)
                    {
                        maincode = acc_code;
                        sql = "";
                        sql += " select 'A' as rowtype, 0 as slno, ";
                        sql += " null as jv_vrno, cast('OP' as nvarchar2(10)) as jv_docno, null as jv_date, null as jv_type,'' as jv_drcr,  ";
                        sql += " nvl(sum(b.jv_debit - b.jv_credit),0) as op, 0 as jv_debit, 0 as jv_credit, 0 as bal, cast('OPENING' as nvarchar2(10)) as jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += " from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                        sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " where a.jvh_year = {YEAR}  ";
                        sql += " and acc_main_code  = '" + acc_code + "' ";
                        if (hide_ho_entries == "Y")
                            sql += "  and a.jvh_type not in('HO','IN-ES') ";
                        sql += " and a.jvh_date < '{SDATE}'  ";
                        sql += " and a.jvh_posted ='Y' and a.jvh_type not in('OB','OC','OI') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";

                        sql += "union all ";

                        sql += "select 'B' as rowtype,  0 as slno,";
                        sql += "a.jvh_vrno as jv_vrno,a.jvh_docno as jv_docno, a.jvh_date as jv_date, a.jvh_type as jv_type, '' as jv_drcr, 0 as op,  sum(b.jv_debit) as jv_debit, sum(b.jv_credit) as jv_credit,  0 as bal, a.jvh_narration as  jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += "from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                        sql += "inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " where a.jvh_year = {YEAR}  ";
                        sql += " and acc_main_code = '" + acc_code + "'";
                        sql += " and a.jvh_posted ='Y' and a.jvh_type not in('OB','OC','OI') ";
                        if (hide_ho_entries == "Y")
                            sql += "  and a.jvh_type not in('HO','IN-ES') ";
                        sql += " and a.jvh_date >= '{SDATE}' and a.jvh_date <= '{EDATE}' ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";
                        sql += " group by a.jvh_vrno,a.jvh_type, a.jvh_docno, a.jvh_date,a.jvh_narration  ";
                        sql += " order by rowtype, jv_date,jv_vrno ";
                    }
                    else
                    {

                        sql = " select * from ( ";

                        
                        sql += " select 'A' as rowtype, 0 as slno, ";
                        sql += " null as jv_vrno, cast('OP' as nvarchar2(10)) as jv_docno, null as jv_date, null as jv_type,'' as jv_drcr,null as acgrp_name, null as  acc_code, null as acc_name, cast(' ' as nvarchar2(40)) as jv_acc_id, ";
                        sql += " nvl(sum( b.jv_debit -  b.jv_credit),0) as op, 0 as jv_debit, 0 as jv_credit, 0 as bal, cast('OPENING' as nvarchar2(10)) as jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += " from ledgerh a inner join ledgert b1 on a.jvh_pkid = b1.jv_parent_id ";
                        sql += " inner join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                        sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " left join acgroupm d on c.acc_group_id = acgrp_pkid ";
                        sql += " where a.jvh_year = {YEAR}  ";
                        sql += " and b1.jv_acc_id = '{ACCID}' ";
                        sql += "  and a.jvh_type ='OP' ";
                        sql += " and a.jvh_date <= '{EDATE}' ";
                        sql += " and a.jvh_posted ='Y' and a.jvh_type not in('OB','OC','OI') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";

                        sql += "union all ";
                        

                        sql += " select 'B' as rowtype, 0 as slno, ";
                        sql += " null as jv_vrno, cast('OP' as nvarchar2(10)) as jv_docno, null as jv_date, null as jv_type,'' as jv_drcr,null as acgrp_name, null as  acc_code, null as acc_name, cast(' ' as nvarchar2(40)) as jv_acc_id, ";
                        sql += " 0 as op, sum(case when b.jv_acc_id <> '{ACCID}' then b.jv_debit else 0 end ) as jv_debit, sum(case when b.jv_acc_id <> '{ACCID}' then b.jv_credit else 0 end ) as jv_credit, 0 as bal, cast('OPENING' as nvarchar2(10)) as jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += " from ledgerh a inner join ledgert b1 on a.jvh_pkid = b1.jv_parent_id ";
                        sql += " inner join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                        sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " left join acgroupm d on c.acc_group_id = acgrp_pkid ";
                        sql += " where a.jvh_year = {YEAR}  ";
                        sql += " and b1.jv_acc_id = '{ACCID}' ";
                        if (hide_ho_entries == "Y")
                            sql += "  and a.jvh_type not in('HO', 'IN-ES' ) ";
                        sql += " and a.jvh_date < '{SDATE}' ";
                        sql += " and a.jvh_posted ='Y' and a.jvh_type not in('OB','OC','OI', 'OP') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";

                        sql += "union all ";


                        sql += "select 'C' as rowtype,  0 as slno,";
                        sql += "a.jvh_vrno as jv_vrno,a.jvh_docno as jv_docno, a.jvh_date as jv_date, a.jvh_type as jv_type, '' as jv_drcr, acgrp_name,acc_code,acc_name, b.jv_acc_id, ";
                        sql += "0 as op,  b.jv_debit, b.jv_credit,  0 as bal, a.jvh_narration as  jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += "from ledgerh a inner join ledgert b1 on a.jvh_pkid = b1.jv_parent_id ";
                        sql += "inner join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                        sql += "inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += "left join acgroupm d on c.acc_group_id = acgrp_pkid ";
                        sql += " where a.jvh_year = {YEAR}  ";
                        sql += " and b1.jv_acc_id = '{ACCID}' ";
                        if (hide_ho_entries == "Y")
                            sql += "  and a.jvh_type not in('HO', 'IN-ES' ) ";
                        sql += " and a.jvh_date >= '{SDATE}' and a.jvh_date <= '{EDATE}' ";
                        sql += " and a.jvh_posted ='Y' and a.jvh_type not in('OB','OC','OI') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";
                        sql += " ) a where jv_acc_id !=  '{ACCID}'";
                        sql += " order by  rowtype, jv_date,jv_vrno ";
                    }

                    sql = sql.Replace("{COMPANY}", company_code);
                    sql = sql.Replace("{BRANCH}", branch_code);
                    sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{ACCID}", acc_id);
                    sql = sql.Replace("{SDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        if (Dr["rowtype"].ToString() !=  "A")
                        {
                            break;
                        }

                       
                        if (Dr["rowtype"].ToString() == "A")
                        {
                            if (Lib.Conv2Decimal(Dr["OP"].ToString()) == 0)
                            {
                                Dt_List.Rows.Remove(Dr);
                                break;
                            }
                        }
                        
                    }

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        iCtr++;
                        if (Dr["rowtype"].ToString() == "A")
                        {
                            nAmt = Lib.Conv2Decimal(Dr["OP"].ToString());
                            if (nAmt > 0)
                            {
                                Dr["jv_debit"] = nAmt;
                                Dr["row_debit"] = nAmt;

                                // this change is for day book
                                Dr["jv_debit"] = 0;
                                Dr["jv_credit"] = nAmt;
                                Dr["row_credit"] = nAmt;

                                nCr = nAmt;
                            }
                            if (nAmt < 0)
                            {
                                nDr = Math.Abs(nAmt);
                                Dr["jv_credit"] = nDr;
                                Dr["row_credit"] = nDr;

                                // this change is for day book
                                Dr["jv_credit"] = 0;
                                Dr["jv_debit"] = nDr;
                                Dr["row_debit"] = nDr;
                            }
                        }
                        else
                        {
                            nDr = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                            nCr = Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                            Dr["row_debit"] = nDr;
                            Dr["row_credit"] = nCr;

                            if (nDr > nCr)
                            {
                                nDr = nDr - nCr;
                                nCr = 0;

                                Dr["jv_debit"] = nDr;
                                Dr["jv_credit"] = nCr;
                            }
                            else if (nCr > nDr)
                            {
                                nCr = nCr - nDr;
                                nDr = 0;

                                Dr["jv_debit"] = nDr;
                                Dr["jv_credit"] = nCr;
                            }
                        }

                        row_bal = Lib.Conv2Decimal(Dr["row_debit"].ToString()) - Lib.Conv2Decimal(Dr["row_credit"].ToString());
                        tot_row_debit  += Lib.Conv2Decimal(Dr["row_debit"].ToString());
                        tot_row_credit += Lib.Conv2Decimal(Dr["row_credit"].ToString());
                        Dr["row_bal"] = row_bal;
                        tot_row_bal += row_bal;

                        DrTot += nDr;
                        CrTot += nCr;

                        nBal += nDr;
                        nBal -= nCr;
                        if (nBal > 0)
                            Dr["jv_drcr"] = "DR";
                        if (nBal < 0)
                            Dr["jv_drcr"] = "CR";

                        Dr["bal"] = Math.Abs(nBal);
                        Dr["slno"] = iCtr;
                    }

                    if (Dt_List.Rows.Count > 0)
                    {
                        iCtr++; // this is for total row
                        page_rowcount = iCtr;
                        page_count = page_rowcount / page_rows;
                        if ((page_rowcount % page_rows) != 0)
                            page_count++;
                        page_current = 1;

                        DataRow Dr = Dt_List.NewRow();
                        Dr["slno"] = page_rowcount;
                        Dr["rowtype"] = "TOTAL";
                        Dr["jv_docno"] = "TOTAL";

                        Dr["jv_debit"] = DrTot;
                        Dr["jv_credit"] = CrTot;
                        Dr["bal"] = Math.Abs(nBal);
                        if (nBal > 0)
                            Dr["jv_drcr"] = "DR";
                        if (nBal < 0)
                            Dr["jv_drcr"] = "CR";


                        Dr["row_debit"] = tot_row_debit;
                        Dr["row_credit"] = tot_row_credit;
                        Dr["row_bal"] = tot_row_bal;

                        Dt_List.Rows.Add(Dr);

                        if (Lib.CreateFolder(report_folder))
                        {
                            Dt_List.TableName = "REPORT";
                            Dt_List.WriteXml(File_Name + ".xml", XmlWriteMode.WriteSchema);
                        }
                    }
                }
                else
                {
                    Dt_List.ReadXml(File_Name + ".xml");
                }

                if (type == "EXCEL")
                {
                    ProcessExcelFile();
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

                    if (Dt_List.Rows.Count > 0)
                    {
                        startrow = (page_current - 1) * page_rows + 1;
                        endrow = (startrow + page_rows) - 1;
                    }

                    sFilter = "slno >= " + startrow.ToString() + " and slno <= " + endrow.ToString();

                    foreach (DataRow Dr in Dt_List.Select(sFilter))
                    {
                        mrow = new LedgerReport();
                        mrow.acc_pkid = acc_id;
                        mrow.jv_docno = Dr["jv_docno"].ToString();

                        mrow.jv_type = Dr["jv_type"].ToString();
                        mrow.jv_vrno = Dr["jv_vrno"].ToString();


                        mrow.grp_name = Dr["acgrp_name"].ToString();
                        mrow.acc_code = Dr["acc_code"].ToString();
                        mrow.acc_name = Dr["acc_name"].ToString();


                        mrow.rec_company_code = company_code;
                        mrow.rec_branch_code = branch_code;
                        mrow.jv_year = year_code;

                        mrow.jv_date = Lib.DatetoStringDisplayformat(Dr["jv_date"]);
                        mrow.debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString(), "NULL");
                        mrow.credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString(), "NULL");
                        mrow.bal = Lib.Conv2Decimal(Dr["bal"].ToString(), "NULL");
                        mrow.jv_drcr = Dr["jv_drcr"].ToString();
                        mrow.jv_narration = Dr["jv_narration"].ToString();

                        mList.Add(mrow);
                    }
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
            RetData.Add("reportfile", File_Name);
            RetData.Add("list", mList);

            return RetData;
        }

        private void ProcessExcelFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 0;

            string sTitle = "";

            string sDate = "";

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];
            WS.ViewOptions.ShowGridLines = false;
            WS.PrintOptions.Portrait = false;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;

            WS.Columns[0].Width = 256;

            WS.Columns[1].Width = 256 * 15; // DATE
            WS.Columns[2].Width = 256 * 15; // TYPE
            WS.Columns[3].Width = 256 * 15; // VRNO

            WS.Columns[4].Width = 256 * 25; // GRP
            WS.Columns[5].Width = 256 * 25; // CODE
            WS.Columns[6].Width = 256 * 25; // NAME

            WS.Columns[7].Width = 256 * 15; // DEBIT
            WS.Columns[8].Width = 256 * 15; // CREDIT
            WS.Columns[9].Width = 256 * 15; // BALANCE
            WS.Columns[10].Width = 256 * 5; // TYPE
            WS.Columns[11].Width = 256 * 45; // NARRATION


            WS.Columns[12].Width = 256 * 20; // ROW DEIBT
            WS.Columns[13].Width = 256 * 20; // ROW CREDIT
            WS.Columns[14].Width = 256 * 20; // ROW BAL


            WS.Columns[7].Style.NumberFormat = "#,0.00";
            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";

            WS.Columns[10].Style.NumberFormat = "#,0.00";
            WS.Columns[11].Style.NumberFormat = "#,0.00";
            WS.Columns[12].Style.NumberFormat = "#,0.00";

            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            sTitle = "DAY BOOK OF " + acc_name + " PERIOD FROM " + Lib.getFrontEndDate(from_date) + " TO " + Lib.getFrontEndDate(to_date);

            if (ismaincode)
                sTitle += " (MAIN CODE WISE)";
            else
                sTitle += " (SUB CODE WISE)";

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;


            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "GROUP", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "NARRATION", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            if (showtotaldrcr)
            {
                Lib.WriteData(WS, iRow, iCol++, "TOT-DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOT-CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOT-BALANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            }

            foreach (DataRow Dr in Dt_List.Rows)
            {
                iRow++; iCol = 1;
                _Border = "";
                _Bold = false;
                _Color = Color.Black;

                if (Dr["rowtype"].ToString() == "TOTAL")
                {
                    _Border = "TB";
                    _Bold = true;
                    _Color = Color.DarkBlue;
                }
                sDate = Lib.DatetoStringDisplayformat(Dr["jv_date"]);

                Lib.WriteData(WS, iRow, iCol++, sDate, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, Dr["jv_type"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_vrno"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, Dr["acgrp_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acc_code"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acc_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, Dr["jv_debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["bal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_drcr"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_narration"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                if (showtotaldrcr)
                {
                    Lib.WriteData(WS, iRow, iCol++, Dr["row_debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["row_credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["row_bal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                }

            }
            WB.SaveXls(File_Name + ".xls");
        }



    }
}

