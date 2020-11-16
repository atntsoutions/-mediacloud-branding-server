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
    public class LedgerBalService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_tmp = new DataTable();

        ExcelFile WB;
        ExcelWorksheet WS = null;

        int iRow = 0;
        int iCol = 0;

        Boolean bFirst = false;

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
        Boolean transdet = false;

        Boolean showtotaldrcr = false;

        string hide_ho_entries = "N";

        /**All ledger Report print**/
        private int iGroupPage = 0;
        private int iGroupPG = 0;
        private int iPrintedRow = 0;
        private int iTotalWidth = 0;
        private Boolean IsBreakPrinted = false;
        private StreamWriter SW;
        private DataTable DT_GRPTABLE = null;
        private String Report_Main_Code = "";
        private String Report_Main_Code_Desc = "";
        private string Report_Dt_From;
        private string Report_Dt_To;
        private string FinYear_Start_Date;
        private string FinYear_End_Date;
        private int TOT_LENGTH = 116;
        private int LINES_PER_PAGE = 60;
        private int LEDGER_COL1 = 12;
        private int LEDGER_COL2 = 5;
        private int LEDGER_COL3 = 6;
        private int LEDGER_COL4 = 15;
        private int LEDGER_COL5 = 15;
        private int LEDGER_COL6 = 15;
        private int LEDGER_COL7 = 3;
        private int LEDGER_COL8_F1 = 15;
        private int LEDGER_COL8_F2 = 45;
        private string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "", comp_tel = "", comp_fax = "", comp_email = "";
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
            transdet = (Boolean)SearchData["transdet"];
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
                        sql += " select 'A' as rowtype, 0 as slno, null as jvh_pkid, ";
                        sql += " null as jv_vrno, cast('OP' as nvarchar2(10)) as jv_docno, null as jv_date, null as jv_type,'' as jv_drcr,  ";
                        sql += " nvl(sum(jv_debit - jv_credit),0) as op, 0 as jv_debit, 0 as jv_credit, 0 as bal, cast('OPENING' as nvarchar2(10)) as jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " where jvh_year = {YEAR}  ";
                        sql += " and acc_main_code  = '" + acc_code + "' ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in('HO','IN-ES') ";
                        sql += " and jvh_date < '{SDATE}'  ";
                        sql += " and jvh_posted ='Y' and jvh_type not in('OB','OC','OI') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";

                        sql += "union all ";

                        sql += "select 'B' as rowtype,  0 as slno, jvh_pkid,";
                        sql += "jvh_vrno as jv_vrno,jvh_docno as jv_docno, jvh_date as jv_date, jvh_type as jv_type, '' as jv_drcr, 0 as op,  sum(jv_debit) as jv_debit, sum(jv_credit) as jv_credit,  0 as bal, jvh_narration as  jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += "from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += "inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " where jvh_year = {YEAR}  ";
                        sql += " and acc_main_code = '" + acc_code + "'";
                        sql += " and jvh_posted ='Y' and jvh_type not in('OB','OC','OI') ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in('HO','IN-ES') ";
                        sql += " and jvh_date >= '{SDATE}' and jvh_date <= '{EDATE}' ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";
                        sql += " group by jvh_pkid,jvh_vrno,jvh_type, jvh_docno, jvh_date,jvh_narration  ";
                        sql += " order by rowtype, jv_date,jv_vrno ";
                    }
                    else
                    {
                        sql = "";
                        sql += " select 'A' as rowtype, 0 as slno,null as jvh_pkid, ";
                        sql += " null as jv_vrno, cast('OP' as nvarchar2(10)) as jv_docno, null as jv_date, null as jv_type,'' as jv_drcr,  ";
                        sql += " nvl(sum(jv_debit - jv_credit),0) as op, 0 as jv_debit, 0 as jv_credit, 0 as bal, cast('OPENING' as nvarchar2(10)) as jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " where jvh_year = {YEAR}  ";
                        sql += " and jv_acc_id = '{ACCID}' ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in ('HO', 'IN-ES') ";
                        sql += " and jvh_date < '{SDATE}'  ";
                        sql += " and jvh_posted ='Y' and jvh_type not in('OB','OC','OI') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";

                        sql += "union all ";

                        sql += "select 'B' as rowtype,  0 as slno,jvh_pkid,";
                        sql += "jvh_vrno as jv_vrno,jvh_docno as jv_docno, jvh_date as jv_date, jvh_type as jv_type, '' as jv_drcr, 0 as op,  jv_debit, jv_credit,  0 as bal, jvh_narration as  jv_narration, 0 as row_debit, 0 as row_credit, 0 as row_bal ";
                        sql += "from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += "inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += " where jvh_year = {YEAR}  ";
                        sql += " and jv_acc_id = '{ACCID}' ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in('HO', 'IN-ES' ) ";
                        sql += " and jvh_date >= '{SDATE}' and jvh_date <= '{EDATE}' ";
                        sql += " and jvh_posted ='Y' and jvh_type not in('OB','OC','OI') ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " and a.rec_deleted = 'N' ";
                        sql += " order by  rowtype, jv_date,jv_vrno ";
                    }

                    sql = sql.Replace("{COMPANY}", company_code);
                    sql = sql.Replace("{BRANCH}", branch_code);
                    sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{ACCID}", acc_id);
                    sql = sql.Replace("{SDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    

                    if (transdet)
                    {
                        sql = "";

                        sql += " select a.jvh_pkid,a.jvh_vrno as jv_vrno,a.jvh_docno as jv_docno,a.jvh_date as jv_date,a.jvh_type as jv_type,a.jvh_narration as  jv_narration, 0 as slno, d.acc_name, c.jv_debit , c.jv_credit ";
                        sql += " from ledgerh a  ";
                        sql += " inner join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                        sql += " inner join ledgert c on b.jv_parent_id = c.jv_parent_id ";
                        sql += " inner join acctm d on (c.jv_acc_id = d.acc_pkid) ";
                        sql += " where b.jv_acc_id = '{ACCID}' ";
                        sql += " and c.jv_acc_id <> '{ACCID}' ";
                        sql += " and a.jvh_type not in ('OP', 'OB', 'OI','OC', 'IN', 'IN-ES') ";
                        if (hide_ho_entries == "Y")
                            sql += "  and jvh_type not in('HO', 'IN-ES' ) ";
                        sql += " and a.jvh_year = {YEAR} and a.jvh_date between '{SDATE}' and '{EDATE}' ";
                        sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                        sql += " order by a.jvh_pkid, c.jv_ctr ";

                        sql = sql.Replace("{COMPANY}", company_code);
                        sql = sql.Replace("{BRANCH}", branch_code);
                        sql = sql.Replace("{YEAR}", year_code);
                        sql = sql.Replace("{ACCID}", acc_id);
                        sql = sql.Replace("{SDATE}", from_date);
                        sql = sql.Replace("{EDATE}", to_date);

                        Dt_tmp = Con_Oracle.ExecuteQuery(sql);

                    }
                    Con_Oracle.CloseConnection();


                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        if (Dr["rowtype"].ToString() == "B")
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
                                nDr = nAmt;
                            }
                            if (nAmt < 0)
                            {
                                nCr = Math.Abs(nAmt);
                                Dr["jv_credit"] = nCr;
                                Dr["row_credit"] = nCr;
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
                            if (transdet)
                            {
                                Dt_tmp.TableName = "REPORTDET";
                                Dt_tmp.WriteXml(File_Name + "DET.xml", XmlWriteMode.WriteSchema);
                            }
                        }
                    }
                }
                else
                {
                    Dt_List.ReadXml(File_Name + ".xml");
                    try
                    {
                        if (transdet)
                            Dt_tmp.ReadXml(File_Name + "DET.xml");
                    }
                    catch (Exception )
                    {}
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

                        if (transdet)
                        {
                            bFirst = true;
                            foreach ( DataRow Dr1 in Dt_tmp.Select ("JVH_PKID='" + Dr["JVH_PKID"].ToString() + "'" ))
                            {
                                if (!bFirst)
                                    mrow = new LedgerReport();

                                mrow.cb_desc = Dr1["acc_name"].ToString();
                                mrow.cb_dr = Lib.Conv2Decimal(Dr1["jv_debit"].ToString(), "NULL");

                                mrow.cb_cr = Lib.Conv2Decimal(Dr1["jv_credit"].ToString(), "NULL");
                                if (!bFirst)
                                {
                                    mrow.rowtype = "B";
                                    mrow.acc_pkid = acc_id;
                                    mrow.rec_company_code = company_code;
                                    mrow.rec_branch_code = branch_code;
                                    mrow.jv_year = year_code;

                                    mrow.jv_docno = Dr1["jv_docno"].ToString();
                                    mrow.jv_type = Dr1["jv_type"].ToString();
                                    mrow.jv_vrno = Dr1["jv_vrno"].ToString();
                                    mrow.jv_date = Lib.DatetoStringDisplayformat(Dr1["jv_date"]);
                                    mrow.jv_narration = Dr1["jv_narration"].ToString();
                                }

                                if (!bFirst)
                                    mList.Add(mrow);

                                bFirst = false;
                            }
                        }
                        
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

            WS.Columns[4].Width = 256 * 15; // DEBIT
            WS.Columns[5].Width = 256 * 15; // CREDIT
            WS.Columns[6].Width = 256 * 15; // BALANCE
            WS.Columns[7].Width = 256 * 5; // TYPE
            WS.Columns[8].Width = 256 * 45; // NARRATION

            WS.Columns[4].Style.NumberFormat = "#,0.00";
            WS.Columns[5].Style.NumberFormat = "#,0.00";
            WS.Columns[6].Style.NumberFormat = "#,0.00";

            int iCol = 9;

            if (showtotaldrcr)
            {
                WS.Columns[9].Width = 256 * 20; // ROW DEIBT
                WS.Columns[10].Width = 256 * 20; // ROW CREDIT
                WS.Columns[11].Width = 256 * 20; // ROW BAL

                WS.Columns[9].Style.NumberFormat = "#,0.00";
                WS.Columns[10].Style.NumberFormat = "#,0.00";
                WS.Columns[11].Style.NumberFormat = "#,0.00";

                iCol = 12;
            }

            if ( transdet)
            {
                WS.Columns[iCol++].Width = 256 * 30; // ROW NAME
                WS.Columns[iCol++].Width = 256 * 20; // ROW DR
                WS.Columns[iCol].Style.NumberFormat = "#,0.00";
                WS.Columns[iCol++].Width = 256 * 20; // ROW CR
                WS.Columns[iCol].Style.NumberFormat = "#,0.00";
            }




            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            sTitle = "LEDGER OF " + acc_name + " PERIOD FROM " + Lib.getFrontEndDate(from_date) + " TO " + Lib.getFrontEndDate(to_date);

            if (ismaincode)
                sTitle += " (MAIN CODE WISE)";
            else
                sTitle += " (SUB CODE WISE)";

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;



            int _iCol = 0;

            int tcol = 0;



            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
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
            if ( transdet)
            {
                Lib.WriteData(WS, iRow, iCol++, "A/C NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
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

                if (transdet)
                {
                    bFirst = true;
                    _iCol = iCol;
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    iCol = _iCol;

                    foreach (DataRow Dr1 in Dt_tmp.Select("JVH_PKID='" + Dr["JVH_PKID"].ToString() + "'"))
                    {
                        if (!bFirst)
                            iRow++;
                        iCol = _iCol;
                        if (!bFirst)
                        {
                            sDate = Lib.DatetoStringDisplayformat(Dr1["jv_date"]);
                            Lib.WriteData(WS, iRow, 1, sDate, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, 2, Dr1["jv_type"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, 3, Dr1["jv_vrno"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, 8, Dr1["jv_narration"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Dr1["acc_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "#,0.00", true);
                        Lib.WriteData(WS, iRow, iCol++, Dr1["jv_debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                        Lib.WriteData(WS, iRow, iCol++, Dr1["jv_credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                        bFirst = false;
                    }
                }

                if (Dr["rowtype"].ToString() == "TOTAL")
                {

                }


            }
            WB.SaveXls(File_Name + ".xls");
        }

        public IDictionary<string, object> GenerateLedger(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
   
            report_folder = SearchData["report_folder"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            Con_Oracle = new DBConnection();
            try
            {
                FinYear_Start_Date = "";
                FinYear_End_Date = "";
                sql = "select year_start_date,year_end_date from yearm where rec_company_code ='" + company_code + "' and year_code =" + year_code;
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                if (Dt_List.Rows.Count > 0)
                {
                    FinYear_Start_Date = Lib.DatetoStringDisplayformat(Dt_List.Rows[0]["YEAR_START_DATE"]);
                    FinYear_End_Date = Lib.DatetoStringDisplayformat(Dt_List.Rows[0]["YEAR_END_DATE"]);
                }

                ReadCompanyDetails();

                sql = "select acgrp_pkid, acgrp_name,acgrp_order from acgroupm where rec_company_code='" + company_code + "' and acgrp_parent_id is not null order by acgrp_order, acgrp_name";
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
 
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    SaveToText(Dr["acgrp_pkid"].ToString(), Dr["acgrp_name"].ToString(), Dr["acgrp_order"].ToString().Trim().PadLeft(3, '0'));
                }

                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", type);
            return RetData;
        }




        public void SaveToText(string ID, string GrpName, string sOrder)
        {
            try
            {
                DataRow Dr_Tot = null;
                int iCtr = 0;
                decimal nDr = 0;
                decimal nCr = 0;
                decimal nBal = 0;

                decimal DrTot = 0;
                decimal CrTot = 0;

                iGroupPage = 0;

                GrpName = GrpName.Replace("\\", "");
                GrpName = GrpName.Replace("/", "");

                string sFileName = report_folder + "\\LEDGER\\" + branch_code;

                if (!System.IO.Directory.Exists(sFileName))
                    System.IO.Directory.CreateDirectory(sFileName);

                sFileName += "\\" + sOrder + "_" + GrpName + ".txt";

                SW = File.CreateText(sFileName);

                CreateGroupTable();

                sql = " select distinct acc_main_id, acc_main_code,acc_main_name from acctm ";
                sql += " where rec_company_code = '" + company_code + "' and acc_group_id='" + ID + "'";
                sql += " order by acc_main_name";

                DataTable Dt_MainCode = new DataTable();
                Dt_MainCode = Con_Oracle.ExecuteQuery(sql);
             
                sql = "  select  acc_main_code, case when jvh_type ='OP' then 1 else 2 end as roworder,";
                sql += " jvh_date as jv_date, jvh_vrno as jv_vrno, jvh_type as jv_type,jvh_narration as jv_narration,";
                sql += " sum(jv_debit) as jv_debit, sum(jv_credit) as jv_credit, 0 as opening,";
                sql += " '' as type,0 as balance,'' as rowtype ";
                sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                sql += " where a.rec_company_code = '" + company_code + "'";
                sql += " and a.rec_branch_code = '" + branch_code + "'";
                sql += " and jvh_year = " + year_code;
                sql += " and jvh_type not in('OB','OC','OI') ";
                sql += " and jvh_posted ='Y'";
                sql += " and a.rec_deleted = 'N'";
                sql += " and acc_group_id = '" + ID + "'";
                sql += " group by jvh_date,jvh_vrno,jvh_type,jvh_narration,acc_main_code";
                sql += " order by acc_main_code,roworder,jvh_date,jvh_vrno,jvh_type";

                DataTable Dt_Ledger = new DataTable();
                Dt_Ledger = Con_Oracle.ExecuteQuery(sql);

                DataTable Dt_SubLedger;

                foreach (DataRow Dr in Dt_MainCode.Rows)
                {
                    Dt_SubLedger = new DataTable();
                    Dt_SubLedger = Dt_Ledger.Clone();

                    nDr = 0; nCr = 0; nBal = 0; DrTot = 0; CrTot = 0;
                    foreach (DataRow dr in Dt_Ledger.Select("acc_main_code='" + Dr["acc_main_code"].ToString() + "'", "roworder,jv_date,jv_vrno,jv_type"))
                    {
                        nDr = Lib.Conv2Decimal(dr["jv_debit"].ToString());
                        nCr = Lib.Conv2Decimal(dr["jv_credit"].ToString());
                        if (nDr > nCr)
                        {
                            nDr = nDr - nCr;
                            nCr = 0;
                            dr["jv_debit"] = nDr;
                            dr["jv_credit"] = nCr;
                        }
                        else if (nCr > nDr)
                        {
                            nCr = nCr - nDr;
                            nDr = 0;
                            dr["jv_debit"] = nDr;
                            dr["jv_credit"] = nCr;
                        }

                        DrTot += nDr;
                        CrTot += nCr;

                        nBal += nDr;
                        nBal -= nCr;
                        if (nBal > 0)
                            dr["type"] = "DR";
                        if (nBal < 0)
                            dr["type"] = "CR";

                        dr["balance"] = Math.Abs(nBal);

                        dr.AcceptChanges();
                        Dt_SubLedger.ImportRow(dr);

                    }

                    if (Dt_SubLedger.Rows.Count > 0)
                    {
                        Report_Main_Code = Dr["acc_main_code"].ToString();
                        Report_Main_Code_Desc = Dr["acc_main_name"].ToString();
                        Report_Dt_From = FinYear_Start_Date;
                        Report_Dt_To = FinYear_End_Date;
                        iGroupPage++;
                        iGroupPG = iGroupPage;

                        Dr_Tot = Dt_SubLedger.NewRow();
                        Dr_Tot["rowtype"] = "TOTAL";

                        Dr_Tot["jv_debit"] = DrTot;
                        Dr_Tot["jv_credit"] = CrTot;
                        Dr_Tot["balance"] = Math.Abs(nBal);
                        if (nBal > 0)
                            Dr_Tot["type"] = "DR";
                        if (nBal < 0)
                            Dr_Tot["type"] = "CR";

                        Dt_SubLedger.Rows.Add(Dr_Tot);
                        Dt_SubLedger.AcceptChanges();

                        WriteToFile(Dt_SubLedger);
                    }
                    iCtr++;
                }
                if (Dt_MainCode.Rows.Count > 0)
                    PrintSummary();

                SW.Close();
            }
            catch (Exception Ex)
            {
                if (SW != null)
                    SW.Close();
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void CreateGroupTable()
        {
            DT_GRPTABLE = new DataTable();
            DT_GRPTABLE.Columns.Add("ACC_PAGE", typeof(System.Decimal));
            DT_GRPTABLE.Columns.Add("ACC_CODE", typeof(System.String));
            DT_GRPTABLE.Columns.Add("ACC_NAME", typeof(System.String));
            DT_GRPTABLE.Columns.Add("ACC_DEBIT", typeof(System.Decimal));
            DT_GRPTABLE.Columns.Add("ACC_CREDIT", typeof(System.Decimal));
            DT_GRPTABLE.Columns.Add("ACC_BALANCE", typeof(System.Decimal));
        }

        private void WriteToFile(DataTable Dt_Report)
        {
            int iPage = 1;

            int LEN_NARR = LEDGER_COL8_F2;
            int BLANK_LEFT = getBlankColWidth();

            int iRow = 0;
            iPrintedRow = 0;
            SetColWidth();

            string sLine = "";
            string sNarration = "";
            string sData = "";

            int iLen = 0;
            int TotalRows = Dt_Report.Rows.Count;
            TotalRows += 5;
            iPrintedRow += 5;
            GetAddress();
            getHeading();
            iRow = 7;
            DataRow Drt = DT_GRPTABLE.NewRow();
            foreach (DataRow row in Dt_Report.Rows)
            {
                sLine = "";
                if (!row["JV_DATE"].Equals(DBNull.Value))
                    sLine += Lib.DatetoStringDisplayformat(row["JV_DATE"]).ToString().Replace(" 12:00", "").PadRight(LEDGER_COL1);
                else
                    sLine += "".PadRight(LEDGER_COL1);
                sLine += row["JV_TYPE"].ToString().PadRight(LEDGER_COL2);
                sLine += row["JV_VRNO"].ToString().PadLeft(LEDGER_COL3);
                sLine += Lib.NumericFormat(row["JV_DEBIT"].ToString(), 2).PadLeft(LEDGER_COL4);
                sLine += Lib.NumericFormat(row["JV_CREDIT"].ToString(), 2).PadLeft(LEDGER_COL5);
                sLine += Lib.NumericFormat(row["BALANCE"].ToString(), 2).PadLeft(LEDGER_COL6);
                sLine += row["TYPE"].ToString().PadRight(LEDGER_COL7);
                sNarration = row["JV_NARRATION"].ToString().Trim();
                sData = sNarration;
                iLen = sNarration.Length;
                if (iLen > LEN_NARR)
                {
                    sData = sNarration.Substring(0, LEN_NARR);
                    sNarration = sNarration.Substring(LEN_NARR);
                }
                sLine += sData.PadRight(LEN_NARR);

                //  string sPdfNarrtn = sData;
                SW.WriteLine(sLine);
                //if (ChkPdf.Checked)
                //{
                //    string sJvdate = "";
                //    if (row.Cells["JV_DATE"].Value != null)
                //        sJvdate = row.Cells["JV_DATE"].FormattedValue.ToString().Replace(" 12:00", "");

                //    PDF.WriteDetailsData(sJvdate, row.Cells["JV_TYPE"].Value.ToString(), row.Cells["JV_VRNO"].Value.ToString(),
                //        row.Cells["JV_DEBIT"].FormattedValue.ToString(), row.Cells["JV_CREDIT"].FormattedValue.ToString(),
                //        row.Cells["BALANCE"].FormattedValue.ToString() + row.Cells["TYPE"].Value.ToString()
                //        , sPdfNarrtn);
                //}

                IsBreakPrinted = false;
                iRow++;
                checkPageBreak(TotalRows, ref iRow, ref iPage);
                while (sNarration != "")
                {
                    iLen = sNarration.Length;
                    sData = sNarration;
                    if (iLen > LEN_NARR)
                    {
                        sData = sNarration.Substring(0, LEN_NARR);
                        sNarration = sNarration.Substring(LEN_NARR);
                    }
                    else
                        sNarration = "";
                    sLine = "";
                    sLine += "".PadLeft(BLANK_LEFT);
                    sLine += sData.PadRight(LEN_NARR);
                    SW.WriteLine(sLine);
                    //if (ChkPdf.Checked)
                    //    PDF.WriteDetailsData("", "", "", "", "", "", sData);
                    IsBreakPrinted = false;
                    iRow++;
                    checkPageBreak(TotalRows, ref iRow, ref iPage);
                }

                if (row["ROWTYPE"].ToString() == "TOTAL")
                {
                    Drt["ACC_CODE"] = Report_Main_Code;
                    Drt["ACC_NAME"] = Report_Main_Code_Desc;
                    Drt["ACC_DEBIT"] = Lib.Convert2Decimal(Lib.NumericFormat(row["JV_DEBIT"].ToString(), 2)).ToString();
                    Drt["ACC_CREDIT"] = Lib.Convert2Decimal(Lib.NumericFormat(row["JV_CREDIT"].ToString(), 2)).ToString();
                    Drt["ACC_BALANCE"] = Lib.Convert2Decimal(Lib.NumericFormat(row["BALANCE"].ToString(), 2)).ToString();
                    Drt["ACC_PAGE"] = iGroupPG.ToString();
                }
            }
            DT_GRPTABLE.Rows.Add(Drt);
            if (IsBreakPrinted == false)
            {
                iRow = LINES_PER_PAGE + 1;
                PutPageBreak(iPage);
            }
        }
        private void checkPageBreak(int iTotRow, ref int iRow, ref int iPage)
        {
            string str = "";
            //iPrintedRow++;
            if (iRow >= LINES_PER_PAGE)
            {
                iRow = 2;
                str = "Page# " + iPage.ToString();
                str += "".PadLeft(20) + "Group Page# " + iGroupPage.ToString() + "\f";
                SW.WriteLine(getCenter(str));
                //if (ChkPdf.Checked)
                //    PDF.WritePageNo(iPage.ToString(), iGroupPage.ToString());
                if (iPrintedRow < iTotRow)
                    getHeading();
                iPage++;
                iGroupPage++;
                IsBreakPrinted = true;
            }
        }

        private void PutPageBreak(int iPage)
        {
            string str = "";
            iPrintedRow++;
            iRow = 2;
            str = "Page " + iPage.ToString();
            str += "".PadLeft(20) + "Group Page# " + iGroupPage.ToString() + "\f";
            SW.WriteLine(getCenter(str));
            //if (ChkPdf.Checked)
            //{
            //    PDF.WritePageNo(iPage.ToString(), iGroupPage.ToString());
            //    PDF.GetNewPage();
            //}
            IsBreakPrinted = true;
        }

        private string getCenter(string Str)
        {
            int iLen = Str.Length;
            iLen = (iTotalWidth - iLen) / 2;
            if (iLen < 0)
                iLen = 0;
            return "".PadLeft(iLen) + Str;
        }

        private void SetColWidth()
        {
            int nTot = 0;
            nTot += LEDGER_COL1;
            nTot += LEDGER_COL2;
            nTot += LEDGER_COL3;
            nTot += LEDGER_COL4;
            nTot += LEDGER_COL5;
            nTot += LEDGER_COL6;
            nTot += LEDGER_COL7;
            nTot += LEDGER_COL8_F1;
            iTotalWidth = nTot;
        }
        private int getBlankColWidth()
        {
            int nTot = 0;
            nTot += LEDGER_COL1;
            nTot += LEDGER_COL2;
            nTot += LEDGER_COL3;
            nTot += LEDGER_COL4;
            nTot += LEDGER_COL5;
            nTot += LEDGER_COL6;
            nTot += LEDGER_COL7;

            return nTot;
        }
        private void GetAddress()
        {
            string str = "";
            string LINE1 = comp_name.ToUpper();
            str = comp_add1.ToUpper();
            if(comp_add2!="")
            {
                if (str != "" && !str.Trim().EndsWith(","))
                    str += ",";
                str += comp_add2.ToUpper();
            }
            if (comp_add3 != "")
            {
                if (str != "" && !str.Trim().EndsWith(","))
                    str += ",";
                str += comp_add3.ToUpper();
            }

            string LINE2 = str;
            str = "TEL : " + comp_tel;
            str += " FAX : " + comp_fax;
            string LINE3 = str;
            string LINE4 = "Email :" + comp_email.ToLower();
            SW.WriteLine(getCenter(LINE1));
            SW.WriteLine(getCenter(LINE2));
            SW.WriteLine(getCenter(LINE3));
            SW.WriteLine(getCenter(LINE4));
            //SW.WriteLine("");
            //if (ChkPdf.Checked)
            //    PDF.GetAddress();
        }
        private void getHeading()
        {
            string str = "LEDGER: ";
            //if (Report_Code.Length > 0)
            //    str += Report_DESC; // + "-" + Report_Code;
            //else
            str += Report_Main_Code_Desc; // +"-" + Report_Main_Code;
            str += " FROM " + FinYear_Start_Date;
            str += " TO " + FinYear_End_Date;
            SW.WriteLine(str);
            //if (ChkPdf.Checked)
            //    PDF.Report_Caption = str;

            str = "";
            SW.WriteLine("".PadLeft(TOT_LENGTH, '-'));
            str += "DATE".PadRight(LEDGER_COL1);
            str += "TYPE".PadRight(LEDGER_COL2);
            str += "VRNO".PadLeft(LEDGER_COL3);
            str += "DEBIT".PadLeft(LEDGER_COL4);
            str += "CREDIT".PadLeft(LEDGER_COL5);
            str += "BALANCE".PadLeft(LEDGER_COL6);
            str += "   NARRATION".PadRight(LEDGER_COL7);
            SW.WriteLine(str);
            SW.WriteLine("".PadLeft(TOT_LENGTH, '-'));

            //if (ChkPdf.Checked)
            //    PDF.getHeading();
        }

        private void ReadCompanyDetails()
        {
            comp_name = ""; comp_add1 = ""; comp_add2 = ""; comp_add3 = ""; comp_tel = ""; comp_fax = ""; comp_email = "";
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
                    comp_tel = Dr["COMP_TEL"].ToString();
                    comp_fax = Dr["COMP_FAX"].ToString();
                    //comp_web = Dr["COMP_WEB"].ToString();
                    comp_email = Dr["COMP_EMAIL"].ToString();
                    //comp_cinno = Dr["COMP_CINNO"].ToString();
                    //comp_gstin = Dr["COMP_GSTIN"].ToString();
                    break;
                }
            }
        }

        private void PrintSummary()
        {
            decimal nDr = 0;
            decimal nCr = 0;
            decimal nNet = 0;

            decimal nDrTot = 0;
            decimal nCrTot = 0;
            try
            {
                string sLine = "";


                GetAddress();
                SW.WriteLine("ACCOUNT SUMMARY");
                SW.WriteLine("".PadLeft(TOT_LENGTH, '-'));

                SW.WriteLine("PAGE".PadRight(6) + "CODE".PadRight(15) + "NAME".PadRight(20) + "DEBIT".PadLeft(14) + "CREDIT".PadLeft(14) + "BAL-DR".PadLeft(14) + "BAL-CR".PadLeft(14));
                SW.WriteLine("".PadLeft(TOT_LENGTH, '-'));

                //if (ChkPdf.Checked)
                //    PDF.WriteSummaryHeaderData();
                foreach (DataRow Dr in DT_GRPTABLE.Rows)
                {
                    sLine = "";
                    sLine += Dr["ACC_PAGE"].ToString().PadRight(6);
                    sLine += Dr["ACC_CODE"].ToString().PadRight(15);
                    sLine += Dr["ACC_NAME"].ToString().PadRight(30).Substring(0, 20);
                    sLine += Dr["ACC_DEBIT"].ToString().PadLeft(14);
                    sLine += Dr["ACC_CREDIT"].ToString().PadLeft(14);

                    //sLine += Dr["ACC_BALANCE"].ToString().PadLeft(14);

                    nDr = Lib.Convert2Decimal(Dr["ACC_DEBIT"].ToString());
                    nCr = Lib.Convert2Decimal(Dr["ACC_CREDIT"].ToString());

                    if (nDr > nCr)
                    {
                        sLine += Dr["ACC_BALANCE"].ToString().PadLeft(14);
                        sLine += "".ToString().PadLeft(14);
                        nDrTot += Lib.Convert2Decimal(Dr["ACC_BALANCE"].ToString());
                    }
                    else if (nCr > nDr)
                    {
                        sLine += "".ToString().PadLeft(14);
                        sLine += Dr["ACC_BALANCE"].ToString().PadLeft(14);
                        nCrTot += Lib.Convert2Decimal(Dr["ACC_BALANCE"].ToString());
                    }
                    SW.WriteLine(sLine);

                    //if (ChkPdf.Checked)
                    //{

                    //    if (nDr > nCr)
                    //    {
                    //        PDF.WriteSummaryData(Dr["ACC_PAGE"].ToString(), Dr["ACC_CODE"].ToString(), Dr["ACC_NAME"].ToString().PadRight(30).Substring(0, 20), Dr["ACC_DEBIT"].ToString(), Dr["ACC_CREDIT"].ToString(), Dr["ACC_BALANCE"].ToString(), "");
                    //    }
                    //    else if (nCr > nDr)
                    //    {
                    //        PDF.WriteSummaryData(Dr["ACC_PAGE"].ToString(), Dr["ACC_CODE"].ToString(), Dr["ACC_NAME"].ToString().PadRight(30).Substring(0, 20), Dr["ACC_DEBIT"].ToString(), Dr["ACC_CREDIT"].ToString(), "", Dr["ACC_BALANCE"].ToString());
                    //    }
                    //    else
                    //    {
                    //        PDF.WriteSummaryData(Dr["ACC_PAGE"].ToString(), Dr["ACC_CODE"].ToString(), Dr["ACC_NAME"].ToString().PadRight(30).Substring(0, 20), Dr["ACC_DEBIT"].ToString(), Dr["ACC_CREDIT"].ToString(), "", "");
                    //    }
                    //}
                }

                SW.WriteLine("".PadLeft(TOT_LENGTH, '-'));
                SW.WriteLine("TOTAL".PadRight(6) + "".PadRight(15) + "".PadRight(20) + "".PadLeft(14) + "".PadLeft(14) + nDrTot.ToString().PadLeft(14) + nCrTot.ToString().PadLeft(14));
                SW.WriteLine("".PadLeft(TOT_LENGTH, '-'));
                //if (ChkPdf.Checked)
                //    PDF.WriteSummaryData("TOTAL", "", "", "", "", nDrTot.ToString(), nCrTot.ToString());
                nNet = nDrTot - nCrTot;
                if (nNet > 0)
                {
                    SW.WriteLine("Net - " + nNet.ToString() + " DR");
                    //if (ChkPdf.Checked)
                    //    PDF.WriteSummaryData("Net - ", nNet.ToString() + " DR", "", "", "", "", "");
                }
                if (nNet < 0)
                {
                    nNet = Math.Abs(nNet);
                    SW.WriteLine("Net - " + nNet.ToString() + " CR");
                    //if (ChkPdf.Checked)
                    //    PDF.WriteSummaryData("Net - ", nNet.ToString() + " CR", "", "", "", "", "");
                }
                SW.WriteLine("".PadLeft(TOT_LENGTH, '-'));
                SW.WriteLine("\f");
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

