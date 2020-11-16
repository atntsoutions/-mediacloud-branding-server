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
    public class FcBalService : BL_Base
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
        string acc_name = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string curr_code = "";
        string from_date = "";
        string to_date = "";
        Boolean ismaincode = false;
        string hide_ho_entries = "N";

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sFilter = "";

            decimal nAmt = 0;
            decimal nDr = 0;
            decimal nCr = 0;
            decimal nBal = 0;

            decimal DrTot = 0;
            decimal CrTot = 0;



            decimal nfAmt = 0;
            decimal nfDr = 0;
            decimal nfCr = 0;
            decimal nfBal = 0;

            decimal DrfTot = 0;
            decimal CrfTot = 0;




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
            acc_name = SearchData["acc_name"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            curr_code = SearchData["curr_code"].ToString().ToUpper();

            from_date = SearchData["from_date"].ToString();
            to_date = SearchData["to_date"].ToString();
            ismaincode = (Boolean)SearchData["ismaincode"];

            hide_ho_entries = SearchData["hide_ho_entries"].ToString();


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


                if (type == "NEW")
                {


                    sql += " select 'A' as rowtype, 0 as slno, ";
                    sql += " null as jv_vrno, cast('OP' as nvarchar2(10)) as jv_docno, null as jv_date, null as jv_type,'' as jv_drcr, '' as fdrcr, cast('' as nvarchar2(10)) as curr_code, 0 as exrate, ";
                    sql += " nvl(sum(jv_debit - jv_credit),0) as op, 0 as jv_debit, 0 as jv_credit, 0 as bal, ";

                    sql += " sum(case when jv_drcr = 'DR' then jv_ftotal else 0 end) as fdr,";
                    sql += " sum(case when jv_drcr = 'CR' then jv_ftotal else 0 end) as fcr,";
                    sql += " 0 as  fbal,";

                    sql += " cast('OPENING' as nvarchar2(10)) as jv_narration "; 

                    sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                    sql += "inner join param d on b.jv_curr_id = d.param_pkid ";
                    sql += " where jvh_year = {YEAR}  ";
                    sql += " and jv_acc_id = '{ACCID}' ";

                    if (hide_ho_entries == "Y")
                        sql += "  and jvh_type not in( 'HO' ,'IN-ES' ) ";

                    sql += " and jvh_date < '{SDATE}'  ";
                    sql += " and jvh_posted ='Y' and param_code ='{CURR_CODE}'"  ;
                    sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                    sql += " and a.rec_deleted = 'N' ";

                    sql += "union all ";

                    sql += "select 'B' as rowtype,  0 as slno,";
                    sql += "jvh_vrno as jv_vrno,jvh_docno as jv_docno, jvh_date as jv_date, jvh_type as jv_type, '' as jv_drcr, '' as fdrcr, param_code as curr_code, jv_exrate, ";
                    sql += " 0 as op,  jv_debit, jv_credit,  0 as bal, ";

                    sql += " case when jv_drcr = 'DR' then jv_ftotal else 0 end as fdr,";
                    sql += " case when jv_drcr = 'CR' then jv_ftotal else 0 end as fcr,";
                    sql += " 0 as  fbal,";

                    sql += " jvh_narration as  jv_narration ";

                    sql += "from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += "inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                    sql += "inner join param d on b.jv_curr_id = d.param_pkid ";
                    sql += " where jvh_year = {YEAR}  ";
                    sql += " and jv_acc_id = '{ACCID}' ";
                    if (hide_ho_entries == "Y")
                        sql += "  and jvh_type not in( 'HO', 'IN-ES' ) ";
                    sql += " and jvh_date >= '{SDATE}' and jvh_date <= '{EDATE}' ";
                    sql += " and jvh_posted ='Y' and  param_code ='{CURR_CODE}'";
                    sql += " and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                    sql += " and a.rec_deleted = 'N' ";
                    sql += " order by  rowtype, jv_date ";


                    sql = sql.Replace("{COMPANY}", company_code);
                    sql = sql.Replace("{BRANCH}", branch_code);
                    sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{ACCID}", acc_id);
                    sql = sql.Replace("{SDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{CURR_CODE}", curr_code);

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
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
                                nDr = nAmt;
                            }
                            if (nAmt < 0)
                            {
                                nCr = Math.Abs(nAmt);
                                Dr["jv_credit"] = nCr;
                            }


                            nfAmt = Lib.Conv2Decimal(Dr["fdr"].ToString()) - Lib.Conv2Decimal(Dr["fcr"].ToString());
                            if (nfAmt > 0)
                            {
                                nfDr = nfAmt;
                                Dr["fdr"] = nfDr;
                                Dr["fcr"] = 0;
                            }
                            if (nfAmt < 0)
                            {
                                nfCr = Math.Abs(nfAmt);
                                Dr["fdr"] = 0;
                                Dr["fcr"] = nfCr;
                            }

                        }
                        else
                        {
                            nDr = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                            nCr = Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                            nfDr = Lib.Conv2Decimal(Dr["fdr"].ToString());
                            nfCr = Lib.Conv2Decimal(Dr["fcr"].ToString());

                        }

                        Dr["curr_code"] = Dr["curr_code"].ToString();
                        Dr["exrate"] = Lib.Convert2Decimal(Dr["exrate"].ToString());


                        DrTot += nDr;
                        CrTot += nCr;

                        nBal += nDr;
                        nBal -= nCr;
                        if (nBal > 0)
                            Dr["jv_drcr"] = "DR";
                        if (nBal < 0)
                            Dr["jv_drcr"] = "CR";

                        Dr["bal"] = Math.Abs(nBal);



                        DrfTot += nfDr;
                        CrfTot += nfCr;

                        nfBal += nfDr;
                        nfBal -= nfCr;
                        if (nfBal > 0)
                            Dr["fdrcr"] = "DR";
                        if (nfBal < 0)
                            Dr["fdrcr"] = "CR";

                        Dr["fbal"] = Math.Abs(nfBal);



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


                        Dr["fdr"] = DrfTot;
                        Dr["fcr"] = CrfTot;
                        Dr["fbal"] = Math.Abs(nfBal);
                        if (nfBal > 0)
                            Dr["fdrcr"] = "DR";
                        if (nfBal < 0)
                            Dr["fdrcr"] = "CR";


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
                        mrow.jv_date = Lib.DatetoStringDisplayformat(Dr["jv_date"]);
                        mrow.debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString(), "NULL");
                        mrow.credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString(), "NULL");
                        mrow.bal = Lib.Conv2Decimal(Dr["bal"].ToString(), "NULL");
                        mrow.jv_drcr = Dr["jv_drcr"].ToString();


                        mrow.curr_code = Dr["curr_code"].ToString();
                        mrow.exrate = Lib.Conv2Decimal(Dr["exrate"].ToString(), "NULL");

                        mrow.fdr = Lib.Conv2Decimal(Dr["fdr"].ToString(), "NULL");
                        mrow.fcr = Lib.Conv2Decimal(Dr["fcr"].ToString(), "NULL");
                        mrow.fbal = Lib.Conv2Decimal(Dr["fbal"].ToString(), "NULL");
                        mrow.fdrcr = Dr["fdrcr"].ToString();

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
            WS.Columns[1].Width = 256 * 15; // VRNO
            WS.Columns[2].Width = 256 * 15; // DATE

            WS.Columns[3].Width = 256 * 15; // CURR
            WS.Columns[4].Width = 256 * 15; // EXRATE

            WS.Columns[5].Width = 256 * 15; // DEBIT
            WS.Columns[6].Width = 256 * 15; // CREDIT
            WS.Columns[7].Width = 256 * 15; // BALANCE
            WS.Columns[8].Width = 256 * 5; // TYPE

            WS.Columns[9].Width = 256 * 15; // DEBIT
            WS.Columns[10].Width = 256 * 15; // CREDIT
            WS.Columns[11].Width = 256 * 15; // BALANCE
            WS.Columns[12].Width = 256 * 5; // TYPE

            WS.Columns[13].Width = 256 * 45; // NARRATION


            WS.Columns[4].Style.NumberFormat = "#,0.000";
            WS.Columns[5].Style.NumberFormat = "#,0.00";
            WS.Columns[6].Style.NumberFormat = "#,0.00";
            WS.Columns[6].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";
            WS.Columns[10].Style.NumberFormat = "#,0.00";
            WS.Columns[11].Style.NumberFormat = "#,0.00";


            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            sTitle = curr_code +  " STATEMENT " + acc_name + " PERIOD FROM " + Lib.getFrontEndDate(from_date) + " TO " + Lib.getFrontEndDate(to_date);


            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);


            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;

            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "CURR", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "EXRATE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "DEBIT-FC", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT-FC", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BAL-FC", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "DEBIT-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT-INR", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BAL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);



            Lib.WriteData(WS, iRow, iCol++, "NARRATION", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

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
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_docno"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, sDate, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, Dr["curr_code"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["exrate"].ToString(), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.000", true);

                Lib.WriteData(WS, iRow, iCol++, Dr["fdr"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["fcr"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["fbal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["fdrcr"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, Dr["jv_debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["bal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_drcr"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, Dr["jv_narration"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            }
            WB.SaveXls(File_Name + ".xls");
        }
    }
}

