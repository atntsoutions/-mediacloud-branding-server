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
    public class TrialBalanceService : BL_Base
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

        Boolean shownote = false;


        string hide_ho_entries = "N";



        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            
            

            string sFilter = "";



            string sqlDet = "";


            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<LedgerReport> mList = new List<LedgerReport>();

            LedgerReport mrow;

            type = SearchData["type"].ToString();
            subtype = SearchData["subtype"].ToString();
            report_folder = SearchData["report_folder"].ToString();
            PKID = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            from_date = SearchData["from_date"].ToString();
            to_date = SearchData["to_date"].ToString();
            ismaincode = (Boolean ) SearchData["ismaincode"];
            shownote = (Boolean)SearchData["shownote"];
            hide_ho_entries = SearchData["hide_ho_entries"].ToString();


            from_date = Lib.StringToDate(from_date);
            to_date = Lib.StringToDate(to_date);

            decimal nBal = 0;
            decimal nDr = 0;
            decimal nCr = 0;

            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine( report_folder, PKID);

            try
            {
                
                if (type == "NEW")
                {

                    sqlDet = "";

                    sqlDet += "  select jv_acc_id, ";
                    sqlDet += "  sum(case when jvh_type  = 'OP' or  jvh_date < '{FDATE}'  then jv_debit - jv_credit else 0 end) as opbal, ";
                    sqlDet += "  sum(case when jvh_type <> 'OP' and jvh_date between '{FDATE}' and '{EDATE}' then jv_debit  else 0 end) as debit, ";
                    sqlDet += "  sum(case when jvh_type <> 'OP' and jvh_date between '{FDATE}' and '{EDATE}' then jv_credit else 0 end) as credit, ";
                    sqlDet += "  sum(jv_debit) - sum(jv_credit) as bal ";
                    sqlDet += "  from ledgerh a ";
                    sqlDet += "  inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sqlDet += "  where jvh_year = {YEAR}  and jvh_date <= '{EDATE}' and jvh_posted ='Y' and jvh_type not in('OB','OC','OI')  ";
                    if (hide_ho_entries == "Y")
                        sqlDet += "  and jvh_type  not in('HO','IN-ES')  ";
                    sqlDet += "  and a.rec_company_code = '{COMPANY}' and a.rec_branch_code = '{BRANCH}' ";
                    sqlDet += "  and a.rec_deleted = 'N'   ";
                    sqlDet += "  group by jv_acc_id ";



                    if (ismaincode)
                    {
                        sql = "";
                        sql += " select row_number() over(order  by acgrp_order, acc_main_name) as slno, '' as ROWTYPE,";
                        sql += " acgrp_order, level1, level2, acgrp_name, acc_main_id as acc_pkid, acc_main_code as acc_code, acc_main_name as acc_name, ";

                        if (shownote)
                            sql += " note_no, main_head, sub_head,sub_note, ";

                        sql += " sum(opbal) as opbal, sum(debit) as debit, sum(credit) as credit,  ";
                        sql += " sum(case when bal > 0 then bal else 0 end) as drbal,";
                        sql += " sum(case when bal < 0 then abs(bal) else 0 end) as crbal ";
                        sql += " from( ";
                        sql += sqlDet;
                        sql += " ) a ";
                        sql += " inner join acctm on jv_acc_id = acc_pkid ";
                        sql += " inner join view_acgroupm on acc_group_id = acgrp_pkid ";
                        if (shownote)
                        {
                            sql += " left join bshead on acc_bs_id = pkid ";
                            sql += " group by acgrp_order, level1, level2, acgrp_name,  acc_main_id, acc_main_code, acc_main_name, note_no, main_head, sub_head,sub_note ";
                        }
                        else
                        {
                            sql += " group by acgrp_order, level1, level2, acgrp_name,  acc_main_id, acc_main_code, acc_main_name ";
                        }

                        sql += " order by acgrp_order, acc_main_name ";
                    }
                    else 
                    {
                        sql = "";
                        sql += " select row_number() over(order  by acgrp_order, acc_name) as slno, '' as ROWTYPE,";
                        sql += " acgrp_order, level1, level2, acgrp_name,  acc_pkid, acc_code, acc_name, ";

                        if (shownote)
                            sql += " note_no, main_head, sub_head,sub_note, ";

                        sql += " opbal, debit, credit,  ";
                        sql += " case when bal > 0 then bal else 0 end as drbal,";
                        sql += " case when bal < 0 then abs(bal) else 0 end as crbal ";
                        sql += " from( ";
                        sql += sqlDet;
                        sql += " ) a ";
                        sql += " inner join acctm on jv_acc_id = acc_pkid ";
                        sql += " inner join view_acgroupm on acc_group_id = acgrp_pkid ";

                        if (shownote)
                            sql += " left join bshead on acc_bs_id = pkid ";

                        sql += " order by acgrp_order, acc_name ";
                    }

                    sql = sql.Replace("{COMPANY}", company_code);
                    sql = sql.Replace("{BRANCH}", branch_code);
                    sql = sql.Replace("{YEAR}", year_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    if ( ismaincode)
                    {
                        foreach ( DataRow Dr in Dt_List.Rows)
                        {
                            nDr = Lib.Conv2Decimal(Dr["drbal"].ToString());
                            nCr = Lib.Conv2Decimal(Dr["crbal"].ToString());
                            if (nDr > nCr)
                            {
                                nBal = nDr - nCr;
                                Dr["drbal"] = nBal;
                                Dr["crbal"] = 0;
                            }
                            else if (nCr > nDr)
                            {
                                nBal = nCr - nDr;
                                Dr["drbal"] = 0;
                                Dr["crbal"] = nBal;
                            }
                        }
                    }


                    if (Dt_List.Rows.Count > 0)
                    {
                        page_rowcount = Lib.Conv2Integer(Dt_List.Rows[Dt_List.Rows.Count - 1]["slno"].ToString());
                        page_rowcount++; // this is for total row
                        page_count = page_rowcount / page_rows;
                        if ((page_rowcount % page_rows) != 0)
                            page_count++;
                        page_current = 1;

                        DataRow Dr = Dt_List.NewRow();
                        Dr["slno"] = page_rowcount;
                        Dr["rowtype"] = "TOTAL";
                        Dr["acc_name"] = "TOTAL";

                        Dr["opbal"] = Dt_List.Compute("sum(opbal)", "1=1");
                        Dr["debit"] = Dt_List.Compute("sum(debit)", "1=1");
                        Dr["credit"] = Dt_List.Compute("sum(credit)", "1=1");

                        Dr["drbal"] = Dt_List.Compute("sum(drbal)", "1=1");
                        Dr["crbal"] = Dt_List.Compute("sum(crbal)", "1=1");
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
                        mrow.level1 = Dr["level1"].ToString();
                        mrow.level2 = Dr["level2"].ToString();
                        mrow.grp_name = Dr["acgrp_name"].ToString();

                        mrow.acc_pkid = Dr["acc_pkid"].ToString();
                        mrow.acc_code = Dr["acc_code"].ToString();
                        mrow.acc_name = Dr["acc_name"].ToString();

                        if (shownote)
                        {
                            mrow.bs_note_no = Dr["note_no"].ToString();
                            mrow.bs_main_head = Dr["main_head"].ToString();
                            mrow.bs_sub_head = Dr["sub_head"].ToString();
                            mrow.bs_sub_note = Dr["sub_note"].ToString();
                        }


                        mrow.opbal = Lib.Conv2Decimal(Dr["opbal"].ToString(), "NULL");
                        mrow.debit = Lib.Conv2Decimal(Dr["debit"].ToString(), "NULL");
                        mrow.credit = Lib.Conv2Decimal(Dr["credit"].ToString(), "NULL");
                        mrow.drbal = Lib.Conv2Decimal(Dr["drbal"].ToString(), "NULL");
                        mrow.crbal = Lib.Conv2Decimal(Dr["crbal"].ToString(), "NULL");

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

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];
            WS.ViewOptions.ShowGridLines = false;
            WS.PrintOptions.Portrait = false;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;

            WS.Columns[0].Width = 256;
            WS.Columns[1].Width = 256 * 17;
            WS.Columns[2].Width = 256 * 25;
            WS.Columns[3].Width = 256 * 20;
            WS.Columns[4].Width = 256 * 50;
            WS.Columns[5].Width = 256 * 15;
            WS.Columns[6].Width = 256 * 15;
            WS.Columns[7].Width = 256 * 15;
            WS.Columns[8].Width = 256 * 15;
            WS.Columns[9].Width = 256 * 15;

            if (shownote)
            {
                WS.Columns[10].Width = 256 * 10;
                WS.Columns[11].Width = 256 * 22;
                WS.Columns[12].Width = 256 * 30;
                WS.Columns[13].Width = 256 * 40;
            }


            WS.Columns[5].Style.NumberFormat = "#,0.00";
            WS.Columns[6].Style.NumberFormat = "#,0.00";
            WS.Columns[7].Style.NumberFormat = "#,0.00";
            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";


            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS,branch_code, iRow,iCol);

            sTitle = "TRIAL BALANCE FROM " + Lib.getFrontEndDate(from_date) + " TO " + Lib.getFrontEndDate(to_date);

            if (ismaincode)
                sTitle += " (MAIN CODE WISE)";
            else
                sTitle += " (SUB CODE WISE)";

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;

            Lib.WriteData(WS, iRow, iCol++, "MAIN-GROUP", _Color, true, _Border,"L","", _Size, false,325,"",true);
            Lib.WriteData(WS, iRow, iCol++, "SUB-GROUP", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "OP-BAL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DR-BAL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CR-BAL", _Color, true, _Border, "R", "", _Size, false, 325, "", true);

            if (shownote)
            {
                Lib.WriteData(WS, iRow, iCol++, "NOTE-NO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAIN-HEAD", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUB-HEAD", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SUB-NOTE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
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
                Lib.WriteData(WS, iRow, iCol++, Dr["level2"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acgrp_name"].ToString(), _Color,_Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acc_code"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acc_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["opbal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["drbal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["crbal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                if (shownote)
                {
                    Lib.WriteData(WS, iRow, iCol++, Dr["note_no"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["main_head"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["sub_head"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["sub_note"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                }
            }
            WB.SaveXls(File_Name + ".xls");
        }

    }
}

