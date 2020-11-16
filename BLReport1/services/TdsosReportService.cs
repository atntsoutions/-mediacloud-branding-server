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


namespace BLReport1
{
    public class TdsosReportService : BL_Base
    {

        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<TdsosReport> mList = new List<TdsosReport>();
        TdsosReport mrow;
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
        string searchexpid = "";
        string format_type = "";
        string from_date = "";
        string to_date = "";
        string ErrorMessage = "";
        string party_name = "";

        Boolean bCompany = false;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<TdsosReport>();
            ErrorMessage = "";
            DataRow Dr_Target = null;
            decimal cert_amt = 0;
            try
            {
                type = SearchData["type"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                format_type = SearchData["format_type"].ToString();

                if (SearchData.ContainsKey("party_name"))
                    party_name = SearchData["party_name"].ToString();

                if (SearchData.ContainsKey("iscompany"))
                    bCompany = (Boolean)SearchData["iscompany"];

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (format_type == "BRANCH-WISE")
                {
                    sql = " select a.rec_branch_code,";
                    sql += "    sum(nvl(jv_debit,0)-nvl(jv_credit,0)) as tds_amt,";
                    sql += "    sum(nvl(tds_cert_alloc_amt,0)) as cert_amt_collected,";
                    sql += "    sum(nvl(jv_debit,0))- sum(nvl(tds_cert_alloc_amt,0)) as pending_amt,";
                    sql += "    sum(case when extract(month from jvh_date) in (4, 5, 6)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q1, ";
                    sql += "    sum(case when extract(month from jvh_date) in (7, 8, 9)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q2, ";
                    sql += "    sum(case when extract(month from jvh_date) in (10, 11, 12)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q3, ";
                    sql += "    sum(case when extract(month from jvh_date) in (1, 2, 3)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q4 ";
                    sql += "  from (";
                    sql += "     select max(a.rec_branch_code) as rec_branch_code,max(jvh_date) as jvh_date,";
                    sql += " 	 max(nvl(jv_debit,0)) as jv_debit,max(nvl(jv_credit,0)) as jv_credit,  ";
                    sql += "     sum(nvl(tds_cert_alloc_amt ,0)) as tds_cert_alloc_amt";
                    sql += "     from tdspaidm a ";
                    sql += "     where a.rec_company_code = '{COMPCODE}' ";
                    if (!bCompany)
                        sql += "     and a.rec_branch_code = '{BRANCHCODE}' ";
                    sql += "     and jvh_year = {YEARCODE} ";
                    sql += "     and jvh_type <> 'OP'";
                    sql += " 	 group by a.jv_pkid";
                    sql += "    ) a group by a.rec_branch_code";
                    sql += "    order by a.rec_branch_code";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{YEARCODE}", year_code);
                    sql = sql.Replace("{BRANCHCODE}", branch_code);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    if (bCompany)
                    {
                        //total certificate amt company received.
                        sql = "select sum(tds_amt) as certamt from tdscertm where rec_company_code='" + company_code + "' and tds_year= " + year_code;
                        Object sVal = Con_Oracle.ExecuteScalar(sql);
                        cert_amt = Lib.Conv2Decimal(Lib.NumericFormat(sVal.ToString(), 2));
                    }

                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TdsosReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "black";
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mrow.tds_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                        mrow.collected_amt = Lib.Conv2Decimal(Dr["cert_amt_collected"].ToString());
                        mrow.pending_amt = Lib.Conv2Decimal(Dr["pending_amt"].ToString());
                        mrow.q1_amt = Lib.Conv2Decimal(Dr["q1"].ToString());
                        mrow.q2_amt = Lib.Conv2Decimal(Dr["q2"].ToString());
                        mrow.q3_amt = Lib.Conv2Decimal(Dr["q3"].ToString());
                        mrow.q4_amt = Lib.Conv2Decimal(Dr["q4"].ToString());
                        mList.Add(mrow);
                    }
                    if (Dt_List.Rows.Count > 0)
                    {
                        mrow = new TdsosReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "Red";
                        mrow.branch = "TOTAL";
                        mrow.tds_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(tds_amt)", "1=1").ToString(), 2));
                        mrow.collected_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(cert_amt_collected)", "1=1").ToString(), 2));
                        mrow.pending_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(pending_amt)", "1=1").ToString(), 2));
                        mrow.q1_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q1)", "1=1").ToString(), 2));
                        mrow.q2_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q2)", "1=1").ToString(), 2));
                        mrow.q3_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q3)", "1=1").ToString(), 2));
                        mrow.q4_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q4)", "1=1").ToString(), 2));
                        mList.Add(mrow);
                    }

                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintBranchWiseReport();
                    }
                }
                if (format_type == "PARTY-WISE")
                {
                    sql = " select a.party_name,a.rec_branch_code,max(tan_code) as tan_code,";
                    sql += "    sum(nvl(jv_debit,0)-nvl(jv_credit,0)) as tds_amt,";
                    sql += "    sum(nvl(tds_cert_alloc_amt,0)) as cert_amt_collected,";
                    sql += "    sum(nvl(jv_debit,0))- sum(nvl(tds_cert_alloc_amt,0)) as pending_amt,";
                    sql += "    sum(case when extract(month from jvh_date) in (4, 5, 6)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q1, ";
                    sql += "    sum(case when extract(month from jvh_date) in (7, 8, 9)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q2, ";
                    sql += "    sum(case when extract(month from jvh_date) in (10, 11, 12)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q3, ";
                    sql += "    sum(case when extract(month from jvh_date) in (1, 2, 3)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end) as q4 ";
                    sql += "  from (";
                    sql += "     select max(a.party_name) as party_name,max(a.rec_branch_code) as rec_branch_code,max(jvh_date) as jvh_date,max(tan) as tan_code,";
                    sql += " 	 max(nvl(jv_debit,0)) as jv_debit,max(nvl(jv_credit,0)) as jv_credit,  ";
                    sql += "     sum(nvl(tds_cert_alloc_amt ,0)) as tds_cert_alloc_amt";
                    sql += "     from tdspaidm a ";
                    sql += "     where a.rec_company_code = '{COMPCODE}'";
                    sql += "     and a.rec_branch_code = '{BRANCHCODE}'";
                    sql += "     and jvh_year = {YEARCODE} ";
                    sql += "     and jvh_type <> 'OP'";
                    sql += " 	 group by a.jv_pkid";
                    sql += "    ) a group by a.party_name,a.rec_branch_code";
                    sql += "    order by a.party_name,a.rec_branch_code";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRANCHCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TdsosReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "black";
                        mrow.party_name = Dr["party_name"].ToString();
                        mrow.tan_code = Dr["tan_code"].ToString();
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mrow.tds_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                        mrow.collected_amt = Lib.Conv2Decimal(Dr["cert_amt_collected"].ToString());
                        mrow.pending_amt = Lib.Conv2Decimal(Dr["pending_amt"].ToString());
                        mrow.q1_amt = Lib.Conv2Decimal(Dr["q1"].ToString());
                        mrow.q2_amt = Lib.Conv2Decimal(Dr["q2"].ToString());
                        mrow.q3_amt = Lib.Conv2Decimal(Dr["q3"].ToString());
                        mrow.q4_amt = Lib.Conv2Decimal(Dr["q4"].ToString());
                        mList.Add(mrow);
                    }
                    if (Dt_List.Rows.Count > 0)
                    {
                        mrow = new TdsosReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "Red";
                        mrow.party_name = "TOTAL";
                        mrow.tan_code = "";
                        mrow.branch = "";
                        mrow.tds_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(tds_amt)", "1=1").ToString(), 2));
                        mrow.collected_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(cert_amt_collected)", "1=1").ToString(), 2));
                        mrow.pending_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(pending_amt)", "1=1").ToString(), 2));
                        mrow.q1_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q1)", "1=1").ToString(), 2));
                        mrow.q2_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q2)", "1=1").ToString(), 2));
                        mrow.q3_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q3)", "1=1").ToString(), 2));
                        mrow.q4_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q4)", "1=1").ToString(), 2));
                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintPartyWiseReport();
                    }
                }
                if (format_type == "TDS-DETAILS")
                {
                    sql = " select a.rec_branch_code,jvh_docno,jvh_date,party_name,sman_name, tan as tan_code,tan_name, tds_cert_no,";
                    sql += "     nvl(jv_debit,0)-nvl(jv_credit,0) as tds_amt,";
                    sql += "     nvl(tds_cert_alloc_amt,0) as cert_amt_collected,";
                    sql += "     nvl(jv_debit,0)- nvl(tds_cert_alloc_amt,0) as pending_amt,";
                    sql += "     case when extract(month from jvh_date) in (4, 5, 6)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end as q1, ";
                    sql += "     case when extract(month from jvh_date) in (7, 8, 9)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end as q2, ";
                    sql += "     case when extract(month from jvh_date) in (10, 11, 12)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end as q3, ";
                    sql += "     case when extract(month from jvh_date) in (1, 2, 3)  then nvl(jv_debit,0) - nvl(tds_cert_alloc_amt,0) else 0 end as q4 ";
                    sql += "    from tdspaidm a";
                    sql += "    where a.rec_company_code = '{COMPCODE}'";
                    sql += "    and a.rec_branch_code = '{BRANCHCODE}'";
                    sql += "    and jvh_year = {YEARCODE} ";
                    sql += "    and jvh_type <> 'OP'";
                    sql += "    and a.party_name = '{PARTYNAME}'";
                    sql += "    order by a.rec_branch_code,jvh_date,jvh_type,jvh_vrno";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRANCHCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);
                    sql = sql.Replace("{PARTYNAME}", party_name);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    Dictionary<string, string> dicDocNos = new Dictionary<string, string>();
                    decimal  tot_tds_amt = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TdsosReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "black";
                        mrow.branch = Dr["rec_branch_code"].ToString();
                        mrow.party_name = Dr["party_name"].ToString();
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.sman_name = Dr["sman_name"].ToString();
                        mrow.tan_code = Dr["tan_code"].ToString();
                        mrow.tan_name = Dr["tan_name"].ToString();
                        mrow.tds_cert_no = Dr["tds_cert_no"].ToString();
                        mrow.tds_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                        mrow.collected_amt = Lib.Conv2Decimal(Dr["cert_amt_collected"].ToString());
                        mrow.pending_amt = Lib.Conv2Decimal(Dr["pending_amt"].ToString());
                        mrow.q1_amt = Lib.Conv2Decimal(Dr["q1"].ToString());
                        mrow.q2_amt = Lib.Conv2Decimal(Dr["q2"].ToString());
                        mrow.q3_amt = Lib.Conv2Decimal(Dr["q3"].ToString());
                        mrow.q4_amt = Lib.Conv2Decimal(Dr["q4"].ToString());
                        mList.Add(mrow);

                        if (!dicDocNos.ContainsKey(Dr["jvh_docno"].ToString())) //Jv will repeat (same jv allocate different cert) so to take distinct tds amt
                        {
                            dicDocNos.Add(Dr["jvh_docno"].ToString(), "");
                            tot_tds_amt += Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                        }
                    }
                    if (Dt_List.Rows.Count > 0)
                    {
                        mrow = new TdsosReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "Red";
                        mrow.party_name = "TOTAL";
                        mrow.branch = "";
                        mrow.jvh_docno = "";
                        mrow.jvh_date = "";
                        mrow.sman_name = "";
                        mrow.tan_code = "";
                        mrow.tan_name = "";
                        mrow.tds_cert_no = "";
                        //mrow.tds_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(tds_amt)", "1=1").ToString(), 2));
                        mrow.tds_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_tds_amt.ToString(), 2));
                        mrow.collected_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(cert_amt_collected)", "1=1").ToString(), 2));
                        mrow.pending_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(pending_amt)", "1=1").ToString(), 2));
                        mrow.q1_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q1)", "1=1").ToString(), 2));
                        mrow.q2_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q2)", "1=1").ToString(), 2));
                        mrow.q3_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q3)", "1=1").ToString(), 2));
                        mrow.q4_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(q4)", "1=1").ToString(), 2));
                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintTdsDetReport();
                    }
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
            RetData.Add("cert_amt", cert_amt);
            return RetData;
        }

        private void PrintPartyWiseReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "TdsPartyReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 40;
                WS.Columns[2].Width = 256 * 12;
                WS.Columns[3].Width = 256 * 12;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 12;
                WS.Columns[6].Width = 256 * 12;
                WS.Columns[7].Width = 256 * 12;
                WS.Columns[8].Width = 256 * 12;
                WS.Columns[9].Width = 256 * 12;
                WS.Columns[10].Width = 256 * 12;
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
                Lib.WriteData(WS, iRow, 1, "TDS PARTY WISE OS REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-PAID", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ALLOCATED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PENDING", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q1.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q2.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q3.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q4.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (TdsosReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.party_name, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.collected_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q1_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q2_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q3_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q4_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }else
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.collected_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q1_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q2_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q3_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q4_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void PrintTdsDetReport()
        {

            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "TdsDetReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 12;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 10;
                WS.Columns[4].Width = 256 * 45;
                WS.Columns[5].Width = 256 * 20;
                WS.Columns[6].Width = 256 * 12;
                WS.Columns[7].Width = 256 * 12;
                WS.Columns[8].Width = 256 * 12;
                WS.Columns[9].Width = 256 * 12;
                WS.Columns[10].Width = 256 * 10;
                WS.Columns[11].Width = 256 * 10;
                WS.Columns[12].Width = 256 * 10;
                WS.Columns[13].Width = 256 * 12;
                WS.Columns[14].Width = 256 * 19;
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
                Lib.WriteData(WS, iRow, 1, "TDS DETAIL OS REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-PAID", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ALLOCATED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PENDING", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q1.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q2.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q3.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q4.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN.CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN.NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CERTIFICATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                foreach (TdsosReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.party_name, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.collected_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q1_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q2_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q3_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q4_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                    }
                    else
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_docno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.collected_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q1_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q2_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q3_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q4_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_cert_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }
                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        private void PrintBranchWiseReport()
        {

            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
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
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "TdsBrReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 20;
                WS.Columns[2].Width = 256 * 12;
                WS.Columns[3].Width = 256 * 12;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 12;
                WS.Columns[6].Width = 256 * 12;
                WS.Columns[7].Width = 256 * 12;
             
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
                Lib.WriteData(WS, iRow, 1, "TDS BRANCH WISE OS REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-PAID", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ALLOCATED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PENDING", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q1.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q2.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q3.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Q4.BAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
               
                foreach (TdsosReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.collected_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q1_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q2_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q3_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q4_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    else
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.collected_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q1_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q2_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q3_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.q4_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                }
                iRow++;

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
