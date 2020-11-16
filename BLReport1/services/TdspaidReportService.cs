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
    public class TdspaidReportService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<TdspaidReport> mList = new List<TdspaidReport>();
        TdspaidReport mrow;
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

 
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<TdspaidReport>();
            ErrorMessage = "";
            DataRow Dr_Target = null;
            try
            {
                type = SearchData["type"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
                format_type = SearchData["format_type"].ToString();
                
                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (format_type == "TDS-PAID-DETAILS")
                {
   
                    sql = " select a.rec_branch_code,jv_pkid,  jvh_vrno as jv_vrno, a.jvh_type as jv_type ,jvh_date as jv_date, ";
                    sql += "  party_code,party_name,";
                    sql += "  tan as tan_code,tan_name,";
                    sql += "  tds_cert_no,tds_cert_qtr,tds_cert_brcode as cert_recvd_at,tds_cert_amt as cert_amt,";
                    sql += "  jv_gross_bill_amt as gross_bill_amt,tds_cert_gross as gross_cert_amt,";
                    sql += "  jv_credit,jv_debit as tds_amt,tds_cert_alloc_amt as cert_alloc_amt,  nvl(jv_debit,0)- nvl(tds_cert_alloc_amt,0) as pending_amt ";
                    sql += "  from tdspaidm a";
                    sql += "   where a.rec_company_code = '{COMPCODE}'";
                    sql += "   and jvh_year = {YEARCODE} ";
                    sql += "   and jvh_type <> 'OP'";
                    if (type == "SUMMARY-DETAIL")
                    {
                        if (PKID.Trim() == "")
                            sql += " and jv_tan_id is null ";
                        else
                            sql += " and jv_tan_id = '" + PKID + "'";
                    }
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(party_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(tan_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "   order by tan_code, a.rec_branch_code,a.jvh_vrno, a.jvh_date,a.jvh_type";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TdspaidReport();
                        mrow.row_type = "DETAIL";
                        mrow.row_colour = "blue";
                        mrow.branch_code = Dr["rec_branch_code"].ToString();
                        mrow.jv_pkid = Dr["jv_pkid"].ToString();
                        mrow.jv_vrno = Dr["jv_vrno"].ToString();
                        mrow.jv_type = Dr["jv_type"].ToString();
                        if (!Dr["jv_date"].Equals(DBNull.Value))
                            mrow.jv_date = ((DateTime)Dr["jv_date"]).ToString("dd-MMM-yyyy");
                        else
                            mrow.jv_date = "";
                        mrow.party_code = Dr["party_code"].ToString();
                        mrow.party_name = Dr["party_name"].ToString();
                        mrow.tan_code = Dr["tan_code"].ToString();
                        mrow.tan_name = Dr["tan_name"].ToString();
                        mrow.tds_cert_no = Dr["tds_cert_no"].ToString();
                        mrow.tds_cert_qtr = Dr["tds_cert_qtr"].ToString();
                        mrow.cert_recvd_at = Dr["cert_recvd_at"].ToString();
                        mrow.gross_bill_amt = Lib.Conv2Decimal(Dr["gross_bill_amt"].ToString());
                        mrow.gross_cert_amt = Lib.Conv2Decimal(Dr["gross_cert_amt"].ToString());
                        mrow.tds_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                        mrow.cert_amt = Lib.Conv2Decimal(Dr["cert_amt"].ToString());
                        mrow.cert_alloc_amt = Lib.Conv2Decimal(Dr["cert_alloc_amt"].ToString());
                        mrow.pending_amt = Lib.Conv2Decimal(Dr["pending_amt"].ToString());
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mList.Add(mrow);
                    }
                    if(type == "SUMMARY-DETAIL" && Dt_List.Rows.Count>0)
                    {
                        mrow = new TdspaidReport();
                        mrow.row_type = "TOTAL";
                        mrow.row_colour = "Red";
                        mrow.branch_code ="";
                        mrow.jv_pkid = "";
                        mrow.jv_vrno = "";
                        mrow.jv_type = "";
                        mrow.jv_date = "";
                        mrow.party_code = "";
                        mrow.party_name = "";
                        mrow.tan_code = "";
                        mrow.tan_name = "";
                        mrow.tds_cert_no = "";
                        mrow.tds_cert_qtr = "";
                        mrow.cert_recvd_at = "TOTAL";
                        mrow.gross_bill_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(gross_bill_amt)", "1=1").ToString(),2));
                        mrow.gross_cert_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(gross_cert_amt)", "1=1").ToString(), 2));
                        mrow.tds_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(tds_amt)", "1=1").ToString(), 2));
                        mrow.cert_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(cert_amt)", "1=1").ToString(), 2));
                        mrow.cert_alloc_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(cert_alloc_amt)", "1=1").ToString(), 2));
                        mrow.pending_amt = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(pending_amt)", "1=1").ToString(), 2));
                        mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(Dt_List.Compute("sum(jv_credit)", "1=1").ToString(), 2));
                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintPaidDetailReport();
                    }
                }
                if (format_type == "TDS-PAID-SUMMARY")
                {
                    sql = " select a.*,c.tds_amt as cert_amt,'DETAIL' as row_type,'Black' as row_colour from (";

                    sql += " select jv_tan_id,";
                    sql += "    max(tan_code) as tan_code,";
                    sql += "    max(tan_name) as tan_name, ";
                    sql += "    sum(nvl(jv_credit,0)) as jv_credit, ";
                    sql += "    sum(nvl(jv_debit,0)) as tds_amt,sum(nvl(tds_cert_alloc_amt,0)) as cert_alloc_amt,   sum(nvl(jv_debit,0))- sum(nvl(tds_cert_alloc_amt,0)) as pending_amt";
                    sql += "    from (";

                    sql += "      select max(a.jv_tan_id) as jv_tan_id,max(tan) as tan_code,max(tan_name) as tan_name,";
                    sql += "  	 max(nvl(jv_debit,0)) as jv_debit,max(nvl(jv_credit,0)) as jv_credit,  ";
                    sql += "      sum(nvl(tds_cert_alloc_amt ,0)) as tds_cert_alloc_amt";
                    sql += "      from tdspaidm a ";
                    sql += "   where a.rec_company_code = '{COMPCODE}'";
                    sql += "   and a.rec_branch_code != 'KOLAF' ";
                    sql += "   and jvh_year = {YEARCODE} ";
                    sql += "   and jvh_type <> 'OP' ";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(tan) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(tan_name) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += "  	 group by a.jv_pkid";

                    sql += "    ) a";

                    sql += "   group by jv_tan_id";

                    sql += " )a";
                    sql += "   left join (";
                    sql += "   select tds_tan_id, sum(tds_amt) as tds_amt from tdscertm a";
                    sql += "   where a.rec_company_code = '{COMPCODE}'";
                    sql += "   and tds_year = {YEARCODE} ";
                    sql += "   group by tds_tan_id";
                    sql += "   )c on a.jv_tan_id = c.tds_tan_id";
                    sql += "   order by tan_code ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    if (Dt_List.Rows.Count > 0)
                    {
                        Dr_Target = Dt_List.NewRow();
                        Dr_Target["row_type"] = "TOTAL";
                        Dr_Target["row_colour"] = "Red";
                        Dr_Target["tan_name"] = "TOTAL";
                        Dr_Target["jv_credit"] = Lib.NumericFormat(Dt_List.Compute("sum(jv_credit)", "1=1").ToString(), 2);
                        Dr_Target["tds_amt"] = Lib.NumericFormat(Dt_List.Compute("sum(tds_amt)", "1=1").ToString(), 2);
                        Dr_Target["cert_alloc_amt"] = Lib.NumericFormat(Dt_List.Compute("sum(cert_alloc_amt)", "1=1").ToString(), 2);
                        Dr_Target["cert_amt"] = Lib.NumericFormat(Dt_List.Compute("sum(cert_amt)", "1=1").ToString(), 2);
                        Dr_Target["pending_amt"] = Lib.NumericFormat(Dt_List.Compute("sum(pending_amt)", "1=1").ToString(), 2);
                        Dt_List.Rows.Add(Dr_Target);
                    }
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TdspaidReport();
                        mrow.row_type = Dr["row_type"].ToString();
                        mrow.row_colour = Dr["row_colour"].ToString();
                        mrow.tan_id = Dr["jv_tan_id"].ToString();
                        mrow.tan_code = Dr["tan_code"].ToString();
                        mrow.tan_name = Dr["tan_name"].ToString();
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.tds_amt = Lib.Conv2Decimal(Dr["tds_amt"].ToString());
                        mrow.cert_alloc_amt = Lib.Conv2Decimal(Dr["cert_alloc_amt"].ToString()); 
                        mrow.cert_amt = Lib.Conv2Decimal(Dr["cert_amt"].ToString());
                        mrow.pending_amt = Lib.Conv2Decimal(Dr["pending_amt"].ToString());
                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            PrintPaidSummaryReport();
                    }
                }
                if (format_type == "26AS-MASTERS")
                {

                    sql = " select ASM_SLNO,ASM_TAN,ASM_TAN_NAME,ASM_GROSS,ASM_DEDUCTED,ASM_TDS,'DETAIL' as row_type,'Black' as row_colour ";
                    sql += " from TDS26ASM a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and asm_year = {YEARCODE}";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(ASM_TAN) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(ASM_TAN_NAME) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += " order by asm_slno,asm_tan ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    if (Dt_List.Rows.Count > 0)
                    {
                        Dr_Target = Dt_List.NewRow();
                        Dr_Target["row_type"] = "TOTAL";
                        Dr_Target["row_colour"] = "Red";
                        Dr_Target["ASM_TAN_NAME"] = "TOTAL";
                        Dr_Target["ASM_GROSS"] = Lib.NumericFormat(Dt_List.Compute("sum(ASM_GROSS)", "1=1").ToString(), 2);
                        Dr_Target["ASM_DEDUCTED"] = Lib.NumericFormat(Dt_List.Compute("sum(ASM_DEDUCTED)", "1=1").ToString(), 2);
                        Dr_Target["ASM_TDS"] = Lib.NumericFormat(Dt_List.Compute("sum(ASM_TDS)", "1=1").ToString(), 2);
                        Dt_List.Rows.Add(Dr_Target);
                    }

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TdspaidReport();
                        mrow.row_type = Dr["row_type"].ToString();
                        mrow.row_colour = Dr["row_colour"].ToString();
                        mrow.tan_id = "";
                        mrow.tan_code = Dr["ASM_TAN"].ToString();
                        mrow.tan_name = Dr["ASM_TAN_NAME"].ToString();
                        mrow.asm_gross = Lib.Conv2Decimal(Dr["ASM_GROSS"].ToString());
                        mrow.asm_deducted = Lib.Conv2Decimal(Dr["ASM_DEDUCTED"].ToString());
                        mrow.asm_tds = Lib.Conv2Decimal(Dr["ASM_TDS"].ToString());
                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            Print26asMasterReport();
                    }
                }
                if (format_type == "26AS-DETAILS")
                {

                    sql = " select ASM_SLNO,ASM_TAN,ASM_TAN_NAME,ASD_SECTION,ASD_TRANS_DATE,ASD_BOOK_DATE ,ASD_GROSS,ASD_DEDUCTED ,ASD_TDS,'DETAIL' as row_type,'Black' as row_colour ";
                    sql += " from TDS26ASM a ";
                    sql += " inner join TDS26ASD b on a.ASM_PKID = b.ASD_PARENT_ID";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.asm_year = {YEARCODE}";
                    if (searchstring != "")
                    {
                        sql += " and (";
                        sql += "  upper(ASM_TAN) like '%" + searchstring.ToUpper() + "%'";
                        sql += " or ";
                        sql += "  upper(ASM_TAN_NAME) like '%" + searchstring.ToUpper() + "%'";
                        sql += " )";
                    }
                    sql += " order by asm_slno,asm_tan";
                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                  
                    if (Dt_List.Rows.Count > 0)
                    {
                        Dr_Target = Dt_List.NewRow();
                        Dr_Target["row_type"] = "TOTAL";
                        Dr_Target["row_colour"] = "Red";
                        Dr_Target["ASM_TAN_NAME"] = "TOTAL";
                        Dr_Target["ASD_GROSS"] = Lib.NumericFormat(Dt_List.Compute("sum(ASD_GROSS)", "1=1").ToString(), 2);
                        Dr_Target["ASD_DEDUCTED"] = Lib.NumericFormat(Dt_List.Compute("sum(ASD_DEDUCTED)", "1=1").ToString(), 2);
                        Dr_Target["ASD_TDS"] = Lib.NumericFormat(Dt_List.Compute("sum(ASD_TDS)", "1=1").ToString(), 2);
                        Dt_List.Rows.Add(Dr_Target);
                    }
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new TdspaidReport();
                        mrow.row_type = Dr["row_type"].ToString();
                        mrow.row_colour = Dr["row_colour"].ToString();
                        mrow.tan_id = "";
                        mrow.tan_code = Dr["ASM_TAN"].ToString();
                        mrow.tan_name = Dr["ASM_TAN_NAME"].ToString();
                        mrow.asd_section = Dr["ASD_SECTION"].ToString();
                        mrow.asd_trans_date = Lib.DatetoStringDisplayformat(Dr["ASD_TRANS_DATE"]);
                        mrow.asd_book_date = Lib.DatetoStringDisplayformat(Dr["ASD_BOOK_DATE"]);
                        mrow.asd_gross = Lib.Conv2Decimal(Dr["ASD_GROSS"].ToString());
                        mrow.asd_deducted = Lib.Conv2Decimal(Dr["ASD_DEDUCTED"].ToString());
                        mrow.asd_tds = Lib.Conv2Decimal(Dr["ASD_TDS"].ToString());

                        mList.Add(mrow);
                    }
                    if (type == "EXCEL")
                    {
                        if (mList != null)
                            Print26asDetailReport();
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
            return RetData;
        }

        private void PrintPaidDetailReport()
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


                File_Display_Name = "TdsPaidReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 10;
                WS.Columns[2].Width = 256 * 10;
                WS.Columns[3].Width = 256 * 12;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 12;
                WS.Columns[6].Width = 256 * 40;
                WS.Columns[7].Width = 256 * 12;
                WS.Columns[8].Width = 256 * 40;
                WS.Columns[9].Width = 256 * 12;
                WS.Columns[10].Width = 256 * 8;
                WS.Columns[11].Width = 256 * 13;
                WS.Columns[12].Width = 256 * 14;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 12;
                WS.Columns[15].Width = 256 * 12;
                WS.Columns[16].Width = 256 * 12;
                WS.Columns[17].Width = 256 * 12;
                WS.Columns[18].Width = 256 * 12;

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
                Lib.WriteData(WS, iRow, 1, "TDS PAID DETAIL REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTY-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTY-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CERT-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CERT-QTR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CERT-RECVD-AT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GROSS-BILL-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GROSS-CERT-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CERT-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ALLOCATED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PENDING", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                

                foreach (TdspaidReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    Lib.WriteData(WS, iRow, iCol++, Rec.branch_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_vrno, _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.party_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.tds_cert_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.tds_cert_qtr, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cert_recvd_at, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.gross_bill_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.gross_cert_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cert_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cert_alloc_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
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

        private void PrintPaidSummaryReport()
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



                File_Display_Name = "TdsPaidReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 12;
                WS.Columns[2].Width = 256 * 40;
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
                Lib.WriteData(WS, iRow, 1, "TDS PAID SUMMARY REPORT - TAN WISE ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
               
                Lib.WriteData(WS, iRow, iCol++, "TAN-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-DR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-CR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CERT-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ALLOCATED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PENDING", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
               
                foreach (TdspaidReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cert_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cert_alloc_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    else
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tds_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cert_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cert_alloc_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.pending_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
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
        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            //Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "SALES EXECUTIVE");
            //RetData.Add("smanlist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "CITY");
            //RetData.Add("citylist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "STATE");
            //RetData.Add("statelist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;
        }

        private void Print26asMasterReport()
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


                File_Display_Name = "TdsPaidReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 12;
                WS.Columns[2].Width = 256 * 40;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                
                


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
                Lib.WriteData(WS, iRow, 1, "TDS26AS - MASTER", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "TAN-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAID-CREDITED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAX-DEDUCTED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-DEPOSITED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (TdspaidReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asm_gross, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asm_deducted, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asm_tds, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }else
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asm_gross, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asm_deducted, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asm_tds, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
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
        private void Print26asDetailReport()
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


                File_Display_Name = "TdsPaidReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 12;
                WS.Columns[2].Width = 256 * 40;
                WS.Columns[3].Width = 256 * 12;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 12;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;


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
                Lib.WriteData(WS, iRow, 1, "TDS26AS - MASTER", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "TAN-CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN-NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SECTION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TRANSACTION-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BOOKING-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAID-CREDITED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAX-DEDUCTED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TDS-DEPOSITED", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (TdspaidReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_section, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_trans_date, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_book_date, _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_gross, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_deducted, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_tds, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    else
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_section, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_trans_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_book_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_gross, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_deducted, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.asd_tds, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
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
