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
    public class AcTransReportService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<AcTransReport> mList = new List<AcTransReport>();
        AcTransReport mrow;
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

        string ErrorMessage = "";

        string hide_ho_entries = "N";

        string sWhere = "";

        string type_date = "";
        string from_date = "";
        string to_date = "";
        string category = "";
        string vrnos = "0";
      
        long page_count = 0;
        long page_current = 0;
        long page_rows = 0;
        long page_rowcount = 0;

        long startrow = 0;
        long endrow = 0;

        Boolean all = false;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<AcTransReport>();
            ErrorMessage = "";
            try
            {
                DataTable Dt_Temp = new DataTable();
                Con_Oracle = new DBConnection();

                page_count = 0;
                page_current = 0;
                page_rows = 0;
                page_rowcount = 0;

                type = SearchData["type"].ToString();
                rec_category = SearchData["rec_category"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
                type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();
                category = SearchData["category"].ToString();
                hide_ho_entries = SearchData["hide_ho_entries"].ToString();
                vrnos = SearchData["vrnos"].ToString().Replace(" ", "");

                //all = (Boolean)SearchData["all"];

                page_count = (long)SearchData["page_count"];
                page_current = (long)SearchData["page_current"];
                page_rows = (long)SearchData["page_rows"];
                page_rowcount = (long)SearchData["page_rowcount"];

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);

                startrow = 0;
                endrow = 0;
                sWhere = "";

                sWhere = " where  a.rec_company_code = '{COMPCODE}' and a.rec_branch_code = '{BRCODE}' ";
                if (from_date != "NULL" && to_date != "NULL")
                {
                    if (type_date == "jvh_date")
                        sWhere += " and a.{TYPEDATE} between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                    else
                        sWhere += " and to_char(a.{TYPEDATE},'DD-MON-YYYY')  between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";
                }
                if (category != "ALL")
                    sWhere += " and a.jvh_type = '{TYPE}'";

                if (hide_ho_entries == "Y")
                    sWhere += "  and jvh_type not in('HO','IN-ES') ";

                if (vrnos != "")
                {
                    if (vrnos.Contains("-"))
                    {
                        string[] sData = vrnos.Split('-');
                        string vrnos_from = sData[0];
                        string vrnos_to = sData[1];

                        sWhere += " and a.jvh_vrno between " + Lib.Conv2Integer(vrnos_from).ToString() + " and " + Lib.Conv2Integer(vrnos_to).ToString();
                    }
                    else
                        sWhere += " and a.jvh_vrno in (" + vrnos + ")";
                }

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM ledgerh a ";
                    sql += " inner join ledgert b on jvh_pkid = b.jv_parent_id";
                    sql += " left join acctm c on jv_acc_id  = c.acc_pkid";
                    sql += sWhere;

                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{TYPEDATE}", type_date);
                    sql = sql.Replace("{TYPE}", category);

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



                sql = "";
                sql += " select * from ( ";
                sql += " select jvh_vrno, jvh_type,jvh_date,";
                sql += " acc_code, acc_name, jv_debit, jv_credit, ";
                sql += " a.rec_created_by,a.rec_created_date,a.rec_edited_by,a.rec_edited_date, jv_ctr,";
                sql += "  row_number() over(order by a.{TYPEDATE},jvh_type, jvh_vrno, jv_ctr) rn ";
                sql += " from ledgerh a ";
                sql += " inner join ledgert b on jvh_pkid = b.jv_parent_id";
                sql += " left join acctm c on jv_acc_id  = c.acc_pkid";
                sql += sWhere;
                sql += ") a ";

                if (type != "EXCEL")
                {
                    sql += " where rn between {startrow} and {endrow}";
                }

                sql += " order by a.{TYPEDATE}, jvh_type, jvh_vrno, jv_ctr";


                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);
                sql = sql.Replace("{TYPEDATE}", type_date);
                sql = sql.Replace("{TYPE}", category);
                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());


                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();


                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new AcTransReport();
                    mrow.jvh_vrno = Dr["jvh_vrno"].ToString();
                    mrow.jvh_type = Dr["jvh_type"].ToString();
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.acc_code = Dr["acc_code"].ToString();
                    mrow.acc_name = Dr["acc_name"].ToString();
                    mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                    mrow.rec_createdby = Dr["rec_created_by"].ToString();
                    mrow.rec_createddate = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mrow.rec_editedby = Dr["rec_edited_by"].ToString();
                    mrow.rec_editeddate = Lib.DatetoStringDisplayformat(Dr["rec_edited_date"]);

                    mList.Add(mrow);

                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintAcTransReport();
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
            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);
            return RetData;
        }

        private void PrintAcTransReport()
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

                File_Display_Name = "TranscationReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 7;
                WS.Columns[2].Width = 256 * 5;
                WS.Columns[3].Width = 256 * 14;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 31;
                WS.Columns[6].Width = 256 * 12;
                WS.Columns[7].Width = 256 * 11;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 14;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 14;
                WS.Columns[12].Width = 256 * 15;


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
                Lib.WriteData(WS, iRow, 1, "TRANSCATION REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 11;
                iCol = 1;


                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED-BY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREATED-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EDITED-BY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EDITED-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);


                _Size = 10;

                foreach (AcTransReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;

                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_vrno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.acc_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_debit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.rec_createdby, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.rec_createddate, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.rec_editedby, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.rec_editeddate, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                }

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }



        public IDictionary<string, object> TransDetList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<AcTransReport>();

            List<AcTransReport> xrefList = new List<AcTransReport>();

            ErrorMessage = "";
            string narration = "";
            string JVH_PKID = "";
            try
            {
                DataTable Dt_Temp = new DataTable();
                Con_Oracle = new DBConnection();


                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["jvh_year"].ToString();
                type = SearchData["jvh_type"].ToString();
                string vrno = SearchData["jvh_vrno"].ToString();
                string acc_code = SearchData["acc_code"].ToString();
                string str = "";

                decimal nDr = 0;
                decimal nCr = 0;

                sql = "";
                
                sql += " select jvh_pkid,jvh_vrno, jvh_type,jvh_date,";
                sql += " acc_code, acc_name, jv_debit, jv_credit, jvh_narration,";
                sql += " a.rec_created_by,a.rec_created_date,a.rec_edited_by,a.rec_edited_date, jv_ctr";
                sql += " from ledgerh a ";
                sql += " inner join ledgert b on jvh_pkid = b.jv_parent_id";
                sql += " left join acctm c on jv_acc_id  = c.acc_pkid";
                sql += " where  a.rec_company_code = '{COMPCODE}' and a.rec_branch_code = '{BRCODE}' ";
                sql += " and jvh_year = {YEAR} and jvh_type ='{TYPE}' and jvh_vrno  = {VRNO}";
                sql += " order by jv_ctr";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEAR}", year_code);
                sql = sql.Replace("{TYPE}", type);
                sql = sql.Replace("{VRNO}", vrno);

                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();


                foreach (DataRow Dr in Dt_List.Rows)
                {
                    JVH_PKID= Dr["jvh_pkid"].ToString();

                    mrow = new AcTransReport();
                    mrow.jvh_vrno = Dr["jvh_vrno"].ToString();
                    mrow.jvh_type = Dr["jvh_type"].ToString();
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.acc_code = Dr["acc_code"].ToString();
                    mrow.acc_name = Dr["acc_name"].ToString();
                    mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                    nDr += Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    nCr += Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                    narration = Dr["jvh_narration"].ToString();

                    str = "";
                    if (!Dr["rec_created_by"].Equals(DBNull.Value))
                    {
                        str += Dr["rec_created_by"].ToString();
                        str += "-" + Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    }
                    if (!Dr["rec_edited_by"].Equals(DBNull.Value))
                    {
                        str += " / " + Dr["rec_edited_by"].ToString();
                        str += "-" + Lib.DatetoStringDisplayformat(Dr["rec_edited_date"]);
                    }
                    mrow.rec_createdby = str;

                    //mrow.rec_createdby = Dr["rec_created_by"].ToString();
                    //mrow.rec_createddate = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    //mrow.rec_editedby = Dr["rec_edited_by"].ToString();
                    //mrow.rec_editeddate = Lib.DatetoStringDisplayformat(Dr["rec_edited_date"]);

                    mrow.rowcolor = "";
                    if ( Dr["acc_code"].ToString().StartsWith(acc_code) )
                        mrow.rowcolor = "RED";

                    mList.Add(mrow);
                }

                if ( nDr != 0 || nCr != 0)
                {
                    mrow = new AcTransReport();
                    mrow.acc_name = "TOTAL";
                    mrow.jv_debit = nDr ;
                    mrow.jv_credit = nCr;
                    mList.Add(mrow);
                }
                Dt_List.Rows.Clear();


                if(JVH_PKID != "")
                {
                    sql = " select x.xref_pkid, b.jv_ctr as slno,  ac.acc_code, ac.acc_name, a.jvh_docno ,b.jv_debit, b.jv_credit,h.jvh_docno as xref_no,h.jvh_date as xref_date, xref_amt ";
                    sql += " 	from ledgerh a ";
                    sql += " 	inner join ledgert b on a.jvh_pkid = b.jv_parent_id";
                    sql += " 	inner join acctm ac on jv_acc_id = acc_pkid";
                    sql += " 	inner join ledgerxref x on b.jv_pkid =  xref_dr_jv_id";
                    sql += " 	inner join ledgert d on x.xref_cr_jv_id = d.jv_pkid";
                    sql += " 	inner join ledgerh h on d.jv_parent_id = h.jvh_pkid";
                    sql += " 	where a.jvh_pkid ='{JVHID}'";
                    sql += " 	";
                    sql += " 	union all";
                    sql += " 	";
                    sql += " 	select x.xref_pkid, b.jv_ctr as slno, ac.acc_code, ac.acc_name, a.jvh_docno ,b.jv_debit, b.jv_credit,h.jvh_docno as xref_no,h.jvh_date as xref_date, xref_amt";
                    sql += " 	from ledgerh a ";
                    sql += " 	inner join ledgert b on a.jvh_pkid = b.jv_parent_id";
                    sql += " 	inner join acctm ac on jv_acc_id = acc_pkid";
                    sql += " 	inner join ledgerxref x on b.jv_pkid =  xref_cr_jv_id";
                    sql += " 	inner join ledgert d on x.xref_dr_jv_id = d.jv_pkid";
                    sql += " 	inner join ledgerh h on d.jv_parent_id = h.jvh_pkid";
                    sql += " 	where a.jvh_pkid ='{JVHID}'";
                    sql += " 	order by slno";

                    sql = sql.Replace("{JVHID}", JVH_PKID);

                    Con_Oracle = new DBConnection();
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    decimal xref_tot = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {

                        mrow = new AcTransReport();
                        mrow.rowcolor = "Black";
                        mrow.jvh_pkid = JVH_PKID;
                        mrow.xref_pkid = Dr["xref_pkid"].ToString();
                        mrow.slno= Lib.Conv2Integer(Dr["slno"].ToString());
                        mrow.acc_code = Dr["acc_code"].ToString();
                        mrow.acc_name = Dr["acc_name"].ToString();
                        mrow.jvh_docno = Dr["jvh_docno"].ToString();
                        mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                        mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.xref_no  = Dr["xref_no"].ToString();
                        mrow.xref_date = Lib.DatetoStringDisplayformat(Dr["xref_date"]);
                        mrow.xref_amt = Lib.Conv2Decimal(Dr["xref_amt"].ToString());

                        xref_tot += Lib.Conv2Decimal(Dr["xref_amt"].ToString());
                        xrefList.Add(mrow);
                    }

                    if (xref_tot != 0)
                    {
                        mrow = new AcTransReport();
                        mrow.jvh_pkid = JVH_PKID;
                        mrow.rowcolor = "Red";
                        mrow.acc_name = "TOTAL";
                        mrow.xref_amt = xref_tot;
                        xrefList.Add(mrow);
                    }

                    Dt_List.Rows.Clear();
                }

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("list", mList);
            RetData.Add("narration", narration);
            RetData.Add("xreflist", xrefList);
            return RetData;
        }

        public Dictionary<string, object> DeleteRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string id = SearchData["pkid"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string user_code = SearchData["user_code"].ToString();
            string user_pkid = SearchData["user_pkid"].ToString();
            string type = SearchData["type"].ToString();
            try
            {
                Con_Oracle = new DBConnection();

                
                if (type == "ROW-DELETE")
                    sql = " delete from ledgerxref where  xref_pkid = '" + id + "'";
                else
                    sql = " delete from ledgerxref where  xref_dr_jvh_id = '" + id + "' or xref_cr_jvh_id ='" + id + "'";

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();


                Con_Oracle.CloseConnection();

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
            return RetData;
        }

    }


}
