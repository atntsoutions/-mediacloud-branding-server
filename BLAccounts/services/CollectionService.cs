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
    public class CollectionService : BL_Base
    {



        DataTable Dt_List = new DataTable();

        ExcelFile WB;
        ExcelWorksheet WS = null;

        int iRow = 0;
        int iCol = 0;


        string type = "";
        string report_folder = "";
        string File_Name = "";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string to_date = "";
        string from_date = "";
        string ACC_ID = "";


        List<CollectionReport> mList = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sql = "";
            //string SID = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            mList = new List<CollectionReport>();

            CollectionReport mrow;

            type = SearchData["type"].ToString();

            report_folder = SearchData["report_folder"].ToString();
            PKID = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            string fdate = SearchData["from_date"].ToString();
            from_date = Lib.StringToDate(fdate).ToUpper();

            string edate = SearchData["to_date"].ToString();
            to_date = Lib.StringToDate(edate).ToUpper();

            ACC_ID = SearchData["acc_id"].ToString();

            string searchstring = SearchData["searchstring"].ToString().ToUpper();



            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine(report_folder, PKID);


            string fy_start = "31-MAR-" + year_code;
            DateTime dt = DateTime.Parse(fdate);

            dt = new DateTime(dt.Year, dt.Month, 1 );

            DateTime p3_last = dt.AddDays(-1);
            DateTime p3_first = new DateTime(p3_last.Year, p3_last.Month, 1);

            DateTime p2_last = p3_first.AddDays(-1);
            DateTime p2_first = new DateTime(p2_last.Year, p2_last.Month, 1);

            DateTime p1_last = p2_first.AddDays(-1);
            DateTime p1_first = new DateTime(p1_last.Year, p1_last.Month, 1);


            try
            {

                sql += " select  a.* from( ";
                sql += " select 'A' as slno, branch_code,  ";
                sql += " sum(case when jvh_date <= '{31-MAR-2019}' then bal else 0 end) as col1, ";
                sql += " sum(case when jvh_date between '{01-APR-2019}' and '{30-APR-2019}' then bal else 0 end) as col2, ";
                sql += " sum(case when jvh_date between '{01-MAY-2019}' and '{31-MAY-2019}' then bal else 0 end) as col3, ";
                sql += " sum(case when jvh_date between '{01-JUN-2019}' and '{30-JUN-2019}' then bal else 0 end) as col4 ";
                sql += " from( ";
                sql += " select h.rec_branch_code as branch_code, jvh_date, jv_debit - nvl(sum(xref_amt), 0) as bal ";
                sql += " from ledgerh h ";
                sql += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id) ";
                sql += " inner join Acctm a on (L.jv_acc_id = A.acc_pkid) ";
                sql += " inner join acgroupm g on a.acc_group_id = g.acgrp_pkid and g.rec_company_code = 'CPL'  and acgrp_name = 'SUNDRY DEBTORS' ";
                sql += " left  join ledgerxref X on(L.jv_pkid = X.xref_dr_jv_id   and X.XREF_CR_JV_DATE < '{01-JUL-2019}') ";
                sql += " left  join param s on(jv_acc_id = param_pkid) ";
                sql += " where ";
                //sql += " jv_acc_id IN('76F20A3E-7CFB-4F89-AF30-826911AB2985', '5958DA86-6C20-C2B6-B7B3-38A8C90C974B') and ";
                sql += " h.rec_company_code = 'CPL'  and jvh_date < '{01-JUL-2019}'  and L.jv_debit > 0 and h.rec_deleted = 'N' and acc_against_invoice = 'D' ";
                sql += " and jvh_type not in ('OP', 'OB', 'OC') ";
                sql += " group by h.rec_branch_code, jv_pkid, jvh_date, jv_debit ";
                sql += " ) a group by branch_code ";

                sql += " union all ";

                sql += " select 'B' as slno, branch_code,  ";
                sql += " sum(case when jvh_date <= '{31-MAR-2019}' then bal else 0 end) as col1, ";
                sql += " sum(case when jvh_date between '{01-APR-2019}' and '{30-APR-2019}' then bal else 0 end) as col2, ";
                sql += " sum(case when jvh_date between '{01-MAY-2019}' and '{31-MAY-2019}' then bal else 0 end) as col3, ";
                sql += " sum(case when jvh_date between '{01-JUN-2019}' and '{30-JUN-2019}' then bal else 0 end) as col4 ";

                sql += " from( ";
                sql += " select h.rec_branch_code as branch_code, jvh_date, jv_debit - nvl(sum(xref_amt), 0) as bal ";
                sql += " from ledgerh h ";
                sql += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id) ";

                sql += " inner join Acctm a on (L.jv_acc_id = A.acc_pkid) ";

                sql += " inner join acgroupm g on a.acc_group_id = g.acgrp_pkid and g.rec_company_code = 'CPL'  and acgrp_name = 'SUNDRY DEBTORS' ";
                sql += " left  join ledgerxref X on(L.jv_pkid = X.xref_dr_jv_id   and X.XREF_CR_JV_DATE <= '{31-JUL-2019}') ";
                sql += " left  join param s on(jv_acc_id = param_pkid) ";

                sql += " where ";
                //sql += " jv_acc_id IN('76F20A3E-7CFB-4F89-AF30-826911AB2985', '5958DA86-6C20-C2B6-B7B3-38A8C90C974B') and ";
                sql += " h.rec_company_code = 'CPL'  and jvh_date < '{01-JUL-2019}'  and L.jv_debit > 0 and h.rec_deleted = 'N' and acc_against_invoice = 'D' ";
                sql += " and jvh_type not in ('OP', 'OB', 'OC') ";
                sql += " group by h.rec_branch_code, jv_pkid, jvh_date, jv_debit ";
                sql += " ) a group by branch_code ";

                sql += " ) a order by branch_code, slno ";


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{EDATE}", to_date);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{PKID}", ACC_ID);

                sql = sql.Replace("{31-MAR-2019}", fy_start);
                sql = sql.Replace("{01-APR-2019}", p1_first.ToString("dd-MMM-yyyy"));
                sql = sql.Replace("{30-APR-2019}", p1_last.ToString("dd-MMM-yyyy"));

                sql = sql.Replace("{01-MAY-2019}", p2_first.ToString("dd-MMM-yyyy"));
                sql = sql.Replace("{31-MAY-2019}", p2_last.ToString("dd-MMM-yyyy"));

                sql = sql.Replace("{01-JUN-2019}", p3_first.ToString("dd-MMM-yyyy"));
                sql = sql.Replace("{30-JUN-2019}", p3_last.ToString("dd-MMM-yyyy"));

                sql = sql.Replace("{01-JUL-2019}", from_date);
                sql = sql.Replace("{31-JUL-2019}", to_date);


                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                // SID = "";
                // string jvh_docno = "";
                // decimal cr_total = 0;

                mrow = new CollectionReport();
                mrow.row_type = "HEADER";
                mrow.row_color = "BLACK";
                mrow.desc1 = "BRANCH";
                mrow.desc2 = "UPTO";
                mrow.col1 = fy_start;
                mrow.col2 = p1_first.ToString("MMM").ToUpper();
                mrow.col3 = p2_first.ToString("MMM").ToUpper();
                mrow.col4 = p3_first.ToString("MMM").ToUpper();
                mList.Add(mrow);

                string id = "";
                decimal c1 = 0;
                decimal c2 = 0;
                decimal c3 = 0;
                decimal c4 = 0;

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    if (id == "")
                        id = Dr["branch_code"].ToString();

                    if ( id != Dr["branch_code"].ToString())
                    {
                        mrow = new CollectionReport();
                        mrow.row_type = "HEADER";
                        mrow.row_color = "RED";
                        mrow.desc1 = "COLLECTION";
                        mrow.desc2 = "";
                        mrow.col1 = c1.ToString();
                        mrow.col2 = c2.ToString();
                        mrow.col3 = c3.ToString();
                        mrow.col4 = c4.ToString();
                        mList.Add(mrow);
                        c1 = 0;c2 = 0;c3 = 0;c4 = 0;
                    }
                    mrow = new CollectionReport();
                    mrow.row_type = "DETAIL";
                    mrow.row_color = "BLACK";
                    mrow.desc1 = Dr["branch_code"].ToString();
                    if (Dr["slno"].ToString() == "A")
                        mrow.desc2 = from_date;
                    if (Dr["slno"].ToString() == "B")
                        mrow.desc2 = to_date;

                    mrow.col1 = Dr["col1"].ToString();
                    mrow.col2 = Dr["col2"].ToString();
                    mrow.col3 = Dr["col3"].ToString();
                    mrow.col4 = Dr["col4"].ToString();
                    mList.Add(mrow);

                    id = Dr["branch_code"].ToString();
                    if ( Dr["slno"].ToString() == "A")
                    {
                        c1 += Lib.Conv2Decimal(Dr["col1"].ToString());
                        c2 += Lib.Conv2Decimal(Dr["col2"].ToString());
                        c3 += Lib.Conv2Decimal(Dr["col3"].ToString());
                        c4 += Lib.Conv2Decimal(Dr["col4"].ToString());
                    }
                    if (Dr["slno"].ToString() == "B")
                    {
                        c1 -= Lib.Conv2Decimal(Dr["col1"].ToString());
                        c2 -= Lib.Conv2Decimal(Dr["col2"].ToString());
                        c3 -= Lib.Conv2Decimal(Dr["col3"].ToString());
                        c4 -= Lib.Conv2Decimal(Dr["col4"].ToString());
                    }

                }

                if (id != "")
                {
                    mrow = new CollectionReport();
                    mrow.row_type = "HEADER";
                    mrow.row_color = "RED";
                    mrow.desc1 = "COLLECTION";
                    mrow.desc2 = "";
                    mrow.col1 = c1.ToString();
                    mrow.col2 = c2.ToString();
                    mrow.col3 = c3.ToString();
                    mrow.col4 = c4.ToString();
                    mList.Add(mrow);
                }


                if (Lib.CreateFolder(report_folder))
                    ProcessExcelFile();
                Dt_List.Rows.Clear();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", type);
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
            WS.Columns[1].Width = 256 * 25;
            WS.Columns[2].Width = 256 * 15;
            WS.Columns[3].Width = 256 * 15;
            WS.Columns[4].Width = 256 * 15;
            WS.Columns[5].Width = 256 * 15;
            WS.Columns[6].Width = 256 * 15;

            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            sTitle = "COLLECTION REPORT";

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;


            foreach (CollectionReport Dr in mList)
            {
                iRow++; iCol = 1;
                _Border = "";
                _Bold = false;

                if (Dr.row_color == "BLACK")
                    _Color = Color.Black;
                else
                    _Color = Color.Red;

                if (Dr.row_type == "HEADER")
                {
                    _Border = "TB";
                    _Bold = true;
                }

                Lib.WriteData(WS, iRow, iCol++, Dr.desc1, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr.desc2, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr.col1, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, Dr.col2, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, Dr.col3, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, Dr.col4, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

            }
            WB.SaveXls(File_Name + ".xls");


        }

        public object nvl(object svalue, object sret)
        {
            if (svalue == null)
                return sret;
            else
                return svalue;
        }


    }


}

