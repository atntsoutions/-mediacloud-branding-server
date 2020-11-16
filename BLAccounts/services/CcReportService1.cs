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
    public class CcReportService1 : BL_Base
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
        string CC_ID = "";
        string CC_CODE = "";
        string CC_NAME = "";
        string CC_TYPE = "";
        string CC_UPDATE = "N";
        string showIncExpOnly = "";
        string hide_ho_entries = "N";

        decimal nDr = 0, nCr = 0;


        List<LedgerReport> mList = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string SQL = "";

            string SID = "";

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            mList = new List<LedgerReport>();

            LedgerReport mrow;

            type = SearchData["type"].ToString();

            report_folder = SearchData["report_folder"].ToString();
            PKID = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();
            from_date = SearchData["from_date"].ToString();
            from_date = Lib.StringToDate(from_date).ToUpper();
            to_date = SearchData["to_date"].ToString();
            to_date = Lib.StringToDate(to_date).ToUpper();
            CC_ID = SearchData["cc_id"].ToString();
            CC_CODE = SearchData["cc_code"].ToString();
            CC_NAME = SearchData["cc_name"].ToString();
            CC_TYPE = SearchData["cc_type"].ToString();
            CC_UPDATE = SearchData["cc_update"].ToString();

            hide_ho_entries = SearchData["hide_ho_entries"].ToString();


            showIncExpOnly = SearchData["showIncExpOnly"].ToString();

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine(report_folder, PKID);


            try
            {
                if (CC_UPDATE == "Y")
                {
                    if (CC_TYPE == "MBL SEA EXPORT" || CC_TYPE == "MBL SEA IMPORT" || CC_TYPE == "MAWB AIR EXPORT" || CC_TYPE == "MAWB AIR IMPORT")
                        Lib.UpdateCCusingMBLID(CC_ID);
                    if (CC_TYPE == "SI SEA EXPORT" || CC_TYPE == "SI SEA IMPORT" || CC_TYPE == "SI AIR EXPORT" || CC_TYPE == "SI AIR IMPORT")
                        Lib.UpdateCCusingHBLID(CC_ID);
                    if (CC_TYPE == "JOB SEA EXPORT" || CC_TYPE == "JOB AIR EXPORT")
                        Lib.UpdateCCusingJOBID(CC_ID);
                }

                if (CC_TYPE != "ACC CODE")
                {
                    SQL = "";
                    SQL += " select  jvh_docno, jvh_date, ";
                    SQL += " cc_code, cc_name, cc_chwt, cc_cbm,";
                    SQL += " acc_code,acc_name,ct_category,";
                    SQL += " case when jv_drcr = 'DR' then ct_amount else 0 end as dr_amt,";
                    SQL += " case when jv_drcr = 'CR' then ct_amount else 0 end as cr_amt,";
                    SQL += " ct_remarks, jv_remarks  ";
                    SQL += " from ledgerh a ";
                    SQL += " inner join ledgert b on jvh_pkid = b.jv_parent_id";
                    SQL += " inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                    SQL += " inner join acctm d on jv_acc_id = acc_pkid";
                    SQL += " inner join costcenterm e on ct_cost_id = cc_pkid  ";
                    if (showIncExpOnly == "Y")
                        SQL += " inner join acgroupm f  on d.acc_group_id = f.acgrp_pkid ";
                    SQL += " where a.rec_company_code = '{COMPANY_CODE}' ";
                    if (CC_TYPE != "EMPLOYEE")
                        SQL += " and a.rec_branch_code ='{BRANCH_CODE}' ";
                    SQL += " and jvh_year = {FYEAR} ";
                    if (hide_ho_entries == "Y")
                        SQL += "  and jvh_type not in ('HO', 'IN-ES' ) ";
                    if (showIncExpOnly == "Y")
                        SQL += "and acgrp_name in('DIRECT INCOME','DIRECT EXPENSE', 'INDIRECT INCOME', 'INDIRECT EXPENSE') ";

                    //if (CC_TYPE.Trim().StartsWith("M"))
                    //    SQL += " and ct_cost_id  in (select cc_pkid from costcenterm where cc_parent_id='{PKID}' )";
                    //else
                    //    SQL += " and ct_cost_id ='{PKID}' ";

                    if (CC_ID.Trim().Length > 0)
                    {
                        if (CC_TYPE.Trim().StartsWith("M"))
                            SQL += " and ct_cost_id  in (select cc_pkid from costcenterm where cc_parent_id='{PKID}' )";
                        else
                            SQL += " and ct_cost_id ='{PKID}' ";
                    }

                        SQL += " and ct_year = {FYEAR} ";
                    if (from_date != "NULL" && to_date != "NULL")
                        SQL += " and cc_date >= '{FDATE}' and cc_date <= '{EDATE}' ";

                    if (CC_TYPE == "CNTR SEA EXPORT")
                        SQL += " and ct_posted  ='N'";
                    else
                        SQL += " and ct_posted  ='Y'";

                    if (CC_ID.Trim().Length > 0)
                        SQL += " order by jvh_date, jv_ctr ";
                    else
                        SQL += " order by  cc_code,jvh_date, jv_ctr ";
                }
                if (CC_TYPE == "ACC CODE")
                {
                    SQL = "";
                    SQL += " select  jvh_docno, jvh_date, ";
                    SQL += " cc_code, cc_name, cc_chwt, cc_cbm,";
                    SQL += " acc_code,acc_name,ct_category,";
                    SQL += " jv_debit, jv_credit,";
                    SQL += " case when jv_drcr = 'DR' and ct_posted = 'Y' then ct_amount else 0 end as dr_amt,";
                    SQL += " case when jv_drcr = 'CR' and ct_posted = 'Y' then ct_amount else 0 end as cr_amt,";
                    SQL += " ct_remarks, jv_remarks  ";
                    SQL += " from ledgerh a ";
                    SQL += " inner join ledgert b on jvh_pkid = b.jv_parent_id";
                    SQL += " left join costcentert c on b.jv_pkid = c.ct_jv_id";
                    SQL += " left join acctm d on jv_acc_id = acc_pkid";
                    SQL += " left join costcenterm e on ct_cost_id = cc_pkid  ";
                    if (showIncExpOnly == "Y")
                        SQL += " inner join acgroupm f  on d.acc_group_id = f.acgrp_pkid ";
                    SQL += " where a.rec_company_code = '{COMPANY_CODE}' ";
                    SQL += " and a.rec_branch_code ='{BRANCH_CODE}' ";
                    SQL += " and jvh_year = {FYEAR} and jvh_type not in ( 'OI','OC','OB') ";
                    if (hide_ho_entries == "Y")
                        SQL += "  and jvh_type not in ('HO', 'IN-ES') ";
                    if (showIncExpOnly == "Y")
                        SQL += "and acgrp_name in('DIRECT INCOME','DIRECT EXPENSE', 'INDIRECT INCOME', 'INDIRECT EXPENSE') ";
                    SQL += " and jv_acc_id ='{PKID}' ";
                    if (from_date != "NULL" && to_date != "NULL")
                        SQL += " and jvh_date >= '{FDATE}' and jvh_date <= '{EDATE}' ";
                    SQL += " order by ct_category, cc_code, jvh_date, jv_ctr ";
                }

                SQL = SQL.Replace("{COMPANY_CODE}", company_code);
                SQL = SQL.Replace("{BRANCH_CODE}", branch_code);

                SQL = SQL.Replace("{PKID}", CC_ID);
                SQL = SQL.Replace("{FYEAR}", year_code);
                SQL = SQL.Replace("{FDATE}", from_date);
                SQL = SQL.Replace("{EDATE}", to_date);


                Dt_List = Con_Oracle.ExecuteQuery(SQL);
                Con_Oracle.CloseConnection();

                nDr = 0; nCr = 0;
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new LedgerReport();
                    mrow.rowtype = "ROW";
                    mrow.rowcolor = "BLACK";
                    mrow.jv_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.jv_docno = Dr["jvh_docno"].ToString();
                    mrow.cc_code = Dr["cc_code"].ToString();
                    mrow.cc_name = Dr["cc_name"].ToString();
                    mrow.acc_code = Dr["acc_code"].ToString();
                    mrow.acc_name = Dr["acc_name"].ToString();
                    mrow.cc_category = Dr["ct_category"].ToString();
                    mrow.cc_remarks = Dr["ct_remarks"].ToString();
                    mrow.jv_remarks = Dr["jv_remarks"].ToString();
                    if ( Lib.Conv2Decimal(Dr["dr_amt"].ToString()) == 0 && Lib.Conv2Decimal(Dr["cr_amt"].ToString()) ==0)
                    {
                        Dr["dr_amt"] = Lib.Conv2Decimal(Dr["jv_Debit"].ToString());
                        Dr["cr_amt"] = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.cc_remarks = "COST CENTER NOT ALLOCATED";
                    }

                    
                    mrow.debit = Lib.Conv2Decimal(Dr["dr_amt"].ToString(), "NULL");
                    mrow.credit = Lib.Conv2Decimal(Dr["cr_amt"].ToString(), "NULL");
                    mrow.cc_chwt = Lib.Conv2Decimal(Dr["cc_chwt"].ToString(), "NULL");
                    mrow.cc_cbm = Lib.Conv2Decimal(Dr["cc_cbm"].ToString(), "NULL");
                    mList.Add(mrow);

                    // Grand Total
                    nDr += Lib.Conv2Decimal(Dr["dr_amt"].ToString());
                    nCr += Lib.Conv2Decimal(Dr["cr_amt"].ToString());

                }
                if (nDr != 0 || nCr != 0)
                {
                    mrow = new LedgerReport();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.acc_name = "TOTAL";
                    mrow.debit = nDr;
                    mrow.credit = nCr;
                    mList.Add(mrow);

                    mrow = new LedgerReport();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.acc_name = "BALANCE";
                    mrow.debit = nDr - nCr;
                    mList.Add(mrow);

                }
                if (type == "EXCEL")
                {
                    if (Lib.CreateFolder(report_folder))
                        ProcessExcelFile();
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
            RetData.Add("reportfile", File_Name);
            if (type != "EXCEL")
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
            WS.Columns[2].Width = 256 * 13;
            WS.Columns[3].Width = 256 * 14;
            WS.Columns[4].Width = 256 * 12;
            WS.Columns[5].Width = 256 * 30;
            WS.Columns[6].Width = 256 * 15;

            WS.Columns[7].Width = 256 * 30;
            WS.Columns[8].Width = 256 * 14;
            WS.Columns[9].Width = 256 * 14;
            WS.Columns[10].Width = 256 * 30;
            WS.Columns[11].Width = 256 * 30;

            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";


            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            //sTitle = "COST CENTER REPORT " + Lib.getFrontEndDate(to_date) ;
            sTitle = "COST CENTER REPORT - " + CC_TYPE + " - " + CC_CODE + " - " + CC_NAME ;

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;

            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CC.CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CC.NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "A/C.CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "A/C.NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "JV-REMARKS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            foreach (LedgerReport Dr in mList)
            {
                iRow++; iCol = 1;
                _Border = "";
                _Bold = false;
                _Color = Color.Black;

                if (Dr.rowtype.ToString() == "TOTAL" )
                {
                    _Border = "TB";
                    _Bold = true;
                    _Color = Color.DarkBlue;
                }

                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jv_docno,""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jv_date,""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cc_category, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cc_code,""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cc_name,""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.debit,0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.credit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cc_remarks, 0), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jv_remarks, 0), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

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

