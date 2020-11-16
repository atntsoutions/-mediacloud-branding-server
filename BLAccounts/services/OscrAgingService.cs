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
    public class OscrAgingService : BL_Base
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
        string ACC_ID = "";

        decimal tot_age1 = 0;
        decimal tot_age2 = 0;
        decimal tot_age3 = 0;
        decimal tot_age4 = 0;
        decimal tot_age5 = 0;
        decimal tot_age6 = 0;
        decimal tot_oneyear = 0;
        decimal tot_overdue = 0;
        decimal tot_balance = 0;
        decimal tot_advance = 0;
        Boolean all = false;
        Boolean IsTillDate = false;
        Boolean IsOverDue = false;

        List<LedgerReport> mList = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string SQL = "";
            string branch = "";

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
            string edate = SearchData["to_date"].ToString();
            to_date = Lib.StringToDate(edate).ToUpper();
            ACC_ID = SearchData["acc_id"].ToString();

            IsOverDue = (Boolean)SearchData["isoverdue"];
            all = (Boolean)SearchData["all"];

            //  string searchstring = SearchData["searchstring"].ToString().ToUpper();

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine(report_folder, PKID);

            IsTillDate = false;
            if (System.DateTime.Now.ToString(Lib.BACK_END_DATE_FORMAT).ToUpper() == to_date)
                IsTillDate = true;

            try
            {

                SQL += " select  branch_code,cust_name,cust_code,";
                SQL += "sum( case when os_days <= 15 then  balance  else 0 end) as age1,  ";
                SQL += "sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2,   ";
                SQL += "sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                SQL += "sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                SQL += "sum( case when os_days between 91 and 180 then  balance  else 0 end) as age5,   ";
                SQL += "sum( case when os_days > 180  then  balance  else 0 end) as age6, ";
                SQL += "sum(balance) as balance,  ";
                SQL += "sum(adv) as advance, ";
                SQL += "sum(case when overdue > 0  then  balance  else 0 end) as overdue, ";
                SQL += "sum( case when os_days >= 365  then  balance  else 0 end) as oneyear ";

                SQL += "from ( ";

                SQL += " select branch_code,jv_pkid,jv_acc_id,jvh_cc_id,jvh_vrno,jvh_docno,jvh_date,cust_code, cust_name, jv_debit, jv_credit,adv, balance, invtype,cust_crdays, cust_crlimit, os_days,os_days - cust_crdays as overdue ";


                SQL += " from (";
                SQL += " select h.rec_branch_code as branch_code, jv_pkid,jv_acc_id,jvh_cc_id, jvh_vrno,jvh_docno,jvh_date,max(L.rec_category) as INVTYPE,";
                SQL += " jv_credit, nvl(sum(xref_amt),0) as jv_debit, jv_credit - nvl(sum(xref_amt),0) as balance ,";
                SQL += "  0 as adv, ";

                //SQL += " trunc(sysdate - jvh_date,0)  as os_days   ";

                if (IsTillDate)
                    SQL += " max(trunc(sysdate -  nvl(nvl(jvh_org_invdt,jvh_reference_date),jvh_date)  ,0))  as os_days   ";
                else
                    SQL += " max(trunc( to_date('{EDATE}','DD-MON-YYYY') -  nvl(nvl(jvh_org_invdt,jvh_reference_date),jvh_date)  ,0))  as os_days   ";


                SQL += " from ledgerh h ";
                SQL += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id)";
                SQL += " inner join  Acctm a on (L.jv_acc_id = A.acc_pkid )";
                SQL += " inner join acgroupm g on (a.acc_group_id = g.acgrp_pkid and g.rec_company_code ='{COMPCODE}'  and acgrp_name  = 'SUNDRY CREDITORS' ) ";
                SQL += " left  join ledgerxref X on (L.jv_pkid=X.xref_cr_jv_id )";
                SQL += " left  join param s on ( jv_acc_id = param_pkid)";
                SQL += " where  ";

                SQL += " h.rec_company_code = '{COMPCODE}' ";
                if (!all)
                {
                    SQL += " and h.REC_BRANCH_CODE = '{BRCODE}' ";
                }

                if (ACC_ID != "")
                    SQL += " and jv_acc_id = '{PKID}'";
                if (IsTillDate == false)
                    SQL += " and jvh_date <= '{EDATE}' ";
                SQL += " and L.jv_credit >0 and h.rec_deleted  ='N' and acc_against_invoice ='C' and jvh_type not in('OP','OB','OC')    ";
                SQL += " group by h.rec_branch_code,jv_pkid,jv_acc_id,jvh_cc_id,jvh_date,jvh_vrno,jvh_docno,jv_credit  ";
                SQL += " having (jv_credit - nvl(sum(xref_amt),0)) !=0";
                SQL += "   ";

                SQL += " union all";
                SQL += "   ";
                SQL += " select h.rec_branch_code as branch_code,jv_pkid,jv_acc_id,null as jvh_cc_id,jvh_vrno,jvh_docno,jvh_date,max(L.rec_category) as INVTYPE,";
                SQL += " 0 as jv_debit, 0 as jv_credit,0 as balance,  ";
                SQL += " nvl(sum(xref_amt),0) - jv_debit   as adv,  ";
                SQL += " 0 as os_days ";
                SQL += " from ledgerh h ";
                SQL += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id )";
                SQL += " inner join  Acctm a on (L.jv_acc_id = A.acc_pkid )   ";
                SQL += " inner join acgroupm g on (a.acc_group_id = g.acgrp_pkid and g.rec_company_code ='{COMPCODE}' and acgrp_name  = 'SUNDRY CREDITORS')";
                SQL += " left join ledgerxref X on (L.jv_pkid=X.xref_dr_jv_id  )";
                SQL += " where ";

                SQL += " h.rec_company_code = '{COMPCODE}' ";
                if (!all)
                {
                    SQL += " and h.REC_BRANCH_CODE = '{BRCODE}' ";
                }

                if (ACC_ID != "")
                    SQL += " and jv_acc_id = '{PKID}'";
                if (IsTillDate == false)
                    SQL += " and jvh_date <= '{EDATE}' ";
                SQL += " and L.jv_credit >0 and h.REC_DELETED = 'N'  and acc_against_invoice ='C' and jvh_type not in('OP','OB','OC') ";
                SQL += " group by h.rec_branch_code,jv_pkid,jv_acc_id,jvh_date,jvh_vrno,jvh_docno,jv_debit";
                SQL += " having (jv_debit - nvl(sum(xref_amt),0)) !=0 ";
                SQL += " ) a ";
                SQL += " left join customerm cust on a.jv_acc_id = cust_pkid";

                if (all)
                {
                    SQL += " where branch_code not in('KOLAF','HOCPL')";
                }
                SQL += " ) group by branch_code,cust_code,cust_name ";
                SQL += " order by branch_code,cust_code,cust_name";


                SQL = SQL.Replace("{COMPCODE}", company_code);
                SQL = SQL.Replace("{BRCODE}", branch_code);
                SQL = SQL.Replace("{EDATE}", to_date);
                SQL = SQL.Replace("{PKID}", ACC_ID);

                Dt_List = Con_Oracle.ExecuteQuery(SQL);
                Con_Oracle.CloseConnection();

                tot_age1 = 0;
                tot_age2 = 0;
                tot_age3 = 0;
                tot_age4 = 0;
                tot_age5 = 0;
                tot_age6 = 0;
                tot_balance = 0;
                tot_oneyear = 0;
                tot_overdue = 0;
                tot_advance = 0;
                foreach (DataRow Dr in Dt_List.Rows)
                {

                    mrow = new LedgerReport();
                    mrow.rowtype = "ROW";
                    mrow.rowcolor = "BLACK";

                    mrow.cust_name = Dr["cust_name"].ToString();
                    mrow.cust_code = Dr["cust_code"].ToString();
                    mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());
                    mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());//total os
                    mrow.advance = Lib.Conv2Decimal(Dr["advance"].ToString());

                    branch = Dr["branch_code"].ToString();

                    mrow.branch = branch;
                    mList.Add(mrow);
                    tot_age1 += Lib.Conv2Decimal(mrow.age1.ToString());
                    tot_age2 += Lib.Conv2Decimal(mrow.age2.ToString());
                    tot_age3 += Lib.Conv2Decimal(mrow.age3.ToString());
                    tot_age4 += Lib.Conv2Decimal(mrow.age4.ToString());
                    tot_age5 += Lib.Conv2Decimal(mrow.age5.ToString());
                    tot_age6 += Lib.Conv2Decimal(mrow.age6.ToString());
                    tot_oneyear += Lib.Conv2Decimal(mrow.oneyear.ToString());
                    tot_overdue += Lib.Conv2Decimal(mrow.overdue.ToString());
                    tot_balance += Lib.Conv2Decimal(mrow.balance.ToString());
                    tot_advance += Lib.Conv2Decimal(mrow.advance.ToString());


                }
                if (mList.Count > 1)
                {
                    mrow = new LedgerReport();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.cust_code = "TOTAL";
                    mrow.age1 = Lib.Conv2Decimal(Lib.NumericFormat(tot_age1.ToString(), 2));
                    mrow.age2 = Lib.Conv2Decimal(Lib.NumericFormat(tot_age2.ToString(), 2));
                    mrow.age3 = Lib.Conv2Decimal(Lib.NumericFormat(tot_age3.ToString(), 2));
                    mrow.age4 = Lib.Conv2Decimal(Lib.NumericFormat(tot_age4.ToString(), 2));
                    mrow.age5 = Lib.Conv2Decimal(Lib.NumericFormat(tot_age5.ToString(), 2));
                    mrow.age6 = Lib.Conv2Decimal(Lib.NumericFormat(tot_age6.ToString(), 2));
                    mrow.oneyear = Lib.Conv2Decimal(Lib.NumericFormat(tot_oneyear.ToString(), 2));
                    mrow.overdue = Lib.Conv2Decimal(Lib.NumericFormat(tot_overdue.ToString(), 2));
                    mrow.balance = Lib.Conv2Decimal(Lib.NumericFormat(tot_balance.ToString(), 2));
                    mrow.advance = Lib.Conv2Decimal(Lib.NumericFormat(tot_advance.ToString(), 2));

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
            WS.Columns[1].Width = 256 * 15;
            WS.Columns[2].Width = 256 * 25;
            WS.Columns[3].Width = 256 * 15;
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


            WS.Columns[2].Style.NumberFormat = "#,0.00";
            WS.Columns[3].Style.NumberFormat = "#,0.00";
            WS.Columns[4].Style.NumberFormat = "#,0.00";
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



            iRow = 1; iCol = 1;
            if (all)
            {
                iRow = Lib.WriteAddress(WS, "HOCPL", iRow, iCol);
            }
            else
            {
                iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);
            }
            sTitle = "AGING REPORT CREDITORS AS ON " + Lib.getFrontEndDate(to_date);

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;
            if (all)
            {
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            }
            Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "0-15", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "16-30", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TOTAL OS", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "OVERDUE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "1YEAR+", _Color, true, _Border, "R", "", _Size, false, 325, "", true);


            foreach (LedgerReport Dr in mList)
            {
                iRow++; iCol = 1;
                _Border = "";
                _Bold = false;
                _Color = Color.Black;

                if (Dr.rowtype.ToString() == "TOTAL")
                {
                    _Border = "TB";

                    _Bold = true;
                    _Color = Color.DarkBlue;
                }
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.branch, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cust_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cust_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age1, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age2, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age3, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age4, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age5, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age6, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.balance, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.advance, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.overdue, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.oneyear, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);


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

