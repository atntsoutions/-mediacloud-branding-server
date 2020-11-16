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
    public class OsReportService : BL_Base
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
        decimal _nDr = 0, _nCr = 0, _nBal = 0, _nAdv = 0;
        decimal nDr = 0, nCr = 0, nBal = 0, nAdv = 0;
        Boolean all = false;
        Boolean IsTillDate = false;
        Boolean do_not_use_credit_date = false;
        Boolean IsOverDue = false;

        DateTime jv_date;
        DateTime due_date;
        int cr_days = 0;
        List<LedgerReport> mList = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string SQL = "";


            string SID = "";


            Decimal age1 = 0,age2=0, age3=0,age4=0,age5=0,age6 = 0,oneyear = 0;
            Decimal _age1 = 0, _age2 = 0, _age3 = 0, _age4 = 0, _age5 = 0,_age6 = 0, _oneyear = 0;

            string sname = "";
            string sman_name = "";
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

            IsOverDue = (Boolean) SearchData["isoverdue"];
            all = (Boolean)SearchData["all"];
            do_not_use_credit_date = (Boolean)SearchData["do_not_use_credit_date"];

            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine(report_folder, PKID);

            IsTillDate = false;
            if (System.DateTime.Now.ToString(Lib.BACK_END_DATE_FORMAT).ToUpper() == to_date)
                IsTillDate = true;

            string DRSQL = "";
            string CRSQL = "";

            try
            {
                SQL = "";
                SQL += " select branch_code,jv_pkid,jv_acc_id,jvh_cc_id,jvh_vrno,jvh_docno,jvh_date,cust_code, cust_name, jv_debit, jv_credit,adv, balance, invtype,cust_crdays, cust_crlimit, os_days,os_days - cust_crdays as overdue, ";
                SQL += " nvl( sman1.param_name,nvl(sman2.param_name, sman.param_name))  as sman_name,  jv_od_type, jv_od_remarks ";
                SQL += " from (";
                SQL += " select h.rec_branch_code as branch_code, jv_pkid,jv_acc_id,jvh_cc_id, jvh_vrno,jvh_docno,jvh_date,max(h.rec_category) as INVTYPE,";
                SQL += " jv_debit, nvl(sum(xref_amt),0) as jv_credit, jv_debit - nvl(sum(xref_amt),0) as balance ,  ";
                SQL += "  0 as adv, max(jv_od_type) as jv_od_type, max(jv_od_remarks) as jv_od_remarks,";
                SQL += " trunc( to_date('{EDATE}') - jvh_date,0)  as os_days   ";
                SQL += " from ledgerh h ";
                SQL += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id )";
                SQL += " inner join Acctm a on (L.jv_acc_id = A.acc_pkid )";
                SQL += " inner join acgroupm g on a.acc_group_id = g.acgrp_pkid and g.rec_company_code ='{COMPCODE}'  and acgrp_name  = 'SUNDRY DEBTORS'";

                SQL += " left  join ledgerxref X on (L.jv_pkid=X.xref_dr_jv_id  {DRSQL} )";

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
                {
                    SQL += " and jvh_date <= '{EDATE}' ";
                    if(do_not_use_credit_date  == false)
                    {
                        DRSQL =  " and X.XREF_CR_JV_DATE <= '{EDATE}' ";
                    }
                }


                SQL += " and L.jv_debit >0 and h.rec_deleted  ='N' and acc_against_invoice ='D' and jvh_type not in('OP','OB','OC')    ";
                SQL += " group by h.rec_branch_code,jv_pkid,jv_acc_id,jvh_cc_id,jvh_date,jvh_vrno,jvh_docno,jv_debit  ";
                SQL += " having (jv_debit - nvl(sum(xref_amt),0)) !=0";
                SQL += "   ";

                if (IsOverDue == false)
                {
                    SQL += " union all";
                    SQL += "   ";
                    SQL += " select h.rec_branch_code as branch_code,jv_pkid,jv_acc_id,null as jvh_cc_id,jvh_vrno,jvh_docno,jvh_date,max(h.rec_category) as INVTYPE,";
                    SQL += " 0 as jv_debit, 0 as jv_credit,0 as balance,  ";
                    SQL += " nvl(sum(xref_amt),0) - jv_credit   as adv,   max(jv_od_type) as jv_od_type, max(jv_od_remarks) as jv_od_remarks, ";
                    SQL += " 0 as os_days ";
                    SQL += " from ledgerh h ";
                    SQL += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id )";
                    SQL += " inner join Acctm a on (L.jv_acc_id = A.acc_pkid )   ";
                    SQL += " inner join acgroupm g on a.acc_group_id = g.acgrp_pkid and g.rec_company_code ='{COMPCODE}' and acgrp_name  = 'SUNDRY DEBTORS'";
                    SQL += " left join ledgerxref X on (L.jv_pkid=X.xref_cr_jv_id {CRSQL} )";
                    SQL += " where ";
                    
                    SQL += " h.rec_company_code = '{COMPCODE}' ";
                    if (!all)
                    {
                        SQL += " and h.REC_BRANCH_CODE = '{BRCODE}' ";
                    }
                    
                    if (ACC_ID != "")
                        SQL += " and jv_acc_id = '{PKID}'";
                    if (IsTillDate == false)
                    {
                        SQL += " and jvh_date <= '{EDATE}' ";

                        if (do_not_use_credit_date == false)
                        {
                            CRSQL = " and X.XREF_DR_JV_DATE <= '{EDATE}' ";
                            
                        }

                    }
                    SQL += " and L.jv_credit >0 and h.REC_DELETED = 'N'  and acc_against_invoice ='D' and jvh_type not in('OP','OB','OC') ";
                    SQL += " group by h.rec_branch_code,jv_pkid,jv_acc_id,jvh_date,jvh_vrno,jvh_docno,jv_credit";
                    SQL += " having (jv_credit - nvl(sum(xref_amt),0)) !=0 ";
                }

                SQL += " ) a ";
                SQL += " left join customerm cust on a.jv_acc_id = cust_pkid";
                SQL += " left join hblm on jvh_cc_id = hbl_pkid";
                SQL += " left join custdet cd on a.jv_acc_id = cd.det_cust_id and a.branch_code =  cd.det_branch_code ";
                SQL += " left join param sman  on cust_sman_id = sman.param_pkid";
                SQL += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                SQL += " left join param sman1   on hbl_salesman_id = sman1.param_pkid";

                if (IsOverDue)
                    SQL += " where (os_days - cust_crdays) > 0 ";

                SQL += " order by branch_code,cust_name, jvh_date, jvh_vrno";

                SQL = SQL.Replace("{CRSQL}", CRSQL);
                SQL = SQL.Replace("{DRSQL}", DRSQL);

                SQL = SQL.Replace("{COMPCODE}",company_code);
                SQL = SQL.Replace("{BRCODE}", branch_code);
                SQL = SQL.Replace("{EDATE}", to_date);
                SQL = SQL.Replace("{PKID}", ACC_ID);

                

                Dt_List = Con_Oracle.ExecuteQuery(SQL);
                Con_Oracle.CloseConnection();

                SID = "";
                cr_days = 0;
                jv_date = new DateTime();
                foreach (DataRow Dr in Dt_List.Rows)
                {
                   
                    if (SID != Dr["jv_acc_id"].ToString())
                    {
                        if (SID != "")
                        {
                            mrow = new LedgerReport();
                            mrow.rowtype = "TOTAL";
                            mrow.rowcolor = "RED";
                            //  mrow.jv_docno = "TOTAL";
                            mrow.acc_code = "TOTAL";


                            mrow.pkid = "";
                            mrow.jv_od_type = "";
                            mrow.jv_od_remarks = "";

                            mrow.rec_category = "";


                            mrow.acc_name = sname;
                            mrow.sman_name = sman_name;
                            mrow.branch = branch;
                            mrow.debit = _nDr;
                            mrow.credit = _nCr;
                            mrow.bal = _nBal;
                            mrow.advance = _nAdv;
                            mrow.age1 = _age1;
                            mrow.age2 = _age2;
                            mrow.age3 = _age3;
                            mrow.age4 = _age4;
                            mrow.age5 = _age5;
                            mrow.age6 = _age6;
                            mrow.oneyear = _oneyear;

                            mList.Add(mrow);
                            _nDr = 0; _nCr = 0; _nBal = 0; _nAdv = 0;
                            _age1 = 0; _age2 = 0;_age3 = 0;_age4 = 0;_age5 = 0;_age6 = 0;_oneyear = 0;
                            sname = "";
                            sman_name = "";
                            branch = "";
                        }
                        SID = Dr["jv_acc_id"].ToString();
                    }
                    mrow = new LedgerReport();

                    jv_date = new DateTime();
                    due_date = new DateTime();
                    cr_days = 0;

                    mrow.rowtype = "ROW";
                    mrow.rowcolor = "BLACK";

                    mrow.pkid = Dr["jv_pkid"].ToString();
                    mrow.acc_code = Dr["cust_code"].ToString();
                    sname = Dr["cust_name"].ToString();
                    mrow.acc_name = sname;

                    sman_name = Dr["sman_name"].ToString();
                    mrow.sman_name = sman_name;

                    mrow.jv_od_type = Dr["jv_od_type"].ToString();
                    mrow.jv_od_remarks = Dr["jv_od_remarks"].ToString();


                    mrow.rec_category = Dr["invtype"].ToString();

                    //  mrow.jv_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    // mrow.crdays = Lib.Conv2Decimal(Dr["cust_crdays"].ToString());

                    jv_date =   (DateTime)(Dr["jvh_date"]);
                    mrow.jv_date = Lib.DatetoStringDisplayformat(jv_date);
                    cr_days = Lib.Conv2Integer(Dr["cust_crdays"].ToString());
                    mrow.crdays = cr_days;
                    due_date = jv_date.AddDays(cr_days);
                    mrow.due_date = Lib.DatetoStringDisplayformat(due_date);
                   
                    mrow.jv_docno = Dr["jvh_docno"].ToString();
                    mrow.debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    mrow.credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                    // LATEST CHANGE - BENNY
                    /*
                    if (Lib.Conv2Decimal(Dr["balance"].ToString()) > 0)
                        mrow.bal = Lib.Conv2Decimal(Dr["balance"].ToString());
                    */
                    mrow.bal = Lib.Conv2Decimal(Dr["balance"].ToString());

                    mrow.advance = Lib.Conv2Decimal(Dr["adv"].ToString());
                  
                    mrow.crlimit = Lib.Conv2Decimal(Dr["cust_crlimit"].ToString());
                    mrow.osdays = Lib.Conv2Decimal(Dr["os_days"].ToString());
                    mrow.overduedays = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    branch = Dr["branch_code"].ToString();
                    mrow.branch = branch;

                   
                    if (mrow.osdays <= 15)
                    {
                        mrow.age1 = mrow.bal;
                        _age1 += Lib.Conv2Decimal(Dr["balance"].ToString());
                        age1 += Lib.Conv2Decimal(Dr["balance"].ToString());
                    }
                    if (mrow.osdays >= 16 && mrow.osdays <= 30)
                    {
                        mrow.age2 = mrow.bal;
                        _age2 += Lib.Conv2Decimal(Dr["balance"].ToString());
                        age2 += Lib.Conv2Decimal(Dr["balance"].ToString());
                    }
                    if (mrow.osdays >=31 && mrow.osdays <= 60)
                    {
                        mrow.age3 = mrow.bal;
                        _age3 += Lib.Conv2Decimal(Dr["balance"].ToString());
                        age3 += Lib.Conv2Decimal(Dr["balance"].ToString());
                    }
                    if (mrow.osdays >= 61 && mrow.osdays <= 90)
                    {
                        mrow.age4 = mrow.bal;
                        _age4 += Lib.Conv2Decimal(Dr["balance"].ToString());
                        age4 += Lib.Conv2Decimal(Dr["balance"].ToString());
                    }
                    if (mrow.osdays >= 91 && mrow.osdays <= 180)
                    {
                        mrow.age5 = mrow.bal;
                        _age5 += Lib.Conv2Decimal(Dr["balance"].ToString());
                        age5 += Lib.Conv2Decimal(Dr["balance"].ToString());
                    }
                    if (mrow.osdays > 180 )
                    {
                        mrow.age6 = mrow.bal;
                        _age6 += Lib.Conv2Decimal(Dr["balance"].ToString());
                        age6 += Lib.Conv2Decimal(Dr["balance"].ToString());
                    }
                    if (mrow.osdays >= 365)
                    {
                        mrow.oneyear = mrow.bal;
                        _oneyear += Lib.Conv2Decimal(Dr["balance"].ToString());
                        oneyear += Lib.Conv2Decimal(Dr["balance"].ToString());
                    }
                    mList.Add(mrow);

                    // Customer Total
                    _nDr += Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    _nCr += Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                    // LATEST CHANGE - BENNY
                    /*
                    if (Lib.Conv2Decimal(Dr["balance"].ToString()) > 0)
                        _nBal += Lib.Conv2Decimal(Dr["balance"].ToString());
                    */
                    _nBal += Lib.Conv2Decimal(Dr["balance"].ToString());


                    _nAdv += Lib.Conv2Decimal(Dr["adv"].ToString());

                    // Grand Total
                    nDr += Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    nCr += Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                    // LATEST CHANGE - BENNY
                    /*
                    if (Lib.Conv2Decimal(Dr["balance"].ToString()) > 0)
                        nBal += Lib.Conv2Decimal(Dr["balance"].ToString());
                    */

                    nBal += Lib.Conv2Decimal(Dr["balance"].ToString());

                    nAdv += Lib.Conv2Decimal(Dr["adv"].ToString());
                }
                if (SID != "")
                {
                    mrow = new LedgerReport();

                    mrow.rowtype = "TOTAL";

                    mrow.rowcolor = "RED";
                    // mrow.jv_docno = "TOTAL";

                    mrow.pkid = "";
                    mrow.jv_od_type  = "";
                    mrow.jv_od_remarks = "";
                    mrow.rec_category = "";

                    mrow.acc_code = "TOTAL";
                    mrow.branch = branch;
                    mrow.acc_name = sname;
                    mrow.sman_name = sman_name;
                    mrow.debit = _nDr;
                    mrow.credit = _nCr;
                    mrow.bal = _nBal;
                    mrow.advance = _nAdv;
                    mrow.age1 = _age1;
                    mrow.age2 = _age2;
                    mrow.age3 = _age3;
                    mrow.age4 = _age4;
                    mrow.age5 = _age5;
                    mrow.age6 = _age6;
                    mrow.oneyear = _oneyear;
                    
                    mList.Add(mrow);

                    mrow = new LedgerReport();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "ORANGE";
                  //  mrow.jv_docno = "TOTAL";
                    mrow.acc_code = "TOTAL";
                    mrow.acc_name = "GRAND TOTAL";

                    mrow.pkid = "";
                    mrow.jv_od_type = "";
                    mrow.jv_od_remarks = "";
                    mrow.rec_category = "";

                    mrow.debit = nDr;
                    mrow.credit = nCr;
                    mrow.bal = nBal;
                    mrow.advance = nAdv;

                    mrow.age1 = age1;
                    mrow.age2 = age2;
                    mrow.age3 = age3;
                    mrow.age4 = age4;
                    mrow.age5 = age5;
                    mrow.age6 = age6;
                    mrow.oneyear = oneyear;

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


            

            Boolean IsOverDue = false;

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
            WS.Columns[16].Width = 256 * 15;
            WS.Columns[17].Width = 256 * 15;
            WS.Columns[18].Width = 256 * 15;
            WS.Columns[19].Width = 256 * 15;
            WS.Columns[20].Width = 256 * 15;
            WS.Columns[21].Width = 256 * 15;
            WS.Columns[22].Width = 256 * 15;
            WS.Columns[23].Width = 256 * 15;



            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";
            WS.Columns[10].Style.NumberFormat = "#,0.00";
         //   WS.Columns[10].Style.NumberFormat = "#,0.00";
            WS.Columns[11].Style.NumberFormat = "#,0.00";

            WS.Columns[12].Style.NumberFormat = "#,0.00";
            WS.Columns[13].Style.NumberFormat = "#,0.00";
            WS.Columns[14].Style.NumberFormat = "#,0.00";
            WS.Columns[15].Style.NumberFormat = "#,0.00";
            WS.Columns[16].Style.NumberFormat = "#,0.00";
            WS.Columns[17].Style.NumberFormat = "#,0.00";
            WS.Columns[18].Style.NumberFormat = "#,0.00";


            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            sTitle = "DEBTORS OUTSTANDING REPORT AS ON " + Lib.getFrontEndDate(to_date);

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
            Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CR-LIMIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CR-DAYS", _Color, true, _Border, "C", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            
            Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "OS-DAYS", _Color, true, _Border, "C", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "OVERDUE", _Color, true, _Border, "C", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DUE-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "0-15", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "16-30", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            


            foreach (LedgerReport Dr in mList)
            {
                iRow++; iCol = 1;
                _Border = "";
                _Bold = false;
                _Color = Color.Black;

                if (Dr.rowtype.ToString() == "TOTAL" || Dr.rowtype.ToString() == "GRANDTOTAL")
                {
                    _Border = "TB";
                   
                    _Bold = true;
                    _Color = Color.DarkBlue;
                }


                IsOverDue = false;
                if ( Dr.overduedays > 0)
                    IsOverDue = true;

                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.branch, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_code, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.sman_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.rec_category, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.crlimit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.crdays, 0), _Color, _Bold, _Border, "C", "", _Size, false, 325, "#;(#);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jv_docno, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


                Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Dr.jv_date,"", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, "dd/MM/yyyy", true);
               

                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.debit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.credit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                if ( IsOverDue)
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.bal, 0), Color.Red, true, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                else
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.bal, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.advance, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.osdays, 0), _Color, _Bold, _Border, "C", "", _Size, false, 325, "#;(#);#", false);

                if (IsOverDue)
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.overduedays, 0), Color.Red, true, _Border, "C", "", _Size, false, 325, "#;(#);#", false);
                else
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.overduedays, 0), _Color, _Bold, _Border, "C", "", _Size, false, 325, "#;(#);#", false);

                Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Dr.due_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, "dd/MM/yyyy", true);

                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age1, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age2, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age3, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age4, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age5, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, nvl(Dr.age6, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
               

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

