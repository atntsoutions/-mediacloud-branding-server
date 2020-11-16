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
    public class PayHistoryService : BL_Base
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
        string intrest ="";
        string credit_days = "";
      
        Boolean all = false;
        Boolean detail = false;


        Boolean IsOverDue = false;

        List<PayHistroyReport> mList = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sql = "";
            //string SID = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            mList = new List<PayHistroyReport>();

            PayHistroyReport mrow;

            type = SearchData["type"].ToString();

            report_folder = SearchData["report_folder"].ToString();
            PKID = SearchData["pkid"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            string edate = SearchData["to_date"].ToString();
            to_date = Lib.StringToDate(edate).ToUpper();
            string fdate = SearchData["from_date"].ToString();
            from_date = Lib.StringToDate(fdate).ToUpper();

            ACC_ID = SearchData["acc_id"].ToString();

            IsOverDue = (Boolean) SearchData["isoverdue"];
            all = (Boolean)SearchData["all"];
            detail = (Boolean)SearchData["detail"];

            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            if(SearchData.ContainsKey("intrest"))
                intrest = SearchData["intrest"].ToString();

            if (SearchData.ContainsKey("credit_days"))
                credit_days = SearchData["credit_days"].ToString();

            Dt_List = new DataTable();
            report_folder = System.IO.Path.Combine(report_folder, PKID);
            File_Name = System.IO.Path.Combine(report_folder, PKID);



            try
            {
                if (detail)
                {
                    sql = " select   acc_name,rec_branch_code,jv_pkid,slno ,jvh_vrno, jvh_type,jvh_date, jv_debit,cr_date, {INTREST} as interest,   ";
                    sql += " case when xref_amt >0 and days > 0 then days else 0 end as days,  xref_amt,cr_total,   ";
                    sql += " case when balance> 0  and bal_days > 0 and slno = rows_count then bal_days else 0 end as bal_days,   ";
                    sql += " case when balance> 0  and bal_days > 0 and slno = rows_count then balance else 0 end as balance,   ";
                    sql += " case when xref_amt > 0 and days >0 then round(xref_amt * {INTREST}/100/365 *  days ,0) else 0 end  as int1,  ";
                    sql += " case when balance  > 0 and bal_days > 0 and slno = rows_count then round(balance  * {INTREST}/100 /365 *  bal_days,0) else 0 end as int2  from  ";
                    sql += " (  	 ";
                    sql += " 	select  type,rec_branch_code,jvh_vrno, jvh_type,acc_name,a.jv_pkid, a.jvh_date,  a.jv_debit,cr_date, xref_amt,  	 round(cr_date - a.jvh_date2 ) as days, ";
                    sql += " 	round(sysdate - a.jvh_date2 ) as bal_days,  	 sum(xref_amt) over( partition by a.jv_pkid order by type,a.jv_pkid,a.jvh_date,cr_date) as cr_total, ";
                    sql += " 	sum(a.jv_debit - xref_amt) over( partition by a.jv_pkid order by a.jv_pkid) as balance, ";
                    sql += " 	ROW_NUMBER() OVER(PARTITION BY a.jv_pkid ";
                    sql += " 	order by a.jv_pkid,type,a.jvh_date, cr_date ) as slno,   ";
                    sql += " 	count(jv_pkid) OVER(PARTITION BY a.jv_pkid  ) as rows_count from   ";
                    sql += " 	(  	 	  ";
                    sql += " 		select 'A' as type,a.rec_branch_code,jv_pkid,jvh_vrno, jvh_type, jvh_date, jvh_date + {CRDAYS} as jvh_date2, acc_name, null as cr_date, ";
                    sql += " 		jv_debit, 0 as xref_amt  ";
                    sql += " 		from ledgerh a  ";
                    sql += " 		inner join ledgert b on  jvh_pkid = jv_parent_id";
                    sql += " 		inner join acctm c on (b.jv_acc_id = c.acc_pkid )";
                    sql += " 		inner join acgroupm g on c.acc_group_id = g.acgrp_pkid and g.rec_company_code ='{COMPCODE}'  and acgrp_name  = 'SUNDRY DEBTORS' ";
                    sql += " 		where  a.rec_company_code = '{COMPCODE}' ";
                    if (!all)
                    {
                        sql += " and  a.rec_branch_code = '{BRCODE}' ";
                    }

                    if (ACC_ID != "")
                        sql += " and b.jv_acc_id = '{PKID}'";

                    sql += " 		and jvh_date between '{FDATE}' and '{EDATE}' and jv_debit >0 and jvh_type <> 'OP'  ";
                    sql += " 		union all   ";
                    sql += " 		select 'B' as type,a.rec_branch_code,xref_dr_jv_id,jvh_vrno, jvh_type,jvh_date, jvh_date + {CRDAYS} as jvh_date2, acc_name,xref_cr_jv_date , ";
                    sql += " 		0 as jv_debit, sum(xref_amt) as xref_amt ";
                    sql += " 		from ledgerxref  xref   ";
                    sql += " 		inner join ledgert b on xref_dr_jv_id = jv_pkid";
                    sql += " 		inner join ledgerh a on jv_parent_id = jvh_pkid	 ";
                    sql += " 		inner join acctm c on (b.jv_acc_id = c.acc_pkid ) ";
                    sql += " 		inner join acgroupm g on c.acc_group_id = g.acgrp_pkid and g.rec_company_code ='{COMPCODE}'  and acgrp_name  = 'SUNDRY DEBTORS' ";

                    sql += " 		where   a.rec_company_code = '{COMPCODE}' ";
                    if (!all)
                    {
                        sql += " and   a.rec_branch_code = '{BRCODE}'    ";
                    }

                    if (ACC_ID != "")
                        sql += " and b.jv_acc_id = '{PKID}'";

                    sql += " 		and xref_cr_jv_date between '{FDATE}' and '{EDATE}' ";


                    sql += " 		group by a.rec_branch_code,xref_dr_jv_id,jvh_vrno,acc_name,jvh_date, jvh_type, xref_cr_jv_date ";
                    sql += " 	) a  ";
                    sql += " ) a  order by rec_branch_code,jvh_date,jvh_vrno,jvh_type,slno";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{PKID}", ACC_ID);
                    sql = sql.Replace("{INTREST}", intrest);
                    sql = sql.Replace("{CRDAYS}", credit_days);

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                   // SID = "";
                    // string jvh_docno = "";
                    // decimal cr_total = 0;
                    foreach (DataRow Dr in Dt_List.Rows)
                    {


                        mrow = new PayHistroyReport();
                        mrow.rowtype = "ROW";
                        mrow.rowcolor = "BLACK";

                        mrow.cr_total = Lib.Conv2Decimal(Dr["cr_total"].ToString());

                        mrow.jvh_vrno = Dr["jvh_vrno"].ToString();
                        mrow.jvh_type = Dr["jvh_type"].ToString();
                        mrow.sl_no = Dr["slno"].ToString();
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.intrest = Lib.Conv2Decimal(Dr["interest"].ToString());
                        mrow.acc_name = Dr["acc_name"].ToString();
                        mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                        mrow.cr_date = Lib.DatetoStringDisplayformat(Dr["cr_date"]);

                        mrow.days = Lib.Conv2Decimal(Dr["days"].ToString());
                        mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                        mrow.bal_days = Lib.Conv2Decimal(Dr["bal_days"].ToString());
                        mrow.int1 = Lib.Conv2Decimal(Dr["int1"].ToString());
                        mrow.int2 = Lib.Conv2Decimal(Dr["int2"].ToString());

                        mrow.branch = Dr["rec_branch_code"].ToString();

                        mList.Add(mrow);

                    }
                    if (type == "EXCEL")
                    {
                        if (Lib.CreateFolder(report_folder))
                            ProcessExcelFile();
                    }
                    Dt_List.Rows.Clear();
                }
                if (!detail)
                {

                    sql = "    select a.*,  ";
                    sql += " case when (pending - nvl(cust_crdays,0)) >0 then pending - nvl(cust_crdays,0) else null end as overdue ";
                    sql += " from ( ";
                    sql += " select a.rec_branch_code as branch,a.jvh_pkid,a.jvh_vrno,a.jvh_type,a.jvh_date,c.acc_code, c.acc_name,b.jv_debit,b.jv_credit,xref_crdate,xref_amt,  ";
                    sql += " cust_crlimit,cust_crdays, nvl(sman2.param_name, sman.param_name)  as sman_name,  round(nvl(xref_crdate,sysdate) - jvh_date,0) as pending, ";
                    sql += " case when (jv_debit - nvl(xref_amt,0)) <> 0 then 'PENDING' else null end as status ";
                    sql += " from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id  ";
                    sql += " left join   (";
                    sql += " select xref_dr_jv_id, max(xref_cr_jv_date) as XREF_CRDATE, sum(xref_amt) as XREF_AMT  ";
                    sql += " from ledgerxref where   xref_cr_jv_date >= '{FDATE}'  ";
                    sql += " and xref_cr_jv_date <= '{EDATE}'  ";

                    if (!all)
                    {
                        sql += " and rec_branch_code = '{BRCODE}'";       
                    }

                    sql += " group by xref_dr_jv_id";
                    sql += " ) b on (b.jv_pkid = b.xref_dr_jv_id)";
                    sql += " inner join acctm c on (b.jv_acc_id = c.acc_pkid )  ";

                    sql += " inner join acgroupm g on c.acc_group_id = g.acgrp_pkid and g.rec_company_code ='{COMPCODE}'  and acgrp_name  = 'SUNDRY DEBTORS' ";

                    sql += " left join customerm shpr on (b.jv_acc_id = shpr.cust_pkid)  ";
                    sql += "  left join custdet cd on b.jv_acc_id = cd.det_cust_id and b.rec_branch_code =  cd.det_branch_code ";
                    sql += " left join param  sman on (shpr.cust_sman_id = sman.param_pkid)";
                    sql += "  left join param sman2 on (cd.det_sman_id = sman2.param_pkid)";
                    
                    sql += " left join yearm y on(a.jvh_year = y.year_code and y.rec_company_code ='{COMPCODE}')  ";
                    sql += " where  ";
                    sql += "  a.jvh_date   >= '{FDATE}' and  a.jvh_date <= '{EDATE}' and jv_debit >0 and jvh_type <> 'OP'  ";
                    if(!all)
                    {
                        sql += " and  a.rec_branch_code = '{BRCODE}'";
                    }

                    if (ACC_ID != "")
                        sql += " and b.jv_acc_id = '{PKID}'";

                   // sql += " and a.rec_deleted = 'N'  and ";
                  //  sql += " b.jv_debit >0 and ( (a.jvh_type like 'IN%' ) or a.jvh_type = 'OI')";
                    sql += "  ";
                    sql += " ) a  order by branch,jvh_date,jvh_vrno,jvh_type ";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{EDATE}", to_date);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{PKID}", ACC_ID);
                    //sql = sql.Replace("{INTREST}", intrest);
                    //sql = sql.Replace("{CRDAYS}", credit_days);

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {


                        mrow = new PayHistroyReport();
                        mrow.rowtype = "ROW";
                        mrow.rowcolor = "BLACK";
                        
                        mrow.jvh_vrno = Dr["jvh_vrno"].ToString();                   
                        mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);                  
                        mrow.acc_name = Dr["acc_name"].ToString();
                        mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                        mrow.xref_crdate = Lib.DatetoStringDisplayformat(Dr["xref_crdate"]);
                        mrow.xref_amt = Lib.Conv2Decimal(Dr["xref_amt"].ToString());
                        mrow.status = Dr["status"].ToString();
                        mrow.pending = Lib.Conv2Decimal(Dr["pending"].ToString());
                        mrow.cust_crdays = Lib.Conv2Decimal(Dr["cust_crdays"].ToString());
                        mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                        mrow.cust_crlimit = Lib.Conv2Decimal(Dr["cust_crlimit"].ToString());
                        mrow.sman_name = Dr["sman_name"].ToString();
                        mrow.branch = Dr["branch"].ToString();

                        mList.Add(mrow);

                    }
                    if (type == "EXCEL")
                    {
                        if (Lib.CreateFolder(report_folder))
                            ProcessExcelFile();
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

            if (detail)
            {
                WS.Columns[0].Width = 256;
                WS.Columns[1].Width = 256 * 12;
                WS.Columns[2].Width = 256 * 8;
                WS.Columns[3].Width = 256 * 8;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 12;

                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 12;
                WS.Columns[9].Width = 256 * 12;
                WS.Columns[10].Width = 256 * 12;
                WS.Columns[11].Width = 256 * 12;
                WS.Columns[12].Width = 256 * 12;
                WS.Columns[13].Width = 256 * 12;

                WS.Columns[14].Width = 256 * 12;
                WS.Columns[15].Width = 256 * 12;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;
                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;

                iRow = 1; iCol = 1;

                iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

                sTitle = "PAYMENT HISTORY REPORT FROM " + Lib.getFrontEndDate(from_date) + " TO " + Lib.getFrontEndDate(to_date);

                Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

                iCol = 1;
                _Color = Color.DarkBlue;
                _Border = "TB";
                _Size = 10;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "INV.NO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SL.NO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INTEREST%", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-AMT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DAYS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BAL-DAYS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INT-1", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INT-2", _Color, true, _Border, "R", "", _Size, false, 325, "", true);



                foreach (PayHistroyReport Dr in mList)
                {
                    iRow++; iCol = 1;
                    _Border = "";
                    _Bold = false;
                    _Color = Color.Black;

                    //if (Dr.rowtype.ToString() == "TOTAL" || Dr.rowtype.ToString() == "GRANDTOTAL")
                    //{
                    //    _Border = "TB";

                    //    _Bold = true;
                    //    _Color = Color.DarkBlue;
                    //}




                    if (all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.branch, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jvh_vrno, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jvh_type, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.sl_no, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Dr.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.intrest, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jv_debit, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Dr.cr_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cr_total, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.days, 0), _Color, _Bold, _Border, "L", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.balance, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.bal_days, 0), _Color, _Bold, _Border, "L", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.int1, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.int2, 0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                }
                WB.SaveXls(File_Name + ".xls");
            }
            if(!detail)
            {

                if(!all)
                {
                    WS.Columns[0].Width = 256;
                    WS.Columns[1].Width = 256 * 7;
                    WS.Columns[2].Width = 256 * 10;
                    WS.Columns[3].Width = 256 * 27;
                    WS.Columns[4].Width = 256 * 11;
                    WS.Columns[5].Width = 256 * 11;

                    WS.Columns[6].Width = 256 * 12;
                    WS.Columns[7].Width = 256 * 9;
                    WS.Columns[8].Width = 256 * 11;
                    WS.Columns[9].Width = 256 * 9;
                    WS.Columns[10].Width = 256 * 9;
                    WS.Columns[11].Width = 256 * 13;
                    WS.Columns[12].Width = 256 * 21;
                    WS.Columns[13].Width = 256 * 15;

                    WS.Columns[14].Width = 256 * 15;
                }
                else
                {
                    WS.Columns[0].Width = 256;
                    WS.Columns[1].Width = 256 * 10;
                    WS.Columns[2].Width = 256 * 7;
                    WS.Columns[3].Width = 256 * 10;
                    WS.Columns[4].Width = 256 * 27;
                    WS.Columns[5].Width = 256 * 11;
                    WS.Columns[6].Width = 256 * 11;

                    WS.Columns[7].Width = 256 * 12;
                    WS.Columns[8].Width = 256 * 9;
                    WS.Columns[9].Width = 256 * 11;
                    WS.Columns[10].Width = 256 * 9;
                    WS.Columns[11].Width = 256 * 9;
                    WS.Columns[12].Width = 256 * 13;
                    WS.Columns[13].Width = 256 * 21;
                    WS.Columns[14].Width = 256 * 15;

                    WS.Columns[15].Width = 256 * 15;
                }

               
               

                iRow = 1; iCol = 1;

                iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

                sTitle = "PAYMENT HISTORY REPORT FROM " + Lib.getFrontEndDate(from_date) + " TO " + Lib.getFrontEndDate(to_date);

                Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

                iCol = 1;
                _Color = Color.DarkBlue;
                _Border = "TB";
                _Size = 10;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "INV.NO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
          
          
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
          
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AMOUNT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-AMT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "PENDING", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-DAYS", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OVER-DUE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-LIMIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, _Border, "L", "", _Size, false, 325, "", true);



                foreach (PayHistroyReport Dr in mList)
                {
                    iRow++; iCol = 1;
                    _Border = "";
                    _Bold = false;
                    _Color = Color.Black;


                    if (all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, nvl(Dr.branch, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.jvh_vrno, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "#;(#);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Dr.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.acc_name, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr.jv_debit, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#",false);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Dr.xref_crdate, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, _Bold, _Border, "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.xref_amt,0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.status, ""), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.pending, 0), _Color, _Bold, _Border, "L", "", _Size, false, 325,"#;(#);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cust_crdays, 0), _Color, _Bold, _Border, "L", "", _Size, false, 325, "#;(#);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.overdue,0), _Color, _Bold, _Border, "L", "", _Size, false, 325, "#;(#);#", false);
                    Lib.WriteData(WS, iRow, iCol++, nvl(Dr.cust_crlimit,0), _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Dr.sman_name, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                }
                WB.SaveXls(File_Name + ".xls");


            }

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

