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
    public class ReconService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_List_Reconciled = new DataTable();
        DataTable Dt_Bank = new DataTable();


        string reconciled = "";
        string unreconciled = "";

        ExcelFile WB;
        ExcelWorksheet WS = null;

        int iRow = 0;
        int iCol = 0;

        //string type = "";
        //string subtype = "";
        //string report_folder = "";
        string File_Name = "";
        string PKID = "";
        string acc_id = "";
        string acc_name = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string from_date = "";
        string to_date = "";
        Boolean ismaincode = false;
        string user_code = "";


        Boolean basedonreconcileddate = false;

        string report_folder = "";

        public IDictionary<string, object> UpdateRecon(Dictionary<string, object> SearchData)
        {
            string pkid = SearchData["pkid"].ToString();
            string inputdate = SearchData["inputdate"].ToString();
            string user_code = SearchData["user_code"].ToString();
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string oDate = "";

            try
            {

                Con_Oracle = new DBConnection();

                if (inputdate != "")
                {
                    DateTime dt = Convert.ToDateTime(inputdate);

                    sql = " select jvh_pkid ";
                    sql += " from ledgerh inner join ledgert on jvh_pkid = jv_parent_id ";
                    sql += " where jv_pkid = '" + pkid + "' and jvh_date >'" + dt.ToString("dd-MMM-yyyy") + "'";

                    if (Con_Oracle.IsRowExists(sql))
                        throw new Exception("Date Below Entry Date");

                    if (Convert.ToDateTime(inputdate) > System.DateTime.Today)
                        throw new Exception("Future Date Cannot Be Entered");
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("ledgert", "EDIT", "jv_pkid", pkid);
                Rec.InsertString("jv_recon_by", user_code);
                Rec.InsertDate("jv_recon_date", inputdate);


                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
                if (inputdate != "")
                    oDate = DateTime.Parse(inputdate).ToString("dd/MM/yyyy");

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                    throw Ex;
                }
            }
            RetData.Add("status", "OK");
            RetData.Add("displaydate", oDate);
            return RetData;
        }

        public IDictionary<string, object> UpdateOsRemarks(Dictionary<string, object> SearchData)
        {
            string pkid = SearchData["pkid"].ToString();
            string type = SearchData["type"].ToString();
            string remarks = SearchData["remarks"].ToString();
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            try
            {

                if (remarks.Length > 60)
                {
                    remarks = remarks.Substring(0, 60);
                }

                Con_Oracle = new DBConnection();

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("ledgert", "EDIT", "jv_pkid", pkid);
                Rec.InsertString("jv_od_type", type);
                Rec.InsertString("jv_od_remarks", remarks);


                sql = Rec.UpdateRow();

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
                    throw Ex;
                }
            }
            RetData.Add("status", "OK");
            return RetData;
        }

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {


            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string sWhere = "";
            Con_Oracle = new DBConnection();
            List<Recon> mList = new List<Recon>();
            Recon mrow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            report_folder = SearchData["report_folder"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            string fromdate = SearchData["from_date"].ToString();
            string todate = SearchData["to_date"].ToString();




            basedonreconcileddate = (Boolean)SearchData["basedonreconcileddate"];

            reconciled = SearchData["reconciled"].ToString();
            unreconciled = SearchData["unreconciled"].ToString();

            PKID = SearchData["pkid"].ToString();
            acc_id = SearchData["acc_id"].ToString();
            acc_name = SearchData["acc_name"].ToString();
            company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            year_code = SearchData["year_code"].ToString();

            user_code = SearchData["user_code"].ToString();

            fromdate = Lib.StringToDate(fromdate);
            todate = Lib.StringToDate(todate);


            from_date = fromdate;
            to_date = todate;

            long startrow = 0;
            long endrow = 0;




            try
            {

                if (type != "EXCEL")
                {
                    sWhere = " where  1=1  ";

                    sWhere += " and (";

                    //sql += " jv_recon_date <='" + todate + "'";

                    if (basedonreconcileddate)
                        sWhere += " jv_recon_date >= '" + fromdate + "'  and jv_recon_date <='" + todate + "'";
                    else
                        sWhere += " jvh_date >= '" + fromdate + "'  and jvh_date <='" + todate + "'";

                    //changed jv_year cond
                    sWhere += " and jv_acc_id ='" + acc_id + "'";
                    sWhere += " and jvh_year <= " + year_code + "";
                    sWhere += " and a.rec_company_code = '" + company_code + "' and a.rec_branch_code = '" + branch_code + "'";
                    sWhere += " and jvh_type not in('OP','OC','OI') ";

                    if (searchstring != "")
                    {
                        sWhere += " and (";
                        sWhere += " jvh_docno like '%" + searchstring + "%' or jv_chqno like '%" + searchstring + "%'";
                        sWhere += "  )";
                    }

                    if (reconciled == "Y" && unreconciled == "N")
                        sWhere += " and b.jv_recon_date is not null  ";
                    if (unreconciled == "Y" && reconciled == "N")
                        sWhere += " and b.jv_recon_date is null  ";

                    sWhere += " and a.rec_deleted = 'N' ";
                    sWhere += " )";

                    if (type == "NEW")
                    {
                        sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM ledgerh a ";
                        sql += "inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += "inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                        sql += sWhere;
                        DataTable Dt_Temp = new DataTable();
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


                    DataTable Dt_List = new DataTable();
                    sql = "";
                    sql += " select * from ( ";
                    sql += " select  jv_pkid,";
                    sql += " jvh_vrno as jv_vrno,jvh_docno as jv_docno, jvh_year, jvh_date as jv_date, jvh_type as jv_type, jv_acc_id, ";
                    sql += "  jv_debit, jv_credit,jv_chqno,jv_due_date,jv_recon_date,jv_paid_to,jv_bank,jvh_narration as  jv_narration,nvl(jvh_not_over_chq,'N') as jvh_not_over_chq,";
                    sql += " row_number() over(order by jvh_date) rn ";
                    sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                    sql += " inner join acctm c on b.jv_acc_id = c.acc_pkid ";
                    sql += sWhere;
                    sql += ") a ";
                    sql += " where rn between {startrow} and {endrow}";
                    sql += " order by  jv_date ";


                    sql = sql.Replace("{startrow}", startrow.ToString());
                    sql = sql.Replace("{endrow}", endrow.ToString());

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new Recon();
                        mrow.jv_pkid = Dr["jv_pkid"].ToString();
                        mrow.recon_acc_pkid = Dr["jv_acc_id"].ToString();
                        mrow.recon_jv_vrno = Dr["jv_vrno"].ToString();
                        mrow.recon_jv_type = Dr["jv_type"].ToString();
                        mrow.recon_jv_docno = Dr["jv_docno"].ToString();
                        mrow.recon_jv_date = Lib.DatetoString(Dr["jv_date"]);
                        mrow.recon_jv_year = Dr["jvh_year"].ToString();
                        mrow.debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                        mrow.credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.recon_type = Dr["jv_type"].ToString();
                        mrow.recon_chqno = Dr["jv_chqno"].ToString();
                        mrow.recon_due_date = Dr["jv_due_date"].ToString();
                        mrow.recon_date = Lib.DatetoString(Dr["jv_recon_date"]);
                        mrow.recon_display_date = Lib.DatetoStringDisplayformat(Dr["jv_recon_date"]);
                        mrow.recon_jv_narration = Dr["jv_narration"].ToString();
                        mrow.recon_paid_to = Dr["jv_paid_to"].ToString();
                        mrow.recon_bank = Dr["jv_bank"].ToString();
                        mrow.jvh_not_over_chq = Dr["jvh_not_over_chq"].ToString();
                        mList.Add(mrow);

                    }
                }

                if (type == "EXCEL" && user_code == "")
                {

                    report_folder = System.IO.Path.Combine(report_folder, PKID);
                    File_Name = System.IO.Path.Combine(report_folder, PKID);


                    if (Lib.CreateFolder(report_folder))
                    {

                        sql = "";
                        sql += " select jvh_vrno ,jvh_docno , jvh_year, jvh_date , jvh_type ,  ";
                        sql += " jv_debit, jv_credit,jv_chqno,jv_due_date, jv_recon_date, jv_paid_to,jv_bank,jvh_narration,nvl(jvh_not_over_chq,'N') as jvh_not_over_chq ";
                        sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " where ";

                        if (basedonreconcileddate)
                            sql += " jv_recon_date <='" + todate + "'";
                        else
                            sql += " jvh_date <='" + todate + "'";

                        //changed jv_year cond

                        sql += " and jv_acc_id ='" + acc_id + "'";
                        sql += " and jvh_year <= " + year_code + "";
                        sql += " and a.rec_company_code = '" + company_code + "' and a.rec_branch_code = '" + branch_code + "'";
                        sql += " and  jvh_type not in('OP','OC','OI') and jv_recon_date is not null ";
                        if (basedonreconcileddate)
                            sql += " order by jvh_date, jvh_type, jvh_vrno ";
                        else
                            sql += " order by jvh_date, jvh_type, jvh_vrno ";

                        Dt_List_Reconciled = new DataTable();
                        Dt_List_Reconciled = Con_Oracle.ExecuteQuery(sql);


                        sql = "";
                        sql += " select jvh_vrno ,jvh_docno , jvh_year, jvh_date , jvh_type ,  ";
                        sql += " jv_debit, jv_credit,jv_chqno,jv_due_date, jv_recon_date,jv_paid_to,jv_bank,jvh_narration,nvl(jvh_not_over_chq,'N') as jvh_not_over_chq ";
                        sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
                        sql += " where ";

                        sql += " jvh_date <='" + todate + "'";

                        //changed jv_year cond

                        sql += " and jv_acc_id ='" + acc_id + "'";
                        sql += " and jvh_year <= " + year_code + "";
                        sql += " and a.rec_company_code = '" + company_code + "' and a.rec_branch_code = '" + branch_code + "'";
                        sql += " and jvh_type not in('OP','OC','OI') and jv_recon_date is null ";
                        sql += " order by jvh_date, jvh_type, jvh_vrno ";


                        Dt_List = new DataTable();
                        Dt_List = Con_Oracle.ExecuteQuery(sql);

                        Con_Oracle.CloseConnection();
                        ProcessExcelFile();

                        Dt_List = null;
                        Dt_List_Reconciled = null;
                    }
                }

                if (type == "EXCEL" && user_code != "")
                {

                    report_folder = System.IO.Path.Combine(report_folder, PKID);
                    File_Name = System.IO.Path.Combine(report_folder, PKID);

                    if (Lib.CreateFolder(report_folder))
                    {
                        mList = ProcessReport();
                    }
                }



            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);

            return RetData;

        }

        private List<Recon> ProcessReport()
        {
            string _Border = "";
            Color _Color = Color.Black;
            int _Size = 0;
            string sTitle = "";
            string sName = "Report";

            // Report
            List<Recon> mList = new List<Recon>();
            
            DataTable table_Rep = new DataTable();
            DataTable table_Tmp = new DataTable();
            string sql = "";
            decimal nBank = 0, nDebit = 0, nCredit = 0, nBalance = 0, nLedger = 0;
            // Report



            // Bank Balance
            sql = "";
            sql += " select ";
            sql += " sum(jv_Debit) as Debit, sum(jv_Credit) as Credit ";
            sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
            sql += " where jv_acc_id ='" + acc_id + "'";
            sql += " and jvh_year <= " + year_code + "";
            sql += " and a.rec_company_code = '" + company_code + "' and a.rec_branch_code = '" + branch_code + "'";
            sql += " and  jvh_type not in('OP','OC','OI') ";
            sql += " and  jv_recon_date  <='" + to_date + "' and  jv_recon_date is not null ";
            table_Tmp = Con_Oracle.ExecuteQuery(sql);
            if (table_Tmp.Rows.Count > 0)
            {
                nDebit = Lib.Convert2Decimal(table_Tmp.Rows[0]["DEBIT"].ToString());
                nCredit = Lib.Convert2Decimal(table_Tmp.Rows[0]["CREDIT"].ToString());
                nBank = nDebit - nCredit;
            }




            // Ledger Balance

            sql = "";
            sql += " select ";
            sql += " sum(jv_Debit) as Debit, sum(jv_Credit) as Credit ";
            sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";
            sql += " where jv_acc_id ='" + acc_id + "'";
            sql += " and jvh_year <= " + year_code + "";
            sql += " and a.rec_company_code = '" + company_code + "' and a.rec_branch_code = '" + branch_code + "'";
            sql += " and  jvh_type not in('OP','OC','OI') ";
            sql += " and  jvh_date  <='" + to_date + "'";
            table_Tmp = Con_Oracle.ExecuteQuery(sql);
            if (table_Tmp.Rows.Count > 0)
            {
                nDebit = Lib.Convert2Decimal(table_Tmp.Rows[0]["DEBIT"].ToString());
                nCredit = Lib.Convert2Decimal(table_Tmp.Rows[0]["CREDIT"].ToString());
                nLedger = nDebit - nCredit;
            }
            // Un-Reconciled Detail
            nDebit = 0; nCredit = 0;
            sql = "";

            sql += " select jvh_vrno ,jvh_docno , jvh_year, jvh_date , jvh_type ,jvh_not_over_chq , ";
            sql += " case when jv_debit >0 then 1 else  2 end as jv_order, ";
            sql += " jv_debit, jv_credit,jv_chqno,jv_due_date, jv_recon_date,jv_paid_to,jv_bank,jvh_narration,nvl(jvh_not_over_chq,'N') as jvh_not_over_chq, ";
            sql += " '' as rowtype, 0 as recon_amount, '' as recon_amount_type ";
            sql += " from ledgerh a inner join ledgert b on jvh_pkid = jv_parent_id ";

            sql += " where jv_acc_id ='" + acc_id + "'";
            sql += " and jvh_year <= " + year_code + "";
            sql += " and a.rec_company_code = '" + company_code + "' and a.rec_branch_code = '" + branch_code + "'";
            sql += " and jvh_type not in('OP','OC','OI')  ";
            sql += " and JVH_DATE          <='" + to_date + "'";
            sql += " and (jv_recon_date is null or jv_recon_date  >'" + to_date + "' ) ";
            sql += " order by jv_order,JVH_DATE,JVH_VRNO, JVH_TYPE";
            table_Rep = Con_Oracle.ExecuteQuery(sql);

            foreach (DataRow Dr in table_Rep.Rows)
            {

                if (Lib.Convert2Decimal(Dr["jv_debit"].ToString()) > 0)
                {
                    Dr["recon_amount"] = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    Dr["recon_amount_type"] = "DR";
                    nDebit += Lib.Convert2Decimal(Dr["jv_debit"].ToString());
                    nBalance += nBalance;
                }
                else
                {
                    Dr["recon_amount"] = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                    Dr["recon_amount_type"] = "CR";
                    nCredit += Lib.Convert2Decimal(Dr["jv_credit"].ToString());
                    nBalance -= nBalance;
                }


            }



            DataRow row;


                       

            if (nDebit != 0)
            {
                row = table_Rep.NewRow();
                row["rowtype"] = "SUMMARY";
                row["jv_bank"] = "TOTAL UN-RECONCILED DEBIT";
                row["recon_amount"] = nDebit;
                row["recon_amount_type"] = "DR";
                table_Rep.Rows.Add(row);
            }
            if (nCredit != 0)
            {
                row = table_Rep.NewRow();
                row["rowtype"] = "SUMMARY";
                row["jv_bank"] = "TOTAL UN-RECONCILED CREDIT";
                row["recon_amount"] = nCredit;
                row["recon_amount_type"] = "CR";
                table_Rep.Rows.Add(row);
            }


            row = table_Rep.NewRow();
            row["rowtype"] = "SUMMARY";
            row["jv_bank"] = "BANK BOOK BALANCE ";
            row["recon_amount"] = Math.Abs(nBank);
            if (nBank > 0)
                row["recon_amount_type"] = "CR";
            else if (nBank < 0)
                row["recon_amount_type"] = "DR";
            table_Rep.Rows.Add(row);
            


            nBank += nDebit;
            nBank -= nCredit;

            if (nDebit != 0 || nCredit != 0)
            {
                row = table_Rep.NewRow();
                row["rowtype"] = "SUMMARY";
                row["jv_bank"] = "BANK BOOK BALANCE + UN-RECONCILED";
                row["recon_amount"] = Math.Abs(nBank);
                if (nBank > 0)
                    row["recon_amount_type"] = "DR";
                else
                    row["recon_amount_type"] = "CR";
                table_Rep.Rows.Add(row);
            }

            row = table_Rep.NewRow();
            row["rowtype"] = "SUMMARY";
            row["jv_bank"] = "LEDGER BALANCE";
            row["recon_amount"] = Math.Abs(nLedger);
            if (nLedger > 0)
                row["recon_amount_type"] = "DR";
            else
                row["recon_amount_type"] = "CR";
            table_Rep.Rows.Add(row);





            // Output to Excel

            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];
            WS.ViewOptions.ShowGridLines = false;
            WS.PrintOptions.Portrait = false;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;

            WS.Columns[0].Width = 256;
            WS.Columns[1].Width = 256 * 8;
            WS.Columns[2].Width = 256 * 6;
            WS.Columns[3].Width = 256 * 10;
            WS.Columns[4].Width = 256 * 20;
            WS.Columns[5].Width = 256 * 12;
            WS.Columns[6].Width = 256 * 10;
            WS.Columns[7].Width = 256 * 10;
            WS.Columns[8].Width = 256 * 10;
            WS.Columns[9].Width = 256 * 15;
            WS.Columns[10].Width = 256 * 6;
            WS.Columns[11].Width = 256 * 50;
            WS.Columns[12].Width = 256 * 100;

            WS.Columns[9].Style.NumberFormat = "#,0.00";

            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            sTitle = "BANK RECONCILE " + acc_name + " AS ON " + Lib.getFrontEndDate(to_date);

            //jvh_not_over_chq

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            // recocniled

            iCol = 1;
            _Color = Color.DarkBlue;
            _Border = "TB";
            _Size = 10;

            iRow++;
            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BANK", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "NOT-OVR-CHQ", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CHQ#", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DUE-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "RECNS-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "AMOUNT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "PAID TO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "NARRATION", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, _Border, "L", "", _Size, false, 325, "", true);

            foreach (DataRow Dr in table_Rep.Rows)
            {
                iRow++;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_vrno"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_type"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jvh_date"]), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_bank"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_not_over_chq"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, Dr["jv_chqno"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jv_due_date"]), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jv_recon_date"]) , _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["recon_amount"], _Color, false, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["recon_amount_type"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jv_paid_to"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_narration"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, _Border, "L", "", _Size, false, 325, "", true);

                if ( Dr["rowtype"].ToString() == "SUMMARY")
                {
                    Lib.WriteData(WS, iRow, 4, Dr["jv_bank"].ToString(), _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 5, DBNull.Value , _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 6, DBNull.Value, _Color, false, _Border, "L", "", _Size, false, 325, "", true);
                }
            }


            WB.SaveXls(File_Name + ".xls");

            table_Tmp = null;
            table_Rep = null;


            return mList;
        }


        private void ProcessExcelFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 0;

            decimal nTotal = 0;

            string sTitle = "";

            string sDate = "";

            decimal nUnRecBal = 0;
            decimal nRecBal = 0;
            decimal nDr = 0;
            decimal nCr = 0;

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];
            WS.ViewOptions.ShowGridLines = false;
            WS.PrintOptions.Portrait = false;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;

            WS.Columns[0].Width = 256;
            WS.Columns[1].Width = 256 * 10;
            WS.Columns[2].Width = 256 * 10;
            WS.Columns[3].Width = 256 * 15;
            WS.Columns[4].Width = 256 * 15;
            WS.Columns[5].Width = 256 * 15;
            WS.Columns[6].Width = 256 * 15;
            WS.Columns[7].Width = 256 * 15;
            WS.Columns[8].Width = 256 * 15;
            WS.Columns[9].Width = 256 * 15;
            WS.Columns[10].Width = 256 * 70;
            WS.Columns[11].Width = 256 * 50;

            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";

            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);

            sTitle = "BANK RECONCILE " + acc_name + " AS ON " + Lib.getFrontEndDate(to_date);

            if (basedonreconcileddate)
                sTitle += " (as per Reconciled Date)";

            Lib.WriteData(WS, iRow++, iCol, sTitle, Color.Brown, true, "", "L", "Calibri", 12, false);

            // recocniled
            nDr = 0; nCr = 0;
            if (reconciled == "Y")
            {
                iCol = 1;
                _Color = Color.DarkBlue;
                _Border = "TB";
                _Size = 10;

                iRow++;
                Lib.WriteData(WS, iRow++, iCol, "RECONCILED LIST", Color.Brown, true, "", "L", "Calibri", 12, false);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FIN-YEAR", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHQ#", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DUE-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RECNS-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NARRATION", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAID TO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            }
            nRecBal = 0;
            foreach (DataRow Dr in Dt_List_Reconciled.Rows)
            {

                _Border = "";
                _Bold = false;
                _Color = Color.Black;
                if (reconciled == "Y")
                {
                    iRow++; iCol = 1;
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_vrno"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_type"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_year"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jvh_date"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_chqno"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jv_due_date"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jv_recon_date"]), _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_narration"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_paid_to"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                }
                nRecBal += Lib.Convert2Decimal(Dr["jv_debit"].ToString());
                nRecBal -= Lib.Convert2Decimal(Dr["jv_credit"].ToString());
                nDr += Lib.Convert2Decimal(Dr["jv_debit"].ToString());
                nCr += Lib.Convert2Decimal(Dr["jv_credit"].ToString());
            }

            if (reconciled == "Y")
            {
                iRow++;
                if (nDr != 0 || nCr != 0)
                {
                    Lib.WriteData(WS, iRow, 5, "TOTAL", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 8, nDr, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 9, nCr, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    nTotal = nDr - nCr;
                    iRow++;
                    Lib.WriteData(WS, iRow, 6, "BALANCE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 8, nTotal, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                }
                iRow += 2;
            }

            // unrecocniled
            nDr = 0; nCr = 0; nUnRecBal = 0;
            if (unreconciled == "Y")
            {
                iCol = 1;
                _Color = Color.DarkBlue;
                _Border = "TB";
                _Size = 10;

                iRow++;
                Lib.WriteData(WS, iRow++, iCol, "UN-RECONCILED LIST", Color.Brown, true, "", "L", "Calibri", 12, false);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FIN-YEAR", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CHQ#", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DUE-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "RECNS-DATE", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, _Border, "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NARRATION", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAID TO", _Color, true, _Border, "L", "", _Size, false, 325, "", true);
            }

            foreach (DataRow Dr in Dt_List.Rows)
            {

                _Border = "";
                _Bold = false;
                _Color = Color.Black;
                if (unreconciled == "Y")
                {
                    iRow++; iCol = 1;
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_vrno"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_type"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_year"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jvh_date"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_chqno"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jv_due_date"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_narration"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jv_paid_to"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                }
                nUnRecBal += Lib.Convert2Decimal(Dr["jv_debit"].ToString());
                nUnRecBal -= Lib.Convert2Decimal(Dr["jv_credit"].ToString());
                nDr += Lib.Convert2Decimal(Dr["jv_debit"].ToString());
                nCr += Lib.Convert2Decimal(Dr["jv_credit"].ToString());
            }

            if (unreconciled == "Y")
            {
                iRow++;
                if (nDr != 0 || nCr != 0)
                {
                    Lib.WriteData(WS, iRow, 6, "TOTAL", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 8, nDr, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 9, nCr, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    nTotal = nDr - nCr;
                    iRow++;
                    Lib.WriteData(WS, iRow, 6, "BALANCE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, 8, nTotal, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                }
                iRow += 2;
            }


            _Bold = true;
            _Size = 12;


            Lib.WriteData(WS, iRow, 5, "BALANCE AS PER BANK BOOK", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 8, nRecBal, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            iRow += 1;
            Lib.WriteData(WS, iRow, 5, "UN-RECONCILED BALANCE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 8, nUnRecBal, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            nTotal = nRecBal + nUnRecBal;

            iRow += 1;
            Lib.WriteData(WS, iRow, 5, "LEDGER BALANCE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 8, nTotal, _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


            WB.SaveXls(File_Name + ".xls");
        }




    }
}

