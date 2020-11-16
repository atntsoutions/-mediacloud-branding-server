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
    public class AllocationService : BL_Base
    {


        public Dictionary<string, object> GetPendingList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();


            Con_Oracle = new DBConnection();

            string accid = SearchData["accid"].ToString();
            string stype = SearchData["type"].ToString();
            string comp_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();


            List<PendingList> sList = new List<PendingList>();

            List<PendingList> mList = new List<PendingList>();
            PendingList cRow;

            DataTable Dt_Rec = new DataTable();


            StringBuilder sql = new StringBuilder();

            string sql1 = "";

            try
            {
                sql1 = "";
                // Find Source List for debtors credit amount and creditors debit amount
                // This Is For Creditors 
                if (stype == "DR")
                {
                    sql1 += " select case when jv_debit >0 then 'DR' else 'CR' end as drcr, b.jv_parent_id, b.jv_pkid,b.jv_acc_id,a.jvh_date,a.jvh_year, ";
                    sql1 += " a.jvh_vrno,a.jvh_docno, a.jvh_type, a.jvh_reference,jvh_narration,";
                    sql1 += " jv_debit as Amount, xref_amt, (jv_debit - nvl(xref_amt,0)) as balance ";
                    sql1 += " from ledgerh a inner join ledgert b on jvh_pkid = b.jv_parent_id left join ";
                    sql1 += " (select x.xref_dr_jv_id,sum(xref_amt) as xref_amt from ledgerxref x where ";
                    sql1 += " x.xref_acc_id ='" + accid + "' and x.rec_branch_code = '" + branch_code + "' ";
                    sql1 += " group by x.xref_dr_jv_id ";
                    sql1 += " ) xref on ( b.jv_pkid = xref.xref_dr_jv_id )";
                    sql1 += " where  b.jv_acc_id ='" + accid + "' and a.rec_branch_code = '" + branch_code + "' and ";
                    sql1 += " b.jv_debit > 0 and  (b.jv_debit - nvl(xref_amt,0)) >0 and a.jvh_type not in('OP','OB','OC')";
                    sql1 += " order by a.jvh_date,a.jvh_type";


                }
                // This Is For Debtros
                if (stype == "CR")
                {


                    sql1 += " select case when jv_debit >0 then 'DR' else 'CR' end as drcr,  b.jv_parent_id, b.jv_pkid,b.jv_acc_id,a.jvh_date,a.jvh_year, ";
                    sql1 += " a.jvh_vrno,a.jvh_docno, a.jvh_type, a.jvh_reference,jvh_narration,";
                    sql1 += " jv_credit as amount, xref_amt, (jv_credit - nvl(xref_amt,0)) as balance ";
                    sql1 += " from ledgerh a inner join ledgert b on jvh_pkid = b.jv_parent_id left join ";
                    sql1 += " (select x.xref_cr_jv_id,sum(xref_amt) as xref_amt from ledgerxref x where ";
                    sql1 += " x.xref_acc_id ='" + accid + "' and x.rec_branch_code = '" + branch_code + "' ";
                    sql1 += " group by x.xref_cr_jv_id ";
                    sql1 += " ) xref on ( b.jv_pkid = xref.xref_cr_jv_id )";
                    sql1 += " where  b.jv_acc_id ='" + accid + "' and a.rec_branch_code = '" + branch_code + "' and ";
                    sql1 += " b.jv_credit > 0 and  (b.jv_credit - nvl(xref_amt,0)) >0 and a.jvh_type not in('OP','OB','OC') ";
                    sql1 += " order by a.jvh_date,a.jvh_type";



                }

                sql1.Replace("{ACCOUNT}", accid);
                sql1.Replace("{COMPANY}", comp_code);
                sql1.Replace("{BRANCH}", branch_code);

                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql1.ToString());
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    cRow = new PendingList();
                    cRow.jv_parent_id = Dr["jv_parent_id"].ToString();
                    cRow.jv_pkid = Dr["jv_pkid"].ToString();
                    cRow.jv_drcr = Dr["drcr"].ToString();
                    cRow.jv_acc_id = Dr["jv_acc_id"].ToString();
                    cRow.jv_year = Lib.Conv2Integer(Dr["jvh_year"].ToString());
                    cRow.jv_reference = Dr["jvh_reference"].ToString();
                    cRow.jv_vrno = Dr["jvh_vrno"].ToString();
                    cRow.jv_docno = Dr["jvh_docno"].ToString();
                    cRow.jv_type = Dr["jvh_type"].ToString();
                    cRow.jv_display_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    cRow.jv_date = Lib.DatetoString(Dr["jvh_date"]);
                    cRow.jv_total = Lib.Conv2Decimal(Dr["amount"].ToString());
                    cRow.jv_balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    cRow.jv_allocation = 0;
                    sList.Add(cRow);
                }
                Dt_Rec.Rows.Clear();


                // Pending List for debtors debit balance and creditors credit balance
                // This Is For Creditors
                if (stype == "DR")
                {
                    sql.Append(" select * from  ( ");
                    sql.Append(" select 'CR' as drcr, jv_parent_id, jv_pkid,jv_acc_id,jvh_year,jvh_Vrno,jvh_docno, jvh_type, jvh_date, ");
                    sql.Append(" jvh_reference,jv_credit as Amount,  ");
                    sql.Append(" jv_credit - nvl(xref_Amt,0) as Balance,");
                    sql.Append(" 0 as Allocation");
                    sql.Append(" from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id");
                    sql.Append(" left join ");
                    sql.Append(" (	select 	 xref_cr_jv_id,");
                    sql.Append("	sum(xref_amt) as xref_Amt	");
                    sql.Append("    from	 ledgerxref x");
                    sql.Append("    where 	 x.xref_acc_id =  '{ACCOUNT}' and ");
                    sql.Append("    x.rec_company_code= '{COMPANY}' and x.rec_branch_code= '{BRANCH}'");
                    sql.Append("    group by xref_cr_jv_id");
                    sql.Append(" )  b");
                    sql.Append(" on b.jv_pkid = b.xref_cr_jv_id");
                    sql.Append(" where b.jv_acc_id = '{ACCOUNT}' and a.jvh_type not in ('OP','OB', 'OC') and ");
                    sql.Append("       a.rec_branch_code= '{BRANCH}' and b.jv_credit > 0");
                    sql.Append(" )  jv");
                    sql.Append(" where (Balance) > 0  ");
                    sql.Append(" order by jvh_date,jvh_type, jvh_vrno");
                }
                // This Is For Debtros
                if (stype == "CR")
                {
                    sql.Append(" select * from  ( ");
                    sql.Append(" select 'DR'as drcr, jv_parent_id, jv_pkid,jv_acc_id,jvh_year,jvh_Vrno, jvh_docno, jvh_type, jvh_date, ");
                    sql.Append(" jvh_reference,jv_debit as Amount,  ");
                    sql.Append(" jv_debit - nvl(xref_Amt,0) as Balance,");
                    sql.Append(" 0 as Allocation");
                    sql.Append(" from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id");
                    sql.Append(" left join ");
                    sql.Append(" (	select 	 xref_dr_jv_id,");
                    sql.Append("	sum(xref_amt) as xref_Amt	");
                    sql.Append("    from ledgerxref x");
                    sql.Append("    where  x.xref_acc_id =  '{ACCOUNT}' and ");
                    sql.Append("    x.rec_company_code= '{COMPANY}' and x.rec_branch_code= '{BRANCH}'");
                    sql.Append("    group by xref_dr_jv_id");
                    sql.Append(" )  b");
                    sql.Append(" on b.jv_pkid = b.xref_dr_jv_id");
                    sql.Append(" where b.jv_acc_id = '{ACCOUNT}' and a.jvh_type not in ('OP','OB', 'OC') and ");
                    sql.Append("       a.rec_branch_code= '{BRANCH}' and b.jv_debit > 0");
                    sql.Append(" )  jv");
                    sql.Append(" where (Balance) > 0  ");
                    sql.Append(" order by jvh_date,jvh_type, jvh_vrno");
                }


                sql.Replace("{ACCOUNT}", accid);
                sql.Replace("{COMPANY}", comp_code);
                sql.Replace("{BRANCH}", branch_code);

                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql.ToString());
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    cRow = new PendingList();
                    cRow.jv_parent_id = Dr["jv_parent_id"].ToString();
                    cRow.jv_pkid = Dr["jv_pkid"].ToString();
                    cRow.jv_drcr = Dr["drcr"].ToString();
                    cRow.jv_acc_id = Dr["jv_acc_id"].ToString();
                    cRow.jv_year = Lib.Conv2Integer(Dr["jvh_year"].ToString());
                    cRow.jv_reference = Dr["jvh_reference"].ToString();
                    cRow.jv_vrno = Dr["jvh_vrno"].ToString();
                    cRow.jv_docno = Dr["jvh_docno"].ToString();
                    cRow.jv_type = Dr["jvh_type"].ToString();
                    cRow.jv_display_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    cRow.jv_date = Lib.DatetoString(Dr["jvh_date"]);
                    cRow.jv_total = Lib.Conv2Decimal(Dr["amount"].ToString());
                    cRow.jv_balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    cRow.jv_allocation = 0;
                    mList.Add(cRow);
                }

                Dt_Rec.Rows.Clear();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("sourcelist", sList);
            RetData.Add("pendinglist", mList);

            return RetData;
        }

        public Dictionary<string, object> Save(XList Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            DBRecord Rec;
            try
            {
                Con_Oracle = new DBConnection();
                Con_Oracle.BeginTransaction();
                int iCtr = 0;
                foreach (LedgerXref Row in Record.XrefList)
                {
                    iCtr++;
                    Rec = new DBRecord();
                    Rec.CreateRow("LedgerXref", "ADD", "xref_pkid", Row.xref_pkid);

                    Rec.InsertString("xref_jvh_id", Row.xref_jvh_id);
                    Rec.InsertNumeric("xref_year", Row.xref_cr_jv_year.ToString());
                    
                    Rec.InsertString("xref_jv_id", Row.xref_jv_id.ToString());
                    Rec.InsertString("xref_acc_id", Row.xref_acc_id.ToString());
                    Rec.InsertString("xref_drcr", Row.xref_drcr.ToString());

                    Rec.InsertString("xref_dr_jvh_id", Row.xref_dr_jvh_id.ToString());
                    Rec.InsertString("xref_dr_jv_id", Row.xref_dr_jv_id.ToString());
                    Rec.InsertNumeric("xref_dr_jv_year", Row.xref_dr_jv_year.ToString());
                    Rec.InsertDate("xref_dr_jv_date", Row.xref_dr_jv_date);

                    Rec.InsertString("xref_cr_jvh_id", Row.xref_cr_jvh_id.ToString());
                    Rec.InsertString("xref_cr_jv_id", Row.xref_cr_jv_id.ToString());
                    Rec.InsertNumeric("xref_cr_jv_year", Row.xref_cr_jv_year.ToString());

                    Rec.InsertDate("xref_cr_jv_date", Row.xref_cr_jv_date);
                    Rec.InsertNumeric("xref_amt", Row.xref_amt.ToString());
                    Rec.InsertNumeric("xref_adv_amt", Row.xref_adv_amt.ToString());

                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);

                    sql = Rec.UpdateRow();
                    Con_Oracle.ExecuteNonQuery(sql);
                }
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch ( Exception Ex)
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


