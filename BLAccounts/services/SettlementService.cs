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
    public class SettlementService : BL_Base
    {
       
        public Dictionary<string, object> GetSettlementList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();


            Con_Oracle = new DBConnection();


            string jvhid = "";
            if (SearchData.ContainsKey("jvhid"))
                jvhid = SearchData["jvhid"].ToString();

            string jvid = "";
            if (SearchData.ContainsKey("jvid"))
                jvid = SearchData["jvid"].ToString();

            string accid = SearchData["accid"].ToString();
            string stype = SearchData["type"].ToString();
            string comp_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            List<PendingList> mList = new List<PendingList>();
            PendingList cRow;

            StringBuilder sql = new StringBuilder();


            try
            {
                // This Is For Creditors
                
                
                    sql.Append(" select * from  ( ");
                    sql.Append(" select 'CR' as drcr, jv_parent_id, jv_pkid,jv_acc_id,jvh_year,jvh_Vrno,jvh_docno, jvh_type, jvh_date, ");
                    sql.Append(" jvh_reference,jv_credit as Amount,  ");
                    sql.Append(" jv_credit - nvl(xref_Amt,0) as Balance,");
                    sql.Append(" nvl(Xref_Allocated_Amt,0) as Allocation");
                    sql.Append(" from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id");
                    sql.Append(" left join ");
                    sql.Append(" (	select 	 xref_cr_jv_id,");
                    sql.Append("	sum(xref_amt) as xref_Amt,	");
                    sql.Append("    sum(0) as xref_Allocated_Amt");
                    sql.Append(" from	 ledgerxref x");
                    sql.Append(" where 	 x.xref_Acc_id =  '{ACCOUNT}' and ");
                    sql.Append("         x.rec_company_Code= '{COMPANY}' and x.rec_branch_code= '{BRANCH}'");
                    sql.Append(" group by xref_cr_jv_id");
                    sql.Append(" )  b");
                    sql.Append(" on b.jv_pkid = b.xref_cr_jv_id");
                    sql.Append(" where b.jv_acc_id = '{ACCOUNT}' and a.jvh_type not in ('OP','OB', 'OC') and ");
                    sql.Append("       a.rec_branch_code= '{BRANCH}' and b.jv_credit > 0");


                

                    sql.Append(" union all ");

                    
                    sql.Append(" select 'DR' as drcr, jv_parent_id, jv_pkid,jv_acc_id,jvh_year,jvh_Vrno, jvh_docno, jvh_type, jvh_date, ");
                    sql.Append(" jvh_reference,jv_debit as Amount,  ");
                    sql.Append(" jv_debit - nvl(xref_Amt,0) as Balance,");
                    sql.Append(" nvl(Xref_Allocated_Amt,0) as Allocation");
                    sql.Append(" from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id");
                    sql.Append(" left join ");
                    sql.Append(" (	select 	 xref_dr_jv_id,");
                    sql.Append("	sum(xref_amt) as xref_Amt,	");
                    sql.Append("    sum(0) as xref_Allocated_Amt");
                    sql.Append(" from	 ledgerxref x");
                    sql.Append(" where 	 x.xref_Acc_id =  '{ACCOUNT}' and ");
                    sql.Append("         x.rec_company_Code= '{COMPANY}' and x.rec_branch_code= '{BRANCH}'");
                    sql.Append(" group by xref_dr_jv_id");
                    sql.Append(" )  b");
                    sql.Append(" on b.jv_pkid = b.xref_dr_jv_id");
                    sql.Append(" where b.jv_acc_id = '{ACCOUNT}' and a.jvh_type not in ('OP','OB', 'OC') and ");
                    sql.Append("       a.rec_branch_code= '{BRANCH}' and b.jv_debit > 0");
                    sql.Append(" )  jv");
                    sql.Append(" where (Balance) > 0  ");
                    sql.Append(" order by jvh_date,jvh_type, jvh_vrno");
                


                sql.Replace("{XREF_JV_PKID}", jvid);
                sql.Replace("{ACCOUNT}", accid);
                sql.Replace("{COMPANY}", comp_code);
                sql.Replace("{BRANCH}", branch_code);
                sql.Replace("{ENTITY_ID}", jvhid);

                DataTable Dt_Rec = new DataTable();
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
                    cRow.jv_allocation = Lib.Conv2Decimal(Dr["allocation"].ToString());
                    mList.Add(cRow);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("list", mList);
            return RetData;
        }


    }
}


