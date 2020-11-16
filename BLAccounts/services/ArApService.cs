using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLAccounts
{
    public class ArApService : BL_Base
    {

        LovService lov = null;
        DataRow lovRow_cgst = null;
        DataRow lovRow_sgst = null;
        DataRow lovRow_igst = null;
        DataRow lovRow_cgst_rc_dr = null;
        DataRow lovRow_sgst_rc_dr = null;
        DataRow lovRow_igst_rc_dr = null;
        DataRow lovRow_cgst_rc_cr = null;
        DataRow lovRow_sgst_rc_cr = null;
        DataRow lovRow_igst_rc_cr = null;

        DataRow lovRow_Local_Currency = null;
        DataRow lovRow_Doc_Prefix = null;


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            Con_Oracle = new DBConnection();
            List<Ledgerh> mList = new List<Ledgerh>();
            Ledgerh mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();


            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();

            string cc_category = SearchData["cc_category"].ToString();
            string cc_id = SearchData["cc_id"].ToString();



            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where  1=1 ";


                sWhere += " and (";
                sWhere += " a.rec_company_code = '{COMP}'";
                sWhere += " and a.rec_branch_code = '{BRANCH}'";
                sWhere += " and a.jvh_year =  {FINYEAR}";
                if ( rowtype.Contains (","))
                    sWhere += " and a.jvh_type in( {ROWTYPE} )";
                else 
                sWhere += " and a.jvh_type =  '{ROWTYPE}'";

                sWhere += " ) ";

                if (searchstring != "")
                {
                    sWhere += " and ( ";
                    sWhere += "jvh_docno like '%" + searchstring + "%'";
                    sWhere += "or cc_code like '%" + searchstring + "%'";
                    sWhere += "or cc_name like '%" + searchstring + "%'";
                    sWhere += "or acc_name like '%" + searchstring + "%'";
                    sWhere += ")";
                }

                if ( cc_id != "")
                {
                    sWhere += " and jvh_cc_id = '" + cc_id + "'";
                }


                sWhere = sWhere.Replace("{COMP}", company_code);
                sWhere = sWhere.Replace("{BRANCH}", branch_code);
                sWhere = sWhere.Replace("{FINYEAR}", year_code);
                sWhere = sWhere.Replace("{ROWTYPE}", rowtype);


                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM Ledgerh  a ";
                    sql += " left join acctm b on a.jvh_acc_id = b.acc_pkid ";
                    sql += " left join costcenterm c on a.jvh_cc_id = c.cc_pkid ";

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
                sql += " select  jvh_pkid,jvh_vrno,jvh_docno, jvh_type,jvh_date, jvh_gstin,jvh_gst_type,acc_name, ";
                sql += " jvh_tot_amt, jvh_gst_amt,jvh_rc,jvh_net_amt,cc_type,cc_code, cc_name,";
                sql += " jvh_subtype,jvh_rec_source, jvh_narration,a.rec_aprvd_status,a.rec_aprvd_by,a.rec_aprvd_remark, ";
                sql += " row_number() over(order by jvh_vrno) rn ";
                sql += " from ledgerh a ";
                sql += " left join acctm b on a.jvh_acc_id = b.acc_pkid ";
                sql += " left join costcenterm c on a.jvh_cc_id = c.cc_pkid ";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by jvh_vrno";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Ledgerh();
                    mRow.jvh_pkid = Dr["jvh_pkid"].ToString();
                    mRow.jvh_docno = Dr["jvh_docno"].ToString();

                    mRow.jvh_type = Dr["jvh_type"].ToString();
                    mRow.jvh_subtype = Dr["jvh_subtype"].ToString();
                    mRow.jvh_rec_source = Dr["jvh_rec_source"].ToString();

                    mRow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);

                    mRow.jvh_cc_category = Dr["cc_type"].ToString();
                    mRow.jvh_cc_code = Dr["cc_code"].ToString();
                    mRow.jvh_cc_name = Dr["cc_name"].ToString();
                    mRow.jvh_acc_name = Dr["acc_name"].ToString();
                    mRow.jvh_gstin = Dr["jvh_gstin"].ToString();
                    mRow.jvh_gst_type = Dr["jvh_gst_type"].ToString();

                    mRow.jvh_narration = Dr["jvh_narration"].ToString();

                    mRow.jvh_rc = false;
                    if (Dr["jvh_rc"].ToString() == "Y")
                        mRow.jvh_rc = true;
                    mRow.jvh_tot_amt = Lib.Conv2Decimal(Dr["jvh_tot_amt"].ToString());
                    mRow.jvh_gst_amt = Lib.Conv2Decimal(Dr["jvh_gst_amt"].ToString());
                    mRow.jvh_net_amt = Lib.Conv2Decimal(Dr["jvh_net_amt"].ToString());
                    mRow.rec_aprvd_status = Dr["rec_aprvd_status"].ToString();
                    mRow.rec_aprvd_remark = Dr["rec_aprvd_remark"].ToString();
                    mRow.rec_aprvd_by = Dr["rec_aprvd_by"].ToString();

                    mList.Add(mRow);
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

        public Dictionary<string, object> GetPendingList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();


            Con_Oracle = new DBConnection();

            string jvhid = SearchData["jvhid"].ToString();
            string jvid = SearchData["jvid"].ToString();
            string accid = SearchData["accid"].ToString();
            string stype = SearchData["type"].ToString();
            string comp_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            List<PendingList> mList = new List<PendingList>();
            PendingList cRow;

            StringBuilder sql = new StringBuilder();

            try
            {


                if (stype == "DR")
                {
                    sql.Append(" select * from  ( ");
                    sql.Append(" select 'CR' as drcr, jv_parent_id, jv_pkid,jv_acc_id,jvh_year,jvh_Vrno, jvh_type, jvh_date, ");
                    sql.Append(" jvh_reference,jv_credit as Amount,  ");
                    sql.Append(" jv_credit - nvl(xref_Amt,0) as Balance,");
                    sql.Append(" nvl(Xref_Allocated_Amt,0) as Allocation");
                    sql.Append(" from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id");
                    sql.Append(" left join ");
                    sql.Append(" (	select 	 xref_cr_jv_id,");
                    sql.Append("	sum(case when xref_jv_id='{XREF_JV_PKID}' then 0        else xref_amt end ) as xref_Amt,	");
                    sql.Append("    sum(case when xref_jv_id='{XREF_JV_PKID}' then xref_amt else 0 end ) as xref_Allocated_Amt");
                    sql.Append(" from	 ledgerxref x");
                    sql.Append(" where 	 x.xref_Acc_id =  '{ACCOUNT}' and ");
                    sql.Append("         x.rec_company_Code= '{COMPANY}' and x.rec_branch_code= '{BRANCH}'");
                    sql.Append(" group by xref_cr_jv_id");
                    sql.Append(" )  b");
                    sql.Append(" on b.jv_pkid = b.xref_cr_jv_id");
                    sql.Append(" where a.jvh_pkid <> '{ENTITY_ID}' and b.jv_acc_id = '{ACCOUNT}' and a.jvh_type not in ('OP','OB', 'OC') and ");
                    sql.Append("       a.rec_branch_code= '{BRANCH}' and b.jv_credit > 0");
                    sql.Append(" )  jv");
                    sql.Append(" where (Balance) > 0  ");
                    sql.Append(" order by jvh_date,jvh_type, jvh_vrno");
                }

                if (stype == "CR")
                {
                    sql.Append(" select * from  ( ");
                    sql.Append(" select 'DR'as drcr, jv_parent_id, jv_pkid,jv_acc_id,jvh_year,jvh_Vrno, jvh_type, jvh_date, ");
                    sql.Append(" jvh_reference,jv_debit as Amount,  ");
                    sql.Append(" jv_debit - nvl(xref_Amt,0) as Balance,");
                    sql.Append(" nvl(Xref_Allocated_Amt,0) as Allocation");
                    sql.Append(" from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id");
                    sql.Append(" left join ");
                    sql.Append(" (	select 	 xref_dr_jv_id,");
                    sql.Append("	sum(case when xref_jv_id='{XREF_JV_PKID}' then 0        else xref_amt end ) as xref_Amt,	");
                    sql.Append("    sum(case when xref_jv_id='{XREF_JV_PKID}' then xref_amt else 0 end ) as xref_Allocated_Amt");
                    sql.Append(" from	 ledgerxref x");
                    sql.Append(" where 	 x.xref_Acc_id =  '{ACCOUNT}' and ");
                    sql.Append("         x.rec_company_Code= '{COMPANY}' and x.rec_branch_code= '{BRANCH}'");
                    sql.Append(" group by xref_dr_jv_id");
                    sql.Append(" )  b");
                    sql.Append(" on b.jv_pkid = b.xref_dr_jv_id");
                    sql.Append(" where a.jvh_pkid <> '{ENTITY_ID}' and b.jv_acc_id = '{ACCOUNT}' and a.jvh_type not in ('OP','OB', 'OC') and ");
                    sql.Append("       a.rec_branch_code= '{BRANCH}' and b.jv_debit > 0");
                    sql.Append(" )  jv");
                    sql.Append(" where (Balance) > 0  ");
                    sql.Append(" order by jvh_date,jvh_type, jvh_vrno");
                }



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
                    cRow.jv_type = Dr["jvh_type"].ToString();
                    cRow.jv_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
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

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();
            string lockedmsg = "";
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();


                sql = "";
                sql += " select ";
                sql += " jvh_pkid, jvh_date, jvh_year, jvh_type, jvh_vrno, jvh_docno, ";
                sql += " jvh_acc_id, acc_code as jvh_acc_code, acc_name as jvh_acc_name,";
                sql += " addr.add_line1||'\n'||addr.add_line2||'\n'||addr.add_line3 as  jvh_acc_br_addr,addr.add_email as jvh_acc_br_email,";
                sql += " jvh_gst ,jvh_gstin ,jvh_state_id,state.param_code as jvh_state_code,state.param_name as jvh_state_name, ";
                sql += " jvh_gst_type,jvh_org_invno, jvh_org_invdt,jvh_no_brok,jvh_brok_remarks,jvh_remarks,jvh_brok_amt,jvh_brok_per,jvh_basic_frt,";
                sql += " jvh_rc,jvh_sez,  jvh_is_export, jvh_igst_exception, jvh_exrate,jvh_sman_id ,jvh_reference ,jvh_reference_date,jvh_narration,a.rec_category, ";
                sql += " jvh_cgst_amt, jvh_sgst_amt, jvh_igst_amt, jvh_gst_amt,";
                sql += " jvh_debit, jvh_credit, jvh_curr_id, jvh_curr_code, jvh_acc_br_id, add_branch_slno as jvh_acc_br_slno, jvh_cc_category, ";
                sql += " jvh_cc_id, cc_code as jvh_cc_code,cc_name as jvh_cc_name, cc_year as jvh_cc_year, jvh_rec_source, ";
                sql += " jvh_edit_code, jvh_edit_date, a.rec_locked, a.rec_company_code, a.rec_branch_code,jvh_banktype,  ";
                sql += " jvh_brok_exrate,jvh_brok_amt_inr ";
                sql += " from ledgerh a left join acctm on jvh_acc_id = acc_pkid ";
                sql += " left join addressm addr on jvh_acc_br_id = add_pkid ";
                sql += " left join param state on jvh_state_id = state.param_pkid ";
                sql += " left join costcenterm cc on jvh_cc_id =cc.cc_pkid ";
                sql += " where  a.jvh_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Ledgerh();
                    mRow.jvh_pkid = Dr["jvh_pkid"].ToString();
                    mRow.jvh_vrno = Lib.Conv2Integer(Dr["jvh_vrno"].ToString());
                    mRow.jvh_docno = Dr["jvh_docno"].ToString();
                    mRow.jvh_type = Dr["jvh_type"].ToString();
                    mRow.jvh_date = Lib.DatetoString(Dr["jvh_date"]);

                    mRow.jvh_year = Lib.Conv2Integer(Dr["jvh_year"].ToString());

                    mRow.jvh_rec_source  = Dr["jvh_rec_source"].ToString();

                    mRow.jvh_reference = Dr["jvh_reference"].ToString();
                    mRow.jvh_reference_date = Lib.DatetoString(Dr["jvh_reference_date"]);

                    mRow.jvh_org_invno = Dr["jvh_org_invno"].ToString();
                    mRow.jvh_org_invdt = Lib.DatetoString(Dr["jvh_org_invdt"]);

                    mRow.jvh_acc_id = Dr["jvh_acc_id"].ToString();

                    mRow.jvh_acc_code = Dr["jvh_acc_code"].ToString();
                    mRow.jvh_acc_name = Dr["jvh_acc_name"].ToString();
                    mRow.jvh_acc_br_id = Dr["jvh_acc_br_id"].ToString();
                    mRow.jvh_acc_br_slno = Dr["jvh_acc_br_slno"].ToString();
                    mRow.jvh_acc_br_address = Dr["jvh_acc_br_addr"].ToString();
                    mRow.jvh_acc_br_email = Dr["jvh_acc_br_email"].ToString();

                    mRow.jvh_gstin = Dr["jvh_gstin"].ToString();
                    mRow.jvh_gst_type = Dr["jvh_gst_type"].ToString();


                    mRow.rec_locked = false;
                    mRow.jvh_edit_code = Dr["jvh_edit_code"].ToString();
                    mRow.jvh_edit_date = Dr["jvh_edit_date"].ToString();
                    if (Dr["rec_locked"].ToString() == "Y" && Dr["jvh_edit_date"].ToString() != System.DateTime.Today.ToString("yyyyMMdd"))
                    {
                        mRow.jvh_edit_code = "";
                        mRow.rec_locked = true;
                    }



                    mRow.jvh_gst = false;
                    if (Dr["jvh_gst"].ToString() == "Y")
                        mRow.jvh_gst = true;

                    mRow.jvh_rc = false;
                    if (Dr["jvh_rc"].ToString() == "Y")
                        mRow.jvh_rc = true;

                    mRow.jvh_sez = false;
                    if (Dr["jvh_sez"].ToString() == "Y")
                        mRow.jvh_sez = true;

                    mRow.jvh_is_export = false;
                    if (Dr["jvh_is_export"].ToString() == "Y")
                        mRow.jvh_is_export = true;

                    mRow.jvh_igst_exception = false;
                    if (Dr["jvh_igst_exception"].ToString() == "Y")
                        mRow.jvh_igst_exception  = true;


                    mRow.jvh_no_brok = false;
                    if (Dr["jvh_no_brok"].ToString() == "Y")
                        mRow.jvh_no_brok = true;

                    mRow.jvh_brok_amt = Lib.Conv2Decimal(Dr["jvh_brok_amt"].ToString());
                    mRow.jvh_brok_per = Lib.Conv2Decimal(Dr["jvh_brok_per"].ToString());
                    mRow.jvh_basic_frt = Lib.Conv2Decimal(Dr["jvh_basic_frt"].ToString());
                    mRow.jvh_brok_remarks = Dr["jvh_brok_remarks"].ToString();
                    mRow.jvh_remarks = Dr["jvh_remarks"].ToString();

                    mRow.jvh_state_id = Dr["jvh_state_id"].ToString();
                    mRow.jvh_state_code = Dr["jvh_state_code"].ToString();
                    mRow.jvh_state_name = Dr["jvh_state_name"].ToString();

                    mRow.jvh_curr_id = Dr["jvh_curr_id"].ToString();
                    mRow.jvh_curr_code = Dr["jvh_curr_code"].ToString();
                    mRow.jvh_exrate = Lib.Conv2Decimal(Dr["jvh_exrate"].ToString());

                    mRow.rec_category = Dr["rec_category"].ToString();
                    mRow.jvh_cc_category = Dr["jvh_cc_category"].ToString();
                    mRow.jvh_cc_id = Dr["jvh_cc_id"].ToString();
                    mRow.jvh_cc_code = Dr["jvh_cc_code"].ToString();
                    mRow.jvh_cc_name = Dr["jvh_cc_name"].ToString();
                    mRow.jvh_cc_year = Lib.Conv2Integer(Dr["jvh_cc_year"].ToString());

                    if (mRow.jvh_cc_year > 0 && mRow.jvh_cc_year != mRow.jvh_year)
                        mRow.jvh_cc_code = mRow.jvh_cc_code + "/" + mRow.jvh_cc_year.ToString();


                    mRow.jvh_narration = Dr["jvh_narration"].ToString();
                    mRow.jvh_debit = Lib.Conv2Decimal(Dr["jvh_debit"].ToString());
                    mRow.jvh_credit = Lib.Conv2Decimal(Dr["jvh_credit"].ToString());

                    mRow.jvh_cgst_amt = Lib.Conv2Decimal(Dr["jvh_cgst_amt"].ToString());
                    mRow.jvh_sgst_amt = Lib.Conv2Decimal(Dr["jvh_sgst_amt"].ToString());
                    mRow.jvh_igst_amt = Lib.Conv2Decimal(Dr["jvh_igst_amt"].ToString());
                    mRow.jvh_gst_amt = Lib.Conv2Decimal(Dr["jvh_gst_amt"].ToString());
                    if (Dr["jvh_banktype"].Equals(DBNull.Value))
                        mRow.jvh_banktype = "NA";
                    else
                        mRow.jvh_banktype = Dr["jvh_banktype"].ToString();

                    mRow.jvh_brok_exrate = Lib.Conv2Decimal(Dr["jvh_brok_exrate"].ToString());
                    mRow.jvh_brok_amt_inr = Lib.Conv2Decimal(Dr["jvh_brok_amt_inr"].ToString());


                    string JvhDate = Lib.StringToDate(Dr["jvh_date"]);
                    lockedmsg = Lib.IsDateLocked(JvhDate, Dr["jvh_type"].ToString(),
                        Dr["rec_company_code"].ToString(),
                        Dr["rec_branch_code"].ToString(), Dr["jvh_year"].ToString());

                    break;
                }

                List<Ledgert> mList = new List<Ledgert>();
                Ledgert aRow;

                sql = "select b.jv_pkid, b.jv_parent_id, ";
                sql += " b.jv_acc_id, c.acc_main_code as jv_acc_main_code, c.acc_code as jv_acc_code, b.jv_acc_name,act.actype_name, c.acc_against_invoice , c.acc_cost_centre, ";
                sql += " b.jv_curr_id, d.param_code as jv_curr_code,";
                sql += " b.jv_sac_id, sac.param_code as jv_sac_code,jv_is_taxable,";
                sql += " jv_qty,jv_rate,jv_ftotal, jv_exrate, jv_total, jv_taxable_amt, jv_drcr, jv_debit, jv_credit, ";
                sql += " jv_cgst_rate,jv_cgst_amt, jv_sgst_rate, jv_sgst_amt, jv_igst_rate, jv_igst_amt, jv_gst_amt,jv_gst_rate, ";
                sql += " jv_cntr_type_id, cntrtype.param_code as jv_cntr_type_code,";
                sql += " jv_net_total, jv_gst_type,jv_gst_edited,jv_is_gst_item, ";
                sql += " jv_doc_type, jv_bank, jv_branch,jv_chqno, jv_due_date, ";
                sql += " b.jv_pay_reason, b.jv_supp_docs, b.jv_paid_to, b.jv_remarks";
                sql += " from ledgert b ";
                sql += " left join acctm c on b.jv_acc_id = c.acc_pkid ";
                sql += " left join param d on b.jv_curr_id = d.param_pkid ";
                sql += " left join param sac on b.jv_sac_id = sac.param_pkid ";
                sql += " left join param cntrtype on b.jv_cntr_type_id = cntrtype.param_pkid ";
                sql += " left join actypem act on c.acc_type_id = act.actype_pkid ";
                sql += " where b.jv_parent_id ='{ID}' and nvl(jv_row_type,'JV') not in('HEADER','GST') ";
                sql += " order by jv_ctr ";

                sql = sql.Replace("{ID}", id);

                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    aRow = new Ledgert();
                    aRow.jv_pkid = Dr["jv_pkid"].ToString();
                    aRow.jv_parent_id = Dr["jv_parent_id"].ToString();

                    aRow.jv_acc_id = Dr["jv_acc_id"].ToString();
                    aRow.jv_acc_main_code = Dr["jv_acc_main_code"].ToString();
                    aRow.jv_acc_code = Dr["jv_acc_code"].ToString();
                    aRow.jv_acc_name = Dr["jv_acc_name"].ToString();
                    aRow.jv_acc_type_name = Dr["actype_name"].ToString();

                    aRow.jv_acc_against_invoice = Dr["acc_against_invoice"].ToString();
                    aRow.jv_acc_cost_centre = Dr["acc_cost_centre"].ToString();
                    if (Dr["jv_is_taxable"].ToString() == "Y")
                        aRow.jv_is_taxable = true;
                    else
                        aRow.jv_is_taxable = false;

                    if (Dr["jv_is_gst_item"].ToString() == "Y")
                        aRow.jv_is_gst_item = true;
                    else
                        aRow.jv_is_gst_item = false;



                    aRow.jv_cntr_type_id = Dr["jv_cntr_type_id"].ToString();
                    aRow.jv_cntr_type_code = Dr["jv_cntr_type_code"].ToString();

                    aRow.jv_curr_id = Dr["jv_curr_id"].ToString();
                    aRow.jv_curr_code = Dr["jv_curr_code"].ToString();

                    aRow.jv_sac_id = Dr["jv_sac_id"].ToString();
                    aRow.jv_sac_code = Dr["jv_sac_code"].ToString();

                    
                    if (Dr["jv_gst_edited"].ToString() == "Y")
                        aRow.jv_gst_edited = true;
                    else
                        aRow.jv_gst_edited = false;

                    aRow.jv_qty = Lib.Conv2Decimal(Dr["jv_qty"].ToString());
                    aRow.jv_rate = Lib.Conv2Decimal(Dr["jv_rate"].ToString());

                    aRow.jv_ftotal = Lib.Conv2Decimal(Dr["jv_ftotal"].ToString());
                    aRow.jv_exrate = Lib.Conv2Decimal(Dr["jv_exrate"].ToString());
                    aRow.jv_total = Lib.Conv2Decimal(Dr["jv_total"].ToString());
                    aRow.jv_taxable_amt = Lib.Conv2Decimal(Dr["jv_taxable_amt"].ToString());

                    aRow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    aRow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());

                    aRow.jv_cgst_rate = Lib.Conv2Decimal(Dr["jv_cgst_rate"].ToString());
                    aRow.jv_cgst_amt = Lib.Conv2Decimal(Dr["jv_cgst_amt"].ToString());

                    aRow.jv_sgst_rate = Lib.Conv2Decimal(Dr["jv_sgst_rate"].ToString());
                    aRow.jv_sgst_amt = Lib.Conv2Decimal(Dr["jv_sgst_amt"].ToString());

                    aRow.jv_igst_rate = Lib.Conv2Decimal(Dr["jv_igst_rate"].ToString());
                    aRow.jv_igst_amt = Lib.Conv2Decimal(Dr["jv_igst_amt"].ToString());

                    aRow.jv_gst_amt = Lib.Conv2Decimal(Dr["jv_gst_amt"].ToString());
                    aRow.jv_gst_rate = Lib.Conv2Decimal(Dr["jv_gst_rate"].ToString());

                    aRow.jv_net_total = Lib.Conv2Decimal(Dr["jv_net_total"].ToString());

                    aRow.jv_drcr = Dr["jv_drcr"].ToString();

                    aRow.jv_gst_type = Dr["jv_gst_type"].ToString();

                    aRow.jv_doc_type = Dr["jv_doc_type"].ToString();
                    aRow.jv_bank = Dr["jv_bank"].ToString();
                    aRow.jv_branch = Dr["jv_branch"].ToString();
                    aRow.jv_chqno = Lib.Conv2Integer(Dr["jv_chqno"].ToString());
                    aRow.jv_due_date = Lib.DatetoString(Dr["jv_due_date"]);

                    aRow.jv_pay_reason = Dr["jv_pay_reason"].ToString();
                    aRow.jv_supp_docs = Dr["jv_supp_docs"].ToString();
                    aRow.jv_paid_to = Dr["jv_paid_to"].ToString();
                    aRow.jv_remarks = Dr["jv_remarks"].ToString();

                    mList.Add(aRow);
                }
                mRow.LedgerList = mList;


                // Any Allocation Exists against this Record
                mRow.jvh_allocation_found = false;
                sql = "select xref_jvh_id from ledgerxref where (xref_dr_jvh_id = '{ID}' or xref_cr_jvh_id = '{ID}') and xref_jvh_id<> '{ID}'";
                sql = sql.Replace("{ID}", id);
                if (Con_Oracle.IsRowExists(sql))
                    mRow.jvh_allocation_found = true;

                List<CostCentert> mCostCenterList = new List<CostCentert>();
                CostCentert cRow;

                sql = "";
                sql += " select ct_jvh_id, ct_pkid, ct_jv_id, ct_acc_id, ct_year, ct_category, ct_cost_id, ct_cost_year, ";
                sql += " cc_code as ct_cost_code,cc_name as ct_cost_name ,ct_amount, ct_ctr ";
                sql += " from costcentert a inner join costcenterm b on a.ct_cost_id = b.cc_pkid ";
                sql += " where ct_jvh_id = '{ID}' and ct_type = 'M' ";
                sql += " order by ct_jv_id,ct_ctr ";
                sql = sql.Replace("{ID}", id);

                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    cRow = new CostCentert();
                    cRow.ct_pkid = Dr["ct_pkid"].ToString();
                    cRow.ct_jvh_id = Dr["ct_jvh_id"].ToString();
                    cRow.ct_jv_id = Dr["ct_jv_id"].ToString();
                    cRow.ct_acc_id = Dr["ct_acc_id"].ToString();
                    cRow.ct_year = Lib.Conv2Integer(Dr["ct_year"].ToString());
                    cRow.ct_category = Dr["ct_category"].ToString();
                    cRow.ct_cost_year = Lib.Conv2Integer(Dr["ct_cost_year"].ToString());
                    cRow.ct_cost_id = Dr["ct_cost_id"].ToString();
                    cRow.ct_cost_code = Dr["ct_cost_code"].ToString();
                    cRow.ct_cost_name = Dr["ct_cost_name"].ToString();
                    cRow.ct_amount = Lib.Conv2Decimal(Dr["ct_amount"].ToString());
                    cRow.ct_ctr = Lib.Conv2Integer(Dr["ct_ctr"].ToString());
                    mCostCenterList.Add(cRow);
                }
                mRow.CostCenterList = mCostCenterList;




                Con_Oracle.CloseConnection();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("lockedmsg", lockedmsg);
            RetData.Add("record", mRow);
            return RetData;
        }

        public string AllValid(Ledgerh Record)
        {
            string str = "";
            try
            {
                sql = "";
                /*
                if (Con_Oracle.IsRowExists(sql))
                    str += "|Code/Name Exists";
                */

                //Transaction Locking
                string jvhdate = Lib.StringToDate(Record.jvh_date.ToString());
                str += Lib.IsDateLocked(jvhdate, Record.jvh_type.ToString(),
                        Record._globalvariables.comp_code,
                        Record._globalvariables.branch_code, Record._globalvariables.year_code);

                if (Lib.IsValidGST(Record.jvh_gst, Record.jvh_gstin, Record.jvh_state_code, Record.jvh_igst_exception) == false)
                {
                    str += " | Invalid GST  (Invalid GST Number  or Mismatch between GST Number and State Code)";
                }

                if (Record.jvh_type == "PN") //Folder closed checking for payment invoice. changed on 27/08/2018
                {
                    sql = "";
                    sql += "select hbl_pkid from hblm a where ";
                    sql += " nvl(length(hbl_edit_code),0) = 0 ";
                    sql += " and hbl_pkid = '" + Record.jvh_cc_id + "'";
                    if (Con_Oracle.IsRowExists(sql))
                    {
                        str += " | Master Locked. ";
                    }


                    if (Record.jvh_cc_category == "MAWB AIR EXPORT" || Record.jvh_cc_category == "MBL SEA EXPORT" || Record.jvh_cc_category == "MAWB AIR IMPORT" || Record.jvh_cc_category == "MBL SEA IMPORT")
                    {
                        if (Record.rec_mode == "ADD")
                        {
                            if (Lib.IsJobLinked(Record.jvh_cc_id, "MASTER") == false)
                            {
                                str += " | House Not Linked ";
                            }
                        }

                        sql = "";
                        sql += "select hbl_bl_no from hblm where nvl(length(hbl_bl_no),0) > 0 and hbl_pkid =  '" + Record.jvh_cc_id + "'";
                        if (!Con_Oracle.IsRowExists(sql))
                        {
                            if (Record.jvh_cc_category.Contains("SEA"))
                                str += " | MBL Number Not Found ";
                            else
                                str += " | MAWB Number Not Found ";
                        }
                    }
                }

                if (Record.rec_mode == "ADD")
                {
                    if (Record.jvh_type == "IN")
                    {
                        sql = "";
                        sql += "select jvh_pkid  from ledgerh a where ";
                        sql += " rec_company_code = '" + Record._globalvariables.comp_code + "'";
                        sql += " and rec_branch_code = '" + Record._globalvariables.branch_code + "'";
                        sql += " and jvh_type ='" + Record.jvh_type + "'";
                        sql += " and jvh_year ='" + Record.jvh_year + "'";
                        sql += " and jvh_date > '" + Lib.StringToDate(Record.jvh_date) + "'";
                        if (Con_Oracle.IsRowExists(sql))
                        {
                            str += " | Back Dated Entry Not Possible ";
                        }
                    }

                    if (Record.jvh_cc_category == "SI AIR EXPORT" || Record.jvh_cc_category == "SI SEA EXPORT")
                        if (Lib.IsJobLinked(Record.jvh_cc_id, "HOUSE") == false)
                        {
                            str += " | Job Not Linked ";
                        }
                }
            

                if (!Lib.IsInFinYear(Record.jvh_date, Record._globalvariables.year_start_date, Record._globalvariables.year_end_date, true))
                {
                    str += " | Invalid Date (Future Date or Date not in Financial Year)";
                }

                sql = "";
                if (Record.jvh_acc_id.Trim().Length > 0 || Record.jvh_acc_br_id.Trim().Length > 0)
                {
                    sql = "select add_pkid from addressm where add_pkid = '" + Record.jvh_acc_br_id + "' and add_parent_id = '" + Record.jvh_acc_id + "'";
                    if (!Con_Oracle.IsRowExists(sql))
                    {
                        str += "| Invalid Party Code/Address";
                    }
                }

                lovRow_Local_Currency = lov.getSettings(Record._globalvariables.comp_code, "LOCAL-CURRENCY");
                lovRow_Doc_Prefix = lov.getSettings(Record._globalvariables.branch_code, "DOC-PREFIX");

                if (lovRow_Doc_Prefix == null)
                    str += "| Doc Prefix Not Found";

                if (Record.jvh_rc == false)
                {
                    if (Record.jvh_cgst_amt > 0 || Record.jvh_sgst_amt > 0)
                    {
                        lovRow_cgst = lov.getSettings(Record._globalvariables.comp_code, "CGST");
                        lovRow_sgst = lov.getSettings(Record._globalvariables.comp_code, "SGST");
                        if (lovRow_cgst == null)
                            str += "| CGST Code Not Found";
                        if (lovRow_sgst == null)
                            str += "| SGST Code Not Found";
                    }
                    if (Record.jvh_igst_amt > 0)
                    {
                        lovRow_igst = lov.getSettings(Record._globalvariables.comp_code, "IGST");
                        if (lovRow_igst == null)
                            str += "| IGST Code Not Found";
                    }
                }
                if (Record.jvh_rc == true)
                {
                    if (Record.jvh_cgst_amt > 0 || Record.jvh_sgst_amt > 0)
                    {
                        lovRow_cgst_rc_dr = lov.getSettings(Record._globalvariables.comp_code, "CGST-RC-DR");
                        lovRow_sgst_rc_dr = lov.getSettings(Record._globalvariables.comp_code, "SGST-RC-DR");
                        lovRow_cgst_rc_cr = lov.getSettings(Record._globalvariables.comp_code, "CGST-RC-CR");
                        lovRow_sgst_rc_cr = lov.getSettings(Record._globalvariables.comp_code, "SGST-RC-CR");

                        if (lovRow_cgst_rc_dr == null)
                            str += "| CGST RC DR Code Not Found";
                        if (lovRow_sgst_rc_dr == null)
                            str += "| SGST RC DR Code Not Found";
                        if (lovRow_cgst_rc_cr == null)
                            str += "| CGST RC CR Code Not Found";
                        if (lovRow_sgst_rc_cr == null)
                            str += "| SGST RC CR Code Not Found";
                    }
                    if (Record.jvh_igst_amt > 0)
                    {
                        lovRow_igst_rc_dr = lov.getSettings(Record._globalvariables.comp_code, "IGST-RC-DR");
                        lovRow_igst_rc_cr = lov.getSettings(Record._globalvariables.comp_code, "IGST-RC-CR");
                        if (lovRow_igst_rc_dr == null)
                            str += "| IGST RC DR Code Not Found";
                        if (lovRow_igst_rc_cr == null)
                            str += "| IGST RC CR Code Not Found";
                    }

                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(Ledgerh Record)
        {
            DataTable Dt_cc = new DataTable();

            DataTable Dt_Hbl = new DataTable();

            DataTable Dt_cc_jobs = new DataTable();
            DataTable Dt_cc_hbls = new DataTable();
            DataTable Dt_cc_cntr = new DataTable();


            decimal cc_amt = 0;

            int iCtr = 0;

            Boolean bOk = false;
            int iVrNo = 0;
            string DocNo = "";

            string rowtype = "";
            decimal nFAmt = 0;
            decimal nAmt = 0;
            string currid = "";

            decimal nGstRowAmt = 0;
            string gstid = "";
            string gstname = "";

            string gstrc_drid = "";
            string gstrc_drname = "";

            string gstrc_crid = "";
            string gstrc_crname = "";

            string doc_prefix = "";

            string gst_dr_or_cr = "";
            lov = new LovService();

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";


            string oldccid = "";
            try
            {
                Con_Oracle = new DBConnection();

                /*
                if (Record.cust_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");
                */

                ErrorMessage = AllValid(Record);

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if (!Lib.IsValidGst(Record.jvh_acc_br_id, Record.jvh_gstin))
                    throw new Exception("GSTIN should be same as master gstin");

                if (Record.jvh_cc_category != "NA")
                {
                    Dt_cc_jobs = Lib.getCCJOBS(Record.jvh_cc_category, Record.jvh_cc_id);
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                    Dt_cc_cntr = Lib.getCCCntrs(Record.jvh_cc_category,Record.jvh_cc_id);
                }

                if (lovRow_Doc_Prefix != null)
                    doc_prefix = lovRow_Doc_Prefix["name"].ToString();

                if (Record.rec_mode == "ADD")
                {
                    sql = "select nvl(max(jvh_vrno),1000) + 1 as jvh_vrno  from ledgerh where ";
                    sql += " rec_company_code = '" + Record._globalvariables.comp_code + "'";
                    sql += " and rec_branch_code = '" + Record._globalvariables.branch_code + "'";
                    sql += " and jvh_year = " + Record._globalvariables.year_code;
                    sql += " and jvh_type ='" + Record.jvh_type + "'";
                    iVrNo = Lib.Conv2Integer(Con_Oracle.ExecuteScalar(sql).ToString());
                    DocNo = Record.jvh_type + "/" + doc_prefix + "/" + Record._globalvariables.year_prefix + "/" + iVrNo.ToString();
                }
                else
                {
                    sql = "select jvh_cc_category,jvh_cc_id from ledgerh where jvh_pkid = '" + Record.jvh_pkid + "'";
                    DataTable dt_c1 = new DataTable();
                    dt_c1 = Con_Oracle.ExecuteQuery(sql);
                    foreach ( DataRow Dr in dt_c1.Rows)
                    {
                        if (Dr["jvh_cc_category"].ToString() != "EMPLOYEE"  && Dr["jvh_cc_id"].ToString() != "")
                        {
                            oldccid = Dr["jvh_cc_id"].ToString();
                            if (oldccid == Record.jvh_cc_id)
                                oldccid = "";
                        }
                    }
                }

                


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("Ledgerh", Record.rec_mode, "jvh_pkid", Record.jvh_pkid);
                if (Record.rec_mode == "ADD")
                {
                    Record.jvh_docno = DocNo.ToString();
                    Rec.InsertNumeric("jvh_vrno", iVrNo.ToString());
                    Rec.InsertString("jvh_docno", DocNo.ToString());
                    Rec.InsertString("jvh_type", Record.jvh_type);
                    Rec.InsertString("jvh_subtype", Record.jvh_subtype);
                    Rec.InsertString("jvh_posted", "Y");

                    if ( Record.jvh_rec_source == "")
                        Rec.InsertString("jvh_rec_source", "JV");
                    else
                        Rec.InsertString("jvh_rec_source", Record.jvh_rec_source);

                    Rec.InsertNumeric("jvh_year", Record._globalvariables.year_code);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);

                    Rec.InsertString("rec_locked", "N");
                    Rec.InsertString("rec_deleted", "N");
                    Rec.InsertString("jvh_edit_code", "{S}");
                    Rec.InsertString("jvh_edit_date", System.DateTime.Today.ToString("yyyyMMdd"));
                

                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "SYSDATE");
                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "SYSDATE");
                }


                Rec.InsertDate("jvh_date", Record.jvh_date);

                Rec.InsertString("jvh_reference", Record.jvh_reference);
                Rec.InsertDate("jvh_reference_date", Record.jvh_reference_date);

                Rec.InsertString("jvh_acc_id", Record.jvh_acc_id);
                Rec.InsertString("jvh_acc_br_id", Record.jvh_acc_br_id);

                Rec.InsertNumeric("jvh_version", "10");

                Rec.InsertString("jvh_gstin", Record.jvh_gstin);
                Rec.InsertString("jvh_state_id", Record.jvh_state_id);
                Rec.InsertString("jvh_gst_type", Record.jvh_gst_type);

                Rec.InsertString("jvh_org_invno", Record.jvh_org_invno);
                Rec.InsertDate("jvh_org_invdt", Record.jvh_org_invdt);

                if (Record.jvh_gst)
                    Rec.InsertString("jvh_gst", "Y");
                else
                    Rec.InsertString("jvh_gst", "N");

                if (Record.jvh_rc)
                    Rec.InsertString("jvh_rc", "Y");
                else
                    Rec.InsertString("jvh_rc", "N");

                if (Record.jvh_sez)
                    Rec.InsertString("jvh_sez", "Y");
                else
                    Rec.InsertString("jvh_sez", "N");

                if (Record.jvh_is_export)
                    Rec.InsertString("jvh_is_export", "Y");
                else
                    Rec.InsertString("jvh_is_export", "N");


                if (Record.jvh_igst_exception)
                    Rec.InsertString("jvh_igst_exception", "Y");
                else
                    Rec.InsertString("jvh_igst_exception", "N");



                if (Record.jvh_no_brok)
                    Rec.InsertString("jvh_no_brok", "Y");
                else
                    Rec.InsertString("jvh_no_brok", "N");

                Rec.InsertNumeric("jvh_basic_frt", Record.jvh_basic_frt.ToString());
                Rec.InsertNumeric("jvh_brok_per", Record.jvh_brok_per.ToString());
                Rec.InsertNumeric("jvh_brok_amt", Record.jvh_brok_amt.ToString());
                Rec.InsertString("jvh_brok_remarks", Record.jvh_brok_remarks);
                Rec.InsertString("jvh_remarks", Record.jvh_remarks);


                Rec.InsertString("jvh_curr_id", Record.jvh_curr_id);
                Rec.InsertString("jvh_curr_code", Record.jvh_curr_code);

                Rec.InsertNumeric("jvh_exrate", Record.jvh_exrate.ToString());


                Rec.InsertNumeric("jvh_tot_famt", Record.jvh_tot_famt.ToString());
                Rec.InsertNumeric("jvh_cgst_famt", Record.jvh_cgst_famt.ToString());
                Rec.InsertNumeric("jvh_sgst_famt", Record.jvh_sgst_famt.ToString());
                Rec.InsertNumeric("jvh_igst_famt", Record.jvh_igst_famt.ToString());
                Rec.InsertNumeric("jvh_gst_famt", Record.jvh_gst_famt.ToString());
                Rec.InsertNumeric("jvh_net_famt", Record.jvh_net_famt.ToString());

                Rec.InsertNumeric("jvh_tot_amt", Record.jvh_tot_amt.ToString());
                Rec.InsertNumeric("jvh_cgst_amt", Record.jvh_cgst_amt.ToString());
                Rec.InsertNumeric("jvh_sgst_amt", Record.jvh_sgst_amt.ToString());
                Rec.InsertNumeric("jvh_igst_amt", Record.jvh_igst_amt.ToString());
                Rec.InsertNumeric("jvh_gst_amt", Record.jvh_gst_amt.ToString());
                Rec.InsertNumeric("jvh_net_amt", Record.jvh_net_amt.ToString());

                Rec.InsertString("jvh_narration", Record.jvh_narration);
                Rec.InsertNumeric("jvh_debit", Record.jvh_debit.ToString());
                Rec.InsertNumeric("jvh_credit", Record.jvh_credit.ToString());

                

                if (Record.jvh_cc_category.Contains("SEA EXPORT"))
                    Rec.InsertString("rec_category", "SEA EXPORT");
                else if (Record.jvh_cc_category.Contains("AIR EXPORT"))
                    Rec.InsertString("rec_category", "AIR EXPORT");
                else if (Record.jvh_cc_category.Contains("SEA IMPORT"))
                    Rec.InsertString("rec_category", "SEA IMPORT");
                else if (Record.jvh_cc_category.Contains("AIR IMPORT"))
                    Rec.InsertString("rec_category", "AIR IMPORT");
                else if (Record.jvh_cc_category.Contains("GENERAL"))
                    Rec.InsertString("rec_category", "GENERAL");
                else
                    Rec.InsertString("rec_category", "OTHERS");


                Rec.InsertString("jvh_cc_category", Record.jvh_cc_category);
                Rec.InsertString("jvh_cc_id", Record.jvh_cc_id);
                Rec.InsertString("jvh_banktype", Record.jvh_banktype);
                Rec.InsertNumeric("jvh_brok_exrate", Record.jvh_brok_exrate.ToString());
                Rec.InsertNumeric("jvh_brok_amt_inr", Record.jvh_brok_amt_inr.ToString());

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);


                sql = "delete from  ledgert where jv_parent_id ='" + Record.jvh_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                decimal nDr = 0;
                decimal nCr = 0;

                if (Record.jvh_drcr == "DR")
                {
                    nDr = Record.jvh_net_amt;
                    nCr = 0;
                }
                if (Record.jvh_drcr == "CR")
                {
                    nDr = 0;
                    nCr = Record.jvh_net_amt;
                }

                iCtr++;
                rowtype = "HEADER";
                SaveLedgerRecord(
                "ADD",
                Record.jvh_pkid, Record.jvh_pkid, Record.jvh_acc_id, Record.jvh_acc_name, false,
                Record.jvh_curr_id, "", 1, Record.jvh_tot_famt , false, false,
                Record.jvh_drcr, Record.jvh_tot_famt, Record.jvh_exrate, Record.jvh_tot_amt, Record.jvh_tot_amt, "",
                0, 0, 0,0, 0, 0,0, 0,"",nDr, nCr, Record.jvh_net_amt,Record.jvh_net_famt, 0, 0, 0, 0, Record.jvh_net_famt,
                "", "", "",0, "","", "", "", "", "HEADER",
                Record._globalvariables.year_code, Record._globalvariables.comp_code, Record._globalvariables.branch_code, iCtr
                );


                gst_dr_or_cr = "";
                foreach (Ledgert Row in Record.LedgerList)
                {
                    iCtr++;
                    if (Row.jv_debit > 0)
                        rowtype = "DR-LEDGER";
                    else
                        rowtype = "CR-LEDGER";

                    SaveLedgerRecord(
                    "ADD",
                    Record.jvh_pkid, Row.jv_pkid, Row.jv_acc_id, Row.jv_acc_name, Row.jv_is_gst_item,
                    Row.jv_curr_id, Row.jv_sac_id, Row.jv_qty, Row.jv_rate, Row.jv_is_taxable, Row.jv_gst_edited,
                    Row.jv_drcr, Row.jv_ftotal, Row.jv_exrate, Row.jv_total, Row.jv_taxable_amt, Row.jv_gst_type,
                    Row.jv_cgst_rate, Row.jv_sgst_rate, Row.jv_igst_rate,
                    Row.jv_cgst_amt, Row.jv_sgst_amt, Row.jv_igst_amt, Row.jv_gst_amt,Row.jv_gst_rate,Row.jv_cntr_type_id,
                    Row.jv_debit, Row.jv_credit, Row.jv_net_total,
                    Row.jv_total_fc, Row.jv_cgst_famt, Row.jv_sgst_famt, Row.jv_igst_famt, Row.jv_gst_famt, Row.jv_net_ftotal,
                    Row.jv_doc_type, Row.jv_bank, Row.jv_branch,
                    Row.jv_chqno, Row.jv_due_date,
                    Row.jv_pay_reason, Row.jv_supp_docs, Row.jv_paid_to, Row.jv_remarks, rowtype,
                    Record._globalvariables.year_code, Record._globalvariables.comp_code, Record._globalvariables.branch_code, iCtr
                    );

                    if (Row.jv_cgst_amt != 0 || Row.jv_sgst_amt != 0 || Row.jv_igst_amt != 0)
                    {
                        if (Row.jv_debit > 0)
                            gst_dr_or_cr = "DR";
                        else
                            gst_dr_or_cr = "CR";
                    }
                }


                for (int k = 1; k <= 3; k++)
                {
                    bOk = false;
                    if (k == 1 && Record.jvh_igst_amt > 0)
                    {
                        bOk = true;
                        nAmt = Record.jvh_igst_amt;
                        nFAmt = Record.jvh_igst_amt;
                        nGstRowAmt = nAmt;
                        if (Record.jvh_rc)
                        {
                            gstrc_drid = lovRow_igst_rc_dr["id"].ToString();
                            gstrc_drname = lovRow_igst_rc_dr["name"].ToString();
                            gstrc_crid = lovRow_igst_rc_cr["id"].ToString();
                            gstrc_crname = lovRow_igst_rc_cr["name"].ToString();
                        }
                        else
                        {
                            gstid = lovRow_igst["id"].ToString();
                            gstname = lovRow_igst["name"].ToString();
                        }

                    }
                    if (k == 2 && Record.jvh_cgst_amt > 0)
                    {
                        bOk = true;
                        nAmt = Record.jvh_cgst_amt;
                        nFAmt = Record.jvh_cgst_amt;
                        nGstRowAmt = nAmt;
                        if (Record.jvh_rc)
                        {
                            gstrc_drid = lovRow_cgst_rc_dr["id"].ToString();
                            gstrc_drname = lovRow_cgst_rc_dr["name"].ToString();
                            gstrc_crid = lovRow_cgst_rc_cr["id"].ToString();
                            gstrc_crname = lovRow_cgst_rc_cr["name"].ToString();
                        }
                        else
                        {
                            gstid = lovRow_cgst["id"].ToString();
                            gstname = lovRow_cgst["name"].ToString();
                        }

                    }
                    if (k == 3 && Record.jvh_sgst_amt > 0)
                    {
                        bOk = true;
                        nAmt = Record.jvh_sgst_amt;
                        nFAmt = Record.jvh_sgst_amt;
                        nGstRowAmt = nAmt;
                        if (Record.jvh_rc)
                        {
                            gstrc_drid = lovRow_sgst_rc_dr["id"].ToString();
                            gstrc_drname = lovRow_sgst_rc_dr["name"].ToString();
                            gstrc_crid = lovRow_sgst_rc_cr["id"].ToString();
                            gstrc_crname = lovRow_sgst_rc_cr["name"].ToString();
                        }
                        else
                        {
                            gstid = lovRow_sgst["id"].ToString();
                            gstname = lovRow_sgst["name"].ToString();
                        }
                    }
                    if (bOk)
                    {
                        iCtr++;
                        if (lovRow_Local_Currency != null)
                            currid = lovRow_Local_Currency["ID"].ToString();
                        if (Record.jvh_exrate > 1)
                        {
                            nFAmt = nAmt / Record.jvh_exrate;
                            nFAmt = Lib.RoundNumber_Latest(nFAmt.ToString(), 2, true);
                        }
                        if (Record.jvh_rc)
                        {
                            iCtr++;
                            SaveLedgerRecord("ADD", Record.jvh_pkid, System.Guid.NewGuid().ToString().ToUpper(),
                                gstrc_drid, gstrc_drname, false, currid,
                                "", 1, nAmt, false, false, "DR", nAmt, 1,
                                nGstRowAmt, nGstRowAmt, "", 0, 0, 0,
                                0, 0, 0, 0,0,"", nGstRowAmt, nGstRowAmt,
                                nGstRowAmt, nFAmt, 0, 0, 0, 0, nFAmt,
                                "", "", "", 0, "", "",
                                "", "", "", "GST", Record._globalvariables.year_code, Record._globalvariables.comp_code,
                                Record._globalvariables.branch_code, iCtr
                            );
                            iCtr++;
                            SaveLedgerRecord("ADD", Record.jvh_pkid, System.Guid.NewGuid().ToString().ToUpper(),
                                gstrc_crid, gstrc_crname, false, currid,
                                "", 1, nAmt, false, false, "CR", nAmt, 1,
                                nGstRowAmt, nGstRowAmt, "", 0, 0, 0,
                                0, 0, 0, 0, 0, "", nGstRowAmt, nGstRowAmt,
                                nGstRowAmt, nFAmt, 0, 0, 0, 0, nFAmt,
                                "", "", "", 0, "", "",
                                "", "", "", "GST", Record._globalvariables.year_code, Record._globalvariables.comp_code,
                                Record._globalvariables.branch_code, iCtr
                            );
                        }
                        else
                        {
                            iCtr++;
                            SaveLedgerRecord("ADD", Record.jvh_pkid, System.Guid.NewGuid().ToString().ToUpper(),
                                gstid, gstname, false, currid,
                                "", 1, nAmt, false, false, gst_dr_or_cr, nAmt, 1,
                                nGstRowAmt, nGstRowAmt, "", 0, 0, 0,
                                0, 0, 0, 0, 0, "", nGstRowAmt, nGstRowAmt,
                                nGstRowAmt, nFAmt, 0, 0, 0, 0, nFAmt,
                                "", "", "", 0, "", "",
                                "", "", "", "GST", Record._globalvariables.year_code, Record._globalvariables.comp_code,
                                Record._globalvariables.branch_code, iCtr
                            );

                        }
                    }
                }


                #region CostCenter, LedgerXref

                sql = "delete from  costcentert where ct_jvh_id ='" + Record.jvh_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                if (Record.jvh_cc_category != "NA")
                {
                    if (Record.jvh_cc_category == "GENERAL JOB")
                    {
                        foreach (CostCentert Row in Record.CostCenterList)
                        {
                            iCtr++;
                            Rec.CreateRow("CostCentert", "ADD", "ct_pkid", Row.ct_pkid);
                            Rec.InsertString("ct_jvh_id", Record.jvh_pkid);

                            Rec.InsertNumeric("ct_year", Row.ct_year.ToString());
                            Rec.InsertString("ct_jv_id", Row.ct_jv_id);
                            Rec.InsertString("ct_acc_id", Row.ct_acc_id);
                            Rec.InsertString("ct_category", Row.ct_category);
                            Rec.InsertString("ct_cost_id", Row.ct_cost_id);
                            Rec.InsertNumeric("ct_cost_year", Row.ct_cost_year.ToString());
                            Rec.InsertNumeric("ct_amount", Row.ct_amount.ToString());

                            Rec.InsertString("ct_type", "M");
                            if (Row.ct_category == "CNTR SEA EXPORT")
                                Rec.InsertString("ct_posted", "N");
                            else
                                Rec.InsertString("ct_posted", "Y");
                            Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                            Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                            Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);

                            sql = Rec.UpdateRow();
                            Con_Oracle.ExecuteNonQuery(sql);
                        }

                        foreach (CostCentert Row in Record.CostCenterList)
                        {
                            if (Row.ct_category == "CNTR SEA EXPORT")
                            {
                                sql = " select hbl_id as id, count(*) over() as tot ";
                                sql += " from hblcontainer where hbl_cntr_id = '" + Row.ct_cost_id + "' ";
                                sql += " group by hbl_id ";
                                Dt_Hbl = new DataTable();
                                Dt_Hbl = Con_Oracle.ExecuteQuery(sql);
                                foreach (DataRow Dr in Dt_Hbl.Rows)
                                {
                                    iCtr++;
                                    cc_amt = Row.ct_amount / Lib.Conv2Integer(Dr["tot"].ToString());
                                    cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);

                                    Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                                    Rec.InsertString("ct_jvh_id", Record.jvh_pkid);
                                    Rec.InsertNumeric("ct_year", Row.ct_year.ToString());
                                    Rec.InsertString("ct_jv_id", Row.ct_jv_id);
                                    Rec.InsertString("ct_acc_id", Row.ct_acc_id);
                                    Rec.InsertString("ct_category", "SI SEA EXPORT");
                                    Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                                    Rec.InsertNumeric("ct_cost_year", Row.ct_cost_year.ToString());
                                    Rec.InsertNumeric("ct_amount", cc_amt.ToString());
                                    Rec.InsertString("ct_type", "S");
                                    Rec.InsertString("ct_posted", "Y");
                                    Rec.InsertString("ct_remarks", Row.ct_cost_code + " (" + Row.ct_amount.ToString() + "/" + Dr["tot"].ToString() + " HBL)");
                                    Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                                    sql = Rec.UpdateRow();
                                    Con_Oracle.ExecuteNonQuery(sql);
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (Ledgert Row in Record.LedgerList)
                        {
                            if (Lib.getCCType(Row.jv_acc_main_code) == "JOB WISE")
                            {
                                foreach (DataRow Dr in Dt_cc_jobs.Rows)
                                {
                                    iCtr++;
                                    cc_amt = Row.jv_total / Lib.Conv2Integer(Dr["tot"].ToString());
                                    cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                                    Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                                    Rec.InsertString("ct_jvh_id", Record.jvh_pkid);
                                    Rec.InsertNumeric("ct_year", Record.jvh_year.ToString());
                                    Rec.InsertString("ct_jv_id", Row.jv_pkid);
                                    Rec.InsertString("ct_acc_id", Row.jv_acc_id);
                                    Rec.InsertString("ct_category", Dr["cc_category"].ToString());
                                    Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                                    Rec.InsertNumeric("ct_cost_year", Row.jv_year.ToString());
                                    Rec.InsertNumeric("ct_amount", cc_amt.ToString());
                                    Rec.InsertString("ct_type", "M");
                                    Rec.InsertString("ct_posted", "Y");
                                    Rec.InsertString("ct_remarks", "");
                                    Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                                    sql = Rec.UpdateRow();
                                    Con_Oracle.ExecuteNonQuery(sql);
                                }
                            }
                            if (Lib.getCCType(Row.jv_acc_main_code) == "HBL WISE")
                            {

                                foreach (DataRow Drc in Dt_cc_cntr.Rows)
                                {
                                    iCtr++;

                                    cc_amt = Row.jv_total / Lib.Conv2Integer(Drc["tot"].ToString());
                                    cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);

                                    Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                                    Rec.InsertString("ct_jvh_id", Record.jvh_pkid);
                                    Rec.InsertNumeric("ct_year", Record.jvh_year.ToString());
                                    Rec.InsertString("ct_jv_id", Row.jv_pkid);
                                    Rec.InsertString("ct_acc_id", Row.jv_acc_id);
                                    Rec.InsertString("ct_category", Drc["cc_category"].ToString());
                                    Rec.InsertString("ct_cost_id", Drc["id"].ToString());
                                    Rec.InsertNumeric("ct_cost_year", Row.jv_year.ToString());
                                    Rec.InsertNumeric("ct_amount", cc_amt.ToString());
                                    Rec.InsertString("ct_type", "M");
                                    Rec.InsertString("ct_posted", "N");
                                    Rec.InsertString("ct_remarks", "");
                                    Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                                    sql = Rec.UpdateRow();
                                    Con_Oracle.ExecuteNonQuery(sql);
                                }

                                foreach (DataRow Dr in Dt_cc_hbls.Rows)
                                {
                                    iCtr++;

                                    cc_amt = Row.jv_total / Lib.Conv2Integer(Dr["tot"].ToString());
                                    cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);

                                    cc_amt = Lib.getCCAmt(Dr, Row.jv_total, cc_amt);

                                    Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                                    Rec.InsertString("ct_jvh_id", Record.jvh_pkid);
                                    Rec.InsertNumeric("ct_year", Record.jvh_year.ToString());
                                    Rec.InsertString("ct_jv_id", Row.jv_pkid);
                                    Rec.InsertString("ct_acc_id", Row.jv_acc_id);
                                    Rec.InsertString("ct_category", Dr["cc_category"].ToString());
                                    Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                                    Rec.InsertNumeric("ct_cost_year", Row.jv_year.ToString());
                                    Rec.InsertNumeric("ct_amount", cc_amt.ToString());
                                    Rec.InsertString("ct_type", "S");
                                    Rec.InsertString("ct_posted", "Y");
                                    Rec.InsertString("ct_remarks", "");
                                    Rec.InsertNumeric("ct_ctr", iCtr.ToString());
                                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                                    sql = Rec.UpdateRow();
                                    Con_Oracle.ExecuteNonQuery(sql);
                                }


                            }

                        }
                    }
                }

                //sql = "delete from  ledgerxref where xref_jvh_id ='" + Record.jvh_pkid + "'";
                //Con_Oracle.ExecuteNonQuery(sql);

                #endregion

                sql = "";
                sql += " select sum(jv_debit) - sum(jv_credit) as total  from ledgert ";
                sql += " where jv_parent_id ='" + Record.jvh_pkid + "'";
                sql += " group by jv_parent_id ";
                sql += " having sum(jv_debit) <> sum(jv_credit) ";

                if (Con_Oracle.IsRowExists(sql))
                {
                    throw new System.Exception("Total Debit and Credit Not Equal");
                }

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                if (oldccid != "")
                    Lib.UpdateArApInvNos(oldccid);
                if (Record.jvh_cc_category != "EMPLOYEE")
                {
                    if (Record.jvh_cc_id.ToString() != "")
                        Lib.UpdateArApInvNos(Record.jvh_cc_id);
                }


                string str = "TOT " + Record.jvh_tot_amt.ToString() + ", GST " + Record.jvh_gst_amt.ToString() + ", NET " + Record.jvh_net_amt.ToString() + ", " + Record.jvh_acc_name;
                Lib.AuditLog("ARAP", Record.jvh_type, Record.rec_mode, Record._globalvariables.comp_code, Record._globalvariables.branch_code, Record._globalvariables.user_code, Record.jvh_pkid, Record.jvh_docno, str);

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
            Con_Oracle.CloseConnection();

            RetData.Add("jvh_vrno", iVrNo.ToString());
            RetData.Add("jvh_docno", DocNo);

            return RetData;
        }




        private void SaveLedgerRecord(string mode,
            string EntID, string pkid,
            string acc_id, string acc_name, Boolean is_gst_item, string currid, string sacid,
            decimal qty, decimal rate,
            Boolean Taxable, Boolean gst_edited, string drcr,
            decimal ftotal, decimal exrate, decimal total,
            decimal taxable_amt, string gst_type,
            decimal cgst_rate, decimal sgst_rate, decimal igst_rate,
            decimal cgst_amt, decimal sgst_amt, decimal igst_amt, decimal gst_amt, decimal gst_rate, string cntr_type_id,
            decimal debit, decimal credit, decimal net_total,
            decimal total_fc, decimal cgst_famt, decimal sgst_famt, decimal igst_famt, decimal gst_famt, decimal net_ftotal,
            string doc_type, string bank, string branch, int chqno, string due_date, string pay_reason,
            string supp_docs, string paid_to, string remarks, string row_type, string year_code, string comp_code, string branch_code,
            int iCtr
            )
        {

            DBRecord Rec = new DBRecord();
            Rec.CreateRow("Ledgert", mode, "jv_pkid", pkid);
            Rec.InsertString("jv_parent_id", EntID);

            Rec.InsertString("jv_acc_id", acc_id);
            Rec.InsertString("jv_acc_name", acc_name);
            Rec.InsertString("jv_curr_id", currid);
            Rec.InsertString("jv_sac_id", sacid);

            Rec.InsertString("jv_posted", "Y");

            Rec.InsertNumeric("jv_qty", qty.ToString());
            Rec.InsertNumeric("jv_rate", rate.ToString());

            if (Taxable)
                Rec.InsertString("jv_is_taxable", "Y");
            else
                Rec.InsertString("jv_is_taxable", "N");

            if (is_gst_item)
                Rec.InsertString("jv_is_gst_item", "Y");
            else
                Rec.InsertString("jv_is_gst_item", "N");

            if (gst_edited)
                Rec.InsertString("jv_gst_edited", "Y");
            else
                Rec.InsertString("jv_gst_edited", "N");


            Rec.InsertString("jv_cntr_type_id", cntr_type_id);

            Rec.InsertString("jv_drcr", drcr);

            Rec.InsertNumeric("jv_ftotal", ftotal.ToString());
            Rec.InsertNumeric("jv_exrate", exrate.ToString());
            Rec.InsertNumeric("jv_total", total.ToString());
            Rec.InsertNumeric("jv_taxable_amt", taxable_amt.ToString());

            Rec.InsertNumeric("jv_total_fc", total_fc.ToString());

            Rec.InsertString("jv_gst_type", gst_type);

            Rec.InsertNumeric("jv_cgst_rate", cgst_rate.ToString());
            Rec.InsertNumeric("jv_cgst_amt", cgst_amt.ToString());

            Rec.InsertNumeric("jv_sgst_rate", sgst_rate.ToString());
            Rec.InsertNumeric("jv_sgst_amt", sgst_amt.ToString());

            Rec.InsertNumeric("jv_igst_rate", igst_rate.ToString());
            Rec.InsertNumeric("jv_igst_amt", igst_amt.ToString());

            Rec.InsertNumeric("jv_gst_amt", gst_amt.ToString());
            Rec.InsertNumeric("jv_gst_rate", gst_rate.ToString());

            if (drcr == "DR")
            {
                Rec.InsertNumeric("jv_debit", debit.ToString());
                Rec.InsertNumeric("jv_credit", "0");
            }
            if (drcr == "CR")
            {
                Rec.InsertNumeric("jv_debit", "0");
                Rec.InsertNumeric("jv_credit", credit.ToString());
            }

            Rec.InsertNumeric("jv_net_total", net_total.ToString());


            Rec.InsertNumeric("jv_cgst_famt", cgst_famt.ToString());
            Rec.InsertNumeric("jv_sgst_famt", sgst_famt.ToString());
            Rec.InsertNumeric("jv_igst_famt", igst_famt.ToString());
            Rec.InsertNumeric("jv_gst_famt", gst_famt.ToString());

            Rec.InsertNumeric("jv_net_ftotal", net_ftotal.ToString());

            Rec.InsertString("jv_doc_type", doc_type);
            Rec.InsertString("jv_bank", bank);
            Rec.InsertString("jv_branch", branch);
            Rec.InsertNumeric("jv_chqno", chqno.ToString());
            Rec.InsertDate("jv_due_date", due_date);

            Rec.InsertString("jv_pay_reason", pay_reason);
            Rec.InsertString("jv_supp_docs", supp_docs);
            Rec.InsertString("jv_paid_to", paid_to);
            Rec.InsertString("jv_remarks", remarks);

            Rec.InsertString("jv_row_type", row_type);

            Rec.InsertNumeric("jv_year", year_code);

            Rec.InsertString("rec_deleted", "N");

            Rec.InsertString("rec_company_code", comp_code);
            Rec.InsertString("rec_branch_code", branch_code);

            Rec.InsertNumeric("jv_ctr", iCtr.ToString());

            sql = Rec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);

        }

        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "SALES EXECUTIVE");
            RetData.Add("smanlist", lovservice.Lov(parameter)["param"]);

            return RetData;
        }
    }
}


