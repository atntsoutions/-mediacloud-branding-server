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
    public class BillingService : BL_Base
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
            string parentid = SearchData["parentid"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();

            try
            {
                sWhere = " where  1=1 ";
                sWhere += " and (";
                sWhere += " a.rec_company_code = '{COMP}'";
                sWhere += " and a.rec_branch_code = '{BRANCH}'";
                sWhere += " and a.jvh_year =  {FINYEAR}";
                sWhere += " and a.jvh_cc_id = '{PARENTID}'";
                sWhere += " ) ";

                sWhere = sWhere.Replace("{COMP}", company_code);
                sWhere = sWhere.Replace("{BRANCH}", branch_code);
                sWhere = sWhere.Replace("{FINYEAR}", year_code);
                sWhere = sWhere.Replace("{ROWTYPE}", rowtype);
                sWhere = sWhere.Replace("{PARENTID}", parentid);

                DataTable Dt_List = new DataTable();
                sql = "";
                sql += " select  jvh_pkid,jvh_vrno,jvh_docno, jvh_type,jvh_date, jvh_gstin,jvh_gst_type,acc_name, ";
                sql += " jvh_tot_amt, jvh_gst_amt,jvh_rc,jvh_net_amt,";
                sql += " jvh_subtype,jvh_rec_source,jvh_exwork ";
                sql += " from ledgerh a ";
                sql += " left join acctm b on a.jvh_acc_id = b.acc_pkid ";
                sql += sWhere;
                sql += " order by jvh_vrno";


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
                    mRow.jvh_acc_name = Dr["acc_name"].ToString();
                    mRow.jvh_gstin = Dr["jvh_gstin"].ToString();
                    mRow.jvh_gst_type = Dr["jvh_gst_type"].ToString();
                    mRow.jvh_rc = false;
                    if (Dr["jvh_rc"].ToString() == "Y")
                        mRow.jvh_rc = true;

                    mRow.jvh_exwork = false;
                    if (Dr["jvh_exwork"].ToString() == "Y")
                        mRow.jvh_exwork = true;

                    mRow.jvh_tot_amt = Lib.Conv2Decimal(Dr["jvh_tot_amt"].ToString());
                    mRow.jvh_gst_amt = Lib.Conv2Decimal(Dr["jvh_gst_amt"].ToString());
                    mRow.jvh_net_amt = Lib.Conv2Decimal(Dr["jvh_net_amt"].ToString());

                    mList.Add(mRow);
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

        public Dictionary<string, object> GetPendingList(Dictionary<string, object> SearchData)
        {

            

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();

            string Narration = "";
            string Shipper = "";

            string id = SearchData["parentid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                List<Ledgert> mList = new List<Ledgert>();
                Ledgert aRow;


                Con_Oracle = new DBConnection();

                sql = "select a.hbl_type,a.hbl_no as SINO ,a.hbl_bl_no as hbl , a.hbl_job_nos as jobnos ,a.hbl_book_cntr as cntr, ";
                sql += " b.hbl_bl_no as mbl, ";
                sql += " c.acc_name as shipper1,d.acc_name as shipper2  ";
                sql += " from hblm a ";
                sql += " left join hblm b on a.hbl_mbl_id = b.hbl_pkid";
                sql += " left join acctm c on a.hbl_exp_id = c.acc_pkid";
                sql += " left join acctm d on a.hbl_imp_id = d.acc_pkid";
                sql += " where a.hbl_pkid = '" + id + "'";

                DataTable Dt_Data =  Con_Oracle.ExecuteQuery(sql); 
                foreach (DataRow Dr1 in Dt_Data.Rows)
                {
                    Narration = "";
                    if (Dr1["hbl_type"].ToString() == "HBL-SE" || Dr1["hbl_type"].ToString() == "HBL-AE")
                    {
                        Shipper = Dr1["SHIPPER1"].ToString();
                        Narration += " job# " + Dr1["JOBNOS"].ToString() + ",";
                        Narration += " SI# " + Dr1["SINO"].ToString() + ",";
                        if (Dr1["hbl_type"].ToString() == "HBL-SE")
                        {
                            Narration += " MBL# " + Dr1["mbl"].ToString() + ",";
                            Narration += " HBL# " + Dr1["hbl"].ToString() + ",";
                            Narration += " Cntr# " + Dr1["cntr"].ToString();
                        }
                        if (Dr1["hbl_type"].ToString() == "HBL-AE")
                        {
                            Narration += " MAWB# " + Dr1["mbl"].ToString() + ",";
                            Narration += " HWWB# " + Dr1["hbl"].ToString() + ",";
                        }
                    }
                    if (Dr1["hbl_type"].ToString() == "HBL-SI" || Dr1["hbl_type"].ToString() == "HBL-AI")
                    {
                        Shipper = Dr1["SHIPPER2"].ToString();
                        if (Dr1["hbl_type"].ToString() == "HBL-SI")
                        {
                            Narration += " SI# " + Dr1["SINO"].ToString() + ",";
                            Narration += " MBL# " + Dr1["mbl"].ToString() + ",";
                            Narration += " HBL# " + Dr1["hbl"].ToString() + ",";
                        }
                        if (Dr1["hbl_type"].ToString() == "HBL-AI")
                        {
                            Narration += " SI# " + Dr1["SINO"].ToString() + ",";
                            Narration += " MAWB# " + Dr1["mbl"].ToString() + ",";
                            Narration += " HABW# " + Dr1["hbl"].ToString() + ",";
                        }
                    }
                    Narration = Narration.ToUpper();
                }
                Dt_Data.Rows.Clear();

                sql = "select inv_pkid, inv_jvid, inv_source, nvl(inv_posted,'N') as inv_posted, '' as jv_parent_id,";
                sql += " inv_acc_id as jv_acc_id, c.acc_code as jv_acc_code, c.acc_main_code, inv_acc_name as jv_acc_name,c.acc_cost_centre, ";
                sql += " d.param_pkid as jv_curr_id, d.param_code as jv_curr_code,";
                sql += " sac.param_pkid as jv_sac_id, sac.param_code as jv_sac_code, ";
                sql += " inv_qty as jv_qty, inv_rate as jv_rate,inv_ftotal as jv_ftotal, inv_exrate as jv_exrate, inv_total as jv_total, ";
                sql += " inv_total as jv_taxable_amt, inv_drcr as jv_drcr, ";
                sql += " case when inv_drcr = 'DR' then inv_total else 0 end as jv_debit, ";
                sql += " case when inv_drcr = 'CR' then inv_total else 0 end as jv_credit, ";
                sql += " 0 as jv_cgst_rate, 0 as jv_cgst_amt, 0 as jv_sgst_rate, 0 as jv_sgst_amt, 0 as jv_igst_rate, 0 as jv_igst_amt, 0 as jv_gst_amt,0 as jv_gst_rate, ";
                sql += " inv_cntr_type_id as jv_cntr_type_id, cntrtype.param_code as jv_cntr_type_code,";
                sql += " inv_total as jv_net_total, '' as jv_gst_type, 'N' as jv_gst_edited,  ";
                sql += " 'N' as jv_is_gst_item, c.acc_taxable as jv_is_taxable ";
                sql += " from jobincome a ";
                sql += " left join acctm c on a.inv_acc_id = c.acc_pkid ";
                sql += " left join param d on a.inv_curr_id = d.param_pkid";
                sql += " left join param sac on c.acc_sac_id = sac.param_pkid ";
                sql += " left join param cntrtype on a.inv_cntr_type_id = cntrtype.param_pkid ";
                sql += " where inv_parent_id = '{ID}' and inv_source in ('CLEARING INCOME', 'FREIGHT MEMO', 'LOCAL CHARGES','EX-WORK') ";
                sql += " and inv_type ='PREPAID' and nvl(inv_posted,'N') = 'N' ";
                sql += " order by inv_ctr ";
                sql = sql.Replace("{ID}", id);


                Dt_Rec = new DataTable();

                



                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    aRow = new Ledgert();
                    aRow.jv_pkid = Dr["inv_pkid"].ToString();
                    aRow.jv_parent_id = Dr["jv_parent_id"].ToString() ; //Dr["jv_parent_id"].ToString();

                    aRow.jv_income_type = Dr["inv_source"].ToString();

                    aRow.jv_acc_id = Dr["jv_acc_id"].ToString();
                    aRow.jv_acc_code = Dr["jv_acc_code"].ToString();
                    aRow.jv_acc_name = Dr["jv_acc_name"].ToString();

                    aRow.jv_acc_main_code = Dr["acc_main_code"].ToString();

                    aRow.jv_acc_against_invoice = "N";
                    aRow.jv_acc_cost_centre = Dr["acc_cost_centre"].ToString();


                    if (Dr["jv_is_gst_item"].ToString() == "Y")
                        aRow.jv_is_gst_item = true;
                    else
                        aRow.jv_is_gst_item = false;


                    if (Dr["jv_is_taxable"].ToString() == "Y")
                        aRow.jv_is_taxable = true;
                    else
                        aRow.jv_is_taxable = false;

                    if (Dr["inv_posted"].ToString() == "Y")
                        aRow.jv_selected = true;
                    else
                        aRow.jv_selected = false;

                    if (Dr["jv_gst_edited"].ToString() == "Y")
                        aRow.jv_gst_edited = true;
                    else
                        aRow.jv_gst_edited = false;


                    aRow.jv_cntr_type_id = Dr["jv_cntr_type_id"].ToString();
                    aRow.jv_cntr_type_code = Dr["jv_cntr_type_code"].ToString();

                    aRow.jv_curr_id = Dr["jv_curr_id"].ToString();
                    aRow.jv_curr_code = Dr["jv_curr_code"].ToString();

                    aRow.jv_sac_id = Dr["jv_sac_id"].ToString();
                    aRow.jv_sac_code = Dr["jv_sac_code"].ToString();




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



                    mList.Add(aRow);
                }
                mRow.LedgerList = mList;

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            RetData.Add("narration", Narration);
            RetData.Add("shipper", Shipper);
            return RetData;
        }



        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Ledgerh mRow = new Ledgerh();

            string lockedmsg = "";
            string parentid = SearchData["parentid"].ToString();
            string id = SearchData["pkid"].ToString();

            //string mode = "EDIT";

            string Narration = "";

            try
            {
                Con_Oracle = new DBConnection();





                DataTable Dt_Rec = new DataTable();

                sql = "";
                sql += " select ";
                sql += " jvh_pkid, jvh_date, jvh_year, jvh_type, jvh_vrno, jvh_docno,jvh_exwork, ";
                sql += " jvh_acc_id,acc_code as jvh_acc_code, acc_name as jvh_acc_name,";
                sql += " addr.add_line1||'\n'||addr.add_line2||'\n'||addr.add_line3 as  jvh_acc_br_addr,addr.add_email as jvh_acc_br_email,";
                sql += " jvh_gst ,jvh_gstin ,jvh_state_id,state.param_code as jvh_state_code,state.param_name as jvh_state_name, ";
                sql += " jvh_gst_type,jvh_org_invno, jvh_org_invdt,  ";
                sql += " jvh_rc,jvh_sez, jvh_is_export, jvh_exrate,jvh_sman_id ,jvh_reference ,jvh_reference_date,jvh_narration,a.rec_category, ";
                sql += " jvh_cgst_amt, jvh_sgst_amt, jvh_igst_amt, jvh_gst_amt,";
                sql += " jvh_debit, jvh_credit, jvh_curr_id, jvh_curr_code, jvh_acc_br_id, add_branch_slno as jvh_acc_br_slno, jvh_cc_category, ";
                sql += " jvh_cc_id, cc_code as jvh_cc_code,cc_name as jvh_cc_name, jvh_rec_source, ";
                sql += " jvh_edit_code, jvh_edit_date, a.rec_locked,a.rec_company_code, a.rec_branch_code,jvh_banktype ";
                sql += " from ledgerh a left join acctm on jvh_acc_id = acc_pkid ";
                sql += " left join addressm addr on jvh_acc_br_id = add_pkid ";
                sql += " left join param state on jvh_state_id = state.param_pkid ";
                sql += " left join costcenterm cc on jvh_cc_id =cc.cc_pkid ";
                sql += " where  a.jvh_pkid ='" + id + "'";

                
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Ledgerh();
                    mRow.jvh_pkid = Dr["jvh_pkid"].ToString();
                    mRow.jvh_vrno = Lib.Conv2Integer(Dr["jvh_vrno"].ToString());
                    mRow.jvh_docno = Dr["jvh_docno"].ToString();
                    mRow.jvh_date = Lib.DatetoString(Dr["jvh_date"]);

                    mRow.jvh_type = Dr["jvh_type"].ToString();

                    mRow.jvh_year = Lib.Conv2Integer(Dr["jvh_year"].ToString());

                    mRow.jvh_rec_source = Dr["jvh_rec_source"].ToString();

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


                    mRow.jvh_exwork = false;
                    if (Dr["jvh_exwork"].ToString() == "Y")
                        mRow.jvh_exwork = true;


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

                    string JvhDate = Lib.StringToDate(Dr["jvh_date"]);
                    lockedmsg = Lib.IsDateLocked(JvhDate, Dr["jvh_type"].ToString(),
                        Dr["rec_company_code"].ToString(),
                        Dr["rec_branch_code"].ToString(), Dr["jvh_year"].ToString());

                    break;
                }


                sql = "select inv_pkid, inv_jvid, inv_source, jv_pkid, jv_parent_id, nvl(inv_posted,'N') as inv_posted,";
                sql += " inv_acc_id as jv_acc_id, jv_acc_id as old_acc_id, c.acc_main_code as jv_acc_main_code, c.acc_code as jv_acc_code, inv_acc_name as jv_acc_name,c.acc_cost_centre, ";
                sql += " d.param_pkid as jv_curr_id, d.param_code as jv_curr_code,";
                sql += " sac.param_pkid as jv_sac_id, sac.param_code as jv_sac_code,  c.acc_taxable , jv_is_taxable,";
                sql += " inv_qty as jv_qty, inv_rate as jv_rate,inv_ftotal as jv_ftotal, inv_exrate as jv_exrate, inv_total as jv_total, ";
                sql += " inv_total as jv_taxable_amt, inv_drcr as jv_drcr, ";
                sql += " case when inv_drcr = 'DR' then inv_total else 0 end as jv_debit, ";
                sql += " case when inv_drcr = 'CR' then inv_total else 0 end as jv_credit, ";
                sql += " jv_cgst_rate, jv_cgst_amt, jv_sgst_rate, jv_sgst_amt, jv_igst_rate, jv_igst_amt, jv_gst_amt,jv_gst_rate, ";
                sql += " inv_cntr_type_id as jv_cntr_type_id, cntrtype.param_code as jv_cntr_type_code,";
                sql += " jv_net_total,  jv_gst_type,  jv_gst_edited,  jv_is_gst_item, inv_posted, ";
                sql += " tax_cgst_rate, tax_sgst_rate, tax_igst_rate ";
                sql += " from jobincome a ";
                sql += " left join ledgert b on a.inv_pkid = b.jv_pkid and b.jv_parent_id = '{ID}'";
                sql += " left join acctm c on a.inv_acc_id = c.acc_pkid ";
                sql += " left join param d on a.inv_curr_id = d.param_pkid";
                sql += " left join param sac on c.acc_sac_id = sac.param_pkid ";
                sql += " left join taxm on a.inv_acc_id = taxm.tax_acc_id";
                sql += " left join param cntrtype on a.inv_cntr_type_id = cntrtype.param_pkid ";
                sql += " where inv_parent_id = '{PARENTID}' and inv_source in ('CLEARING INCOME', 'FREIGHT MEMO', 'LOCAL CHARGES','EX-WORK') ";
                sql += " and inv_type ='PREPAID' and nvl(inv_jvid,'{ID}') = '{ID}' ";
                sql += " order by inv_ctr ";

                sql = sql.Replace("{ID}", id);
                sql = sql.Replace("{PARENTID}",parentid);

                List<Ledgert> mList = new List<Ledgert>();
                Ledgert aRow;


                Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    aRow = new Ledgert();
                    aRow.jv_pkid = Dr["inv_pkid"].ToString();
                    aRow.jv_parent_id = Dr["jv_parent_id"].ToString();

                    aRow.jv_income_type = Dr["inv_source"].ToString();

                    aRow.jv_acc_id = Dr["jv_acc_id"].ToString();
                    aRow.jv_acc_main_code = Dr["jv_acc_main_code"].ToString();
                    aRow.jv_acc_code = Dr["jv_acc_code"].ToString();
                    aRow.jv_acc_name = Dr["jv_acc_name"].ToString();

                    aRow.jv_acc_against_invoice = "N";
                    aRow.jv_acc_cost_centre = Dr["acc_cost_centre"].ToString();


                    if (Dr["jv_is_gst_item"].ToString() == "Y")
                        aRow.jv_is_gst_item = true;
                    else
                        aRow.jv_is_gst_item = false;


                    if (Dr["jv_gst_edited"].ToString() == "Y")
                        aRow.jv_gst_edited = true;
                    else
                        aRow.jv_gst_edited = false;


                    aRow.jv_cntr_type_id = Dr["jv_cntr_type_id"].ToString();
                    aRow.jv_cntr_type_code = Dr["jv_cntr_type_code"].ToString();

                    aRow.jv_curr_id = Dr["jv_curr_id"].ToString();
                    aRow.jv_curr_code = Dr["jv_curr_code"].ToString();

                    aRow.jv_sac_id = Dr["jv_sac_id"].ToString();
                    aRow.jv_sac_code = Dr["jv_sac_code"].ToString();

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


                    if (Dr["inv_posted"].ToString() == "Y")
                    {
                        aRow.jv_selected = true;
                        if (Dr["acc_taxable"].ToString() == "Y")
                            aRow.jv_is_taxable = true;
                        else
                            aRow.jv_is_taxable = false;

                        aRow.jv_cgst_rate = Lib.Conv2Decimal(Dr["tax_cgst_rate"].ToString());
                        aRow.jv_sgst_rate = Lib.Conv2Decimal(Dr["tax_sgst_rate"].ToString());
                        aRow.jv_igst_rate = Lib.Conv2Decimal(Dr["tax_igst_rate"].ToString());

                        if (mRow.jvh_gst_type == "INTRA-STATE")
                        {
                            aRow.jv_cgst_rate = Lib.Conv2Decimal(Dr["tax_cgst_rate"].ToString());
                            aRow.jv_sgst_rate = Lib.Conv2Decimal(Dr["tax_sgst_rate"].ToString());
                            aRow.jv_igst_rate = 0;
                        }
                        if (mRow.jvh_gst_type == "INTER-STATE")
                        {
                            aRow.jv_cgst_rate = 0;
                            aRow.jv_sgst_rate = 0;
                            aRow.jv_igst_rate = Lib.Conv2Decimal(Dr["tax_igst_rate"].ToString());
                        }
                    }
                    else
                    {
                        aRow.jv_selected = false;
                        if (Dr["acc_taxable"].ToString() == "Y")
                            aRow.jv_is_taxable = true;
                        else
                            aRow.jv_is_taxable = false;
                    }

                    mList.Add(aRow);
                }


                mRow.LedgerList = mList;

                // Any Allocation Exists against this Record
                mRow.jvh_allocation_found = false;
                sql = "select xref_jvh_id from ledgerxref where (xref_dr_jvh_id = '{ID}' or xref_cr_jvh_id = '{ID}') and xref_jvh_id<> '{ID}'";
                sql = sql.Replace("{ID}", id);
                if (Con_Oracle.IsRowExists(sql))
                    mRow.jvh_allocation_found = true;


                sql = "select hbl_type,hbl_no,hbl_bl_no, hbl_job_nos,hbl_book_cntr  from hblm where hbl_pkid = '" + parentid + "'";
                DataTable Dt_Data = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr1 in Dt_Data.Rows)
                {
                    Narration += "SI# " + Dr1["hbl_no"].ToString();
                    if (Dr1["hbl_bl_no"].ToString().Length > 0)
                        Narration += " BL# " + Dr1["hbl_bl_no"].ToString();
                    if (Dr1["hbl_job_nos"].ToString().Length > 0)
                        Narration += " JOB# " + Dr1["hbl_job_nos"].ToString();
                    if (Dr1["hbl_book_cntr"].ToString().Length > 0)
                        Narration += " CNTR# " + Dr1["hbl_book_cntr"].ToString();
                    if (Narration.Length > 255)
                        Narration = Narration.Substring(0, 255);
                }
                Dt_Data.Rows.Clear();




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
            RetData.Add("narration", Narration);
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

                if (Lib.IsValidGST(Record.jvh_gst, Record.jvh_gstin, Record.jvh_state_code,Record.jvh_igst_exception) == false)
                {
                    str += " | Invalid GST (Invalid GST Number  or Mismatch between GST Number and State Code)";
                }

                if (Record.rec_mode == "ADD")
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
                        str +=  " | Back Dated Entry Not Possible ";
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


            Boolean jobWise = false;

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


                Dt_cc_jobs = Lib.getCCJOBS(Record.jvh_cc_category, Record.jvh_cc_id);
                Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                Dt_cc_cntr = Lib.getCCCntrs(Record.jvh_cc_category, Record.jvh_cc_id);




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
                    foreach (DataRow Dr in dt_c1.Rows)
                    {
                        if (Dr["jvh_cc_category"].ToString() != "EMPLOYEE" && Dr["jvh_cc_id"].ToString() != "")
                        {
                            oldccid = Dr["jvh_cc_id"].ToString();
                            if (oldccid == Record.jvh_cc_id)
                                oldccid = "";
                        }
                    }
                }

                jobWise = false;
                if (Record.jvh_cc_category == "SI SEA EXPORT")
                {
                    sql = "select hbl_mbl_id from hblm where  hbl_pkid = '" + Record.jvh_cc_id + "' and hbl_type ='HBL-SE' and hbl_mbl_id is null";
                    if (Con_Oracle.IsRowExists(sql))
                        jobWise = true;
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
                    Rec.InsertString("jvh_rec_source", "OP");
                    Rec.InsertString("rec_locked", "N");

                    Rec.InsertString("rec_deleted", "N");
                    Rec.InsertNumeric("jvh_year", Record._globalvariables.year_code);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);

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


                Rec.InsertString("jvh_gstin", Record.jvh_gstin);
                Rec.InsertString("jvh_state_id", Record.jvh_state_id);
                Rec.InsertString("jvh_gst_type", Record.jvh_gst_type);

                Rec.InsertString("jvh_org_invno", Record.jvh_org_invno);
                Rec.InsertDate("jvh_org_invdt", Record.jvh_org_invdt);


                if (Record.jvh_exwork)
                    Rec.InsertString("jvh_exwork", "Y");
                else
                    Rec.InsertString("jvh_exwork", "N");

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

                if ( Record.jvh_cc_category.Contains("SEA EXPORT"))
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

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);


                sql = "update jobincome set inv_jvid='', inv_posted= 'N' where inv_jvid ='" + Record.jvh_pkid + "'";
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
                Record.jvh_curr_id, "", 1, Record.jvh_tot_famt, false, false,
                Record.jvh_drcr, Record.jvh_tot_famt, Record.jvh_exrate, Record.jvh_tot_amt, Record.jvh_tot_amt, "",
                0, 0, 0, 0, 0, 0, 0, 0, "", nDr, nCr, Record.jvh_net_amt, Record.jvh_net_famt, 0, 0, 0, 0, Record.jvh_net_famt,
                "", "", "", 0, "", "", "", "", "", "HEADER",
                Record._globalvariables.year_code, Record._globalvariables.comp_code, Record._globalvariables.branch_code, iCtr
                );


                gst_dr_or_cr = "";
                foreach (Ledgert Row in Record.LedgerList)
                {
                    if (Row.jv_selected)
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
                        Row.jv_cgst_amt, Row.jv_sgst_amt, Row.jv_igst_amt, Row.jv_gst_amt, Row.jv_gst_rate, Row.jv_cntr_type_id,
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

                        sql = "update jobincome set inv_jvid='" + Record.jvh_pkid  + "', inv_posted= 'Y' where inv_pkid ='" + Row.jv_pkid + "'";
                        Con_Oracle.ExecuteNonQuery(sql);

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
                                0, 0, 0, 0, 0, "", nGstRowAmt, nGstRowAmt,
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




                sql = "delete from  costcentert where ct_jvh_id ='" + Record.jvh_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);


                foreach (Ledgert Row in Record.LedgerList)
                {
                    if (Lib.getCCType(Row.jv_acc_main_code, jobWise) == "JOB WISE")
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
                    else if (Lib.getCCType(Row.jv_acc_main_code) == "HBL WISE")
                    {
                        foreach (DataRow  Drc in Dt_cc_cntr.Rows)
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

                if (Record.jvh_exwork)
                {
                    if (Record.jvh_acc_code.StartsWith("1105") || Record.jvh_acc_code.StartsWith("1205") || Record.jvh_acc_code.StartsWith("1305") || Record.jvh_acc_code.StartsWith("1405"))
                    {
                        iCtr = 0;
                        foreach (DataRow Drc in Dt_cc_cntr.Rows)
                        {
                            iCtr++;

                            cc_amt = Record.jvh_net_amt / Lib.Conv2Integer(Drc["tot"].ToString());
                            cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);

                            // This is header Record

                            Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                            Rec.InsertString("ct_jvh_id", Record.jvh_pkid);
                            Rec.InsertNumeric("ct_year", Record.jvh_year.ToString());
                            Rec.InsertString("ct_jv_id", Record.jvh_pkid);
                            Rec.InsertString("ct_acc_id", Record.jvh_acc_id);
                            Rec.InsertString("ct_category", Drc["cc_category"].ToString());
                            Rec.InsertString("ct_cost_id", Drc["id"].ToString());
                            Rec.InsertNumeric("ct_cost_year", Record.jvh_year.ToString());
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
                            cc_amt = Record.jvh_tot_amt / Lib.Conv2Integer(Dr["tot"].ToString());
                            cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                            Rec.CreateRow("CostCentert", "ADD", "ct_pkid", System.Guid.NewGuid().ToString().ToUpper());
                            Rec.InsertString("ct_jvh_id", Record.jvh_pkid);
                            Rec.InsertNumeric("ct_year", Record.jvh_year.ToString());
                            Rec.InsertString("ct_jv_id", Record.jvh_pkid);
                            Rec.InsertString("ct_acc_id", Record.jvh_acc_id);
                            Rec.InsertString("ct_category", Dr["cc_category"].ToString());
                            Rec.InsertString("ct_cost_id", Dr["id"].ToString());
                            Rec.InsertNumeric("ct_cost_year", Record.jvh_year.ToString());
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
                }



                //sql = "delete from  ledgerxref where xref_jvh_id ='" + Record.jvh_pkid + "'";
                //Con_Oracle.ExecuteNonQuery(sql);


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

                string str = " TOT " + Record.jvh_tot_amt.ToString() + ", GST " + Record.jvh_gst_amt.ToString() + ", NET " + Record.jvh_net_amt.ToString() + ", " + Record.jvh_acc_name;
                Lib.AuditLog("INVOICE", Record.jvh_type, Record.rec_mode, Record._globalvariables.comp_code, Record._globalvariables.branch_code, Record._globalvariables.user_code, Record.jvh_pkid, Record.jvh_docno, str);

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


