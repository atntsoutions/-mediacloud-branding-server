using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

using System.Drawing;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;
namespace BLAccounts
{
    public class MoneyTransferService : BL_Base
    {
        Dictionary<string, string> DicPayMode = new Dictionary<string, string>();
        Dictionary<string, string> DicPaytype = new Dictionary<string, string>();
        Dictionary<string, string> DicCategory = new Dictionary<string, string>();

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            moneytransfer mRow = new moneytransfer();
            string ErrorMessage = "";
            string Mode = "ADD";
            string report_folder = "";
            string jvhid = "";
            string jvid = "";
           
            try
            {
                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();
                if (SearchData.ContainsKey("jvhid"))
                    jvhid = SearchData["jvhid"].ToString();
                if (SearchData.ContainsKey("jvid"))
                    jvid = SearchData["jvid"].ToString();
                
                Con_Oracle = new DBConnection();

                sql = "select mt_pkid from moneytransfer where mt_jv_id = '{PKID}' ";
                sql = sql.Replace("{PKID}", jvid);
                Mode = "ADD";
                if (Con_Oracle.IsRowExists(sql))
                    Mode = "EDIT";

                if (Mode == "ADD")
                {
                    mRow = InitRecord(jvid);
                }
                else
                {
                    sql = "select mt_pkid,mt_jvh_id ,mt_jv_id,mt_jv_acc_id,mt_jvh_docno,mt_type ,";
                    sql += " mt_txn_mode ,mt_corp_code,mt_cust_cfno,mt_cust_uniq_ref,";
                    sql += " mt_corp_acc_no,mt_value_date ,mt_txn_curr ,mt_txn_amt,";
                    sql += " mt_party_id ,b.cust_code as mt_party_code,b.cust_name as mt_party_name, mt_ben_id ,mt_ben_name ,mt_ben_code ,mt_ben_acc_no ,";
                    sql += " mt_ben_acc_type ,mt_ben_addr1,mt_ben_addr2,mt_ben_addr3,mt_ben_city ,";
                    sql += " mt_ben_state,mt_ben_pin,mt_ben_ifsc ,mt_ben_bank_name,mt_base_code,";
                    sql += " mt_chq_no ,mt_chq_date ,mt_payable_loc,mt_print_loc,mt_ben_email1,";
                    sql += " mt_ben_email2,mt_ben_mob,mt_corp_batch_no,mt_company_code ,mt_product_code ,";
                    sql += " mt_enrichment1,mt_enrichment2,mt_enrichment3,mt_enrichment4,mt_enrichment5,";
                    sql += " mt_pay_type,mt_corp_email,mt_transmission_date,mt_user_id,mt_user_dept,mt_lock ";
                    sql += " from moneytransfer a";
                    sql += " left join customerm b on a.mt_party_id = b.cust_pkid";
                    sql += " where mt_jv_id ='" + jvid + "'";

                    DataTable Dt_Rec = new DataTable();
                    Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_Rec.Rows)
                    {
                        mRow = new moneytransfer();
                        mRow.mt_pkid = Dr["mt_pkid"].ToString();
                        mRow.mt_jvh_id = Dr["mt_jvh_id"].ToString();
                        mRow.mt_jv_id = Dr["mt_jv_id"].ToString();
                        mRow.mt_jv_acc_id = Dr["mt_jv_acc_id"].ToString();
                        mRow.mt_jvh_docno = Dr["mt_jvh_docno"].ToString();
                        mRow.mt_type = Dr["mt_type"].ToString();
                        mRow.mt_txn_mode = Dr["mt_txn_mode"].ToString();
                        mRow.mt_corp_code = Dr["mt_corp_code"].ToString();
                        mRow.mt_cust_uniq_ref = Dr["mt_cust_uniq_ref"].ToString();
                        mRow.mt_corp_acc_no = Dr["mt_corp_acc_no"].ToString();
                        mRow.mt_value_date = Lib.DatetoString(Dr["mt_value_date"]);
                        mRow.mt_txn_curr = Dr["mt_txn_curr"].ToString();
                        mRow.mt_txn_amt = Lib.Convert2Decimal(Dr["mt_txn_amt"].ToString());
                        mRow.mt_party_id = Dr["mt_party_id"].ToString();
                        mRow.mt_party_code = Dr["mt_party_code"].ToString();
                        mRow.mt_party_name = Dr["mt_party_name"].ToString();
                        mRow.mt_ben_id = Dr["mt_ben_id"].ToString();
                        mRow.mt_ben_name = Dr["mt_ben_name"].ToString();
                        mRow.mt_ben_code = Dr["mt_ben_code"].ToString();
                        mRow.mt_ben_acc_no = Dr["mt_ben_acc_no"].ToString();
                        mRow.mt_ben_acc_type = Dr["mt_ben_acc_type"].ToString();
                        mRow.mt_ben_addr1 = Dr["mt_ben_addr1"].ToString();
                        mRow.mt_ben_addr2 = Dr["mt_ben_addr2"].ToString();
                        mRow.mt_ben_addr3 = Dr["mt_ben_addr3"].ToString();
                        mRow.mt_ben_city = Dr["mt_ben_city"].ToString();
                        mRow.mt_ben_state = Dr["mt_ben_state"].ToString();
                        mRow.mt_ben_pin = Dr["mt_ben_pin"].ToString();
                        mRow.mt_ben_ifsc = Dr["mt_ben_ifsc"].ToString();
                        mRow.mt_ben_bank_name = Dr["mt_ben_bank_name"].ToString();
                        mRow.mt_base_code = Dr["mt_base_code"].ToString();
                        mRow.mt_chq_no = Dr["mt_chq_no"].ToString();
                        mRow.mt_chq_date = Lib.DatetoString(Dr["mt_chq_date"]);
                        mRow.mt_payable_loc = Dr["mt_payable_loc"].ToString();
                        mRow.mt_print_loc = Dr["mt_print_loc"].ToString();
                        mRow.mt_ben_email1 = Dr["mt_ben_email1"].ToString();
                        mRow.mt_ben_email2 = Dr["mt_ben_email2"].ToString();
                        mRow.mt_ben_mob = Dr["mt_ben_mob"].ToString();
                        mRow.mt_corp_batch_no = Dr["mt_corp_batch_no"].ToString();
                        mRow.mt_company_code = Dr["mt_company_code"].ToString();
                        mRow.mt_product_code = Dr["mt_product_code"].ToString();
                        mRow.mt_enrichment1 = Dr["mt_enrichment1"].ToString();
                        mRow.mt_enrichment2 = Dr["mt_enrichment2"].ToString();
                        mRow.mt_enrichment3 = Dr["mt_enrichment3"].ToString();
                        mRow.mt_enrichment4 = Dr["mt_enrichment4"].ToString();
                        mRow.mt_enrichment5 = Dr["mt_enrichment5"].ToString();
                        mRow.mt_pay_type = Dr["mt_pay_type"].ToString();
                        mRow.mt_corp_email = Dr["mt_corp_email"].ToString();
                        mRow.mt_transmission_date = Dr["mt_transmission_date"].ToString();
                        mRow.mt_user_id = Dr["mt_user_id"].ToString();
                        mRow.mt_user_dept = Dr["mt_user_dept"].ToString();
                        mRow.mt_lock = Dr["mt_lock"].ToString();
                        break;
                    }
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("recmode", Mode);
            RetData.Add("record", mRow);
            return RetData;
        }

        private moneytransfer InitRecord(string jv_id)
        {
            moneytransfer Rec = new moneytransfer();
            Rec.mt_pkid = Guid.NewGuid().ToString().ToUpper();
            Rec.mt_jvh_id ="";
            Rec.mt_jv_id = jv_id;
            Rec.mt_jv_acc_id = "";
            Rec.mt_jvh_docno = "";
            Rec.mt_type = "P";
            Rec.mt_txn_mode = "";
            Rec.mt_corp_code = "CARGOMAR";
            Rec.mt_cust_cfno = 0;
            Rec.mt_cust_uniq_ref ="";
            Rec.mt_corp_acc_no = "";
            Rec.mt_value_date = DateTime.Now.ToString(Lib.FRONT_END_DATE_FORMAT);
            Rec.mt_txn_curr = "";
            Rec.mt_txn_amt = 0;
            Rec.mt_party_id = "";
            Rec.mt_party_code = "";
            Rec.mt_party_name = "";
            Rec.mt_ben_id = "";
            Rec.mt_ben_name = "";
            Rec.mt_ben_code = "";
            Rec.mt_ben_acc_no = "";
            Rec.mt_ben_acc_type = "";
            Rec.mt_ben_addr1 = "";
            Rec.mt_ben_addr2 = "";
            Rec.mt_ben_addr3 = "";
            Rec.mt_ben_city = "";
            Rec.mt_ben_state = "";
            Rec.mt_ben_pin = "";
            Rec.mt_ben_ifsc = "";
            Rec.mt_ben_bank_name = "";
            Rec.mt_base_code = "";
            Rec.mt_chq_no = "";
            Rec.mt_chq_date = "";
            Rec.mt_payable_loc = "";
            Rec.mt_print_loc = "";
            Rec.mt_ben_email1 = "";
            Rec.mt_ben_email2 = "";
            Rec.mt_ben_mob = "";
            Rec.mt_corp_batch_no = "";
            Rec.mt_company_code = "";
            Rec.mt_product_code = "";
            Rec.mt_enrichment1 = "";
            Rec.mt_enrichment2 = "";
            Rec.mt_enrichment3 = "";
            Rec.mt_enrichment4 = "";
            Rec.mt_enrichment5 = "";
            Rec.mt_pay_type = "";
            Rec.mt_corp_email = "utib562@cargomar.in";
            Rec.mt_transmission_date = "";
            Rec.mt_user_id = "";
            Rec.mt_user_dept = "";
            Rec.mt_lock = "";
            Rec.mt_slno = 0;

            sql = " select jv_parent_id, jv_acc_id,jvh_docno,c.acc_name as jv_acc_name,curr.param_code as currency,jv_total ";
            sql += " from ledgerh a inner join ledgert b on a.jvh_pkid = b.jv_parent_id";
            sql += " left join acctm c on b.jv_acc_id = c.acc_pkid";
            sql += " left join param curr on b.jv_curr_id = curr.param_pkid";
            sql += " where jv_pkid='" + jv_id + "' ";
            DataTable Dt_Rec = new DataTable();
            Dt_Rec = Con_Oracle.ExecuteQuery(sql);
            if (Dt_Rec.Rows.Count > 0)
            {
                Rec.mt_jvh_id = Dt_Rec.Rows[0]["jv_parent_id"].ToString();
                Rec.mt_jv_id = jv_id;
                Rec.mt_jv_acc_id = Dt_Rec.Rows[0]["jv_acc_id"].ToString();
                Rec.mt_jvh_docno = Dt_Rec.Rows[0]["jvh_docno"].ToString();
                Rec.mt_txn_curr = Dt_Rec.Rows[0]["currency"].ToString();
                Rec.mt_txn_amt = Lib .Conv2Decimal(Dt_Rec.Rows[0]["jv_total"].ToString());
                string[] sadta = Dt_Rec.Rows[0]["jv_acc_name"].ToString().Split('#');
                if (sadta.Length > 1)
                    Rec.mt_corp_acc_no = sadta[1].Trim();
            }

            return Rec;
        }

        public Dictionary<string, object> Save(moneytransfer Record)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string Mode = "";
            string sql = "";
            string ErrorMessage = "";
            string cfno = "";
            try
            {

                Con_Oracle = new DBConnection();

                Mode = "ADD";
                Record.mt_pkid = Guid.NewGuid().ToString().ToUpper();

                sql = " select mt_pkid from moneytransfer where mt_jv_id = '" + Record.mt_jv_id + "'";
                DataTable Dt_Rec = new DataTable();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Rec.Rows.Count > 0)
                {
                    Mode = "EDIT";
                    Record.mt_pkid = Dt_Rec.Rows[0]["mt_pkid"].ToString();
                }
                Record.rec_mode = Mode;

                if ((ErrorMessage = AllValid(Record)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (Record.rec_mode == "ADD")
                {
                    sql = "select nvl(max(mt_cust_cfno)+1,1001) as cfno from moneytransfer a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code = '{BRCODE}'";

                    sql = sql.Replace("{COMPCODE}", Record._globalvariables.comp_code);
                    sql = sql.Replace("{BRCODE}", Record._globalvariables.branch_code);

                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        cfno = Dt_Temp.Rows[0]["cfno"].ToString();
                        Record.mt_cust_cfno = Lib.Conv2Integer(Dt_Temp.Rows[0]["cfno"].ToString());
                    }
                    else
                    {
                        ErrorMessage = "Ref Number Not Found Try again";

                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception(ErrorMessage);
                    }
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("moneytransfer", Record.rec_mode, "mt_pkid", Record.mt_pkid);
                Rec.InsertString("mt_jv_acc_id", Record.mt_jv_acc_id);
                Rec.InsertString("mt_txn_mode", Record.mt_txn_mode);
                Rec.InsertString("mt_corp_code", Record.mt_corp_code);
                Rec.InsertString("mt_corp_acc_no", Record.mt_corp_acc_no);
                Rec.InsertString("mt_txn_curr", Record.mt_txn_curr);
                Rec.InsertNumeric("mt_txn_amt", Record.mt_txn_amt.ToString());
                Rec.InsertString("mt_party_id", Record.mt_party_id);
                Rec.InsertString("mt_ben_id", Record.mt_ben_id);
                Rec.InsertString("mt_ben_name", Record.mt_ben_name);
                Rec.InsertString("mt_ben_code", Record.mt_ben_code);
                Rec.InsertString("mt_ben_acc_no", Record.mt_ben_acc_no);
                Rec.InsertString("mt_ben_acc_type", Record.mt_ben_acc_type);
                Rec.InsertString("mt_ben_addr1", Record.mt_ben_addr1);
                Rec.InsertString("mt_ben_addr2", Record.mt_ben_addr2);
                Rec.InsertString("mt_ben_addr3", Record.mt_ben_addr3);
                Rec.InsertString("mt_ben_city", Record.mt_ben_city);
                Rec.InsertString("mt_ben_state", Record.mt_ben_state);
                Rec.InsertString("mt_ben_pin", Record.mt_ben_pin);
                Rec.InsertString("mt_ben_ifsc", Record.mt_ben_ifsc);
                Rec.InsertString("mt_ben_bank_name", Record.mt_ben_bank_name);
                Rec.InsertString("mt_base_code", Record.mt_base_code);
                Rec.InsertString("mt_chq_no", Record.mt_chq_no);
                Rec.InsertDate("mt_chq_date", Record.mt_chq_date);
                Rec.InsertString("mt_payable_loc", Record.mt_payable_loc);
                Rec.InsertString("mt_print_loc", Record.mt_print_loc);
                Rec.InsertString("mt_ben_email1", Record.mt_ben_email1);
                Rec.InsertString("mt_ben_email2", Record.mt_ben_email2);
                Rec.InsertString("mt_ben_mob", Record.mt_ben_mob);
                Rec.InsertString("mt_corp_batch_no", Record.mt_corp_batch_no);
                Rec.InsertString("mt_company_code", Record.mt_company_code);
                Rec.InsertString("mt_product_code", Record.mt_product_code);
                Rec.InsertString("mt_enrichment1", Record.mt_enrichment1);
                Rec.InsertString("mt_enrichment2", Record.mt_enrichment2);
                Rec.InsertString("mt_enrichment3", Record.mt_enrichment3);
                Rec.InsertString("mt_enrichment4", Record.mt_enrichment4);
                Rec.InsertString("mt_enrichment5", Record.mt_enrichment5);
                Rec.InsertString("mt_pay_type", Record.mt_pay_type);

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("mt_corp_email", Record.mt_corp_email);
                    Rec.InsertString("mt_user_id", Record._globalvariables.user_code);
                    Rec.InsertString("mt_user_dept", "ACCOUNTS");
                    Rec.InsertNumeric("mt_cust_cfno", Record.mt_cust_cfno.ToString());
                    Rec.InsertString("mt_jvh_id", Record.mt_jvh_id);
                    Rec.InsertString("mt_jv_id", Record.mt_jv_id);
                    Rec.InsertString("mt_jvh_docno", Record.mt_jvh_docno);
                    Rec.InsertString("mt_type", Record.mt_type);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_created_date", "SYSDATE");
                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "SYSDATE");
                }


                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
 
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("recmode", Mode);
            return RetData;
        }


        public string AllValid(moneytransfer Record)
        {
            string str = "";
            try
            {
                //if (Record.bl_shipper_br_id.Trim().Length > 0 || Record.bl_shipper_id.Trim().Length > 0)
                //{
                //    sql = "select add_pkid from addressm where add_pkid = '{ADD_BRID}'";
                //    sql += " and  add_parent_id = '{PARENT_ID}'";
                //    sql = sql.Replace("{ADD_BRID}", Record.bl_shipper_br_id);
                //    sql = sql.Replace("{PARENT_ID}", Record.bl_shipper_id);
                //    if (!Con_Oracle.IsRowExists(sql))
                //        Lib.AddError(ref str, " Invalid Shipper Address ");
                //}


            }
            catch (Exception Ex)
            {
                Lib.AddError(ref str, Ex.Message.ToString());
            }
            return str;
        }

        public IDictionary<string, object> MtReport(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();


            DicPayMode = new Dictionary<string, string>();
            DicPayMode.Add("RT", "RTGS");
            DicPayMode.Add("NE", "NEFT");
            DicPayMode.Add("FT", "Fund Transfer (Axis to Axis)");
            DicPayMode.Add("CC", "Corporate Cheques");
            DicPayMode.Add("DD", "Demand Draft");
            DicPayMode.Add("PA", "IMPS");

            DicPaytype = new Dictionary<string, string>();
            DicPaytype.Add("CUST", "Customer Payment");
            DicPaytype.Add("MERC", "Merchant Payment");
            DicPaytype.Add("DIST", "Distributor Payment");
            DicPaytype.Add("INTN", "Internal Payment");
            DicPaytype.Add("VEND", "Vendor Payment");

            DicCategory = new Dictionary<string, string>();
            DicCategory.Add("10", "SAVINGS BANK");
            DicCategory.Add("11", "CURRENT ACCOUNT");
            DicCategory.Add("13", "CASH CREDIT");

            Con_Oracle = new DBConnection();
            List<moneytransfer> mList = new List<moneytransfer>();
            moneytransfer mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {

                sWhere = " where a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.rec_branch_code = '{BRCODE}'";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  a.mt_ben_acc_no  like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(mt_ben_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM moneytransfer a ";
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
                sql += " select mt_pkid,mt_jvh_id ,mt_jv_id,mt_jvh_docno,mt_type ,";
                sql += "  mt_txn_mode,mt_corp_code,mt_cust_uniq_ref,";
                sql += "  mt_corp_acc_no,mt_value_date,mt_txn_curr,mt_txn_amt,";
                sql += "  mt_ben_name ,mt_ben_code ,mt_ben_acc_no ,";
                sql += "  mt_ben_acc_type ,mt_ben_addr1,mt_ben_addr2,mt_ben_addr3,mt_ben_city ,";
                sql += "  mt_ben_state,mt_ben_pin,mt_ben_ifsc ,mt_ben_bank_name,";
                sql += "  mt_payable_loc,mt_print_loc,mt_ben_email1,";
                sql += "  mt_ben_email2,mt_ben_mob,";
                sql += "  mt_pay_type,mt_corp_email,to_char(mt_transmission_date,'DD/MM/YYYY HH24:MI:SS') as mt_transmission_date, ";
                sql += "  rec_created_date, row_number() over(order by mt_cust_uniq_ref,a.rec_created_date) rn ";
                sql += "  from moneytransfer a";

                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by mt_cust_uniq_ref,rec_created_date";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new moneytransfer();
                    mRow.mt_pkid = Dr["mt_pkid"].ToString();
                    mRow.mt_jvh_id = Dr["mt_jvh_id"].ToString();
                    mRow.mt_jv_id = Dr["mt_jv_id"].ToString();
                    mRow.mt_jvh_docno = Dr["mt_jvh_docno"].ToString();
                    mRow.mt_type = Dr["mt_type"].ToString();

                    mRow.mt_txn_mode = Dr["mt_txn_mode"].ToString();
                    if (DicPayMode.ContainsKey(mRow.mt_txn_mode))
                        mRow.mt_txn_mode = DicPayMode[mRow.mt_txn_mode];

                    mRow.mt_corp_code = Dr["mt_corp_code"].ToString();
                    mRow.mt_cust_uniq_ref = Dr["mt_cust_uniq_ref"].ToString();
                    mRow.mt_corp_acc_no = Dr["mt_corp_acc_no"].ToString();
                    mRow.mt_value_date = Lib.DatetoStringDisplayformat(Dr["mt_value_date"]);
                    mRow.mt_txn_curr = Dr["mt_txn_curr"].ToString();
                    mRow.mt_txn_amt = Lib.Conv2Decimal(Dr["mt_txn_amt"].ToString());
                    mRow.mt_ben_name = Dr["mt_ben_name"].ToString();
                    mRow.mt_ben_code = Dr["mt_ben_code"].ToString();
                    mRow.mt_ben_acc_no = Dr["mt_ben_acc_no"].ToString();

                    mRow.mt_ben_acc_type = Dr["mt_ben_acc_type"].ToString();
                    if (DicCategory.ContainsKey(mRow.mt_ben_acc_type))
                        mRow.mt_ben_acc_type = DicCategory[mRow.mt_ben_acc_type];

                    mRow.mt_ben_addr1 = Dr["mt_ben_addr1"].ToString();
                    mRow.mt_ben_addr2 = Dr["mt_ben_addr2"].ToString();
                    mRow.mt_ben_addr3 = Dr["mt_ben_addr3"].ToString();
                    mRow.mt_ben_city = Dr["mt_ben_city"].ToString();
                    mRow.mt_ben_state = Dr["mt_ben_state"].ToString();
                    mRow.mt_ben_pin = Dr["mt_ben_pin"].ToString();
                    mRow.mt_ben_ifsc = Dr["mt_ben_ifsc"].ToString();
                    mRow.mt_ben_bank_name = Dr["mt_ben_bank_name"].ToString();
                    mRow.mt_ben_email1 = Dr["mt_ben_email1"].ToString();
                    mRow.mt_ben_email2 = Dr["mt_ben_email2"].ToString();
                    mRow.mt_ben_mob = Dr["mt_ben_mob"].ToString();
                    mRow.mt_pay_type = Dr["mt_pay_type"].ToString();
                    if (DicPaytype.ContainsKey(mRow.mt_pay_type))
                        mRow.mt_pay_type = DicPaytype[mRow.mt_pay_type];
                    mRow.mt_corp_email = Dr["mt_corp_email"].ToString();
                    mRow.mt_transmission_date = Dr["mt_transmission_date"].ToString();

                    mList.Add(mRow);
                }

                Dt_List.Rows.Clear();
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

    }
}
