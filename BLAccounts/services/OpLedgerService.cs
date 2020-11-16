using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts
{
    public class OpLedgerService : BL_Base
    {


        LovService lov = null;
        DataRow lovRow_Local_Currency = null;
        DataRow lovRow_Doc_Prefix = null;

        string report_folder = "";

        string report_caption = "";

        string folderid = "";
        string File_Name = "";
        string branch_code = "";

        private DataTable Dt_Op;

        ExcelFile WB;
        ExcelWorksheet WS = null;

        int iRow = 0;
        int iCol = 0;

        string showCurrency = "N";


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            decimal nDr = 0;
            decimal nCr = 0;
            decimal nBal = 0;


            

            Con_Oracle = new DBConnection();
            List<OpLedger> mList = new List<OpLedger>();
            OpLedger mRow;

            string type = SearchData["type"].ToString();
            string rowtype = SearchData["rowtype"].ToString();

            string subtype = "";

            string company_code = SearchData["company_code"].ToString();
            branch_code = SearchData["branch_code"].ToString();
            string year_code = SearchData["year_code"].ToString();

            if (SearchData.ContainsKey("report_folder"))
                report_folder = SearchData["report_folder"].ToString();
            if (SearchData.ContainsKey("report_caption"))
                report_caption = SearchData["report_caption"].ToString();

            if (SearchData.ContainsKey("folderid"))
                folderid = SearchData["folderid"].ToString();

            if (SearchData.ContainsKey("subtype"))
                subtype = SearchData["subtype"].ToString();

            if (SearchData.ContainsKey("showcurrency"))
                showCurrency = SearchData["showcurrency"].ToString();

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

                if ( subtype == "")
                    sWhere += " and a.jvh_type =  '{ROWTYPE}'";

                sWhere += " ) ";

                if (searchstring != "")
                {
                    if (subtype == "")
                    {
                        sWhere += " and (";
                        sWhere += " jvh_docno like '%{SEARCHSTRING}%' or acc_code like '%{SEARCHSTRING}%' or acc_name like '%{SEARCHSTRING}%' ";
                        sWhere += ")";
                    }
                }

                sWhere = sWhere.Replace("{COMP}", company_code);
                sWhere = sWhere.Replace("{BRANCH}", branch_code);
                sWhere = sWhere.Replace("{FINYEAR}", year_code);
                sWhere = sWhere.Replace("{ROWTYPE}", rowtype);

                sWhere = sWhere.Replace("{SEARCHSTRING}", searchstring);

                if ( type == "EXCEL")
                {

                    report_folder = System.IO.Path.Combine(report_folder, folderid);
                    File_Name = System.IO.Path.Combine(report_folder, folderid);
                    if (Lib.CreateFolder(report_folder))
                    {

                        if (subtype == "DIFFERENCE")
                        {
                            sql = "";
                            sql += " select rec_branch_code, acc_code, acc_name,acgrp_name, op_bal, oi_bal, op_bal - oi_bal as diff from ";
                            sql += " ( ";
                            sql += " select a.rec_branch_code, acc_code, acc_name,acgrp_name, ";
                            sql += " sum(case when jvh_type = 'OP' then jv_debit - jv_credit else 0 end) as op_bal, ";
                            sql += " sum(case when jvh_type in( 'OI', 'OB','OC') then jv_debit - jv_credit else 0 end) as oi_bal ";
                            sql += " from ledgerh a inner ";
                            sql += " join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                            sql += " left join acctm on jv_acc_id = acc_pkid ";
                            sql += " left join acgroupm on acc_group_id = acgrp_pkid ";
                            sql += sWhere;
                            sql += " and jvh_type in ('OP', 'OI', 'OB', 'OC') and acgrp_name in('SUNDRY DEBTORS', 'SUNDRY CREDITORS','EXPENSE PAYABLE','INTERNATIONAL DEBTORS')";
                            sql += " group by  a.rec_branch_code, acc_code, acc_name,acgrp_name ";
                            sql += " ) a ";
                            sql += " where op_bal<> oi_bal ";
                            sql += " order by rec_branch_code,acc_code, acc_name ";

                            Dt_Op = Con_Oracle.ExecuteQuery(sql);
                            Con_Oracle.CloseConnection();
                            ProcessExcelFile2();

                        }
                        else
                        {
                            sql = "";
                            sql += " select  jvh_pkid,jvh_vrno, jvh_type, jvh_docno,jvh_date, jvh_reference, jvh_narration, jvh_debit, jvh_credit,";
                            sql += " jvh_curr_code, jvh_exrate,jvh_tot_famt, ";
                            sql += " acc_code, acc_name, acgrp_name, actype_name ";
                            sql += "  from ledgerh a  left join acctm on jvh_acc_id = acc_pkid ";
                            sql += " left join acgroupm on acc_group_id = acgrp_pkid ";
                            sql += " left join actypem on acc_type_id = actype_pkid ";

                            sql += sWhere;
                            sql += " order by jvh_vrno";

                            Dt_Op = Con_Oracle.ExecuteQuery(sql);
                            Con_Oracle.CloseConnection();
                            ProcessExcelFile();
                        }
                        RetData.Add("list", mList);
                        return RetData;
                    }
                }

                nDr = 0;
                nCr = 0;

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, sum(jvh_Debit) as dr,sum(jvh_credit)  as cr, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM Ledgerh  a ";
                    sql += " left join acctm on jvh_acc_id = acc_pkid ";
                    sql += sWhere;
                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        page_rowcount = Lib.Conv2Integer(Dt_Temp.Rows[0]["total"].ToString());
                        page_count = Lib.Conv2Integer(Dt_Temp.Rows[0]["page_total"].ToString());
                        nDr = Lib.Conv2Decimal(Dt_Temp.Rows[0]["dr"].ToString());
                        nCr = Lib.Conv2Decimal(Dt_Temp.Rows[0]["cr"].ToString());
                        nBal = nDr - nCr;
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
                sql += "  select  jvh_pkid,jvh_vrno, jvh_type, jvh_docno,jvh_date, jvh_reference, jvh_narration, jvh_debit, jvh_credit,";
                sql += " acc_code, acc_name, acgrp_name, actype_name,";
                sql += " jvh_curr_code, jvh_exrate,jvh_tot_famt, ";
                sql += " row_number() over(order by jvh_vrno) rn ";
                sql += "  from ledgerh a  ";
                sql += " left join acctm on jvh_acc_id = acc_pkid ";

                sql += " left join acgroupm on acc_group_id = acgrp_pkid ";
                sql += " left join actypem on acc_type_id = actype_pkid ";

                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow}";
                sql += " order by jvh_vrno";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new OpLedger();
                    mRow.jvh_pkid = Dr["jvh_pkid"].ToString();
                    mRow.jvh_docno = Dr["jvh_docno"].ToString();
                    mRow.jvh_vrno = Lib.Convert2Decimal(Dr["jvh_vrno"].ToString());
                    mRow.jvh_type = Dr["jvh_type"].ToString();
                    mRow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mRow.jvh_reference = Dr["jvh_reference"].ToString();
                    mRow.jvh_narration = Dr["jvh_narration"].ToString();

                    mRow.jvh_acc_code = Dr["acc_code"].ToString();
                    mRow.jvh_acc_name = Dr["acc_name"].ToString();

                    mRow.jvh_group_name = Dr["acgrp_name"].ToString();
                    mRow.jvh_type_name = Dr["actype_name"].ToString();

                    mRow.jvh_curr_code = Dr["jvh_curr_code"].ToString();
                    mRow.jvh_exrate = Lib.Conv2Decimal(Dr["jvh_exrate"].ToString());
                    mRow.jvh_ftotal = Lib.Conv2Decimal(Dr["jvh_tot_famt"].ToString());
                    

                    mRow.jvh_debit = Lib.Conv2Decimal(Dr["jvh_debit"].ToString());
                    mRow.jvh_credit = Lib.Conv2Decimal(Dr["jvh_credit"].ToString());

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
            RetData.Add("dr", nDr.ToString());
            RetData.Add("cr", nCr.ToString());
            RetData.Add("bal", nBal.ToString());
            RetData.Add("list", mList);

            return RetData;
        }

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            OpLedger mRow = new OpLedger();
            string lockedmsg = "";
            string id = SearchData["pkid"].ToString();
            try
            {
                Con_Oracle = new DBConnection();



                DataTable Dt_Rec = new DataTable();

                sql = "";
                sql += " select ";
                sql += " jvh_pkid, jvh_date, jvh_year, jvh_type, jvh_vrno, jvh_docno,jvh_rec_source, jvh_location,";
                sql += " jvh_acc_id,acc_code as jvh_acc_code, acc_name as jvh_acc_name, ";
                sql += " jvh_exrate,jvh_reference ,jvh_reference_date,jvh_narration,a.rec_category, ";
                sql += " jvh_debit, jvh_credit, jvh_curr_id, jvh_curr_code,jvh_cc_category, jvh_remarks, ";
                sql += " jvh_tot_famt,jvh_tot_amt, ";
                sql += " jv_drcr,jv_bank,jv_branch,jv_chqno,jv_due_date,jv_remarks,";
                sql += " jvh_edit_code, jvh_edit_date, a.rec_locked,a.rec_company_code,a.rec_branch_code ";
                sql += " from ledgerh a ";
                sql += " inner join ledgert b on a.jvh_pkid = b.jv_parent_id ";
                sql += " left join acctm on jv_acc_id = acc_pkid ";
                sql += " where  a.jvh_pkid ='" + id + "'";


                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new OpLedger();
                    mRow.jvh_pkid = Dr["jvh_pkid"].ToString();
                    mRow.jvh_vrno = Lib.Conv2Integer(Dr["jvh_vrno"].ToString());
                    mRow.jvh_type = Dr["jvh_type"].ToString();
                    mRow.jvh_docno = Dr["jvh_docno"].ToString();
                    mRow.jvh_date = Lib.DatetoString(Dr["jvh_date"]);
                    mRow.jvh_year = Lib.Conv2Integer(Dr["jvh_year"].ToString());
                    mRow.jvh_reference = Dr["jvh_reference"].ToString();
                    mRow.jvh_reference_date = Lib.DatetoString(Dr["jvh_reference_date"]);
                    mRow.jvh_rec_source = Dr["jvh_rec_source"].ToString();
                    mRow.jvh_type = Dr["jvh_type"].ToString();

                    mRow.jvh_acc_id = Dr["jvh_acc_id"].ToString();
                    mRow.jvh_acc_code = Dr["jvh_acc_code"].ToString();
                    mRow.jvh_acc_name = Dr["jvh_acc_name"].ToString();

                    mRow.jvh_curr_id = Dr["jvh_curr_id"].ToString();
                    mRow.jvh_curr_code = Dr["jvh_curr_code"].ToString();
                    mRow.jvh_exrate = Lib.Conv2Decimal(Dr["jvh_exrate"].ToString());

                    mRow.jvh_location = Dr["jvh_location"].ToString();
                    mRow.jvh_remarks = Dr["jvh_remarks"].ToString();

                    mRow.rec_category = Dr["rec_category"].ToString();
                    mRow.jvh_cc_category = Dr["jvh_cc_category"].ToString();

                    mRow.jvh_debit = Lib.Conv2Decimal(Dr["jvh_debit"].ToString());
                    mRow.jvh_credit = Lib.Conv2Decimal(Dr["jvh_credit"].ToString());


                    mRow.jvh_ftotal = Lib.Conv2Decimal(Dr["jvh_tot_famt"].ToString());
                    mRow.jvh_total = Lib.Conv2Decimal(Dr["jvh_tot_amt"].ToString());

                    mRow.jvh_drcr = Dr["jv_drcr"].ToString();
                    mRow.jvh_bank = Dr["jv_bank"].ToString();
                    mRow.jvh_branch = Dr["jv_branch"].ToString();
                    mRow.jvh_chqno = Lib.Conv2Integer(Dr["jv_chqno"].ToString());
                    mRow.jvh_due_date = Lib.DatetoString(Dr["jv_due_date"]);
                    mRow.jvh_remarks = Dr["jv_remarks"].ToString();

                    mRow.jvh_allocation_found = false;

                    mRow.rec_locked = false;
                    mRow.jvh_edit_code = Dr["jvh_edit_code"].ToString();
                    mRow.jvh_edit_date = Dr["jvh_edit_date"].ToString();
                    if (Dr["rec_locked"].ToString() == "Y" && Dr["jvh_edit_date"].ToString() != System.DateTime.Today.ToString("yyyyMMdd"))
                    {
                        mRow.jvh_edit_code = "";
                        mRow.rec_locked = true;
                    }

                    string JvhDate = Lib.StringToDate(Dr["jvh_date"]);//Transaction Locking for opening ref to LOCK JV DATE
                    lockedmsg = Lib.IsDateLocked(JvhDate, "JV",
                        Dr["rec_company_code"].ToString(),
                        Dr["rec_branch_code"].ToString(), Dr["jvh_year"].ToString());
                    break;
                }
                sql = "";
                //sql = " select xref_jvh_id from ledgerxref where  (xref_dr_jv_id ='{PKID}' or xref_cr_jv_id ='{PKID}')  and xref_jvh_id <> '{PKID}' ";
                sql = " select xref_jvh_id from ledgerxref where  (xref_dr_jv_id ='{PKID}' or xref_cr_jv_id ='{PKID}')  ";

                sql = sql.Replace("{PKID}", id);
                if (Con_Oracle.IsRowExists(sql))
                {
                    mRow.jvh_allocation_found = true;
                }
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


        public string AllValid(OpLedger Record)
        {
            string str = "";
            try
            {
                //Transaction Locking for opening ref to LOCK JV DATE
                string jvhdate = Lib.StringToDate(Record.jvh_reference_date.ToString());

                str = Lib.IsDateLocked(jvhdate, "JV",
                        Record._globalvariables.comp_code,
                        Record._globalvariables.branch_code, Record._globalvariables.year_code);

                if (Record.jvh_type == "OI" && Record.rec_mode == "EDIT")
                {
                    sql = "";
                    sql = " select * from ledgerxref where xref_dr_jv_id ='{PKID}' or xref_cr_jv_id ='{PKID}' ";
                    sql = sql.Replace("{PKID}", Record.jvh_pkid);
                    if (Con_Oracle.IsRowExists(sql))
                    {
                        str += "| Allocation Exist, Cannot Edit";
                    }
                }

                if (Record.jvh_type == "OP" && !Lib.IsInFinYear(Record.jvh_reference_date, Record._globalvariables.year_start_date, Record._globalvariables.year_end_date))
                {
                    str += "| Invalid Date";
                }

                //if ((Record.jvh_type == "OB" || Record.jvh_type == "OI") && !Lib.IsBeforeFinYear(Record.jvh_reference_date, Record._globalvariables.year_start_date))
                //{
                //    str += "| Invalid Date";
                //}

                lovRow_Local_Currency = lov.getSettings(Record._globalvariables.comp_code, "LOCAL-CURRENCY");
                lovRow_Doc_Prefix = lov.getSettings(Record._globalvariables.branch_code, "DOC-PREFIX");

                if (lovRow_Doc_Prefix == null)
                    str += "| Doc Prefix Not Found";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(OpLedger Record)
        {
            int iCtr = 0;

            DataTable Dt_Hbl = new DataTable();
            Boolean bOk = false;
            int iVrNo = 0;
            string DocNo = "";
            string doc_prefix = "";
            lov = new LovService();

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
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

                if (Record.jvh_type == "OI")
                {
                    DocNo = Record.jvh_reference;
                    if (DocNo.Length > 30)
                        DocNo = DocNo.Substring(0, 30);

                    sql = "";
                    sql = "select jvh_pkid from (";
                    sql += "select jvh_pkid  from ledgerh a where a.rec_company_code = '" + Record._globalvariables.comp_code + "' ";
                    sql += " and a.rec_branch_code = '" + Record._globalvariables.branch_code + "'";
                    sql += " and a.jvh_docno = '{DOCNO}'  ";
                    sql += ") a where jvh_pkid <> '{PKID}'";

                    sql = sql.Replace("{DOCNO}", DocNo);
                    sql = sql.Replace("{PKID}", Record.jvh_pkid);

                    if (Con_Oracle.IsRowExists(sql))
                        throw new Exception("Invoice Reference Exists");
                }

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("Ledgerh", Record.rec_mode, "jvh_pkid", Record.jvh_pkid);
                if (Record.rec_mode == "ADD")
                {
                    Record.jvh_docno = DocNo.ToString();
                    Rec.InsertNumeric("jvh_vrno", iVrNo.ToString());
                    Rec.InsertString("jvh_docno", DocNo.ToString());
                    Rec.InsertString("jvh_type", Record.jvh_type);
                    Rec.InsertString("jvh_subtype", Record.jvh_type);
                    if (Record.jvh_type == "OP")
                        Rec.InsertString("jvh_posted", "Y");
                    else
                        Rec.InsertString("jvh_posted", "N");


                    Rec.InsertString("jvh_location", Record.jvh_location);

                    Rec.InsertString("jvh_rec_source", "JV");
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
                    if (Record.jvh_type == "OI")
                        Rec.InsertString("jvh_docno", DocNo.ToString());
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    Rec.InsertFunction("rec_edited_date", "SYSDATE");
                }

                Rec.InsertDate("jvh_date", Record.jvh_reference_date);

                Rec.InsertString("jvh_reference", Record.jvh_reference);

                if (Record.jvh_remarks != null)
                    Rec.InsertString("jvh_remarks", Record.jvh_remarks);

                Rec.InsertDate("jvh_reference_date", Record.jvh_reference_date);

                Rec.InsertString("jvh_acc_id", Record.jvh_acc_id);

                Rec.InsertString("jvh_curr_id", Record.jvh_curr_id);
                Rec.InsertString("jvh_curr_code", Record.jvh_curr_code);

                Rec.InsertNumeric("jvh_exrate", Record.jvh_exrate.ToString());
                Rec.InsertNumeric("jvh_tot_famt", Record.jvh_ftotal.ToString());
                Rec.InsertNumeric("jvh_net_famt", Record.jvh_ftotal.ToString());
                Rec.InsertNumeric("jvh_tot_amt", Record.jvh_total.ToString());
                Rec.InsertNumeric("jvh_net_amt", Record.jvh_total.ToString());
                Rec.InsertString("jvh_narration", Record.jvh_narration);

                if (Record.jvh_drcr == "DR")
                {
                    Rec.InsertNumeric("jvh_debit", Record.jvh_debit.ToString());
                    Rec.InsertNumeric("jvh_credit", "0");
                }
                else
                {
                    Rec.InsertNumeric("jvh_debit", "0");
                    Rec.InsertNumeric("jvh_credit", Record.jvh_credit.ToString());
                }
                Rec.InsertString("rec_category", Record.rec_category);
                Rec.InsertString("jvh_cc_category", Record.jvh_cc_category);


                Rec.InsertString("jvh_sez", "N");
                Rec.InsertString("jvh_rc", "N");
                Rec.InsertString("jvh_gst", "N");

                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                sql = "delete from  ledgert where jv_parent_id ='" + Record.jvh_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);


                SaveLedgerRecord(
                    "ADD",
                    Record.jvh_pkid, Record.jvh_pkid, Record.jvh_acc_id, Record.jvh_acc_name, false,
                    Record.jvh_curr_id, "", 1, Record.jvh_ftotal, false, false,
                    Record.jvh_drcr, Record.jvh_ftotal, Record.jvh_exrate, Record.jvh_total, 0, "",
                    0, 0, 0,
                    0, 0, 0, 0,
                    Record.jvh_debit, Record.jvh_credit, Record.jvh_total,
                    Record.jvh_ftotal, 0, 0, 0, 0, Record.jvh_ftotal,
                    "NA", Record.jvh_bank, Record.jvh_branch,
                    Record.jvh_chqno, Record.jvh_due_date,
                    "", "", "", Record.jvh_remarks, Record.jvh_type,
                    Record._globalvariables.year_code, Record._globalvariables.comp_code, Record._globalvariables.branch_code,
                    "", 0, 0, "", 0, "",
                    "", "", iCtr
                    );

                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();


                string str = "DR " + Record.jvh_debit.ToString() + ", CR " + Record.jvh_credit.ToString() + ", " + Record.jvh_acc_name;
                Lib.AuditLog("OPENING", Record.jvh_type, Record.rec_mode, Record._globalvariables.comp_code, Record._globalvariables.branch_code, Record._globalvariables.user_code, Record.jvh_pkid, Record.jvh_docno, str);


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
            decimal cgst_amt, decimal sgst_amt, decimal igst_amt, decimal gst_amt,
            decimal debit, decimal credit, decimal net_total,
            decimal total_fc, decimal cgst_famt, decimal sgst_famt, decimal igst_famt, decimal gst_famt, decimal net_ftotal,
            string doc_type, string bank, string branch, int chqno, string due_date, string pay_reason,
            string supp_docs, string paid_to, string remarks, string row_type, string year_code, string comp_code, string branch_code,
            string pan_id, decimal tds_per, decimal tds_amt, string tan_id, decimal gross_bill_amt, string tan_party_id, string recon_by, string recon_date,
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
            if (row_type == "OP")
                Rec.InsertString("jv_posted", "Y");
            else
                Rec.InsertString("jv_posted", "N");

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

            if (drcr == "DR")
                Rec.InsertString("jv_row_type", "DR-LEDGER");
            else if (drcr == "CR")
                Rec.InsertString("jv_row_type", "CR-LEDGER");

            Rec.InsertNumeric("jv_year", year_code);

            Rec.InsertString("rec_deleted", "N");

            Rec.InsertString("rec_company_code", comp_code);
            Rec.InsertString("rec_branch_code", branch_code);

            Rec.InsertNumeric("jv_ctr", iCtr.ToString());

            Rec.InsertString("jv_recon_by", recon_by);
            Rec.InsertDate("jv_recon_date", recon_date);

            Rec.InsertString("jv_pan_id", pan_id);
            Rec.InsertNumeric("jv_tds_rate", tds_per.ToString());
            Rec.InsertNumeric("jv_tds_gross_amt", tds_amt.ToString());
            Rec.InsertString("jv_tan_id", tan_id);
            Rec.InsertNumeric("jv_gross_bill_amt", gross_bill_amt.ToString());
            Rec.InsertString("jv_tan_party_id", tan_party_id);

            sql = Rec.UpdateRow();
            Con_Oracle.ExecuteNonQuery(sql);
        }

        public Dictionary<string, object> DeleteRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string id = SearchData["pkid"].ToString();
            string comp_code = SearchData["comp_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string user_code = SearchData["user_code"].ToString();
            string jvh_type = SearchData["jvh_type"].ToString();
            string jvh_docno = SearchData["jvh_docno"].ToString();
            string jvh_narration = SearchData["jvh_narration"].ToString();

            try
            {
                Con_Oracle = new DBConnection();

                sql = "";
                sql = " select xref_jvh_id from ledgerxref where  (xref_dr_jv_id ='{PKID}' or xref_cr_jv_id ='{PKID}')  and xref_jvh_id <> '{PKID}' ";
                sql = sql.Replace("{PKID}", id);
                if (Con_Oracle.IsRowExists(sql))
                {
                    Con_Oracle.CloseConnection();
                    throw new Exception("Cannot Delete, Allocation Exists");
                }

                Con_Oracle.BeginTransaction();
                sql = " Delete from ledgert where jv_parent_id ='" + id + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                sql = " Delete from ledgerh where jvh_pkid ='" + id + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                Lib.AuditLog("OP", jvh_type, "DELETE", comp_code,branch_code, user_code, id, jvh_docno, jvh_narration);

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
            return RetData;
        }


        private void ProcessExcelFile()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;


            decimal nDr = 0;
            decimal nCr = 0;
            decimal nBal = 0;

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];

            WS.Columns[0].Width = 256;
            WS.Columns[1].Width = 256 * 10;
            WS.Columns[2].Width = 256 * 10;
            WS.Columns[3].Width = 256 * 15;
            WS.Columns[4].Width = 256 * 15;
            WS.Columns[5].Width = 256 * 15;
            WS.Columns[6].Width = 256 * 40;
            WS.Columns[7].Width = 256 * 15;
            WS.Columns[8].Width = 256 * 15;
            WS.Columns[9].Width = 256 * 15;
            WS.Columns[10].Width = 256 * 15;
            WS.Columns[11].Width = 256 * 15;
            WS.Columns[12].Width = 256 * 15;

            WS.Columns[7].Style.NumberFormat = "#,0.00";
            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";
            WS.Columns[10].Style.NumberFormat = "#,0.00";
            WS.Columns[11].Style.NumberFormat = "#,0.00";
            WS.Columns[12].Style.NumberFormat = "#,0.00";

            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);


            Lib.WriteData(WS, iRow++, iCol, report_caption , Color.Brown, true, "", "L", "Calibri", 12, false);

            nDr = 0; nCr = 0;

            _Border = "TB";

            _Bold = true;

            Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "REF#", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "GROUP", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            if (showCurrency == "Y")
            {
                Lib.WriteData(WS, iRow, iCol++, "CURR", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AMT", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXRATE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            }

            Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
            

            _Border = "";
            foreach (DataRow Dr in Dt_Op.Rows)
            {
                _Bold = false;
                _Color = Color.Black;
                iRow++; iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_vrno"], _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_type"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_reference"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["jvh_date"]), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acc_code"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acc_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, Dr["acgrp_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["actype_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


                if (showCurrency == "Y")
                {
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_curr_code"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_tot_famt"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["jvh_exrate"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00000", true);
                }


                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_debit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["jvh_credit"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                

                nDr += Lib.Convert2Decimal(Dr["jvh_debit"].ToString());
                nCr += Lib.Convert2Decimal(Dr["jvh_credit"].ToString());
                nBal = nDr - nCr ;
            }

            iRow++;
            iCol = 1;
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            if (showCurrency == "Y")
            {
                Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            }

            Lib.WriteData(WS, iRow, iCol++, nDr, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            Lib.WriteData(WS, iRow, iCol++, nCr, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            


            iRow++;
            iCol = 1;
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);


            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);

            if (showCurrency == "Y")
            {
                Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            }

            Lib.WriteData(WS, iRow, iCol++, nBal, _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            Lib.WriteData(WS, iRow, iCol++, "", _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);



            WB.SaveXls(File_Name + ".xls");
        }


        private void ProcessExcelFile2()
        {
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;


            decimal nDr = 0;
            decimal nCr = 0;
            decimal nBal = 0;

            string sName = "Report";
            WB = new ExcelFile();
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];

            WS.Columns[0].Width = 256;
            WS.Columns[1].Width = 256 * 10;
            WS.Columns[2].Width = 256 * 10;
            WS.Columns[3].Width = 256 * 10;
            WS.Columns[4].Width = 256 * 15;
            WS.Columns[5].Width = 256 * 15;
            WS.Columns[6].Width = 256 * 15;
            WS.Columns[7].Width = 256 * 40;
            WS.Columns[8].Width = 256 * 15;
            WS.Columns[9].Width = 256 * 15;
            WS.Columns[10].Width = 256 * 15;

            WS.Columns[8].Style.NumberFormat = "#,0.00";
            WS.Columns[9].Style.NumberFormat = "#,0.00";
            WS.Columns[10].Style.NumberFormat = "#,0.00";

            iRow = 1; iCol = 1;

            iRow = Lib.WriteAddress(WS, branch_code, iRow, iCol);


            Lib.WriteData(WS, iRow++, iCol, report_caption, Color.Brown, true, "", "L", "Calibri", 12, false);

            nDr = 0; nCr = 0;

            _Border = "TB";

            _Bold = true;


            Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "GROUP", _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "OP", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "INVOICE", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "DIFFERENCE", _Color, _Bold, _Border, "R", "", _Size, false, 325, "", true);

            _Border = "";
            foreach (DataRow Dr in Dt_Op.Rows)
            {
                _Bold = false;
                _Color = Color.Black;
                iRow++; iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, Dr["acc_code"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acc_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["acgrp_name"].ToString(), _Color, _Bold, _Border, "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["op_bal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["oi_bal"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["diff"], _Color, _Bold, _Border, "R", "", _Size, false, 325, "#,0.00", true);
            }


            WB.SaveXls(File_Name + ".xls");
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


