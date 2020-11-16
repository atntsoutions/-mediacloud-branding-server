using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DataBase;
using DataBase_Oracle.Connections;

using BLAccounts;

namespace BLCosting
{
    public class PostingService : BL_Base
    {
        string BrErrorMessage = "";
        string HoErrorMessage = "";


        string HO_VRNO = "";
        string BR_VRNO = "";
        string BR_INVNO = "";


        string frt_jv_id = "";
        Boolean set_frt_id = false;

        int GLOB_CTR = 0;
        string GLOB_JVID = "";

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            string mType = "";
            string str = "";


            string brcode = "";

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Posting mRow = new Posting();


            DataTable Dt_Sac = new DataTable();

            string lockedmsg = "";
            string id = SearchData["pkid"].ToString();

            string str1 = "";
            string str2 = "";
            string str3 = "";

            string nAmt = "";

            try
            {
                DataTable Dt_Rec = new DataTable();


                sql = "";
                sql += " select ";
                sql += " cost_pkid,cost_date,mbl.hbl_type as mbl_type, mbl.hbl_pkid as mbl_pkid, mbl.hbl_book_cntr,hbl_bl_no, cost_type,cost_refno,cost_folderno,cost_currency_id,c.param_code as curr_code, cost_exrate, ";
                sql += " cost_drcr, cost_drcr_amount, cost_drcr_amount_inr, cost_jv_posted, cost_jv_agent_id, cost_jv_agent_br_id, cost_jv_ho_id, cost_jv_br_id,cost_jv_br_inv_id, ";
                sql += " agnt.acc_code as agent_code,agnt.acc_name as agent_name, cost_cntr,";
                sql += " a.rec_branch_code,a.rec_category,cost_source,cost_category, a.cost_type, a.cost_prefix,cost_remarks ";

                sql += " from costingm a ";
                sql += " left join hblm mbl on a.cost_mblid = mbl.hbl_pkid ";
                sql += " left join param c on a.cost_currency_id = c.param_pkid ";
                sql += " left join acctm agnt on a.cost_jv_agent_id = agnt.acc_pkid ";

                sql += " where  a.cost_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Posting();

                    mRow.jv_cost_id = id;
                    mRow.mbl_pkid = Dr["mbl_pkid"].ToString();
                    mRow.mbl_type = Dr["mbl_type"].ToString();
                    mRow.jv_reference = Dr["cost_refno"].ToString();

                    mRow.jv_date = Lib.DatetoString(Dr["cost_date"]);


                    mType = Dr["REC_CATEGORY"].ToString();
                    if (Dr["cost_source"].ToString() == "SEA EXPORT COSTING")
                        mType = "SEA EXPORT";
                    else if (Dr["cost_source"].ToString() == "AIR EXPORT COSTING")
                        mType = "AIR EXPORT";
                    else if (Dr["cost_source"].ToString() == "DRCR ISSUE")
                    {
                        if (Dr["cost_category"].ToString().Contains("SEA EXPORT"))
                            mType = "SEA EXPORT";
                        else if (Dr["cost_category"].ToString().Contains("AIR EXPORT"))
                            mType = "AIR EXPORT";
                        else if (Dr["cost_category"].ToString().Contains("SEA IMPORT"))
                            mType = "SEA IMPORT";
                        else if (Dr["cost_category"].ToString().Contains("AIR IMPORT"))
                            mType = "AIR IMPORT";
                    }
                    else if (Dr["cost_source"].ToString() == "AGENT INVOICE")
                    {
                        if (Dr["cost_category"].ToString().Contains("SEA EXPORT"))
                            mType = "SEA EXPORT";
                        else if (Dr["cost_category"].ToString().Contains("AIR EXPORT"))
                            mType = "AIR EXPORT";
                        else if (Dr["cost_category"].ToString().Contains("SEA IMPORT"))
                            mType = "SEA IMPORT";
                        else if (Dr["cost_category"].ToString().Contains("AIR IMPORT"))
                            mType = "AIR IMPORT";
                    }

                    if (Dr["cost_category"].ToString().Contains("GENERAL JOB"))
                    {
                        mRow.mbl_type = "GEN-JOB";
                    }

                    /*
                    if (mType == "SEA EXPORT")
                        str = Dr["hbl_book_cntr"].ToString();
                    else 
                        str = Dr["hbl_bl_no"].ToString();
                    */
                    str = Dr["cost_cntr"].ToString();
                    mRow.jv_remarks = str;
                    if (str.Length > 300)
                        mRow.jv_remarks = str.Substring(0, 300);


                    if (Dr["cost_type"].ToString() == "SEA")
                        str = "REF# " + Dr["cost_refno"].ToString() + ",CNTR# " + Dr["cost_cntr"].ToString() + ", BL# " + Dr["hbl_bl_no"].ToString() + ", AGENT-" + Dr["agent_name"].ToString();
                    else
                        str = "REF# " + Dr["cost_refno"].ToString() + ",BL# " + Dr["hbl_bl_no"].ToString() + " AGENT-" + Dr["agent_name"].ToString();

                    if (!Dr["cost_remarks"].Equals(DBNull.Value))
                        str += "," + Dr["cost_remarks"].ToString();

                    mRow.jv_narration = str;
                    if (str.Length > 300)
                        mRow.jv_narration = str.Substring(0, 300);


                    if (Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString()) > 0)
                        mRow.jv_drcr = "DR";
                    if (Lib.Conv2Decimal(Dr["cost_drcr_amount"].ToString()) < 0)
                        mRow.jv_drcr = "CR";

                    mRow.jv_exrate = Lib.Convert2Decimal(Dr["cost_exrate"].ToString());
                    mRow.jv_ftotal = Lib.Convert2Decimal(Dr["cost_drcr_amount"].ToString());
                    mRow.jv_total = Lib.Convert2Decimal(Dr["cost_drcr_amount_inr"].ToString());

                    if (Dr["cost_source"].ToString() == "DRCR ISSUE" || Dr["cost_source"].ToString() == "AGENT INVOICE")
                    {
                        if (Dr["cost_drcr"].ToString() == "DR")
                        {
                            mRow.jv_drcr = "DR";
                        }
                        if (Dr["cost_drcr"].ToString() == "CR")
                        {
                            mRow.jv_drcr = "CR";
                            mRow.jv_ftotal = Lib.Convert2Decimal(Dr["cost_drcr_amount"].ToString()) * -1;
                            mRow.jv_total = Lib.Convert2Decimal(Dr["cost_drcr_amount_inr"].ToString()) * -1;
                        }
                    }
                    mRow.jv_agent_id = Dr["cost_jv_agent_id"].ToString();
                    mRow.jv_agent_code = Dr["agent_code"].ToString();
                    mRow.jv_agent_name = Dr["agent_name"].ToString();

                    mRow.jv_agent_br_id = Dr["cost_jv_agent_br_id"].ToString();


                    mRow.jv_curr_code = Dr["curr_code"].ToString();
                    mRow.jv_curr_id = Dr["cost_currency_id"].ToString();


                    mRow.jv_ho_id = "A68D3A49-B4D7-2B52-0D01-D49B54E23703";
                    mRow.jv_ho_code = "CPLHO";
                    mRow.jv_ho_name = "CARGOMAR PVT LTD H.O.";

                    mRow.jv_br_id = "";
                    mRow.jv_br_code = "";
                    mRow.jv_br_name = "";

                    AssignBrCode(Dr["rec_branch_code"].ToString(), mRow);

                    brcode = Dr["rec_branch_code"].ToString();

                    AssignFrtCode(mType, mRow);

                    Dt_Sac = new DataTable();
                    Dt_Sac = Con_Oracle.ExecuteQuery("select acc_sac_id from acctm where acc_pkid = '" + mRow.jv_frt_id + "'");
                    foreach (DataRow Dr1 in Dt_Sac.Rows)
                    {
                        mRow.jv_sac_id = Dr1["acc_sac_id"].ToString();
                    }

                    mRow.jv_ho_record_pkid = Dr["cost_jv_ho_id"].ToString();
                    mRow.jv_br_record_pkid = Dr["cost_jv_br_id"].ToString();
                    mRow.jv_br_inv_record_pkid = Dr["cost_jv_br_inv_id"].ToString();


                    mRow.jv_cost_prefix = Dr["cost_prefix"].ToString();
                    mRow.rec_category = Dr["rec_category"].ToString();
                    mRow.rec_category = mType;

                    mRow.rec_mode = "";
                    mRow.jv_posted = "N";

                    if (Dr["cost_jv_posted"].ToString() == "Y")
                        mRow.jv_posted = "Y";

                    // pls remove after posting all kolkatta books
                    if (brcode == "KOLAF")
                        mRow.jv_posted = "N";

                    mRow.jv_posted_details = "";
                    break;
                }


                sql = " select std_pkid from stmtd where std_jv_entityid = '" + mRow.jv_ho_record_pkid + "'";
                if (Con_Oracle.IsRowExists(sql))
                    mRow.jv_posted = "Y";

                // pls remove after posting all kolkatta books
                if (brcode == "KOLAF")
                    mRow.jv_posted = "N";


                sql = "";
                sql += " select jvh_type,jvh_date,jvh_year, a.rec_company_code, a.rec_branch_code ";
                sql += " ,jvh_vrno, jvh_docno, acc_pkid, acc_code, acc_name, jv_drcr";
                sql += " , jv_debit, jv_credit, jv_ftotal, jv_ctr, jv_drcr ";
                sql += " from ledgerh a ";
                sql += " inner join ledgert b on a.jvh_pkid = jv_parent_id ";
                sql += " inner join acctm on jv_acc_id = acc_pkid ";
                sql += " where jvh_pkid in ('{ID1}','{ID2}','{ID3}') ";
                sql += " order by jvh_pkid, jv_ctr ";

                sql = sql.Replace("{ID1}", mRow.jv_br_inv_record_pkid);
                sql = sql.Replace("{ID2}", mRow.jv_br_record_pkid);
                sql = sql.Replace("{ID3}", mRow.jv_ho_record_pkid);


                DataTable Dt_test = new DataTable();
                Dt_test = Con_Oracle.ExecuteQuery(sql);


                string JvhDate = "", JvhType = "", JvhBranch = "", JvhCompany = "", JvhYear = "";
                foreach (DataRow Dr in Dt_test.Rows)
                {
                    JvhDate = Lib.StringToDate(Dr["jvh_date"]);
                    JvhType = Dr["jvh_type"].ToString();
                    JvhYear = Dr["jvh_year"].ToString();


                    if (str1 == "")
                        str1 = " IN : ";
                    if (str2 == "")
                        str2 = " HO : ";
                    if (str3 == "")
                        str3 = " BR : ";

                    if (Dr["JVH_TYPE"].ToString() == "IN-ES")
                    {
                        if (Lib.Conv2Decimal(Dr["jv_debit"].ToString()) > 0)
                        {
                            str1 += " DR : " + Dr["acc_code"].ToString();
                        }
                        if (Lib.Conv2Decimal(Dr["jv_credit"].ToString()) > 0)
                        {
                            str1 += " CR : " + Dr["acc_code"].ToString();
                        }
                    }
                    else if (Dr["REC_BRANCH_CODE"].ToString() == "HOCPL")
                    {
                        if (Lib.Conv2Decimal(Dr["jv_debit"].ToString()) > 0)
                        {
                            str2 += " DR : " + Dr["acc_code"].ToString();
                        }
                        if (Lib.Conv2Decimal(Dr["jv_credit"].ToString()) > 0)
                        {
                            str2 += " CR : " + Dr["acc_code"].ToString();
                        }
                    }
                    else
                    {
                        JvhCompany = Dr["rec_company_code"].ToString();
                        JvhBranch = Dr["rec_branch_code"].ToString();

                        if (Lib.Conv2Decimal(Dr["jv_debit"].ToString()) > 0)
                        {
                            str3 += " DR : " + Dr["acc_code"].ToString();
                        }
                        if (Lib.Conv2Decimal(Dr["jv_credit"].ToString()) > 0)
                        {
                            str3 += " CR : " + Dr["acc_code"].ToString();
                        }
                    }

                    nAmt = Dr["jv_ftotal"].ToString();

                }
                Dt_test.Rows.Clear();

                if (str1 == " IN : ")
                    str1 = "";

                mRow.jv_posted_details = str1 + str2 + str3 + " Amt " + nAmt.ToString();


                if (JvhCompany != "" && JvhBranch != "")
                    lockedmsg = Lib.IsDateLocked(JvhDate, JvhType, JvhCompany, JvhBranch, JvhYear);

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

        private void AssignFrtCode(string mType, Posting mRow)
        {
            if (mType == "SEA EXPORT")
            {
                mRow.jv_frt_id = "657943E1-B3D1-466F-8C91-F524D0F0512B";
                mRow.jv_frt_code = "1105001";
                mRow.jv_frt_name = "OCEAN FREIGHT CHARGES";
            }
            if (mType == "AIR EXPORT")
            {
                mRow.jv_frt_id = "E1556A5A-71B0-49C0-A7F8-D5F581C628EE";
                mRow.jv_frt_code = "1205001";
                mRow.jv_frt_name = "AIR FRIEGHT CHARGES";
            }
            if (mType == "SEA IMPORT")
            {
                mRow.jv_frt_id = "F7FA9F67-A627-4AB6-A285-CEE605F12A27";
                mRow.jv_frt_code = "1305001";
                mRow.jv_frt_name = "OCEAN FREIGHT CHARGES - IMPORT";
            }
            if (mType == "AIR IMPORT")
            {
                mRow.jv_frt_id = "53D15175-C40C-492E-A434-AAE99314585C";
                mRow.jv_frt_code = "1405001";
                mRow.jv_frt_name = "AIR FREIGHT CHARGES IMPORT";
            }

        }

        private void AssignBrCode(string brcode, Posting mRow)
        {
            if (brcode == "ABDSF")
            {
                mRow.jv_br_id = "ABF700D7-B261-93B9-1B58-05E7B03B1330";
                mRow.jv_br_code = "CPLAHM";
                mRow.jv_br_name = "CARGOMAR PVT LTD AHMEDABAD";
            }
            if (brcode == "BLRAF")
            {
                mRow.jv_br_id = "A3D57FC1-670A-7332-6E41-F6093A57F902";
                mRow.jv_br_code = "CPLBLR";
                mRow.jv_br_name = "CARGOMAR PVT LTD BANGALORE";
            }
            if (brcode == "CHNAF")
            {
                mRow.jv_br_id = "90B2F36A-C883-2C4B-218B-90B14A0DA221";
                mRow.jv_br_code = "CPLMDA";
                mRow.jv_br_name = "CARGOMAR PVT LTD CHENNAI AIR";
            }
            if (brcode == "CHNSF")
            {
                mRow.jv_br_id = "D9D3E36A-B53B-F456-3DC1-3E8EDD5D7334";
                mRow.jv_br_code = "CPLMDS";
                mRow.jv_br_name = "CARGOMAR PVT LTD CHENNAI SEA";
            }
            if (brcode == "COKAF")
            {
                mRow.jv_br_id = "7D061CED-E704-E696-EA08-7D0C50E88F89";
                mRow.jv_br_code = "CPLCHA";
                mRow.jv_br_name = "CARGOMAR PVT LTD KOCHI AIR";
            }
            if (brcode == "COKSF")
            {
                mRow.jv_br_id = "755FF53F-4B3B-8177-8D53-79CDB59A59C4";
                mRow.jv_br_code = "CPLCHS";
                mRow.jv_br_name = "CARGOMAR PVT LTD KOCHI SEA";
            }
            if (brcode == "DELAF")
            {
                mRow.jv_br_id = "AA1C2924-A8E8-9011-4A1F-9D99529DDA2D";
                mRow.jv_br_code = "CPLDLA";
                mRow.jv_br_name = "CARGOMAR PVT LTD DELHI AIR";
            }
            if (brcode == "DELSF")
            {
                mRow.jv_br_id = "38230166-4861-5B7C-7194-51DF5E866E28";
                mRow.jv_br_code = "CPLDLS";
                mRow.jv_br_name = "CARGOMAR PVT LTD DELHI SEA";
            }
            if (brcode == "KOLAF")
            {
                mRow.jv_br_id = "6FEDB24E-7A84-01A0-126B-095803A55BB6";
                mRow.jv_br_code = "CKPL";
                mRow.jv_br_name = "CARGOMAR (KOLKATTA)PVT LTD";
            }
            if (brcode == "MBISF")
            {
                mRow.jv_br_id = "5B8303AC-B04B-A03E-527D-74DCF7B7848D";
                mRow.jv_br_code = "CPLMBS";
                mRow.jv_br_name = "CARGOMAR PVT LTD MUMBAI SEA";
            }

            if (brcode == "MBYAF")
            {
                mRow.jv_br_id = "1239CA89-9630-515F-2A45-D41751584B6F";
                mRow.jv_br_code = "CPLMBA";
                mRow.jv_br_name = "CARGOMAR PVT LTD MUMBAI AIR";
            }
            if (brcode == "SEZSF")
            {
                mRow.jv_br_id = "02D5033A-9FF6-0E4B-E206-7D1E770E39DA";
                mRow.jv_br_code = "CPLCSEZ";
                mRow.jv_br_name = "CARGOMAR PVT LTD CEPZ";
            }
            if (brcode == "TUTSF")
            {
                mRow.jv_br_id = "E51A03FC-5432-F8DC-E174-2036C030356D";
                mRow.jv_br_code = "CPLTUT";
                mRow.jv_br_name = "CARGOMAR PVT LTD TUTICORIN";
            }
            if (brcode == "COKPR")
            {
                mRow.jv_br_id = "B3860948-7208-6FB4-0E50-E267628A52BC";
                mRow.jv_br_code = "CPLPRJ";
                mRow.jv_br_name = "CARGOMAR PVT LTD PROJECTS";
            }


        }

        public string AllValid(Posting Record)
        {
            string str = "";
            try
            {
                str = "";
                if (Record.jv_br_record_pkid.ToString() == "")
                {
                    string jvhdate = Lib.StringToDate(Record.jv_date.ToString());
                    str += Lib.IsDateLocked(jvhdate, "HO",
                            Record._globalvariables.comp_code,
                            Record._globalvariables.branch_code, Record._globalvariables.year_code);
                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }


        public Dictionary<string, object> Save(Posting pRec)
        {
            Boolean CanSaveBranch = true;

            Boolean bTrans = false;
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();
                if ((ErrorMessage = AllValid(pRec)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }
                /*
                if (pRec._globalvariables.branch_code != "KOLAF")
                {
                    CanSaveBranch = false;    
                }
                */
                // Save Branch Invoice - only for debit note
                if (CanSaveBranch && Lib.Conv2Integer(pRec._globalvariables.year_code) >= 2019)
                {

                    if (pRec.jv_drcr == "DR")
                    {

                        pRec.rec_mode = "EDIT";
                        if (pRec.jv_br_inv_record_pkid.ToString() == "")
                        {
                            pRec.jv_br_inv_record_pkid = System.Guid.NewGuid().ToString().ToUpper();
                            pRec.rec_mode = "ADD";
                        }
                        if (pRec.rec_mode == "ADD")
                        {
                            sql = "select cost_pkid from costingm where cost_pkid ='" + pRec.jv_cost_id + "' and cost_jv_br_inv_id is not null";
                            if (Con_Oracle.IsRowExists(sql))
                            {
                                if (Con_Oracle != null)
                                    Con_Oracle.CloseConnection();
                                ErrorMessage = "Already Posted, Try Again";
                                throw new Exception(ErrorMessage);
                            }
                        }

                        if (SaveBranchInvoice(pRec))
                        {
                            sql = "update costingm set cost_jv_br_inv_id = '" + pRec.jv_br_inv_record_pkid + "' ";
                            if (pRec.rec_mode == "ADD")
                                sql += ", cost_jv_br_invno = " + BR_INVNO.ToString();
                            sql += " where cost_pkid = '" + pRec.jv_cost_id + "'";
                            Con_Oracle.BeginTransaction();
                            bTrans = true;
                            Con_Oracle.ExecuteNonQuery(sql);
                            Con_Oracle.CommitTransaction();
                            Con_Oracle.CloseConnection();
                            bTrans = false;
                        }
                    }
                    else
                    {
                        if (pRec.jv_br_inv_record_pkid != "")
                        {

                            Con_Oracle.BeginTransaction();
                            bTrans = true;

                            sql = "delete from  costcentert where ct_jvh_id ='" + pRec.jv_br_inv_record_pkid + "'";
                            Con_Oracle.ExecuteNonQuery(sql);
                            sql = "delete from  ledgerxref where xref_jvh_id ='" + pRec.jv_br_inv_record_pkid + "'";
                            Con_Oracle.ExecuteNonQuery(sql);
                            sql = "delete from  ledgert where jv_parent_id ='" + pRec.jv_br_inv_record_pkid + "'";
                            Con_Oracle.ExecuteNonQuery(sql);
                            sql = "delete from  ledgerh where jvh_pkid ='" + pRec.jv_br_inv_record_pkid + "'";
                            Con_Oracle.ExecuteNonQuery(sql);

                            sql = "update costingm set cost_jv_br_inv_id = '', cost_jv_br_invno = 0 ";
                            sql += " where cost_pkid = '" + pRec.jv_cost_id + "'";
                            Con_Oracle.ExecuteNonQuery(sql);

                            Con_Oracle.CommitTransaction();
                            Con_Oracle.CloseConnection();
                            bTrans = false;

                        }
                    }


                }



                //Save  Branch JV
                if (CanSaveBranch)
                {

                    pRec.rec_mode = "EDIT";
                    if (pRec.jv_br_record_pkid.ToString() == "")
                    {
                        pRec.jv_br_record_pkid = System.Guid.NewGuid().ToString().ToUpper();
                        pRec.rec_mode = "ADD";
                    }
                    if (pRec.rec_mode == "ADD")
                    {
                        sql = "select cost_pkid from costingm where cost_pkid ='" + pRec.jv_cost_id + "' and cost_jv_br_id is not null";
                        if (Con_Oracle.IsRowExists(sql))
                        {
                            if (Con_Oracle != null)
                                Con_Oracle.CloseConnection();
                            ErrorMessage = "Already Posted, Try Again";
                            throw new Exception(ErrorMessage);
                        }
                    }
                    if (SaveBranch(pRec))
                    {
                        sql = "update costingm set cost_jv_br_id = '" + pRec.jv_br_record_pkid + "' ";
                        if (pRec.rec_mode == "ADD")
                            sql += ", cost_jv_br_vrno = " + BR_VRNO.ToString();
                        sql += " where cost_pkid = '" + pRec.jv_cost_id + "'";
                        Con_Oracle.BeginTransaction();
                        bTrans = true;
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                        Con_Oracle.CloseConnection();
                        bTrans = false;
                    }
                }


                //if (pRec._globalvariables.branch_code != "KOLAF"){

                    pRec.rec_mode = "EDIT";
                    if (pRec.jv_ho_record_pkid.ToString() == "")
                    {
                        pRec.jv_ho_record_pkid = System.Guid.NewGuid().ToString().ToUpper();
                        pRec.rec_mode = "ADD";
                    }
                    if (pRec.rec_mode == "ADD")
                    {
                        sql = "select cost_pkid from costingm where cost_pkid ='" + pRec.jv_cost_id + "' and cost_jv_ho_id is not null";
                        if (Con_Oracle.IsRowExists(sql))
                        {
                            if (Con_Oracle != null)
                                Con_Oracle.CloseConnection();
                            ErrorMessage = "Already Posted, Try Again";
                            throw new Exception(ErrorMessage);
                        }
                    }

                    if (SaveHO(pRec))
                    {
                        sql = " update costingm set cost_jv_ho_id = '" + pRec.jv_ho_record_pkid + "' ";
                        if (pRec.rec_mode == "ADD")
                            sql += " ,cost_jv_ho_vrno = " + HO_VRNO.ToString();
                        sql += " where cost_pkid = '" + pRec.jv_cost_id + "'";
                        Con_Oracle.BeginTransaction();
                        bTrans = true;
                        Con_Oracle.ExecuteNonQuery(sql);
                        Con_Oracle.CommitTransaction();
                        Con_Oracle.CloseConnection();
                        bTrans = false;
                    }

                    if (pRec._globalvariables.branch_code == "KOLAF")
                        sql = "update costingm set cost_jv_posted = 'Y' where cost_pkid = '" + pRec.jv_cost_id + "' and cost_jv_ho_id is not null";
                    else
                        sql = "update costingm set cost_jv_posted = 'Y' where cost_pkid = '" + pRec.jv_cost_id + "' and cost_jv_br_id is not null and  cost_jv_ho_id is not null";

                    Con_Oracle.BeginTransaction();
                    bTrans = true;
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    Con_Oracle.CloseConnection();
                    bTrans = false;
                //}

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    if (bTrans)
                        Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            return RetData;
        }

        private Boolean SaveBranchInvoice(Posting pRec)
        {
            Boolean bRet = false;
            LedgerService LedService = new LedgerService();


            DataTable Dt_cc_hbls = new DataTable();


            decimal ftotal = Math.Abs(pRec.jv_ftotal);
            decimal total = Math.Abs(pRec.jv_total);

            try
            {

                // Save Branch JV
                Ledgerh Record = null;
                Record = new Ledgerh();
                Record.jvh_pkid = pRec.jv_br_inv_record_pkid;
                Record.jvh_type = "IN-ES";
                Record.jvh_subtype = "AR";

                Record._globalvariables = new GlobalVariables();
                Record._globalvariables.user_code = pRec._globalvariables.user_code;
                Record._globalvariables.comp_code = pRec._globalvariables.comp_code;
                Record._globalvariables.branch_code = pRec._globalvariables.branch_code;
                Record._globalvariables.year_code = pRec._globalvariables.year_code;
                Record._globalvariables.year_prefix = pRec._globalvariables.year_prefix;
                Record._globalvariables.year_start_date = pRec._globalvariables.year_start_date;
                Record._globalvariables.year_end_date = pRec._globalvariables.year_end_date;


                Record.jvh_year = Lib.Conv2Integer(pRec._globalvariables.year_code);
                Record.jvh_date = pRec.jv_date;
                Record.jvh_reference = pRec.jv_reference;
                Record.jvh_reference_date = pRec.jv_date;

                Record.jvh_narration = pRec.jv_narration;

                Record.jvh_remarks = pRec.jv_remarks;

                Record.jvh_allocation_found = false;


                Record.jvh_rec_source = "HC";

                Record.jvh_acc_id = pRec.jv_agent_id;
                Record.jvh_acc_code = pRec.jv_agent_code;
                Record.jvh_acc_name = pRec.jv_agent_name;
                Record.jvh_acc_br_id = pRec.jv_agent_br_id;
                Record.jvh_sez = false;

                Record.jvh_state_id = "6CDE9A9A-7FC9-473C-A827-209A49BD6DCC";
                Record.jvh_state_code = "";
                Record.jvh_state_name = "";

                Record.jvh_gst = true;
                Record.jvh_rc = false;
                Record.jvh_exwork = false;

                Record.jvh_gstin = "";
                Record.jvh_gst_type = "INTER-STATE";

                Record.jvh_curr_id = pRec.jv_curr_id;
                Record.jvh_curr_code = pRec.jv_curr_code;
                Record.jvh_curr_name = "";

                Record.jvh_exrate = pRec.jv_exrate;

                Record.rec_category = pRec.rec_category;

                Record.jvh_cc_category = "NA";
                Record.jvh_cc_id = "";
                Record.jvh_cc_code = "";
                Record.jvh_cc_name = "";
                if (pRec.mbl_type == "MBL-SE")
                {
                    Record.jvh_cc_category = "MBL SEA EXPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCCntrs(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "MBL-SI")
                {
                    Record.jvh_cc_category = "MBL SEA IMPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "MBL-AE")
                {
                    Record.jvh_cc_category = "MAWB AIR EXPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "MBL-AI")
                {
                    Record.jvh_cc_category = "MAWB AIR IMPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "GEN-JOB")
                {
                    Record.jvh_cc_category = "GENERAL JOB";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }


                Record.jvh_org_invno = "";
                Record.jvh_org_invdt = "";
                Record.jvh_cgst_amt = 0;
                Record.jvh_sgst_amt = 0;
                Record.jvh_igst_amt = 0;
                Record.jvh_gst_amt = 0;



                Record.jvh_tot_famt = ftotal;
                Record.jvh_net_famt = ftotal;

                Record.jvh_tot_amt = total;
                Record.jvh_net_amt = total;
                Record.jvh_debit = total;
                Record.jvh_credit = total;
                Record.jvh_diff = 0;

                Record.jvh_location = pRec.jv_cost_prefix;

                Record.rec_mode = pRec.rec_mode;
                Record.rec_category = pRec.rec_category;


                Record.CostCenterList = new List<CostCentert>();

                Record.XrefList = new List<LedgerXref>();

                set_frt_id = false;
                if (pRec.jv_drcr == "DR")
                {
                    Record.LedgerList = new List<Ledgert>();
                    Record.LedgerList.Add(AddRecord(pRec.jv_agent_id, pRec.jv_agent_code, pRec.jv_agent_name, pRec.jv_curr_id, pRec.jv_curr_code, "DR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category, false, false, "HEADER"));
                    set_frt_id = true;
                    Record.LedgerList.Add(AddRecord(pRec.jv_frt_id, pRec.jv_frt_code, pRec.jv_frt_name, pRec.jv_curr_id, pRec.jv_curr_code, "CR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category, true, true, "", pRec.jv_sac_id));
                    Dt_cc_hbls = new DataTable();
                }

                decimal cc_amt = 0;
                CostCentert mcc;
                int iCtr = 0;
                foreach (DataRow Dr in Dt_cc_hbls.Rows)
                {
                    iCtr++;
                    cc_amt = pRec.jv_total / Lib.Conv2Integer(Dr["tot"].ToString());
                    cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                    if (pRec.mbl_type != "MBL-SE")
                        cc_amt = Lib.getCCAmt(Dr, pRec.jv_total, cc_amt);

                    if (cc_amt < 0)
                        cc_amt = Math.Abs(cc_amt);
                    mcc = new CostCentert();
                    mcc.ct_ctr = iCtr;
                    mcc.ct_pkid = System.Guid.NewGuid().ToString().ToUpper();
                    mcc.ct_jv_id = frt_jv_id;
                    mcc.ct_acc_id = pRec.jv_frt_id;
                    mcc.ct_category = Dr["cc_category"].ToString();
                    mcc.ct_cost_id = Dr["id"].ToString();
                    mcc.ct_cost_year = Lib.Conv2Integer(pRec._globalvariables.year_code);
                    mcc.ct_year = Lib.Conv2Integer(pRec._globalvariables.year_code);
                    mcc.ct_amount = cc_amt;
                    Record.CostCenterList.Add(mcc);
                }

                Dictionary<string, object> mobj = LedService.Save(Record);
                if (mobj.ContainsKey("jvh_vrno"))
                    BR_INVNO = mobj["jvh_vrno"].ToString();

                bRet = true;

                //this.CCList = new Array<CostCentert>();
                //this.PendingListRecords = new Array<pendinglist>();

            }
            catch (Exception Ex)
            {
                bRet = false;
                BrErrorMessage = Ex.Message.ToString();
                throw Ex;
            }
            return bRet;
        }

        private Boolean SaveBranch(Posting pRec)
        {
            Boolean bRet = false;
            LedgerService LedService = new LedgerService();

            DataTable Dt_cc_hbls = new DataTable();


            decimal ftotal = Math.Abs(pRec.jv_ftotal);
            decimal total = Math.Abs(pRec.jv_total);

            try
            {

                // Save Branch JV
                Ledgerh Record = null;
                Record = new Ledgerh();
                Record.jvh_pkid = pRec.jv_br_record_pkid;
                Record.jvh_type = "HO";

                Record._globalvariables = new GlobalVariables();
                Record._globalvariables.user_code = pRec._globalvariables.user_code;
                Record._globalvariables.comp_code = pRec._globalvariables.comp_code;
                Record._globalvariables.branch_code = pRec._globalvariables.branch_code;
                Record._globalvariables.year_code = pRec._globalvariables.year_code;
                Record._globalvariables.year_prefix = pRec._globalvariables.year_prefix;
                Record._globalvariables.year_start_date = pRec._globalvariables.year_start_date;
                Record._globalvariables.year_end_date = pRec._globalvariables.year_end_date;


                Record.jvh_year = Lib.Conv2Integer(pRec._globalvariables.year_code);
                Record.jvh_date = pRec.jv_date;
                Record.jvh_reference = pRec.jv_reference;
                Record.jvh_reference_date = pRec.jv_date;

                Record.jvh_narration = pRec.jv_narration;

                Record.jvh_remarks = pRec.jv_remarks;

                Record.jvh_allocation_found = false;


                Record.jvh_rec_source = "HC";

                Record.jvh_acc_id = "";
                Record.jvh_acc_code = "";
                Record.jvh_acc_name = "";
                Record.jvh_acc_br_id = "";
                Record.jvh_sez = false;

                Record.jvh_state_id = "";
                Record.jvh_state_code = "";
                Record.jvh_state_name = "";

                Record.jvh_gstin = "";
                Record.jvh_gst_type = "";

                Record.jvh_curr_id = pRec.jv_curr_id;
                Record.jvh_curr_code = pRec.jv_curr_code;
                Record.jvh_curr_name = "";

                Record.jvh_exrate = pRec.jv_exrate;

                Record.rec_category = pRec.rec_category;

                Record.jvh_cc_category = "NA";
                Record.jvh_cc_id = "";
                Record.jvh_cc_code = "";
                Record.jvh_cc_name = "";
                if (pRec.mbl_type == "MBL-SE")
                {
                    Record.jvh_cc_category = "MBL SEA EXPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCCntrs(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "MBL-SI")
                {
                    Record.jvh_cc_category = "MBL SEA IMPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "MBL-AE")
                {
                    Record.jvh_cc_category = "MAWB AIR EXPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "MBL-AI")
                {
                    Record.jvh_cc_category = "MAWB AIR IMPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }
                if (pRec.mbl_type == "GEN-JOB")
                {
                    Record.jvh_cc_category = "GENERAL JOB";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                    Dt_cc_hbls = Lib.getCCHBLS(Record.jvh_cc_category, Record.jvh_cc_id);
                }


                Record.jvh_org_invno = "";
                Record.jvh_org_invdt = "";
                Record.jvh_cgst_amt = 0;
                Record.jvh_sgst_amt = 0;
                Record.jvh_igst_amt = 0;
                Record.jvh_gst_amt = 0;



                Record.jvh_tot_famt = ftotal;
                Record.jvh_net_famt = ftotal;

                Record.jvh_tot_amt = total;
                Record.jvh_net_amt = total;
                Record.jvh_debit = total;
                Record.jvh_credit = total;
                Record.jvh_diff = 0;

                Record.jvh_location = pRec.jv_cost_prefix;

                Record.rec_mode = pRec.rec_mode;
                Record.rec_category = pRec.rec_category;


                Record.CostCenterList = new List<CostCentert>();

                Record.XrefList = new List<LedgerXref>();

                set_frt_id = false;
                if (pRec.jv_drcr == "DR")
                {
                    Record.LedgerList = new List<Ledgert>();
                    if (Lib.Conv2Integer(pRec._globalvariables.year_code) >= 2019)
                    {
                        Record.LedgerList.Add(AddRecord(pRec.jv_ho_id, pRec.jv_ho_code, pRec.jv_ho_name, pRec.jv_curr_id, pRec.jv_curr_code, "DR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                        set_frt_id = false;
                        Record.LedgerList.Add(AddRecord(pRec.jv_agent_id, pRec.jv_agent_code, pRec.jv_agent_name, pRec.jv_curr_id, pRec.jv_curr_code, "CR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                        Dt_cc_hbls = new DataTable();
                    }
                    else
                    {
                        Record.LedgerList.Add(AddRecord(pRec.jv_ho_id, pRec.jv_ho_code, pRec.jv_ho_name, pRec.jv_curr_id, pRec.jv_curr_code, "DR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                        set_frt_id = true;
                        Record.LedgerList.Add(AddRecord(pRec.jv_frt_id, pRec.jv_frt_code, pRec.jv_frt_name, pRec.jv_curr_id, pRec.jv_curr_code, "CR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                    }
                }
                if (pRec.jv_drcr == "CR")
                {
                    Record.LedgerList = new List<Ledgert>();
                    Record.LedgerList.Add(AddRecord(pRec.jv_ho_id, pRec.jv_ho_code, pRec.jv_ho_name, pRec.jv_curr_id, pRec.jv_curr_code, "CR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                    set_frt_id = true;
                    Record.LedgerList.Add(AddRecord(pRec.jv_frt_id, pRec.jv_frt_code, pRec.jv_frt_name, pRec.jv_curr_id, pRec.jv_curr_code, "DR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                }
                decimal cc_amt = 0;
                CostCentert mcc;
                int iCtr = 0;
                foreach (DataRow Dr in Dt_cc_hbls.Rows)
                {
                    iCtr++;
                    cc_amt = pRec.jv_total / Lib.Conv2Integer(Dr["tot"].ToString());
                    cc_amt = Lib.RoundNumber_Latest(cc_amt.ToString(), 2, true);
                    if (pRec.mbl_type != "MBL-SE")
                        cc_amt = Lib.getCCAmt(Dr, pRec.jv_total, cc_amt);

                    if (cc_amt < 0)
                        cc_amt = Math.Abs(cc_amt);
                    mcc = new CostCentert();
                    mcc.ct_ctr = iCtr;
                    mcc.ct_pkid = System.Guid.NewGuid().ToString().ToUpper();
                    mcc.ct_jv_id = frt_jv_id;
                    mcc.ct_acc_id = pRec.jv_frt_id;
                    mcc.ct_category = Dr["cc_category"].ToString();
                    mcc.ct_cost_id = Dr["id"].ToString();
                    mcc.ct_cost_year = Lib.Conv2Integer(pRec._globalvariables.year_code);
                    mcc.ct_year = Lib.Conv2Integer(pRec._globalvariables.year_code);
                    mcc.ct_amount = cc_amt;
                    Record.CostCenterList.Add(mcc);
                }

                Dictionary<string, object> mobj = LedService.Save(Record);
                if (mobj.ContainsKey("jvh_vrno"))
                    BR_VRNO = mobj["jvh_vrno"].ToString();

                bRet = true;

                //this.CCList = new Array<CostCentert>();
                //this.PendingListRecords = new Array<pendinglist>();

            }
            catch (Exception Ex)
            {
                bRet = false;
                BrErrorMessage = Ex.Message.ToString();
                throw Ex;
            }
            return bRet;
        }

        private Boolean SaveHO(Posting pRec)
        {
            Boolean bRet = false;
            LedgerService LedService = new LedgerService();


            decimal ftotal = Math.Abs(pRec.jv_ftotal);
            decimal total = Math.Abs(pRec.jv_total);



            try
            {
                // Save Branch JV
                Ledgerh Record = null;
                Record = new Ledgerh();
                Record.jvh_pkid = pRec.jv_ho_record_pkid;
                Record.jvh_type = "HO";

                Record._globalvariables = new GlobalVariables();
                Record._globalvariables.user_code = pRec._globalvariables.user_code;
                Record._globalvariables.comp_code = pRec._globalvariables.comp_code;
                Record._globalvariables.branch_code = "HOCPL";
                Record._globalvariables.year_code = pRec._globalvariables.year_code;

                Record._globalvariables.year_prefix = pRec._globalvariables.year_prefix;
                Record._globalvariables.year_start_date = pRec._globalvariables.year_start_date;
                Record._globalvariables.year_end_date = pRec._globalvariables.year_end_date;

                Record.jvh_year = Lib.Conv2Integer(pRec._globalvariables.year_code);

                Record.jvh_date = pRec.jv_date;
                Record.jvh_reference = pRec.jv_reference;
                Record.jvh_reference_date = pRec.jv_date;
                Record.jvh_narration = pRec.jv_narration;

                Record.jvh_remarks = pRec.jv_remarks;

                Record.jvh_allocation_found = false;

                Record.jvh_rec_source = "HC";

                Record.jvh_acc_id = "";
                Record.jvh_acc_code = "";
                Record.jvh_acc_name = "";
                Record.jvh_acc_br_id = "";
                Record.jvh_sez = false;

                Record.jvh_state_id = "";
                Record.jvh_state_code = "";
                Record.jvh_state_name = "";

                Record.jvh_gstin = "";
                Record.jvh_gst_type = "";

                Record.jvh_curr_id = pRec.jv_curr_id;
                Record.jvh_curr_code = pRec.jv_curr_code;
                Record.jvh_curr_name = "";

                Record.jvh_exrate = pRec.jv_exrate;

                Record.rec_category = pRec.rec_category;

                Record.jvh_cc_category = "NA";
                Record.jvh_cc_id = "";
                Record.jvh_cc_code = "";
                Record.jvh_cc_name = "";
                if (pRec.mbl_type == "MBL-SE")
                {
                    Record.jvh_cc_category = "MBL SEA EXPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                }
                if (pRec.mbl_type == "MBL-SI")
                {
                    Record.jvh_cc_category = "MBL SEA IMPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                }
                if (pRec.mbl_type == "MBL-AE")
                {
                    Record.jvh_cc_category = "MAWB AIR EXPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                }
                if (pRec.mbl_type == "MBL-AI")
                {
                    Record.jvh_cc_category = "MAWB AIR IMPORT";
                    Record.jvh_cc_id = pRec.mbl_pkid;
                }





                Record.jvh_org_invno = "";
                Record.jvh_org_invdt = "";
                Record.jvh_cgst_amt = 0;
                Record.jvh_sgst_amt = 0;
                Record.jvh_igst_amt = 0;
                Record.jvh_gst_amt = 0;


                Record.jvh_tot_famt = ftotal;
                Record.jvh_net_famt = ftotal;
                Record.jvh_tot_amt = total;
                Record.jvh_net_amt = total;
                Record.jvh_debit = total;
                Record.jvh_credit = total;
                Record.jvh_diff = 0;

                Record.jvh_location = pRec.jv_cost_prefix;
                Record.rec_mode = pRec.rec_mode;
                Record.rec_category = pRec.rec_category;

                Record.CostCenterList = new List<CostCentert>();
                Record.XrefList = new List<LedgerXref>();

                set_frt_id = false;

                if (pRec.jv_drcr == "DR")
                {
                    Record.LedgerList = new List<Ledgert>();
                    Record.LedgerList.Add(AddRecord(pRec.jv_agent_id, pRec.jv_agent_code, pRec.jv_agent_name, pRec.jv_curr_id, pRec.jv_curr_code, "DR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                    Record.LedgerList.Add(AddRecord(pRec.jv_br_id, pRec.jv_br_code, pRec.jv_br_name, pRec.jv_curr_id, pRec.jv_curr_code, "CR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                }
                if (pRec.jv_drcr == "CR")
                {
                    Record.LedgerList = new List<Ledgert>();
                    Record.LedgerList.Add(AddRecord(pRec.jv_agent_id, pRec.jv_agent_code, pRec.jv_agent_name, pRec.jv_curr_id, pRec.jv_curr_code, "CR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                    Record.LedgerList.Add(AddRecord(pRec.jv_br_id, pRec.jv_br_code, pRec.jv_br_name, pRec.jv_curr_id, pRec.jv_curr_code, "DR", pRec.jv_ftotal, pRec.jv_total, pRec.jv_exrate, pRec.rec_category));
                }

                Dictionary<string, object> mobj = LedService.Save(Record);
                if (mobj.ContainsKey("jvh_vrno"))
                    HO_VRNO = mobj["jvh_vrno"].ToString();

                bRet = true;
            }
            catch (Exception Ex)
            {
                bRet = false;
                BrErrorMessage = Ex.Message.ToString();
                throw Ex;
            }
            return bRet;

            //this.CCList = new Array<CostCentert>();
            //this.PendingListRecords = new Array<pendinglist>();
        }

        private Ledgert AddRecord(string accid, string accode, string accname, string curr_id, string curr_code, string drcr, decimal ftotal, decimal total, decimal exrate, string rec_category, Boolean Jv_Is_Taxable = false, Boolean jv_gst_item = false, string jv_row_type = "", string jv_sac_id = "")
        {
            ftotal = Math.Abs(ftotal);
            total = Math.Abs(total);

            string jv_id = System.Guid.NewGuid().ToString().ToUpper();

            GLOB_JVID = jv_id;

            // this is for setting costcenter when frt code is selected
            if (set_frt_id)
                frt_jv_id = jv_id;

            Ledgert Rec = new Ledgert();
            Rec.jv_pkid = jv_id;
            Rec.jv_acc_id = accid;
            Rec.jv_acc_code = accode;
            Rec.jv_acc_name = accname;
            Rec.jv_curr_id = curr_id;
            Rec.jv_curr_code = curr_code;
            Rec.jv_curr_name = "";
            Rec.jv_taxable_amt = 0;
            Rec.jv_is_taxable = Jv_Is_Taxable;
            Rec.jv_is_gst_item = jv_gst_item;
            Rec.jv_gst_rate = 0;
            Rec.jv_cgst_rate = 0;
            Rec.jv_sgst_rate = 0;
            Rec.jv_igst_rate = 0;
            Rec.jv_cgst_amt = 0;
            Rec.jv_sgst_amt = 0;
            Rec.jv_igst_amt = 0;
            Rec.jv_gst_amt = 0;
            Rec.jv_qty = 1;
            Rec.jv_rate = ftotal;
            Rec.jv_ftotal = ftotal;
            Rec.jv_total_fc = ftotal;
            Rec.jv_total = total;
            Rec.jv_net_total = total;
            Rec.jv_exrate = exrate;
            Rec.jv_taxable_amt = total;
            if (drcr == "DR")
            {
                Rec.jv_debit = total;
                Rec.jv_credit = 0;
                Rec.jv_drcr = "DR";
            }
            if (drcr == "CR")
            {
                Rec.jv_debit = 0;
                Rec.jv_credit = total;
                Rec.jv_drcr = "CR";
            }

            Rec.jv_doc_type = "NA";
            Rec.jv_bank = "";
            Rec.jv_branch = "";
            Rec.jv_due_date = "";
            Rec.jv_pay_reason = "";
            Rec.jv_supp_docs = "";
            Rec.jv_paid_to = "";
            Rec.jv_remarks = "";
            Rec.jv_sac_id = jv_sac_id;
            Rec.jv_sac_code = "";
            Rec.jv_gst_edited = false;
            Rec.jv_recon_by = "";
            Rec.jv_recon_date = "";
            Rec.jv_pan_id = "";
            Rec.jv_pan_code = "";
            Rec.jv_pan_name = "";
            Rec.jv_tds_rate = 0;
            Rec.jv_tds_gross_amt = 0;
            Rec.jv_tan_id = "";
            Rec.jv_tan_code = "";
            Rec.jv_tan_name = "";
            Rec.jv_gross_bill_amt = 0;
            Rec.jv_tan_party_id = "";
            Rec.jv_tan_party_code = "";
            Rec.jv_tan_party_name = "";
            Rec.rec_category = rec_category;
            Rec.jv_row_type = jv_row_type;
            return Rec;
        }

        // Save Rebate 
        public Dictionary<string, object> SaveRebate(Rebatem pRec)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            int iRows = 0;

            LedgerService LedService = new LedgerService();

            string rbt_id = "";
            string rbt_code = "";
            string rbt_name = "";

            Boolean bFirstRow = true;


            DataTable Dt_cc_jobs = new DataTable();
            DataTable Dt_cc_hbls = new DataTable();
            DataTable Dt_cc_cntr = new DataTable();


            string JVHID = System.Guid.NewGuid().ToString().ToUpper();
            try
            {
                sql = " select acc_pkid,acc_code, acc_name from acctm where rec_company_code = '" + pRec._globalvariables.comp_code + "' and acc_code in ('RBTPAY',";
                sql += " '1101100','1102100','1103100','1105100','1106100','1107100',";
                sql += " '1201100','1202100','1203100','1205100',";
                sql += " '1301100','1302100','1303100','1305100','1306100','1307100',";
                sql += " '1401100','1402100','1403100','1405100'";
                sql += ")";

                Con_Oracle = new DBConnection();
                DataTable dt_acc = new DataTable();
                dt_acc = Con_Oracle.ExecuteQuery(sql);

                if (pRec.rec_mode == "EDIT")
                {
                    sql = "select count(*) as tot from jobincome where inv_rebate_jvid = '" + pRec.jvhid + "'";
                    DataTable dt_count = new DataTable();
                    dt_count = Con_Oracle.ExecuteQuery(sql);
                    if (dt_count.Rows.Count > 0)
                        iRows = Lib.Conv2Integer(dt_count.Rows[0]["tot"].ToString());
                    if (iRows != pRec.RebateList.Count)
                    {
                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception("All Previously Saved Rows Are Not Selected");
                    }
                }

                // Save Branch JV
                Ledgerh Record = null;
                Record = new Ledgerh();

                if (pRec.rec_mode == "ADD")
                    Record.jvh_pkid = JVHID;
                else
                    Record.jvh_pkid = pRec.jvhid;

                JVHID = Record.jvh_pkid;

                Record.jvh_type = "JV";

                Record._globalvariables = new GlobalVariables();

                Record._globalvariables.user_code = pRec._globalvariables.user_code;

                Record._globalvariables.comp_code = pRec._globalvariables.comp_code;
                Record._globalvariables.branch_code = pRec._globalvariables.branch_code;
                Record._globalvariables.year_code = pRec._globalvariables.year_code;

                Record._globalvariables.year_prefix = pRec._globalvariables.year_prefix;
                Record._globalvariables.year_start_date = pRec._globalvariables.year_start_date;
                Record._globalvariables.year_end_date = pRec._globalvariables.year_end_date;


                Record.jvh_year = Lib.Conv2Integer(pRec._globalvariables.year_code);

                Record.jvh_date = pRec.jvh_date;
                Record.jvh_narration = pRec.jvh_narration;

                Record.jvh_remarks = "";

                Record.jvh_allocation_found = false;

                Record.jvh_rec_source = "RP";

                Record.jvh_acc_id = "";
                Record.jvh_acc_code = "";
                Record.jvh_acc_name = "";
                Record.jvh_acc_br_id = "";
                Record.jvh_sez = false;

                Record.jvh_state_id = "";
                Record.jvh_state_code = "";
                Record.jvh_state_name = "";

                Record.jvh_gstin = "";
                Record.jvh_gst_type = "";

                Record.jvh_curr_id = "";
                Record.jvh_curr_code = "";
                Record.jvh_curr_name = "";

                Record.jvh_exrate = 1;

                Record.rec_category = pRec.rec_category;

                Record.jvh_cc_category = "NA";
                Record.jvh_cc_id = "";


                Record.jvh_cc_code = "";
                Record.jvh_cc_name = "";



                Record.jvh_org_invno = "";
                Record.jvh_org_invdt = "";
                Record.jvh_cgst_amt = 0;
                Record.jvh_sgst_amt = 0;
                Record.jvh_igst_amt = 0;
                Record.jvh_gst_amt = 0;

                Record.jvh_tot_famt = pRec.jvh_amount;
                Record.jvh_net_famt = pRec.jvh_amount;

                Record.jvh_tot_amt = pRec.jvh_amount;
                Record.jvh_net_amt = pRec.jvh_amount;
                Record.jvh_debit = pRec.jvh_amount;
                Record.jvh_credit = pRec.jvh_amount;
                Record.jvh_diff = 0;

                Record.jvh_location = "";
                Record.rec_mode = pRec.rec_mode;

                Record.rec_category = "";

                Record.CostCenterList = new List<CostCentert>();

                Record.XrefList = new List<LedgerXref>();

                Record.LedgerList = new List<Ledgert>();

                foreach (DataRow Dr in dt_acc.Rows)
                {
                    if (Dr["acc_code"].ToString() == "RBTPAY")
                    {
                        rbt_id = Dr["acc_pkid"].ToString();
                        rbt_code = Dr["acc_code"].ToString();
                        rbt_name = Dr["acc_name"].ToString();
                    }
                }

                if (rbt_id == "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception("Rebate Payable Code Not Found");
                }

                Record.LedgerList.Add(AddRecord(rbt_id, rbt_code, rbt_name,
                "FE6015A3-001A-4A1F-AE13-BCFCF53F3A5A", "INR", "CR",
                Lib.Conv2Decimal(pRec.jvh_amount.ToString()),
                Lib.Conv2Decimal(pRec.jvh_amount.ToString()),
                1, pRec.rec_category));

                foreach (Rebate Rec in pRec.RebateList)
                {
                    if (bFirstRow)
                    {
                        Record.jvh_narration = "Rebate Payable to " + Rec.shipper_name;
                        Record.jvh_narration += " JOB# " + Rec.jobnos;
                        Record.jvh_narration += " SI# " + Rec.hbl_no;
                        Record.jvh_narration += " MBL# " + Rec.mbl;
                        bFirstRow = false;

                        Record.jvh_cc_id = Rec.hbl_pkid;
                        if (Rec.hbl_type == "HBL-SE")
                            Record.jvh_cc_category = "SI SEA EXPORT";
                        if (Rec.hbl_type == "HBL-AE")
                            Record.jvh_cc_category = "SI AIR EXPORT";
                        if (Rec.hbl_type == "HBL-SI")
                            Record.jvh_cc_category = "SI SEA IMPORT";
                        if (Rec.hbl_type == "HBL-AI")
                            Record.jvh_cc_category = "SI AIR IMPORT";
                        if (Rec.hbl_type == "JOB-GN")
                            Record.jvh_cc_category = "GENERAL JOB";
                    }

                    rbt_id = ""; rbt_code = ""; rbt_name = "";
                    foreach (DataRow Dr in dt_acc.Rows)
                    {
                        if (Dr["acc_code"].ToString() == Rec.acc_main_code + "100")
                        {
                            rbt_id = Dr["acc_pkid"].ToString();
                            rbt_code = Dr["acc_code"].ToString();
                            rbt_name = Dr["acc_name"].ToString();
                        }
                    }
                    if (rbt_id == "")
                        break;
                    Record.LedgerList.Add(AddRecord(rbt_id, rbt_code, rbt_name,
                       "FE6015A3-001A-4A1F-AE13-BCFCF53F3A5A", "INR", "DR",
                       Lib.Conv2Decimal(Rec.inv_rebate_amt_inr.ToString()),
                       Lib.Conv2Decimal(Rec.inv_rebate_amt_inr.ToString()),
                       1, pRec.rec_category));
                }

                if (rbt_id == "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception("Rebate Payable Code Not Found");
                }
                if (Record.jvh_cc_category == "" || Record.jvh_cc_id == "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception("Invalid Cost Center");
                }


                Dictionary<string, object> mobj = LedService.Save(Record);
                if (mobj.ContainsKey("jvh_vrno"))
                    BR_VRNO = mobj["jvh_vrno"].ToString();

                if (pRec.rec_mode == "EDIT")
                    BR_VRNO = pRec.jvh_vrno.ToString();


                Lib.UpdateCC(JVHID);


                if (pRec.rec_mode == "ADD")
                {
                    Con_Oracle.BeginTransaction();

                    foreach (Rebate Rec in pRec.RebateList)
                    {
                        sql = " update jobincome set inv_rebate_jvid = '" + JVHID.ToString() + "',";
                        sql += " inv_rebate_jvno =" + BR_VRNO.ToString();
                        sql += " where inv_pkid = '" + Rec.inv_pkid + "'";
                        Con_Oracle.ExecuteNonQuery(sql);
                    }

                    Con_Oracle.CommitTransaction();
                }

                Con_Oracle.CloseConnection();


                RetData.Add("jvid", JVHID);
                RetData.Add("jvno", BR_VRNO);
            }
            catch (Exception Ex)
            {

                BrErrorMessage = Ex.Message.ToString();
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            return RetData;
        }


        private DataTable Dt_SalHead = new DataTable();


        public Dictionary<string, object> SavePayRoll(Dictionary<string, object> SaveData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string sNarration = "";
            string JVHID = "";
            string JVHDT = "";
            string rec_mode = "";
            string JVHVRNO = "";

            string sal_year = SaveData["sal_year"].ToString();
            string sal_month = SaveData["sal_month"].ToString();

            string user_code = SaveData["user_code"].ToString();
            string comp_code = SaveData["company_code"].ToString();
            string branch_code = SaveData["branch_code"].ToString();
            string year_code = SaveData["year_code"].ToString();
            string year_prefix = SaveData["year_prefix"].ToString();
            string year_start_date = SaveData["year_start_date"].ToString();
            string year_end_date = SaveData["year_end_date"].ToString();
            string jvh_year = SaveData["year_code"].ToString();

            decimal nTotAmt = 0;

            LedgerService LedService = new LedgerService();

            DataTable Dt_cc_jobs = new DataTable();
            DataTable Dt_cc_hbls = new DataTable();
            DataTable Dt_cc_cntr = new DataTable();

            try
            {

                sql = "";

                int iYear = Lib.Conv2Integer(sal_year);
                int iMonth = Lib.Conv2Integer(sal_month) + 1;
                if (iMonth == 13)
                {
                    iYear += 1;
                    iMonth = 1;
                }
                System.DateTime myDt = new DateTime(iYear, iMonth, 01);
                myDt = myDt.AddDays(-1);
                if (myDt.DayOfWeek == DayOfWeek.Sunday)
                    myDt = myDt.AddDays(-1);

                if (myDt > System.DateTime.Today)
                    myDt = System.DateTime.Today;

                JVHDT = myDt.ToString("yyyy-MM-dd");
                string JVHDT1 = myDt.ToString("dd-MMM-yyyy");

                Con_Oracle = new DBConnection();

                //DataTable dt_acc = new DataTable();
                //dt_acc = Con_Oracle.ExecuteQuery(sql);


                string curr_id = "";
                string curr_code = "";
                DataTable Dt_Curr = new DataTable();
                sql = "select id, code from settings where tablename = 'PARAM' and caption = 'LOCAL-CURRENCY' and parentid = '" + comp_code + "'";
                Dt_Curr = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Curr.Rows.Count <= 0)
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception("Local Currency Not Defined");
                }
                else
                {
                    curr_id = Dt_Curr.Rows[0]["id"].ToString();
                    curr_code = Dt_Curr.Rows[0]["code"].ToString();
                }



                rec_mode = "";

                sql = "select salh_pkid, salh_jvid from salaryh where ";
                sql += " rec_company_code = '{COMP_CODE}' and rec_branch_code = '{BRANCH_CODE}'";
                sql += " and salh_year = {SAL_YEAR} and salh_month = {SAL_MONTH} ";

                sql = sql.Replace("{COMP_CODE}", comp_code);
                sql = sql.Replace("{BRANCH_CODE}", branch_code);
                sql = sql.Replace("{SAL_YEAR}", sal_year);
                sql = sql.Replace("{SAL_MONTH}", sal_month);

                DataTable dt_rec = new DataTable();
                dt_rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in dt_rec.Rows)
                {
                    JVHID = Dr["salh_pkid"].ToString();
                    if (Dr["salh_jvid"].ToString() == "")
                        rec_mode = "ADD";
                    else
                        rec_mode = "EDIT";
                }

                sql = "select jvh_pkid, jvh_vrno, to_char(jvh_date,'YYYY-MM-DD') as dt, to_char(jvh_date,'DD-MON-YYYY') as dt1   from ledgerh where jvh_pkid = '" + JVHID + "'";
                dt_rec = Con_Oracle.ExecuteQuery(sql);
                if (rec_mode == "EDIT")
                {
                    if (dt_rec.Rows.Count <= 0)
                    {
                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception("Cannot Find JV posted entries");
                    }
                    else
                    {
                        JVHDT = dt_rec.Rows[0]["dt"].ToString();
                        JVHDT1 = dt_rec.Rows[0]["dt1"].ToString();
                        JVHVRNO = dt_rec.Rows[0]["jvh_vrno"].ToString();
                    }
                }
                else
                {
                    if (dt_rec.Rows.Count > 0)
                    {
                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception("JV Already Exists");
                    }
                }

                DataTable DT = new DataTable();
                DT.Columns.Add("ID", typeof(System.String));

                sql = "";
                sql += " select sal_pkid, sal_emp_id, sal_date, sal_fin_year, sal_year, sal_month, ";
                sql += " emp_no, emp_name, param_code as dept,a.rec_category,sal_gross_earn ,  ";
                sql += " D01,D02,D03,D04,D05,D06, D07, D08,D09,D10,D11, D12, D13,D14,D15,D16,D17,D18,D19,D20,D21,D22,D23,D24,D25 ";
                sql += " from salarym a ";
                sql += " left join empm b on a.sal_emp_id = b.emp_pkid ";
                sql += " left join param p on emp_department_id = param_pkid ";
                sql += " where a.rec_company_code = '{COMP_CODE}' and a.rec_branch_code = '{BRANCH_CODE}' and sal_year = {SAL_YEAR} and sal_month = {SAL_MONTH} ";
                sql += " order by emp_no ";

                sql = sql.Replace("{COMP_CODE}", comp_code);
                sql = sql.Replace("{BRANCH_CODE}", branch_code);
                sql = sql.Replace("{SAL_YEAR}", sal_year);
                sql = sql.Replace("{SAL_MONTH}", sal_month);

                DataTable Dt_Sal = new DataTable();
                Dt_Sal = Con_Oracle.ExecuteQuery(sql);

                sql = "";
                sql += " select sal_code, sal_desc, sal_head, sal_acc_id, acc_pkid, acc_code, acc_name,acc_cost_centre ";
                sql += " from salaryheadm a ";
                sql += " left join acctm b on a.sal_acc_id = b.acc_pkid ";
                sql += " where a.rec_company_code = '{COMP_CODE}' ";
                sql += " order by sal_code ";
                sql = sql.Replace("{COMP_CODE}", comp_code);

                Dt_SalHead = new DataTable();
                Dt_SalHead = Con_Oracle.ExecuteQuery(sql);




                // Save Branch JV
                Ledgerh Record = null;
                Record = new Ledgerh();
                Record.jvh_pkid = JVHID;
                Record.jvh_type = "JV";

                Record._globalvariables = new GlobalVariables();

                Record._globalvariables.user_code = user_code;
                Record._globalvariables.comp_code = comp_code;
                Record._globalvariables.branch_code = branch_code;
                Record._globalvariables.year_code = year_code;
                Record._globalvariables.year_prefix = year_prefix;
                Record._globalvariables.year_start_date = year_start_date;
                Record._globalvariables.year_end_date = year_end_date;
                Record.jvh_year = Lib.Conv2Integer(year_code);
                Record.jvh_date = JVHDT;

                sNarration = "";

                Record.jvh_narration = "Payroll";
                Record.jvh_remarks = "";
                Record.jvh_allocation_found = false;

                Record.jvh_rec_source = "JV";

                Record.jvh_acc_id = "";
                Record.jvh_acc_code = "";
                Record.jvh_acc_name = "";
                Record.jvh_acc_br_id = "";
                Record.jvh_sez = false;

                Record.jvh_state_id = "";
                Record.jvh_state_code = "";
                Record.jvh_state_name = "";

                Record.jvh_gstin = "";
                Record.jvh_gst_type = "";

                Record.jvh_curr_id = curr_id;
                Record.jvh_curr_code = curr_code;
                Record.jvh_curr_name = "";

                Record.jvh_exrate = 1;

                Record.rec_category = "OTHERS";

                Record.jvh_cc_category = "NA";
                Record.jvh_cc_id = "";


                Record.jvh_cc_code = "";
                Record.jvh_cc_name = "";



                Record.jvh_org_invno = "";
                Record.jvh_org_invdt = "";
                Record.jvh_cgst_amt = 0;
                Record.jvh_sgst_amt = 0;
                Record.jvh_igst_amt = 0;
                Record.jvh_gst_amt = 0;


                Record.jvh_diff = 0;

                Record.jvh_location = "";
                Record.rec_mode = rec_mode;

                Record.rec_category = "";

                Record.CostCenterList = new List<CostCentert>();
                Record.XrefList = new List<LedgerXref>();
                Record.LedgerList = new List<Ledgert>();

                object Amt = 0;
                decimal nSAl1 = 0;
                decimal nAmt1 = 0;
                DataRow DrAcc = null;
                string sCol = "";

                decimal ntmpSal = 0;

                nTotAmt = 0;


                // MANAGEMENT SALARY
                Amt = Dt_Sal.Compute("sum(SAL_GROSS_EARN)", "DEPT='MANAGEMENT'");
                nSAl1 = Lib.Conv2Decimal(Amt.ToString());
                nTotAmt += nSAl1;
                if (nSAl1 > 0)
                {
                    sNarration = "DIRECTORS REMUNERATION ";
                    DrAcc = getID("A01");
                    Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                    curr_id, curr_code, "DR",
                    nSAl1, nSAl1, 1, "OTHERS"));
                    //Cost Center
                    AddCCList(Record, Dt_Sal, "MANAGEMENT", true, year_code, "SAL_GROSS_EARN", DrAcc["acc_cost_centre"].ToString(), DrAcc["acc_pkid"].ToString());

                    for (int i = 1; i <= 25; i++)
                    {
                        sCol = (i < 10) ? "D0" + i.ToString() : "D" + i.ToString();
                        if (sCol != "D06")
                        {
                            Amt = Dt_Sal.Compute("sum(" + sCol + ")", "DEPT='MANAGEMENT'");
                            nAmt1 = Lib.Conv2Decimal(Amt.ToString());
                            if (nAmt1 > 0)
                            {
                                DrAcc = getID(sCol);
                                Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                                curr_id, curr_code, "CR",
                                nAmt1, nAmt1, 1, "OTHERS"));
                                nSAl1 = nSAl1 - nAmt1;
                                // Cost Center
                                AddCCList(Record, Dt_Sal, "MANAGEMENT", true, year_code, sCol, DrAcc["acc_cost_centre"].ToString(), DrAcc["acc_pkid"].ToString());

                            }
                        }
                    }
                    // Loan
                    foreach (DataRow Dr in Dt_Sal.Rows)
                    {
                        if (Dr["DEPT"].ToString() == "MANAGEMENT")
                        {
                            DrAcc = getLoanAccount(Con_Oracle, comp_code, Dr["EMP_NO"].ToString());
                            nAmt1 = Lib.Conv2Decimal(Dr["D06"].ToString());
                            if (nAmt1 > 0)
                            {
                                Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                                curr_id, curr_code, "CR",
                                nAmt1, nAmt1, 1, "OTHERS"));
                                nSAl1 = nSAl1 - nAmt1;
                            }
                        }
                    }
                    if (nSAl1 > 0)
                    {
                        DrAcc = getID("A02");
                        Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                        curr_id, curr_code, "CR",
                        nSAl1, nSAl1, 1, "OTHERS"));
                    }
                }


                // SALARIES AND ALLOWANCE (CONFIRMED AND UNCONFIRMED STAFF)
                Amt = Dt_Sal.Compute("sum(SAL_GROSS_EARN)", "DEPT<>'MANAGEMENT' ");
                nSAl1 = Lib.Conv2Decimal(Amt.ToString());
                nTotAmt += nSAl1;
                if (nSAl1 > 0)
                {
                    //CONFIRMED STAFF
                    sNarration = (sNarration == "") ? " SALARY PAYABLE " : " & SALARY PAYABLE ";

                    Amt = Dt_Sal.Compute("sum(SAL_GROSS_EARN)", "DEPT<>'MANAGEMENT' AND REC_CATEGORY ='CONFIRMED' ");
                    ntmpSal = Lib.Conv2Decimal(Amt.ToString());
                    
                    if (ntmpSal > 0)
                    {
                        DrAcc = getID("A03"); // SALARIES ALLOWANCE
                        Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                        curr_id, curr_code, "DR",
                        ntmpSal, ntmpSal, 1, "OTHERS"));

                        // Cost Center
                        AddCCList(Record, Dt_Sal, "MANAGEMENT", false, year_code, "SAL_GROSS_EARN", DrAcc["acc_cost_centre"].ToString(), DrAcc["acc_pkid"].ToString(), "CONFIRMED");

                    }


                    Amt = Dt_Sal.Compute("sum(SAL_GROSS_EARN)", "DEPT<>'MANAGEMENT' AND REC_CATEGORY ='UNCONFIRM' ");
                    ntmpSal = Lib.Conv2Decimal(Amt.ToString());
                    if (ntmpSal > 0)
                    {
                        DrAcc = getID("A05"); // STIPEND CODE
                        Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                        curr_id, curr_code, "DR",
                        ntmpSal, ntmpSal, 1, "OTHERS"));

                        // Cost Center
                        AddCCList(Record, Dt_Sal, "MANAGEMENT", false, year_code, "SAL_GROSS_EARN", DrAcc["acc_cost_centre"].ToString(), DrAcc["acc_pkid"].ToString(), "UNCONFIRM");

                    }


                    for (int i = 1; i <= 25; i++)
                    {
                        sCol = (i < 10) ? "D0" + i.ToString() : "D" + i.ToString();
                        if (sCol != "D06")
                        {
                            Amt = Dt_Sal.Compute("sum(" + sCol + ")", "DEPT<>'MANAGEMENT' ");
                            nAmt1 = Lib.Conv2Decimal(Amt.ToString());
                            if (nAmt1 > 0)
                            {
                                DrAcc = getID(sCol);
                                Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                                curr_id, curr_code, "CR",
                                nAmt1, nAmt1, 1, "OTHERS"));
                                nSAl1 = nSAl1 - nAmt1;
                                // Cost Center
                                AddCCList(Record, Dt_Sal, "MANAGEMENT", false, year_code, sCol, DrAcc["acc_cost_centre"].ToString(), DrAcc["acc_pkid"].ToString());
                            }
                        }
                    }



                    // Staff Loan
                    foreach (DataRow Dr in Dt_Sal.Rows)
                    {
                        if (Dr["DEPT"].ToString() != "MANAGEMENT")
                        {
                            DrAcc = getLoanAccount(Con_Oracle, comp_code, Dr["EMP_NO"].ToString());
                            nAmt1 = Lib.Conv2Decimal(Dr["D06"].ToString());
                            if (nAmt1 > 0 && DrAcc != null)
                            {
                                Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                                curr_id, curr_code, "CR",
                                nAmt1, nAmt1, 1, "OTHERS"));
                                nSAl1 = nSAl1 - nAmt1;
                            }
                        }
                    }

                    if (nSAl1 > 0)
                    {
                        DrAcc = getID("A04");
                        Record.LedgerList.Add(AddRecord(DrAcc["acc_pkid"].ToString(), DrAcc["acc_code"].ToString(), DrAcc["acc_name"].ToString(),
                        curr_id, curr_code, "CR",
                        nSAl1, nSAl1, 1, "OTHERS"));
                    }
                }


                Record.jvh_tot_famt = nTotAmt;
                Record.jvh_net_famt = nTotAmt;
                Record.jvh_tot_amt = nTotAmt;
                Record.jvh_net_amt = nTotAmt;
                Record.jvh_debit = nTotAmt;
                Record.jvh_credit = nTotAmt;

                sNarration = "BEING PROVISIONS MADE FOR " + sNarration;
                sNarration += " FOR " + myDt.DayOfWeek + "/" + sal_year.ToString();
                sNarration += " AFTER DEDUCTING PF/TDS/OTHERS AS PER STATEMENT ";

                Record.jvh_narration = sNarration;

                Dictionary<string, object> mobj = LedService.Save(Record);
                if (mobj.ContainsKey("jvh_vrno"))
                    BR_VRNO = mobj["jvh_vrno"].ToString();

                if (rec_mode == "EDIT")
                    BR_VRNO = JVHVRNO.ToString();


                Con_Oracle.BeginTransaction();

                sql = " update salaryh set salh_jvid = '" + JVHID.ToString() + "',";
                sql += " salh_jvdate = '" + JVHDT1 + "',";
                sql += " salh_jvno =" + BR_VRNO.ToString() + ",";
                sql += " salh_posted ='Y' ";
                sql += " where salh_pkid = '" + JVHID + "'";
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                RetData.Add("jvno", BR_VRNO);
            }
            catch (Exception Ex)
            {

                BrErrorMessage = Ex.Message.ToString();
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            return RetData;
        }

        private DataRow getID(string Code)
        {
            DataRow Dr = null;
            foreach (DataRow Dr1 in Dt_SalHead.Rows)
            {
                if (Dr1["SAL_CODE"].ToString() == Code)
                {
                    Dr = Dr1;
                    break;
                }
            }

            if (Dr == null)
                throw new Exception("A/c Code Not Found : " + Code);

            if ( Dr["acc_pkid"].ToString().Trim().Length <=0 )
                throw new Exception("A/c Code Not Found : " + Code);

            return Dr;
        }


        private DataRow getLoanAccount(DBConnection Con_Oracle, string comp_code, string EmpCode)
        {
            sql = "select acc_pkid , acc_code, acc_name from acctm where rec_company_code = '" + comp_code + "'";
            sql += " and acc_code  = 'SL" + EmpCode + "'";
            DataTable Dt_acc = new DataTable();
            Dt_acc = Con_Oracle.ExecuteQuery(sql);
            if (Dt_acc.Rows.Count > 0)
                return Dt_acc.Rows[0];
            else
                return null;
        }

        private void AddCCList(Ledgerh Record, DataTable Dt_Sal, string Dept, Boolean Cond, string Year_Code, string sCol, string IsCC, string AccID, string Catg = "CONFIRMED")
        {

            if (IsCC == "N")
                return;
            decimal nAmt1 = 0;
            foreach (DataRow Dr in Dt_Sal.Rows)
            {
                if (  ((Dr["DEPT"].ToString() == Dept) == Cond) && Dr["REC_CATEGORY"].ToString() == Catg  ) 
                {
                    nAmt1 = Lib.Conv2Decimal(Dr[sCol].ToString());
                    if (nAmt1 > 0)
                    {
                        Record.CostCenterList.Add(addCC(GLOB_JVID, AccID, "EMPLOYEE", Dr["sal_emp_id"].ToString(), Year_Code, nAmt1));
                    }
                }
            }
        }


        private CostCentert addCC(string JVID, string ACCID, string CATEGORY, string COSTID, string fin_year, decimal cc_amt)
        {
            GLOB_CTR++;
            CostCentert mcc = new CostCentert();
            mcc.ct_ctr = GLOB_CTR;
            mcc.ct_pkid = System.Guid.NewGuid().ToString().ToUpper();
            mcc.ct_jv_id = JVID;
            mcc.ct_acc_id = ACCID;
            mcc.ct_category = CATEGORY;
            mcc.ct_cost_id = COSTID;
            mcc.ct_cost_year = Lib.Conv2Integer(fin_year);
            mcc.ct_year = Lib.Conv2Integer(fin_year);
            mcc.ct_amount = cc_amt;
            return mcc;
        }


    }
}
