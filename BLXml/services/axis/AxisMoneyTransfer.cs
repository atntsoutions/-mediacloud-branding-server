using System;
using System.Data;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Collections;
using DataBase;
using DataBase_Oracle.Connections;
using System.Text;

namespace BLXml
{
    public class AxisMoneyTransfer : BaseReport
    {
        private DataTable DT_MT = new DataTable();
        public Boolean IsError = false;
        public string ErrorMessage = "";
        private string sql = "";
        public string PKID = "";
        public string InvokeType = "";
        public string File_Name = "";
        public string File_Display_Name = "";
        public string report_folder = "";
        public int ProcessOrdCount = 0;
        public string File_Subject = "";
        public string Ftp_updtsql = "";
        public string branch_code = "";
        public string company_code = "";
        public string user_code = "";
        public string cust_uniq_ref = "";
        public string mt_lock = "";
        public string sRemarks = "";

        DBConnection Con_Oracle = null;
        StringBuilder sb = new StringBuilder();
        private StreamWriter sw;
        char SEP_CHAR;
        public void Generate()
        {
            try
            {
                ErrorMessage = "";
                IsError = false;
                UpdateTransferDate();
                ReadData();
                if (IsError)
                    return;
                SetCustomerUniqRefno();
                if (IsError)
                    return;
                WriteFiles();
                DT_MT.Rows.Clear();
            }
            catch (Exception ex)
            {
                IsError = true;
                ErrorMessage += " |" + ex.Message.ToString();
            }
        }
       
        private void UpdateTransferDate()
        {
            try
            {
                Con_Oracle = new DBConnection();
                sql = "update moneytransfer set mt_value_date ='" + DateTime.Now.ToString(Lib.BACK_END_DATE_FORMAT) + "', mt_transmission_date = sysdate where mt_jv_id ='" + PKID + "'";
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
                }
                throw Ex;
            }
        }

        private void ReadData()
        {

            Con_Oracle = new DBConnection();

            sql = "select mt_type ,";
            sql += " mt_txn_mode ,mt_corp_code,mt_cust_cfno,mt_cust_uniq_ref,";
            sql += " mt_corp_acc_no,mt_value_date ,mt_txn_curr ,mt_txn_amt,";
            sql += " mt_ben_name ,mt_ben_code ,mt_ben_acc_no ,";
            sql += " mt_ben_acc_type ,mt_ben_addr1,mt_ben_addr2,mt_ben_addr3,mt_ben_city ,";
            sql += " mt_ben_state,mt_ben_pin,mt_ben_ifsc ,mt_ben_bank_name,mt_base_code,";
            sql += " mt_chq_no ,mt_chq_date ,mt_payable_loc,mt_print_loc,mt_ben_email1,";
            sql += " mt_ben_email2,mt_ben_mob,mt_corp_batch_no,mt_company_code ,mt_product_code ,";
            sql += " mt_enrichment1,mt_enrichment2,mt_enrichment3,mt_enrichment4,mt_enrichment5,";
            sql += " mt_pay_type,mt_corp_email,mt_transmission_date,mt_user_id,mt_user_dept,a.rec_branch_code,a.rec_created_by ";
            sql += " from moneytransfer a";
            sql += " where mt_jv_id ='" + PKID + "'";

            DT_MT = new DataTable();
            DT_MT = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
          
            if (DT_MT.Rows.Count <= 0)
            {
                IsError = true;
                ErrorMessage += " | Details not Found";
            }
            
            foreach (DataRow dr in DT_MT.Rows)
            {
                if (dr["mt_type"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Identifier Cannot be Blank ";
                }
                if (dr["mt_txn_mode"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Payment Mode Cannot be Blank ";
                }
                if (dr["mt_corp_code"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Corporate Code Cannot be Blank ";
                }
                
                for (int i = 0; i < dr["mt_corp_code"].ToString().Length; i++)
                {
                    if (!Char.IsLetterOrDigit(dr["mt_corp_code"].ToString(), i) && !Char.IsWhiteSpace(dr["mt_corp_code"].ToString(), i))
                    {
                        IsError = true;
                        ErrorMessage += " | No Special characters allowed in Corporate Code ";
                        break;
                    }
                }
                if (dr["mt_corp_acc_no"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Debit Account Number Cannot be Blank ";
                }
                for (int i = 0; i < dr["mt_corp_acc_no"].ToString().Length; i++)
                {
                    if (!Char.IsDigit(dr["mt_corp_acc_no"].ToString(), i))
                    {
                        IsError = true;
                        ErrorMessage += " | Only numeric field allowed in Debit Account Number ";
                        break;
                    }
                }
                if (dr["mt_txn_curr"].ToString().Trim() != "INR")
                {
                    IsError = true;
                    ErrorMessage += " |  INR – Only Indian Rupees allowed ";
                }

                if (dr["mt_ben_name"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Beneficiary Name Cannot be Blank ";
                }
                if (dr["mt_ben_code"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Beneficiary Code Cannot be Blank ";
                }
                if (dr["mt_ben_acc_no"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Beneficiary Account Number Cannot be Blank ";
                }
                if (dr["mt_ben_acc_type"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Benefciary Account Type Cannot be Blank ";
                }
                if (dr["mt_ben_ifsc"].ToString().Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Beneficiary IFSC Code Cannot be Blank ";
                }

            }
            
        }
        private void SetCustomerUniqRefno()
        {
            try
            {
                string cfno = "";
                cust_uniq_ref = "";
                Con_Oracle = new DBConnection();

                sql = "select mt_jvh_docno from moneytransfer where mt_jv_id = '" + PKID + "'";
                Object sVal = Con_Oracle.ExecuteScalar(sql);
                if (sVal.ToString().Trim().Length > 0)
                {
                    string[] jvdocno = sVal.ToString().Split('/');

                    sql = "select nvl(max(mt_slno)+1,1001) as cfno from moneytransfer a ";
                    sql += " where a.rec_company_code = '{COMPCODE}'";
                    sql += " and a.rec_branch_code = '{BRCODE}'";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);

                    DataTable Dt_Temp = new DataTable();
                    Dt_Temp = Con_Oracle.ExecuteQuery(sql);
                    if (Dt_Temp.Rows.Count > 0)
                    {
                        cfno = Dt_Temp.Rows[0]["cfno"].ToString();
                    }
                    else
                    {
                        IsError = true;
                        ErrorMessage = "Ref Number Not Found Try again";

                        if (Con_Oracle != null)
                            Con_Oracle.CloseConnection();
                        throw new Exception(ErrorMessage);
                    }

                    for (int i = 0; i < jvdocno.Length - 1; i++)
                        cust_uniq_ref += jvdocno[i];
                    if (cust_uniq_ref != "")
                        cust_uniq_ref += cfno;
                }

                if (cust_uniq_ref.Length <= 0)
                {
                    IsError = true;
                    ErrorMessage += " | Customer Unique Reference Cannot be Blank ";
                }
                else
                {
                    sql = "update moneytransfer set mt_slno =" + cfno + ", mt_cust_uniq_ref ='" + cust_uniq_ref + "' where mt_jv_id ='" + PKID + "'";
                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                }
                Con_Oracle.CloseConnection();
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
        }
        private void WriteFiles()
        {
            try
            {
                Con_Oracle = new DBConnection();
                sql = "select mt_cust_uniq_ref from moneytransfer where mt_jv_id = '" + PKID + "'";
                Object sVal = Con_Oracle.ExecuteScalar(sql);
                cust_uniq_ref = sVal.ToString();
                Con_Oracle.CloseConnection();

                if (IsError || cust_uniq_ref == "")
                {
                    IsError = true;
                    ErrorMessage += " | File Not Generated.";
                    return;
                }

                File_Display_Name = cust_uniq_ref + ".TXT";
                File_Name = report_folder + File_Display_Name;

                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                sw = new StreamWriter(File_Name, false);
                SEP_CHAR = '^';

                DateTime transdate = DateTime.Now;
                sRemarks = "";
                foreach (DataRow dr in DT_MT.Rows)
                {
                    sRemarks += "Ref:" + cust_uniq_ref;
                    sRemarks += ",Mode:" + dr["mt_txn_mode"].ToString();
                    sRemarks += ",Benf:" + dr["mt_ben_name"].ToString();
                    sRemarks += ",Dr A/c:" + dr["mt_corp_acc_no"].ToString();
                    sRemarks += ",Cr A/c:" + dr["mt_ben_acc_no"].ToString();
                    sRemarks += ",Bank:" + dr["mt_ben_bank_name"].ToString();
                    sRemarks += ",IFSC:" + dr["mt_ben_ifsc"].ToString();
                    sRemarks += ",A/c type:" + dr["mt_ben_acc_type"].ToString();
                    sRemarks += ",Email:" + dr["mt_corp_email"].ToString();
                    

                    sw.Write(dr["mt_type"].ToString()); sw.Write(SEP_CHAR);//1
                    sw.Write(dr["mt_txn_mode"].ToString()); sw.Write(SEP_CHAR);//2
                    sw.Write(dr["mt_corp_code"].ToString()); sw.Write(SEP_CHAR);//3
                    sw.Write(cust_uniq_ref); sw.Write(SEP_CHAR);//4
                    sw.Write(dr["mt_corp_acc_no"].ToString()); sw.Write(SEP_CHAR);//5
                    if (!dr["mt_value_date"].Equals(DBNull.Value))
                    {
                        transdate = (DateTime)dr["mt_value_date"];
                        sw.Write(transdate.ToString("yyyy-MM-dd")); sw.Write(SEP_CHAR);//6
                    }
                    else
                    {
                        sw.Write(DateTime.Now.ToString("yyyy-MM-dd")); sw.Write(SEP_CHAR);//6
                    }
                    sw.Write(dr["mt_txn_curr"].ToString()); sw.Write(SEP_CHAR);//7
                    sw.Write(dr["mt_txn_amt"].ToString()); sw.Write(SEP_CHAR);//8
                    sw.Write(dr["mt_ben_name"].ToString()); sw.Write(SEP_CHAR);//9
                    sw.Write(dr["mt_ben_code"].ToString()); sw.Write(SEP_CHAR);//10
                    sw.Write(dr["mt_ben_acc_no"].ToString()); sw.Write(SEP_CHAR);//11
                    sw.Write(dr["mt_ben_acc_type"].ToString()); sw.Write(SEP_CHAR);//12
                    sw.Write(dr["mt_ben_addr1"].ToString()); sw.Write(SEP_CHAR);//13
                    sw.Write(dr["mt_ben_addr2"].ToString()); sw.Write(SEP_CHAR);//14
                    sw.Write(dr["mt_ben_addr3"].ToString()); sw.Write(SEP_CHAR);//15
                    sw.Write(dr["mt_ben_city"].ToString()); sw.Write(SEP_CHAR);//16
                    sw.Write(dr["mt_ben_state"].ToString()); sw.Write(SEP_CHAR);//17
                    sw.Write(dr["mt_ben_pin"].ToString()); sw.Write(SEP_CHAR);//18
                    sw.Write(dr["mt_ben_ifsc"].ToString()); sw.Write(SEP_CHAR);//19
                    sw.Write(dr["mt_ben_bank_name"].ToString()); sw.Write(SEP_CHAR);//20
                    sw.Write(dr["mt_base_code"].ToString()); sw.Write(SEP_CHAR);//21
                    sw.Write(dr["mt_chq_no"].ToString()); sw.Write(SEP_CHAR);//22
                    sw.Write(dr["mt_chq_date"].ToString()); sw.Write(SEP_CHAR);//23
                    sw.Write(dr["mt_payable_loc"].ToString()); sw.Write(SEP_CHAR);//24
                    sw.Write(dr["mt_print_loc"].ToString()); sw.Write(SEP_CHAR);//25
                    sw.Write(dr["mt_ben_email1"].ToString()); sw.Write(SEP_CHAR);//26
                    sw.Write(dr["mt_ben_email2"].ToString()); sw.Write(SEP_CHAR);//27
                    sw.Write(dr["mt_ben_mob"].ToString()); sw.Write(SEP_CHAR);//28
                    sw.Write(dr["mt_corp_batch_no"].ToString()); sw.Write(SEP_CHAR);//29
                    sw.Write(dr["mt_company_code"].ToString()); sw.Write(SEP_CHAR);//30
                    sw.Write(dr["mt_product_code"].ToString()); sw.Write(SEP_CHAR);//31
                    sw.Write(dr["mt_enrichment1"].ToString()); sw.Write(SEP_CHAR);//32
                    sw.Write(dr["mt_enrichment2"].ToString()); sw.Write(SEP_CHAR);//33
                    sw.Write(dr["mt_enrichment3"].ToString()); sw.Write(SEP_CHAR);//34
                    sw.Write(dr["mt_enrichment4"].ToString()); sw.Write(SEP_CHAR);//35
                    sw.Write(dr["mt_enrichment5"].ToString()); sw.Write(SEP_CHAR);//36
                    sw.Write(dr["mt_pay_type"].ToString()); sw.Write(SEP_CHAR);//37
                    sw.Write(dr["mt_corp_email"].ToString().ToLower()); sw.Write(SEP_CHAR);//38
                    if (!dr["mt_transmission_date"].Equals(DBNull.Value))
                    {
                        transdate = (DateTime)dr["mt_transmission_date"];
                        sw.Write(transdate.ToString("yyyy-MM-dd HH:mm:ss")); sw.Write(SEP_CHAR);//39
                    }
                    else
                    {
                        sw.Write(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")); sw.Write(SEP_CHAR);//39
                    }
                    sw.Write(dr["mt_user_id"].ToString()); sw.Write(SEP_CHAR);//40
                    sw.Write(dr["mt_user_dept"].ToString());//41
                    sw.WriteLine();
                    break;
                }

                sw.Flush();
                sw.Close();
                UpdateStatus();
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }

        private void UpdateStatus()
        {
            try
            {
                mt_lock = "G";
                Con_Oracle = new DBConnection();
                sql = "update moneytransfer set mt_lock ='G' where mt_jv_id ='" + PKID + "'";
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();

                Lib.AuditLog("FUND-TRANSFER", "GENERATE", "SENT", company_code, branch_code, user_code, PKID, cust_uniq_ref, sRemarks);
            }
            catch (Exception Ex)
            {
                mt_lock = "";
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
        }

        private void GenerateCsvFiles()
        {

            if (DT_MT.Rows.Count <= 0)
                return;

            sb = new StringBuilder();
            sb.Append("Identifier"); sb.Append(",");
            sb.Append("Payment Mode"); sb.Append(",");
            sb.Append("Corporate Code"); sb.Append(",");
            sb.Append("Customer Reference Number"); sb.Append(",");
            sb.Append("Debit Account Number"); sb.Append(",");
            sb.Append("Value Date"); sb.Append(",");
            sb.Append("Transaction Currency"); sb.Append(",");
            sb.Append("Transaction Amount"); sb.Append(",");
            sb.Append("Beneficiary Name"); sb.Append(",");
            sb.Append("Beneficiary Code / Vendor Code"); sb.Append(",");
            sb.Append("Beneficiary Account Number"); sb.Append(",");
            sb.Append("Benefciary Account Type"); sb.Append(",");
            sb.Append("Beneficiary Address 1"); sb.Append(",");
            sb.Append("Beneficiary Address 2"); sb.Append(",");
            sb.Append("Beneficiary Address 3"); sb.Append(",");
            sb.Append("Beneficiary City"); sb.Append(",");
            sb.Append("Beneficiary State"); sb.Append(",");
            sb.Append("Beneficiary Pin Code"); sb.Append(",");
            sb.Append("Beneficiary IFSC Code"); sb.Append(",");
            sb.Append("Beneficiary Bank Name"); sb.Append(",");
            sb.Append("Base Code"); sb.Append(",");
            sb.Append("Cheque Number"); sb.Append(",");
            sb.Append("Cheque Date"); sb.Append(",");
            sb.Append("Payable location"); sb.Append(",");
            sb.Append("Print Location"); sb.Append(",");
            sb.Append("Beneficiary Email address 1"); sb.Append(",");
            sb.Append("Beneficiary Email address 2"); sb.Append(",");
            sb.Append("Beneficiary Mobile Number"); sb.Append(",");
            sb.Append("Corp Batch No"); sb.Append(",");
            sb.Append("Company Code"); sb.Append(",");
            sb.Append("Product Code"); sb.Append(",");
            sb.Append("Extra 1"); sb.Append(",");
            sb.Append("Extra 2"); sb.Append(",");
            sb.Append("Extra 3"); sb.Append(",");
            sb.Append("Extra 4"); sb.Append(",");
            sb.Append("Extra 5"); sb.Append(",");
            sb.Append("PayType"); sb.Append(",");
            sb.Append("CORP_EMAIL_ADDR"); sb.Append(",");
            sb.Append("TRANSMISSION DATE"); sb.Append(",");
            sb.Append("User ID"); sb.Append(",");
            sb.Append("USER DEPARTMENT");
            foreach (DataRow Dr in DT_MT.Rows)
            {
                //sb.AppendLine();
                //sb.Append(Dr["poid"].ToString()); sb.Append(",");
                //sb.Append(Lib.GetTruncated(Dr["po"].ToString(), 20)); sb.Append(",");
                //sb.Append(Lib.GetTruncated(Dr["posupplier"].ToString(), 20)); sb.Append(",");
                //sb.Append(Lib.GetTruncated(Dr["oldpo"].ToString(), 20)); sb.Append(",");
                //sb.Append(Lib.GetTruncated(Dr["styleno"].ToString(), 20)); sb.Append(",");
                //sb.Append(Lib.NumericFormat(Dr["gw"].ToString(), 2)); sb.Append(",");
                //sb.Append(Lib.NumericFormat(Dr["packageno"].ToString(), 0)); sb.Append(",");
                //sb.Append(Lib.NumericFormat(Dr["cbm"].ToString(), 3)); sb.Append(",");
                //sb.Append(Lib.NumericFormat(Dr["piecenumber"].ToString(), 0)); sb.Append(",");
                //sb.Append(Lib.GetTruncated(Dr["Description"].ToString(), 100)); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["BookingRequestDate"])); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["RandomInspectionDate"])); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["poreleasedate"])); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["cargoreadydate"])); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["fcrdate"])); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["inspectiondate"])); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["stuffingdate"])); sb.Append(",");
                //sb.Append(Lib.DatetoStringDisplayformat(Dr["warehousedeparturedate"]));
            }
        }

        private void WriteCsvFiles()
        {
            try
            {
                if (sb == null || IsError)
                {
                    IsError = true;
                    ErrorMessage += " | Money Transfer Details Not Found.";
                    return;
                }

                if (File.Exists(File_Name))
                    File.Delete(File_Name);

                System.IO.File.AppendAllText(File_Name, sb.ToString());
            }
            catch (Exception Ex)
            {
                IsError = true;
                ErrorMessage += " |" + Ex.Message.ToString();
            }
        }
    }
}
