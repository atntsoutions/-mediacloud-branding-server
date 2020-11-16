using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
 
using DataBase;
using DataBase_Oracle.Connections;

namespace BLHr
{
    public class PayrollsettingService : BL_Base
    {
        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            payrollsetting mRow = new payrollsetting();
            mRow.ps_pkid = "";
            mRow.ps_admin_per = 0;
            mRow.ps_admin_amt = 0;
            mRow.ps_admin_based_on = "";
            mRow.ps_edli_per = 0;
            mRow.ps_edli_amt = 0;
            mRow.ps_edli_based_on = "";
            mRow.ps_esi_emplr_per = 0;
            mRow.ps_esi_limit = 0;
            mRow.ps_pf_emplr_pension_per = 0;
            mRow.ps_pf_cel_limit = 0;
            mRow.ps_pf_cel_limit_amt = 0;
            mRow.ps_esi_emply_per = 0;
            mRow.ps_pf_per = 0;
            mRow.ps_pf_col_excluded = "";
            mRow.ps_pf_br_region = "SOUTH";
            string comp_code = SearchData["comp_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = " select ps_pkid,ps_admin_per,ps_admin_amt,ps_admin_based_on";
                sql += " ,ps_edli_per,ps_edli_amt,ps_edli_based_on,ps_esi_emplr_per";
                sql += " ,ps_esi_limit,ps_pf_emplr_pension_per,ps_pf_cel_limit";
                sql += " ,ps_pf_cel_limit_amt,ps_esi_emply_per,ps_pf_per,ps_pf_col_excluded,ps_pf_br_region ";
                sql += " from payroll_setting a ";
                sql += " where a.rec_company_code = '" + comp_code + "'";
                sql += " and a.rec_branch_code = '" + branch_code + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new payrollsetting();
                    mRow.ps_pkid = Dr["ps_pkid"].ToString();
                    mRow.ps_admin_per = Lib.Conv2Decimal(Dr["ps_admin_per"].ToString());
                    mRow.ps_admin_amt = Lib.Conv2Decimal(Dr["ps_admin_amt"].ToString());
                    mRow.ps_admin_based_on = Dr["ps_admin_based_on"].ToString();
                    mRow.ps_edli_per = Lib.Conv2Decimal(Dr["ps_edli_per"].ToString());
                    mRow.ps_edli_amt = Lib.Conv2Decimal(Dr["ps_edli_amt"].ToString());
                    mRow.ps_edli_based_on = Dr["ps_edli_based_on"].ToString();
                    mRow.ps_esi_emplr_per = Lib.Conv2Decimal(Dr["ps_esi_emplr_per"].ToString());
                    mRow.ps_esi_limit = Lib.Conv2Decimal(Dr["ps_esi_limit"].ToString());
                    mRow.ps_pf_emplr_pension_per = Lib.Conv2Decimal(Dr["ps_pf_emplr_pension_per"].ToString());
                    mRow.ps_pf_cel_limit = Lib.Conv2Decimal(Dr["ps_pf_cel_limit"].ToString());
                    mRow.ps_pf_cel_limit_amt = Lib.Conv2Decimal(Dr["ps_pf_cel_limit_amt"].ToString());

                    mRow.ps_esi_emply_per = Lib.Conv2Decimal(Dr["ps_esi_emply_per"].ToString());
                    mRow.ps_pf_per = Lib.Conv2Decimal(Dr["ps_pf_per"].ToString());
                    mRow.ps_pf_col_excluded = Dr["ps_pf_col_excluded"].ToString();
                    mRow.ps_pf_br_region = Dr["ps_pf_br_region"].ToString();
                    break;
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(payrollsetting Record)
        {
            string str = "";
            try
            {


            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(payrollsetting Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {

                Con_Oracle = new DBConnection();

                if ((ErrorMessage = AllValid(Record)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                sql = "select ps_pkid from payroll_setting a ";
                sql += " where a.rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and a.rec_branch_code = '" + Record._globalvariables.branch_code + "'";

                DataTable Dt_temp = Con_Oracle.ExecuteQuery(sql);
                Record.rec_mode = "ADD";
                Record.ps_pkid = Guid.NewGuid().ToString().ToUpper();
                if (Dt_temp.Rows.Count > 0)
                {
                    Record.ps_pkid = Dt_temp.Rows[0]["ps_pkid"].ToString();
                    Record.rec_mode = "EDIT";
                }

                DBRecord Rec = new DBRecord();

                Rec.CreateRow("payroll_setting", Record.rec_mode, "ps_pkid", Record.ps_pkid);
                Rec.InsertNumeric("ps_admin_per", Record.ps_admin_per.ToString());
                Rec.InsertString("ps_admin_based_on", Record.ps_admin_based_on);
                Rec.InsertNumeric("ps_admin_amt", Record.ps_admin_amt.ToString());
                Rec.InsertNumeric("ps_edli_per", Record.ps_edli_per.ToString());
                Rec.InsertString("ps_edli_based_on", Record.ps_edli_based_on);
                Rec.InsertNumeric("ps_edli_amt", Record.ps_edli_amt.ToString());
                Rec.InsertNumeric("ps_esi_emplr_per", Record.ps_esi_emplr_per.ToString());
                Rec.InsertNumeric("ps_esi_limit", Record.ps_esi_limit.ToString());
                Rec.InsertNumeric("ps_pf_emplr_pension_per", Record.ps_pf_emplr_pension_per.ToString());
                Rec.InsertNumeric("ps_pf_cel_limit", Record.ps_pf_cel_limit.ToString());
                Rec.InsertNumeric("ps_pf_cel_limit_amt", Record.ps_pf_cel_limit_amt.ToString());
                Rec.InsertNumeric("ps_esi_emply_per", Record.ps_esi_emply_per.ToString());
                Rec.InsertNumeric("ps_pf_per", Record.ps_pf_per.ToString());
                Rec.InsertString("ps_pf_col_excluded", Record.ps_pf_col_excluded);
                Rec.InsertString("ps_pf_br_region", Record.ps_pf_br_region);
                if (Record.rec_mode == "ADD")
                {
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
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            return RetData;
        }

        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            // Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "LINER BOOKING STATUS");
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("statuslist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "actypem");
            //RetData.Add("actypem", lovservice.Lov(parameter)["actypem"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "acgroupm");
            //RetData.Add("acgroupm", lovservice.Lov(parameter)["acgroupm"]);

            return RetData;
        }

    }
}
