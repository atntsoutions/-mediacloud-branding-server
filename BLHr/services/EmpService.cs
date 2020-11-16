using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using XL.XSheet;
using DataBase_Oracle.Connections;
using System.Drawing;

namespace BLHr
{
    public class EmpService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;
        int iRow = 0;
        int iCol = 0;
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string report_folder = "";
        
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            
            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            DataTable Dt_List = new DataTable();
            Con_Oracle = new DBConnection();
            List<Emp> mList = new List<Emp>();
            Emp mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_id = SearchData["branch_id"].ToString();
            string company_id = SearchData["company_id"].ToString();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            report_folder = SearchData["report_folder"].ToString();
            

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                 
                sWhere = " where a.rec_company_code ='" + company_code + "'";
                sWhere += " and a.rec_branch_code ='" + branch_code + "'";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.emp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                    sWhere += " or (";
                    sWhere += "  upper(a.emp_no) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                    sWhere += " or (";
                    sWhere += "  upper(a.emp_pfno) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                    sWhere += " or (";
                    sWhere += "  upper(a.emp_esino) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                    sWhere += " or (";
                    sWhere += "  upper(a.emp_adhar_no) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }

                if (type == "EXCEL")
                {
                    sql = " select emp_pkid, emp_no,emp_name, emp_alias, emp_father_name, emp_spouse_name, emp_blood_group,";
                    sql += "  emp_local_address1, emp_local_address2, emp_local_address3, emp_local_city,";
                    sql += "  st.param_name as emp_local_state_name, emp_local_pin, emp_local_pobox,";
                    sql += "  emp_home_address1, emp_home_address2, emp_home_address3, emp_home_city,";
                    sql += "  sthm.param_name as emp_home_state_name, emp_home_pin, emp_home_pobox,   ";
                    sql += "  emp_tel_resi,emp_tel_office,emp_mobile,emp_mobile_office,emp_email_personal,emp_email_office,";
                    sql += "  emp_bank_acno,emp_bank_name,emp_bank_branch,emp_ifsc_code,";
                    sql += "  emp_pfno,emp_esino,emp_pan,emp_adhar_no,emp_uan_no,";
                    sql += "  emp_fuel_type,emp_fuel_limit,emp_bus_limit,emp_train_limit,";
                    sql += "  emp_vehi_maint_limit,emp_drive_vehi_type,emp_mobile_limit,emp_datacard_limit,";
                    sql += "  grade.param_name as emp_grade_name,dept.param_name as emp_department_name,desig.param_name as emp_designation_name ,";
                    sql += "  status.param_name as emp_status_name,emp_in_payroll,";
                    sql += "  emp_marital_status,emp_gender,emp_comp_mediclaim,emp_premium_amt,emp_mediclaim_provider,";
                    sql += "  emp_remarks,emp_do_birth,emp_marrige_date,emp_do_joining,";
                    sql += "  emp_do_confirmation,emp_do_relieve,emp_is_relieved,emp_trans_date,a.rec_branch_code ";
                    sql += "  from empm a  ";
                    sql += "  left join param st on a.emp_local_state_id = st.param_pkid";
                    sql += "  left join param sthm on a.emp_home_state_id = sthm.param_pkid";
                    sql += "  left join param grade on a.emp_grade_id = grade.param_pkid";
                    sql += "  left join param dept on a.emp_department_id = dept.param_pkid";
                    sql += "  left join param desig on a.emp_designation_id = desig.param_pkid";
                    sql += "  left join param status on a.emp_status_id = status.param_pkid";
                    sql += sWhere;
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mRow = new Emp();
                        mRow.emp_pkid = Dr["emp_pkid"].ToString();
                        mRow.emp_no = Dr["emp_no"].ToString();
                        mRow.emp_name = Dr["emp_name"].ToString();
                        mRow.emp_alias = Dr["emp_alias"].ToString();
                        mRow.emp_father_name = Dr["emp_father_name"].ToString();
                        mRow.emp_spouse_name = Dr["emp_spouse_name"].ToString();
                        if (Dr["emp_blood_group"].ToString().Trim() == "")
                            mRow.emp_blood_group = "N/A";
                        else
                            mRow.emp_blood_group = Dr["emp_blood_group"].ToString();
                        mRow.emp_local_address1 = Dr["emp_local_address1"].ToString();
                        mRow.emp_local_address2 = Dr["emp_local_address2"].ToString();
                        mRow.emp_local_address3 = Dr["emp_local_address3"].ToString();
                        mRow.emp_local_city = Dr["emp_local_city"].ToString();
                        mRow.emp_local_state_name = Dr["emp_local_state_name"].ToString();
                        mRow.emp_local_pin = Dr["emp_local_pin"].ToString();
                        mRow.emp_local_pobox = Dr["emp_local_pobox"].ToString();
                        
                        mRow.emp_home_address1 = Dr["emp_home_address1"].ToString();
                        mRow.emp_home_address2 = Dr["emp_home_address2"].ToString();
                        mRow.emp_home_address3 = Dr["emp_home_address3"].ToString();
                        mRow.emp_home_city = Dr["emp_home_city"].ToString();
                        mRow.emp_home_state_name = Dr["emp_home_state_name"].ToString();
                        mRow.emp_home_pin = Dr["emp_home_pin"].ToString();
                        mRow.emp_home_pobox = Dr["emp_home_pobox"].ToString();
                        
                        mRow.emp_tel_resi = Dr["emp_tel_resi"].ToString();
                        mRow.emp_tel_office = Dr["emp_tel_office"].ToString();
                        mRow.emp_mobile = Dr["emp_mobile"].ToString();
                        mRow.emp_mobile_office = Dr["emp_mobile_office"].ToString();
                        mRow.emp_email_personal = Dr["emp_email_personal"].ToString();
                        mRow.emp_email_office = Dr["emp_email_office"].ToString();
                        mRow.emp_bank_acno = Dr["emp_bank_acno"].ToString();
                        mRow.emp_bank_name = Dr["emp_bank_name"].ToString();
                        mRow.emp_bank_branch = Dr["emp_bank_branch"].ToString();
                        mRow.emp_ifsc_code = Dr["emp_ifsc_code"].ToString();
                        mRow.emp_pfno = Dr["emp_pfno"].ToString();
                        mRow.emp_esino = Dr["emp_esino"].ToString();
                        mRow.emp_pan = Dr["emp_pan"].ToString();
                        mRow.emp_adhar_no = Dr["emp_adhar_no"].ToString();
                        mRow.emp_uan_no = Dr["emp_uan_no"].ToString();
                        mRow.emp_fuel_type = Dr["emp_fuel_type"].ToString();
                        mRow.emp_fuel_limit = Lib.Conv2Decimal(Dr["emp_fuel_limit"].ToString());
                        mRow.emp_bus_limit = Lib.Conv2Decimal(Dr["emp_bus_limit"].ToString());
                        mRow.emp_train_limit = Lib.Conv2Decimal(Dr["emp_train_limit"].ToString());
                        mRow.emp_vehi_maint_limit = Lib.Conv2Decimal(Dr["emp_vehi_maint_limit"].ToString());
                        mRow.emp_drive_vehi_type = Dr["emp_drive_vehi_type"].ToString();
                        mRow.emp_mobile_limit = Lib.Conv2Decimal(Dr["emp_mobile_limit"].ToString());
                        mRow.emp_datacard_limit = Dr["emp_datacard_limit"].ToString();
                        mRow.emp_grade_name = Dr["emp_grade_name"].ToString();
                        mRow.emp_department_name = Dr["emp_department_name"].ToString();
                        mRow.emp_designation_name = Dr["emp_designation_name"].ToString();
                        mRow.emp_status_name = Dr["emp_status_name"].ToString();
                        if (Dr["emp_in_payroll"].ToString() == "Y")
                            mRow.emp_in_payroll = true;
                        else
                            mRow.emp_in_payroll = false;

                        mRow.emp_marital_status = Dr["emp_marital_status"].ToString();
                        mRow.emp_gender = Dr["emp_gender"].ToString();
                        if (Dr["emp_comp_mediclaim"].ToString() == "Y")
                            mRow.emp_comp_mediclaim = true;
                        else
                            mRow.emp_comp_mediclaim = false;

                        mRow.emp_premium_amt = Lib.Conv2Decimal(Dr["emp_premium_amt"].ToString());
                        mRow.emp_mediclaim_provider = Dr["emp_mediclaim_provider"].ToString();
                        mRow.emp_remarks = Dr["emp_remarks"].ToString();
                        mRow.emp_do_birth = Lib.DatetoStringDisplayformat(Dr["emp_do_birth"]);
                        mRow.emp_marrige_date = Lib.DatetoStringDisplayformat(Dr["emp_marrige_date"]);
                        mRow.emp_do_joining = Lib.DatetoStringDisplayformat(Dr["emp_do_joining"]);
                        mRow.emp_do_confirmation = Lib.DatetoStringDisplayformat(Dr["emp_do_confirmation"]);
                        mRow.emp_do_relieve = Lib.DatetoStringDisplayformat(Dr["emp_do_relieve"]);
                        if (Dr["emp_is_relieved"].ToString() == "Y")
                            mRow.emp_is_relieved = true;
                        else
                            mRow.emp_is_relieved = false;
                        mRow.emp_trans_date = Lib.DatetoStringDisplayformat(Dr["emp_trans_date"]);
                        mRow.rec_branch_code = Dr["rec_branch_code"].ToString();
                        
                        mList.Add(mRow);
                    }

                    if (mList != null)
                    {
                        PrintEmpReport(mList, company_code);
                    }
                }
                else
                {
                    if (type == "NEW")
                    {
                        sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM empm  a ";
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


                    Dt_List = new DataTable();
                    sql = "";
                    sql += " select * from ( ";
                    sql += " select emp_pkid,emp_no, emp_name,b.comp_name as branch_name,emp_blood_group,emp_mobile,emp_email_personal,";
                    sql += " emp_bank_acno,emp_pfno,emp_do_joining,emp_do_confirmation,emp_do_relieve,emp_trans_date,";
                    sql += " row_number() over(order by emp_no, emp_name) rn ";
                    sql += "  from empm a ";
                    sql += " left join companym b on a.rec_branch_code = b.comp_code ";
                    sql += sWhere;
                    sql += ") a where rn between {startrow} and {endrow}";
                    sql += " order by emp_no,emp_name";

                    sql = sql.Replace("{startrow}", startrow.ToString());
                    sql = sql.Replace("{endrow}", endrow.ToString());

                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mRow = new Emp();
                        mRow.emp_pkid = Dr["emp_pkid"].ToString();
                        mRow.emp_no = Dr["emp_no"].ToString();
                        mRow.emp_name = Dr["emp_name"].ToString();
                        mRow.emp_branch_name = Dr["branch_name"].ToString();
                        mRow.emp_blood_group = Dr["emp_blood_group"].ToString();
                        mRow.emp_mobile = Dr["emp_mobile"].ToString();
                        mRow.emp_email_personal = Dr["emp_email_personal"].ToString();
                        mRow.emp_bank_acno = Dr["emp_bank_acno"].ToString();
                        mRow.emp_pfno = Dr["emp_pfno"].ToString();
                        mRow.emp_do_joining = Lib.DatetoStringDisplayformat(Dr["emp_do_joining"]);
                        mRow.emp_do_confirmation = Lib.DatetoStringDisplayformat(Dr["emp_do_confirmation"]);
                        mRow.emp_do_relieve = Lib.DatetoStringDisplayformat(Dr["emp_do_relieve"]);
                        mRow.emp_trans_date = Lib.DatetoStringDisplayformat(Dr["emp_trans_date"]);
                        mList.Add(mRow);
                    }
                }
                
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }


    


        public Dictionary<string, object>  GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
           Emp mRow =new Emp();
            
            string id = SearchData["pkid"].ToString();

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql  = " select emp_pkid, emp_no,emp_name, emp_alias, emp_father_name, emp_spouse_name, emp_blood_group,";
                sql += " emp_local_address1, emp_local_address2, emp_local_address3, emp_local_city,";
                sql += " emp_local_state_id, emp_local_pin, emp_local_pobox, emp_local_country_id ,";
                sql += " emp_home_address1, emp_home_address2, emp_home_address3, emp_home_city,"; 
                sql += " emp_home_state_id, emp_home_pin, emp_home_pobox, emp_home_country_id, ";
                sql += " emp_tel_resi,emp_tel_office,emp_mobile,emp_mobile_office,emp_email_personal,emp_email_office,";
                sql += " emp_bank_acno,emp_bank_name,emp_bank_branch,emp_ifsc_code,";
                sql += " emp_pfno,emp_esino,emp_pan,emp_adhar_no,emp_uan_no,";
                sql += " emp_fuel_type,emp_fuel_limit,emp_bus_limit,emp_train_limit,";
                sql += " emp_vehi_maint_limit,emp_drive_vehi_type,emp_mobile_limit,emp_datacard_limit,";
                sql += " emp_grade_id,emp_department_id,emp_designation_id,emp_status_id,emp_in_payroll,";
                sql += " emp_marital_status,emp_gender,emp_comp_mediclaim,emp_premium_amt,emp_mediclaim_provider,";
                sql += " emp_remarks,emp_do_birth,emp_marrige_date,emp_do_joining,";
                sql += " emp_do_confirmation,emp_do_relieve,emp_is_relieved,emp_trans_date,a.rec_branch_code,emp_branch_group,emp_is_retired ";
                sql += " from empm a  ";
                sql += " where  a.emp_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Emp();
                    mRow.emp_pkid = Dr["emp_pkid"].ToString();
                    mRow.emp_no = Dr["emp_no"].ToString();
                    mRow.emp_name = Dr["emp_name"].ToString();
                    mRow.emp_alias = Dr["emp_alias"].ToString();
                    mRow.emp_father_name = Dr["emp_father_name"].ToString();
                    mRow.emp_spouse_name = Dr["emp_spouse_name"].ToString();
                    if (Dr["emp_blood_group"].ToString().Trim() == "")
                        mRow.emp_blood_group = "N/A";
                    else
                        mRow.emp_blood_group = Dr["emp_blood_group"].ToString();
                    mRow.emp_local_address1 = Dr["emp_local_address1"].ToString();
                    mRow.emp_local_address2 = Dr["emp_local_address2"].ToString();
                    mRow.emp_local_address3 = Dr["emp_local_address3"].ToString();
                    mRow.emp_local_city = Dr["emp_local_city"].ToString();
                    mRow.emp_local_state_id = Dr["emp_local_state_id"].ToString();
                    mRow.emp_local_pin = Dr["emp_local_pin"].ToString();
                    mRow.emp_local_pobox = Dr["emp_local_pobox"].ToString();
                    mRow.emp_local_country_id = Dr["emp_local_country_id"].ToString();
                    mRow.emp_home_address1 = Dr["emp_home_address1"].ToString();
                    mRow.emp_home_address2 = Dr["emp_home_address2"].ToString();
                    mRow.emp_home_address3 = Dr["emp_home_address3"].ToString();
                    mRow.emp_home_city = Dr["emp_home_city"].ToString();
                    mRow.emp_home_state_id = Dr["emp_home_state_id"].ToString();
                    mRow.emp_home_pin = Dr["emp_home_pin"].ToString();
                    mRow.emp_home_pobox = Dr["emp_home_pobox"].ToString();
                    mRow.emp_home_country_id = Dr["emp_home_country_id"].ToString();
                    mRow.emp_tel_resi = Dr["emp_tel_resi"].ToString();
                    mRow.emp_tel_office = Dr["emp_tel_office"].ToString();
                    mRow.emp_mobile = Dr["emp_mobile"].ToString();
                    mRow.emp_mobile_office = Dr["emp_mobile_office"].ToString();
                    mRow.emp_email_personal = Dr["emp_email_personal"].ToString();
                    mRow.emp_email_office = Dr["emp_email_office"].ToString();
                    mRow.emp_bank_acno = Dr["emp_bank_acno"].ToString();
                    mRow.emp_bank_name = Dr["emp_bank_name"].ToString();
                    mRow.emp_bank_branch = Dr["emp_bank_branch"].ToString();
                    mRow.emp_ifsc_code = Dr["emp_ifsc_code"].ToString();
                    mRow.emp_pfno = Dr["emp_pfno"].ToString();
                    mRow.emp_esino = Dr["emp_esino"].ToString();
                    mRow.emp_pan = Dr["emp_pan"].ToString();
                    mRow.emp_adhar_no = Dr["emp_adhar_no"].ToString();
                    mRow.emp_uan_no = Dr["emp_uan_no"].ToString();
                    mRow.emp_fuel_type = Dr["emp_fuel_type"].ToString();
                    mRow.emp_fuel_limit = Lib.Conv2Decimal(Dr["emp_fuel_limit"].ToString());
                    mRow.emp_bus_limit = Lib.Conv2Decimal(Dr["emp_bus_limit"].ToString());
                    mRow.emp_train_limit = Lib.Conv2Decimal(Dr["emp_train_limit"].ToString());
                    mRow.emp_vehi_maint_limit = Lib.Conv2Decimal(Dr["emp_vehi_maint_limit"].ToString());
                    mRow.emp_drive_vehi_type = Dr["emp_drive_vehi_type"].ToString();
                    mRow.emp_mobile_limit = Lib.Conv2Decimal(Dr["emp_mobile_limit"].ToString());
                    mRow.emp_datacard_limit = Dr["emp_datacard_limit"].ToString();
                    mRow.emp_grade_id = Dr["emp_grade_id"].ToString();
                    mRow.emp_department_id = Dr["emp_department_id"].ToString();
                    mRow.emp_designation_id = Dr["emp_designation_id"].ToString();
                    mRow.emp_status_id = Dr["emp_status_id"].ToString();
                    if (Dr["emp_in_payroll"].ToString() == "Y")
                        mRow.emp_in_payroll = true;
                    else
                        mRow.emp_in_payroll = false;

                    mRow.emp_marital_status = Dr["emp_marital_status"].ToString();
                    mRow.emp_gender = Dr["emp_gender"].ToString();
                    if (Dr["emp_comp_mediclaim"].ToString() == "Y")
                        mRow.emp_comp_mediclaim = true;
                    else
                        mRow.emp_comp_mediclaim = false;

                    mRow.emp_premium_amt = Lib.Conv2Decimal(Dr["emp_premium_amt"].ToString());
                    mRow.emp_mediclaim_provider = Dr["emp_mediclaim_provider"].ToString();
                    mRow.emp_remarks = Dr["emp_remarks"].ToString();
                    mRow.emp_do_birth = Lib.DatetoString(Dr["emp_do_birth"]);
                    mRow.emp_marrige_date = Lib.DatetoString(Dr["emp_marrige_date"]);
                    mRow.emp_do_joining = Lib.DatetoString(Dr["emp_do_joining"]);
                    mRow.emp_do_confirmation = Lib.DatetoString(Dr["emp_do_confirmation"]);
                    mRow.emp_do_relieve = Lib.DatetoString(Dr["emp_do_relieve"]);
                    if (Dr["emp_is_relieved"].ToString() == "Y")
                        mRow.emp_is_relieved = true;
                    else
                        mRow.emp_is_relieved = false;
                    mRow.emp_trans_date = Lib.DatetoString(Dr["emp_trans_date"]);
                    mRow.rec_branch_code = Dr["rec_branch_code"].ToString();
                    mRow.emp_branch_group = Lib.Conv2Integer(Dr["emp_branch_group"].ToString());
                    if (Dr["emp_is_retired"].ToString() == "Y")
                        mRow.emp_is_retired = true;
                    else
                        mRow.emp_is_retired = false;

                    break;
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            return RetData;
        }


        public string AllValid(Emp Record)
        { 
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {

                if (Record.emp_name.Trim().Length <= 0)
                    Lib.AddError(ref str, "| Name Cannot Be Empty");
                if (Record.rec_branch_code.Trim().Length <= 0)
                    Lib.AddError(ref str, " | Branch Cannot Be Empty");
                if (Record.emp_no.Trim().Length > 0)
                {
                    sql = "select emp_pkid from (";
                    sql += "select emp_pkid  from empm a where a.rec_company_code = '" + Record._globalvariables.comp_code + "' and a.emp_no = '{CODE}'  ";
                    sql += ") a where emp_pkid <> '{PKID}'";

                    sql = sql.Replace("{CODE}", Record.emp_no);
                    sql = sql.Replace("{PKID}", Record.emp_pkid);

                    if (Con_Oracle.IsRowExists(sql))
                        Lib.AddError(ref str, " | Code Exists");
                }

                if (Record.emp_name.Trim().Length > 0)
                {
                    sql = "select emp_pkid from (";
                    sql += "select emp_pkid  from empm a where a.rec_company_code = '" + Record._globalvariables.comp_code + "' and a.emp_name = '{NAME}'  ";
                    sql += ") a where emp_pkid <> '{PKID}'";

                    sql = sql.Replace("{NAME}", Record.emp_name);
                    sql = sql.Replace("{PKID}", Record.emp_pkid);


                    if (Con_Oracle.IsRowExists(sql))
                       Lib.AddError(ref str, " | Name Exists");
                }

                if (Record.emp_do_joining.Trim().Length > 0 && Record.emp_do_birth.Trim().Length > 0)
                {
                    DateTime dob = DateTime.Parse(Record.emp_do_birth);
                    DateTime doj = DateTime.Parse(Record.emp_do_joining);

                    if (dob > doj)
                        Lib.AddError(ref str, " | Joining Date should be greater than DOB ");
                }

                if (Record.emp_do_confirmation.Trim().Length > 0 && Record.emp_do_birth.Trim().Length > 0)
                {
                    
                    DateTime dob = DateTime.Parse(Record.emp_do_birth);
                    DateTime doc = DateTime.Parse(Record.emp_do_confirmation);


                    if (dob > doc)
                        Lib.AddError(ref str, " | Confirmation Date should be greater than DOB");

                }

                if (Record.emp_do_relieve.Trim().Length > 0 && Record.emp_do_birth.Trim().Length > 0)
                {

                    DateTime dob = DateTime.Parse(Record.emp_do_birth);
                    DateTime dor = DateTime.Parse(Record.emp_do_relieve);
                    
                    if (dob > dor)
                        Lib.AddError(ref str, " | Relieve Date should be greater than DOB");

                }

                //if (Record.rec_branch_code != Record._globalvariables.branch_code)
                //{
                //    Lib.AddError(ref str, " | selected branch and login branch are mismatch");
                //}

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        
        public Dictionary<string, object> Save(Emp Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

               

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);
                if (Record.emp_blood_group == "N/A")
                    Record.emp_blood_group = "";
                if (Record.emp_branch_group <= 0)
                    Record.emp_branch_group = 1;


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("empm", Record.rec_mode, "emp_pkid", Record.emp_pkid);
                Rec.InsertString("emp_no", Record.emp_no);
                Rec.InsertString("emp_name", Record.emp_name);
                Rec.InsertString("emp_alias", Record.emp_alias);
                Rec.InsertString("emp_father_name", Record.emp_father_name);
                Rec.InsertString("emp_spouse_name", Record.emp_spouse_name);
                Rec.InsertString("emp_blood_group", Record.emp_blood_group);
                Rec.InsertString("emp_local_address1", Record.emp_local_address1);
                Rec.InsertString("emp_local_address2", Record.emp_local_address2);
                Rec.InsertString("emp_local_address3", Record.emp_local_address3);
                Rec.InsertString("emp_local_city", Record.emp_local_city);
                Rec.InsertString("emp_local_state_id", Record.emp_local_state_id);
                Rec.InsertString("emp_local_pin", Record.emp_local_pin);
                Rec.InsertString("emp_local_pobox", Record.emp_local_pobox);
                Rec.InsertString("emp_local_country_id", Record.emp_local_country_id);
                Rec.InsertString("emp_home_address1", Record.emp_home_address1);
                Rec.InsertString("emp_home_address2", Record.emp_home_address2);
                Rec.InsertString("emp_home_address3", Record.emp_home_address3);
                Rec.InsertString("emp_home_city", Record.emp_home_city);
                Rec.InsertString("emp_home_state_id", Record.emp_home_state_id);
                Rec.InsertString("emp_home_pin", Record.emp_home_pin);
                Rec.InsertString("emp_home_pobox", Record.emp_home_pobox);
                Rec.InsertString("emp_home_country_id", Record.emp_home_country_id);
                Rec.InsertString("emp_tel_resi", Record.emp_tel_resi);
                Rec.InsertString("emp_tel_office", Record.emp_tel_office);
                Rec.InsertString("emp_mobile", Record.emp_mobile);
                Rec.InsertString("emp_mobile_office", Record.emp_mobile_office);
                Rec.InsertString("emp_email_personal", Record.emp_email_personal);
                Rec.InsertString("emp_email_office", Record.emp_email_office);
                Rec.InsertString("emp_bank_acno", Record.emp_bank_acno);
                Rec.InsertString("emp_bank_name", Record.emp_bank_name);
                Rec.InsertString("emp_bank_branch", Record.emp_bank_branch);
                Rec.InsertString("emp_ifsc_code", Record.emp_ifsc_code);
                Rec.InsertString("emp_pfno", Record.emp_pfno);
                Rec.InsertString("emp_esino", Record.emp_esino);
                Rec.InsertString("emp_pan", Record.emp_pan);
                Rec.InsertString("emp_adhar_no", Record.emp_adhar_no);
                Rec.InsertString("emp_uan_no", Record.emp_uan_no);
                Rec.InsertString("emp_fuel_type", Record.emp_fuel_type);
                Rec.InsertNumeric("emp_fuel_limit", Record.emp_fuel_limit.ToString());
                Rec.InsertNumeric("emp_bus_limit", Record.emp_bus_limit.ToString());
                Rec.InsertNumeric("emp_train_limit", Record.emp_train_limit.ToString());
                Rec.InsertNumeric("emp_vehi_maint_limit", Record.emp_vehi_maint_limit.ToString());
                Rec.InsertString("emp_drive_vehi_type", Record.emp_drive_vehi_type);
                Rec.InsertNumeric("emp_mobile_limit", Record.emp_mobile_limit.ToString());
                Rec.InsertString("emp_datacard_limit", Record.emp_datacard_limit);
                Rec.InsertString("emp_grade_id", Record.emp_grade_id);
                Rec.InsertString("emp_department_id", Record.emp_department_id);
                Rec.InsertString("emp_designation_id", Record.emp_designation_id);
                Rec.InsertString("emp_status_id", Record.emp_status_id);
                if (Record.emp_in_payroll == true)

                    Rec.InsertString("emp_in_payroll", "Y");
                else
                    Rec.InsertString("emp_in_payroll", "N");

                Rec.InsertString("emp_marital_status", Record.emp_marital_status);
                Rec.InsertString("emp_gender", Record.emp_gender);
                if (Record.emp_comp_mediclaim == true)
                    Rec.InsertString("emp_comp_mediclaim", "Y");
                else
                    Rec.InsertString("emp_comp_mediclaim", "N");

                Rec.InsertNumeric("emp_premium_amt", Record.emp_premium_amt.ToString());
                Rec.InsertString("emp_mediclaim_provider", Record.emp_mediclaim_provider);
                Rec.InsertString("emp_remarks", Record.emp_remarks);
                Rec.InsertDate("emp_do_birth", Record.emp_do_birth);
                Rec.InsertDate("emp_marrige_date", Record.emp_marrige_date);
                Rec.InsertDate("emp_do_joining", Record.emp_do_joining);
                Rec.InsertDate("emp_do_confirmation", Record.emp_do_confirmation);
                Rec.InsertDate("emp_do_relieve", Record.emp_do_relieve);
             
                if (Record.emp_is_relieved == true)
                    Rec.InsertString("emp_is_relieved", "Y");
                else
                    Rec.InsertString("emp_is_relieved", "N");

                Rec.InsertDate("emp_trans_date", Record.emp_trans_date);
                Rec.InsertString("rec_branch_code", Record.rec_branch_code);
                Rec.InsertNumeric("emp_branch_group", Record.emp_branch_group.ToString());
                if (Record.emp_is_retired == true)
                    Rec.InsertString("emp_is_retired", "Y");
                else
                    Rec.InsertString("emp_is_retired", "N");

                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                   // Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
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

                if(Record.rec_branch_code != Record._globalvariables.branch_code)
                {
                    sql = "update salarym set rec_branch_code='" + Record.rec_branch_code + "' where sal_emp_id='" + Record.emp_pkid + "' and sal_month=sal_year";
                    Con_Oracle.ExecuteNonQuery(sql);
                }
                //Costcenter Updation 
                sql = Lib.GetCostCenterSQL(Record.rec_mode, Record.emp_pkid, Record.emp_no, Record.emp_name,Record.emp_do_joining.ToString(), "EMPLOYEE",
                    Record._globalvariables.year_code, "", Record._globalvariables.comp_code, Record._globalvariables.branch_code, "EMPLOYEE");
                Con_Oracle.ExecuteNonQuery(sql);

                Con_Oracle.CommitTransaction();
                

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
            return RetData;
        }

        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            string comp_code = "";
            if (SearchData.ContainsKey("comp_code"))
                comp_code = SearchData["comp_code"].ToString();

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "COUNTRY");
            parameter.Add("comp_code", comp_code);
            RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "STATE");
            parameter.Add("comp_code", comp_code);
            RetData.Add("statelist", lovservice.Lov(parameter)["param"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "branch");
            parameter.Add("comp_code", comp_code);
            RetData.Add("branchlist", lovservice.Lov(parameter)["branch"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "EMPLOYEE GRADE");
            parameter.Add("comp_code", comp_code);
            RetData.Add("gradelist", lovservice.Lov(parameter)["param"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "EMPLOYEE DEPARTMENT");
            parameter.Add("comp_code", comp_code);
            RetData.Add("departmentlist", lovservice.Lov(parameter)["param"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "EMPLOYEE DESIGNATION");
            parameter.Add("comp_code", comp_code);
            RetData.Add("designationlist", lovservice.Lov(parameter)["param"]);

            parameter = new Dictionary<string, object>();
            parameter.Add("table", "param");
            parameter.Add("param_type", "EMPLOYEE STATUS");
            parameter.Add("comp_code", comp_code);
            RetData.Add("statuslist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }

        private void PrintEmpReport(List<Emp> mList, string comp_code)
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPADD3 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";
            string FolderId = "";
            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = "EMPLOYEE LIST";

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "COMP_ADDRESS");
                mSearchData.Add("comp_code", comp_code);
                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPADD3 = Dr["COMP_ADDRESS3"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                FolderId = Guid.NewGuid().ToString().ToUpper();
                File_Display_Name = "EmpReport.xls";
                File_Name = Lib.GetFileName(report_folder, FolderId, File_Display_Name);

                WB = new ExcelFile();
                WB.Worksheets.Add("Report");
                WS = WB.Worksheets["Report"];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 5;
                WS.Columns[2].Width = 256 * 7;
                WS.Columns[3].Width = 256 * 26;
                WS.Columns[4].Width = 256 * 5;
                WS.Columns[5].Width = 256 * 24;
                WS.Columns[6].Width = 256 * 27;
                WS.Columns[7].Width = 256 * 13;
                WS.Columns[8].Width = 256 * 36;
                WS.Columns[9].Width = 256 * 28;
                WS.Columns[10].Width = 256 * 32;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 13;
                WS.Columns[13].Width = 256 * 9;
                WS.Columns[14].Width = 256 * 12;
                WS.Columns[15].Width = 256 * 35;
                WS.Columns[16].Width = 256 * 41;
                WS.Columns[17].Width = 256 * 32;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 13;
                WS.Columns[20].Width = 256 * 9;
                WS.Columns[21].Width = 256 * 12;
                WS.Columns[22].Width = 256 * 13;
                WS.Columns[23].Width = 256 * 9;
                WS.Columns[24].Width = 256 * 11;
                WS.Columns[25].Width = 256 * 13;
                WS.Columns[26].Width = 256 * 37;
                WS.Columns[27].Width = 256 * 25;
                WS.Columns[28].Width = 256 * 18;
                WS.Columns[29].Width = 256 * 20;
                WS.Columns[30].Width = 256 * 24;
                WS.Columns[31].Width = 256 * 12;
                WS.Columns[32].Width = 256 * 25;
                WS.Columns[33].Width = 256 * 11;
                WS.Columns[34].Width = 256 * 12;
                WS.Columns[35].Width = 256 * 14;
                WS.Columns[36].Width = 256 * 13;
                WS.Columns[37].Width = 256 * 9;
                WS.Columns[38].Width = 256 * 9;
                WS.Columns[39].Width = 256 * 9;
                WS.Columns[40].Width = 256 * 10;
                WS.Columns[41].Width = 256 * 16;
                WS.Columns[42].Width = 256 * 12;
                WS.Columns[43].Width = 256 * 15;
                WS.Columns[44].Width = 256 * 22;
                WS.Columns[45].Width = 256 * 18;
                WS.Columns[46].Width = 256 * 24;
                WS.Columns[47].Width = 256 * 12;
                WS.Columns[48].Width = 256 * 10;
                WS.Columns[49].Width = 256 * 15;
                WS.Columns[50].Width = 256 * 7;
                WS.Columns[51].Width = 256 * 16;
                WS.Columns[52].Width = 256 * 13;
                WS.Columns[53].Width = 256 * 19;
                WS.Columns[54].Width = 256 * 61;
                WS.Columns[55].Width = 256 * 10;
                WS.Columns[56].Width = 256 * 13;
                WS.Columns[57].Width = 256 * 10;
                WS.Columns[58].Width = 256 * 17;
                WS.Columns[59].Width = 256 * 10;

                iRow = 0; iCol = 1;

                _Size = 14;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 12;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;
                if (str == "")
                    str = COMPADD3;
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, REPORT_CAPTION, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 383, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 326, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ALIAS", _Color, true, "BT", "L", "", _Size, false, 327, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FATHER-NAME", _Color, true, "BT", "L", "", _Size, false, 328, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SPOUSE-NAME", _Color, true, "BT", "L", "", _Size, false, 329, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BLOOD-GROUP", _Color, true, "BT", "L", "", _Size, false, 330, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-ADDRESS1", _Color, true, "BT", "L", "", _Size, false, 331, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-ADDRESS2", _Color, true, "BT", "L", "", _Size, false, 332, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-ADDRESS3", _Color, true, "BT", "L", "", _Size, false, 333, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-CITY", _Color, true, "BT", "L", "", _Size, false, 334, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-STATE", _Color, true, "BT", "L", "", _Size, false, 335, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-PIN", _Color, true, "BT", "L", "", _Size, false, 336, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LOCAL-POBOX", _Color, true, "BT", "L", "", _Size, false, 337, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOME-ADDRESS1", _Color, true, "BT", "L", "", _Size, false, 338, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOME-ADDRESS2", _Color, true, "BT", "L", "", _Size, false, 339, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOME-ADDRESS3", _Color, true, "BT", "L", "", _Size, false, 340, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOME-CITY", _Color, true, "BT", "L", "", _Size, false, 341, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOME-STATE", _Color, true, "BT", "L", "", _Size, false, 342, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOME-PIN", _Color, true, "BT", "L", "", _Size, false, 343, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HOME-POBOX", _Color, true, "BT", "L", "", _Size, false, 344, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TEL-RESI", _Color, true, "BT", "L", "", _Size, false, 345, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TEL-OFFICE", _Color, true, "BT", "L", "", _Size, false, 346, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MOBILE", _Color, true, "BT", "L", "", _Size, false, 347, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MOBILE-OFFICE", _Color, true, "BT", "L", "", _Size, false, 348, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MAIL-PERSONAL", _Color, true, "BT", "L", "", _Size, false, 349, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EMAIL-OFFICE", _Color, true, "BT", "L", "", _Size, false, 350, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BANK-ACNO", _Color, true, "BT", "L", "", _Size, false, 351, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BANK-NAME", _Color, true, "BT", "L", "", _Size, false, 352, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BANK-BRANCH", _Color, true, "BT", "L", "", _Size, false, 353, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IFSC-CODE", _Color, true, "BT", "L", "", _Size, false, 354, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PFNO", _Color, true, "BT", "L", "", _Size, false, 355, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ESINO", _Color, true, "BT", "L", "", _Size, false, 356, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PAN", _Color, true, "BT", "L", "", _Size, false, 357, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ADHAR-NO", _Color, true, "BT", "L", "", _Size, false, 358, "", true);
                Lib.WriteData(WS, iRow, iCol++, "UAN-NO", _Color, true, "BT", "L", "", _Size, false, 359, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FUEL-TYPE", _Color, true, "BT", "L", "", _Size, false, 360, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FUEL-LIMIT", _Color, true, "BT", "L", "", _Size, false, 361, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BUS-LIMIT", _Color, true, "BT", "L", "", _Size, false, 362, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TRAIN-LIMIT", _Color, true, "BT", "L", "", _Size, false, 363, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VEHI-MAINT-LIMIT", _Color, true, "BT", "L", "", _Size, false, 364, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MOBILE-LIMIT", _Color, true, "BT", "L", "", _Size, false, 365, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATACARD-LIMIT", _Color, true, "BT", "L", "", _Size, false, 366, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GRADE-NAME", _Color, true, "BT", "L", "", _Size, false, 367, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEPARTMENT-NAME", _Color, true, "BT", "L", "", _Size, false, 368, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DESIGNATION-NAME", _Color, true, "BT", "L", "", _Size, false, 369, "", true);
                Lib.WriteData(WS, iRow, iCol++, "STATUS-NAME", _Color, true, "BT", "L", "", _Size, false, 370, "", true);
                Lib.WriteData(WS, iRow, iCol++, "IN-PAYROLL", _Color, true, "BT", "L", "", _Size, false, 371, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MARITAL-STATUS", _Color, true, "BT", "L", "", _Size, false, 372, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GENDER", _Color, true, "BT", "L", "", _Size, false, 373, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COMP-MEDICLAIM", _Color, true, "BT", "L", "", _Size, false, 374, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PREMIUM-AMT", _Color, true, "BT", "L", "", _Size, false, 375, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MEDICLAIM-PROVIDER", _Color, true, "BT", "L", "", _Size, false, 376, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, true, "BT", "L", "", _Size, false, 377, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DO-BIRTH", _Color, true, "BT", "L", "", _Size, false, 378, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MARRIGE-DATE", _Color, true, "BT", "L", "", _Size, false, 379, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DO-JOINING", _Color, true, "BT", "L", "", _Size, false, 380, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DO-CONFIRMATION", _Color, true, "BT", "L", "", _Size, false, 381, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DO-RELIEVE", _Color, true, "BT", "L", "", _Size, false, 382, "", true);
                foreach (Emp Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.rec_branch_code, _Color, false, "", "L", "", _Size, false, 383, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_name, _Color, false, "", "L", "", _Size, false, 326, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_alias, _Color, false, "", "L", "", _Size, false, 327, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_father_name, _Color, false, "", "L", "", _Size, false, 328, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_spouse_name, _Color, false, "", "L", "", _Size, false, 329, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_blood_group, _Color, false, "", "L", "", _Size, false, 330, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_local_address1, _Color, false, "", "L", "", _Size, false, 331, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_local_address2, _Color, false, "", "L", "", _Size, false, 332, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_local_address3, _Color, false, "", "L", "", _Size, false, 333, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_local_city, _Color, false, "", "L", "", _Size, false, 334, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_local_state_name, _Color, false, "", "L", "", _Size, false, 335, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_local_pin, _Color, false, "", "L", "", _Size, false, 336, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_local_pobox, _Color, false, "", "L", "", _Size, false, 337, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_home_address1, _Color, false, "", "L", "", _Size, false, 338, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_home_address2, _Color, false, "", "L", "", _Size, false, 339, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_home_address3, _Color, false, "", "L", "", _Size, false, 340, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_home_city, _Color, false, "", "L", "", _Size, false, 341, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_home_state_name, _Color, false, "", "L", "", _Size, false, 342, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_home_pin, _Color, false, "", "L", "", _Size, false, 343, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_home_pobox, _Color, false, "", "L", "", _Size, false, 344, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_tel_resi, _Color, false, "", "L", "", _Size, false, 345, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_tel_office, _Color, false, "", "L", "", _Size, false, 346, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_mobile, _Color, false, "", "L", "", _Size, false, 347, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_mobile_office, _Color, false, "", "L", "", _Size, false, 348, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_email_personal, _Color, false, "", "L", "", _Size, false, 349, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_email_office, _Color, false, "", "L", "", _Size, false, 350, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_bank_acno, _Color, false, "", "L", "", _Size, false, 351, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_bank_name, _Color, false, "", "L", "", _Size, false, 352, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_bank_branch, _Color, false, "", "L", "", _Size, false, 353, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_ifsc_code, _Color, false, "", "L", "", _Size, false, 354, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_pfno, _Color, false, "", "L", "", _Size, false, 355, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_esino, _Color, false, "", "L", "", _Size, false, 356, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_pan, _Color, false, "", "L", "", _Size, false, 357, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_adhar_no, _Color, false, "", "L", "", _Size, false, 358, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_uan_no, _Color, false, "", "L", "", _Size, false, 359, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_fuel_type, _Color, false, "", "L", "", _Size, false, 360, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_fuel_limit, _Color, false, "", "L", "", _Size, false, 361, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_bus_limit, _Color, false, "", "L", "", _Size, false, 362, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_train_limit, _Color, false, "", "L", "", _Size, false, 363, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_vehi_maint_limit, _Color, false, "", "L", "", _Size, false, 364, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_mobile_limit, _Color, false, "", "L", "", _Size, false, 365, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_datacard_limit, _Color, false, "", "L", "", _Size, false, 366, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_grade_name, _Color, false, "", "L", "", _Size, false, 367, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_department_name, _Color, false, "", "L", "", _Size, false, 368, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_designation_name, _Color, false, "", "L", "", _Size, false, 369, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_status_name, _Color, false, "", "L", "", _Size, false, 370, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_in_payroll ? "Y" : "N", _Color, false, "", "L", "", _Size, false, 371, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_marital_status, _Color, false, "", "L", "", _Size, false, 372, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_gender, _Color, false, "", "L", "", _Size, false, 373, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_comp_mediclaim, _Color, false, "", "L", "", _Size, false, 374, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_premium_amt, _Color, false, "", "L", "", _Size, false, 375, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_mediclaim_provider, _Color, false, "", "L", "", _Size, false, 376, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_remarks, _Color, false, "", "L", "", _Size, false, 377, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_do_birth, _Color, false, "", "L", "", _Size, false, 378, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_marrige_date, _Color, false, "", "L", "", _Size, false, 379, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_do_joining, _Color, false, "", "L", "", _Size, false, 380, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_do_confirmation, _Color, false, "", "L", "", _Size, false, 381, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.emp_do_relieve, _Color, false, "", "L", "", _Size, false, 382, "", true);

                }

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }
    }
}
