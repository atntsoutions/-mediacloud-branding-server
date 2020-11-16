using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;

using DataBase_Oracle.Connections;
using XL.XSheet;
using System.Drawing;

namespace BLHr
{
    public class TaxplanDetService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;
       
        string report_folder = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string branch_name = "";
        string name = "";
        string year_name = "";
        int iRow = 0;
        int iCol = 0;
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            string user_id = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<TaxPlanm> mList = new List<TaxPlanm>();
            TaxPlanm mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string company_code = SearchData["company_code"].ToString();
            string branch_code = SearchData["branch_code"].ToString();

            string year_code = SearchData["year_code"].ToString();
            string user_pkid = SearchData["user_pkid"].ToString();
            string user_code = SearchData["user_code"].ToString();

            string is_admin = SearchData["is_admin"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            PKID = SearchData["file_pkid"].ToString();
            report_folder = SearchData["report_folder"].ToString();
            branch_name = SearchData["branch_name"].ToString();
            year_name = SearchData["year_name"].ToString();
           

            try
            {
                sWhere = " where  a.rec_company_code = '{COMPCODE}'";
                sWhere += " and a.tpm_year =  {FYEAR} ";
                if (is_admin != "Y")
                    sWhere += " and a.rec_created_by ='{USERCODE}'";

                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(b.user_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " )";
                }
                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{FYEAR}", year_code);
                sWhere = sWhere.Replace("{USERCODE}", user_code);

                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM taxplanm  a ";
                    sql += " left join userm b on a.tpm_user_id = b.user_pkid ";
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
                sql += " select tpm_pkid, comp_name, tpm_user_id, b.user_name as tpm_user_name,tpm_year,a.rec_created_by,a.rec_created_date, ";
                sql += " row_number() over(order by b.user_name) rn ";
                sql += " from taxplanm a ";
                sql += " left join userm b on a.tpm_user_id = b.user_pkid ";
                sql += " left join companym c on b.user_branch_id = c.comp_pkid";
                sql += sWhere;
                sql += ") a where rn between {startrow} and {endrow} ";
                sql += " order by tpm_user_name";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new TaxPlanm();
                    mRow.tpm_pkid = Dr["tpm_pkid"].ToString();
                    mRow.tpm_year = Lib.Conv2Integer(Dr["tpm_year"].ToString());
                    mRow.tpm_branch = Dr["comp_name"].ToString();
                    mRow.tpm_user_name = Dr["tpm_user_name"].ToString();
                    mRow.tpm_user_id = Dr["tpm_user_id"].ToString();
                    mRow.rec_created_by= Dr["rec_created_by"].ToString();
                    mRow.rec_created_date = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mList.Add(mRow);

                    if (user_id != "")
                        user_id += ",";
                    user_id += Dr["tpm_user_id"].ToString();
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                    {
                        PrintTaxplanDetReport(user_id, company_code, year_code);

                    }
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
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);

            return RetData;
        }

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            TaxPlanm mRow = new TaxPlanm();

            string id = SearchData["pkid"].ToString();
            string user_pkid = SearchData["user_pkid"].ToString();
            string type = SearchData["type"].ToString();
            string company_code = SearchData["comp_code"].ToString();
            string year = SearchData["year"].ToString();
            string action = SearchData["action"].ToString();

            PKID = SearchData["file_pkid"].ToString();
            report_folder = SearchData["report_folder"].ToString();
            branch_name = SearchData["branch_name"].ToString();
            name = SearchData["name"].ToString();
            year_name = SearchData["year_name"].ToString();

            try
            {
                Con_Oracle = new DBConnection();

                sql = "select tpm_pkid,a.rec_created_by from taxplanm a";
                sql += " where a.rec_company_code = '" + SearchData["comp_code"].ToString() + "'";
                sql += " and a.tpm_year = " + SearchData["year"].ToString() + "";
                sql += " and a.tpm_user_id = '" + user_pkid + "'";
                DataTable Dt_temp = Con_Oracle.ExecuteQuery(sql);
                if (Dt_temp.Rows.Count > 0)
                {
                    mRow.rec_mode = "EDIT";
                    mRow.tpm_pkid = Dt_temp.Rows[0]["tpm_pkid"].ToString();
                    mRow.rec_created_by= Dt_temp.Rows[0]["rec_created_by"].ToString();
                }
                else
                {
                    mRow.rec_mode = "ADD";
                    mRow.tpm_pkid = id;
                }
                mRow.tpm_user_id = user_pkid;

                sql = "  select tp_pkid ,tp_year,tp_group_ctr,tp_ctr,tp_desc,tp_limit,tp_editable,tp_bold,";
                sql += " tpd_amt_invested, tpd_amt_tot,tpd_amt_before_dec31,tpd_amt_after_dec31 ";
                sql += " from taxplan a ";
                sql += " left join taxpland b on a.tp_pkid = b.tpd_plan_id and b.tpd_user_id ='" + user_pkid + "'";
                sql += " where a.rec_company_code = '" + SearchData["comp_code"].ToString() + "'";
                sql += " and a.tp_year = " + SearchData["year"].ToString() + "";
                sql += " order by a.tp_group_ctr,a.tp_ctr ";

                List<TaxPland> mList = new List<TaxPland>();
                TaxPland Row;
                DataTable Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    Row = new TaxPland();
                    Row.tpd_plan_id = Dr["tp_pkid"].ToString();
                    Row.tpd_year = Lib.Conv2Integer(Dr["tp_year"].ToString());
                    Row.tpd_plan_group_ctr = Lib.Conv2Integer(Dr["tp_group_ctr"].ToString());
                    Row.tpd_plan_ctr = Lib.Conv2Integer(Dr["tp_ctr"].ToString());
                    Row.tpd_plan_desc = Dr["tp_desc"].ToString();
                    Row.tpd_plan_limit = Lib.Conv2Decimal(Dr["tp_limit"].ToString());
                    Row.tpd_plan_editable = Dr["tp_editable"].ToString() == "Y" ? true : false;
                    Row.tpd_plan_bold = Dr["tp_bold"].ToString() == "Y" ? true : false;
                    Row.tpd_amt_invested = Lib.Conv2Decimal(Dr["tpd_amt_invested"].ToString());
                    Row.tpd_amt_before_dec31 = Lib.Conv2Decimal(Dr["tpd_amt_before_dec31"].ToString());
                    Row.tpd_amt_after_dec31 = Lib.Conv2Decimal(Dr["tpd_amt_after_dec31"].ToString());
                    Row.tpd_amt_tot = Lib.Conv2Decimal(Dr["tpd_amt_tot"].ToString());

                    mList.Add(Row);
                }
                mRow.DetailList = mList;

                if (type == "EXCEL")
                {
                    if (mList != null)
                    {
                        PrintTaxplanDetReport(user_pkid, company_code, year);
                    }
                }
                Dt_Rec.Rows.Clear();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("record", mRow);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }


        public string AllValid(TaxPlanm Record)
        {
            string str = "";
            //DateTime tdate = DateTime.Now;
            //try
            //{
            //    if (Record.emp_name.Trim().Length <= 0)
            //        Lib.AddError(ref str, "Name Cannot Be Empty");

            //    if (Record.emp_no.Trim().Length > 0)
            //    {

            //        sql = "select emp_pkid from (";
            //        sql += "select emp_pkid  from empm a where (a.emp_no = '{CODE}')  ";
            //        sql += ") a where emp_pkid <> '{PKID}'";

            //        sql = sql.Replace("{CODE}", Record.emp_no);
            //        sql = sql.Replace("{PKID}", Record.emp_pkid);

            //        if (Con_Oracle.IsRowExists(sql))
            //            Lib.AddError(ref str, "Code Exists");
            //    }

            //    if (Record.emp_name.Trim().Length > 0)
            //    {

            //        sql = "select emp_pkid from (";
            //        sql += "select emp_pkid  from empm a where (a.emp_name = '{NAME}')  ";
            //        sql += ") a where emp_pkid <> '{PKID}'";

            //        sql = sql.Replace("{NAME}", Record.emp_name);
            //        sql = sql.Replace("{PKID}", Record.emp_pkid);


            //        if (Con_Oracle.IsRowExists(sql))

            //           Lib.AddError(ref str, "Name Exists");
            //    }

            //    if (Record.emp_do_joining.Trim().Length > 0 && Record.emp_do_birth.Trim().Length > 0)
            //    {
            //        DateTime dob = DateTime.Parse(Record.emp_do_birth);
            //        DateTime doj = DateTime.Parse(Record.emp_do_joining);

            //        if (dob > doj)
            //            Lib.AddError(ref str, "  Joining Date should be greater than DOB ");
            //    }

            //    if (Record.emp_do_confirmation.Trim().Length > 0 && Record.emp_do_birth.Trim().Length > 0)
            //    {

            //        DateTime dob = DateTime.Parse(Record.emp_do_birth);
            //        DateTime doc = DateTime.Parse(Record.emp_do_confirmation);


            //        if (dob > doc)
            //            Lib.AddError(ref str, " Confirmation Date should be greater than DOB");

            //    }

            //    if (Record.emp_do_relieve.Trim().Length > 0 && Record.emp_do_birth.Trim().Length > 0)
            //    {

            //        DateTime dob = DateTime.Parse(Record.emp_do_birth);
            //        DateTime dor = DateTime.Parse(Record.emp_do_relieve);

            //        if (dob > dor)
            //            Lib.AddError(ref str, " Relieve Date should be greater than DOB");

            //    }

            //    if (Record.emp_branch_id != Record._globalvariables.branch_pkid)
            //    {
            //        Lib.AddError(ref str, " selected branch and login branch are mismatch");
            //    }

            //}
            //catch (Exception Ex)
            //{
            //    str = Ex.Message.ToString();
            //}
            return str;
        }


        public Dictionary<string, object> Save(TaxPlanm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                sql = "select tpm_pkid from taxplanm a ";
                sql += " where a.rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and a.tpm_year = " + Record._globalvariables.year_code + "";
                sql += " and a.tpm_user_id = '" + Record.tpm_user_id + "'";

                DataTable Dt_temp = Con_Oracle.ExecuteQuery(sql);
                Record.rec_mode = "ADD";
                if (Dt_temp.Rows.Count > 0)
                {
                    Record.tpm_pkid = Dt_temp.Rows[0]["tpm_pkid"].ToString();
                    Record.rec_mode = "EDIT";
                }

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("taxplanm", Record.rec_mode, "tpm_pkid", Record.tpm_pkid);
                Rec.InsertString("tpm_user_id", Record.tpm_user_id);
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertNumeric("tpm_year", Record._globalvariables.year_code);
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
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

                sql = "delete from taxpland where tpd_parent_id ='" + Record.tpm_pkid + "'";
                Con_Oracle.ExecuteNonQuery(sql);

                foreach (TaxPland Row in Record.DetailList)
                {
                    if (Lib.Conv2Decimal(Row.tpd_amt_tot.ToString()) != 0 || Lib.Convert2Decimal(Row.tpd_amt_after_dec31.ToString()) != 0)
                    {
                        
                        Row.tpd_pkid = Guid.NewGuid().ToString().ToUpper();
                        Rec = new DBRecord();
                        Rec.CreateRow("taxpland", "ADD", "tpd_pkid", Row.tpd_pkid);
                        Rec.InsertString("tpd_parent_id", Record.tpm_pkid);
                        Rec.InsertNumeric("tpd_year", Record._globalvariables.year_code);
                        Rec.InsertString("tpd_user_id", Record.tpm_user_id);

                        Rec.InsertString("tpd_plan_id", Row.tpd_plan_id);
                        Rec.InsertNumeric("tpd_amt_invested", Row.tpd_amt_invested.ToString());
                        Rec.InsertNumeric("tpd_amt_before_dec31", Row.tpd_amt_before_dec31.ToString());
                        Rec.InsertNumeric("tpd_amt_after_dec31", Row.tpd_amt_after_dec31.ToString());
                        Rec.InsertNumeric("tpd_amt_tot", Row.tpd_amt_tot.ToString());

                        sql = Rec.UpdateRow();
                        Con_Oracle.ExecuteNonQuery(sql);
                    }
                }

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

            string year = "";
            if (SearchData.ContainsKey("year"))
                year = SearchData["year"].ToString();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "taxplanlist");
            //parameter.Add("year", year);
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("taxplanlist", lovservice.Lov(parameter)["taxplanlist"]);

            return RetData;

        }




        private void PrintTaxplanDetReport(string user_id, string company_code, string year)
        {
            
            try
            {
                string[] userid_array = null;
                
                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                
                File_Display_Name = "TaxInvestmentsReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);
                
                WB = new ExcelFile();

                if(user_id.Contains(","))
                {
                    userid_array = user_id.Split(',');
                    for(int i =0;i < userid_array.Length;i++)
                    {
                        PrintTaxplan(userid_array[i], company_code, year);
                    }
                }
                else
                {
                    PrintTaxplan(user_id, company_code, year);
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

        private void PrintTaxplan( string user_id,string company_code,string year)
        {
            string str = "";
            
            string user_name = "";
            string sName = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            iRow = 0;
            iCol = 0;
            int i = 0;

            Con_Oracle = new DBConnection();
            DataTable Dt_taxplan = new DataTable();

            sql = "  select tp_pkid ,tp_year,tp_group_ctr,tp_ctr,tp_desc,tp_limit,tp_editable,tp_bold,";
            sql += " tpd_amt_invested, tpd_amt_tot,tpd_amt_before_dec31,tpd_amt_after_dec31";
            sql += " from taxplan a ";
            sql += " left join taxpland b on a.tp_pkid = b.tpd_plan_id and b.tpd_user_id ='" + user_id + "'";
            sql += " where a.rec_company_code = '" + company_code + "'";
            sql += " and a.tp_year = " + year + "";
            sql += " order by a.tp_group_ctr,a.tp_ctr ";

            Dt_taxplan = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            Con_Oracle = new DBConnection();
            DataTable Dt_username = new DataTable();

            sql = "";
            sql = " select user_name,user_code, comp_name from userm ";
            sql += " left join companym c on user_branch_id = c.comp_pkid";
            sql += " where user_pkid = '" + user_id + "'";

            Dt_username = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();

            if (Dt_username.Rows.Count > 0)
            {
                user_name = Dt_username.Rows[0]["user_name"].ToString();
                sName = Dt_username.Rows[0]["user_code"].ToString();
                branch_name = Dt_username.Rows[0]["comp_name"].ToString();
            }
            
            WB.Worksheets.Add(sName);
            WS = WB.Worksheets[sName];

            // WS.ViewOptions.ShowGridLines = false;
            WS.PrintOptions.FitWorksheetWidthToPages = 1;


            WS.Columns[0].Width = 256 * 2;
            WS.Columns[1].Width = 256 * 9;
            WS.Columns[2].Width = 256 * 85;
            WS.Columns[3].Width = 256 * 15;
            WS.Columns[4].Width = 256 * 15;
            WS.Columns[5].Width = 256 * 15;
            WS.Columns[6].Width = 256 * 15;
            WS.Columns[7].Width = 256 * 15;
            WS.Columns[8].Width = 256 * 15;


            iRow = 0; iCol = 1;

            _Size = 14;
            iRow++;
            
            Lib.WriteData(WS, iRow, 1, "INCOME TAX CALCULATION ", _Color, true, "", "L", "", 15, false, 325, "", true);
            iRow++;
            iRow++;
            _Size = 10;
            iCol = 1;


            Lib.WriteData(WS, iRow, 1, "BRANCH", _Color, true, "", "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, branch_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow, 1, "NAME", _Color, true, "", "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, user_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
            iRow++;
            Lib.WriteData(WS, iRow, 1, "YEAR", _Color, true, "", "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, 2, year_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
            iRow++;

            iRow++;
            iRow++;
            Lib.WriteData(WS, iRow, iCol++, "SL-NO", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "PARTICULARS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "LIMIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TOTAL AMOUNT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "TILL 15-SEP", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "BEFORE 31-DEC", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
            Lib.WriteData(WS, iRow, iCol++, "AFTER 31-DEC", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

            string GRP_NO = "";
            foreach (DataRow Dr in Dt_taxplan.Rows)
            {
                iRow++;
                iCol = 1;
                i++;
                _Bold = false;

                if (Dr["tp_bold"].ToString() == "Y")
                {
                    _Bold = true;
                }
                if (GRP_NO != Dr["tp_group_ctr"].ToString())
                {
                    Lib.WriteData(WS, iRow, iCol++, Dr["tp_group_ctr"], _Color, false, "", "C", "", _Size, false, 325, "#;(#);#", true);
                }
                else
                {
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "C", "", _Size, false, 325, "", true);
                }

                Lib.WriteData(WS, iRow, iCol++, Dr["tp_desc"], _Color, _Bold, "", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, Dr["tp_limit"], _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, Dr["tpd_amt_tot"], _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, Dr["tpd_amt_invested"], _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, Dr["tpd_amt_before_dec31"], _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                Lib.WriteData(WS, iRow, iCol++, Dr["tpd_amt_after_dec31"], _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);

                GRP_NO = Dr["tp_group_ctr"].ToString();
            }



        }

    }
}
