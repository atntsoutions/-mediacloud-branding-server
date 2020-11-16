using System;
using System.Data;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLHr
{
    public class LeaveMasterService : BL_Base
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

            Con_Oracle = new DBConnection();
            List<Leavem> mList = new List<Leavem>();
            Leavem mRow;

            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            report_folder = SearchData["report_folder"].ToString();
            int levmonth = Lib.Conv2Integer(SearchData["levmonth"].ToString());
            int levyear = Lib.Conv2Integer(SearchData["levyear"].ToString());
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where a.rec_company_code = '" + company_code + "'";
                sWhere += " and a.rec_branch_code = '" + branch_code + "'";
                sWhere += " and a.lev_year = " + levyear.ToString();
                sWhere += " and a.lev_month = 0 ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(b.emp_name) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  b.emp_no like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " ) ";
                }
               
                if (type == "NEW")
                {
                    sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM Leavem  a ";
                    sql += " left join empm b on a.lev_emp_id = b.emp_pkid ";
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
                sql += " select lev_pkid,lev_emp_id,lev_year,lev_month,lev_fin_year,lev_sl,lev_cl,lev_pl,lev_others,lev_pl_carry ";
                sql += " ,lev_tot_sl as lev_sl_tkn,lev_tot_cl as lev_cl_tkn,lev_tot_pl as lev_pl_tkn";
                sql += " ,(lev_sl-nvl(lev_tot_sl,0)) as lev_sl_bal,(lev_cl- nvl(lev_tot_cl,0))as lev_cl_bal, ((lev_pl + nvl(LEV_PL_CARRY,0)) - nvl(lev_tot_pl,0)) as lev_pl_bal ";
                sql += " ,emp_no,emp_name ";
                sql += " ,row_number() over(order by emp_no) rn ";
                sql += " from leavem a ";
                sql += " left join empm b on a.lev_emp_id = b.emp_pkid ";
                sql += sWhere;
                sql += ") a ";
                if (type != "EXCEL")
                    sql += " where rn between {startrow} and {endrow}";
                sql += " order by emp_no";

                sql = sql.Replace("{startrow}", startrow.ToString());
                sql = sql.Replace("{endrow}", endrow.ToString());

                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Leavem();
                    mRow.lev_pkid = Dr["lev_pkid"].ToString();
                    mRow.lev_emp_id = Dr["lev_emp_id"].ToString();
                    mRow.lev_emp_code = Dr["emp_no"].ToString();
                    mRow.lev_emp_name = Dr["emp_name"].ToString();
                    mRow.lev_fin_year= Lib.Conv2Integer(Dr["lev_fin_year"].ToString());
                    mRow.lev_year = Lib.Conv2Integer(Dr["lev_year"].ToString());
                    mRow.lev_month = Lib.Conv2Integer(Dr["lev_month"].ToString());
                    mRow.lev_sl = Lib.Conv2Decimal(Dr["lev_sl"].ToString());
                    mRow.lev_cl = Lib.Conv2Decimal(Dr["lev_cl"].ToString());
                    mRow.lev_pl = Lib.Conv2Decimal(Dr["lev_pl"].ToString());
                    mRow.lev_others = Lib.Conv2Decimal(Dr["lev_others"].ToString());
                    mRow.lev_pl_carry= Lib.Conv2Decimal(Dr["lev_pl_carry"].ToString());
                    mRow.lev_sl_tkn  = Lib.Conv2Decimal(Dr["lev_sl_tkn"].ToString());
                    mRow.lev_cl_tkn = Lib.Conv2Decimal(Dr["lev_cl_tkn"].ToString());
                    mRow.lev_pl_tkn = Lib.Conv2Decimal(Dr["lev_pl_tkn"].ToString());
                    mRow.lev_sl_bal = Lib.Conv2Decimal(Dr["lev_sl_bal"].ToString());
                    mRow.lev_cl_bal = Lib.Conv2Decimal(Dr["lev_cl_bal"].ToString());
                    mRow.lev_pl_bal = Lib.Conv2Decimal(Dr["lev_pl_bal"].ToString());

                    mList.Add(mRow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintLeaveReport(mList, branch_code);
                }
            }
            catch (Exception Ex)
            {
                if ( Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("page_count", page_count);
            RetData.Add("page_current", page_current);
            RetData.Add("page_rowcount", page_rowcount);
            RetData.Add("list", mList);

            return RetData;
        }

        private void PrintLeaveReport(List<Leavem> mList, string branch_code)
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            string Folderid = "";
            decimal tot = 0;
            try
            {

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                mSearchData.Add("branch_code", branch_code);

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["COMP_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                Folderid = Guid.NewGuid().ToString().ToUpper();
                File_Display_Name = branch_code + "_LEV.xls";
                File_Name = Lib.GetFileName(report_folder, Folderid, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 10;
                WS.Columns[2].Width = 256 * 35;
                WS.Columns[3].Width = 256 * 10;
                WS.Columns[4].Width = 256 * 10;
                WS.Columns[5].Width = 256 * 6;
                WS.Columns[6].Width = 256 * 6;
                WS.Columns[7].Width = 256 * 6;
                WS.Columns[8].Width = 256 * 10;
                WS.Columns[9].Width = 256 * 10;

                iRow = 0; iCol = 1;


                iRow++;
                _Size = 14;
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

                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);

                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "LEAVE STATUS REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);

                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "YEAR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OTHERS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                foreach (Leavem Rec in mList)
                {
                    iRow++;
                    iCol = 1;

                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_emp_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_emp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    str = (Rec.lev_fin_year.ToString() + "-" + (Rec.lev_fin_year + 1).ToString()).ToString();
                    Lib.WriteData(WS, iRow, iCol++, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "OPENING", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_pl, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_cl, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_sl, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_others, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    tot = Lib.Convert2Decimal(Rec.lev_pl.ToString()) + Lib.Convert2Decimal(Rec.lev_cl.ToString()) + Lib.Convert2Decimal(Rec.lev_sl.ToString()) + Lib.Convert2Decimal(Rec.lev_others.ToString());
                    Lib.WriteData(WS, iRow, iCol++, tot, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    iRow++;
                    iCol = 1;

                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "TAKEN", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_pl_tkn, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_cl_tkn, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_sl_tkn, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++,"", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    tot = Lib.Convert2Decimal(Rec.lev_pl_tkn.ToString()) + Lib.Convert2Decimal(Rec.lev_cl_tkn.ToString()) + Lib.Convert2Decimal(Rec.lev_sl_tkn.ToString()) ;
                    Lib.WriteData(WS, iRow, iCol++, tot, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    iRow++;
                    iCol = 1;

                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_pl_bal, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_cl_bal, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.lev_sl_bal, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++,"", _Color, false, "", "L", "", _Size, false, 325, "", true);
                    tot = Lib.Convert2Decimal(Rec.lev_pl_bal.ToString()) + Lib.Convert2Decimal(Rec.lev_cl_bal.ToString()) + Lib.Convert2Decimal(Rec.lev_sl_bal.ToString()) ;
                    Lib.WriteData(WS, iRow, iCol++, tot, _Color, false, "", "L", "", _Size, false, 325, "", true);

                }
                iRow++;

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

        public Dictionary<string, object>  GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Leavem mRow =new Leavem();
            string id = SearchData["pkid"].ToString();
            try
            {
                DataTable Dt_Rec = new DataTable();
     
                sql = " select lev_pkid,lev_emp_id,lev_year,lev_month, ";
                sql += " lev_sl,lev_cl,lev_pl,lev_others, ";
                sql += " lev_pl_carry,lev_fin_year, ";
                sql += " b.emp_no as lev_emp_code, b.emp_name as  lev_emp_name,lev_edit_code ";
                sql += " from leavem a  ";
                sql += " left join empm b on a.lev_emp_id = b.emp_pkid ";
                sql += " where  a.lev_pkid ='" + id + "'";

                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Leavem();
                    mRow.lev_pkid = Dr["lev_pkid"].ToString();
                    mRow.lev_emp_id = Dr["lev_emp_id"].ToString();
                    mRow.lev_emp_code = Dr["lev_emp_code"].ToString();
                    mRow.lev_emp_name = Dr["lev_emp_name"].ToString();
                    mRow.lev_year = Lib.Conv2Integer(Dr["lev_year"].ToString());
                    mRow.lev_month = Lib.Conv2Integer(Dr["lev_month"].ToString());
                    mRow.lev_sl = Lib.Conv2Decimal(Dr["lev_sl"].ToString());
                    mRow.lev_cl = Lib.Conv2Decimal(Dr["lev_cl"].ToString());
                    mRow.lev_pl = Lib.Conv2Decimal(Dr["lev_pl"].ToString());
                    mRow.lev_others = Lib.Conv2Decimal(Dr["lev_others"].ToString());
                    mRow.lev_pl_carry = Lib.Conv2Decimal(Dr["lev_pl_carry"].ToString());
                    mRow.lev_fin_year = Lib.Conv2Integer(Dr["lev_fin_year"].ToString());
                    mRow.lev_edit_code = Dr["lev_edit_code"].ToString();
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

        public string AllValid(Leavem Record)
        { 
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {
                if (Record.lev_emp_id.Trim().Length <= 0)
                    Lib.AddError(ref str, " | Employee Cannot Be Empty");

                sql = " select lev_pkid from leavem a";
                sql += " where a.rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and a.rec_branch_code= '" + Record._globalvariables.branch_code + "'";
                sql += " and a.lev_emp_id = '" + Record.lev_emp_id + "'";
                sql += " and lev_month = 0" ;
                sql += " and lev_fin_year = " + Record._globalvariables.year_code;
                sql += " and lev_edit_code is null";
                if (Con_Oracle.IsRowExists(sql))
                    Lib.AddError(ref str, " Details Closed, Can't Edit ");

                //if (Record.sal_code.Trim().Length > 0)
                //{

                //    sql = "select sal_pkid from (";
                //    sql += "select sal_pkid  from salaryheadm a where (a.sal_code = '{CODE}')  ";
                //    sql += ") a where sal_pkid <> '{PKID}'";

                //    sql = sql.Replace("{CODE}", Record.sal_code);
                //    sql = sql.Replace("{PKID}", Record.sal_pkid);

                //    if (Con_Oracle.IsRowExists(sql))
                //        Lib.AddError(ref str, " | Code Exists");
                //}

                //if (Record.sal_desc.Trim().Length > 0)
                //{

                //    sql = "select sal_pkid from (";
                //    sql += "select sal_pkid  from salaryheadm a where (a.sal_desc = '{NAME}')  ";
                //    sql += ") a where sal_pkid <> '{PKID}'";

                //    sql = sql.Replace("{NAME}", Record.sal_desc);
                //    sql = sql.Replace("{PKID}", Record.sal_pkid);


                //    if (Con_Oracle.IsRowExists(sql))

                //        Lib.AddError(ref str, " | Description Exists");
                //}

            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }
        
        public Dictionary<string, object> Save(Leavem Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();
                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("Leavem", Record.rec_mode, "lev_pkid", Record.lev_pkid);
                Rec.InsertString("lev_emp_id", Record.lev_emp_id);
                Rec.InsertNumeric("lev_sl", Record.lev_sl.ToString());
                Rec.InsertNumeric("lev_cl", Record.lev_cl.ToString());
                Rec.InsertNumeric("lev_pl", Record.lev_pl.ToString());
                Rec.InsertNumeric("lev_others", Record.lev_others.ToString());
                Rec.InsertNumeric("lev_pl_carry", Record.lev_pl_carry.ToString());
                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("lev_edit_code", "{S}");
                    Rec.InsertNumeric("lev_year", Record._globalvariables.year_code);
                    Rec.InsertNumeric("lev_fin_year", Record._globalvariables.year_code);
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

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //parameter.Add("comp_code", comp_code);
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }

        public IDictionary<string, object> Generate(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            string ErrorMessage = "";
            string type = SearchData["type"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string branch_code = SearchData["branch_code"].ToString();
            string company_code = SearchData["company_code"].ToString();
            string year_code = SearchData["year_code"].ToString();
            string user_code = SearchData["user_code"].ToString();
            int FinYear = Lib.Conv2Integer(SearchData["levyear"].ToString());
            string Lev_Emp_Ids = "";
            string ErrMsg = "";
            try
            {

                if ((ErrorMessage = GenerateValid(FinYear, 0, branch_code, company_code)) != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }


                sql = "  select lev_emp_id from leavem a";
                sql += " where a.rec_company_code ='" + company_code + "'";
                sql += " and a.rec_branch_code= '" + branch_code + "'";
                sql += " and lev_fin_year = " + year_code;
                sql += " and lev_month = 0";

                DataTable Dt_Lev = new DataTable();
                Dt_Lev = Con_Oracle.ExecuteQuery(sql);
                Lev_Emp_Ids = "";
                foreach (DataRow dr in Dt_Lev.Rows)
                {
                    if (Lev_Emp_Ids != "")
                        Lev_Emp_Ids += ",";
                    Lev_Emp_Ids += dr["lev_emp_id"].ToString();
                }

                sql = "  select emp_pkid from empm a ";
                sql += " inner join param grd on a.emp_grade_id = grd.param_pkid";
                sql += " where a.rec_company_code ='" + company_code + "'";
                sql += " and a.rec_branch_code = '" + branch_code + "'";
                sql += " and grd.param_name not in('MANAGING DIRECTOR','DIRECTOR') ";
                sql += " and emp_in_payroll = 'Y' ";
                if (Lev_Emp_Ids != "")//during re-generation
                {
                    Lev_Emp_Ids = Lev_Emp_Ids.Replace(",", "','");
                    sql += " and emp_pkid not in ('" + Lev_Emp_Ids + "')";
                }

                DataTable Dt_Emp = new DataTable();
                Dt_Emp = Con_Oracle.ExecuteQuery(sql);
                if (Dt_Emp.Rows.Count <= 0)
                    ErrMsg = "No New Records to Generate";
                foreach (DataRow dr in Dt_Emp.Rows)
                {
                    InsertLeavem(dr["emp_pkid"].ToString(), 0, FinYear, company_code, branch_code, user_code);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("error", ErrMsg);
            return RetData;
        }
        public string GenerateValid(int syear, int smonth, string brcode, string compcode)
        {
            string str = "";
            DateTime tdate = DateTime.Now;
            try
            {
                int tempYear = 0;
                DataTable Dt_TEMP;
    
                sql = " select distinct(lev_year) ";
                sql += " from leavem a ";
                sql += " inner join empm b on a.lev_emp_id = b.emp_pkid";
                sql += " where a.rec_company_code = '" + compcode + "'";
                sql += " and a.rec_branch_code = '" + brcode + "'";
                sql += " and lev_month = 0 ";
                sql += " and b.emp_in_payroll = 'Y'";
                sql += " order by lev_year desc ";

                Dt_TEMP = new DataTable();
                Dt_TEMP = Con_Oracle.ExecuteQuery(sql);
                if (Dt_TEMP.Rows.Count > 0)
                {
                    if (syear > 0)
                    {
                        tempYear = Lib.Conv2Integer(Dt_TEMP.Rows[0]["LEV_YEAR"].ToString());
                        if (syear > tempYear + 1)
                        {
                            Lib.AddError(ref str, " | Previous not Generated");
                        }
                        tempYear = Lib.Conv2Integer(Dt_TEMP.Rows[Dt_TEMP.Rows.Count - 1]["LEV_YEAR"].ToString());
                        if (syear < tempYear)
                        {
                            Lib.AddError(ref str, " | Invalid Generate");
                        }
                    }
                }

                sql = " select lev_pkid from leavem a";
                sql += " inner join empm b on a.lev_emp_id = b.emp_pkid";
                sql += " where a.rec_company_code = '" + compcode + "'";
                sql += " and a.rec_branch_code = '" + brcode + "'";
                sql += " and lev_month = 0 ";
                sql += " and lev_year = " + syear;
                sql += " and lev_edit_code is null";
                sql += " and b.emp_in_payroll = 'Y'";

                Dt_TEMP = new DataTable();
                Dt_TEMP = Con_Oracle.ExecuteQuery(sql);
                if (Dt_TEMP.Rows.Count > 0)
                {
                    Lib.AddError(ref str, " | Year(" + syear.ToString() + ") Closed");
                }
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }
        private void InsertLeavem(string Emp_Id, int lev_month, int lev_Year, string company_code, string branch_code, string UserCode)
        {
            string SQL = "";
            DataTable Dt_LeaveData;
            decimal CarryPL = 0;

            SQL = " select lev_pkid from leavem a";
            SQL += " where a.rec_company_code ='" + company_code + "'";
            SQL += " and a.rec_branch_code = '" + branch_code + "'";
            SQL += " and lev_emp_id = '" + Emp_Id + "'";
            SQL += " and lev_month = 0 ";
            SQL += " and lev_fin_year = " + lev_Year.ToString();
            Dt_LeaveData = new DataTable();
            Dt_LeaveData = Con_Oracle.ExecuteQuery(SQL);
            if (Dt_LeaveData.Rows.Count <= 0)
            {
                DataTable Dt_PreYear;
                int pYear = lev_Year - 1;
                SQL = "";
                SQL = " select lev_pl,lev_pl_carry,lev_tot_pl as tot_pl_taken from leavem a ";
                SQL += " where a.rec_company_code ='" + company_code + "'";
                SQL += " and a.rec_branch_code = '" + branch_code + "'";
                SQL += " and lev_emp_id = '" + Emp_Id + "'";
                SQL += " and lev_month = 0 ";
                SQL += " and lev_fin_year = " + pYear.ToString();
                Dt_PreYear = new DataTable();
                Dt_PreYear = Con_Oracle.ExecuteQuery(SQL);
                if (Dt_PreYear.Rows.Count > 0)
                {
                    CarryPL = Lib.Convert2Decimal(Dt_PreYear.Rows[0]["lev_pl_carry"].ToString());//Previous yr PL Carry
                    CarryPL += Lib.Convert2Decimal(Dt_PreYear.Rows[0]["lev_pl"].ToString());//Previous yr PL
                    CarryPL -= Lib.Convert2Decimal(Dt_PreYear.Rows[0]["tot_pl_taken"].ToString());
                    if (CarryPL > 40)
                        CarryPL = 40; //since carryforward + curent pl <=60; ie 40 +20 <=60
                }

                SQL = "  Insert into LEAVEM ";
                SQL += " (LEV_PKID,LEV_EMP_ID";
                SQL += " ,LEV_YEAR,LEV_FIN_YEAR";
                SQL += " ,LEV_MONTH,LEV_PL,LEV_SL,LEV_CL";
                SQL += " ,LEV_PL_CARRY";
                SQL += " ,REC_COMPANY_CODE,REC_BRANCH_CODE,REC_CREATED_BY,REC_CREATED_DATE,LEV_EDIT_CODE)";
                SQL += " values ('" + Guid.NewGuid().ToString().ToUpper() + "','" + Emp_Id + "'";
                SQL += "," + lev_Year.ToString() + "," + lev_Year.ToString();
                SQL += ",0,20,8,8";
                SQL += "," + CarryPL;
                SQL += ",'" + company_code + "'";
                SQL += ",'" + branch_code + "'";
                SQL += ",'" + UserCode + "'";
                SQL += ",(SYSDATE),'{S}')";
                try
                {
                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(SQL);
                    Con_Oracle.CommitTransaction();
                }
                catch (Exception ex)
                {
                    if (Con_Oracle != null)
                    {
                        Con_Oracle.RollbackTransaction();
                        Con_Oracle.CloseConnection();
                    }
                    throw ex;
                }
            }
        }

    }
}
