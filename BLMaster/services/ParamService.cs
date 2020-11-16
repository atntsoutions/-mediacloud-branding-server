using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase.Connections;

using XL.XSheet;
using System.IO;
namespace BLMaster
{
    public class ParamService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;
        int iRow = 0;
        int iCol = 0;
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string report_folder = "";
        DataTable Dt_param = null;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            string sWhere = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            DataTable Dt_List = null;
            Con_Oracle = new DBConnection();
            List<Param> mList = new List<Param>();
            Param mRow;

            string company_code = SearchData["company_code"].ToString();
            string type = SearchData["type"].ToString();
            string param_type = SearchData["rowtype"].ToString();
            string searchstring = SearchData["searchstring"].ToString().ToUpper();
            string sortby = SearchData["sortby"].ToString().ToUpper();
            report_folder = SearchData["report_folder"].ToString();

            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;

            try
            {
                sWhere = " where a.rec_company_code = '{COMPANY_CODE}' and param_type = '{PARAM_TYPE}' ";
                if (searchstring != "")
                {
                    sWhere += " and (";
                    sWhere += "  upper(a.param_code) like '%" + searchstring.ToUpper() + "%'";
                    sWhere += " or ";
                    sWhere += "  upper(a.param_name) like '%" + searchstring.ToUpper() + "%'";
                    if (param_type == "COUNTRY")
                    {
                        sWhere += " or ";//Region
                        sWhere += "  upper(a.param_id1) like '%" + searchstring.ToUpper() + "%'";
                    }
                    sWhere += " )";
                }

                sWhere = sWhere.Replace("{COMPANY_CODE}", company_code);
                sWhere = sWhere.Replace("{PARAM_TYPE}", param_type);


                if (type == "EXCEL")
                {
                    sql = "";
                    sql += "  select  param_pkid, param_type, param_code, param_name, param_rate,param_email,param_id1,param_id3 ";
                    sql += "  from param a " + sWhere;
                    if (sortby == "CODE")
                        sql += " order by param_code";
                    else
                        sql += " order by param_name";

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                }
                else
                {

                    if (type == "NEW")
                    {
                        sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM param  a ";
                        if (Con_Oracle.DB == "SQL")
                            sql = "SELECT count(*) as total, ceiling(COUNT(*) / cast(" + page_rows.ToString() + " as decimal) ) page_total  FROM param  a ";

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


                   
                    sql = "";
                    sql += " select * from ( ";
                    sql += "  select  param_pkid, param_type, param_code, param_name, param_rate,param_email,param_id1,param_id3, rec_locked,  row_number() over(order by param_name) rn ";
                    sql += "  from param a " + sWhere;
                    sql += ") a where rn between {startrow} and {endrow}";
                    if (sortby == "CODE")
                        sql += " order by param_code";
                    else
                        sql += " order by param_name";

                    sql = sql.Replace("{startrow}", startrow.ToString());
                    sql = sql.Replace("{endrow}", endrow.ToString());

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                }

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Param();
                    mRow.param_pkid = Dr["param_pkid"].ToString();
                    mRow.param_type = Dr["param_type"].ToString();
                    mRow.param_code = Dr["param_code"].ToString();
                    mRow.param_name = Dr["param_name"].ToString();
                    mRow.param_email = Dr["param_email"].ToString();
                    mRow.param_rate = Lib.Conv2Decimal(Dr["param_rate"].ToString());
                    mRow.param_id1 = Dr["param_id1"].ToString();
                    mRow.param_id3 = Dr["param_id3"].ToString();
                    mRow.rec_locked = (Dr["rec_locked"].ToString() == "Y" ? true : false);
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
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            return RetData;
        }

        public Dictionary<string, object> GetRecord(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Param mRow = new Param();

            string id = SearchData["pkid"].ToString();

            string comp_code  = SearchData["comp_code"].ToString();

            string ServerImageUrl = Lib.GetSeverImageURL(comp_code);

            try
            {
                DataTable Dt_Rec = new DataTable();

                sql = "select  param_pkid, param_type, param_code, param_name,param_id1, param_id2,param_id3,param_id4, param_email, param_rate, a.rec_locked,param_slno, ";
                sql += "param_file_name ";
                sql += " from param a  ";
                sql += " where  a.param_pkid ='" + id + "'";


                Con_Oracle = new DBConnection();
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                foreach (DataRow Dr in Dt_Rec.Rows)
                {
                    mRow = new Param();
                    mRow.param_pkid = Dr["param_pkid"].ToString();
                    mRow.param_type = Dr["param_type"].ToString();
                    mRow.param_code = Dr["param_code"].ToString();
                    mRow.param_name = Dr["param_name"].ToString();
                    mRow.param_id1 = Dr["param_id1"].ToString();
                    mRow.param_id2 = Dr["param_id2"].ToString();
                    mRow.param_id3 = Dr["param_id3"].ToString();
                    mRow.param_id4 = Dr["param_id4"].ToString();
                    mRow.param_email = Dr["param_email"].ToString();


                    mRow.param_slno = Lib.Conv2Integer(Dr["param_slno"].ToString());

                    mRow.param_file_name = Dr["param_file_name"].ToString();
                    //mRow.param_file = Dr["param_file_name"].ToString();
                    mRow.param_file_uploaded = false;
                    if (Dr["param_file_name"].ToString().Trim().Length >0 )
                        mRow.param_file_uploaded = true;

                    mRow.param_server_folder = Lib.getPath(ServerImageUrl, comp_code, "PARAM", mRow.param_slno.ToString(), false);

                    if (Dr["rec_locked"].ToString() == "Y")
                        mRow.rec_locked = true;
                    else
                        mRow.rec_locked = false;

                    mRow.param_rate = Lib.Conv2Decimal(Dr["param_rate"].ToString());
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

        public string AllValid(Param Record)
        {
            string str = "";
            try
            {
                sql = "select param_pkid from (";
                sql += "select param_pkid  from param a where a.rec_company_code = '{COMPANY_CODE}'  ";
                //sql += " and param_type= '{TYPE}' and (a.param_code = '{CODE}' or a.param_name = '{NAME}')  ";
                sql += " and param_type= '{TYPE}' and (upper(a.param_name) = '{NAME}')  ";
                sql += ") a where param_pkid <> '{PKID}'";

                sql = sql.Replace("{COMPANY_CODE}", Record._globalvariables.comp_code);
                sql = sql.Replace("{TYPE}", Record.param_type);
                sql = sql.Replace("{CODE}", Record.param_code);
                sql = sql.Replace("{NAME}", Record.param_name.ToString().ToUpper());
                sql = sql.Replace("{PKID}", Record.param_pkid);

                if (Con_Oracle.IsRowExists(sql))
                    str = "Code/Name Exists";
            }
            catch (Exception Ex)
            {
                str = Ex.Message.ToString();
            }
            return str;
        }

        public Dictionary<string, object> Save(Param Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            string sql1 = "";

            object iSlno = 0;

            try
            {
                Con_Oracle = new DBConnection();

                /*
                if (Record.param_code.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Code Cannot Be Empty");
                */

                if (Record.param_name.Trim().Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Name Cannot Be Empty");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                if ((ErrorMessage = AllValid(Record)) != "")
                    throw new Exception(ErrorMessage);


                DBRecord Rec = new DBRecord();
                Rec.CreateRow("param", Record.rec_mode, "param_pkid", Record.param_pkid);
                if (Record.rec_mode == "ADD")
                    Rec.InsertString("param_type", Record.param_type);


                Rec.InsertString("param_code", Record.param_code.Replace(" ", ""));
                Rec.InsertString("param_name", Record.param_name, "P");

                Rec.InsertString("param_file_name", Record.param_file_name, "P");


                Rec.InsertString("param_id1", Record.param_id1, "P");
                Rec.InsertString("param_id2", Record.param_id2, "P");
                Rec.InsertString("param_id3", Record.param_id3, "P");
                Rec.InsertString("param_id4", Record.param_id4, "P");

                Rec.InsertString("param_email", Record.param_email, "P");

                Rec.InsertNumeric("param_rate", Record.param_rate.ToString());

                if ( Record.rec_locked)
                    Rec.InsertString("rec_locked", "Y");
                else
                    Rec.InsertString("rec_locked", "N");


                if (Record.rec_mode == "ADD")
                {
                    Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                    Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                    Rec.InsertString("rec_created_by", Record._globalvariables.user_code);

                    if ( Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_created_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_created_date", "getdate()");

                }
                if (Record.rec_mode == "EDIT")
                {
                    Rec.InsertString("rec_edited_by", Record._globalvariables.user_code);
                    if (Con_Oracle.DB == "ORACLE")
                        Rec.InsertFunction("rec_edited_date", "SYSDATE");
                    else
                        Rec.InsertFunction("rec_edited_date", "getdate()");
                }


                sql = Rec.UpdateRow();

                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);

                Con_Oracle.CommitTransaction();

                sql = "select param_slno from param where param_pkid = '" + Record.param_pkid + "'";
                iSlno = Con_Oracle.ExecuteScalar(sql);

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
            Con_Oracle.CloseConnection();

            RetData.Add("slno", iSlno);
            return RetData;
        }

        public IDictionary<string, object> getSettings(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();

            Con_Oracle = new DBConnection();
            List<Settings> mList = new List<Settings>();
            Settings mRow;
            Lockingm lRow = new Lockingm();

            string parentid = SearchData["parentid"].ToString();

            string company_code = "";
            string branch_code = "";
            string year_code = "";
            if (SearchData.ContainsKey("company_code"))
                company_code = SearchData["company_code"].ToString();
            if (SearchData.ContainsKey("branch_code"))
                branch_code = SearchData["branch_code"].ToString();
            if (SearchData.ContainsKey("year_code"))
                year_code = SearchData["year_code"].ToString();

            try
            {

                DataTable Dt_List = new DataTable();
                sql = "";

                sql += " select parentid, tablename, caption, id, code, name from settings  where parentid = '{PARENTID}' and tablename = 'TEXT' ";

                sql = sql.Replace("{PARENTID}", parentid);

                Dt_List = Con_Oracle.ExecuteQuery(sql);

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mRow = new Settings();
                    mRow.parentid = Dr["parentid"].ToString();
                    mRow.tablename = Dr["tablename"].ToString();
                    mRow.caption = Dr["caption"].ToString();
                    mRow.id = Dr["id"].ToString();
                    mRow.code = Dr["code"].ToString();
                    mRow.name = Dr["name"].ToString();
                    mList.Add(mRow);
                }


                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("lockrecord", lRow);
            RetData.Add("list", mList);
            return RetData;
        }

        private Lockingm GetLockingDate(string comp_code, string br_code, string yr_code)
        {
            string sDate = "";

            //sDate = Lib.DatetoString(DateTime.Now);

            Lockingm lRow = new Lockingm();
            lRow = new Lockingm();
            lRow.lock_year = 0;
            lRow.lock_ar = sDate;
            lRow.lock_ap = sDate;
            lRow.lock_drn = sDate;
            lRow.lock_crn = sDate;
            lRow.lock_dri = sDate;
            lRow.lock_cri = sDate;
            lRow.lock_cr = sDate;
            lRow.lock_cp = sDate;
            lRow.lock_br = sDate;
            lRow.lock_bp = sDate;
            lRow.lock_jv = sDate;
            lRow.lock_cjv = sDate;

            sql = " select lock_pkid,lock_year,lock_in as lock_ar ";
            sql += " ,lock_pn as lock_ap,lock_dn as lock_drn,lock_cn as lock_crn ";
            sql += " ,lock_di as lock_dri,lock_ci as lock_cri ,lock_cr,lock_cp,lock_br,lock_bp";
            sql += " ,lock_jv,lock_ho as lock_cjv ";
            sql += " from lockingm a";
            sql += " where a.rec_company_code = '{COMPCODE}'";
            sql += " and a.rec_branch_code = '{BRCODE}'";
            sql += " and a.lock_year = {YEARCODE} ";

            sql = sql.Replace("{COMPCODE}", comp_code);
            sql = sql.Replace("{BRCODE}", br_code);
            sql = sql.Replace("{YEARCODE}", yr_code);

            DataTable Dt_Lock = new DataTable();
            Dt_Lock = Con_Oracle.ExecuteQuery(sql);
            foreach (DataRow Dr in Dt_Lock.Rows)
            {
                lRow.lock_year = Lib.Conv2Integer(Dr["lock_year"].ToString());
                lRow.lock_ar = Lib.DatetoString(Dr["lock_ar"]);
                lRow.lock_ap = Lib.DatetoString(Dr["lock_ap"]);
                lRow.lock_drn = Lib.DatetoString(Dr["lock_drn"]);
                lRow.lock_crn = Lib.DatetoString(Dr["lock_crn"]);
                lRow.lock_dri = Lib.DatetoString(Dr["lock_dri"]);
                lRow.lock_cri = Lib.DatetoString(Dr["lock_cri"]);
                lRow.lock_cr = Lib.DatetoString(Dr["lock_cr"]);
                lRow.lock_cp = Lib.DatetoString(Dr["lock_cp"]);
                lRow.lock_br = Lib.DatetoString(Dr["lock_br"]);
                lRow.lock_bp = Lib.DatetoString(Dr["lock_bp"]);
                lRow.lock_jv = Lib.DatetoString(Dr["lock_jv"]);
                lRow.lock_cjv = Lib.DatetoString(Dr["lock_cjv"]);
                break;
            }

            Dt_Lock.Rows.Clear();
            return lRow;
        }

        public Dictionary<string, object> SaveLockings(Lockingm Record)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();
                
                Record.lock_pkid = Guid.NewGuid().ToString().ToUpper(); 
                Record.rec_mode = "ADD";

                sql = "select lock_pkid from lockingm ";
                sql += " where rec_company_code = '" + Record._globalvariables.comp_code + "'";
                sql += " and rec_branch_code = '" + Record._globalvariables.branch_code + "'";
                sql += " and lock_year = " + Record._globalvariables.year_code;
                DataTable dt_lock = new DataTable();
                dt_lock = Con_Oracle.ExecuteQuery(sql);
                if (dt_lock.Rows.Count > 0)
                {
                    Record.lock_pkid = dt_lock.Rows[0]["lock_pkid"].ToString();
                    Record.rec_mode = "EDIT";
                }
                dt_lock.Rows.Clear();

                DBRecord Rec = new DBRecord();
                Rec.CreateRow("lockingm", Record.rec_mode, "lock_pkid", Record.lock_pkid);
                Rec.InsertDate("lock_in", Record.lock_ar);
                Rec.InsertDate("lock_pn", Record.lock_ap);
                Rec.InsertDate("lock_dn", Record.lock_drn);
                Rec.InsertDate("lock_cn", Record.lock_crn);
                Rec.InsertDate("lock_di", Record.lock_dri);
                Rec.InsertDate("lock_ci", Record.lock_cri);
                Rec.InsertDate("lock_cr", Record.lock_cr);
                Rec.InsertDate("lock_cp", Record.lock_cp);
                Rec.InsertDate("lock_br", Record.lock_br);
                Rec.InsertDate("lock_bp", Record.lock_bp);
                Rec.InsertDate("lock_jv", Record.lock_jv);
                Rec.InsertDate("lock_ho", Record.lock_cjv);
                Rec.InsertString("rec_company_code", Record._globalvariables.comp_code);
                Rec.InsertString("rec_branch_code", Record._globalvariables.branch_code);
                Rec.InsertNumeric("lock_year", Record._globalvariables.year_code.ToString());

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

        public Dictionary<string, object> SaveSettings(Settings_VM RecordVM)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string ErrorMessage = "";
            try
            {
                Con_Oracle = new DBConnection();

                DBRecord Rec = new DBRecord();

                Con_Oracle.BeginTransaction();
                foreach (Settings Record in RecordVM.RecordDet)
                {
                    sql = "delete from settings where parentid = '" + Record.parentid + "' and caption = '" + Record.caption + "'";
                    Con_Oracle.ExecuteQuery(sql);
                    Rec = new DBRecord();
                    Rec.CreateRow("settings", "ADD", "caption", Record.caption);
                    Rec.InsertString("parentid", Record.parentid);
                    Rec.InsertString("tablename", Record.tablename);
                    Rec.InsertString("id", Record.id);
                    Rec.InsertString("code", Record.code);
                    Rec.InsertString("name", Record.name, "P");
                    sql = Rec.UpdateRow();
                    Con_Oracle.ExecuteNonQuery(sql);
                }
                
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
            Con_Oracle.CloseConnection();
            return RetData;
        }

        public IDictionary<string, object> DataTransfer(Dictionary<string, object> SearchData)
        {
            string report_folder = "";
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            //Con_Oracle = new DBConnection();
            string ErrorMessage = "";
            string[] Csvfiles = null;
            string FileText = "";
            string[] strlines = null;
            string TabName = "";
            string[] tColsNames = null;
            string[] tColsValues = null;
            DataTable Dt_TransData = new DataTable();
            DataRow Dr_Target = null;
            int SavedRecords = 0;
            int RowscopiedFrmExcel = 0;
            string SaveMessage = "Transfer Complete ";
            string Target_report_folder = "";
            string sRowid = "";
            char splitChar = '\t';
            try
            {
                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();

                Csvfiles = System.IO.Directory.GetFiles(report_folder);//, "*.txt|*.csv"

                if (Csvfiles.Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Csv or Txt File Not Found");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                sRowid = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                foreach (string strfile in Csvfiles)
                {
                    RowscopiedFrmExcel = 0;
                    FileText = System.IO.File.ReadAllText(strfile);
                    FileText = FileText.Replace("\r", "");
                    strlines = FileText.Split('\n');
                    if (strlines.Length > 1)
                    {
                        TabName = System.IO.Path.GetFileNameWithoutExtension(strfile);
                        tColsNames = strlines[0].Split(splitChar);
                        Dt_TransData = CreateTable(TabName, tColsNames);
                        for (int LineIndex = 1; LineIndex < strlines.Length; LineIndex++)
                        {
                            tColsValues = strlines[LineIndex].Split(splitChar);
                            if (tColsNames.Length == tColsValues.Length)
                            {
                                RowscopiedFrmExcel++;
                                Dr_Target = Dt_TransData.NewRow();
                                for (int i = 0; i < tColsNames.Length; i++)
                                    Dr_Target[tColsNames[i]] = tColsValues[i].Trim();
                                Dt_TransData.Rows.Add(Dr_Target);
                            }
                        }
                        Dt_TransData.AcceptChanges();
                    }
                    SavedRecords = InsertTable(TabName, Dt_TransData, sRowid);
                    SaveMessage += " | Table-DT_" + TabName + ", CopiedRows-" + RowscopiedFrmExcel + ", savedRows-" + SavedRecords;

                    /*
                    //File Moving to targetfolder
                    Target_report_folder = report_folder + "\\Processed";
                    if (!System.IO.Directory.Exists(Target_report_folder))
                        System.IO.Directory.CreateDirectory(Target_report_folder);
                    Target_report_folder += "\\" + System.IO.Path.GetFileName(strfile);
                    if (System.IO.File.Exists(Target_report_folder))
                        System.IO.File.Delete(Target_report_folder);
                    System.IO.File.Move(strfile, Target_report_folder);
                    */
                }
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
            RetData.Add("savemsg", SaveMessage);
            return RetData;
        }

        private DataTable CreateTable(string tableName, string[] tColsNames)
        {
            DataTable Dt_temp = new DataTable();
            //sql = "select * from tab where tname ='DT_" + tableName + "'";
            //if (!Con_Oracle.IsRowExists(sql))
            //{
            //    sql = "Create table {DT-TNAME} (ROW_ID nvarchar2(40))";
            //    sql = sql.Replace("{DT-TNAME}", "DT_" + tableName.ToUpper());
            //    Con_Oracle.ExecuteNonQuery(sql);
            //}
            try
            {
                sql = "Create table {DT-TNAME} (ROW_PKID nvarchar2(40),ROW_ID nvarchar2(40),ROW_STATUS char(1))";
                sql = sql.Replace("{DT-TNAME}", "DT_" + tableName.ToUpper());
                Con_Oracle = new DBConnection();
                Con_Oracle.ExecuteNonQuery(sql);

            }
            catch (Exception)
            {

            }
            foreach (string sColName in tColsNames)
            {
                Dt_temp.Columns.Add(sColName, typeof(string));
                try
                {
                    sql = "Alter table DT_" + tableName.ToUpper() + " add " + sColName.ToUpper() + " Nvarchar2(" + GetColWidth(sColName) + ")";
                    Con_Oracle = new DBConnection();
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CloseConnection();

                }
                catch (Exception)
                {

                }
            }

            return Dt_temp;
        }

        private string GetColWidth(string ColName)
        {
            string sWidth = "100";
            if (ColName.ToUpper() == "CODE")
                sWidth = "15";
            else if (ColName.ToUpper() == "NAME")
                sWidth = "100";

            return sWidth;
        }

        private int InsertTable(string TableName, DataTable Dt_Data, string sRowid)
        {
            string StrFldVal = "";
            string StrFldNam = "";
            int RecCount = 0;
            foreach (DataRow dRow in Dt_Data.Rows)
            {
                RecCount++;
                StrFldVal = "";
                StrFldNam = "";
                foreach (DataColumn dCol in Dt_Data.Columns)
                {
                    if (StrFldNam != "")
                        StrFldNam += ",";
                    StrFldNam += dCol.ColumnName;

                    if (StrFldVal != "")
                        StrFldVal += ",";
                    if (dRow[dCol.ColumnName].ToString().Trim() == "")
                        StrFldVal += "NULL";
                    else
                    {
                        if (dRow[dCol.ColumnName].ToString().Contains("'"))
                            dRow[dCol.ColumnName] = dRow[dCol.ColumnName].ToString().Replace("'", "");
                        StrFldVal += "'" + GetTruncate(dRow[dCol.ColumnName].ToString(), GetColWidth(dCol.ColumnName)) + "'";
                    }
                }

                sql = " insert into DT_" + TableName + " (ROW_PKID,ROW_ID,ROW_STATUS," + StrFldNam + ") values('" + Guid.NewGuid().ToString().ToUpper() + "','" + sRowid + "','N'," + StrFldVal + ")";
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
            }

            return RecCount;
        }

        private string GetTruncate(string sData, string sWidth)
        {
            int sLen = Lib.Conv2Integer(sWidth);
            if (sData.Length > sLen)
                sData = sData.Substring(0, sLen);
            return sData.ToUpper();
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
            parameter.Add("param_type", "IMPORT DATA");
            parameter.Add("comp_code", comp_code);
            RetData.Add("dtlist", lovservice.Lov(parameter)["param"]);

            return RetData;

        }


        public IDictionary<string, object> SaveImportData(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Con_Oracle = new DBConnection();
            string ErrorMessage = "";
            string SaveMessage = "Transfer Complete ";
            string paramcode = "";
            string paramtype = "";
            string Dt_tablename = "";
            string year_start_date = "";
            string year_end_date = "";
            string year_prefix = "";
            string year_code = "";

            int RecCount = 0;

            try
            {
                string comp_code = "";
                if (SearchData.ContainsKey("comp_code"))
                    comp_code = SearchData["comp_code"].ToString();
                string branch_code = "";
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();

                string user_code = "";
                if (SearchData.ContainsKey("user_code"))
                    user_code = SearchData["user_code"].ToString();

                year_start_date = "";
                if (SearchData.ContainsKey("year_start_date"))
                    year_start_date = SearchData["year_start_date"].ToString();
                year_end_date = "";
                if (SearchData.ContainsKey("year_end_date"))
                    year_end_date = SearchData["year_end_date"].ToString();
                year_prefix = "";
                if (SearchData.ContainsKey("year_prefix"))
                    year_prefix = SearchData["year_prefix"].ToString();
                year_code = "";
                if (SearchData.ContainsKey("year_code"))
                    year_code = SearchData["year_code"].ToString();


                if (SearchData.ContainsKey("paramcode"))
                    paramcode = SearchData["paramcode"].ToString();

                if (SearchData.ContainsKey("paramtype"))
                    paramtype = SearchData["paramtype"].ToString();

                if (paramtype.Length <= 0)
                    Lib.AddError(ref ErrorMessage, "Transfer Table Type Not Found");

                if (ErrorMessage != "")
                    throw new Exception(ErrorMessage);

                Dt_tablename = "DT_" + paramtype.Replace(" ", ""); //Import data Display name

                SetParamData(comp_code);

                if (paramcode.StartsWith("PARAM"))
                    RecCount = InsertParam(Dt_tablename, paramtype, comp_code, branch_code, user_code);

                if (paramcode == "AGENT")
                    RecCount = InsertAgent(Dt_tablename, "AGENT", comp_code, branch_code, user_code);

                if (paramcode == "CONSIGNEE")
                    RecCount = InsertConsignee(Dt_tablename, "CONSIGNEE", comp_code, branch_code, user_code);

                if (paramcode == "SHIPPER")
                    RecCount = InsertAgent(Dt_tablename, "SHIPPER", comp_code, branch_code, user_code);

                if (paramcode == "ACCTM")
                    RecCount = InsertAcctm(Dt_tablename, "ACCTM", comp_code, branch_code, user_code);

                if (paramcode == "BRANCHES")
                    RecCount = InsertBranches(Dt_tablename, "BRANCHES", comp_code, branch_code, user_code);


                //if (paramcode == "OPENING")
                //    RecCount = InsertOpening(Dt_tablename, "OP", comp_code, branch_code, user_code, year_start_date, year_end_date, year_prefix, year_code);
                //if (paramcode == "SHIPPER_OPENING")
                //    RecCount = InsertOpening(Dt_tablename, "OI", comp_code, branch_code, user_code, year_start_date, year_end_date, year_prefix, year_code);



                /*
                if (paramcode == "OTHER_OPENING")
                    RecCount = InsertOtherOpening(Dt_tablename, "OP", comp_code, branch_code, user_code, year_start_date, year_end_date, year_prefix, year_code);
                



                if (paramcode == "OPENING")
                    RecCount = InsertAgentOpening(Dt_tablename, "OP", comp_code, branch_code, user_code, year_start_date, year_end_date, year_prefix, year_code);

                if (paramcode == "SHIPPER_OPENING")
                    RecCount = InsertAgentOpening(Dt_tablename, "OC", comp_code, branch_code, user_code, year_start_date, year_end_date, year_prefix, year_code);

                
                 */


                SaveMessage = " | Table-" + Dt_tablename + ", SavedRows-" + RecCount.ToString();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
            Con_Oracle.CloseConnection();
            RetData.Add("savemsg", SaveMessage);
            return RetData;
        }



        private int InsertParam(string tablename, string paramtype, string sCOMPANY_CODE, string sBRANCH_CODE, string sUSER)
        {
            int RecCount = 0;
            sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by code ";
            DataTable Dt_Table = Con_Oracle.ExecuteQuery(sql);

            string ErrName = "";

            foreach (DataRow Dr in Dt_Table.Rows)
            {
                try
                {
                    sql = " INSERT INTO PARAM ( ";
                    sql += " PARAM_PKID,REC_COMPANY_CODE,REC_BRANCH_CODE";
                    sql += " ,PARAM_TYPE,PARAM_CODE,PARAM_NAME,";
                    if (Dr.Table.Columns.Contains("ID1"))
                        sql += " PARAM_ID1,";
                    if (Dr.Table.Columns.Contains("ID2"))
                        sql += " PARAM_ID2,";
                    if (Dr.Table.Columns.Contains("ID3"))
                        sql += " PARAM_ID3,";
                    if (Dr.Table.Columns.Contains("ID4"))
                        sql += " PARAM_ID4,";
                    if (Dr.Table.Columns.Contains("EMAIL"))
                        sql += " PARAM_EMAIL,";
                    if (Dr.Table.Columns.Contains("RATE"))
                        sql += " PARAM_RATE,";
                    sql += " REC_CREATED_BY,REC_CREATED_DATE";
                    sql += " ) VALUES (";
                    sql += " [PARAM_PKID],[REC_COMPANY_CODE],[REC_BRANCH_CODE],";
                    sql += " [PARAM_TYPE],[PARAM_CODE],[PARAM_NAME],";
                    if (Dr.Table.Columns.Contains("ID1"))
                        sql += " [PARAM_ID1],";
                    if (Dr.Table.Columns.Contains("ID2"))
                        sql += " [PARAM_ID2],";
                    if (Dr.Table.Columns.Contains("ID3"))
                        sql += " [PARAM_ID3],";
                    if (Dr.Table.Columns.Contains("ID4"))
                        sql += " [PARAM_ID4],";
                    if (Dr.Table.Columns.Contains("EMAIL"))
                        sql += " [PARAM_EMAIL],";
                    if (Dr.Table.Columns.Contains("RATE"))
                        sql += " [PARAM_RATE],";
                    sql += " [REC_CREATED_BY],[REC_CREATED_DATE]";
                    sql += " )";

                    sql = sql.Replace("[PARAM_PKID]", "'" + Dr["row_pkid"].ToString() + "'");
                    sql = sql.Replace("[REC_COMPANY_CODE]", "'" + sCOMPANY_CODE + "'");
                    sql = sql.Replace("[REC_BRANCH_CODE]", "'" + sBRANCH_CODE + "'");
                    sql = sql.Replace("[PARAM_TYPE]", "'" + paramtype + "'");
                    sql = sql.Replace("[PARAM_CODE]", "'" + Dr["code"].ToString() + "'");
                    sql = sql.Replace("[PARAM_NAME]", "'" + Dr["name"].ToString() + "'");
                    if (Dr.Table.Columns.Contains("ID1"))
                        sql = sql.Replace("[PARAM_ID1]", "'" + Dr["ID1"].ToString() + "'");
                    if (Dr.Table.Columns.Contains("ID2"))
                        sql = sql.Replace("[PARAM_ID2]", "'" + Dr["ID2"].ToString() + "'");
                    if (Dr.Table.Columns.Contains("ID3"))
                        sql = sql.Replace("[PARAM_ID3]", "'" + Dr["ID3"].ToString() + "'");
                    if (Dr.Table.Columns.Contains("ID4"))
                        sql = sql.Replace("[PARAM_ID4]", "'" + Dr["ID4"].ToString() + "'");
                    if (Dr.Table.Columns.Contains("EMAIL"))
                        sql = sql.Replace("[PARAM_EMAIL]", "'" + Dr["EMAIL"].ToString() + "'");
                    if (Dr.Table.Columns.Contains("RATE"))
                        sql = sql.Replace("[PARAM_RATE]", Dr["RATE"].ToString() + "'");
                    sql = sql.Replace("[REC_CREATED_BY]", "'" + sUSER + "'");
                    sql = sql.Replace("[REC_CREATED_DATE]", "sysdate");

                    ErrName = Dr["code"].ToString() + "~" + Dr["name"].ToString() + "~";

                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = "update " + tablename + " set row_status='Y' where row_pkid='" + Dr["row_pkid"].ToString() + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    RecCount++;
                }
                catch (Exception Ex)
                {
                    Con_Oracle.CreateErrorLog(ErrName + Ex.Message.ToString());
                    Con_Oracle.RollbackTransaction();
                }
            }
            Con_Oracle.CloseConnection();
            return RecCount;
        }


        private int InsertAgent(string tablename, string paramtype, string sCOMPANY_CODE, string sBRANCH_CODE, string sUSER)
        {
            int RecCount = 0;
            sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by code ";
            if (paramtype == "SHIPPER")
                sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by code ";
            if (paramtype == "CONSIGNEE")
                sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by name ";

            DataTable Dt_Table = Con_Oracle.ExecuteQuery(sql);


            string sql2 = "";
            string sql3 = "";
            string sqlacc = "";

            string ErrName = "";

            string CODE = "";
            string ID = "";

            int iCtr = 0;

            foreach (DataRow Dr in Dt_Table.Rows)
            {
                try
                {
                    if (CODE != Dr["code"].ToString())
                    {
                        iCtr = 0;
                        CODE = Dr["code"].ToString();
                        ID = Dr["row_pkid"].ToString();

                        sql2 = " insert into customerm (";
                        sql2 += "CUST_PKID,CUST_CODE,CUST_NAME,CUST_IECODE,CUST_TYPE,CUST_CLASS,";
                        sql2 += "CUST_SMAN_ID,CUST_CSD_ID,CUST_PANNO,CUST_TANNO,CUST_CRDAYS,CUST_CRLIMIT, ";
                        sql2 += "CUST_CRDATE,CUST_LINKED,CUST_IS_AGENT,CUST_IS_SHIPPER,CUST_IS_CONSIGNEE,CUST_IS_CHA,CUST_IS_CREDITOR,";
                        sql2 += "cust_adcode, cust_bank, cust_bank_branch, cust_acno, cust_forexacno, cust_bank_address1, cust_bank_address2,cust_bank_address3,";
                        sql2 += "CUST_IS_OTHERS,CUST_SEPZ_UNIT,CUST_NOMINATION,CUST_REFERDBY,";
                        sql2 += "REC_LOCKED,REC_COMPANY_CODE,REC_BRANCH_CODE,REC_CREATED_BY,REC_CREATED_DATE";
                        sql2 += ") VALUES (";
                        sql2 += "[CUST_PKID],[CUST_CODE],[CUST_NAME],[CUST_IECODE],[CUST_TYPE],[CUST_CLASS],";
                        sql2 += "[CUST_SMAN_ID],[CUST_CSD_ID],[CUST_PANNO],[CUST_TANNO],[CUST_CRDAYS],[CUST_CRLIMIT], ";
                        sql2 += "[CUST_CRDATE],[CUST_LINKED],[CUST_IS_AGENT],[CUST_IS_SHIPPER],[CUST_IS_CONSIGNEE],[CUST_IS_CHA],[CUST_IS_CREDITOR],";
                        sql2 += "[CUST_ADCODE],[CUST_BANK],[CUST_BANK_BRANCH],[CUST_ACNO],[CUST_FOREXACNO],[CUST_BANK_ADDRESS1],[CUST_BANK_ADDRESS2],[CUST_BANK_ADDRESS3],";
                        sql2 += "[CUST_IS_OTHERS],[CUST_SEPZ_UNIT],[CUST_NOMINATION],[CUST_REFERDBY],";
                        sql2 += "[REC_LOCKED],[REC_COMPANY_CODE],[REC_BRANCH_CODE],[REC_CREATED_BY],[REC_CREATED_DATE]";
                        sql2 += ")";


                        sql2 = sql2.Replace("[CUST_PKID]", "'" + Dr["row_pkid"].ToString() + "'");
                        sql2 = sql2.Replace("[CUST_CODE]", "'" + Dr["code"].ToString() + "'");
                        sql2 = sql2.Replace("[CUST_NAME]", "'" + Dr["name"].ToString() + "'");
                        sql2 = sql2.Replace("[CUST_IECODE]", "''");

                        ErrName = Dr["code"].ToString() + "~" + Dr["name"].ToString() + "~";

                        if (paramtype == "SHIPPER")
                        {
                            sql2 = sql2.Replace("[CUST_TYPE]", "'" + getData(Dr, "TYPE") + "'");
                            sql2 = sql2.Replace("[CUST_CLASS]", "'" + getData(Dr, "CLASS") + "'");
                            sql2 = sql2.Replace("[CUST_SMAN_ID]", "'" + GetByName("SALESMAN", Dr["SMAN"].ToString()) + "'");
                            sql2 = sql2.Replace("[CUST_CSD_ID]", "'" + GetByName("SALESMAN", Dr["CSD"].ToString()) + "'");
                            sql2 = sql2.Replace("[CUST_PANNO]", "'" + Dr["pan"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_TANNO]", "'" + Dr["tan"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_CRDAYS]", Lib.Conv2Integer(Dr["crdays"].ToString()).ToString());
                            sql2 = sql2.Replace("[CUST_CRLIMIT]", Lib.Conv2Decimal(Dr["crlimit"].ToString()).ToString());

                            sql2 = sql2.Replace("[CUST_CRDATE]", "NULL");
                            sql2 = sql2.Replace("[CUST_LINKED]", "'Y'");

                            sql2 = sql2.Replace("[CUST_ADCODE]", "'" + Dr["adcode"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_ACNO]", "'" + Dr["acno"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_FOREXACNO]", "'" + Dr["forexacno"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_BANK]", "'" + Dr["bank"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_BANK_BRANCH]", "'" + Dr["bank_branch"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_BANK_ADDRESS1]", "'" + Dr["bank_Address1"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_BANK_ADDRESS2]", "'" + Dr["bank_Address2"].ToString() + "'");
                            sql2 = sql2.Replace("[CUST_BANK_ADDRESS3]", "'" + Dr["bank_Address3"].ToString() + "'");



                            //[CUST_DBKACNO],[CUST_ADCODE],[CUST_BANK],[CUST_BANK_BRANCH],[CUST_ACNO],[CUST_FOREXACNO],[CUST_BANK_ADDRESS1],[CUST_BANK_ADDRESS2],[CUST_BANK_ADDRESS3]


                        }
                        else
                        {
                            sql2 = sql2.Replace("[CUST_TYPE]", "'N'");
                            sql2 = sql2.Replace("[CUST_CLASS]", "'N'");
                            sql2 = sql2.Replace("[CUST_SMAN_ID]", "''");
                            sql2 = sql2.Replace("[CUST_CSD_ID]", "''");
                            sql2 = sql2.Replace("[CUST_PANNO]", "''");
                            sql2 = sql2.Replace("[CUST_TANNO]", "''");
                            sql2 = sql2.Replace("[CUST_CRDAYS]", "0");
                            sql2 = sql2.Replace("[CUST_CRLIMIT]", "0");
                            sql2 = sql2.Replace("[CUST_CRDATE]", "NULL");
                            sql2 = sql2.Replace("[CUST_LINKED]", "'N'");

                            sql2 = sql2.Replace("[CUST_ADCODE]", "''");
                            sql2 = sql2.Replace("[CUST_ACNO]", "''");
                            sql2 = sql2.Replace("[CUST_FOREXACNO]", "''");
                            sql2 = sql2.Replace("[CUST_BANK]", "''");
                            sql2 = sql2.Replace("[CUST_BANK_BRANCH]", "''");
                            sql2 = sql2.Replace("[CUST_BANK_ADDRESS1]", "''");
                            sql2 = sql2.Replace("[CUST_BANK_ADDRESS2]", "''");
                            sql2 = sql2.Replace("[CUST_BANK_ADDRESS3]", "''");


                        }
                        if (paramtype == "AGENT")
                            sql2 = sql2.Replace("[CUST_IS_AGENT]", "'Y'");
                        else
                            sql2 = sql2.Replace("[CUST_IS_AGENT]", "'N'");
                        if (paramtype == "SHIPPER")
                            sql2 = sql2.Replace("[CUST_IS_SHIPPER]", "'Y'");
                        else
                            sql2 = sql2.Replace("[CUST_IS_SHIPPER]", "'N'");

                        if (paramtype == "CONSIGNEE")
                            sql2 = sql2.Replace("[CUST_IS_CONSIGNEE]", "'Y'");
                        else
                            sql2 = sql2.Replace("[CUST_IS_CONSIGNEE]", "'N'");

                        sql2 = sql2.Replace("[CUST_IS_CHA]", "'N'");
                        sql2 = sql2.Replace("[CUST_IS_OTHERS]", "'N'");
                        sql2 = sql2.Replace("[CUST_IS_CREDITOR]", "'N'");

                        if (paramtype == "SHIPPER")
                            sql2 = sql2.Replace("[CUST_SEPZ_UNIT]", "'" + Dr["SEZUNIT"].ToString() + "'");
                        else
                            sql2 = sql2.Replace("[CUST_SEPZ_UNIT]", "'N'");

                        if (paramtype == "CONSIGNEE")
                            sql2 = sql2.Replace("[CUST_NOMINATION]", "'" + Dr["NOMINATION"].ToString() + "'");
                        else
                            sql2 = sql2.Replace("[CUST_NOMINATION]", "'NA'");

                        sql2 = sql2.Replace("[CUST_REFERDBY]", "''");
                        sql2 = sql2.Replace("[REC_LOCKED]", "'N'");
                        sql2 = sql2.Replace("[REC_COMPANY_CODE]", "'" + sCOMPANY_CODE + "'");
                        sql2 = sql2.Replace("[REC_BRANCH_CODE]", "'" + sBRANCH_CODE + "'");
                        sql2 = sql2.Replace("[REC_CREATED_BY]", "'" + sUSER + "'");
                        sql2 = sql2.Replace("[REC_CREATED_DATE]", "sysdate");


                        if (paramtype == "SHIPPER")
                        {
                            sqlacc = " insert into acctm ";
                            sqlacc += "(ACC_PKID,ACC_CODE,ACC_NAME,ACC_MAIN_ID,ACC_MAIN_CODE,ACC_MAIN_NAME,ACC_GROUP_ID,ACC_TYPE_ID,ACC_AGAINST_INVOICE,ACC_COST_CENTRE,ACC_SAC_ID,REC_COMPANY_CODE,REC_LOCKED)";
                            sqlacc += " values (";
                            sqlacc += " [ACC_PKID],[ACC_CODE],[ACC_NAME],[ACC_MAIN_ID],[ACC_MAIN_CODE],[ACC_MAIN_NAME],[ACC_GROUP_ID],[ACC_TYPE_ID],[ACC_AGAINST_INVOICE],[ACC_COST_CENTRE],[ACC_SAC_ID],[REC_COMPANY_CODE], [REC_LOCKED] ";
                            sqlacc += " )";

                            sqlacc = sqlacc.Replace("[ACC_PKID]", "'" + Dr["row_pkid"].ToString() + "'");
                            sqlacc = sqlacc.Replace("[ACC_CODE]", "'" + Dr["code"].ToString() + "'");
                            sqlacc = sqlacc.Replace("[ACC_NAME]", "'" + Dr["name"].ToString() + "'");
                            sqlacc = sqlacc.Replace("[ACC_MAIN_ID]", "'" + Dr["row_pkid"].ToString() + "'");
                            sqlacc = sqlacc.Replace("[ACC_MAIN_CODE]", "'" + Dr["code"].ToString() + "'");
                            sqlacc = sqlacc.Replace("[ACC_MAIN_NAME]", "'" + Dr["name"].ToString() + "'");
                            sqlacc = sqlacc.Replace("[ACC_GROUP_ID]", "'" + GetByName("ACGROUP", "SUNDRY DEBTORS") + "'");
                            sqlacc = sqlacc.Replace("[ACC_TYPE_ID]", "'" + GetByName("ACTYPE", "DEBTORS") + "'");
                            sqlacc = sqlacc.Replace("[ACC_AGAINST_INVOICE]", "'D'");
                            sqlacc = sqlacc.Replace("[ACC_COST_CENTRE]", "'N'");
                            sqlacc = sqlacc.Replace("[ACC_SAC_ID]", "''");
                            sqlacc = sqlacc.Replace("[REC_LOCKED]", "'N'");
                            sqlacc = sqlacc.Replace("[REC_COMPANY_CODE]", "'" + sCOMPANY_CODE + "'");
                            sqlacc = sqlacc.Replace("[REC_BRANCH_CODE]", "'" + sBRANCH_CODE + "'");
                            sqlacc = sqlacc.Replace("[REC_CREATED_BY]", "'" + sUSER + "'");
                            sqlacc = sqlacc.Replace("[REC_CREATED_DATE]", "sysdate");
                        }
                    }
                    iCtr++;
                    sql3 = " INSERT INTO ADDRESSM (";
                    sql3 += " ADD_PKID,ADD_PARENT_ID,ADD_SOURCE,ADD_BRANCH_SLNO,ADD_GST_TYPE,ADD_GSTIN, ";
                    sql3 += " ADD_LINE1, ADD_LINE2, ADD_LINE3, ADD_LINE4, ADD_CITY_ID, ADD_STATE_ID, ADD_COUNTRY_ID, ";
                    sql3 += " ADD_PIN,ADD_CONTACT,ADD_TEL,ADD_FAX,ADD_EMAIL,ADD_WEB,ADD_ORDER,ADD_CITY, ";
                    sql3 += " ADD_LOCATION, REC_COMPANY_CODE ";
                    sql3 += ") values (";
                    sql3 += " [ADD_PKID],[ADD_PARENT_ID],[ADD_SOURCE],[ADD_BRANCH_SLNO],[ADD_GST_TYPE],[ADD_GSTIN], ";
                    sql3 += " [ADD_LINE1], [ADD_LINE2], [ADD_LINE3], [ADD_LINE4], [ADD_CITY_ID], [ADD_STATE_ID], [ADD_COUNTRY_ID], ";
                    sql3 += " [ADD_PIN],[ADD_CONTACT],[ADD_TEL],[ADD_FAX],[ADD_EMAIL],[ADD_WEB],[ADD_ORDER],[ADD_CITY], ";
                    sql3 += " [ADD_LOCATION], [REC_COMPANY_CODE]";
                    sql3 += ")";

                    sql3 = sql3.Replace("[ADD_PKID]", "'" + System.Guid.NewGuid().ToString().ToUpper() + "'");
                    sql3 = sql3.Replace("[ADD_PARENT_ID]", "'" + ID + "'");
                    sql3 = sql3.Replace("[ADD_SOURCE]", "'CUSTOMER'");
                    sql3 = sql3.Replace("[ADD_BRANCH_SLNO]", Dr["BRANCH"].ToString());

                    if (paramtype == "SHIPPER")
                    {
                        sql3 = sql3.Replace("[ADD_GSTIN]", "'" + Dr["GSTIN"].ToString() + "'");
                        sql3 = sql3.Replace("[ADD_GST_TYPE]", "'" + Dr["GST_TYPE"].ToString() + "'");
                    }
                    else
                    {
                        sql3 = sql3.Replace("[ADD_GST_TYPE]", "'NA'");
                        sql3 = sql3.Replace("[ADD_GSTIN]", "''");
                    }

                    sql3 = sql3.Replace("[ADD_LINE1]", "'" + Dr["ADDRESS1"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LINE2]", "'" + Dr["ADDRESS2"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LINE3]", "'" + Dr["ADDRESS3"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LINE4]", "'" + Dr["ADDRESS4"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_CITY_ID]", "''");

                    if (paramtype == "SHIPPER")
                    {
                        sql3 = sql3.Replace("[ADD_STATE_ID]", "'" + GetByCode("STATE", Dr["STATE"].ToString()) + "'");
                        sql3 = sql3.Replace("[ADD_PIN]", Lib.Conv2Integer(Dr["PIN"].ToString()).ToString());
                    }
                    else
                    {
                        sql3 = sql3.Replace("[ADD_STATE_ID]", "''");
                        sql3 = sql3.Replace("[ADD_PIN]", "'0'");
                    }
                    sql3 = sql3.Replace("[ADD_COUNTRY_ID]", "'" + GetByCode("COUNTRY", Dr["COUNTRY"].ToString()) + "'");

                    sql3 = sql3.Replace("[ADD_CONTACT]", "'" + getData(Dr, "CONTACT") + "'");
                    sql3 = sql3.Replace("[ADD_TEL]", "'" + Dr["TEL"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_FAX]", "'" + Dr["FAX"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_EMAIL]", "'" + Dr["EMAIL"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_WEB]", "'" + Dr["WEB"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_CITY]", "'" + Dr["CITY"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LOCATION]", "'" + getData(Dr, "LOCATION") + "'");
                    sql3 = sql3.Replace("[ADD_ORDER]", iCtr.ToString());
                    sql3 = sql3.Replace("[REC_COMPANY_CODE]", "'" + sCOMPANY_CODE + "'");

                    Con_Oracle.BeginTransaction();
                    if (sql2 != "")
                    {
                        Con_Oracle.ExecuteNonQuery(sql2);
                        sql2 = "";
                    }
                    if (sqlacc != "")
                    {
                        Con_Oracle.ExecuteNonQuery(sqlacc);
                        sqlacc = "";
                    }

                    Con_Oracle.ExecuteNonQuery(sql3);
                    sql = "update " + tablename + " set row_status='Y' where row_pkid='" + Dr["row_pkid"].ToString() + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    RecCount++;

                }
                catch (Exception Ex)
                {
                    Con_Oracle.CreateErrorLog(ErrName + Ex.Message.ToString());
                    Con_Oracle.RollbackTransaction();
                }
            }
            Con_Oracle.CloseConnection();
            return RecCount;
        }

        private int InsertConsignee(string tablename, string paramtype, string sCOMPANY_CODE, string sBRANCH_CODE, string sUSER)
        {
            int RecCount = 0;

            int SLNO = 10;

            sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by code ";
            if (paramtype == "CONSIGNEE")
                sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by name ";

            DataTable Dt_Table = Con_Oracle.ExecuteQuery(sql);


            DataTable Dt_test = new DataTable();

            string sql2 = "";
            string sql3 = "";

            string ErrName = "";

            string CODE = "";
            string NAME = "";
            string ID = "";


            int iCtr = 0;

            foreach (DataRow Dr in Dt_Table.Rows)
            {
                try
                {
                    if (NAME != Dr["name"].ToString().Trim())
                    {
                        SLNO = 1;
                        iCtr = 0;

                        CODE = Dr["name"].ToString().Trim().Replace(" ", "").PadRight(12, ' ').Substring(0, 10).Trim();
                        sql = " select count(*) as tot from customerm where cust_code ='" + CODE + SLNO.ToString() + "'";
                        string st1 = Con_Oracle.ExecuteScalar(sql).ToString();
                        SLNO = Lib.Conv2Integer(st1);
                        SLNO++;

                        CODE = CODE + SLNO.ToString();

                        NAME = Dr["NAME"].ToString().Trim();
                        ID = Dr["row_pkid"].ToString();

                        sql2 = " insert into customerm (";
                        sql2 += "CUST_PKID,CUST_CODE,CUST_NAME,CUST_IECODE,CUST_TYPE,CUST_CLASS,";
                        sql2 += "CUST_SMAN_ID,CUST_CSD_ID,CUST_PANNO,CUST_TANNO,CUST_CRDAYS,CUST_CRLIMIT, ";
                        sql2 += "CUST_CRDATE,CUST_LINKED,CUST_IS_AGENT,CUST_IS_SHIPPER,CUST_IS_CONSIGNEE,CUST_IS_CHA,CUST_IS_CREDITOR,";
                        sql2 += "cust_adcode, cust_bank, cust_bank_branch, cust_acno, cust_forexacno, cust_bank_address1, cust_bank_address2,cust_bank_address3,";
                        sql2 += "CUST_IS_OTHERS,CUST_SEPZ_UNIT,CUST_NOMINATION,CUST_REFERDBY,";
                        sql2 += "REC_LOCKED,REC_COMPANY_CODE,REC_BRANCH_CODE,REC_CREATED_BY,REC_CREATED_DATE";
                        sql2 += ") VALUES (";
                        sql2 += "[CUST_PKID],[CUST_CODE],[CUST_NAME],[CUST_IECODE],[CUST_TYPE],[CUST_CLASS],";
                        sql2 += "[CUST_SMAN_ID],[CUST_CSD_ID],[CUST_PANNO],[CUST_TANNO],[CUST_CRDAYS],[CUST_CRLIMIT], ";
                        sql2 += "[CUST_CRDATE],[CUST_LINKED],[CUST_IS_AGENT],[CUST_IS_SHIPPER],[CUST_IS_CONSIGNEE],[CUST_IS_CHA],[CUST_IS_CREDITOR],";
                        sql2 += "[CUST_ADCODE],[CUST_BANK],[CUST_BANK_BRANCH],[CUST_ACNO],[CUST_FOREXACNO],[CUST_BANK_ADDRESS1],[CUST_BANK_ADDRESS2],[CUST_BANK_ADDRESS3],";
                        sql2 += "[CUST_IS_OTHERS],[CUST_SEPZ_UNIT],[CUST_NOMINATION],[CUST_REFERDBY],";
                        sql2 += "[REC_LOCKED],[REC_COMPANY_CODE],[REC_BRANCH_CODE],[REC_CREATED_BY],[REC_CREATED_DATE]";
                        sql2 += ")";

                        sql2 = sql2.Replace("[CUST_PKID]", "'" + Dr["row_pkid"].ToString() + "'");

                        sql2 = sql2.Replace("[CUST_CODE]", "'" + CODE + "'");
                        sql2 = sql2.Replace("[CUST_NAME]", "'" + Dr["name"].ToString().Trim() + "'");
                        sql2 = sql2.Replace("[CUST_IECODE]", "''");

                        ErrName = Dr["code"].ToString() + "~" + Dr["name"].ToString() + "~";

                        sql2 = sql2.Replace("[CUST_TYPE]", "'N'");
                        sql2 = sql2.Replace("[CUST_CLASS]", "'N'");
                        sql2 = sql2.Replace("[CUST_SMAN_ID]", "''");
                        sql2 = sql2.Replace("[CUST_CSD_ID]", "''");
                        sql2 = sql2.Replace("[CUST_PANNO]", "''");
                        sql2 = sql2.Replace("[CUST_TANNO]", "''");
                        sql2 = sql2.Replace("[CUST_CRDAYS]", "0");
                        sql2 = sql2.Replace("[CUST_CRLIMIT]", "0");
                        sql2 = sql2.Replace("[CUST_CRDATE]", "NULL");
                        sql2 = sql2.Replace("[CUST_LINKED]", "'N'");

                        sql2 = sql2.Replace("[CUST_ADCODE]", "''");
                        sql2 = sql2.Replace("[CUST_ACNO]", "''");
                        sql2 = sql2.Replace("[CUST_FOREXACNO]", "''");
                        sql2 = sql2.Replace("[CUST_BANK]", "''");
                        sql2 = sql2.Replace("[CUST_BANK_BRANCH]", "''");
                        sql2 = sql2.Replace("[CUST_BANK_ADDRESS1]", "''");
                        sql2 = sql2.Replace("[CUST_BANK_ADDRESS2]", "''");
                        sql2 = sql2.Replace("[CUST_BANK_ADDRESS3]", "''");

                        if (paramtype == "AGENT")
                            sql2 = sql2.Replace("[CUST_IS_AGENT]", "'Y'");
                        else
                            sql2 = sql2.Replace("[CUST_IS_AGENT]", "'N'");

                        sql2 = sql2.Replace("[CUST_IS_SHIPPER]", "'N'");

                        if (paramtype == "CONSIGNEE")
                            sql2 = sql2.Replace("[CUST_IS_CONSIGNEE]", "'Y'");
                        else
                            sql2 = sql2.Replace("[CUST_IS_CONSIGNEE]", "'N'");

                        sql2 = sql2.Replace("[CUST_IS_CHA]", "'N'");
                        sql2 = sql2.Replace("[CUST_IS_OTHERS]", "'N'");
                        sql2 = sql2.Replace("[CUST_IS_CREDITOR]", "'N'");
                        sql2 = sql2.Replace("[CUST_SEPZ_UNIT]", "'N'");
                        if (paramtype == "CONSIGNEE")
                            sql2 = sql2.Replace("[CUST_NOMINATION]", "'" + Dr["NOMINATION"].ToString() + "'");
                        else
                            sql2 = sql2.Replace("[CUST_NOMINATION]", "'NA'");

                        sql2 = sql2.Replace("[CUST_REFERDBY]", "''");
                        sql2 = sql2.Replace("[REC_LOCKED]", "'N'");
                        sql2 = sql2.Replace("[REC_COMPANY_CODE]", "'" + sCOMPANY_CODE + "'");
                        sql2 = sql2.Replace("[REC_BRANCH_CODE]", "'" + sBRANCH_CODE + "'");
                        sql2 = sql2.Replace("[REC_CREATED_BY]", "'" + sUSER + "'");
                        sql2 = sql2.Replace("[REC_CREATED_DATE]", "sysdate");
                    }
                    iCtr++;
                    sql3 = " INSERT INTO ADDRESSM (";
                    sql3 += " ADD_PKID,ADD_PARENT_ID,ADD_SOURCE,ADD_BRANCH_SLNO,ADD_GST_TYPE,ADD_GSTIN, ";
                    sql3 += " ADD_LINE1, ADD_LINE2, ADD_LINE3, ADD_LINE4, ADD_CITY_ID, ADD_STATE_ID, ADD_COUNTRY_ID, ";
                    sql3 += " ADD_PIN,ADD_CONTACT,ADD_TEL,ADD_FAX,ADD_EMAIL,ADD_WEB,ADD_ORDER,ADD_CITY, ";
                    sql3 += " ADD_LOCATION, REC_COMPANY_CODE ";
                    sql3 += ") values (";
                    sql3 += " [ADD_PKID],[ADD_PARENT_ID],[ADD_SOURCE],[ADD_BRANCH_SLNO],[ADD_GST_TYPE],[ADD_GSTIN], ";
                    sql3 += " [ADD_LINE1], [ADD_LINE2], [ADD_LINE3], [ADD_LINE4], [ADD_CITY_ID], [ADD_STATE_ID], [ADD_COUNTRY_ID], ";
                    sql3 += " [ADD_PIN],[ADD_CONTACT],[ADD_TEL],[ADD_FAX],[ADD_EMAIL],[ADD_WEB],[ADD_ORDER],[ADD_CITY], ";
                    sql3 += " [ADD_LOCATION], [REC_COMPANY_CODE]";
                    sql3 += ")";

                    sql3 = sql3.Replace("[ADD_PKID]", "'" + System.Guid.NewGuid().ToString().ToUpper() + "'");
                    sql3 = sql3.Replace("[ADD_PARENT_ID]", "'" + ID + "'");
                    sql3 = sql3.Replace("[ADD_SOURCE]", "'CUSTOMER'");
                    sql3 = sql3.Replace("[ADD_BRANCH_SLNO]", Dr["BRANCH"].ToString());
                    sql3 = sql3.Replace("[ADD_GST_TYPE]", "'NA'");
                    sql3 = sql3.Replace("[ADD_GSTIN]", "''");
                    sql3 = sql3.Replace("[ADD_LINE1]", "'" + Dr["ADDRESS1"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LINE2]", "'" + Dr["ADDRESS2"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LINE3]", "'" + Dr["ADDRESS3"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LINE4]", "'" + Dr["ADDRESS4"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_CITY_ID]", "''");
                    sql3 = sql3.Replace("[ADD_STATE_ID]", "''");
                    sql3 = sql3.Replace("[ADD_PIN]", "'0'");
                    sql3 = sql3.Replace("[ADD_COUNTRY_ID]", "'" + GetByCode("COUNTRY", Dr["COUNTRY"].ToString()) + "'");
                    sql3 = sql3.Replace("[ADD_CONTACT]", "'" + getData(Dr, "CONTACT") + "'");
                    sql3 = sql3.Replace("[ADD_TEL]", "'" + Dr["TEL"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_FAX]", "'" + Dr["FAX"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_EMAIL]", "'" + Dr["EMAIL"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_WEB]", "'" + Dr["WEB"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_CITY]", "'" + Dr["CITY"].ToString() + "'");
                    sql3 = sql3.Replace("[ADD_LOCATION]", "'" + getData(Dr, "LOCATION") + "'");
                    sql3 = sql3.Replace("[ADD_ORDER]", iCtr.ToString());
                    sql3 = sql3.Replace("[REC_COMPANY_CODE]", "'" + sCOMPANY_CODE + "'");

                    Con_Oracle.BeginTransaction();
                    if (sql2 != "")
                    {
                        Con_Oracle.ExecuteNonQuery(sql2);
                        sql2 = "";
                    }

                    Con_Oracle.ExecuteNonQuery(sql3);
                    sql = "update " + tablename + " set row_status='Y' where row_pkid='" + Dr["row_pkid"].ToString() + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    RecCount++;

                }
                catch (Exception Ex)
                {
                    Con_Oracle.CreateErrorLog(ErrName + Ex.Message.ToString());
                    Con_Oracle.RollbackTransaction();
                }
            }
            Con_Oracle.CloseConnection();
            return RecCount;
        }


        private int InsertAcctm(string tablename, string paramtype, string sCOMPANY_CODE, string sBRANCH_CODE, string sUSER)
        {
            int RecCount = 0;
            sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by code ";
            DataTable Dt_Table = Con_Oracle.ExecuteQuery(sql);

            decimal nRate = 0;

            string sqlacc = "";
            string CODE = "";
            int iCtr = 0;

            int iDupTot = 0;

            string ErrName = "";
            string sName = "";
            string sName1 = "";
            foreach (DataRow Dr in Dt_Table.Rows)
            {
                try
                {
                    if (sName != Dr["name"].ToString())
                        iDupTot = 0;
                    else
                        iDupTot++;
                    sName = Dr["name"].ToString();
                    sName1 = sName;
                    if (iDupTot > 0)
                        sName1 = sName.Replace(" ", " ".PadLeft(iDupTot + 1, ' '));

                    sqlacc = " insert into acctm ";
                    sqlacc += "(ACC_PKID,ACC_CODE,ACC_NAME,ACC_MAIN_ID,ACC_MAIN_CODE,ACC_MAIN_NAME,ACC_GROUP_ID,ACC_TYPE_ID,ACC_AGAINST_INVOICE,ACC_COST_CENTRE,ACC_TAXABLE,ACC_SAC_ID,ACC_CGST_RATE,ACC_SGST_RATE, ACC_IGST_RATE,";
                    sqlacc += "ACC_BRANCH_CODE, REC_LOCKED,REC_COMPANY_CODE,REC_BRANCH_CODE,REC_CREATED_BY,REC_CREATED_DATE";
                    sqlacc += ")";
                    sqlacc += " values (";
                    sqlacc += " [ACC_PKID],[ACC_CODE],[ACC_NAME],[ACC_MAIN_ID],[ACC_MAIN_CODE],[ACC_MAIN_NAME],[ACC_GROUP_ID],[ACC_TYPE_ID],[ACC_AGAINST_INVOICE],[ACC_COST_CENTRE],[ACC_TAXABLE],[ACC_SAC_ID],[ACC_CGST_RATE],[ACC_SGST_RATE],[ACC_IGST_RATE], ";
                    sqlacc += "[ACC_BRANCH_CODE],[REC_LOCKED],[REC_COMPANY_CODE],[REC_BRANCH_CODE],[REC_CREATED_BY],[REC_CREATED_DATE]";
                    sqlacc += " )";

                    sqlacc = sqlacc.Replace("[ACC_PKID]", "'" + Dr["row_pkid"].ToString() + "'");
                    sqlacc = sqlacc.Replace("[ACC_CODE]", "'" + Dr["code"].ToString() + "'");


                    //sqlacc = sqlacc.Replace("[ACC_NAME]", "'" + Dr["name"].ToString() + "'");
                    sqlacc = sqlacc.Replace("[ACC_NAME]", "'" + sName1 + "'");


                    sqlacc = sqlacc.Replace("[ACC_GROUP_ID]", "'" + GetByName("ACGROUP", Dr["type"].ToString()) + "'");

                    sqlacc = sqlacc.Replace("[ACC_TYPE_ID]", "'" + GetByName("ACTYPE", Dr["subtype"].ToString()) + "'");

                    sqlacc = sqlacc.Replace("[ACC_AGAINST_INVOICE]", "'N'");

                    if (Dr["maincode"].ToString() == Dr["code"].ToString())
                    {
                        sqlacc = sqlacc.Replace("[ACC_MAIN_ID]", "'" + Dr["row_pkid"].ToString() + "'");
                        sqlacc = sqlacc.Replace("[ACC_MAIN_CODE]", "'" + Dr["maincode"].ToString() + "'");
                        sqlacc = sqlacc.Replace("[ACC_MAIN_NAME]", "'" + sName1 + "'");
                    }
                    else
                    {
                        sqlacc = sqlacc.Replace("[ACC_MAIN_ID]", "'" + GetByCode("ACCOUNTS MAIN CODE", Dr["maincode"].ToString()) + "'");
                        sqlacc = sqlacc.Replace("[ACC_MAIN_CODE]", "'" + Dr["maincode"].ToString() + "'");
                        sqlacc = sqlacc.Replace("[ACC_MAIN_NAME]", "'" + GetByCode("ACCOUNTS MAIN CODE", Dr["maincode"].ToString(), "param_name") + "'");
                    }


                    sqlacc = sqlacc.Replace("[ACC_BRANCH_CODE]", "'" + Dr["branch_code"].ToString() + "'");

                    if (Dr["cc"].ToString() == "Y")
                        sqlacc = sqlacc.Replace("[ACC_COST_CENTRE]", "'Y'");
                    else
                        sqlacc = sqlacc.Replace("[ACC_COST_CENTRE]", "'N'");

                    nRate = Lib.Conv2Decimal(Dr["rate"].ToString());

                    if (nRate > 0)
                    {
                        sqlacc = sqlacc.Replace("[ACC_IGST_RATE]", nRate.ToString());
                        nRate = nRate / 2;
                        sqlacc = sqlacc.Replace("[ACC_CGST_RATE]", nRate.ToString());
                        sqlacc = sqlacc.Replace("[ACC_SGST_RATE]", nRate.ToString());
                        sqlacc = sqlacc.Replace("[ACC_TAXABLE]", "'Y'");
                    }
                    else
                    {
                        sqlacc = sqlacc.Replace("[ACC_IGST_RATE]", "0");
                        sqlacc = sqlacc.Replace("[ACC_CGST_RATE]", "0");
                        sqlacc = sqlacc.Replace("[ACC_SGST_RATE]", "0");
                        sqlacc = sqlacc.Replace("[ACC_TAXABLE]", "'N'");
                    }

                    sqlacc = sqlacc.Replace("[ACC_SAC_ID]", "'" + GetByCode("SAC", Dr["sac"].ToString()) + "'");


                    sqlacc = sqlacc.Replace("[REC_LOCKED]", "'N'");
                    sqlacc = sqlacc.Replace("[REC_COMPANY_CODE]", "'" + sCOMPANY_CODE + "'");
                    sqlacc = sqlacc.Replace("[REC_BRANCH_CODE]", "'" + sBRANCH_CODE + "'");
                    sqlacc = sqlacc.Replace("[REC_CREATED_BY]", "'" + sUSER + "'");
                    sqlacc = sqlacc.Replace("[REC_CREATED_DATE]", "sysdate");

                    ErrName = Dr["code"].ToString() + "~" + Dr["name"].ToString() + "~";

                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sqlacc);
                    sql = "update " + tablename + " set row_status='Y' where row_pkid='" + Dr["row_pkid"].ToString() + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    RecCount++;
                }
                catch (Exception Ex)
                {
                    Con_Oracle.CreateErrorLog(ErrName + Ex.Message.ToString());
                    Con_Oracle.RollbackTransaction();
                }
            }
            Con_Oracle.CloseConnection();
            return RecCount;
        }

        private int InsertBranches(string tablename, string paramtype, string sCOMPANY_CODE, string sBRANCH_CODE, string sUSER)
        {
            int RecCount = 0;
            sql = "select * from " + tablename + "  where  nvl(row_status,'N') !='Y' order by name ";
            DataTable Dt_Table = Con_Oracle.ExecuteQuery(sql);

            string sqlacc = "";
            string CODE = "";
            int iCtr = 0;


            DataTable dt_comp = new DataTable();

            int iDupTot = 0;

            string ErrName = "";
            string sName = "";
            string sName1 = "";
            foreach (DataRow Dr in Dt_Table.Rows)
            {
                try
                {
                    sql = "select comp_pkid from companym where comp_type ='C' and rec_company_code = '" + Dr["comp_code"].ToString() + "'";
                    dt_comp = Con_Oracle.ExecuteQuery(sql);


                    if (sName != Dr["name"].ToString())
                        iDupTot = 0;
                    else
                        iDupTot++;
                    sName = Dr["name"].ToString();
                    sName1 = sName;
                    if (iDupTot > 0)
                        sName1 = sName.Replace(" ", " ".PadLeft(iDupTot + 1, ' '));

                    iCtr++;

                    sql = "";
                    sql += " insert into companym (COMP_PKID,COMP_TYPE,REC_COMPANY_CODE,REC_BRANCH_CODE,COMP_PARENT_ID,COMP_CODE,COMP_NAME,COMP_ADDRESS1,COMP_ADDRESS2,COMP_ADDRESS3,COMP_TEL,COMP_FAX,COMP_EMAIL,COMP_WEB,COMP_PTC,COMP_MOBILE,COMP_PREFIX,COMP_PANNO,COMP_CINNO,COMP_GSTIN,COMP_REG_ADDRESS,COMP_IATA_CODE,COMP_LOCATION,COMP_ORDER,REC_UPDATED,REC_LOCKED) ";
                    sql += " values ( ";
                    sql += " [COMP_PKID],[COMP_TYPE],[REC_COMPANY_CODE],[REC_BRANCH_CODE],[COMP_PARENT_ID],[COMP_CODE],[COMP_NAME],[COMP_ADDRESS1],[COMP_ADDRESS2],[COMP_ADDRESS3],[COMP_TEL],[COMP_FAX],[COMP_EMAIL],[COMP_WEB],[COMP_PTC],[COMP_MOBILE],[COMP_PREFIX],[COMP_PANNO],[COMP_CINNO],[COMP_GSTIN],[COMP_REG_ADDRESS],[COMP_IATA_CODE],[COMP_LOCATION],[COMP_ORDER],[REC_UPDATED],[REC_LOCKED] ";
                    sql += ")";


                    sql = sql.Replace("[COMP_PKID]", "'" + Dr["row_pkid"].ToString() + "'");
                    sql = sql.Replace("[COMP_TYPE]", "'B'");

                    sql = sql.Replace("[REC_COMPANY_CODE]", "'" + Dr["comp_code"].ToString() + "'");
                    sql = sql.Replace("[REC_BRANCH_CODE]", "'" + Dr["code"].ToString() + "'");

                    sql = sql.Replace("[COMP_PARENT_ID]", "'" + dt_comp.Rows[0]["comp_pkid"].ToString() + "'");

                    sql = sql.Replace("[COMP_CODE]", "'" + Dr["code"].ToString() + "'");
                    sql = sql.Replace("[COMP_NAME]", "'" + Dr["name"].ToString() + "'");
                    sql = sql.Replace("[COMP_ADDRESS1]", "'" + Dr["address1"].ToString() + "'");
                    sql = sql.Replace("[COMP_ADDRESS2]", "'" + Dr["address2"].ToString() + "'");
                    sql = sql.Replace("[COMP_ADDRESS3]", "'" + Dr["address3"].ToString() + "'");

                    sql = sql.Replace("[COMP_TEL]", "'" + Dr["tel"].ToString() + "'");
                    sql = sql.Replace("[COMP_FAX]", "'" + Dr["fax"].ToString() + "'");
                    sql = sql.Replace("[COMP_EMAIL]", "'" + Dr["email"].ToString() + "'");
                    sql = sql.Replace("[COMP_WEB]", "'" + Dr["web"].ToString() + "'");

                    sql = sql.Replace("[COMP_PTC]", "'" + Dr["ptc"].ToString() + "'");
                    sql = sql.Replace("[COMP_MOBILE]", "'" + Dr["mobile"].ToString() + "'");
                    sql = sql.Replace("[COMP_PREFIX]", "'" + Dr["prefix"].ToString() + "'");

                    sql = sql.Replace("[COMP_PANNO]", "'" + Dr["pan"].ToString() + "'");
                    sql = sql.Replace("[COMP_CINNO]", "'" + Dr["cin"].ToString() + "'");
                    sql = sql.Replace("[COMP_GSTIN]", "'" + Dr["gstin"].ToString() + "'");
                    sql = sql.Replace("[COMP_REG_ADDRESS]", "'" + Dr["regaddress"].ToString() + "'");
                    sql = sql.Replace("[COMP_IATA_CODE]", "''");
                    sql = sql.Replace("[COMP_LOCATION]", "'" + Dr["location"].ToString() + "'");
                    sql = sql.Replace("[COMP_ORDER]", iCtr.ToString());


                    sql = sql.Replace("[REC_LOCKED]", "'N'");
                    sql = sql.Replace("[REC_UPDATED]", "'N'");

                    ErrName = Dr["code"].ToString() + "~" + Dr["name"].ToString() + "~";

                    Con_Oracle.BeginTransaction();
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = "update " + tablename + " set row_status='Y' where row_pkid='" + Dr["row_pkid"].ToString() + "'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    Con_Oracle.CommitTransaction();
                    RecCount++;
                }
                catch (Exception Ex)
                {
                    Con_Oracle.CreateErrorLog(ErrName + Ex.Message.ToString());
                    Con_Oracle.RollbackTransaction();
                }
            }
            Con_Oracle.CloseConnection();
            return RecCount;
        }


        



        private string getLinkId(string source_name, string source_table, string company_code, string branch_code)
        {
            string LinkId = "";

            string sql = "select targetid from linkm ";
            sql += " where rec_company_code = '" + company_code + "'";
            sql += " and rec_branch_code = '" + branch_code + "'";
            if (source_table == "DT_SHIPPER_OPENING" || source_table == "DT_OPENING")
                sql += " and sourcetable in ('DT_SHIPPER_OPENING','DT_OPENING')";
            else
                sql += " and sourcetable = '" + source_table + "'";
            sql += " and name = '" + source_name + "'";
            DataTable Dt_Link = new DataTable();
            Dt_Link = Con_Oracle.ExecuteQuery(sql);
            if (Dt_Link.Rows.Count > 0)
                LinkId = Dt_Link.Rows[0]["targetid"].ToString();
            return LinkId;
        }

        string getData(DataRow Dr, string ColName)
        {
            if (Dr.Table.Columns.Contains(ColName))
                return Dr[ColName].ToString();
            else
                return "";
        }

        private string GetByCode(string sType, string sdesc, string retfld = "param_pkid")
        {
            string sdata = "";
            foreach (DataRow Dr in Dt_param.Rows)
            {
                if (Dr["param_type"].ToString() == sType && Dr["param_code"].ToString() == sdesc)
                {
                    sdata = Dr[retfld].ToString();
                    break;
                }
            }
            return sdata;
        }

        private string GetByName(string sType, string sdesc, string retfld = "param_pkid")
        {
            string sdata = "";
            foreach (DataRow Dr in Dt_param.Rows)
            {
                if (Dr["param_type"].ToString() == sType && Dr["param_name"].ToString() == sdesc)
                {
                    sdata = Dr[retfld].ToString();
                    break;
                }
            }
            return sdata;
        }

        private void SetParamData(string CompanyCode)
        {
            string sql = "select param_type,param_pkid, param_code, param_name from param ";
            sql += " where rec_company_code = '" + CompanyCode + "' and param_type in('COUNTRY','STATE', 'SALESMAN','ACCOUNTS MAIN CODE', 'SAC','CURRENCY')";
            sql += " union all ";
            sql += "select cast('ACGROUP' as nvarchar2(50)) ,acgrp_pkid, acgrp_name,acgrp_name from acgroupm where rec_company_code = '" + CompanyCode + "' and acgrp_parent_id is not null";
            sql += " union all ";
            sql += "select cast('ACTYPE' as nvarchar2(50)) ,actype_pkid, actype_name,actype_name from actypem where rec_company_code = '" + CompanyCode + "'";
            Dt_param = Con_Oracle.ExecuteQuery(sql);
        }

        public static string ReplaceFirstOccurrence(string Source, string Find, string Replace)
        {
            int Place = Source.IndexOf(Find);
            string result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
            return result;
        }

        private string GetyyyyMMddDate(string sDate, char sepChar = '/', string sFmt = "")
        {
            if (sDate.Contains(sepChar))
            {
                string[] thisdate = sDate.Split(sepChar);
                if (sFmt == "mmddyyyy")
                    sDate = thisdate[2] + "-" + thisdate[0] + "-" + thisdate[1];
                else
                {
                    sDate = thisdate[2] + "-" + thisdate[1] + "-" + thisdate[0];
                }
            }
            return sDate;
        }


        private string GetyyyyMMddDatenew(string sDate, string brcode)
        {

            string[] thisdate = sDate.Split('-');

            if ( brcode  == "CHNSF" || brcode == "CHNAF")
                sDate = thisdate[2] + "-" + thisdate[0] + "-" + thisdate[1];
            else 
                sDate = thisdate[2] + "-" + thisdate[1] + "-" + thisdate[0];

            return sDate;
        }







        public IDictionary<string, object> UpdateData(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            string type = "";
            string subtype = "";
            string comp_code = "";
            string branch_code = "";
            string year_code = "";

            string status = "";

            Boolean ismaster = false;

            int no1 = 0;
            int no2 = 0;


            DataTable Dt_test = null;
            DataTable Dt_House = null;

            try
            {

                if (SearchData.ContainsKey("type"))
                    type = SearchData["type"].ToString();
                if (SearchData.ContainsKey("subtype"))
                    subtype = SearchData["subtype"].ToString();
                if (SearchData.ContainsKey("comp_code"))
                    comp_code = SearchData["comp_code"].ToString();
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();
                if (SearchData.ContainsKey("no1"))
                    no1 = Lib.Conv2Integer(SearchData["no1"].ToString());
                if (SearchData.ContainsKey("no2"))
                    no2 = Lib.Conv2Integer(SearchData["no2"].ToString());

                if (SearchData.ContainsKey("year_code"))
                    year_code = SearchData["year_code"].ToString();


                if (type == "UPDATENARRATION")
                {
                    Con_Oracle = new DBConnection();

                    sql += " select jvh_pkid, jvh_cc_category, a.hbl_type,a.hbl_job_nos as jobnos,a.hbl_no as SINO,a.hbl_bl_no as hbl,b.hbl_bl_no as mbl, a.hbl_book_cntr as cntr,";
                    sql += " c.acc_name as billto,d.acc_name as shipper    ";
                    sql += " from ledgerh h  ";
                    sql += " inner join hblm a on jvh_cc_id = a.hbl_pkid ";
                    sql += " left join hblm b on a.hbl_mbl_id = b.hbl_pkid";
                    sql += " left join acctm c on jvh_acc_id = c.acc_pkid";
                    sql += " left join acctm d on a.hbl_exp_id = d.acc_pkid";
                    sql += " where  h.rec_company_code = '{COMP_CODE}' and  h.rec_branch_code = '{BRANCH_CODE}' and h.jvh_type ='IN' and jvh_cc_category  in ('SI AIR EXPORT', 'SI SEA EXPORT')";
                    sql += " union all";
                    sql += " select  jvh_pkid, jvh_cc_category, a.hbl_type,a.hbl_job_nos as jobnos,a.hbl_no as SINO,a.hbl_bl_no as hbl,b.hbl_bl_no as mbl, a.hbl_book_cntr as cntr,";
                    sql += " c.acc_name as billto,d.acc_name as shipper    ";
                    sql += " from ledgerh h  ";
                    sql += " inner join hblm a on jvh_cc_id = a.hbl_pkid ";
                    sql += " left join hblm b on a.hbl_mbl_id = b.hbl_pkid";
                    sql += " left join acctm c on jvh_acc_id = c.acc_pkid";
                    sql += " left join acctm d on a.hbl_exp_id = d.acc_pkid";
                    sql += " where  h.rec_company_code = '{COMP_CODE}' and  h.rec_branch_code = '{BRANCH_CODE}' and h.jvh_type ='IN' and jvh_cc_category  in ('SI AIR IMPORT', 'SI SEA IMPORT')";
                    sql += " union all";
                    sql += " select  jvh_pkid, jvh_cc_category, a.hbl_type,a.hbl_job_nos as jobnos,a.hbl_no as SINO,a.hbl_bl_no as hbl,b.hbl_bl_no as mbl, a.hbl_book_cntr as cntr,";
                    sql += " c.acc_name as billto,null as shipper    ";
                    sql += " from ledgerh h  ";
                    sql += " inner join hblm a on jvh_cc_id = a.hbl_pkid ";
                    sql += " left join hblm b on a.hbl_mbl_id = b.hbl_pkid";
                    sql += " left join acctm c on jvh_acc_id = c.acc_pkid";
                    sql += " where  h.rec_company_code = '{COMP_CODE}' and  h.rec_branch_code = '{BRANCH_CODE}' and h.jvh_type ='IN' and jvh_cc_category  = 'GENERAL JOB' and nvl(jvh_version,0) <> 10 ";

                    sql = sql.Replace("{COMP_CODE}", comp_code);
                    sql = sql.Replace("{BRANCH_CODE}", branch_code);


                    Dt_test = new DataTable();
                    Dt_test = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_test.Rows)
                    {
                        if (Dr["hbl_type"].ToString() == "HBL-SE")
                        {
                            sql = Dr["BILLTO"].ToString();
                            if (Dr["BILLTO"].ToString() != Dr["SHIPPER"].ToString())
                                sql += " ON A/C OF " + Dr["BILLTO"].ToString();
                            sql += " AGAINST job# " + Dr["JOBNOS"].ToString() + ",";
                            sql += " si# " + Dr["SINO"].ToString() + ",";
                            sql += " MBL# " + Dr["mbl"].ToString() + ",";
                            sql += " HBL# " + Dr["hbl"].ToString() + ",";
                            sql += " Cntr# " + Dr["cntr"].ToString();
                        }
                        if (Dr["hbl_type"].ToString() == "HBL-AE")
                        {
                            sql = Dr["BILLTO"].ToString();
                            if (Dr["BILLTO"].ToString() != Dr["SHIPPER"].ToString())
                                sql += " ON A/C OF " + Dr["BILLTO"].ToString();
                            sql += " AGAINST job# " + Dr["JOBNOS"].ToString() + ",";
                            sql += " si# " + Dr["SINO"].ToString() + ",";
                            sql += " MAWB# " + Dr["mbl"].ToString() + ",";
                            sql += " HAWB# " + Dr["hbl"].ToString();
                        }

                        if (Dr["hbl_type"].ToString() == "HBL-SI")
                        {
                            sql = Dr["BILLTO"].ToString();
                            if (Dr["BILLTO"].ToString() != Dr["SHIPPER"].ToString())
                                sql += " ON A/C OF " + Dr["BILLTO"].ToString();
                            sql += " AGAINST si# " + Dr["SINO"].ToString() + ",";
                            sql += " MBL# " + Dr["mbl"].ToString() + ",";
                            sql += " HBL# " + Dr["hbl"].ToString() + ",";
                            sql += " Cntr# " + Dr["cntr"].ToString();
                        }
                        if (Dr["hbl_type"].ToString() == "HBL-AI")
                        {
                            sql = Dr["BILLTO"].ToString();
                            if (Dr["BILLTO"].ToString() != Dr["SHIPPER"].ToString())
                                sql += " ON A/C OF " + Dr["BILLTO"].ToString();
                            sql += " AGAINST SI# " + Dr["SINO"].ToString() + ",";
                            sql += " MAWB# " + Dr["mbl"].ToString() + ",";
                            sql += " HAWB# " + Dr["hbl"].ToString();
                        }


                        if (Dr["hbl_type"].ToString() == "JOB-GN")
                        {
                            sql = Dr["BILLTO"].ToString();
                            sql += " AGAINST JOB# " + Dr["SINO"].ToString() + ",";
                            if (!Dr["cntr"].Equals(DBNull.Value))
                                sql += " CNTR# " + Dr["cntr"].ToString().Replace("'","");
                        }

                        sql = sql.ToUpper();
                        if (sql.Length > 255)
                            sql = sql.Substring(0, 255);

                        sql = " update ledgerh set jvh_narration = '" + sql + "' where jvh_pkid = '" + Dr["jvh_pkid"].ToString() + "'";

                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CommitTransaction();
                    }
                    Dt_test.Rows.Clear();


                    sql = " select jvh_pkid, jvh_cc_category, a.hbl_type,a.hbl_job_nos as sino, a.hbl_bl_no as mbl, a.hbl_book_cntr as cntr, ";
                    sql += " c.acc_name as billto ";
                    sql += " from ledgerh h ";
                    sql += " inner join hblm a on jvh_cc_id = a.hbl_pkid ";
                    sql += " left join acctm c on jvh_acc_id = c.acc_pkid ";
                    sql += " where h.rec_company_code = '{COMP_CODE}' and  h.rec_branch_code = '{BRANCH_CODE}' and h.jvh_type = 'PN' ";

                    sql = sql.Replace("{COMP_CODE}", comp_code);
                    sql = sql.Replace("{BRANCH_CODE}", branch_code);


                    Dt_test = new DataTable();
                    Dt_test = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();

                    foreach (DataRow Dr in Dt_test.Rows)
                    {

                        if (Dr["hbl_type"].ToString() == "MBL-SE" || Dr["hbl_type"].ToString() == "MBL-SI")
                        {
                            sql = "PAYABLE TO " + Dr["BILLTO"].ToString();
                            sql += " AGAINST ";
                            sql += " SI# " + Dr["SINO"].ToString() + ",";
                            sql += " MBL# " + Dr["mbl"].ToString() + ",";
                            sql += " Cntr# " + Dr["cntr"].ToString();
                        }
                        else if (Dr["hbl_type"].ToString() == "MBL-AE" ||  Dr["hbl_type"].ToString() == "MBL-AI")
                        {
                            sql = "PAYABLE TO " + Dr["BILLTO"].ToString();
                            sql += " AGAINST ";
                            sql += " SI# " + Dr["SINO"].ToString() + ",";
                            sql += " MAWB# " + Dr["mbl"].ToString() + "";
                        }
                        else
                        {
                            sql = Dr["BILLTO"].ToString();
                            sql += " AGAINST ";
                            sql += " JOB# " + Dr["SINO"].ToString() + ",";
                            if (!Dr["cntr"].Equals(DBNull.Value))
                                sql += " CNTR# " + Dr["cntr"].ToString();
                        }
                        sql = sql.ToUpper();
                        if (sql.Length > 255)
                            sql = sql.Substring(0, 255);


                        sql = " update ledgerh set jvh_narration = '" + sql + "' where jvh_pkid = '" + Dr["jvh_pkid"].ToString() + "'";

                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteQuery(sql);
                        Con_Oracle.CommitTransaction();
                    }
                    Con_Oracle.CloseConnection();
                    Dt_test.Rows.Clear();

                    status = "ok";
                }

            }
            catch (Exception Ex)
            {
                status = "erorr";
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            RetData.Add("result", status);
            return RetData;
        }


        

        public Dictionary<string, object> Process(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            string filename = "";
            string serror = "";
            string year_code = "";
            string sMsg = "";
            string user_code = "";
            string comp_code = "";
            Con_Oracle = new DBConnection();
            try
            {
                comp_code = "";
                if (SearchData.ContainsKey("comp_code"))
                    comp_code = SearchData["comp_code"].ToString();

                year_code = "";
                if (SearchData.ContainsKey("year_code"))
                    year_code = SearchData["year_code"].ToString();

                user_code = "";
                if (SearchData.ContainsKey("user_code"))
                    user_code = SearchData["user_code"].ToString();

                sql = "select doc_path,doc_file_name,rec_created_date from documentm where doc_parent_id='BCFAFC91-B127-4D7F-908A-0C7180D1345F' and rec_deleted='N'";
                DataTable Dt_Rec = new DataTable();
               
                Dt_Rec = Con_Oracle.ExecuteQuery(sql);
          
                if (Dt_Rec.Rows.Count <= 0)
                {
                    serror = "No Files to Process";
                }
                if (Dt_Rec.Rows.Count > 1)
                {
                    serror = "More than one File Found to Process";
                }

                if (serror == "")
                {
                    filename = @"D://documents/" + Dt_Rec.Rows[0]["doc_path"].ToString() + "/" + Dt_Rec.Rows[0]["doc_file_name"].ToString();
                    if (!Path.GetExtension(filename).ToUpper().Contains("TXT"))
                    {
                        serror = "Please Upload a Text file and continue....";
                    }
                }
                if (serror == "")
                {
                    Process26AS(filename, year_code, user_code, comp_code, ref sMsg);
                }
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }

            Con_Oracle.CloseConnection();

            RetData.Add("serror", serror);
            RetData.Add("smsg", sMsg);
            RetData.Add("filename", filename);
            return RetData;
        }
        private void Process26AS(string FullPathFileName, string FinYrCode, string user_code,string comp_code, ref string Msg)
        {
            string MissingMastersSL = "";
            string SQL = "";
            string SQL_26ASM = "";
            string SQL_26ASD = "";

            try
            {

                SQL_26ASM = " INSERT INTO TDS26ASM ( ";
                SQL_26ASM += "    ASM_PKID ";
                SQL_26ASM += "   ,ASM_SLNO";
                SQL_26ASM += "   ,ASM_YEAR";
                SQL_26ASM += "   ,ASM_TAN";
                SQL_26ASM += "   ,ASM_TAN_NAME";
                SQL_26ASM += "   ,ASM_GROSS ";
                SQL_26ASM += "   ,ASM_DEDUCTED ";
                SQL_26ASM += "   ,ASM_TDS ";
                SQL_26ASM += "   ,REC_CREATED_BY,REC_CREATED_DATE ";
                SQL_26ASM += "   ,REC_COMPANY_CODE ";
                SQL_26ASM += "   ) VALUES (";
                SQL_26ASM += "    [ASM_PKID]";
                SQL_26ASM += "   ,[ASM_SLNO]";
                SQL_26ASM += "   ,[ASM_YEAR]";
                SQL_26ASM += "   ,[ASM_TAN]";
                SQL_26ASM += "   ,[ASM_TAN_NAME]";
                SQL_26ASM += "   ,[ASM_GROSS]";
                SQL_26ASM += "   ,[ASM_DEDUCTED]";
                SQL_26ASM += "   ,[ASM_TDS]";
                SQL_26ASM += "   ,[REC_CREATED_BY],[REC_CREATED_DATE]";
                SQL_26ASM += "   ,[REC_COMPANY_CODE]";
                SQL_26ASM += "   )";


                SQL_26ASD = " INSERT INTO TDS26ASD ( ";
                SQL_26ASD += "    ASD_PKID ";
                SQL_26ASD += "   ,ASD_PARENT_ID";
                SQL_26ASD += "   ,ASD_SLNO";
                SQL_26ASD += "   ,ASD_SECTION";
                SQL_26ASD += "   ,ASD_TRANS_DATE ";
                SQL_26ASD += "   ,ASD_BOOK_DATE ";
                SQL_26ASD += "   ,ASD_GROSS ";
                SQL_26ASD += "   ,ASD_DEDUCTED ";
                SQL_26ASD += "   ,ASD_TDS ";
                SQL_26ASD += "   ) VALUES (";
                SQL_26ASD += "    [ASD_PKID]";
                SQL_26ASD += "   ,[ASD_PARENT_ID]";
                SQL_26ASD += "   ,[ASD_SLNO]";
                SQL_26ASD += "   ,[ASD_SECTION]";
                SQL_26ASD += "   ,[ASD_TRANS_DATE]";
                SQL_26ASD += "   ,[ASD_BOOK_DATE]";
                SQL_26ASD += "   ,[ASD_GROSS]";
                SQL_26ASD += "   ,[ASD_DEDUCTED]";
                SQL_26ASD += "   ,[ASD_TDS]";
                SQL_26ASD += "   )";


                StreamReader reader = new StreamReader(FullPathFileName);

                string sMasterRefNo = "";
                string sCellData = "";
                decimal Tds_GrAmt = 0;
                decimal Tds_DeductAmt = 0;
                decimal Tds_Amt = 0;
                int _BlankRows = 0;
                string ASM_PKID = "";
                int ASD_SLNO = 0;
                bool IsMasterExist = false;
                int SavedCount = 0;
                MissingMastersSL = "";

                Con_Oracle.BeginTransaction();
                SQL = "delete from tds26asd where asd_parent_id in (select asm_pkid from tds26asm where asm_year=" + FinYrCode.ToString() + ")";
                Con_Oracle.ExecuteNonQuery(SQL);
                SQL = "delete from tds26asm where asm_year=" + FinYrCode.ToString();
                Con_Oracle.ExecuteNonQuery(SQL);
                Con_Oracle.CommitTransaction();

                ASD_SLNO = 0;
                SavedCount = 0;

                int iRow = 0;
                string sline = null;
                string[] sdata = null;
                while ((sline = reader.ReadLine()) != null)
                {
                    iRow++;
                    sline = sline.Trim();
                    sdata = sline.Split(',');

                    SQL = ""; Tds_GrAmt = 0; Tds_DeductAmt = 0; Tds_Amt = 0;
                    if (sdata.Length > 0)
                        sCellData = sdata[0];
                    else
                        sCellData = "";


                    sMasterRefNo = sCellData;

                    if (sCellData.Trim().Length > 0 && Lib.Conv2Integer(sCellData) > 0)
                    {
                        //Master
                        ASD_SLNO = 0;
                        IsMasterExist = false;
                        ASM_PKID = Guid.NewGuid().ToString().ToUpper();

                        if (sdata.Length > 5)
                            Tds_GrAmt = Lib.Convert2Decimal(sdata[5]);
                        else
                            Tds_GrAmt = 0;
                        if (sdata.Length > 6)
                            Tds_DeductAmt = Lib.Convert2Decimal(sdata[6]);
                        else
                            Tds_DeductAmt = 0;
                        if (sdata.Length > 7)
                            Tds_Amt = Lib.Convert2Decimal(sdata[7]);
                        else
                            Tds_Amt = 0;

                        if (Tds_GrAmt != 0 || Tds_Amt != 0)
                        {
                            sMasterRefNo = "";//to identify missing master recs while Inserting
                            _BlankRows = 0;
                            IsMasterExist = true;

                            SQL = SQL_26ASM;
                            SQL = SQL.Replace("[ASM_PKID]", "'" + ASM_PKID + "'");

                            if (sdata.Length > 0)
                                SQL = SQL.Replace("[ASM_SLNO]", Lib.Conv2Integer(sdata[0]).ToString());
                            else
                                SQL = SQL.Replace("[ASM_SLNO]", "0");
                            SQL = SQL.Replace("[ASM_YEAR]", FinYrCode.ToString());

                            if (sdata.Length > 1)
                                SQL = SQL.Replace("[ASM_TAN_NAME]", "'" + sdata[1].Replace("'", "''") + "'");
                            else
                                SQL = SQL.Replace("[ASM_TAN_NAME]", "'" + "" + "'");

                            if (sdata.Length > 2)
                                SQL = SQL.Replace("[ASM_TAN]", "'" + sdata[2] + "'");
                            else
                                SQL = SQL.Replace("[ASM_TAN]", "'" + "" + "'");

                            SQL = SQL.Replace("[ASM_GROSS]", Tds_GrAmt.ToString());
                            SQL = SQL.Replace("[ASM_DEDUCTED]", Tds_DeductAmt.ToString());
                            SQL = SQL.Replace("[ASM_TDS]", Tds_Amt.ToString());
                            SQL = SQL.Replace("[REC_CREATED_BY]", "'" + user_code + "'");
                            SQL = SQL.Replace("[REC_CREATED_DATE]", "sysdate");
                            SQL = SQL.Replace("[REC_COMPANY_CODE]", "'" + comp_code + "'");
                        }

                        _BlankRows++;
                    }
                    else
                    {

                        if (sdata.Length > 2)
                            if (sdata[2].ToUpper() == "SECTION")
                                continue;
                        ASD_SLNO++;

                        if (sdata.Length > 5)
                            Tds_GrAmt = Lib.Convert2Decimal(sdata[5]);
                        else
                            Tds_GrAmt = 0;
                        if (sdata.Length > 6)
                            Tds_DeductAmt = Lib.Convert2Decimal(sdata[6]);
                        else
                            Tds_DeductAmt = 0;
                        if (sdata.Length > 7)
                            Tds_Amt = Lib.Convert2Decimal(sdata[7]);
                        else
                            Tds_Amt = 0;

                        if (IsMasterExist && (Tds_GrAmt != 0 || Tds_Amt != 0))
                        {
                            _BlankRows = 0;

                            SQL = SQL_26ASD;
                            SQL = SQL.Replace("[ASD_PKID]", "'" + Guid.NewGuid().ToString().ToUpper() + "'");
                            SQL = SQL.Replace("[ASD_PARENT_ID]", "'" + ASM_PKID + "'");
                            SQL = SQL.Replace("[ASD_SLNO]", ASD_SLNO.ToString());
                            if (sdata.Length > 2)
                                SQL = SQL.Replace("[ASD_SECTION]", "'" + sdata[2] + "'");
                            else
                                SQL = SQL.Replace("[ASD_SECTION]", "'" + "" + "'");

                            string ThisDate = "";
                            if (sdata.Length > 8)
                                ThisDate = sdata[8];
                            if (ThisDate.Trim().Length > 0)
                                SQL = SQL.Replace("[ASD_TRANS_DATE]", "'" + Lib.StringToDate(ThisDate, "DD-MM-YYYY") + "'");
                            else
                                SQL = SQL.Replace("[ASD_TRANS_DATE]", "'" + DBNull.Value + "'");

                            ThisDate = "";
                            if (sdata.Length > 9)
                                ThisDate = sdata[9];
                            if (ThisDate.Trim().Length > 0)
                                SQL = SQL.Replace("[ASD_BOOK_DATE]", "'" + Lib.StringToDate(ThisDate, "DD-MM-YYYY") + "'");
                            else
                                SQL = SQL.Replace("[ASD_BOOK_DATE]", "'" + DBNull.Value + "'");
                            SQL = SQL.Replace("[ASD_GROSS]", Tds_GrAmt.ToString());
                            SQL = SQL.Replace("[ASD_DEDUCTED]", Tds_DeductAmt.ToString());
                            SQL = SQL.Replace("[ASD_TDS]", Tds_Amt.ToString());
                        }

                        _BlankRows++;
                    }

                    if (sMasterRefNo.Trim() != "")
                    {
                        if (MissingMastersSL.Trim() != "")
                            MissingMastersSL += ", ";
                        MissingMastersSL += sMasterRefNo;
                    }

                    if (SQL.Trim() != "")
                    {
                        SavedCount++;
                        Con_Oracle.BeginTransaction();
                        Con_Oracle.ExecuteNonQuery(SQL);
                        Con_Oracle.CommitTransaction();
                    }

                    //Lbl_Referesh.Text = "Rows : " + iRow.ToString();
                    //Lbl_Referesh.Tag = "Total Rows : " + iRow.ToString() + ", Saved Rows : " + SavedCount.ToString() + (MissingMastersSL.Trim() != "" ? "\n MissingMaster Sr. Nos. " + MissingMastersSL : "");
                    //Lbl_Referesh.Refresh();
                }

                Msg = "\n Total Rows : " + iRow.ToString() + ", Saved Rows : " + SavedCount.ToString() + (MissingMastersSL.Trim() != "" ? "\n MissingMaster Sr. Nos. " + MissingMastersSL : "");
                if (reader != null)
                    reader.Close();
            }
            catch (Exception ex)
            {
                Con_Oracle.CreateErrorLog("TDS26AS-" + ex.Message.ToString());
                Con_Oracle.RollbackTransaction();
            }
        }
    }
}
