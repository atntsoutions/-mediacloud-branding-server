using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

using XL.XSheet;

namespace BLReport1
{
    public class AuditLogService : BL_Base
    {
        ExcelFile WB;
        ExcelWorksheet WS = null;
        int iRow = 0;
        int iCol = 0;
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string report_folder = "";
        string folderid = "";
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            List<Auditlog> mList = new List<Auditlog>();
            Auditlog mRow = null;
            string ErrorMessage = "";
            long page_count = (long)SearchData["page_count"];
            long page_current = (long)SearchData["page_current"];
            long page_rows = (long)SearchData["page_rows"];
            long page_rowcount = (long)SearchData["page_rowcount"];
            long startrow = 0;
            long endrow = 0;
            try
            {
                DataTable Dt_List = null; 
                Con_Oracle = new DBConnection();
                string type = SearchData["type"].ToString();
                string company_code = SearchData["comp_code"].ToString();
                string branch_code = SearchData["branch_code"].ToString();
                string year_code = SearchData["year_code"].ToString();
                string from_date = SearchData["from_date"].ToString();
                string to_date = SearchData["to_date"].ToString();
                string searchstring = SearchData["searchstring"].ToString().ToUpper();
                string searchuser = SearchData["searchuser"].ToString().ToUpper();
                string searchtype = SearchData["searchtype"].ToString().ToUpper();
                string searchmodule = SearchData["searchmodule"].ToString().ToUpper();
                string searchbranch = SearchData["searchbranch"].ToString().ToUpper();
                string searchaction = SearchData["searchaction"].ToString().ToUpper();
                string searchremarks = SearchData["searchremarks"].ToString().ToUpper();
                report_folder = SearchData["report_folder"].ToString();
                folderid = SearchData["pkid"].ToString();

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);


                string sWhere = "";
                sWhere = " where 1=1 ";
                sWhere += " and a.audit_comp_code = '{COMPCODE}'";
                if (searchuser != "")
                    sWhere += " and  a.audit_user_code  like '%" + searchuser + "%'";

                if (searchtype != "")
                    sWhere += " and  a.audit_type  like '%" + searchtype + "%'";

                if (searchmodule != "")
                    sWhere += " and  a.audit_module  like '%" + searchmodule + "%'";

                if (searchbranch != "")
                    sWhere += " and  a.audit_branch_code  like '%" + searchbranch + "%'";

                if (searchaction != "")
                    sWhere += " and  a.audit_action  like '%" + searchaction + "%'";

                if (searchremarks != "")
                    sWhere += " and  a.audit_remarks  like '%" + searchremarks + "%'";

                
                if (from_date != "NULL")
                    sWhere += "  and to_date(to_char(a.audit_date,'DD-MON-YYYY'),'DD-MON-YYYY') >= '{FDATE}' ";
                if (to_date != "NULL")
                    sWhere += "  and to_date(to_char(a.audit_date,'DD-MON-YYYY'),'DD-MON-YYYY') <= '{EDATE}' ";


                sWhere = sWhere.Replace("{COMPCODE}", company_code);
                sWhere = sWhere.Replace("{BRCODE}", branch_code);
                sWhere = sWhere.Replace("{FDATE}", from_date);
                sWhere = sWhere.Replace("{EDATE}", to_date);


                if (type == "EXCEL")
                {
                    sql = "";
                    sql = "";
                    sql += " select to_char(audit_date,'DD/MM/YYYY HH24:MI:SS') as auditdate,audit_action,audit_user_code,audit_branch_code ";
                    sql += " ,audit_module,audit_type,audit_refno,audit_remarks ";
                    sql += " from auditlog a ";
                    sql += sWhere;
                    sql += " order by auditdate";

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                }
                else
                {

                    if (type == "NEW")
                    {
                        sql = "SELECT count(*) as total, ceil(COUNT(*) / " + page_rows.ToString() + ") page_total  FROM auditlog  a ";
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

                    DataTable Dt_ftp = new DataTable();
                    sql = "";
                    sql += " select * from ( ";
                    sql += " select  to_char(audit_date,'DD/MM/YYYY HH24:MI:SS') as auditdate,audit_action,audit_user_code,audit_branch_code ";
                    sql += " ,audit_module,audit_type,audit_refno,audit_remarks ";
                    sql += " ,row_number() over(order by a.audit_date) rn ";
                    sql += " from auditlog a ";
                    sql += sWhere;
                    sql += ") a where rn between {startrow} and {endrow}";
                    sql += " order by auditdate";

                    sql = sql.Replace("{startrow}", startrow.ToString());
                    sql = sql.Replace("{endrow}", endrow.ToString());
                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                }
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mRow = new Auditlog();
                        mRow.audit_date = Dr["auditdate"].ToString();
                        mRow.audit_action = Dr["audit_action"].ToString();
                        mRow.audit_user_code = Dr["audit_user_code"].ToString();
                        mRow.audit_type = Dr["audit_type"].ToString();
                        mRow.audit_refno = Dr["audit_refno"].ToString();
                        mRow.audit_remarks = Dr["audit_remarks"].ToString();
                        mRow.audit_branch_code = Dr["audit_branch_code"].ToString();
                        mRow.audit_module = Dr["audit_module"].ToString();
                        mList.Add(mRow);
                    }
                

                if (type == "EXCEL" && mList != null)
                {
                    PrintAuditReport(mList, company_code);
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

        private void PrintAuditReport(List<Auditlog> mList, string comp_code)
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
            DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = " LIST";

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
                File_Display_Name = "AuditReport.xls";
                File_Name = Lib.GetFileName(report_folder, FolderId, File_Display_Name);

                WB = new ExcelFile();
                WB.Worksheets.Add("Report");
                WS = WB.Worksheets["Report"];

                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 35;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;

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

                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MODULE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "USER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ACTION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REMARKS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                foreach (Auditlog Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    Lib.WriteData(WS, iRow, iCol++, Rec.audit_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.audit_branch_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.audit_module, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.audit_user_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.audit_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.audit_action, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.audit_remarks, _Color, false, "", "L", "", _Size, false, 325, "", true);
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
        public IDictionary<string, object> LoadDefault(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            //Dictionary<string, object> parameter;

            LovService lovservice = new LovService();

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "SALES EXECUTIVE");
            //RetData.Add("smanlist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "CITY");
            //RetData.Add("citylist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "STATE");
            //RetData.Add("statelist", lovservice.Lov(parameter)["param"]);

            //parameter = new Dictionary<string, object>();
            //parameter.Add("table", "param");
            //parameter.Add("param_type", "COUNTRY");
            //RetData.Add("countrylist", lovservice.Lov(parameter)["param"]);

            return RetData;
        }
    }
}
