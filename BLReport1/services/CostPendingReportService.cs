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
    public class CostPendingService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<CostPendingReport> mList = new List<CostPendingReport>();
        CostPendingReport mrow;
        int iRow = 0;
        int iCol = 0;
        string type = "";
        string types = "";
        string report_folder = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "myreport.xls";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string searchtype = "";
        string searchstring = "";
        string category = "";
        string searchexpid = "";
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string ErrorMessage = "";
        string sort_colname = "";
        Boolean all = false;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<CostPendingReport>();
            ErrorMessage = "";
            try
            {

                type = SearchData["type"].ToString();
                category = SearchData["category"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
                type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();
                sort_colname = SearchData["sort_colname"].ToString();

                all = (Boolean)SearchData["all"];

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);


                if (from_date == "NULL" || to_date == "NULL")
                    Lib.AddError(ref ErrorMessage, " | Date Cannot Be Empty");

                //if (type == "SCREEN" && from_date != "NULL" && to_date != "NULL")
                //{
                //    DateTime dt_frm = DateTime.Parse(from_date);
                //    DateTime dt_to = DateTime.Parse(to_date);
                //    int days = (dt_to - dt_frm).Days;

                //    //if (days > 31)
                //    //    Lib.AddError(ref ErrorMessage, " | Only one month data range can be used,use excel to download");
                //}
                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (sort_colname == "")
                    sort_colname = "mbl.rec_created_date";

                sql = " select mbl.hbl_pkid as mbl_pkid, mbl.hbl_bl_no as mbl_bl_no,mbl.hbl_date as mbl_date,";
                sql += "  mbl.hbl_pol_etd as  mbl_sob_date,mbl.hbl_no as book_no, ";
                sql += "  agent.cust_name as mbl_agent_name,mbl.rec_created_date,";
                sql += "  mbl.hbl_folder_no as mbl_folder_no,mbl.hbl_nocosting as mbl_nocosting, ";
                sql += "  mbl.hbl_folder_sent_date as mbl_folder_sent_date,mbl.hbl_book_cntr as mbl_book_cntr,";
                sql += "  mbl.rec_branch_code,";
                sql += "  nvl(cost_date,mbl.hbl_nocosting_date) as cost_date,cost_refno,cost_jv_posted ";
                sql += "  from  hblm mbl ";
                sql += "  left join costingm a on mbl.hbl_pkid = a.cost_mblid ";
                sql += "  left join customerm agent on mbl.hbl_agent_id = agent.cust_pkid";
                sql += "  where mbl.rec_company_code = '{COMPCODE}' ";

                if (!all)
                    sql += "  and mbl.rec_branch_code = '{BRCODE}' ";

                sql += " and mbl.hbl_type = '{CATEGORY}' ";
                sql += "  and to_char(mbl.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                sql += "  order by mbl.rec_branch_code," + sort_colname;

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{CATEGORY}", category);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new CostPendingReport();
                    mrow.mbl_pkid = Dr["mbl_pkid"].ToString();
                    mrow.mbl_hbl_no = Dr["book_no"].ToString();
                    mrow.mbl_bl_no = Dr["mbl_bl_no"].ToString();
                    mrow.mbl_date = Lib.DatetoStringDisplayformat(Dr["mbl_date"]);
                    mrow.mbl_sob_date = Lib.DatetoStringDisplayformat(Dr["mbl_sob_date"]);
                    mrow.mbl_agent_name = Dr["mbl_agent_name"].ToString();
                    mrow.mbl_folder_no = Dr["mbl_folder_no"].ToString();
                    mrow.mbl_folder_sent_date = Lib.DatetoStringDisplayformat(Dr["mbl_folder_sent_date"]);
                    mrow.cost_date = Lib.DatetoStringDisplayformat(Dr["cost_date"]);
                    mrow.cost_refno = Dr["cost_refno"].ToString();
                    mrow.rec_created_date = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mrow.mbl_book_cntr = Dr["mbl_book_cntr"].ToString();
                    mrow.mbl_nocosting = Dr["mbl_nocosting"].ToString();
                    mrow.cost_jv_posted = Dr["cost_jv_posted"].ToString();
                    mrow.rec_branch_code = Dr["rec_branch_code"].ToString();
                    mList.Add(mrow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintPendingListReport();
                }

                Dt_List.Rows.Clear();

            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", type);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", File_Type);
            RetData.Add("filedisplayname", File_Display_Name);
            RetData.Add("list", mList);
            return RetData;
        }

        private void PrintPendingListReport()
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
            try
            {
                REPORT_CAPTION = "";

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if (!all)
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                else
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }


                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        COMPNAME = Dr["BR_NAME"].ToString();
                        COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
                        COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
                        COMPTEL = Dr["COMP_TEL"].ToString();
                        COMPFAX = Dr["COMP_FAX"].ToString();
                        COMPWEB = Dr["COMP_WEB"].ToString();
                        break;
                    }
                }

                File_Display_Name = "CostPending Report.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                if (!all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 15;
                    WS.Columns[2].Width = 256 * 22;
                    WS.Columns[3].Width = 256 * 15;
                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 15;
                    WS.Columns[6].Width = 256 * 20;
                    WS.Columns[7].Width = 256 * 15;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 15;
                    WS.Columns[10].Width = 256 * 15;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 15;
                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 13;
                    WS.Columns[2].Width = 256 * 15;
                    WS.Columns[3].Width = 256 * 22;
                    WS.Columns[4].Width = 256 * 15;
                    WS.Columns[5].Width = 256 * 15;
                    WS.Columns[6].Width = 256 * 15;
                    WS.Columns[7].Width = 256 * 20;
                    WS.Columns[8].Width = 256 * 15;
                    WS.Columns[9].Width = 256 * 15;
                    WS.Columns[10].Width = 256 * 15;
                    WS.Columns[11].Width = 256 * 15;
                    WS.Columns[12].Width = 256 * 15;
                    WS.Columns[13].Width = 256 * 15;
                }

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
                Lib.WriteData(WS, iRow, 1, str, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, false, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "COST PENDING REPORT ", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "CREATED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FOLDER-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SENT-ON", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINERS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SOB-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST-REFNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "COST-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POSTED", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NO COSTING", _Color, true, "BT", "L", "", _Size, false, 325, "", true);


                foreach (CostPendingReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.rec_branch_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }

                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.rec_created_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_agent_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_folder_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_folder_sent_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_cntr, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_sob_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_refno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.cost_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.cost_jv_posted, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.mbl_nocosting, _Color, false, "", "L", "", _Size, false, 325, "", true);

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

    }
}

