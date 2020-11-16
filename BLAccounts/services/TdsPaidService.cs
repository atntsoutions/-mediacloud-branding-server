using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;


using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts

{
    public class TdsPaidService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<TdsPaid> mList = new List<TdsPaid>();
        TdsPaid mrow;
        int iRow = 0;
        int iCol = 0;
        string type = "";

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
        string rec_category = "";
        string searchexpid = "";
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string code = "";
        string id = "";
        string acc_name = "";
       
        string ErrorMessage = "";
        Boolean main_code = false;
        decimal tot_debit = 0;
        decimal tot_credit = 0;
        decimal tot_deference = 0;
        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<TdsPaid>();
            ErrorMessage = "";
            try
            {
                type = SearchData["type"].ToString();

                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
              
                id = SearchData["acc_id"].ToString();
                acc_name = SearchData["acc_name"].ToString();

                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                sql = "";
                sql = "  select  jvh_date, jvh_vrno, jvh_type, ";
                sql += " jv_gross_bill_amt, jv_debit, jv_credit,";
                sql += " cust.cust_code as party_code, ";
                sql += " cust.cust_name as party_name,";
                sql += " tan.param_code as tan,";
                sql += " tan.param_name as tan_name,";
                sql += " nvl(sman2.param_name,sman.param_name) as sman_name,";//sman.param_name as sman_name
                sql += " jvh_narration ";
                sql += " from ledgerh a ";
                sql += " inner join ledgert b on jvh_pkid = b.jv_parent_id";

                sql += " left join customerm cust on b.jv_tan_party_id = cust.cust_pkid";
                sql += "  left join custdet  cd on b.rec_branch_code = cd.det_branch_code and b.jv_tan_party_id = cd.det_cust_id ";

                sql += " left join  param sman on cust.cust_sman_id = sman.param_pkid";
                sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                sql += " left join param tan on jv_tan_id = tan.param_pkid";
                
                sql += " where  jvh_year ='{YEAR}' and jv_acc_id = '{ID}' ";
                sql += " and a.rec_branch_code = '{BRCODE}'";
                sql += " order by a.jvh_date, jvh_type, jvh_vrno";


                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEAR}", year_code);
                sql = sql.Replace("{ID}", id);


                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                tot_credit = 0;
                tot_debit = 0;
                tot_deference = 0;
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new TdsPaid();
                    mrow.rowtype = "DETAIL";
                    mrow.rowcolor = "BLACK";
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.jvh_vrno = Dr["jvh_vrno"].ToString();
                    mrow.jvh_type = Dr["jvh_type"].ToString();
                    mrow.jv_gross_bill_amt = Lib.Conv2Decimal(Dr["jv_gross_bill_amt"].ToString());
                    mrow.jv_debit = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                    mrow.jv_credit = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                    mrow.party_code = Dr["party_code"].ToString();
                    mrow.party_name = Dr["party_name"].ToString();
                    mrow.tan = Dr["tan"].ToString();
                    mrow.tan_name = Dr["tan_name"].ToString();
                    mrow.sman_name = Dr["sman_name"].ToString();
                    mrow.jvh_narration = Dr["jvh_narration"].ToString();
                    mList.Add(mrow);

                    tot_debit += Lib.Conv2Decimal(mrow.jv_debit.ToString());
                    tot_credit += Lib.Conv2Decimal(mrow.jv_credit.ToString());

                }
                if (mList.Count > 1)
                {
                    mrow = new TdsPaid();
                    mrow.rowtype = "TOTAL";
                    mrow.rowcolor = "RED";
                    mrow.jvh_date = "TOTAL";
                    mrow.jv_debit = Lib.Conv2Decimal(Lib.NumericFormat(tot_debit.ToString(), 2));
                    mrow.jv_credit = Lib.Conv2Decimal(Lib.NumericFormat(tot_credit.ToString(), 2));
                    mList.Add(mrow);

                    mrow = new TdsPaid();
                    mrow.rowtype = "BALANCE";
                    mrow.rowcolor = "RED";
                    mrow.jvh_date = "BALANCE";
                   
                    tot_deference = tot_debit - tot_credit;
                    mrow.jv_debit = Lib.Conv2Decimal(Lib.NumericFormat(tot_deference.ToString(), 2));
                   // mrow.jv_debit = 0;                
                    mList.Add(mrow);
                }


                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintTdsPaidReport();
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

        private void PrintTdsPaidReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            string REPORT_CAPTION = "";

            //string _Border = "";
            //Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;
            //DataRow Dr_target = null;
            iRow = 0;
            iCol = 0;
            try
            {
                REPORT_CAPTION = searchtype;

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

                File_Display_Name = "TdsPaidReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 10;
                WS.Columns[3].Width = 256 * 8;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 25;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 25;
                WS.Columns[11].Width = 256 * 25;
                WS.Columns[12].Width = 256 * 150;
                WS.Columns[13].Width = 256 * 15;
              



                iRow = 0; iCol = 1;
                WS.Columns[4].Style.NumberFormat = "#0.00";
                WS.Columns[5].Style.NumberFormat = "#0.00";
                WS.Columns[6].Style.NumberFormat = "#0.00";

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
                Lib.WriteData(WS, iRow, 1, "TDS PAID : "+acc_name, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "GROSS", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NARRATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);



                foreach (TdsPaid Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.rowtype == "DETAIL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_vrno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_type, _Color, false, "", "L", "", _Size, false, 325, "", true);                    
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_gross_bill_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_debit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.party_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.party_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.tan_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.sman_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_narration, _Color, false, "", "L", "", _Size, false, 325, "", true);

                    }
                    if (Rec.rowtype == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, true, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_debit, _Color, true, "T", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_credit, _Color, true, "T", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "T", "L", "", _Size, false, 325, "", true);

                    }
                    if (Rec.rowtype == "BALANCE")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.jvh_date, _Color, true, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.jv_debit, _Color, true, "B", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "B", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "B", "L", "", _Size, false, 325, "", true);

                    }
                }
                // iRow++;

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
