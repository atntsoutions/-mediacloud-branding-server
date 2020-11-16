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
    public class BrokService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<BrokReport> mList = new List<BrokReport>();
        BrokReport mrow;
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
        string cntr_no = "";
        string ErrorMessage = "";
       
        string from_date = "";
        string to_date = "";
        Decimal tot_brok = 0;
        Boolean all = false;


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<BrokReport>();
            ErrorMessage = "";
            try
            {
                type = SearchData["type"].ToString();
                rec_category = SearchData["rec_category"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                //  searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();
                all = (Boolean)SearchData["all"];

                from_date = Lib.StringToDate(from_date);
                to_date = Lib.StringToDate(to_date);


                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                sql = " select  a.rec_branch_code as branch, ";
                sql += " jvh_vrno as our_ref_no,jvh_date as our_ref_date,";
                sql += " hbl_no as bkslno, hbl_bl_no as blno, hbl_date as bldate,";
                sql += " vsl.param_name as vessel, hbl_vessel_no as voyage,";
                sql += " hbl_terms as frt_status, hbl_nature as movement,";
                sql += " carrier.param_name as carrier,";
                sql += " agent.cust_name as agent,";
                sql += " shpr.cust_name as shipper,";
                sql += " cons.cust_name as consignee,";
                sql += " jvh_reference as refno,jvh_reference_date as ref_date,  ";
                sql += " jvh_org_invno as invno, jvh_org_invdt as inv_date,";
                sql += " jvh_basic_frt as frt, jvh_brok_per as brok_per, jvH_brok_amt as brok_amt,   ";
                sql += " jvh_brok_exrate,jvh_brok_amt_inr ";
                sql += " from ledgerh a";
                sql += " left join hblm m on a.jvh_cc_id = m.hbl_pkid";
                sql += " left join customerm shpr on m.hbl_exp_id = shpr.cust_pkid";
                sql += " left join customerm cons on m.hbl_imp_id = cons.cust_pkid";
                sql += " left join customerm agent on m.hbl_agent_id = agent.cust_pkid";
                sql += " left join param carrier on m.hbl_carrier_id = carrier.param_pkid";
                sql += " left join param vsl on m.hbl_vessel_id = vsl.param_pkid";
                sql += " where a.jvh_brok_amt > 0 ";
                if (!all)
                {
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                }
                sql += " and a.jvh_date between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY')";

                sql += " order by a.rec_branch_code,a.jvh_date";

                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

               // tot_brok = 0;

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new BrokReport();
                    mrow.branch = Dr["branch"].ToString();
                    mrow.jvh_vrno = Dr["our_ref_no"].ToString();
                    mrow.jvh_date = Lib.DatetoStringDisplayformat(Dr["our_ref_date"]);
                    mrow.hbl_no = Dr["bkslno"].ToString();
                    mrow.hbl_bl_no = Dr["blno"].ToString();
                    mrow.hbl_date = Lib.DatetoStringDisplayformat(Dr["bldate"]);
                    mrow.vessel = Dr["vessel"].ToString();
                    mrow.hbl_vessel_no = Dr["voyage"].ToString();
                    mrow.hbl_terms = Dr["frt_status"].ToString();
                    mrow.hbl_nature = Dr["movement"].ToString();
                    mrow.carrier = Dr["carrier"].ToString();
                    mrow.agent = Dr["agent"].ToString();
                    mrow.shipper = Dr["shipper"].ToString();
                    mrow.consignee = Dr["consignee"].ToString();
                    mrow.jvh_reference = Dr["refno"].ToString();
                    mrow.jvh_reference_date = Lib.DatetoStringDisplayformat(Dr["ref_date"]);
                    mrow.jvh_org_invno = Dr["invno"].ToString();
                    mrow.jvh_org_invdt = Lib.DatetoStringDisplayformat(Dr["inv_date"]);
                    mrow.jvh_basic_frt = Lib.Conv2Decimal(Dr["frt"].ToString());
                    mrow.jvh_brok_per = Lib.Conv2Decimal(Dr["brok_per"].ToString());
                    mrow.jvh_brok_amt = Lib.Conv2Decimal(Dr["brok_amt"].ToString());
                    mrow.jvh_brok_exrate = Lib.Conv2Decimal(Dr["jvh_brok_exrate"].ToString());
                    mrow.jvh_brok_amt_inr = Lib.Conv2Decimal(Dr["jvh_brok_amt_inr"].ToString());
                    mList.Add(mrow);
                   // tot_brok += Lib.Conv2Decimal(mrow.jvh_brok_amt.ToString());

                }

                //if (mList.Count > 1)
                //{
                //    mrow = new BrokReport();
                //    mrow.row_type = "TOTAL";
                //    mrow.row_colour = "RED";
                //    mrow.jvh_vrno = "TOTAL";
                //    mrow.jvh_brok_amt = Lib.Conv2Decimal(Lib.NumericFormat(tot_brok.ToString(), 3));
                //    mList.Add(mrow);
                //}


                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintBrokReport();
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

        private void PrintBrokReport()
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
                REPORT_CAPTION = searchtype;

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if(all)
                {
                    mSearchData.Add("branch_code", "HOCPL");
                }
                else
                {
                    mSearchData.Add("branch_code", branch_code);
                }
                

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

                File_Display_Name = "EGMReport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 15;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 15;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 25;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
                WS.Columns[19].Width = 256 * 15;
                WS.Columns[20].Width = 256 * 15;
                WS.Columns[21].Width = 256 * 15;
                WS.Columns[22].Width = 256 * 15;
                WS.Columns[23].Width = 256 * 15;
                WS.Columns[24].Width = 256 * 15;

                iRow = 0; iCol = 1;
               if(all)
                {
                    WS.Columns[19].Style.NumberFormat = "#0.00";
                    WS.Columns[20].Style.NumberFormat = "#0.00";
                    WS.Columns[21].Style.NumberFormat = "#0.00";
                }
                else
                {
                    WS.Columns[18].Style.NumberFormat = "#0.00";
                    WS.Columns[19].Style.NumberFormat = "#0.00";
                    WS.Columns[20].Style.NumberFormat = "#0.00";
                }

               

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
                Lib.WriteData(WS, iRow, 1, "BROKERAGE REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 11;
                iCol = 1;
                if (all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "OUR-REF-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OUR-REF-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BKSLNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BLNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BLDATE ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VESSEL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VOYAGE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FRT-STATUS", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MOVEMENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REFNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REF-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "FRT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BROK-PER", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BROK-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BROK-EX-RATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BROK-INR-AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                _Size = 10;
                Decimal si_no = 1;
                foreach (BrokReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (all)
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    }
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_vrno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_bl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.hbl_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.vessel, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_vessel_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_terms, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nature, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.carrier, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.agent, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.shipper, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.consignee, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_reference, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_reference_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_org_invno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.jvh_org_invdt, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);

                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_basic_frt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_brok_per, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_brok_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_brok_exrate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    Lib.WriteData(WS, iRow, iCol++, Rec.jvh_brok_amt_inr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
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

