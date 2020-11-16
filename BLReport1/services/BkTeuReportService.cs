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
    public class BkTeuReportService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<BkTeuReport> mList = new List<BkTeuReport>();
        BkTeuReport mrow;
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
        string searchexpid = "";
        string type_date = "";
        string from_date = "";
        string to_date = "";
        string to_date_month = "";
        string ErrorMessage = "";
        string SobDate = "";
        string shipper_id = "";
        string consignee_id = "";
        string agent_id = "";
        string carrier_id = "";
        string pol_id = "";
        string pod_id = "";

        decimal tot_teu_ason_day = 0;
        decimal tot_teu = 0;
        decimal tot_20tue = 0;
        decimal tot_40tue = 0;
        //decimal tot_grwt = 0;
        decimal tot_cbm = 0;
        Boolean all = false;
        

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<BkTeuReport>();
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
                type_date = SearchData["type_date"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();

                if (SearchData.ContainsKey("shipper_id"))
                    shipper_id = SearchData["shipper_id"].ToString();

                if (SearchData.ContainsKey("consignee_id"))
                    consignee_id = SearchData["consignee_id"].ToString();

                if (SearchData.ContainsKey("agent_id"))
                    agent_id = SearchData["agent_id"].ToString();

                if (SearchData.ContainsKey("carrier_id"))
                    carrier_id = SearchData["carrier_id"].ToString();

                if (SearchData.ContainsKey("pol_id"))
                    pol_id = SearchData["pol_id"].ToString();

                if (SearchData.ContainsKey("pod_id"))
                    pod_id = SearchData["pod_id"].ToString();

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

                //    if (days > 31)
                //        Lib.AddError(ref ErrorMessage, " | Only one month data range can be used,use excel to download");
                //}
                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                to_date_month = DateTime.Parse(to_date).ToString("MMMM");


                sql = " select hbl_pkid ,HBL_BOOK_CNTR as cntr_no,HBL_BOOK_CNTR_TEU as cntr_teu,a.rec_branch_code as branch,";
                sql += " a.hbl_no as cntr_booking_no,";
                sql += " a.hbl_book_cntr_m20 as teu20,a.hbl_book_cntr_m40 as teu40,a.hbl_book_cntr_mteu as mteu,a.hbl_book_cntr_mcbm as mcbm, ";
                sql += " a.hbl_bl_no as mbl_no,a.hbl_book_no as mbl_book_no,a.hbl_pol_etd as mbl_pol_etd,a.hbl_shipment_type as mbl_shipment_type,";
                sql += " a.hbl_nature as mbl_nature,shpr.cust_name as mbl_exp_name,cnge.cust_name as mbl_imp_name,liner.param_name as mbl_carrier_name,";
                sql += " pol.param_name as pol,pod.param_name as pod,pofd.param_name as pofd,agent.cust_name as agent,nvl(a.hbl_nomination,cnge.cust_nomination) as hbl_nomination,";
                sql += " row_number() over(order by hbl_no) rn,a.rec_created_date  ";
                sql += " from hblm a  ";
                sql += " left join customerm shpr on a.hbl_exp_id = shpr.cust_pkid ";
                sql += " left join customerm cnge on a.hbl_imp_id = cnge.cust_pkid";
                sql += " left join param liner on a.hbl_carrier_id = liner.param_pkid ";
                sql += " left join param pol on a.hbl_pol_id = pol.param_pkid";
                sql += " left join param pod on a.hbl_pod_id = pod.param_pkid ";
                sql += " left join param pofd on a.hbl_pofd_id = pofd.param_pkid";
                sql += " left join customerm agent on a.hbl_agent_id = agent.cust_pkid";
                sql += " where a.rec_company_code = '{COMPCODE}'";

                sql += " and a.hbl_status_id not in ('A493CB8A-4041-4DCC-8749-F09034F886F9') ";//for not display cancelled liner booking status.

                if (!all)
                {
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                }
                
                sql += " and a.hbl_type='MBL-SE' ";
                if (type_date == "SOB")
                    sql += "  and a.hbl_pol_etd between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";
                else
                    sql += "  and to_char(a.rec_created_date,'DD-MON-YYYY') between to_date('{FDATE}','DD-MON-YYYY') and to_date('{EDATE}','DD-MON-YYYY') ";

                if (shipper_id.Length > 0)
                    sql += " and a.hbl_exp_id = '" + shipper_id + "' ";
                if (consignee_id.Length > 0)
                    sql += " and a.hbl_imp_id = '" + consignee_id + "' ";
                if (agent_id.Length > 0)
                    sql += " and a.hbl_agent_id = '" + agent_id + "' ";

                if (carrier_id.Length > 0)
                    sql += " and a.hbl_carrier_id = '" + carrier_id + "'";
                if (pol_id.Length > 0)
                    sql += " and a.hbl_pol_id = '" + pol_id + "'";
                if (pod_id.Length > 0)
                    sql += " and a.hbl_pod_id = '" + pod_id + "'";

                if (type_date == "SOB")
                    sql += " order by a.rec_branch_code,a.hbl_pol_etd";
                else
                    sql += " order by a.rec_branch_code,a.rec_created_date";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                
                tot_teu_ason_day = 0;
                tot_teu = 0;
                tot_20tue = 0;
                tot_40tue = 0;               
                tot_cbm = 0;
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new BkTeuReport();
                    mrow.row_type = "DETAIL";
                    mrow.row_colour = "BLACK";                 
                    mrow.cntr_no = Dr["cntr_no"].ToString();
                    if (Dr["mbl_shipment_type"].ToString() == "LCL")
                    {
                        mrow.mbl_book_cntr_mteu = 0;
                        mrow.mbl_book_cntr_m20 = 0;
                        mrow.mbl_book_cntr_m40 = 0;
                    }
                    else
                    {
                        mrow.mbl_book_cntr_mteu = Lib.Conv2Decimal(Dr["mteu"].ToString());
                        mrow.mbl_book_cntr_m20 = Lib.Conv2Decimal(Dr["teu20"].ToString());
                        mrow.mbl_book_cntr_m40 = Lib.Conv2Decimal(Dr["teu40"].ToString());
                    }  
                    
                    mrow.cntr_booking_no = Dr["cntr_booking_no"].ToString();
                    mrow.mbl_book_no = Dr["mbl_book_no"].ToString();
                    mrow.mbl_no = Dr["mbl_no"].ToString();
                    mrow.mbl_pol_etd = Lib.DatetoStringDisplayformat(Dr["mbl_pol_etd"]);
                    mrow.mbl_shipment_type = Dr["mbl_shipment_type"].ToString();
                    mrow.mbl_nature = Dr["mbl_nature"].ToString();                                  
                    mrow.mbl_book_cntr_mcbm = Lib.Conv2Decimal(Dr["mcbm"].ToString());
                    mrow.mbl_exp_name = Dr["mbl_exp_name"].ToString();
                    mrow.mbl_imp_name = Dr["mbl_imp_name"].ToString();
                    mrow.mbl_carrier_name = Dr["mbl_carrier_name"].ToString();
                    mrow.hbl_pol = Dr["pol"].ToString();
                    mrow.hbl_pod = Dr["pod"].ToString();
                    mrow.hbl_pofd = Dr["pofd"].ToString();
                    mrow.hbl_agent = Dr["agent"].ToString();
                    mrow.branch = Dr["branch"].ToString();
                    mrow.hbl_nomination = Dr["hbl_nomination"].ToString();
                    mList.Add(mrow);

                    if (type_date == "SOB")
                        SobDate = Lib.StringToDate(Dr["mbl_pol_etd"]);
                    else
                        SobDate = Lib.StringToDate(Dr["rec_created_date"]);

                    if (SobDate == to_date)
                        tot_teu_ason_day += Lib.Conv2Decimal(mrow.mbl_book_cntr_mteu.ToString());
                    tot_teu += Lib.Conv2Decimal(mrow.mbl_book_cntr_mteu.ToString());
                    tot_20tue += Lib.Conv2Decimal(mrow.mbl_book_cntr_m20.ToString());
                    tot_40tue += Lib.Conv2Decimal(mrow.mbl_book_cntr_m40.ToString());
                    tot_cbm += Lib.Conv2Decimal(mrow.mbl_book_cntr_mcbm.ToString());
                }
                if (mList.Count > 1)
                {
                    mrow = new BkTeuReport();
                    mrow.row_type = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.cntr_booking_no = "TOTAL";
                    mrow.mbl_book_cntr_mteu = Lib.Conv2Decimal(Lib.NumericFormat(tot_teu.ToString(), 2));
                    mrow.mbl_book_cntr_m20 = Lib.Conv2Decimal(Lib.NumericFormat(tot_20tue.ToString(), 0));
                    mrow.mbl_book_cntr_m40 = Lib.Conv2Decimal(Lib.NumericFormat(tot_40tue.ToString(), 3));
                    mrow.mbl_book_cntr_mcbm = Lib.Conv2Decimal(Lib.NumericFormat(tot_cbm.ToString(), 3));
                    mList.Add(mrow);
                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintBkTeuReport();
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
            RetData.Add("tomonth", to_date_month);
            RetData.Add("totteu", Lib.NumericFormat(tot_teu.ToString(),2));
            RetData.Add("totteuday", Lib.NumericFormat(tot_teu_ason_day.ToString(), 2));
            return RetData;
        }

        private void PrintBkTeuReport()
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



                File_Display_Name = "BkTeu.xls";
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
                WS.Columns[4].Width = 256 * 25;
                WS.Columns[5].Width = 256 * 25;
                WS.Columns[6].Width = 256 * 25;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 18;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;
                WS.Columns[15].Width = 256 * 15;
                WS.Columns[16].Width = 256 * 15;
                WS.Columns[17].Width = 256 * 15;
                WS.Columns[18].Width = 256 * 15;
               
                if (all)
                {
                    WS.Columns[19].Width = 256 * 15;
                    WS.Columns[20].Width = 256 * 30;
                }
                else
                {
                    WS.Columns[19].Width = 256 * 30;
                    WS.Columns[20].Width = 256 * 15;
                }
               

                iRow = 0; iCol = 1;

                if(all)
                {
                    WS.Columns[12].Style.NumberFormat = "#0.00";
                    WS.Columns[13].Style.NumberFormat = "#0.00";
                    WS.Columns[14].Style.NumberFormat = "#0.00";
                    WS.Columns[15].Style.NumberFormat = "#0.00";
                }
                else
                {
                    WS.Columns[11].Style.NumberFormat = "#0.00";
                    WS.Columns[12].Style.NumberFormat = "#0.00";
                    WS.Columns[13].Style.NumberFormat = "#0.00";
                    WS.Columns[14].Style.NumberFormat = "#0.00";
                }
               
               


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
                Lib.WriteData(WS, iRow, 1, "BOOKING REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);

                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                Lib.WriteData(WS, iRow, iCol++, "MBLBK#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BOOKING.NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);              
                Lib.WriteData(WS, iRow, iCol++, "MBL#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHIPPER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NOMINATION", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AGENT", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "POFD", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "20", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "40", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TEU", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CBM", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                
                Lib.WriteData(WS, iRow, iCol++, "SOB", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NATURE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SHPMNT.TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONTAINER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
               
                foreach (BkTeuReport Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_booking_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_no, _Color, false, "", "L", "", _Size, false, 325, "", true);                       
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_exp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_imp_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_nomination, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_carrier_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_agent, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pol, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pod, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_pofd, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_cntr_m20, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_cntr_m40,_Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_cntr_mteu, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_book_cntr_mcbm, _Color, false, "", "R", "", _Size, false, 325, "#,0.000;(#,0.000);#", false);
                        
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.mbl_pol_etd, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_nature, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl_shipment_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        
                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.cntr_booking_no, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,"", _Color, true, "BT", "C", "", _Size, false, 325, "", true);                       
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "C", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,Rec.mbl_book_cntr_m20, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,Rec.mbl_book_cntr_m40, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,Rec.mbl_book_cntr_mteu, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++,Rec.mbl_book_cntr_mcbm, _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                       
                    }
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

