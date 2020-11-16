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
    public class RebateService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_List2 = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<Rebate> mList = new List<Rebate>();
        Rebate mrow;

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
        string hbl_type = "";
        string from_date = "";
        string to_date = "";
        string ErrorMessage = "";
        Boolean all = false;

        Boolean showpaid = false;

        decimal nRebate = 0;

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<Rebate>();
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
                hbl_type = SearchData["hbl_type"].ToString();
                from_date = SearchData["from_date"].ToString();
                to_date = SearchData["to_date"].ToString();
                all = (Boolean)SearchData["all"];
                showpaid = (Boolean)SearchData["showpaid"];

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

                sql = " select a.inv_pkid, c.hbl_pkid, h.jvh_vrno,h.jvh_date,c.hbl_no,a.inv_source, c.hbl_type,m.hbl_bl_no as mbl, ";
                sql += " c.hbl_bl_no as hbl, a.rec_branch_code as branch,nvl(sman2.param_name, sman.param_name)  as sman_name,   ";
                sql += " c.rec_created_date, c.hbl_job_nos,  ";
                sql += " d.acc_name as shipper_name,   ";
                sql += " carrier.param_name as carrier_name,   ";
                sql += " a.inv_type, b.acc_main_code, b.acc_code, b.acc_name, a.inv_qty, a.inv_Rate,a.inv_ftotal, a.inv_curr_code, a.inv_Exrate, a.inv_total,  ";
                sql += " a.inv_rebate_amt, a.inv_rebate_curr_code, a.inv_rebate_exrate, a.inv_rebate_amt_inr, ";
                sql += " a.inv_rebate_jvid, a.inv_rebate_jvno ,jv.jvh_date as jv_rebate_date,";
                sql += " cnge.cust_name as consignee_name ";

                sql += " from jobincome a ";
                sql += " left join ledgert t on a.inv_pkid = t.jv_pkid ";
                sql += " left join ledgerh h on t.jv_parent_id  = h.jvh_pkid  ";
                sql += " inner join acctm b on a.inv_acc_id = b.acc_pkid ";
                sql += " inner join hblm c on a.inv_parent_id = c.hbl_pkid ";
                sql += " left  join hblm m on c.hbl_mbl_id = m.hbl_pkid ";
                sql += " left join acctm d on case when c.hbl_type in('HBL-SE', 'HBL-AE') then c.hbl_exp_id else c.hbl_imp_id end = d.acc_pkid ";
                sql += " left join param carrier on c.hbl_carrier_id = carrier.param_pkid ";
                sql += " left join ledgerh jv on a.inv_rebate_jvid  = jv.jvh_pkid  ";

                sql += " left join customerm shpr on d.acc_pkid = cust_pkid ";
                sql += " left join custdet cd on d.acc_pkid = cd.det_cust_id and a.rec_branch_code =  cd.det_branch_code ";
                sql += " left join param sman  on shpr.cust_sman_id = sman.param_pkid";
                sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                sql += " left join customerm cnge on case when c.hbl_type in('HBL-SE', 'HBL-AE') then c.hbl_imp_id else c.hbl_exp_id end = cnge.cust_pkid ";

                sql += " where ";
                sql += " a.rec_company_code = '{COMPCODE}'";
                if (!all)
                {
                    sql += " and a.rec_branch_code = '{BRCODE}'";
                }

                if (hbl_type != "ALL")
                    sql += " and c.hbl_type = '" + hbl_type + "'";

                sql += " and(inv_rebate_amt > 0  or inv_rebate_amt_inr > 0) ";

                sql += " and to_char(c.rec_created_date,'DD-MON-YYYY') between to_date( '{FDATE}', 'DD-MON-YYYY')  and to_date( '{EDATE}', 'DD-MON-YYYY')  ";
                sql += " order by a.rec_branch_code,c.rec_created_date, c.hbl_type,c.hbl_no, inv_ctr ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{FDATE}", from_date);
                sql = sql.Replace("{EDATE}", to_date);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                if (showpaid)
                {

                    sql = " select inv_rebate_jvid, xref_cr_jvh_id,jvh_vrno, jvh_type, jvh_date, xref_amt, 'N' as flag   ";
                    sql += " from ledgerxref a ";
                    sql += " inner join (";
                    sql += " select distinct inv_rebate_jvid from jobincome a";
                    sql += " inner join hblm c on a.inv_parent_id = c.hbl_pkid   ";
                    sql += " where a.rec_company_code = '{COMPCODE}' ";
                    if (!all)
                    {
                        sql += " and a.rec_branch_code = '{BRCODE}' ";
                    }
                    sql += " and to_char(c.rec_created_date,'DD-MON-YYYY') between to_date( '{FDATE}', 'DD-MON-YYYY')  and to_date( '{EDATE}', 'DD-MON-YYYY') ";
                    sql += " and inv_rebate_jvid is not null";
                    sql += " ) b on a.xref_cr_jvh_id = inv_rebate_jvid";
                    sql += " inner join  ledgerh b on xref_dr_jvh_id = jvh_pkid ";
                    sql += " order by inv_rebate_jvid";


                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{FDATE}", from_date);
                    sql = sql.Replace("{EDATE}", to_date);

                    Con_Oracle = new DBConnection();
                    Dt_List2 = new DataTable();
                    Dt_List2 = Con_Oracle.ExecuteQuery(sql);
                    Con_Oracle.CloseConnection();
                }


                nRebate = 0;
                string paid_vrno = "";
                string paid_date = "";
                decimal paid_amount = 0;
                string vrno = "";
                string date = "";

                foreach (DataRow Dr in Dt_List.Rows)
                {

                    mrow = new Rebate();
                    mrow.row_type = "DETAIL";
                    mrow.selected = false;
                    mrow.created = Lib.DatetoStringDisplayformat(Dr["rec_created_date"]);
                    mrow.inv_pkid = Dr["inv_pkid"].ToString();
                    mrow.inv_no = Dr["jvh_vrno"].ToString();
                    mrow.hbl_no = Dr["hbl_no"].ToString();
                    mrow.inv_source = Dr["inv_source"].ToString();
                    mrow.hbl_pkid = Dr["hbl_pkid"].ToString();
                    mrow.hbl_type = Dr["hbl_type"].ToString();
                    mrow.mbl = Dr["mbl"].ToString();
                    mrow.hbl = Dr["hbl"].ToString();
                    mrow.jobnos = Dr["hbl_job_nos"].ToString();

                    mrow.shipper_name = Dr["shipper_name"].ToString();
                    mrow.carrier_name = Dr["carrier_name"].ToString();
                    mrow.inv_type = Dr["inv_type"].ToString();

                    mrow.acc_main_code = Dr["acc_main_code"].ToString();
                    mrow.acc_code = Dr["acc_code"].ToString();
                    mrow.acc_name = Dr["acc_name"].ToString();

                    mrow.inv_qty = Lib.Conv2Decimal(Dr["inv_qty"].ToString());
                    mrow.inv_rate = Lib.Conv2Decimal(Dr["inv_rate"].ToString());
                    mrow.inv_fototal = Lib.Conv2Decimal(Dr["inv_ftotal"].ToString());
                    mrow.inv_curr_code = Dr["inv_curr_code"].ToString();
                    mrow.inv_exrate = Lib.Conv2Decimal(Dr["inv_exrate"].ToString());
                    mrow.inv_total = Lib.Conv2Decimal(Dr["inv_total"].ToString());
                    mrow.inv_rebate_amt = Lib.Conv2Decimal(Dr["inv_rebate_amt"].ToString());
                    mrow.inv_rebate_curr_code = Dr["inv_rebate_curr_code"].ToString();
                    mrow.inv_rebate_exrate = Lib.Conv2Decimal(Dr["inv_rebate_exrate"].ToString());
                    mrow.inv_rebate_amt_inr = Lib.Conv2Decimal(Dr["inv_rebate_amt_inr"].ToString());

                    mrow.inv_rebate_jvid = Dr["inv_rebate_jvid"].ToString();
                    mrow.inv_rebate_jvno = Dr["inv_rebate_jvno"].ToString();

                    mrow.inv_date = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                    mrow.inv_rebate_jvdate = Lib.DatetoStringDisplayformat(Dr["jv_rebate_date"]);
                    mrow.branch = Dr["branch"].ToString();

                    mrow.inv_date_original = Lib.DatetoString(Dr["jvh_date"]);
                    mrow.inv_rebate_jvdate_original = Lib.DatetoString(Dr["jv_rebate_date"]);
                    mrow.salesman = Dr["sman_name"].ToString();
                    mrow.consignee_name = Dr["consignee_name"].ToString();

                    if (showpaid)
                    {
                        paid_amount = 0;
                        paid_vrno = "";
                        paid_date = "";

                        foreach (DataRow Dr2 in Dt_List2.Select("inv_rebate_jvid = '" + Dr["inv_rebate_jvid"].ToString() + "'"))
                        {
                            paid_amount += Lib.Conv2Decimal(Dr2["xref_amt"].ToString());
                            paid_date = Lib.DatetoStringDisplayformat(Dr2["jvh_date"]);
                            paid_vrno = Dr2["jvh_type"].ToString() + "-" + Dr2["jvh_vrno"].ToString();
                            Dr2["xref_amt"] = 0;
                        }

                        mrow.paid_vrno = paid_vrno;
                        mrow.paid_date = paid_date;
                        mrow.paid_amt = paid_amount;


                    }


                    mList.Add(mrow);
                    nRebate += Lib.Conv2Decimal(mrow.inv_rebate_amt_inr.ToString());
                }
                if (mList.Count > 1)
                {
                    mrow = new Rebate();
                    mrow.row_type = "TOTAL";
                    mrow.hbl_pkid = "";
                    mrow.inv_rebate_amt_inr = Lib.Conv2Decimal(Lib.NumericFormat(nRebate.ToString(), 2));
                    mList.Add(mrow);
                }

                
                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintReport();
                }
                Dt_List.Rows.Clear();
                Dt_List2.Rows.Clear();
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

        private void PrintReport()
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


                File_Display_Name = "Rebate.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                // WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;
                if(!all)
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 6;
                    WS.Columns[2].Width = 256 * 12;
                    WS.Columns[3].Width = 256 * 6;
                    WS.Columns[4].Width = 256 * 10;
                    WS.Columns[5].Width = 256 * 16;
                    WS.Columns[6].Width = 256 * 7;
                    WS.Columns[7].Width = 256 * 13;
                    WS.Columns[8].Width = 256 * 14;
                    WS.Columns[9].Width = 256 * 26;
                    WS.Columns[10].Width = 256 * 29;
                    WS.Columns[11].Width = 256 * 9;
                    WS.Columns[12].Width = 256 * 9;
                    WS.Columns[13].Width = 256 * 38;
                    WS.Columns[14].Width = 256 * 10;
                    WS.Columns[15].Width = 256 * 5;
                    WS.Columns[16].Width = 256 * 9;
                    WS.Columns[17].Width = 256 * 11;
                    WS.Columns[18].Width = 256 * 7;
                    WS.Columns[19].Width = 256 * 11;
                    WS.Columns[20].Width = 256 * 31;
                    WS.Columns[21].Width = 256 * 20;
                    WS.Columns[22].Width = 256 * 15;
                    WS.Columns[23].Width = 256 * 15;
                    WS.Columns[24].Width = 256 * 15;
                    WS.Columns[25].Width = 256 * 15;
                    WS.Columns[26].Width = 256 * 15;
                    WS.Columns[27].Width = 256 * 15;
                    WS.Columns[28].Width = 256 * 15;
                }
                else
                {
                    WS.Columns[0].Width = 256 * 2;
                    WS.Columns[1].Width = 256 * 10;
                    WS.Columns[2].Width = 256 * 6;
                    WS.Columns[3].Width = 256 * 12;
                    WS.Columns[4].Width = 256 * 6;
                    WS.Columns[5].Width = 256 * 10;
                    WS.Columns[6].Width = 256 * 16;
                    WS.Columns[7].Width = 256 * 7;
                    WS.Columns[8].Width = 256 * 13;
                    WS.Columns[9].Width = 256 * 14;
                    WS.Columns[10].Width = 256 * 26;
                    WS.Columns[11].Width = 256 * 29;
                    WS.Columns[12].Width = 256 * 9;
                    WS.Columns[13].Width = 256 * 9;
                    WS.Columns[14].Width = 256 * 38;
                    WS.Columns[15].Width = 256 * 10;
                    WS.Columns[16].Width = 256 * 5;
                    WS.Columns[17].Width = 256 * 9;
                    WS.Columns[18].Width = 256 * 11;
                    WS.Columns[19].Width = 256 * 7;
                    WS.Columns[20].Width = 256 * 11;
                    WS.Columns[21].Width = 256 * 31;
                    WS.Columns[22].Width = 256 * 20;
                    WS.Columns[23].Width = 256 * 15;
                    WS.Columns[24].Width = 256 * 15;
                    WS.Columns[25].Width = 256 * 15;
                    WS.Columns[26].Width = 256 * 15;
                    WS.Columns[27].Width = 256 * 15;
                    WS.Columns[28].Width = 256 * 15;
                    WS.Columns[29].Width = 256 * 15;
                }
               


                iRow = 0; iCol = 1;

                if(all)
                {
                    WS.Columns[15].Style.NumberFormat = "#0.00";
                    WS.Columns[17].Style.NumberFormat = "#0.00";
                    WS.Columns[18].Style.NumberFormat = "#0.00";
                }
                else
                {
                    WS.Columns[14].Style.NumberFormat = "#0.00";
                    WS.Columns[16].Style.NumberFormat = "#0.00";
                    WS.Columns[17].Style.NumberFormat = "#0.00";
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
                // Lib.WriteMergeCell(WS, iRow, 1, 9, 1, COMPWEB, "Calibri", 12, true, Color.Blue, "C", "C", "", "");
                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "REBATE LIST", _Color, true, "", "L", "", 15, false, 325, "", true);
                // Lib.WriteMergeCell(WS, iRow, 1, 9, 1,"TEU REPORT", "Calibri", 15, true, Color.Black, "C", "C", "", "");
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                if(all)
                {
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
               
                Lib.WriteData(WS, iRow, iCol++, "INV#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INV-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI#", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SI-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "MBL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "HBL", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CARRIER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PP/CC", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "NAME", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
               
                Lib.WriteData(WS, iRow, iCol++, "REBATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CURR", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EXRATE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "REBATE-INR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JV-NO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "JV-DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CONSIGNEE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                if(showpaid)
                {
                    Lib.WriteData(WS, iRow, iCol++, "PAID-VRNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "AMT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                }
                foreach (Rebate Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.inv_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.created, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_source, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.mbl, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.hbl, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.shipper_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.carrier_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_type, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.acc_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.acc_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                      
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_rebate_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_rebate_curr_code, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_rebate_exrate, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_rebate_amt_inr, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.inv_rebate_jvno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.inv_rebate_jvdate, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.consignee_name, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.salesman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        if(showpaid)
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.paid_vrno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Lib.nvlDate(Rec.paid_date, "", Lib.FRONT_END_DATE_DISPLAY_FORMAT), _Color, false, "", "L", "", _Size, false, 325, Lib.FRONT_END_DATE_DISPLAY_FORMAT, true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.paid_amt, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        }
                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if(all)
                        {
                            Lib.WriteData(WS, iRow, 18, Rec.inv_rebate_amt_inr, _Color, true, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        }
                        else
                        {
                            Lib.WriteData(WS, iRow, 17, Rec.inv_rebate_amt_inr, _Color, true, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        }
                       
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

