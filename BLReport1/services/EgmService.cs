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
    public class EgmService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<EgmReport> mList = new List<EgmReport>();
        EgmReport mrow;
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
        string sb_date = "";
        string egm_date = "";
        
       

        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<EgmReport>();
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
                searchstring = SearchData["searchstring"].ToString().ToUpper().Trim();

                if (searchstring.Length <= 0)
                    Lib.AddError(ref ErrorMessage, " | Container Cannot Be Empty");
                else
                    cntr_no = searchstring.Replace(",", "','");
                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                sql = " select distinct cntr_no,job_docno,pol.param_code as pol_code, cntr_egmno as egno, cntr_egmdt as egdate, ";
                sql += " job_cargo_nature as cargo_nature,job_pkg as quantity,job_grwt as gr_wt,pol.param_name as sb_filed_location,";
                sql += " job_pkg as no_of_pkg,pkgunit.param_code as pkg_unit,pofd.param_code as destination_port,cmdty.param_name as cargo_discription,";
                sql += " exp.cust_name as exporter,expaddr.add_line1 as exp_add1,expaddr.add_line2 as exp_add2,expaddr.add_line3 as exp_add3,";
                sql += " imp.cust_name as importer,impaddr.add_line1 as imp_add1,impaddr.add_line2 as imp_add2,impaddr.add_line3 as imp_add3,";
                sql += " opr_sbill_no, opr_sbill_date ,null as shut_out_quantity";
                sql += " from containerm a  ";
                sql += " inner join packingm b on a.cntr_pkid = pack_cntr_id ";
                sql += " inner join jobm c on pack_job_id = job_pkid ";
                sql += " inner join param pol on job_pol_id = pol.param_pkid ";
                sql += " left join joboperationsm opr  on job_pkid = opr_job_id  ";
                sql += " left join param pkgunit on c.job_pkg_unit_id = pkgunit.param_pkid";
                sql += " left join param pofd on c.job_pofd_id=pofd.param_pkid";
                sql += " left join customerm exp on c.job_exp_id = exp.cust_pkid";
                sql += " left join addressm expaddr on c.job_exp_br_id = expaddr.add_pkid";
                sql += " left join customerm imp on c.job_imp_id = imp.cust_pkid";
                sql += " left join addressm impaddr on c.job_imp_br_id = impaddr.add_pkid";
                sql += " left join param cmdty on c.job_commodity_id = cmdty.param_pkid";
                sql += " where a.rec_branch_code = '{BRCODE}'   and cntr_no in ('" + cntr_no + "')";
                sql += " order by cntr_no, job_docno";


                sql = sql.Replace("{BRCODE}", branch_code);

                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();
                string cargo_nature = "";

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new EgmReport();
                    cargo_nature = Dr["cargo_nature"].ToString();
                    if (cargo_nature == "N")
                    {
                        mrow.job_cargo_nature = "";
                    }
                    else
                    {
                        mrow.job_cargo_nature = Dr["cargo_nature"].ToString();
                    }

                    mrow.job_docno = Dr["job_docno"].ToString();
                    mrow.pol_code = Dr["pol_code"].ToString();
                    mrow.pol = Dr["sb_filed_location"].ToString();
                    mrow.egno = Dr["egno"].ToString();
                    mrow.egdate = Lib.DatetoStringDisplayformat(Dr["egdate"]);
                    mrow.opr_sbill_no = Dr["opr_sbill_no"].ToString();
                    mrow.opr_sbill_date = Lib.DatetoStringDisplayformat(Dr["opr_sbill_date"]);
                    mrow.job_qty = Lib.Conv2Decimal(Dr["quantity"].ToString());
                    mrow.pkg_unit = Dr["pkg_unit"].ToString();
                    mrow.job_pkg = Lib.Conv2Decimal(Dr["no_of_pkg"].ToString());
                    mrow.pofd_code = Dr["destination_port"].ToString();
                    mrow.cntr_no = Dr["cntr_no"].ToString();
                    mrow.exporter = Dr["exporter"].ToString();
                    mrow.exp_add1 = Dr["exp_add1"].ToString();
                    mrow.exp_add2 = Dr["exp_add2"].ToString();
                    mrow.exp_add3 = Dr["exp_add3"].ToString();
                    mrow.importer = Dr["importer"].ToString();
                    mrow.imp_add1 = Dr["imp_add1"].ToString();
                    mrow.imp_add2 = Dr["imp_add2"].ToString();
                    mrow.imp_add3 = Dr["imp_add3"].ToString();
                    mrow.commodity = Dr["cargo_discription"].ToString();
                    mrow.job_grwt = Lib.Conv2Decimal(Dr["gr_wt"].ToString());
                    mrow.shut_out_qty = Lib.Conv2Decimal(Dr["shut_out_quantity"].ToString());
                    mList.Add(mrow);

                }

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintEgmReport();
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

        private void PrintEgmReport()
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
                WS.Columns[4].Width = 256 * 20;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 25;
                WS.Columns[14].Width = 256 * 25;
                WS.Columns[15].Width = 256 * 25;
                WS.Columns[16].Width = 256 * 25;
                WS.Columns[17].Width = 256 * 25;
                WS.Columns[18].Width = 256 * 25;
                WS.Columns[19].Width = 256 * 25;
                WS.Columns[20].Width = 256 * 25;
                WS.Columns[21].Width = 256 * 20;
                WS.Columns[22].Width = 256 * 17;
                WS.Columns[23].Width = 256 * 30;
                WS.Columns[24].Width = 256 * 23;
                WS.Columns[25].Width = 256 * 17;

                iRow = 0; iCol = 1;
                WS.Columns[8].Style.NumberFormat = "#0.00";
                WS.Columns[10].Style.NumberFormat = "#0.00";

                WS.Columns[22].Style.NumberFormat = "#0.000";
                WS.Columns[23].Style.NumberFormat = "#0.000";

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
                Lib.WriteData(WS, iRow, 1, "EGM REPORT", _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 11;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "Sl.No", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB Filed Location", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EGM No", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "EGM Date", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB No. ", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SB Date", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Cargo Nature", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Quantity", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Unit", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "No.of Pkgs", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Destin Port", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Container", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Exporter Name", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Exporter Address-Line 1", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Exporter Address-Line 2", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Exporter Address-Line 3", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Consignee Name", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Consignee Address-Line 1", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Consignee Address-Line 2", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Consignee Address-Line 3", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Cargo Description", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Weight (KGs)", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "Shut out quantity, if any (KGs)", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                _Size = 10;
                Decimal si_no = 1;
                foreach (DataRow Dr in Dt_List.Rows)
                {
                    iRow++;
                    iCol = 1;

                    if (si_no == 1)
                    {
                        for (int i = 1; i <= 23; i++)
                        {
                            Lib.WriteData(WS, iRow, iCol++, "(" + i + ")", _Color, false, "", "C", "", _Size, false, 325, "", true);
                        }
                        iRow++; iCol = 1;
                    }
                    
                    // DateTime dt = ((DateTime) Dr["opr_sbill_date"]);
                    //  str = dt.ToString("MM/dd/yyyy");

                    string dt = Dr["opr_sbill_date"].ToString();
                    if(dt != "")
                    {
                        object sb_dt =  Dr["opr_sbill_date"];
                        sb_date = ((DateTime)sb_dt).ToString("MM/dd/yyyy");
                    }
                    else
                    {
                        sb_date = "";dt = "";
                    }

                    dt = Dr["egdate"].ToString();
                    if (dt != "")
                    {
                        object egm_dt = Dr["egdate"];
                        egm_date = ((DateTime)egm_dt).ToString("MM/dd/yyyy");
                    }
                    else
                    {
                        egm_date = "";dt = "";
                    }

                    string cargo_nature = Dr["cargo_nature"].ToString();
                    if (cargo_nature == "N")
                    {
                        cargo_nature = "";
                    }

                    cntr_no = Dr["cntr_no"].ToString();
                    cntr_no = cntr_no.Replace("-","");

                    Lib.WriteData(WS, iRow, iCol++, si_no++, _Color, false, "", "R", "", _Size, false, 325, "", true);                 
                    Lib.WriteData(WS, iRow, iCol++, Dr["pol_code"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["egno"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, egm_date, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["opr_sbill_no"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, sb_date, _Color, false, "", "L", "", _Size, false, 325, "", true);                 
                    Lib.WriteData(WS, iRow, iCol++, cargo_nature, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["quantity"], _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["pkg_unit"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["no_of_pkg"], _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["destination_port"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, cntr_no, _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["exporter"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["exp_add1"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["exp_add2"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["exp_add3"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["importer"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["imp_add1"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["imp_add2"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["imp_add3"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["cargo_discription"], _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["gr_wt"], _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["shut_out_quantity"], _Color, false, "", "R", "", _Size, false, 325, "", true);

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

