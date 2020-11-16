using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;

namespace BLAccounts
{
    public class InvoiceReportService : BaseReport
    {
        public DataTable DT_MBLINV = null;
        public DataTable DT_MBLOS = null;
        public DataTable DT_INVFOOTER = null;
        public string Report_format = "SUMMARY";
        public string DetailFormat = "SUMMARY";
        public string HeaderFormat = "SEA";
        public string report_folder = "";
        public string folderid = "";
        public string File_Name = "";
        public string File_Type = "";
        public string File_Display_Name = "myreport.pdf";
        public List<InvoiceReport> DetList = new List<InvoiceReport>();
        public InvoiceReport hRow;
        public Dictionary<string, string> BankInfoDic = new Dictionary<string, string>();
        public string company_code = "";
        public string branch_code = "";
        private string ImagePath = "";
        private int CurrentPageNo = 0;
        private int RowsPerPage = 25;
        private int TotalRows = 0;
        private int TotalPages = 0;
        private string str = "";
        public string ReportCaption = "";
        private string InvFormatType = "";
        private string Print_OnAcount = "";


        private float DCOL1 = 0;
        private float DCOL2 = 0;
        private float DCOL3 = 0;
        private float DCOL4 = 0;
        private float DCOL5 = 0;
        private float DCOL6 = 0;
        private float DCOL7 = 0;
        private float DCOL8 = 0;
        private float DCOL9 = 0;
        private float DCOL10 = 0;
        private float DCOL11 = 0;
        private float DCOL12 = 0;
        private float DCOL13 = 0;
        private float DCOL14 = 0;
        private float DCOL15 = 0;

        private string ctr = "";
        private string sac_code = "";
        private string acc_name = "";
        private string qty = "";
        private string rate = "";
        private string curr_code = "";
        private string exrate = "";
        private string ftotal = "";
        private string total = "";
        private string igst_rate = "";
        private string igst_amt = "";
        private string cgst_rate = "";
        private string cgst_amt = "";
        private string sgst_rate = "";
        private string sgst_amt = "";
        private string net_total = "";

        public void Process()
        {
            if (hRow == null)
                return;

            InvFormatType = hRow.jvh_type;//Print Format

            Print_OnAcount = "";

            if (HeaderFormat.Contains("IMPORT"))
            {
                if (hRow.jvh_acc_id.Length > 0 && hRow.hbl_imp_id.Length > 0)
                {
                    if (hRow.jvh_acc_id != hRow.hbl_imp_id)
                        Print_OnAcount = "ON A/C OF " + hRow.hbl_consignee_name;
                }
            }
            else
            {
                if (hRow.jvh_acc_id.Length > 0 && hRow.hbl_exp_id.Length > 0)
                {
                    if (hRow.jvh_acc_id != hRow.hbl_exp_id)
                        Print_OnAcount = "ON A/C OF " + hRow.hbl_exp_name;
                }
            }

            if (InvFormatType == "PN") // For payment invoice On A/c will nnot print
                Print_OnAcount = "";

            TotalRows = DetList.Count;
            TotalPages = Lib.Conv2Integer((Math.Ceiling(TotalRows / float.Parse(RowsPerPage.ToString()))).ToString());
            if (TotalPages <= 0)
                TotalPages = 1;

            ImagePath = report_folder + "\\Images";
            File_Display_Name = Lib.ProperFileName(hRow.jvh_docno) + ".pdf";
            File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
            File_Type = "pdf";

            BeginReport(1100, 800);
            int iIndex = 0;
            for (int i = 0; i < TotalPages; i++)
            {
                NewPage();
                if (i == 0)
                    iIndex = 0;
                else
                    iIndex = (i * RowsPerPage);
                for (int iStart = iIndex; iStart < (iIndex + RowsPerPage); iStart++)
                {
                    if (iStart < TotalRows)
                        WriteDetails(DetList[iStart]);
                    else
                        WriteDetails(null);
                }
                WriteTotal();
                WriteFooter();
            }
            EndReport();
            if (ExportList != null)
            {
                Export2Pdf mypdf = new Export2Pdf();
                mypdf.ExportList = ExportList;
                mypdf.FileName = File_Name;
                mypdf.Page_Height = 1100;
                mypdf.Page_Width = 800;
                mypdf.Process();
            }
        }
        private void NewPage()
        {
            CurrentPageNo++;
            InitPage();
            AddPage(1100, 800);
            if (InvFormatType != "PN")
            {
                if (company_code == "CPL")
                    LoadImage(ImagePath + "\\Logo.gif", 20, Row + 20, 68, 70);
                else if (company_code == "SGT")
                    LoadImage(ImagePath + "\\sgtLogo.png", 20, Row + 20, 68, 70);
            }
            WriteCompanyDetails();
            WriteInvoiceHeader();
            WriteDetailHeader();
        }
        private void InitPage()
        {
            HCOL1 = 10;
            HCOL2 = HCOL1 + 70;
            HCOL3 = HCOL2 + 10;
            HCOL4 = HCOL3 + 350;
            HCOL5 = HCOL4 + 75;
            HCOL6 = HCOL5 + 10;
            HCOL7 = HCOL6 + 225;

            ifontName = "Calibri";
            ifontSize = 9;

            Row = 10;
            ROW_HT = 15;
        }
        private void WriteCompanyDetails()
        {
            if (InvFormatType == "PN")
            {
                AddXYLabel(2, Row, ROW_HT, 800 - 4, hRow.inv_comp_name, ifontName, ifontSize + 6, "", "BC");
                Row += ROW_HT;
                str = hRow.inv_comp_add1;
                str += " " + hRow.inv_comp_add2;
                AddXYLabel(2, Row, ROW_HT, 800 - 4, str.Trim(), ifontName, ifontSize, "", "C");
                Row += ROW_HT;
                Row += ROW_HT;
                AddXYLabel(2, Row - 5, ROW_HT, 800 - 4, "MBL EXPENSE", ifontName, ifontSize + 6, "", "BC");
                AddXYLabel(2, Row, ROW_HT, 800 - 4, "", ifontName, ifontSize + 6, "B", "BC");
                AddXYLabel(650, Row, ROW_HT, 50, "Page# : ", ifontName, ifontSize, "", "BL");
                str = CurrentPageNo.ToString() + "/" + TotalPages;
                AddXYLabel(690, Row, ROW_HT, 50, str, ifontName, ifontSize, "", "L");
            }
            else
            {
                DrawHLine(2, 2, Page_Width - 4);
                DrawHLine(2, Page_Height - 2, Page_Width - 4);
                DrawVLine(2, 2, Page_Height - 4);
                DrawVLine(Page_Width - 2, 2, Page_Height - 4);
                AddXYLabel(2, Row, ROW_HT, 800 - 4, hRow.inv_comp_name, ifontName, ifontSize + 6, "", "BC");
                Row += ROW_HT;
                str = hRow.inv_comp_add1;
                str += " " + hRow.inv_comp_add2;
                AddXYLabel(2, Row, ROW_HT, 800 - 4, str.Trim(), ifontName, ifontSize, "", "C");
                str = hRow.inv_comp_add3;
                if (str != "")
                {
                    Row += ROW_HT;
                    AddXYLabel(2, Row, ROW_HT, 800 - 4, str.Trim(), ifontName, ifontSize, "", "C");
                }
                Row += ROW_HT;
                str = "";
                if (hRow.inv_comp_tel != "")
                    str = "TEL: " + hRow.inv_comp_tel;
                if (hRow.inv_comp_fax != "")
                    str += " FAX: " + hRow.inv_comp_fax;
                AddXYLabel(2, Row, ROW_HT, 800 - 4, str, ifontName, ifontSize, "", "C");
                Row += ROW_HT;
                str = "";
                if (hRow.inv_comp_email != "")
                    str = "E-MAIL: " + hRow.inv_comp_email.ToLower();
                if (hRow.inv_comp_web != "")
                    str += " WEB: " + hRow.inv_comp_web.ToLower();
                AddXYLabel(2, Row, ROW_HT, 800 - 4, str.Trim(), ifontName, ifontSize, "", "C");
                Row += ROW_HT;
                str = "";
                if (hRow.inv_comp_cinno != "")
                    str = "CIN#: " + hRow.inv_comp_cinno;
                AddXYLabel(2, Row, ROW_HT, 800 - 4, str, ifontName, ifontSize, "", "C");

                Row += ROW_HT;
                Row += ROW_HT;
                AddXYLabel(2, Row - 5, ROW_HT, 800 - 4, ReportCaption.ToUpper(), ifontName, ifontSize + 6, "", "BC");
                AddXYLabel(2, Row, ROW_HT, 800 - 4, "", ifontName, ifontSize + 6, "B", "BC");
                AddXYLabel(650, Row, ROW_HT, 50, "Page# : ", ifontName, ifontSize, "", "BL");
                str = CurrentPageNo.ToString() + "/" + TotalPages;
                AddXYLabel(690, Row, ROW_HT, 50, str, ifontName, ifontSize, "", "L");
            }
        }
        private void WriteInvoiceHeader()
        {
            //hrow1
            Row += ROW_HT;
            str = "INVOICE NO";
            if (InvFormatType == "PN")
                str = "OUR.REF#";
            if (InvFormatType == "DN")
                str = "DN. NO";
            if (InvFormatType == "CN")
                str = "CN. NO";

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, str, ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            
            str = hRow.jvh_docno;
            /*
            if (InvFormatType == "PN")
                str = hRow.jvh_reference;
                */
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, str, ifontName, ifontSize, "", "L");
            //if (Report_format == "FC")
            //{
            //    AddXYLabel(HCOL3 + 110, Row, ROW_HT, (HCOL3 + 150) - (HCOL3 + 100), "CURR.", ifontName, ifontSize, "", "LB");
            //    AddXYLabel(HCOL3 + 150, Row, ROW_HT, (HCOL3 + 160) - (HCOL3 + 150), ":", ifontName, ifontSize, "", "LB");
            //    AddXYLabel(HCOL3 + 160, Row, ROW_HT, HCOL4 - (HCOL3 + 160), hRow.jvh_curr_code, ifontName, ifontSize, "", "L");
            //}
            if (InvFormatType == "DN"|| InvFormatType == "CN")
            {
                AddXYLabel(HCOL3 + 110, Row, ROW_HT, (HCOL3 + 150) - (HCOL3 + 100), "INV#", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 150, Row, ROW_HT, (HCOL3 + 160) - (HCOL3 + 150), ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 160, Row, ROW_HT, HCOL4 - (HCOL3 + 160), hRow.jvh_org_invno, ifontName, ifontSize, "", "L");
            }

            str = "";
            if (HeaderFormat.Contains("SEA"))
                str = "MBL NO";
            else if (HeaderFormat.Contains("AIR"))
                str = "MAWB NO";
            if (str != "")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, str, ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                str = hRow.mbl_no;
                if (HeaderFormat.Contains("SEA EXPORT"))
                {
                    if (hRow.mbl_pol_etd.Trim() != "")
                        str += " / " + hRow.mbl_pol_etd;
                }
                else
                {
                    if (hRow.mbl_date.Trim() != "")
                        str += " / " + hRow.mbl_date;
                }
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");
            }
            if (HeaderFormat == "GENERAL JOB")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "BOOKING NO", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL6, Row, ROW_HT, (HCOL6 + 70) - HCOL6, hRow.hbl_genjob_no, ifontName, ifontSize, "", "L");
                /*
                AddXYLabel(HCOL6 + 70, Row, ROW_HT, (HCOL6 + 150) - (HCOL6 + 70), "BOOKING DATE", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL6 + 150, Row, ROW_HT, (HCOL6 + 160) - (HCOL6 + 150), ":", ifontName, ifontSize, "", "LB");
                str = hRow.hbl_date.ToString().Replace(".", "/");
                AddXYLabel(HCOL6 + 160, Row, ROW_HT, HCOL7 - (HCOL6 + 160), str, ifontName, ifontSize, "", "L");
                */
            }

            //hrow2
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "DATE", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, (HCOL3 + 110) - HCOL3, hRow.jvh_date, ifontName, ifontSize, "", "L");//"07/03/2018 SI# 1986"
            if (InvFormatType == "PN")
            {
                AddXYLabel(HCOL3 + 90, Row, ROW_HT, (HCOL3 + 150) - (HCOL3 + 100), "ORG. INV#", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 150, Row, ROW_HT, (HCOL3 + 160) - (HCOL3 + 150), ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 160, Row, ROW_HT, HCOL4 - (HCOL3 + 160), hRow.jvh_org_invno, ifontName, ifontSize, "", "L");
            }
            else
            {
                if (hRow.hbl_no != "" && HeaderFormat != "GENERAL JOB")
                {
                    if (Report_format == "FC")
                        hRow.hbl_no += "     CURR. : " + hRow.jvh_curr_code;
                    AddXYLabel(HCOL3 + 110, Row, ROW_HT, (HCOL3 + 150) - (HCOL3 + 100), "SI#", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL3 + 150, Row, ROW_HT, (HCOL3 + 160) - (HCOL3 + 150), ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL3 + 160, Row, ROW_HT, HCOL4 - (HCOL3 + 160), hRow.hbl_no, ifontName, ifontSize, "", "L");
                }
            }
            str = "";
            if (HeaderFormat == "SI SEA EXPORT" || HeaderFormat == "SI SEA IMPORT")
                str = "HBL NO";
            else if (HeaderFormat == "SI AIR EXPORT" || HeaderFormat == "SI AIR IMPORT")
                str = "HAWB NO";
            if (str != "")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, str, ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                str = hRow.hbl_bl_no;
                if (hRow.hbl_date.Trim() != "")
                    str += " / " + hRow.hbl_date;
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");
            }

            if (HeaderFormat == "GENERAL JOB")
            {
                if (hRow.hbl_genigm_no != "" && hRow.hbl_genjobtype_code == "TRANSPORT")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "IGM / ITEM NO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.hbl_genigm_no, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_hbl_no != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "HBL NO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_hbl_no, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_lr_no != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "LR NO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_lr_no, ifontName, ifontSize, "", "L");
                }
            }
            else
            {
                if (InvFormatType == "PN")// HeaderFormat start wilth M ie Master Costcenter
                {
                    str = "MBLBK#";
                    if (HeaderFormat.Contains("AIR"))
                        str = "MAWBK#";
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, str, ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.mbl_bkno, ifontName, ifontSize, "", "L");
                }

            }

            //hrow3
            Row += ROW_HT;
            str = "M/S";
            if (InvFormatType == "PN")
                str = "FROM";
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, str, ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, hRow.jvh_party_name, ifontName, ifontSize, "", "L");
            if (HeaderFormat == "SI SEA EXPORT")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "CHA", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.mbl_cha_name, ifontName, ifontSize, "", "L");
            }
            else if (HeaderFormat == "SI SEA IMPORT" || HeaderFormat == "SI AIR IMPORT")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "IGM NO", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                str = hRow.mbl_igmno;
                if (str.Trim() != "" && hRow.mbl_igmdate != "")
                    str += " / " + hRow.mbl_igmdate;
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");
            }
            else if (HeaderFormat == "GENERAL JOB")
            {
                if (hRow.gj_cfs != "" && hRow.hbl_genjobtype_code == "TRANSPORT")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "CFS", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_cfs, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_mbl_no != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "MBL NO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_mbl_no, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_pack_list_no != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "PACK LIST NO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_pack_list_no, ifontName, ifontSize, "", "L");
                }
            }

            //hrow4
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, hRow.jvh_party_addr1, ifontName, ifontSize, "", "L");
            if (HeaderFormat == "SI SEA EXPORT")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "CONSIGNEE", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.hbl_consignee_name, ifontName, ifontSize, "", "L");
            }
            else if (HeaderFormat == "SI SEA IMPORT" || HeaderFormat == "SI AIR IMPORT")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "BE NO", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                str = hRow.hbl_beno;
                if (str.Trim() != "" && hRow.hbl_bedate != "")
                    str += " / " + hRow.hbl_bedate;
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");
            }
            else if (HeaderFormat == "GENERAL JOB")
            {
                if (hRow.gj_loaded_on != "" && hRow.hbl_genjobtype_code == "TRANSPORT")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "LOADED ON", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, (HCOL6 + 70) - HCOL6, hRow.gj_loaded_on, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_unloaded_on != "" && hRow.hbl_genjobtype_code == "TRANSPORT")
                {
                    AddXYLabel(HCOL6 + 70, Row, ROW_HT, (HCOL6 + 150) - (HCOL6 + 70), "UNLOADED ON", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6 + 150, Row, ROW_HT, (HCOL6 + 160) - (HCOL6 + 150), ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6 + 160, Row, ROW_HT, HCOL7 - (HCOL6 + 160), hRow.gj_unloaded_on, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_vessel != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "VESSEL", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_vessel, ifontName, ifontSize, "", "L");
                }
            }

            //hrow5
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
            str = hRow.jvh_party_addr2;
            if (str.Trim() == "")
            {
                str = Print_OnAcount;
                Print_OnAcount = "";
            }
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, str, ifontName, ifontSize, "", "L");

            str = "";
            if (HeaderFormat == "SI SEA EXPORT" || HeaderFormat == "SI SEA IMPORT")
                str = "VESSEL";
            else if (HeaderFormat == "SI AIR EXPORT" || HeaderFormat == "SI AIR IMPORT")
                str = "FLIGHT";
            if (str != "")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, str, ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                str = hRow.mbl_vessel_name + " " + hRow.mbl_vessel_voyage;
                if (str.Trim() != "" && hRow.mbl_pol_etd.Trim() != "")
                {
                    str += " / " + hRow.mbl_pol_etd;
                }
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str.Trim(), ifontName, ifontSize, "", "L");
            }
            if (HeaderFormat == "GENERAL JOB")
            {
                if (hRow.gj_frt_status != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "STATUS", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_frt_status, ifontName, ifontSize, "", "L");
                }
            }


            //hrow6
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
            str = hRow.jvh_party_addr3;
            if(str.Trim()=="")
            {
                str = Print_OnAcount;
                Print_OnAcount = "";
            }
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, str, ifontName, ifontSize, "", "L");
            if (hRow.hbl_pol_name != "")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "POL", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.hbl_pol_name, ifontName, ifontSize, "", "L");
            }
            //hrow7
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, Print_OnAcount, ifontName, ifontSize, "", "L");
            if (hRow.hbl_pod_name != "")
            {
                str = hRow.hbl_pod_name;
                if (hRow.hbl_pofd_name != "" && hRow.hbl_pod_name != hRow.hbl_pofd_name)
                    str += " / " + hRow.hbl_pofd_name;
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "POD", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");
            }

            //hrow8
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "", ifontName, ifontSize, "", "L");
            if (HeaderFormat == "GENERAL JOB")
            {
                if (hRow.gj_from != "" && hRow.hbl_genjobtype_code == "TRANSPORT")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "FROM", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_from, ifontName, ifontSize, "", "L");
                }
            }
            else
            {
                if (hRow.hbl_freight_status != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "STATUS", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.hbl_freight_status, ifontName, ifontSize, "", "L");
                }
            }

            //hrow9
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "GST NO", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3, Row, ROW_HT, (HCOL3 + 110) - HCOL3, hRow.jvh_gstin, ifontName, ifontSize, "", "L");
            if (hRow.jvh_state_name != "")
            {
                str = hRow.jvh_state_code + "-" + hRow.jvh_state_name;
                AddXYLabel(HCOL3 + 110, Row, ROW_HT, (HCOL3 + 150) - (HCOL3 + 100), "SC-POS", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 150, Row, ROW_HT, (HCOL3 + 160) - (HCOL3 + 150), ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 160, Row, ROW_HT, HCOL4 - (HCOL3 + 160), str, ifontName, ifontSize, "", "L");
            }
            if (HeaderFormat == "SI AIR EXPORT" || HeaderFormat == "SI AIR IMPORT")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "PACKING", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (Lib.Conv2Decimal(hRow.hbl_pkg.ToString()) != 0)
                {
                    str = hRow.hbl_pkg.ToString();
                    str += " " + hRow.hbl_pkg_unit;
                }

                if (Lib.Conv2Decimal(hRow.hbl_chwt.ToString()) > 0)
                {
                    if (str.Trim() != "")
                        str += " / ";
                    str += "CHWT: ";
                    str += Lib.NumericFormat(hRow.hbl_chwt.ToString(), 3);
                    str += " " + hRow.hbl_wt_unit;
                }
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");

                if (branch_code == "MBYAF")
                {
                    str = "";
                    if (Lib.Conv2Decimal(hRow.hbl_ntwt.ToString()) > 0)
                    {
                        str = "NTWT: ";
                        str += Lib.NumericFormat(hRow.hbl_ntwt.ToString(), 3);
                    }
                    if (Lib.Conv2Decimal(hRow.hbl_grwt.ToString()) > 0)
                    {
                        if (str.Trim() != "")
                            str += " / ";
                        str += "GRWT: ";
                        str += Lib.NumericFormat(hRow.hbl_grwt.ToString(), 3);
                    }
                    if (str.Trim() != "")
                    {
                        str += " " + hRow.hbl_wt_unit;
                        if (hRow.job_comm_invnos != "")
                        {
                            Row += ROW_HT;
                            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");
                        }
                        else
                            AddXYLabel(HCOL6, Row + ROW_HT, ROW_HT, HCOL7 - HCOL6, str, ifontName, ifontSize, "", "L");
                    }
                }
            }
            else if (HeaderFormat == "SI SEA EXPORT" || HeaderFormat == "SI SEA IMPORT")
            {
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "PACKING", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (Lib.Conv2Decimal(hRow.hbl_pkg.ToString()) != 0)
                {
                    str = hRow.hbl_pkg.ToString();
                    if (str.Trim() != "")
                        str += " " + hRow.hbl_pkg_unit;
                }
                AddXYLabel(HCOL6, Row, ROW_HT, (HCOL6 + 100) - HCOL6, str, ifontName, ifontSize, "", "L");

                AddXYLabel(HCOL6 + 100, Row, ROW_HT, (HCOL6 + 150) - (HCOL6 + 100), "VOLUME", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL6 + 150, Row, ROW_HT, (HCOL6 + 160) - (HCOL6 + 150), ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (Lib.Conv2Decimal(hRow.hbl_cbm.ToString()) != 0)
                    str = hRow.hbl_cbm.ToString();
                AddXYLabel(HCOL6 + 160, Row, ROW_HT, HCOL7 - (HCOL6 + 160), str, ifontName, ifontSize, "", "L");
            }
            else if (HeaderFormat == "GENERAL JOB")
            {
                if (hRow.gj_to1 != "" && hRow.hbl_genjobtype_code == "TRANSPORT")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "TO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_to1, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_cha_name != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "CHA", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_cha_name, ifontName, ifontSize, "", "L");
                }
            }

            if (HeaderFormat == "GENERAL JOB")
            {  //hrow10
                if (hRow.gj_to2 != "" && hRow.hbl_genjobtype_code == "TRANSPORT")
                {
                    Row += ROW_HT;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "", ifontName, ifontSize, "", "L");

                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "TO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_to2, ifontName, ifontSize, "", "L");
                }
                if (hRow.gj_consignee_name != "")
                {
                    Row += ROW_HT;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "", ifontName, ifontSize, "", "L");

                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "CONSIGNEE", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.gj_consignee_name, ifontName, ifontSize, "", "L");
                }
                //hrow11
                if (hRow.hbl_carton_nos != "" || Lib.Conv2Decimal(hRow.hbl_grwt.ToString()) > 0)
                {
                    Row += ROW_HT;
                    if (hRow.hbl_carton_nos != "")
                    {
                        AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "CARTONS", ifontName, ifontSize, "", "LB");
                        AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                        AddXYLabel(HCOL6, Row, ROW_HT, (HCOL6 + 70) - HCOL6, hRow.hbl_carton_nos, ifontName, ifontSize, "", "L");
                    }
                    if (Lib.Conv2Decimal(hRow.hbl_grwt.ToString()) > 0)
                    {
                        str = Lib.NumericFormat(hRow.hbl_grwt.ToString(), 3);
                        if (str.Trim() != "")
                            str += " KGS";
                        AddXYLabel(HCOL6 + 70, Row, ROW_HT, (HCOL6 + 150) - (HCOL6 + 70), "GR.WT.", ifontName, ifontSize, "", "LB");
                        AddXYLabel(HCOL6 + 150, Row, ROW_HT, (HCOL6 + 160) - (HCOL6 + 150), ":", ifontName, ifontSize, "", "LB");
                        AddXYLabel(HCOL6 + 160, Row, ROW_HT, HCOL7 - (HCOL6 + 160), str, ifontName, ifontSize, "", "L");
                    }
                }
            }
           
            if (HeaderFormat == "SI SEA EXPORT" || HeaderFormat == "SI AIR EXPORT")
            {
                //hrow11
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "SBILLNOS", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, hRow.job_sbnos, ifontName, ifontSize, "", "L");
                if (hRow.job_comm_invnos != "")
                {
                    AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "COMM/INNO", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, hRow.job_comm_invnos, ifontName, ifontSize, "", "L");
                }
            }
            if (HeaderFormat == "GENERAL JOB")
            {
                //hrow11
                if (hRow.gj_sb_no != "")
                {
                    Row += ROW_HT;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "SBILLNOS", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, hRow.gj_sb_no, ifontName, ifontSize, "", "L");
                }
               
            }
            if (hRow.job_invnos != "" || hRow.gj_shipper_inv_no != "" || hRow.hbl_invoice_nos != "")
            {
                //hrow13
                str = hRow.job_invnos;//Export Job invoice
                if (HeaderFormat == "SI SEA IMPORT" || HeaderFormat == "SI AIR IMPORT")
                    str = hRow.hbl_invoice_nos;
                if (HeaderFormat == "GENERAL JOB")
                    str = hRow.gj_shipper_inv_no;

                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "INVNOS", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                // AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, str, ifontName, ifontSize, "", "L");//Change on 25/02/2019 Ajith
                if (str.Contains(","))
                {
                    int Linecount = 0;
                    int InvPerLine = 5;
                    bool bNextRow = false;
                    string[] InvArry;

                    if (hRow.hbl_invnos_prncount > 0)
                        InvPerLine = hRow.hbl_invnos_prncount;
                   
                    InvArry = str.Split(',');
                    if (InvArry.Length <= 10)//If less than 10 invoice will print in multiple line other wise print in a single line
                    {
                        str = "";
                        for (int i = 1; i <= InvArry.Length; i++)
                        {
                            if (str != "")
                                str += ",";
                            str += InvArry[i - 1].Trim();
                            if (i % InvPerLine == 0)
                            {
                                if (bNextRow)
                                    Row += ROW_HT;
                                AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, str, ifontName, ifontSize, "", "L");
                                str = "";
                                bNextRow = true;
                                Linecount++;
                                if (Linecount >= 2)
                                    break;
                            }
                        }
                    } 
                    if (str != "")
                    {
                        if (bNextRow)
                            Row += ROW_HT;
                        AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, str, ifontName, ifontSize, "", "L");
                    }
                }
                else
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, str, ifontName, ifontSize, "", "L");

            }
            if (HeaderFormat == "GENERAL JOB" && hRow.jvh_narration != "")//&& branch_code == "SEZSF"
            {
                //  hrow12
                string[] RemArry;
                int sWidth = (int)(HCOL7 - HCOL3);
                RemArry = Lib.ConvertString2Lines(hRow.jvh_narration, sWidth, "WORD");
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "REMARKS", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, RemArry.Length > 0 ? RemArry[0].Trim() : "", ifontName, ifontSize, "", "L");
                for (int i = 1; i < RemArry.Length; i++)
                {
                    Row += ROW_HT;
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, RemArry.Length > i ? RemArry[i].Trim() : "", ifontName, ifontSize, "", "L");
                }

            }
            
            if (hRow.gj_seal_no != "")
            {
                //hrow14
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "SEAL NO", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, hRow.gj_seal_no, ifontName, ifontSize, "", "L");
            }
            //if (hRow.hbl_containers != "")
            if (HeaderFormat.Contains("SEA") || HeaderFormat == "GENERAL JOB")
            {
                //hrow15
                if (HeaderFormat == "GENERAL JOB")
                {
                    if (hRow.hbl_containers != "")
                    {
                        Row += ROW_HT;
                        AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "CONTAINER", ifontName, ifontSize, "", "LB");
                        AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                        AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, hRow.hbl_containers, ifontName, ifontSize, "", "L");
                    }
                }
                else
                {
                    Row += ROW_HT;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "CONTAINER", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, hRow.hbl_containers, ifontName, ifontSize, "", "L");
                }
            }
            if (hRow.job_commodity != "")
            {
                //hrow116
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "COMMODITY", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, hRow.job_commodity, ifontName, ifontSize, "", "L");
            }
            Row += ROW_HT;
            //Row += ROW_HT;
        }
        
        private void WriteDetailHeader()
        {
            if (DetailFormat == "SUMMARY")
            {
                DCOL1 = 10;
                DCOL2 = DCOL1 + 40;
                DCOL3 = DCOL2 + 640;
                DCOL4 = DCOL3 + 100;

                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "SL#", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "PARTICULARS", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
            }
            else if (DetailFormat == "SUMMARY-CS-GST")
            {
                //DCOL1 = 10;
                //DCOL2 = DCOL1 + 20;
                //DCOL3 = DCOL2 + 50;
                //DCOL4 = DCOL3 + 370;
                //DCOL5 = DCOL4 + 70;
                //DCOL6 = DCOL5 + 40;
                //DCOL7 = DCOL6 + 60;
                //DCOL8 = DCOL7 + 40;
                //DCOL9 = DCOL8 + 60;
                //DCOL10 = DCOL9 + 70;

                DCOL1 = 10;
                DCOL2 = DCOL1 + 20;
                DCOL3 = DCOL2 + 40;
                DCOL4 = DCOL3 + 370;
                DCOL5 = DCOL4 + 70;
                DCOL6 = DCOL5 + 40;
                DCOL7 = DCOL6 + 65;
                DCOL8 = DCOL7 + 40;
                DCOL9 = DCOL8 + 65;
                DCOL10 = DCOL9 + 70;

                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "SL#", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "SAC", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "PARTICULARS", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "CGST%", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "CGST", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "SGST%", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "SGST", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, "TOTAL", ifontName, ifontSize, "LTBR", "CB");

            }
            else if (DetailFormat == "SUMMARY-I-GST")
            {
                //DCOL1 = 10;
                //DCOL2 = DCOL1 + 20;
                //DCOL3 = DCOL2 + 45;
                //DCOL4 = DCOL3 + 455;
                //DCOL5 = DCOL4 + 80;
                //DCOL6 = DCOL5 + 40;
                //DCOL7 = DCOL6 + 60;
                //DCOL8 = DCOL7 + 80;

                DCOL1 = 10;
                DCOL2 = DCOL1 + 20;
                DCOL3 = DCOL2 + 45;
                DCOL4 = DCOL3 + 455;
                DCOL5 = DCOL4 + 80;
                DCOL6 = DCOL5 + 35;
                DCOL7 = DCOL6 + 65;
                DCOL8 = DCOL7 + 80;

                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "SL#", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "SAC", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "PARTICULARS", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "IGST%", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "IGST", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "TOTAL", ifontName, ifontSize, "LTBR", "CB");
            }
            else if (DetailFormat == "DETAIL")
            {
                DCOL1 = 10;
                DCOL2 = DCOL1 + 20;
                DCOL3 = DCOL2 + 465;
                DCOL4 = DCOL3 + 40;
                DCOL5 = DCOL4 + 50;
                DCOL6 = DCOL5 + 25;
                DCOL7 = DCOL6 + 70;
                DCOL8 = DCOL7 + 40;
                DCOL9 = DCOL8 + 70;

                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "SL#", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "PARTICULARS", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "QTY", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "RATE", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "CUR", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "EXRATE", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
            }
            else if (DetailFormat == "DETAIL-CS-GST")
            {
                //DCOL1 = 10;
                //DCOL2 = DCOL1 + 20;
                //DCOL3 = DCOL2 + 50;
                //DCOL4 = DCOL3 + 150;
                //DCOL5 = DCOL4 + 40;
                //DCOL6 = DCOL5 + 65;
                //DCOL7 = DCOL6 + 25;
                //DCOL8 = DCOL7 + 65;
                //DCOL9 = DCOL8 + 45;
                //DCOL10 = DCOL9 + 70;
                //DCOL11 = DCOL10 + 40;
                //DCOL12 = DCOL11 + 50;
                //DCOL13 = DCOL12 + 40;
                //DCOL14 = DCOL13 + 50;
                //DCOL15 = DCOL14 + 70;

                DCOL1 = 10;
                DCOL2 = DCOL1 + 20;
                DCOL3 = DCOL2 + 40;
                DCOL4 = DCOL3 + 150;
                DCOL5 = DCOL4 + 40;
                DCOL6 = DCOL5 + 65;
                DCOL7 = DCOL6 + 25;
                DCOL8 = DCOL7 + 65;
                DCOL9 = DCOL8 + 45;
                DCOL10 = DCOL9 + 70;
                DCOL11 = DCOL10 + 35;
                DCOL12 = DCOL11 + 60;
                DCOL13 = DCOL12 + 35;
                DCOL14 = DCOL13 + 60;
                DCOL15 = DCOL14 + 70;

                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "SL#", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "SAC", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "PARTICULARS", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "QTY", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "RATE", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "CUR", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "EX-RT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, "CGST%", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, "CGST", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, "SGST%", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL13, Row, ROW_HT, DCOL14 - DCOL13, "SGST", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL14, Row, ROW_HT, DCOL15 - DCOL14, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
            }
            else if (DetailFormat == "DETAIL-I-GST")
            {
                //DCOL1 = 10;
                //DCOL2 = DCOL1 + 20;
                //DCOL3 = DCOL2 + 50;
                //DCOL4 = DCOL3 + 250;
                //DCOL5 = DCOL4 + 40;
                //DCOL6 = DCOL5 + 55;
                //DCOL7 = DCOL6 + 25;
                //DCOL8 = DCOL7 + 70;
                //DCOL9 = DCOL8 + 40;
                //DCOL10 = DCOL9 + 70;
                //DCOL11 = DCOL10 + 40;
                //DCOL12 = DCOL11 + 50;
                //DCOL13 = DCOL12 + 70;

                DCOL1 = 10;
                DCOL2 = DCOL1 + 20;
                DCOL3 = DCOL2 + 40;
                DCOL4 = DCOL3 + 250;
                DCOL5 = DCOL4 + 40;
                DCOL6 = DCOL5 + 55;
                DCOL7 = DCOL6 + 25;
                DCOL8 = DCOL7 + 70;
                DCOL9 = DCOL8 + 40;
                DCOL10 = DCOL9 + 70;
                DCOL11 = DCOL10 + 35;
                DCOL12 = DCOL11 + 65;
                DCOL13 = DCOL12 + 70;

                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "SL#", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "SAC", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "PARTICULARS", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "QTY", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "RATE", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "CUR", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "EXRATE", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, "IGST%", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, "IGST", ifontName, ifontSize, "LTBR", "CB");
                AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, "AMOUNT", ifontName, ifontSize, "LTBR", "CB");
            }
        }

        private void WriteDetails(InvoiceReport Rec)
        {
            ctr = ""; sac_code = ""; acc_name = ""; qty = ""; rate = ""; curr_code = ""; exrate = "";
            ftotal = ""; total = ""; igst_rate = ""; igst_amt = ""; cgst_rate = ""; cgst_amt = ""; sgst_rate = "";
            sgst_amt = ""; net_total = "";
            if (Rec != null)
            {
                ctr = Lib.Conv2Integer(Rec.jv_ctr.ToString()) != 0 ? Rec.jv_ctr.ToString() : "";
                sac_code = Rec.jv_sac_code;
                acc_name = Rec.jv_acc_name;
                curr_code = Rec.jv_curr_code;
                if (Report_format == "FC")
                {
                    total = Lib.Conv2Decimal(Rec.jv_total_fc.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_total_fc.ToString(), 2) : "";
                    igst_rate = Lib.Conv2Decimal(Rec.jv_igst_rate.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_igst_rate.ToString(), 2) : "";
                    igst_amt = Lib.Conv2Decimal(Rec.jv_igst_famt.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_igst_famt.ToString(), 2) : "";
                    cgst_rate = Lib.Conv2Decimal(Rec.jv_cgst_rate.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_cgst_rate.ToString(), 2) : "";
                    cgst_amt = Lib.Conv2Decimal(Rec.jv_cgst_famt.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_cgst_famt.ToString(), 2) : "";
                    sgst_rate = Lib.Conv2Decimal(Rec.jv_sgst_rate.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_sgst_rate.ToString(), 2) : "";
                    sgst_amt = Lib.Conv2Decimal(Rec.jv_sgst_famt.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_sgst_famt.ToString(), 2) : "";
                    net_total = Lib.Conv2Decimal(Rec.jv_net_ftotal.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_net_ftotal.ToString(), 2) : "";
                }
                else
                {
                    qty = Lib.Conv2Decimal(Rec.jv_qty.ToString()) != 0 ? Rec.jv_qty.ToString() : "";
                    rate = Lib.Conv2Decimal(Rec.jv_rate.ToString()) != 0 ? Rec.jv_rate.ToString() : "";
                    ftotal = Lib.Conv2Decimal(Rec.jv_ftotal.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_ftotal.ToString(), 2) : "";
                    exrate = Lib.Conv2Decimal(Rec.jv_exrate.ToString()) != 0 ? Rec.jv_exrate.ToString() : "";
                    total = Lib.Conv2Decimal(Rec.jv_total.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_total.ToString(), 2) : "";
                    igst_rate = Lib.Conv2Decimal(Rec.jv_igst_rate.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_igst_rate.ToString(), 2) : "";
                    igst_amt = Lib.Conv2Decimal(Rec.jv_igst_amt.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_igst_amt.ToString(), 2) : "";
                    cgst_rate = Lib.Conv2Decimal(Rec.jv_cgst_rate.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_cgst_rate.ToString(), 2) : "";
                    cgst_amt = Lib.Conv2Decimal(Rec.jv_cgst_amt.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_cgst_amt.ToString(), 2) : "";
                    sgst_rate = Lib.Conv2Decimal(Rec.jv_sgst_rate.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_sgst_rate.ToString(), 2) : "";
                    sgst_amt = Lib.Conv2Decimal(Rec.jv_sgst_amt.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_sgst_amt.ToString(), 2) : "";
                    net_total = Lib.Conv2Decimal(Rec.jv_net_total.ToString()) != 0 ? Lib.NumericFormat(Rec.jv_net_total.ToString(), 2) : "";
                }
            }

            Row += ROW_HT;
            if (DetailFormat == "SUMMARY")
            {
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, ctr, ifontName, ifontSize, "L", "C");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, acc_name, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, net_total, ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);
            }
            else if (DetailFormat == "SUMMARY-CS-GST")
            {
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, ctr, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, sac_code, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, acc_name, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, total, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, cgst_rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, cgst_amt, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, sgst_rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, sgst_amt, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, net_total, ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);

            }
            else if (DetailFormat == "SUMMARY-I-GST")
            {
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, ctr, ifontName, ifontSize, "L", "C");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, sac_code, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, acc_name, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, total, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, igst_rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, igst_amt, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, net_total, ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);

            }
            else if (DetailFormat == "DETAIL")
            {
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, ctr, ifontName, ifontSize, "L", "C");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, acc_name, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, qty, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, curr_code, ifontName, ifontSize, "L", "C");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, ftotal, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, exrate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, net_total, ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);
            }
            else if (DetailFormat == "DETAIL-CS-GST")
            {
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, ctr, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, sac_code, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, acc_name, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, qty, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, curr_code, ifontName, ifontSize, "L", "C");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, ftotal, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, exrate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, total, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, cgst_rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, cgst_amt, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, sgst_rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL13, Row, ROW_HT, DCOL14 - DCOL13, sgst_amt, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL14, Row, ROW_HT, DCOL15 - DCOL14, net_total, ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);
            }
            else if (DetailFormat == "DETAIL-I-GST")
            {
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, ctr, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, sac_code, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, acc_name, ifontName, ifontSize, "L", "L");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, qty, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, curr_code, ifontName, ifontSize, "L", "C");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, ftotal, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, exrate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, total, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, igst_rate, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, igst_amt, ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, net_total, ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);
            }
        }
        private void WriteTotal()
        {
            string AmtInWords = "";
            string tot_amt = "", cgst_amt = "", sgst_amt = "", igst_amt = "", net_amt = "";
            if (Report_format == "FC")
            {
                tot_amt = Lib.Conv2Decimal(hRow.jvh_tot_famt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_tot_famt.ToString(), 2) : "";
                igst_amt = Lib.Conv2Decimal(hRow.jvh_igst_famt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_igst_famt.ToString(), 2) : "";
                cgst_amt = Lib.Conv2Decimal(hRow.jvh_cgst_famt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_cgst_famt.ToString(), 2) : "";
                sgst_amt = Lib.Conv2Decimal(hRow.jvh_sgst_famt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_sgst_famt.ToString(), 2) : "";
                net_amt = Lib.Conv2Decimal(hRow.jvh_net_famt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_net_famt.ToString(), 2) : "";
                AmtInWords = Number2Word_USD.Convert(net_amt, hRow.jvh_curr_code, Lib.GetDecimalName(hRow.jvh_curr_code));
            }
            else
            {
                tot_amt = Lib.Conv2Decimal(hRow.jvh_tot_amt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_tot_amt.ToString(), 2) : "";
                igst_amt = Lib.Conv2Decimal(hRow.jvh_igst_amt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_igst_amt.ToString(), 2) : "";
                cgst_amt = Lib.Conv2Decimal(hRow.jvh_cgst_amt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_cgst_amt.ToString(), 2) : "";
                sgst_amt = Lib.Conv2Decimal(hRow.jvh_sgst_amt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_sgst_amt.ToString(), 2) : "";
                net_amt = Lib.Conv2Decimal(hRow.jvh_net_amt.ToString()) != 0 ? Lib.NumericFormat(hRow.jvh_net_amt.ToString(), 2) : "";
                AmtInWords = Number2Word_RS.Convert(net_amt, "INR", Lib.GetDecimalName("INR"));
            }
            Row += ROW_HT;
            if (DetailFormat == "SUMMARY")
            {
                if (CurrentPageNo == TotalPages)
                {
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "TOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, net_amt, ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL4 - DCOL1, AmtInWords, ifontName, ifontSize, "LTBR", "L");
                }
                else
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL4 - DCOL1, "", ifontName, ifontSize, "T", "L");
            }
            else if (DetailFormat == "SUMMARY-CS-GST")
            {
                if (CurrentPageNo == TotalPages)
                {
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "TOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, tot_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, cgst_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, sgst_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, net_amt, ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "GRANDTOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, net_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, "", ifontName, ifontSize, "TBR", "LB");
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL10 - DCOL1, AmtInWords, ifontName, ifontSize, "LTBR", "L", 0, 0, 0, 0, 0, 0, 0);
                }
                else
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL10 - DCOL1, "", ifontName, ifontSize, "T", "LB");
            }
            else if (DetailFormat == "SUMMARY-I-GST")
            {
                if (CurrentPageNo == TotalPages)
                {
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "TOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, tot_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, igst_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, net_amt, ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "GRANDTOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, net_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "LB");
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TBR", "LB");
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL8 - DCOL1, AmtInWords, ifontName, ifontSize, "LTBR", "L");
                }
                else
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL8 - DCOL1, "", ifontName, ifontSize, "T", "LB");
            }
            else if (DetailFormat == "DETAIL")
            {
                if (CurrentPageNo == TotalPages)
                {
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "TOTAL", ifontName, ifontSize, "LTB", "L");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, net_amt, ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL9 - DCOL1, AmtInWords, ifontName, ifontSize, "LTBR", "L");
                }
                else
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL9 - DCOL1, "", ifontName, ifontSize, "T", "LB");

            }
            else if (DetailFormat == "DETAIL-CS-GST")
            {
                if (CurrentPageNo == TotalPages)
                {
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "TOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TB", "R");
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, tot_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, "", ifontName, ifontSize, "LTB", "L");
                    AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, cgst_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, "", ifontName, ifontSize, "LTB", "L");
                    AddXYLabel(DCOL13, Row, ROW_HT, DCOL14 - DCOL13, sgst_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL14, Row, ROW_HT, DCOL15 - DCOL14, net_amt, ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);

                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "GRANDTOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, net_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, "", ifontName, ifontSize, "LTB", "L");
                    AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL13, Row, ROW_HT, DCOL14 - DCOL13, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL14, Row, ROW_HT, DCOL15 - DCOL14, "", ifontName, ifontSize, "TBR", "L");
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL15 - DCOL1, AmtInWords, ifontName, ifontSize, "LTBR", "L");
                }
                else
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL15 - DCOL1, "", ifontName, ifontSize, "T", "LB");

            }
            else if (DetailFormat == "DETAIL-I-GST")
            {
                if (CurrentPageNo == TotalPages)
                {
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "TOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, tot_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, "", ifontName, ifontSize, "LTB", "L");
                    AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, igst_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, net_amt, ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);

                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "GRANDTOTAL", ifontName, ifontSize, "LTB", "LB");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, net_amt, ifontName, ifontSize, "LTB", "RB", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL10, Row, ROW_HT, DCOL11 - DCOL10, "", ifontName, ifontSize, "LTB", "L");
                    AddXYLabel(DCOL11, Row, ROW_HT, DCOL12 - DCOL11, "", ifontName, ifontSize, "TB", "L");
                    AddXYLabel(DCOL12, Row, ROW_HT, DCOL13 - DCOL12, "", ifontName, ifontSize, "TBR", "L");
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL13 - DCOL1, AmtInWords, ifontName, ifontSize, "LTBR", "L");
                }
                else
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL13 - DCOL1, "", ifontName, ifontSize, "T", "LB");

            }
        }

        private void WriteFooter()
        {
            if (InvFormatType == "PN")
            {
                WriteExpenseFooter();
                return;
            }
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "E.&O.E", ifontName, ifontSize + 2, "", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, (HCOL7 - HCOL2) - 50, hRow.inv_comp_name, ifontName, ifontSize + 2, "", "RB");
            if (BankInfoDic.Count > 0)
            {
                Row += ROW_HT;
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "BANK ACCOUNT DETAILS", ifontName, ifontSize, "", "LBU");
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "BENEFICIARY NAME", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2 + 50, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (BankInfoDic.ContainsKey("BANK_COMPANY"))
                    str = BankInfoDic["BANK_COMPANY"].ToString();
                AddXYLabel(HCOL3 + 50, Row, ROW_HT, HCOL6 - HCOL3, str, ifontName, ifontSize, "", "L");
                Row += ROW_HT;
                if (Report_format == "FC")
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "USD A/C #", ifontName, ifontSize, "", "LB");
                else
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "A/C #", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2 + 50, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (BankInfoDic.ContainsKey("BANK_ACNO"))
                    str = BankInfoDic["BANK_ACNO"].ToString();
                AddXYLabel(HCOL3 + 50, Row, ROW_HT, HCOL4 - HCOL3, str, ifontName, ifontSize, "", "L");
                if (Report_format == "FC")
                    AddXYLabel(HCOL3 + 190, Row, ROW_HT, HCOL5 - HCOL4, "SWIFT CODE", ifontName, ifontSize, "", "LB");
                else
                    AddXYLabel(HCOL3 + 200, Row, ROW_HT, HCOL5 - HCOL4, "IFSC CODE", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 260, Row, ROW_HT, (HCOL3 + 250) - (HCOL3 + 240), ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (BankInfoDic.ContainsKey("BANK_IFSC_CODE"))
                    str = BankInfoDic["BANK_IFSC_CODE"].ToString();
                AddXYLabel(HCOL3 + 270, Row, ROW_HT, (HCOL6 + 70) - (HCOL3 + 220), str, ifontName, ifontSize, "", "L");

                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "BANK NAME", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2 + 50, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (BankInfoDic.ContainsKey("BANK_NAME"))
                    str = BankInfoDic["BANK_NAME"].ToString();
                AddXYLabel(HCOL3 + 50, Row, ROW_HT, HCOL6 - HCOL3, str, ifontName, ifontSize, "", "L");
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "ADDRESS", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2 + 50, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                str = "";
                if (BankInfoDic.ContainsKey("BANK_ADD1"))
                    str = BankInfoDic["BANK_ADD1"].ToString();
                AddXYLabel(HCOL3 + 50, Row, ROW_HT, HCOL6 - HCOL3, str.Trim(), ifontName, ifontSize, "", "L");
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "", "LB");
                str = "";
                if (BankInfoDic.ContainsKey("BANK_ADD2"))
                    str = BankInfoDic["BANK_ADD2"].ToString();
                AddXYLabel(HCOL3 + 50, Row, ROW_HT, HCOL6 - HCOL3, str, ifontName, ifontSize, "", "L");
                if (Report_format == "FC")
                {
                    Row += ROW_HT;
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "CORRESPONDENT BANK", ifontName, ifontSize, "", "LB");
                    AddXYLabel(HCOL3 + 40, Row, ROW_HT, 10, ":", ifontName, ifontSize, "", "CB");
                    str = "";
                    if (BankInfoDic.ContainsKey("BANK_ADD3"))
                        str = BankInfoDic["BANK_ADD3"].ToString();
                    AddXYLabel(HCOL3 + 50, Row, ROW_HT, HCOL6 - HCOL3, str, ifontName, ifontSize, "", "L");
                }
                Row += ROW_HT;
            }
            else
            {
                Row += ROW_HT;
                Row += ROW_HT;
                Row += ROW_HT;
                Row += ROW_HT;
                Row += ROW_HT;
                Row += ROW_HT;
                Row += ROW_HT;
            }
            AddXYLabel(HCOL6 + 50, Row, ROW_HT, HCOL7 - HCOL6, "AUTHORISED SIGNATORY", ifontName, ifontSize, "", "LB");
            Row += ROW_HT;
            if (hRow.jvh_type == "IN" && DT_INVFOOTER != null)
            {
                foreach (DataRow dr in DT_INVFOOTER.Rows)
                {
                    Row += ROW_HT;
                    str = dr["text_value"].ToString();
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, str, ifontName, ifontSize, "", "L");
                }
            }
            else
            {
                Row += ROW_HT;
                str = "1.Payment is immediate or as per the agreement,any overdue bills will attract penal interest @24% PA.";
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, str, ifontName, ifontSize, "", "L");
                Row += ROW_HT;
                str = "2.The above rates applied are subject to change as per the Orders issued by the Concerned Authorities.";
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, str, ifontName, ifontSize, "", "L");
                Row += ROW_HT;
                str = "3.Any clarification to this bill should be notified in writing and obtained within three working days from the bill date.";
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, str, ifontName, ifontSize, "", "L");
                if (company_code == "SGT")
                {
                    Row += ROW_HT;
                    str = "4.GST Will Be Paid By Recepient Of Service";
                    AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, str, ifontName, ifontSize, "", "L");
                }
            }

            Row += ROW_HT;
            Row += ROW_HT;

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "GST NO", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2 + 30, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3 + 30, Row, ROW_HT, HCOL6 - HCOL3, hRow.inv_comp_gstin, ifontName, ifontSize, "", "LB");
            //AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, str, ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            //str = "PANNO : " + hRow.inv_comp_panno;
            //AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, str, ifontName, ifontSize, "", "L");
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "PAN NO", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL2 + 30, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
            AddXYLabel(HCOL3 + 30, Row, ROW_HT, HCOL6 - HCOL3, hRow.inv_comp_panno, ifontName, ifontSize, "", "LB");
            if (hRow.inv_comp_uamno != "")
            {
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "UAM NO", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL2 + 30, Row, ROW_HT, HCOL3 - HCOL2, ":", ifontName, ifontSize, "", "LB");
                AddXYLabel(HCOL3 + 30, Row, ROW_HT, HCOL6 - HCOL3, hRow.inv_comp_uamno, ifontName, ifontSize, "", "LB");
            }
            Row += ROW_HT;
            Row += ROW_HT;
            str = hRow.inv_comp_reg_address;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, str, ifontName, ifontSize, "", "C");
        }
        private void WriteExpenseFooter()
        {
            decimal invTotAmt = 0;
            if(DT_MBLINV!=null)
            {
   
                DCOL1 = 10;
                DCOL2 = DCOL1 + 30;//90
                DCOL3 = DCOL2 + 130;//160
                DCOL4 = DCOL3 + 35;
                DCOL5 = DCOL4 + 280;
                DCOL6 = DCOL5 + 85;//140
                DCOL7 = DCOL6 + 50;
                DCOL8 = DCOL7 + 50;
                DCOL9 = DCOL8 + 50;
                DCOL10 = DCOL9 + 70;
                

                Row += ROW_HT;
                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL10 - DCOL1, "INVOICE DETAILS", ifontName, ifontSize, "", "LB");
                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "SI#", ifontName, ifontSize, "LTBR", "LB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "HBL#", ifontName, ifontSize, "LTBR", "LB");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "TERMS", ifontName, ifontSize, "LTBR", "LB");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "PARTY", ifontName, ifontSize, "LTBR", "LB");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "INVOICE#-DATE", ifontName, ifontSize, "LTBR", "LB");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "FRT", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "THC", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "TPT", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, "TOTAL", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                foreach (DataRow dr in DT_MBLINV.Rows)
                {
                    Row += ROW_HT;
                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, dr["SI"].ToString(), ifontName, ifontSize, "L", "L");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, dr["HBL_NO"].ToString(), ifontName, ifontSize, "L", "L");
                    str = "";
                    if (dr["TERMS"].ToString() == "FREIGHT COLLECT")
                        str = "CC";
                    if (dr["TERMS"].ToString() == "FREIGHT PREPAID")
                        str = "PP";
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, str, ifontName, ifontSize, "L", "L");
                    str = "";
                    if (dr["HBL_TYPE"].ToString() == "HBL-SE" || dr["HBL_TYPE"].ToString() == "HBL-AE")
                        str = dr["EXP_NAME"].ToString();
                    if (dr["HBL_TYPE"].ToString() == "HBL-SI" || dr["HBL_TYPE"].ToString() == "HBL-AI")
                        str = dr["IMP_NAME"].ToString();
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, str, ifontName, ifontSize, "L", "L", 0, 0, 0, 0, 0, 0, 0);
                    str = dr["JVH_VRNO"].ToString();
                    if (!dr["JVH_DATE"].Equals(DBNull.Value))
                    {
                        DateTime Dt = (DateTime)dr["JVH_DATE"];
                        str += " - " + Dt.ToString("dd/MM/yy");
                    }
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, str, ifontName, ifontSize, "L", "L", 0, 0, 0, 0, 0, 0, 0);
                    AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, dr["INV_FRT"].ToString(), ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, dr["INV_THC"].ToString(), ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, dr["INV_TPT"].ToString(), ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, Lib.NumericFormat(dr["INV_TOTAL"].ToString(), 2), ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);
                    invTotAmt += Lib.Conv2Decimal(dr["INV_TOTAL"].ToString());
                }
                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "TOTAL", ifontName, ifontSize, "LTB", "LB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "TB", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "TB", "L");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "", ifontName, ifontSize, "TB", "L");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "TB", "L");
                AddXYLabel(DCOL6, Row, ROW_HT, DCOL7 - DCOL6, "", ifontName, ifontSize, "TB", "L");
                AddXYLabel(DCOL7, Row, ROW_HT, DCOL8 - DCOL7, "", ifontName, ifontSize, "TB", "L");
                AddXYLabel(DCOL8, Row, ROW_HT, DCOL9 - DCOL8, "", ifontName, ifontSize, "TB", "L");
                AddXYLabel(DCOL9, Row, ROW_HT, DCOL10 - DCOL9, Lib.NumericFormat(invTotAmt.ToString(), 2), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
            }

            if (DT_MBLOS != null)
            {
                DCOL1 = 10;
                DCOL2 = DCOL1 + 160;
                DCOL3 = DCOL2 + 100;
                DCOL4 = DCOL3 + 100;
                DCOL5 = DCOL4 + 100;
                DCOL6 = DCOL5 + 100;

                Row += ROW_HT;
                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL6 - DCOL1, "OUTSTANDING DETAILS", ifontName, ifontSize, "", "LB");
                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "INVOICE#-DATE", ifontName, ifontSize, "LTBR", "LB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "INVOICE-AMT", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "BALANCE", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "CREDIT-DAYS", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "OVERDUE-DAYS", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
                foreach (DataRow dr in DT_MBLOS.Rows)
                {
                    Row += ROW_HT;
                    str = dr["JVH_VRNO"].ToString();
                    if (!dr["JVH_DATE"].Equals(DBNull.Value))
                    {
                        DateTime Dt = (DateTime)dr["JVH_DATE"];
                        str += " - " + Dt.ToString("dd/MM/yy");
                    }

                    AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, str, ifontName, ifontSize, "L", "L");
                    AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, Lib.NumericFormat(dr["JV_DEBIT"].ToString(), 2), ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, Lib.NumericFormat(dr["BALANCE"].ToString(), 2), ifontName, ifontSize, "L", "R", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, dr["CR_DAYS"].ToString(), ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);
                    AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, dr["OS_DAYS"].ToString(), ifontName, ifontSize, "LR", "R", 0, 0, 0, 0, 0, 0, -5);
                }
                Row += ROW_HT;
                AddXYLabel(DCOL1, Row, ROW_HT, DCOL2 - DCOL1, "", ifontName, ifontSize, "T", "LB");
                AddXYLabel(DCOL2, Row, ROW_HT, DCOL3 - DCOL2, "", ifontName, ifontSize, "T", "L");
                AddXYLabel(DCOL3, Row, ROW_HT, DCOL4 - DCOL3, "", ifontName, ifontSize, "T", "L");
                AddXYLabel(DCOL4, Row, ROW_HT, DCOL5 - DCOL4, "", ifontName, ifontSize, "T", "L");
                AddXYLabel(DCOL5, Row, ROW_HT, DCOL6 - DCOL5, "", ifontName, ifontSize, "T", "L");
            }

        }
    }
}


