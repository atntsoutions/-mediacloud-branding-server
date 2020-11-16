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
    public class BalConfirmReportService : BaseReport
    {
        public string report_folder = "";
        public string folderid = "";
        public string File_Name = "";
        public string File_Type = "";
        public string File_Display_Name = "myreport.pdf";
        public string company_code = "";
        public string branch_code = "";
        public string user_name = "";
        public string PKID = "";
        public bool IsAllChked = false;
        public string AsOnDate = "";
        string sql = "";
        private string ImagePath = "";

        private string comp_name = "";
        private string comp_add1 = "";
        private string comp_add2 = "";
        private string comp_add3 = "";

        private string cust_id = "";
        public string cust_name = "";
        public string cust_add1 = "";
        public string cust_add2 = "";
        public string cust_add3 = "";
        private string cust_branch_name = "";
        private string cust_date1 = "";
        private decimal cust_amt1 = 0;
        private string cust_date2 = "";
        private decimal cust_amt2 = 0;

        DBConnection Con_Oracle = null;

        public void Process(object party_id, object party_name, object branch_name, string date1, object amt1, string date2, object amt2)
        {
            try
            {
                if (party_id == null || party_id.ToString() == "")
                    return;

                cust_id = party_id.ToString();
                cust_name = party_name.ToString();
                cust_branch_name = branch_name.ToString();
                cust_date1 = date1;
                cust_amt1 = Lib.Conv2Decimal(amt1.ToString());
                cust_date2 = date2;
                cust_amt2 = Lib.Conv2Decimal(amt2.ToString());
                AsOnDate = date2;

                ReadData();

                File_Name = report_folder + "\\BALREPORT\\" + (IsAllChked == true ? "ALL" : branch_code) + "\\" + DateTime.Now.ToString("yyyy-MM-dd");
                if (!System.IO.Directory.Exists(File_Name))
                    System.IO.Directory.CreateDirectory(File_Name);

                cust_name = Lib.ProperFileName(cust_name);
                if (cust_name.Length > 100)
                    cust_name = cust_name.Substring(0, 100);

                File_Name += "\\" + cust_name + ".pdf";

                ImagePath = report_folder + "\\Images";

                BeginReport(Page_Height, Page_Width);
                PrintData();
                EndReport();
                if (ExportList != null)
                {
                    Export2Pdf mypdf = new Export2Pdf();
                    mypdf.ExportList = ExportList;
                    mypdf.FileName = File_Name;
                    mypdf.Page_Height = Page_Height;
                    mypdf.Page_Width = Page_Width;
                    mypdf.Process();
                }


            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw ex;
            }
        }


        private void ReadData()
        {
            string GstStateId = "";
            LovService mService = new LovService();
            DataRow mRow_gstst = mService.getSettings(branch_code, "GST-STATE");
            if (mRow_gstst != null)
            {
                GstStateId = mRow_gstst["id"].ToString();
            }

            Con_Oracle = new DBConnection();

            sql = "select add_line1,add_line2,add_line3,add_state_id from addressm where add_parent_id = '" + cust_id + "' order by add_branch_slno"; //and add_branch_slno = 0
            DataTable Dt_temp = new DataTable();
            Dt_temp = Con_Oracle.ExecuteQuery(sql);

            foreach (DataRow Dr in Dt_temp.Rows)
            {
                if (cust_add1 == "")//to initialsd first
                {
                    cust_add1 = Dr["add_line1"].ToString();
                    cust_add2 = Dr["add_line2"].ToString();
                    cust_add3 = Dr["add_line3"].ToString();
                }
                if (GstStateId == Dr["add_state_id"].ToString())
                {
                    cust_add1 = Dr["add_line1"].ToString();
                    cust_add2 = Dr["add_line2"].ToString();
                    cust_add3 = Dr["add_line3"].ToString();
                    break;
                }
            }

            Con_Oracle.CloseConnection();

            if (branch_code == "KOLAF")
            {
                comp_name = "CARGOMAR (KOLKATA) PVT. LTD.";
                comp_add1 = "Diamond Chambers 4 Chowringhee Lane, Room: 10A, 10th Floor, Block - III & IV, Kolkata-700 016";
                comp_add2 = "Tel: +91-33-2252 0949, 0950  Fax: +91-33-2252 0952";
                comp_add3 = "Email: infocal@cargomar.in  Website: www.cargomar.in CIN: U63090WB2001PTC093556";

            }
            else
            {

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                mSearchData.Add("table", "COMP_ADDRESS");
                mSearchData.Add("comp_code", company_code);
                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        comp_name = Dr["COMP_NAME"].ToString();
                        comp_add1 = Dr["COMP_ADDRESS1"].ToString() + ", " + Dr["COMP_ADDRESS2"].ToString();
                        comp_add2 = Dr["COMP_ADDRESS3"].ToString();
                        comp_add3 = "Email: " + Dr["COMP_EMAIL"].ToString() + " Website: " + Dr["COMP_WEB"].ToString() + " CIN: " + Dr["COMP_CINNO"].ToString();
                        break;
                    }
                }
                //comp_name = "CARGOMAR PVT. LTD.";
                //comp_add1 = "Regd. Off.No. III/695-C,'Cargomar House', Kottaram Junction, Maradu, Kochi-682 304, India";
                //comp_add2 = "Tel: +91-484-2705995,2706224,2706730  Fax: +91-484-2706224";
                //comp_add3 = "Email: ganesh@cargomar.in Website : www.cargomar.in";
            }
        }


        private void PrintData()
        {
            InitPage();
            AddPage(1100, 800);
            LoadImage(ImagePath + "\\FIATA.jpg", 100, 10, 60, 55);
            LoadImage(ImagePath + "\\Logo.gif", 350, 10, 68, 70);
            LoadImage(ImagePath + "\\IATA.jpg", 600, 10, 65, 63);
            WriteCompanyDetails();
            WriteDetails();
        }
        private void InitPage()
        {
            HCOL1 = 70;
            HCOL2 = HCOL1 + 150;
            HCOL3 = HCOL2 + 150;
            HCOL4 = HCOL3 + 150;
            HCOL5 = HCOL4 + 100;
            HCOL6 = HCOL5 + 100;
            HCOL7 = HCOL6 + 100;

            //ifontName = "Verdana";
            //ifontSize = 10;

            ifontName = "Calibri";
            ifontSize = 14;

            Row = 90;
            ROW_HT = 16;
        }
        private void WriteCompanyDetails()
        {

            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_name, "Times New Roman", 22, "", "BC");
            DrawHLine(2, Row + ROW_HT + 2, Page_Width - 4);
            Row += ROW_HT + 2;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_add1, ifontName, ifontSize, "", "BC");
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_add2, ifontName, ifontSize, "", "BC");
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, comp_add3, ifontName, ifontSize, "", "BC");
            Row += ROW_HT;
            Row += ROW_HT;


        }

        private void WriteDetails()
        {
            string str = "";
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, DateTime.Now.ToString(Lib.FRONT_END_DATE_DISPLAY_FORMAT), ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "To,", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Accounts Manager", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, cust_name, ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, cust_add1, ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, cust_add2, ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, cust_add3, ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Dear Sir,", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Sub: Balance confirmation as on " + AsOnDate, ifontName, ifontSize, "", "LB");
            Row += ROW_HT;
            Row += ROW_HT;

            if (cust_amt2 < 0)
                str = "credit balance";
            else
                str = "debit balance";

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "With reference to the above subject, our books of account show a "+str+ " in your account as below", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "DATE", ifontName, ifontSize, "LTBR", "LB");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "AMOUNT", ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, -5);
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "DR/CR", ifontName, ifontSize, "LTBR", "LB");
            if (cust_amt1 != 0)
            {
                Row += ROW_HT;
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, cust_date1, ifontName, ifontSize, "LTBR", "L");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Lib.NumericFormat(cust_amt1.ToString(), 2), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, -5);
                if (cust_amt1 < 0)
                    str = "CR";
                else
                    str = "DR";
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, str, ifontName, ifontSize, "LTBR", "L");
            }

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, cust_date2, ifontName, ifontSize, "LTBR", "L");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, Lib.NumericFormat(cust_amt2.ToString(), 2), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, -5);
            if (cust_amt2 < 0)
                str = "CR";
            else
                str = "DR";
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, str, ifontName, ifontSize, "LTBR", "L");


            Row += ROW_HT;

            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Please sign the confirmation slip below and forward this to our address mentioned in the letter head.", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "In case you need any further assistance including statement of account to reconcile the differences", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            if (branch_code == "KOLAF")
                str = "infocal@cargomar.in";
            else
                str = "hogen@cargomar.in";
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "if any please do write to us on our email id "+str, ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "We would be grateful if you would give this request your earliest attention. If we do not receive your", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "confirmation within 15 days, we would treat the above balance is correct.", ifontName, ifontSize, "", "L");
  
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Yours faithfully,", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "FOR " + comp_name, ifontName, ifontSize, "", "LB");
            Row += ROW_HT;
            Row += ROW_HT;
            Row += ROW_HT;
            Row += ROW_HT;

            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Authorised Signatory", ifontName, ifontSize, "", "LB");
            Row += ROW_HT;
            DrawHLine(2, Row + ROW_HT , Page_Width - 4);
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 50, "(Do not perforate the form at this point)", ifontName, ifontSize, "", "RB");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(2, Row, ROW_HT, 800 - 4, "Confirmation", ifontName, ifontSize, "", "CUB");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "We confirm the balance as mention above as per our books of accounts as on "+ AsOnDate, ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "Date: ", ifontName, ifontSize, "", "L");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, "Signature with Rubber Stamp:" , ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL7 - HCOL1, "Place: ", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            Row += ROW_HT;
            Row += ROW_HT;
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, "Name of Authorised Person:", ifontName, ifontSize, "", "L");
            Row += ROW_HT;
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL7 - HCOL3, "Designation:", ifontName, ifontSize, "", "L");

        }
    }
}
