using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using System.Drawing;
using XL.XSheet;

namespace BLHr
{
    public class SalarySheetService : BaseReport
    {
        ExcelFile file;
        ExcelWorksheet ws = null;
        ExcelWorksheet ws2 = null;
        CellRange myCell;
        int iCol = 0;
        int iRow = 0;

        public string report_folder = "";
        public string folderid = "";
        public string File_Name = "";
        public string File_Type = "";
        public string File_Display_Name = "myreport.pdf";
        public string company_code = "";
        public string branch_code = "";
        public string year_code = "";
        private string ImagePath = "";
        public string PKID = "";
        public string Invoke_type = "";
        public string Emp_Status = "";
        public string IsAdmin = "N";
        public bool IsConsol = false;
        public int cYear = 0;
        public int cMonth = 0;
        public string fileformat = "";
        public int emp_br_grp = 1;
        public string branch_codes = "";

        //public List<AttachList> mailattachments = new List<AttachList>();
        private string Report_Caption = "";
        private string Report_Header1 = "";
        private string Report_Header2 = "";
        private string Report_Header3 = "";
        private string Report_Header4 = "";
        private int ifontSizesm = 7;
       // private float ROW_HTsm = 0;

        private int RowCount = 1;
        private int RowsPerPage = 43;
        string sql = "";
        DataTable Dt_invoice = new DataTable();
        DataTable Dt_Data = new DataTable();
        DataTable Dt_SalSheet = new DataTable();
        private float Xtolrnce =-5;
       // private DataRow DR = null;

        //private Dictionary<string, string> DR_DOC = null;
        //string str = "";
        DBConnection Con_Oracle = null;
        string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "",
                comp_tel = "", comp_fax = "", comp_web = "", comp_email = "", comp_cinno = "", comp_gstin = "", Comp_br_name = "";

        
        private float HCOL11 = 0;
        private float HCOL12 = 0;
        private float HCOL13 = 0;
        private float HCOL14 = 0;
        private float HCOL15 = 0;
        private float HCOL16 = 0;
        private float HCOL17 = 0;
        private float HCOL18 = 0;
        private float HCOL19 = 0;
        private float HCOL20 = 0;
        private float HCOL21 = 0;
        private float HCOL22 = 0;
        private float HCOL23 = 0;
        private float HCOL24 = 0;
        //private float HCOL25 = 0;

        
        public SalarySheetService()
        {

        }
        public void ProcessData()
        {
            try
            {
                Init();
                ReadData();
                if (Dt_SalSheet.Rows.Count <= 0)
                    throw new Exception("No Details or Invalid Period to Print...Salarysheet");

                if (!AllValid())
                    return;

                string fname = "myreport";
                if (Dt_SalSheet.Rows[0]["SAL_MON_YR"].ToString().Trim().Length > 0)
                {
                    if (emp_br_grp <= 1)
                        fname = "SALSHEET-" + branch_code + "-" + Dt_SalSheet.Rows[0]["SAL_MON_YR"].ToString().Replace(" ", "");
                    else
                        fname = "SALSHEET-TPRSF-" + Dt_SalSheet.Rows[0]["SAL_MON_YR"].ToString().Replace(" ", "");
                }
                if (fname.Length > 30)
                    fname = fname.Substring(0, 30);

                if (fileformat == "PDF")
                {
                    File_Display_Name = Lib.ProperFileName(fname) + ".pdf";
                    File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                    File_Type = "pdf";
                    ImagePath = report_folder + "\\Images";


                    BeginReport(800, 1375);
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
                else
                {
                    File_Display_Name = Lib.ProperFileName(fname) + ".xls";
                    File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                    File_Type = "EXCEL";
                    ImagePath = report_folder + "\\Images";
                    PrintExcelSalSheet();
                }
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw ex;
            }
        }
        private void Init()
        {
        }
        private bool AllValid()
        {
            bool bRet = true;
            return bRet;
        }
        private void ReadData()
        {
            if (emp_br_grp <= 0)
                emp_br_grp = 1;

            PKID = PKID.Replace(",", "','");

            sql = " Select SAL_PKID,SAL_MONTH,SAL_YEAR,SAL_EMP_ID,SAL_DATE,upper(trim(to_char(sal_date, 'MONTH')))||'-'||to_char(sal_date, 'YYYY') as SAL_MON_YR ";
            sql += "  ,A01,A02,A03,A04,A05";
            sql += "  ,A06,A07,A08,A09,A10";
            sql += "  ,A11,A12,A13,A14,A15";
            sql += "  ,A16,A17,A18,A19,A20";
            sql += "  ,A21,A22,A23,A24,A25";
            sql += "  ,D01-nvl(SAL_PF_BAL,0) as D01 ,D02,D03,D04,D05";
            sql += "  ,D06,D07,D08,D09,D10";
            sql += "  ,D11,D12,D13,D14,D15";
            sql += "  ,D16,D17,D18,D19,D20";
            sql += "  ,D21,D22,D23,D24,D25";
            sql += "  ,SAL_NET,SAL_GROSS_EARN,SAL_GROSS_DEDUCT";
            sql += "  ,SAL_LOP_AMT,SAL_DAYS_WORKED,SAL_PF_BAL,SAL_PF_MON_YEAR";
            sql += "  ,EMP_PKID,EMP_NAME,EMP_NO,grd.param_name as EMP_GRADE,desig.param_name as EMP_DESIGNATION";
            sql += "  ,EMP_DO_JOINING,EMP_DO_RELIEVE,SAL_BASIC_RT,SAL_DA_RT";
            sql += "  ,EMP_PAN,EMP_BANK_ACNO,EMP_PFNO,EMP_ESINO, sal_mail_sent,EMP_EMAIL_OFFICE,EMP_EMAIL_PERSONAL,SAL_PF_WAGE_BAL,SAL_PF_LIMIT,SAL_PF_BASE ";
            sql += "  ,EMP_FATHER_NAME,SAL_PAY_DATE,a.REC_BRANCH_CODE,a.REC_CATEGORY,null as branch ";
            sql += "  from salarym a";
            sql += "  inner join empm b on a.sal_emp_id = b.emp_pkid";
            sql += "  left join param grd on b.emp_grade_id = grd.param_pkid";
            sql += "  left join param desig on b.emp_designation_id = desig.param_pkid";
            sql += "  where a.rec_company_code = '" + company_code + "'";
            if (!IsConsol)
                sql += "  and a.sal_emp_branch_group = " + emp_br_grp;
            if (PKID != "")
                sql += "  and a.sal_pkid in ('" + PKID + "')";
            else
            {
                if (!IsConsol)
                    sql += "  and a.rec_branch_code = '" + branch_code + "'";
                else if(branch_codes !="")
                    sql += "  and a.rec_branch_code in ('" + branch_codes + "')";

                sql += "  and a.sal_month = " + cMonth.ToString();
                sql += "  and a.sal_year = " + cYear.ToString();
                if (Emp_Status != "BOTH")
                    sql += " and a.rec_category = '" + Emp_Status.ToString() + "'";
            }
            sql += " order by emp_no";

            Con_Oracle = new DBConnection();
            Dt_SalSheet = new DataTable();
            Dt_SalSheet = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
        }
         
        private void PrintData()
        {
            InitPage();
            ReadCompanyDetails();
            GetNewPage();
            WriteDetails();
        }
        private void InitPage()
        {
            HCOL1 = 30;
            HCOL2 = HCOL1 + 25;
            HCOL3 = HCOL2 + 40;
            HCOL4 = HCOL3 + 215;
            HCOL5 = HCOL4 + 55;//HCOL5 = HCOL4 + 65;//new col
            HCOL6 = HCOL5 + 55;
            HCOL7 = HCOL6 + 35;
            HCOL8 = HCOL7 + 35;
            HCOL9 = HCOL8 + 55;
            HCOL10 = HCOL9 + 45;
            HCOL11 = HCOL10 + 55;
            HCOL12 = HCOL11 + 60;
            HCOL13 = HCOL12 + 65;
            HCOL14 = HCOL13 + 45;
            HCOL15 = HCOL14 + 65;
            HCOL16 = HCOL15 + 45;
            HCOL17 = HCOL16 + 45;
            HCOL18 = HCOL17 + 45;
            HCOL19 = HCOL18 + 45;
            HCOL20 = HCOL19 + 45;
            HCOL21 = HCOL20 + 35;
            HCOL22 = HCOL21 + 55;
            HCOL23 = HCOL22 + 55;
            HCOL24 = HCOL23 + 60;

            ifontName = "Calibri";
            ifontSize = 10;
            
            Row = 0;
            ROW_HT = 15;
        }
        private void ReadCompanyDetails()
        {
            comp_name = ""; comp_add1 = ""; comp_add2 = ""; comp_add3 = "";
            comp_tel = ""; comp_fax = ""; comp_web = ""; comp_email = ""; comp_cinno = ""; comp_gstin = "";
            Comp_br_name = "";
            if (emp_br_grp == 1)
            {
                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                if (branch_code == "")
                    mSearchData.Add("branch_code", "HOCPL");
                else
                    mSearchData.Add("branch_code", branch_code);

                DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
                if (Dt_CompAddress != null)
                {
                    foreach (DataRow Dr in Dt_CompAddress.Rows)
                    {
                        Comp_br_name = Dr["BR_NAME"].ToString();
                        comp_name = Dr["COMP_NAME"].ToString();
                        comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                        comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                        comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                        comp_tel = Dr["COMP_TEL"].ToString();
                        comp_fax = Dr["COMP_FAX"].ToString();
                        comp_web = Dr["COMP_WEB"].ToString();
                        comp_email = Dr["COMP_EMAIL"].ToString();
                        comp_cinno = Dr["COMP_CINNO"].ToString();
                        comp_gstin = Dr["COMP_GSTIN"].ToString();
                        break;
                    }
                }
            }
            if (emp_br_grp == 2)
            {
                Lib.GetEmployeeBranch(2, ref comp_name, ref comp_add1, ref comp_add2, ref comp_add3);
            }

            Report_Header1 = comp_name;
            Report_Header2 = comp_add1;
            Report_Header3 = comp_add2;
            if (Report_Header3.Trim() != "" && comp_add3.Trim() != "")
                Report_Header3 += ",";
            Report_Header3 += comp_add3;

            DateTime dtmstart = new DateTime(cYear, cMonth, 1);
            Report_Caption = "SALARY/ STIPEND";
            if (Emp_Status == "CONFIRMED")
                Report_Caption = "SALARY";
            else if (Emp_Status == "UNCONFIRM")
                Report_Caption = "STIPEND";
            Report_Caption += " DETAILS FOR " + dtmstart.ToString("MMMM").ToUpper() + ", " + cYear.ToString();

        }

        private void WriteHeader()
        {
            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "", ifontName, ifontSize, "LT", "");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL24 - HCOL3, Report_Header1, ifontName, 9, "LTR", "B");
            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "", ifontName, ifontSize, "L", "");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL24 - HCOL3, Report_Header2, ifontName, 9, "LR", "B");
            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL3 - HCOL1, "", ifontName, ifontSize, "LB", "");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL24 - HCOL3, Report_Header3, ifontName, 9, "LBR", "B");

            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL24 - HCOL1, "", ifontName, ifontSize, "LTR", "");
            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL24 - HCOL1, Report_Caption, ifontName, 9, "LR", "CB");
            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL24 - HCOL1, "", ifontName, ifontSize, "LBR", "");


            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSizesm, "LTR", "");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSizesm, "TR", "");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "", ifontName, ifontSize, "TR", "");
           // AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "", ifontName, ifontSize, "TR", "");
            AddXYLabel(HCOL4, Row, ROW_HT, HCOL13 - HCOL4, "EARNINGS", ifontName, ifontSize, "TBR", "CB");
            AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, "", ifontName, ifontSize, "TR", "");
            AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, "", ifontName, ifontSize, "TR", "");
            AddXYLabel(HCOL15, Row, ROW_HT, HCOL22 - HCOL15, "DEDUCTIONS", ifontName, ifontSize, "TBR", "CB");
            AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, "", ifontName, ifontSize, "TR", "");
            AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, "", ifontName, ifontSize, "TR", "");

            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "LR", "C");
            AddXYLabel(HCOL1, Row + 8, ROW_HT, HCOL2 - HCOL1, "SI#", ifontName, ifontSize, "", "C");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL2, Row + 8, ROW_HT, HCOL3 - HCOL2, "CODE", ifontName, ifontSize, "", "C");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL3, Row + 8, ROW_HT, HCOL4 - HCOL3, "EMPLOYEE NAME", ifontName, ifontSize, "", "C");
            //AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "", ifontName, ifontSize, "R", "C");
            //AddXYLabel(HCOL4, Row + 8, ROW_HT, HCOL5 - HCOL4, "BRANCH", ifontName, ifontSize, "", "C");

            AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "BASIC/", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL12, Row, ROW_HT, HCOL13 - HCOL12, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL13, Row + 8, ROW_HT, HCOL14 - HCOL13, "LOP", ifontName, ifontSize, "", "C");
            AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, "GROSS", ifontName, ifontSize, "R", "CB");
            AddXYLabel(HCOL15, Row, ROW_HT, HCOL16 - HCOL15, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL16, Row, ROW_HT, HCOL17 - HCOL16, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL17, Row, ROW_HT, HCOL18 - HCOL17, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL18, Row, ROW_HT, HCOL19 - HCOL18, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL19, Row, ROW_HT, HCOL20 - HCOL19, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL20, Row, ROW_HT, HCOL21 - HCOL20, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL21, Row, ROW_HT, HCOL22 - HCOL21, "OTHER", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, "", ifontName, ifontSize, "R", "CB");
            AddXYLabel(HCOL22, Row - 8, ROW_HT, HCOL23 - HCOL22, "TOTAL", ifontName, ifontSize, "", "CB");
            AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, "NET", ifontName, ifontSize, "R", "CB");

            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "LR", "");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "R", "");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "", ifontName, ifontSize, "R", "");
           // AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "", ifontName, ifontSize, "R", "");

            AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "DA / SPL", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "BASIC", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "DA", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, "SPL", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, "TOTAL", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, "CCA", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, "HRA", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, "", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL12, Row, ROW_HT, HCOL13 - HCOL12, "", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL11, Row - 8, ROW_HT, HCOL12 - HCOL11, "CONVE-", ifontName, ifontSize, "", "C");
            AddXYLabel(HCOL12, Row - 8, ROW_HT, HCOL13 - HCOL12, "OTHER", ifontName, ifontSize, "", "C");
            AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, "", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, Emp_Status == "UNCONFIRM" ? "STIPEND" : "SALARY", ifontName, ifontSize, "R", "CB");
            AddXYLabel(HCOL15, Row, ROW_HT, HCOL16 - HCOL15, "PF", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL16, Row, ROW_HT, HCOL17 - HCOL16, "ESI", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL17, Row, ROW_HT, HCOL18 - HCOL17, "TDS", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL18, Row, ROW_HT, HCOL19 - HCOL18, "PTAX", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL19, Row, ROW_HT, HCOL20 - HCOL19, "LOAN", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL20, Row, ROW_HT, HCOL21 - HCOL20, "LWF", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL21, Row, ROW_HT, HCOL22 - HCOL21, "DEDUC-", ifontName, ifontSize, "R", "C");
            AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, "", ifontName, ifontSize, "R", "CB");
            AddXYLabel(HCOL22, Row - 8, ROW_HT, HCOL23 - HCOL22, "DEDUC-", ifontName, ifontSize, "", "CB");
            AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, Emp_Status == "UNCONFIRM" ? "STIPEND" : "SALARY", ifontName, ifontSize, "R", "CB");

            GetNextRow();
            AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "LBR", "");
            AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, "", ifontName, ifontSize, "BR", "");
           // AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "", ifontName, ifontSize, "BR", "");

            AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "RATE", ifontName, ifontSize, "BR", "C");
            AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, "", ifontName, ifontSize, "BR", "C");
            AddXYLabel(HCOL12, Row, ROW_HT, HCOL13 - HCOL12, "", ifontName, ifontSize, "BR", "C");
            AddXYLabel(HCOL11, Row - 8, ROW_HT, HCOL12 - HCOL11, "YANCE", ifontName, ifontSize, "", "C");
            AddXYLabel(HCOL12, Row - 8, ROW_HT, HCOL13 - HCOL12, "EARNINGS", ifontName, ifontSize, "", "C");
            AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, "", ifontName, ifontSize, "BR", "CB");
            AddXYLabel(HCOL15, Row, ROW_HT, HCOL16 - HCOL15, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL16, Row, ROW_HT, HCOL17 - HCOL16, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL17, Row, ROW_HT, HCOL18 - HCOL17, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL18, Row, ROW_HT, HCOL19 - HCOL18, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL19, Row, ROW_HT, HCOL20 - HCOL19, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL20, Row, ROW_HT, HCOL21 - HCOL20, "", ifontName, ifontSize, "BR", "");
            AddXYLabel(HCOL21, Row, ROW_HT, HCOL22 - HCOL21, "TIONS", ifontName, ifontSize, "BR", "C");
            AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, "", ifontName, ifontSize, "BR", "CB");
            AddXYLabel(HCOL22, Row - 8, ROW_HT, HCOL23 - HCOL22, "TIONS", ifontName, ifontSize, "", "CB");
            AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, "", ifontName, ifontSize, "BR", "CB");

        }
        private void WriteDetails()
        {
            int SiNo = 0;
            decimal TotAmt = 0;
            decimal OtherEarns = 0, OtherDeduts = 0, TotOtherEarns = 0, TotOtherDeduts = 0;
            decimal BasicRt = 0, DaRt = 0, TotBasicRt = 0, TotDaRt = 0;
            foreach (DataRow dr in Dt_SalSheet.Rows)
            {
                GetNextRow();
                OtherEarns = 0;
                OtherDeduts = 0;
                SiNo++;

                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, SiNo.ToString(), ifontName, ifontSize, "LTBR", "C");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, dr["EMP_NO"].ToString(), ifontName, ifontSize, "LTBR", "C");
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, dr["EMP_NAME"].ToString(), ifontName, ifontSizesm, "LTBR", "");
               // AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, dr["BRANCH"].ToString(), ifontName, ifontSizesm - 1, "LTBR", "");
                /*********************BasicRate and Date Rate *******************/
                BasicRt = Lib.Convert2Decimal(dr["SAL_BASIC_RT"].ToString());
                DaRt = Lib.Convert2Decimal(dr["SAL_DA_RT"].ToString());
               
                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, (BasicRt + DaRt).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                TotBasicRt += BasicRt;
                TotDaRt += DaRt;
                
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, (Lib.Convert2Decimal(dr["A01"].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString())).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, dr["A02"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, Lib.Convert2Decimal(dr["A11"].ToString()).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);

                TotAmt = Lib.Convert2Decimal(dr["A01"].ToString()) + Lib.Convert2Decimal(dr["A02"].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString());
                TotAmt += Lib.Convert2Decimal(dr["A11"].ToString());


                AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, Lib.Convert2Decimal(Lib.NumericFormat(TotAmt.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, dr["A03"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, dr["A04"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, dr["A06"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);

                OtherEarns = Lib.Convert2Decimal(dr["A05"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A07"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A08"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A09"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A10"].ToString());

                //OtherEarns += Common.Convert2Decimal(dr["A11"].ToString()); added in col 5

                OtherEarns += Lib.Convert2Decimal(dr["A12"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A13"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A14"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A15"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A16"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A17"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A18"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A19"].ToString());
                //A20 added with basic
                OtherEarns += Lib.Convert2Decimal(dr["A21"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A22"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A23"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A24"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A25"].ToString());
                TotOtherEarns += OtherEarns;
                 
                AddXYLabel(HCOL12, Row, ROW_HT, HCOL13 - HCOL12, Lib.Convert2Decimal(Lib.NumericFormat(OtherEarns.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);

                decimal sLop = 0;
                int sMonth = 0, sYear = 0;
                string eStatus = "";
                sLop = Lib.Convert2Decimal(dr["SAL_LOP_AMT"].ToString());//SAL_LOP_AMT  
                sMonth = Lib.Conv2Integer(dr["SAL_MONTH"].ToString());
                sYear = Lib.Conv2Integer(dr["SAL_YEAR"].ToString());
              
                DateTime dtime;
                if (dr["EMP_DO_JOINING"].ToString().Trim() != "" && Lib.Convert2Decimal(dr["SAL_DAYS_WORKED"].ToString()) + 1 < DateTime.DaysInMonth(sYear, sMonth))
                {
                    dtime = (DateTime)dr["EMP_DO_JOINING"];
                    if (dtime.Month == sMonth && dtime.Year == sYear)
                        eStatus = "DOJ-" + dr["SAL_DAYS_WORKED"].ToString();
                }

                if (dr["EMP_DO_RELIEVE"].ToString().Trim() != "" && Lib.Convert2Decimal(dr["SAL_DAYS_WORKED"].ToString()) + 1 < DateTime.DaysInMonth(sYear, sMonth))
                {
                    dtime = (DateTime)dr["EMP_DO_RELIEVE"];
                    if (dtime.Month == sMonth && dtime.Year == sYear)
                        eStatus = "DOR-" + dr["SAL_DAYS_WORKED"].ToString();
                }

                if (eStatus != "")
                    AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, eStatus, ifontName, ifontSize, "LTBR", "C");
                else
                    AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, sLop.ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, dr["SAL_GROSS_EARN"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);

                AddXYLabel(HCOL15, Row, ROW_HT, HCOL16 - HCOL15, dr["D01"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL16, Row, ROW_HT, HCOL17 - HCOL16, dr["D02"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL17, Row, ROW_HT, HCOL18 - HCOL17, dr["D03"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL18, Row, ROW_HT, HCOL19 - HCOL18, dr["D09"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL19, Row, ROW_HT, HCOL20 - HCOL19, dr["D06"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL20, Row, ROW_HT, HCOL21 - HCOL20, dr["D13"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);

                OtherDeduts = Lib.Convert2Decimal(dr["D04"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D05"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D07"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D08"].ToString());
                //OtherDeduts += Common.Convert2Decimal(dr["D09"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D10"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D11"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D12"].ToString());
                //OtherDeduts += Common.Convert2Decimal(dr["D13"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D14"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D15"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D16"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D17"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D18"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D19"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D20"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D21"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D22"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D23"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D24"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D25"].ToString());

                TotOtherDeduts += OtherDeduts;

                AddXYLabel(HCOL21, Row, ROW_HT, HCOL22 - HCOL21, Lib.Convert2Decimal(Lib.NumericFormat(OtherDeduts.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, dr["SAL_GROSS_DEDUCT"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, dr["SAL_NET"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
            }

            if (Dt_SalSheet.Rows.Count > 0)
            {

                GetNextRow();
                AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "LTBR", "");
                AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "LTBR", "");
                AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, Emp_Status == "UNCONFIRM" ? "STIPEND TOTAL" : "SALARY TOTAL", ifontName, ifontSize, "LTBR", "B");
               // AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, "", ifontName, ifontSize, "LTBR", "B");

                AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, Lib.Convert2Decimal(Lib.NumericFormat((TotBasicRt + TotDaRt).ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, (FindTot("A01") + FindTot("A20")).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, FindTot("A02").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, FindTot("A11").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                TotAmt = FindTot("A01") + FindTot("A02") + FindTot("A20") + FindTot("A11");
                AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, Lib.Convert2Decimal(Lib.NumericFormat(TotAmt.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, FindTot("A03").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, FindTot("A04").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, FindTot("A06").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL12, Row, ROW_HT, HCOL13 - HCOL12, Lib.Convert2Decimal(Lib.NumericFormat(TotOtherEarns.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, FindTot("SAL_LOP_AMT").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);//SAL_LOP_AMT
                AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, FindTot("SAL_GROSS_EARN").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL15, Row, ROW_HT, HCOL16 - HCOL15, FindTot("D01").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL16, Row, ROW_HT, HCOL17 - HCOL16, FindTot("D02").ToString(), ifontName, ifontSize, "LTBR", "RBR", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL17, Row, ROW_HT, HCOL18 - HCOL17, FindTot("D03").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL18, Row, ROW_HT, HCOL19 - HCOL18, FindTot("D09").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL19, Row, ROW_HT, HCOL20 - HCOL19, FindTot("D06").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL20, Row, ROW_HT, HCOL21 - HCOL20, FindTot("D13").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL21, Row, ROW_HT, HCOL22 - HCOL21, Lib.Convert2Decimal(Lib.NumericFormat(TotOtherDeduts.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, FindTot("SAL_GROSS_DEDUCT").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, FindTot("SAL_NET").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
            }
        }
        private decimal FindTot(string FldName)
        {
            return Lib.Convert2Decimal(Lib.NumericFormat(Dt_SalSheet.Compute("sum(" + FldName + ")", "1=1").ToString(), 0));
        }
        private void GetNewPage()
        {
            Row = 20; //Start Row
            RowCount = 1; //Start RowIndex
            ROW_HT = 17;
             ifontSize = 10;
           // ifontSize = 8;
            //ifontSizesm = 7;
            ifontSizesm = 9;
            ifontName = "Calibri";
            // ifontName = "Arial";
            //ifontName = "Verdana";
            AddPage(800, 1375);
            LoadImage(ImagePath + "\\Logo.gif", HCOL1 + 7, Row + 18, 48, 50);
            WriteHeader();
        }
        private void GetNextRow()
        {
            if (RowCount <= RowsPerPage)
            {
                RowCount++;
                Row += ROW_HT;
            }
            else
            {
                GetNewPage();
                RowCount++;
                Row += ROW_HT;
            }
        }

        private void PrintExcelSalSheet()
        {
            //string fname = "myreport";
            //fname = "WAGE-" + branch_code + "-" + new DateTime(salyear, salmonth, 01).ToString("MMMM").ToUpper() + "-" + salyear.ToString();
            //if (fname.Length > 30)
            //    fname = fname.Substring(0, 30);
            //File_Display_Name = Lib.ProperFileName(fname) + ".xls";
            //File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
            //File_Type = "xls";
            // ImagePath = report_folder + "\\Images";

            OpenFile();
            SetColumns();
            WriteHeading();
            FillData();
            file.SaveXls(File_Name);
        }

        private void OpenFile()
        {
            file = new ExcelFile();
            file.Worksheets.Add("Report");
            ws = file.Worksheets["Report"];
            ws.PrintOptions.Portrait = false;
            ws.PrintOptions.FitWorksheetWidthToPages = 1;
        }

        private void SetColumns()
        {
            iRow = 0;
            iCol = 2;
            ws.Columns[0].Width = 255 * 4;
            ws.Columns[1].Width = 255 * 5;
            ws.Columns[2].Width = 255 * 22;
            ws.Columns[3].Width = 255 * 7;
            ws.Columns[4].Width = 255 * 6;
            ws.Columns[5].Width = 255 * 6;
            ws.Columns[6].Width = 255 * 6;
            ws.Columns[7].Width = 255 * 7;
            ws.Columns[8].Width = 255 * 6;
            ws.Columns[9].Width = 255 * 7;
            ws.Columns[10].Width = 255 * 7;
            ws.Columns[11].Width = 255 * 7;
            ws.Columns[12].Width = 255 * 6;
            ws.Columns[13].Width = 255 * 8; 
            ws.Columns[14].Width = 255 * 5;
            ws.Columns[15].Width = 255 * 5;
            ws.Columns[16].Width = 255 * 5;
            ws.Columns[17].Width = 255 * 5;
            ws.Columns[18].Width = 255 * 6;
            ws.Columns[19].Width = 255 * 4;
            ws.Columns[20].Width = 255 * 6;
            ws.Columns[21].Width = 255 * 8;
            ws.Columns[22].Width = 255 * 8;
            for (int s = 0; s < 23; s++)
            {
                ws.Columns[s].Style.Font.Name = "Arial";
                ws.Columns[s].Style.Font.Size = 8 * 20;
            }
        }
        private void WriteHeading()
        {
            string str = "";
            ReadCompanyDetails();


            ws.Cells.GetSubrangeRelative(0, 0, 2, 3).SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
            ws.Cells.GetSubrangeRelative(0, 2, 21, 3).SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
            ws.Cells.GetSubrangeRelative(3, 0, 23, 3).SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);

            WriteData(iRow, 2, comp_name, true);
            iRow++;
            WriteData(iRow, 2, comp_add1, true);
            iRow++;
            WriteData(iRow, 2, comp_add2, true);
            iRow++;
            iRow++;
            
            myCell = ws.Cells.GetSubrangeRelative(iRow, 0, 23, 1);
            myCell.Merged = true;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            myCell.Style.Font.Weight = ExcelFont.BoldWeight;
            myCell.Value = Report_Caption;
            myCell.Style.Font.Name = "Arial";
            myCell.Style.Font.Size = 10 * 20;
            ws.Rows[iRow].Height = 16 * 20;
            iRow++;
            iRow++;

            //str = "Wages period from   ";
            //if (salmonth > 0 && salyear > 0)
            //{
            //    str += "01/" + salmonth.ToString().PadLeft(2, '0') + "/" + salyear.ToString() + "  to  " + DateTime.DaysInMonth(salyear, salmonth).ToString() + "/" + salmonth.ToString().PadLeft(2, '0') + "/" + salyear.ToString();
            //}
            //ws.Rows[iRow].Style.Font.Size = 10 * 20;
            //WriteData(iRow, 0, str);
            //WriteData(iRow, 6, "See Rule 29 (1)", true);
            //WriteData(iRow, 10, "Place");
            //WriteData(iRow, 14, comp_location);
            // ws.Rows[iRow].Style.Font.Size = 8 * 20;

            Merge_Cell(iRow, 0, "SL#", false, 1, 4);
            Merge_Cell(iRow, 1, "CODE", false, 1, 4);
            Merge_Cell(iRow, 2, "EMPLOYEE NAME", false, 1, 4);

            Merge_Cell(iRow, 3, "EARNINGS", true, 9, 1);
            Merge_Cell(iRow + 1, 3, "BASIC / DA / SPL RATE", false, 1, 3);
            Merge_Cell(iRow + 1, 4, "BASIC", false, 1, 3);
            Merge_Cell(iRow + 1, 5, "DA", false, 1, 3);
            Merge_Cell(iRow + 1, 6, "SPL", false, 1, 3);
            Merge_Cell(iRow + 1, 7, "TOTAL", false, 1, 3);
            Merge_Cell(iRow + 1, 8, "CCA", false, 1, 3);
            Merge_Cell(iRow + 1, 9, "HRA", false, 1, 3);
            Merge_Cell(iRow + 1, 10, "CONVE-YANCE", false, 1, 3);
            Merge_Cell(iRow + 1, 11, "OTHER EARNINGS", false, 1, 3);
            Merge_Cell(iRow, 12, "LOP", false, 1, 4);
            Merge_Cell(iRow, 13, "GROSS SALARY", true, 1, 4);
            Merge_Cell(iRow, 14, "DEDUCTIONS", true, 7, 1);
            Merge_Cell(iRow + 1, 14, "PF", false, 1, 3);
            Merge_Cell(iRow + 1, 15, "ESI", false, 1, 3);
            Merge_Cell(iRow + 1, 16, "TDS", false, 1, 3);
            Merge_Cell(iRow + 1, 17, "PTAX", false, 1, 3);
            Merge_Cell(iRow + 1, 18, "LOAN", false, 1, 3);
            Merge_Cell(iRow + 1, 19, "LWF", false, 1, 3);
            Merge_Cell(iRow + 1, 20, "OTHER DEDUC-TIONS", false, 1, 3);
            Merge_Cell(iRow, 21, "TOTAL DEDUC-TIONS", true, 1, 4);
            Merge_Cell(iRow, 22, "NET SALARY", true, 1, 4);

            iRow++;
            iRow++;
            iRow++;
        }
        private void FillData()
        {
            int SiNo = 0;
            decimal TotAmt = 0;
            decimal OtherEarns = 0, OtherDeduts = 0, TotOtherEarns = 0, TotOtherDeduts = 0;
            decimal BasicRt = 0, DaRt = 0, TotBasicRt = 0, TotDaRt = 0;
            foreach (DataRow dr in Dt_SalSheet.Rows)
            {
                iRow++;
                OtherEarns = 0;
                OtherDeduts = 0;
                SiNo++;

                WriteData(iRow, 0, SiNo, false, "ALL");
                WriteData(iRow, 1, dr["EMP_NO"].ToString(), false, "ALL");
                WriteData(iRow, 2, dr["EMP_NAME"].ToString(), false, "ALL");
                BasicRt = Lib.Convert2Decimal(dr["SAL_BASIC_RT"].ToString());
                DaRt = Lib.Convert2Decimal(dr["SAL_DA_RT"].ToString());

                // AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, (BasicRt + DaRt).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 3, (BasicRt + DaRt), false, "ALL");

                TotBasicRt += BasicRt;
                TotDaRt += DaRt;

                //AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, (Lib.Convert2Decimal(dr["A01"].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString())).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, dr["A02"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, Lib.Convert2Decimal(dr["A11"].ToString()).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 4, (Lib.Convert2Decimal(dr["A01"].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString())), false, "ALL");
                WriteData(iRow, 5, dr["A02"], false, "ALL");
                WriteData(iRow, 6, Lib.Convert2Decimal(dr["A11"].ToString()), false, "ALL");

                TotAmt = Lib.Convert2Decimal(dr["A01"].ToString()) + Lib.Convert2Decimal(dr["A02"].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString());
                TotAmt += Lib.Convert2Decimal(dr["A11"].ToString());


                //AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, Lib.Convert2Decimal(Lib.NumericFormat(TotAmt.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, dr["A03"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, dr["A04"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, dr["A06"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 7, Lib.Convert2Decimal(Lib.NumericFormat(TotAmt.ToString(), 0)), false, "ALL");
                WriteData(iRow, 8, dr["A03"], false, "ALL");
                WriteData(iRow, 9, dr["A04"], false, "ALL");
                WriteData(iRow, 10, dr["A06"], false, "ALL");

                OtherEarns = Lib.Convert2Decimal(dr["A05"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A07"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A08"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A09"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A10"].ToString());

                //OtherEarns += Common.Convert2Decimal(dr["A11"].ToString()); added in col 5

                OtherEarns += Lib.Convert2Decimal(dr["A12"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A13"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A14"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A15"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A16"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A17"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A18"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A19"].ToString());
                //A20 added with basic
                OtherEarns += Lib.Convert2Decimal(dr["A21"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A22"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A23"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A24"].ToString());
                OtherEarns += Lib.Convert2Decimal(dr["A25"].ToString());
                TotOtherEarns += OtherEarns;

                //AddXYLabel(HCOL12, Row, ROW_HT, HCOL13 - HCOL12, Lib.Convert2Decimal(Lib.NumericFormat(OtherEarns.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 11, Lib.Convert2Decimal(Lib.NumericFormat(OtherEarns.ToString(), 0)), false, "ALL");

                decimal sLop = 0;
                int sMonth = 0, sYear = 0;
                string eStatus = "";
                sLop = Lib.Convert2Decimal(dr["SAL_LOP_AMT"].ToString());//SAL_LOP_AMT  
                sMonth = Lib.Conv2Integer(dr["SAL_MONTH"].ToString());
                sYear = Lib.Conv2Integer(dr["SAL_YEAR"].ToString());

                DateTime dtime;
                if (dr["EMP_DO_JOINING"].ToString().Trim() != "" && Lib.Convert2Decimal(dr["SAL_DAYS_WORKED"].ToString()) + 1 < DateTime.DaysInMonth(sYear, sMonth))
                {
                    dtime = (DateTime)dr["EMP_DO_JOINING"];
                    if (dtime.Month == sMonth && dtime.Year == sYear)
                        eStatus = "DOJ-" + dr["SAL_DAYS_WORKED"].ToString();
                }

                if (dr["EMP_DO_RELIEVE"].ToString().Trim() != "" && Lib.Convert2Decimal(dr["SAL_DAYS_WORKED"].ToString()) + 1 < DateTime.DaysInMonth(sYear, sMonth))
                {
                    dtime = (DateTime)dr["EMP_DO_RELIEVE"];
                    if (dtime.Month == sMonth && dtime.Year == sYear)
                        eStatus = "DOR-" + dr["SAL_DAYS_WORKED"].ToString();
                }

                //if (eStatus != "")
                //    AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, eStatus, ifontName, ifontSize, "LTBR", "C");
                //else
                //    AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, sLop.ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);

                if (eStatus != "")
                    WriteData(iRow, 12, eStatus, false, "ALL");
                else
                    WriteData(iRow, 12, sLop, false, "ALL");

                //  AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, dr["SAL_GROSS_EARN"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 13, dr["SAL_GROSS_EARN"], false, "ALL");

                //AddXYLabel(HCOL15, Row, ROW_HT, HCOL16 - HCOL15, dr["D01"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL16, Row, ROW_HT, HCOL17 - HCOL16, dr["D02"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL17, Row, ROW_HT, HCOL18 - HCOL17, dr["D03"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL18, Row, ROW_HT, HCOL19 - HCOL18, dr["D09"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL19, Row, ROW_HT, HCOL20 - HCOL19, dr["D06"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL20, Row, ROW_HT, HCOL21 - HCOL20, dr["D13"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 14, dr["D01"], false, "ALL");
                WriteData(iRow, 15, dr["D02"], false, "ALL");
                WriteData(iRow, 16, dr["D03"], false, "ALL");
                WriteData(iRow, 17, dr["D09"], false, "ALL");
                WriteData(iRow, 18, dr["D06"], false, "ALL");
                WriteData(iRow, 19, dr["D13"], false, "ALL");

                OtherDeduts = Lib.Convert2Decimal(dr["D04"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D05"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D07"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D08"].ToString());
                //OtherDeduts += Common.Convert2Decimal(dr["D09"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D10"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D11"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D12"].ToString());
                //OtherDeduts += Common.Convert2Decimal(dr["D13"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D14"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D15"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D16"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D17"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D18"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D19"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D20"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D21"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D22"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D23"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D24"].ToString());
                OtherDeduts += Lib.Convert2Decimal(dr["D25"].ToString());

                TotOtherDeduts += OtherDeduts;

                //AddXYLabel(HCOL21, Row, ROW_HT, HCOL22 - HCOL21, Lib.Convert2Decimal(Lib.NumericFormat(OtherDeduts.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, dr["SAL_GROSS_DEDUCT"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                //AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, dr["SAL_NET"].ToString(), ifontName, ifontSize, "LTBR", "R", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 20, Lib.Convert2Decimal(Lib.NumericFormat(OtherDeduts.ToString(), 0)), false, "ALL");
                WriteData(iRow, 21, dr["SAL_GROSS_DEDUCT"], false, "ALL");
                WriteData(iRow, 22, dr["SAL_NET"], false, "ALL");
                //WriteData(iRow, 23, dr["REC_BRANCH_CODE"], false, "ALL");
            }

            if (Dt_SalSheet.Rows.Count > 0)
            {

                iRow++;
               //AddXYLabel(HCOL1, Row, ROW_HT, HCOL2 - HCOL1, "", ifontName, ifontSize, "LTBR", "");
                WriteData(iRow, 0, "", true, "ALL");
               // AddXYLabel(HCOL2, Row, ROW_HT, HCOL3 - HCOL2, "", ifontName, ifontSize, "LTBR", "");
                WriteData(iRow, 1, "", true, "ALL");
               // AddXYLabel(HCOL3, Row, ROW_HT, HCOL4 - HCOL3, Emp_Status == "UNCONFIRM" ? "STIPEND TOTAL" : "SALARY TOTAL", ifontName, ifontSize, "LTBR", "B");
                WriteData(iRow, 2, Emp_Status == "UNCONFIRM" ? "STIPEND TOTAL" : "SALARY TOTAL", true, "ALL");
              //  AddXYLabel(HCOL4, Row, ROW_HT, HCOL5 - HCOL4, Lib.Convert2Decimal(Lib.NumericFormat((TotBasicRt + TotDaRt).ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 3, Lib.Convert2Decimal(Lib.NumericFormat((TotBasicRt + TotDaRt).ToString(), 0)), true, "ALL");
              //  AddXYLabel(HCOL5, Row, ROW_HT, HCOL6 - HCOL5, (FindTot("A01") + FindTot("A20")).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 4, (FindTot("A01") + FindTot("A20")), true, "ALL");
               // AddXYLabel(HCOL6, Row, ROW_HT, HCOL7 - HCOL6, FindTot("A02").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 5, FindTot("A02"), true, "ALL");
              //  AddXYLabel(HCOL7, Row, ROW_HT, HCOL8 - HCOL7, FindTot("A11").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 6, FindTot("A11"), true, "ALL");
                TotAmt = FindTot("A01") + FindTot("A02") + FindTot("A20") + FindTot("A11");
               // AddXYLabel(HCOL8, Row, ROW_HT, HCOL9 - HCOL8, Lib.Convert2Decimal(Lib.NumericFormat(TotAmt.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 7, Lib.Convert2Decimal(Lib.NumericFormat(TotAmt.ToString(), 0)), true, "ALL");
              //  AddXYLabel(HCOL9, Row, ROW_HT, HCOL10 - HCOL9, FindTot("A03").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 8, FindTot("A03"), true, "ALL");
              //  AddXYLabel(HCOL10, Row, ROW_HT, HCOL11 - HCOL10, FindTot("A04").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 9, FindTot("A04"), true, "ALL");
              //  AddXYLabel(HCOL11, Row, ROW_HT, HCOL12 - HCOL11, FindTot("A06").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 10, FindTot("A06"), true, "ALL");
              //  AddXYLabel(HCOL12, Row, ROW_HT, HCOL13 - HCOL12, Lib.Convert2Decimal(Lib.NumericFormat(TotOtherEarns.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 11, Lib.Convert2Decimal(Lib.NumericFormat(TotOtherEarns.ToString(), 0)), true, "ALL");
              //  AddXYLabel(HCOL13, Row, ROW_HT, HCOL14 - HCOL13, FindTot("SAL_LOP_AMT").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);//SAL_LOP_AMT
                WriteData(iRow, 12, FindTot("SAL_LOP_AMT"), true, "ALL");
              //  AddXYLabel(HCOL14, Row, ROW_HT, HCOL15 - HCOL14, FindTot("SAL_GROSS_EARN").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 13, FindTot("SAL_GROSS_EARN"), true, "ALL");
              //  AddXYLabel(HCOL15, Row, ROW_HT, HCOL16 - HCOL15, FindTot("D01").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 14, FindTot("D01"), true, "ALL");
              //  AddXYLabel(HCOL16, Row, ROW_HT, HCOL17 - HCOL16, FindTot("D02").ToString(), ifontName, ifontSize, "LTBR", "RBR", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 15, FindTot("D02"), true, "ALL");
              //  AddXYLabel(HCOL17, Row, ROW_HT, HCOL18 - HCOL17, FindTot("D03").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 16, FindTot("D03"), true, "ALL");
              //  AddXYLabel(HCOL18, Row, ROW_HT, HCOL19 - HCOL18, FindTot("D09").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 17, FindTot("D09"), true, "ALL");
              //  AddXYLabel(HCOL19, Row, ROW_HT, HCOL20 - HCOL19, FindTot("D06").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 18, FindTot("D06"), true, "ALL");
              //  AddXYLabel(HCOL20, Row, ROW_HT, HCOL21 - HCOL20, FindTot("D13").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 19, FindTot("D13"), true, "ALL");
              //  AddXYLabel(HCOL21, Row, ROW_HT, HCOL22 - HCOL21, Lib.Convert2Decimal(Lib.NumericFormat(TotOtherDeduts.ToString(), 0)).ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 20, Lib.Convert2Decimal(Lib.NumericFormat(TotOtherDeduts.ToString(), 0)), true, "ALL");
              //  AddXYLabel(HCOL22, Row, ROW_HT, HCOL23 - HCOL22, FindTot("SAL_GROSS_DEDUCT").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 21, FindTot("SAL_GROSS_DEDUCT"), true, "ALL");
              //  AddXYLabel(HCOL23, Row, ROW_HT, HCOL24 - HCOL23, FindTot("SAL_NET").ToString(), ifontName, ifontSize, "LTBR", "RB", 0, 0, 0, 0, 0, 0, Xtolrnce);
                WriteData(iRow, 22, FindTot("SAL_NET"), true, "ALL");
            }

        }

        private void WriteData(int _Row, int _Col, Object sData)
        {
            WriteData(_Row, _Col, sData, false, System.Drawing.Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, string BORDERS)
        {
            WriteData(_Row, _Col, sData, false, System.Drawing.Color.Black, BORDERS);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold)
        {
            WriteData(_Row, _Col, sData, bBold, System.Drawing.Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, string BORDERS)
        {
            WriteData(_Row, _Col, sData, bBold, System.Drawing.Color.Black, BORDERS);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, System.Drawing.Color c, string BORDERS)
        {
            ws.Cells[_Row, _Col].Value = sData;
            if (bBold)
                ws.Cells[_Row, _Col].Style.Font.Weight = ExcelFont.BoldWeight;
            ws.Cells[_Row, _Col].Style.Font.Color = c;
            if (BORDERS == "ALL")
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, System.Drawing.Color.Black, LineStyle.Thin);
            else if (BORDERS == "NFORMAT")
                ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
            else if (BORDERS == "R_FORMAT")
            {
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, System.Drawing.Color.Black, LineStyle.Thin);
            }
            else if (BORDERS == "L_FORMAT")
            {
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Left, System.Drawing.Color.Black, LineStyle.Thin);
            }
        }
        private void Merge_Cell(int _Row, int _Col, int _Width, int _Height)
        {
            myCell = ws.Cells.GetSubrangeRelative(_Row, _Col, _Width, _Height);
            myCell.Merged = true;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = "Arial";
            myCell.Style.Font.Size = 8 * 20;
        }
        private void Merge_Cell(int _Row, int _Col, object sData, bool fBold, int _Width, int _Height, string FontName = "Arial", string cBorders = "ALL")
        {
            myCell = ws.Cells.GetSubrangeRelative(_Row, _Col, _Width, _Height);
            myCell.Merged = true;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = FontName;
            myCell.Style.Font.Size = 8 * 20;
            if (fBold)
                myCell.Style.Font.Weight = ExcelFont.BoldWeight;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            if (cBorders == "ALL")
                myCell.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
            myCell.Value = sData;
        }

    }
}
