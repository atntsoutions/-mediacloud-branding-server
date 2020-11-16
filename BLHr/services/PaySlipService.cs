using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;
using System.Drawing;

namespace BLHr
{
    public class PaySlipService : BL_Base
    {
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
        public int cYear = 0;
        public int cMonth = 0;

        ExcelFile file;
        ExcelWorksheet ws = null;
        CellRange myCell;

        DataTable Dt_PaySlip = new DataTable();
        DataTable Dt_HEAD = new DataTable();

        string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "",
                comp_tel = "", comp_fax = "", comp_web = "", comp_email = "", comp_cinno = "", comp_gstin = "", Comp_br_name = "";
         
        
        public PaySlipService()
        {
           
        }
        public void ProcessData()
        {
            try
            {
                Init();
                ReadData();
                if (Dt_PaySlip.Rows.Count <= 0)
                    throw new Exception("No Details to Print Or Already Printed...PaySlip");

                if (!AllValid())
                    return;

                string fname = "myreport";
                if (Dt_PaySlip.Rows[0]["SAL_MON_YR"].ToString().Trim().Length > 0)
                    fname = "PAYSLIP-" + branch_code + "-" + Dt_PaySlip.Rows[0]["SAL_MON_YR"].ToString().Replace(" ", "");
                if (fname.Length > 30)
                    fname = fname.Substring(0, 30);
                File_Display_Name = Lib.ProperFileName(fname) + ".xls";
                File_Name = Lib.GetFileName(report_folder, folderid, File_Display_Name);
                File_Type = "xls";
                ImagePath = report_folder + "\\Images";

                file = new ExcelFile();
                file.Worksheets.Add("Report");
                ws = file.Worksheets["Report"];
                ws.PrintOptions.FitWorksheetWidthToPages = 1;
                ReadCompanyDetails();
                SetColumns();
                FillData();
                file.SaveXls(File_Name);
                UpdatePrinted();
            }
            catch (Exception ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw ex;
            }
        }

        private void UpdatePrinted()
        {
            try
            {
                sql = "update salarym set rec_printed ='Y' where sal_pkid in ('" + PKID + "') ";
                Con_Oracle = new DBConnection();
                Con_Oracle.BeginTransaction();
                Con_Oracle.ExecuteNonQuery(sql);
                Con_Oracle.CommitTransaction();
                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                {
                    Con_Oracle.RollbackTransaction();
                    Con_Oracle.CloseConnection();
                }
                throw Ex;
            }
        }
        private void SetColumns()
        {
            ws.Columns[0].Width = 255 * 36;
            ws.Columns[1].Width = 255 * 12;
            ws.Columns[2].Width = 255 * 36;
            ws.Columns[3].Width = 255 * 12;
            //ws.Columns[4].Width = 255 * 15;
            //ws.Columns[5].Width = 255 * 15;
            for (int s = 0; s < 4; s++)
            {
                ws.Columns[s].Style.Font.Name = "Times New Roman";
                ws.Columns[s].Style.Font.Size = 9 * 20;
            }
        }
        private void Init()
        {
            //GenSINos = "";
            //GenSIType = "";
        }
        private bool AllValid()
        {
            bool bRet = true;
            return bRet;
        }
        private void ReadData()
        {

            PKID = PKID.Replace(",", "','");

            sql = " Select SAL_PKID,SAL_MONTH,SAL_YEAR,SAL_EMP_ID,SAL_DATE,upper(trim(to_char(sal_date, 'MONTH')))||'-'||to_char(sal_date, 'YYYY') as SAL_MON_YR	";
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
            sql += "  ,D16 as SAL_LOP_AMT,SAL_DAYS_WORKED,SAL_PF_BAL,SAL_PF_MON_YEAR";
            sql += "  ,EMP_PKID,EMP_NAME,EMP_NO,grd.param_name as EMP_GRADE,desig.param_name as EMP_DESIGNATION";
            sql += "  ,EMP_DO_JOINING,EMP_DO_RELIEVE,SAL_BASIC_RT,SAL_DA_RT";
            sql += "  ,EMP_PAN,EMP_BANK_ACNO,EMP_PFNO,EMP_ESINO, sal_mail_sent,EMP_EMAIL_OFFICE,EMP_EMAIL_PERSONAL";
            sql += "  ,SAL_PF_WAGE_BAL,SAL_PF_LIMIT,SAL_PF_BASE,SAL_PF_CEL_LIMIT ";
            sql += "  ,EMP_FATHER_NAME,SAL_PAY_DATE,a.REC_BRANCH_CODE,a.REC_CATEGORY,null as branch ";
            sql += "  from salarym a";
            sql += "  inner join empm b on a.sal_emp_id = b.emp_pkid";
            sql += "  left join param grd on b.emp_grade_id = grd.param_pkid";
            sql += "  left join param desig on b.emp_designation_id = desig.param_pkid";
            //sql += "  where a.rec_company_code = '" + company_code + "'";
            //sql += "  and a.rec_branch_code = '" + branch_code + "'";
            //sql += "  and a.sal_month = " + cMonth.ToString();
            //sql += "  and a.sal_year = " + cYear.ToString();
            sql += " where sal_pkid in ('" + PKID + "')";
            if (IsAdmin != "Y")
                sql += " and nvl(a.rec_printed,'N') = 'N' ";
            sql += " order by emp_no,sal_year,sal_month";

            Con_Oracle = new DBConnection();
            Dt_PaySlip = new DataTable();
            Dt_PaySlip = Con_Oracle.ExecuteQuery(sql);
            sql = "select * from salaryheadm where rec_company_code ='"+ company_code + "' and sal_head is not null order by sal_code";
            Dt_HEAD = new DataTable();
            Dt_HEAD = Con_Oracle.ExecuteQuery(sql);
            Con_Oracle.CloseConnection();
        }
       
        private void ReadCompanyDetails()
        {
            comp_name = ""; comp_add1 = ""; comp_add2 = ""; comp_add3 = "";
            comp_tel = ""; comp_fax = ""; comp_web = ""; comp_email = ""; comp_cinno = ""; comp_gstin = "";
            Comp_br_name = "";
            Dictionary<string, object> mSearchData = new Dictionary<string, object>();
            LovService mService = new LovService();
            mSearchData.Add("table", "ADDRESS");
            mSearchData.Add("branch_code", branch_code);
            DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
            if (Dt_CompAddress != null)
            {
                foreach (DataRow Dr in Dt_CompAddress.Rows)
                {
                    Comp_br_name = Dr["BR_NAME"].ToString();
                    comp_name = Dr["COMP_NAME"].ToString();
                    //comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    //comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    //comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    //comp_tel = Dr["COMP_TEL"].ToString();
                    //comp_fax = Dr["COMP_FAX"].ToString();
                    //comp_web = Dr["COMP_WEB"].ToString();
                    //comp_email = Dr["COMP_EMAIL"].ToString();
                    //comp_cinno = Dr["COMP_CINNO"].ToString();
                    //comp_gstin = Dr["COMP_GSTIN"].ToString();
                    break;
                }
            }

            //Report_Header1 = comp_name;
            //Report_Header2 = comp_add1;
            //Report_Header3 = comp_add2;
            //if (Report_Header3.Trim() != "" && comp_add3.Trim() != "")
            //    Report_Header3 += ",";
            //Report_Header3 += comp_add3;

            //DateTime dtmstart = new DateTime(cYear, cMonth, 1);
            //Report_Caption = "SALARY/ STIPEND";
            //if (Emp_Status == "CONFIRMED")
            //    Report_Caption = "SALARY";
            //else if (Emp_Status == "UNCONFIRM")
            //    Report_Caption = "STIPEND";
            //Report_Caption += " DETAILS FOR " +  dtmstart.ToString("MMMM").ToUpper() + ", " + cYear.ToString();
             
        }
        private void FillData()
        {
           int iCol = 0;
            decimal DaysWork = 0;
            int EmpCount = 0;
            int Row_Start = 0;
            int iRow = Row_Start;
            foreach (DataRow dr in Dt_PaySlip.Rows)
            {
                EmpCount++;
                DaysWork = Lib.Convert2Decimal(dr["SAL_DAYS_WORKED"].ToString());
                Merge_Cell(iRow, iCol + 0, GetFormNumber(dr["REC_BRANCH_CODE"].ToString(), "PAYSLIP"), false, 4, 1, "Times New Roman", "NO BORDER");
                iRow++;
                if (dr["SAL_DATE"].ToString().Trim() != "")
                    WriteData(iRow, iCol + 0, "PAYSLIP FOR " + Convert.ToDateTime(dr["SAL_DATE"]).ToString("MMMM").ToUpper() + ", " + dr["SAL_YEAR"].ToString(), true);
                else
                    WriteData(iRow, iCol + 0, "PAYSLIP FOR " + Convert.ToDateTime("01/" + cMonth + "/" + cYear).ToString("MMMM").ToUpper() + ", " + cYear, true);
                ws.Cells.GetSubrangeRelative(iRow++, iCol + 0, 4, 1).SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
                ws.Cells.GetSubrangeRelative(iRow, iCol + 0, 4, 5).SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);

                WriteData(iRow, iCol + 0, "EMPLOYEE NO", "LFORMAT"); WriteData(iRow++, iCol + 1, dr["EMP_NO"]);
                WriteData(iRow, iCol + 0, "NAME", "LFORMAT"); WriteData(iRow++, iCol + 1, dr["EMP_NAME"]);
                WriteData(iRow, iCol + 0, "COMPANY", "LFORMAT"); WriteData(iRow++, iCol + 1, comp_name);
                //WriteData(iRow, iCol + 0, "GRADE", "LFORMAT"); WriteData(iRow++, iCol + 1, dr["EMP_GRADE"]);
                WriteData(iRow, iCol + 0, "DESIGNATION", "LFORMAT"); WriteData(iRow++, iCol + 1, dr["EMP_DESIGNATION"]);

                WriteData(iRow, iCol + 0, "DAYS WORKED", "LFORMAT");
                ws.Cells[iRow, iCol + 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                if ((DaysWork - Decimal.Floor(DaysWork)) != 0)
                    WriteData(iRow++, iCol + 1, Decimal.Floor(DaysWork) + " ½");
                else
                    WriteData(iRow++, iCol + 1, Decimal.Floor(DaysWork));

                WriteData(iRow, iCol + 0, "EARNINGS", true, "ALL");
                ws.Cells[iRow, iCol + 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
                WriteData(iRow, iCol + 1, "AMOUNT", true, "ALL");
                WriteData(iRow, iCol + 2, "DEDUCTIONS", true, "ALL");
                ws.Cells[iRow, iCol + 3].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
                WriteData(iRow, iCol + 3, "AMOUNT", true, "ALL");
                iRow++;

                int EarnRow = iRow;
                int DeductRow = iRow;
                string HedColName = "";
                decimal TotBasicDa = 0;
                foreach (DataRow dh in Dt_HEAD.Select("SAL_CODE LIKE 'A%'", "SAL_HEAD_ORDER"))
                {
                    HedColName = dh["SAL_CODE"].ToString().ToUpper();
                    if (HedColName == "A20")//specialbasic already print with basic
                        continue;

                    if (HedColName == "A01" || HedColName == "A02" || HedColName == "A11")
                        TotBasicDa += Lib.Convert2Decimal(dr[HedColName].ToString());
                   
                    if (Lib.Convert2Decimal(dr[HedColName].ToString()) != 0)
                    {
                        WriteData(EarnRow, iCol + 0, dh["SAL_HEAD"]);
                        if (HedColName == "A01")//for adding special basic(A20) to BASIC
                            WriteData(EarnRow++, iCol + 1, Lib.Convert2Decimal(dr[HedColName].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString()), false, "NFORMAT");
                        else
                            WriteData(EarnRow++, iCol + 1, dr[HedColName], false, "NFORMAT");
                    }
                }
                foreach (DataRow dh in Dt_HEAD.Select("SAL_CODE LIKE 'D%'", "SAL_HEAD_ORDER"))
                {
                    HedColName = dh["SAL_CODE"].ToString().ToUpper();
                    if (Lib.Convert2Decimal(dr[HedColName].ToString()) != 0)
                    {
                        WriteData(DeductRow, iCol + 2, dh["SAL_HEAD"]);
                        WriteData(DeductRow++, iCol + 3, dr[HedColName], false, "NFORMAT");
                    }
                }
                if ((EarnRow - iRow) < 10 && (DeductRow - iRow) < 10)
                    EarnRow = iRow + 10;

                if (EarnRow > DeductRow)
                {
                    ws.Cells.GetSubrangeRelative(iRow, iCol + 0, 1, EarnRow - iRow).SetBorders(MultipleBorders.Vertical, Color.Black, LineStyle.Thin);
                    ws.Cells.GetSubrangeRelative(iRow, iCol + 2, 1, EarnRow - iRow).SetBorders(MultipleBorders.Vertical, Color.Black, LineStyle.Thin);
                    ws.Cells.GetSubrangeRelative(iRow, iCol + 3, 1, EarnRow - iRow).SetBorders(MultipleBorders.Vertical, Color.Black, LineStyle.Thin);
                    iRow = EarnRow;
                }
                else
                {
                    ws.Cells.GetSubrangeRelative(iRow, iCol + 0, 1, DeductRow - iRow).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
                    ws.Cells.GetSubrangeRelative(iRow, iCol + 1, 1, DeductRow - iRow).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
                    ws.Cells.GetSubrangeRelative(iRow, iCol + 2, 1, DeductRow - iRow).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
                    ws.Cells.GetSubrangeRelative(iRow, iCol + 3, 1, DeductRow - iRow).SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
                    iRow = DeductRow;
                }
                WriteData(iRow, iCol + 0, "GROSS SALARY", true, "ALL");
                WriteData(iRow, iCol + 1, dr["SAL_GROSS_EARN"], true, "NALLFORMAT");
                WriteData(iRow, iCol + 2, "TOTAL DEDUCTIONS", true, "ALL");
                WriteData(iRow++, iCol + 3, dr["SAL_GROSS_DEDUCT"], true, "NALLFORMAT");

                ws.Cells.GetSubrangeRelative(iRow, iCol + 0, 4, 3).SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
                WriteData(iRow, iCol + 0, "NET SALARY", true);
                WriteData(iRow++, iCol + 1, dr["SAL_NET"], true, "NFORMAT");
                //WriteData(iRow++, iCol + 0, "PF CONTRIBUTION(12%) CALCULATED ON RS: " + TotBasicDa.ToString());


                decimal Pf_Wage_Bal = 0;
                if (Convert.ToDateTime(dr["SAL_DATE"]) >= Convert.ToDateTime("01/12/2014"))
                {
                    Pf_Wage_Bal = Lib.Convert2Decimal(dr["SAL_PF_WAGE_BAL"].ToString());
                    TotBasicDa = Pf_Wage_Bal;
                    if (Lib.Convert2Decimal(dr["SAL_PF_LIMIT"].ToString()) > 0)//Special pf limit
                        TotBasicDa += Lib.Convert2Decimal(dr["SAL_PF_LIMIT"].ToString());
                    else if (Lib.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString()) > Lib.Convert2Decimal(dr["SAL_PF_CEL_LIMIT"].ToString()))
                        TotBasicDa += Lib.Convert2Decimal(dr["SAL_PF_CEL_LIMIT"].ToString());
                    else
                        TotBasicDa += Lib.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString());
                }

                string str = "PF CONTRIBUTION(12%) CALCULATED ON RS: ";
                    str += Lib.Convert2Decimal(dr["SAL_PF_BASE"].ToString());

                if (Lib.Convert2Decimal(dr["SAL_PF_WAGE_BAL"].ToString()) > 0)
                {
                    string[] pMnth = dr["SAL_PF_MON_YEAR"].ToString().Split(',');
                    str += " (" + Convert.ToDateTime("01/" + pMnth[0] + "/" + pMnth[1]).ToString("MMMM").ToUpper() + ": " + Pf_Wage_Bal.ToString();
                    str += ", " + Convert.ToDateTime(dr["SAL_DATE"]).ToString("MMMM").ToUpper() + ": " + (TotBasicDa - Pf_Wage_Bal).ToString() + ")";
                }

                WriteData(iRow++, iCol + 0, str);


                //WriteData(iRow, iCol + 0, "NET SALARY HAS BEEN CREDITED TO YOUR BANK A/C NO. ");//+ dr["EMP_BANK_ACNO"].ToString();

                WriteData(iRow, iCol + 0, "NET SALARY HAS BEEN CREDITED TO YOUR BANK A/C NO. " + dr["EMP_BANK_ACNO"].ToString());


                WriteData(iRow + 1, iCol + 0, "This is a system generated report and hence signature is not required.");
                iRow = iRow + 12;

                if (Dt_PaySlip.Rows.Count > 1)
                {
                    if (EmpCount % 2 == 0)
                    {
                       // Row_Start += 59;
                        Row_Start += 61;
                        iRow = Row_Start;
                        ws.HorizontalPageBreaks.Add(iRow);
                    }
                    else
                    {
                        iRow = Row_Start + 34;
                        ws.Cells.GetSubrangeRelative(iRow - 4, iCol + 0, 4, 1).SetBorders(MultipleBorders.Bottom, Color.Gray, LineStyle.Dashed);
                    }
                }
            }
        }


        private void WriteData(int _Row, int _Col, Object sData)
        {
            WriteData(_Row, _Col, sData, false, Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, string BORDERS)
        {
            WriteData(_Row, _Col, sData, false, Color.Black, BORDERS);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold)
        {
            WriteData(_Row, _Col, sData, bBold, Color.Black, "");
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, string BORDERS)
        {
            WriteData(_Row, _Col, sData, bBold, Color.Black, BORDERS);
        }
        private void WriteData(int _Row, int _Col, Object sData, Boolean bBold, Color c, string BORDERS)
        {
            ws.Cells[_Row, _Col].Value = sData;
            //ws.Cells[iR, iC].Style.Font.Weight = ExcelFont.BoldWeight;
            if (bBold)
                ws.Cells[_Row, _Col].Style.Font.Weight = ExcelFont.BoldWeight;
            ws.Cells[_Row, _Col].Style.Font.Color = c;
            if (BORDERS == "ALL")
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
            else if (BORDERS == "NFORMAT")
                ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
            else if (BORDERS == "NRFORMAT")
            {
                ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Thin);
            }
            else if (BORDERS == "NALLFORMAT")
            {
                ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
            }
            else if (BORDERS == "LFORMAT")
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, LineStyle.Thin);
            else if (BORDERS == "RFORMAT")
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Thin);
            else if (BORDERS == "LRFORMAT")
            {
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, LineStyle.Thin);
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Thin);
            }
            else if (BORDERS == "NFORMAT7")
            {
                ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
                ws.Cells[_Row, _Col].Style.Font.Size = 20 * 7;
            }
            else if (BORDERS == "NFORMAT7ALL")
            {
                ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
                ws.Cells[_Row, _Col].Style.Font.Size = 20 * 7;
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Outside, Color.Gray, LineStyle.Thin);
            }
            else if (BORDERS == "B_R_NFORMAT")
            {
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Thin);
                ws.Cells[_Row, _Col].Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Thin);
                //ws.Cells[_Row, _Col].Style.NumberFormat = "#0.00";
            }
        }

        private void Merge_Cell(int _Row, int _Col, object sData, bool fBold, int _Width, int _Height, string FontName = "Arial", string cBorders = "ALL")
        {
            myCell = ws.Cells.GetSubrangeRelative(_Row, _Col, _Width, _Height);
            myCell.Merged = true;
            myCell.Style.WrapText = true;
            myCell.Style.Font.Name = FontName;
            myCell.Style.Font.Size = 9 * 20;
            if (fBold)
                myCell.Style.Font.Weight = ExcelFont.BoldWeight;
            myCell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            myCell.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            if (cBorders == "ALL")
                myCell.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin);
            myCell.Value = sData;
        }
        public string GetFormNumber(string sBrCode, string sCategory)
        {
            string sFrmNum = "";
            if (sCategory == "WAGES")
            {
                if (sBrCode == "HOCPL" || sBrCode == "COKSF" ||
                    sBrCode == "COKAF" || sBrCode == "SEZSF")
                    sFrmNum = "FORM No. XI";
            }
            else if (sCategory == "PAYSLIP")
            {
                if (sBrCode == "HOCPL" || sBrCode == "COKSF" ||
                    sBrCode == "COKAF" || sBrCode == "SEZSF")
                    sFrmNum = "FORM NO. XIII";
            }
            return sFrmNum;
        }
    }
}
