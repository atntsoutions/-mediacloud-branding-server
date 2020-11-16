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

namespace BLEmail
{
    public class _SALRECORD
    {
        public string ALLOW { get; set; }
        public decimal ALLOW_AMT { get; set; }
        public string DED { get; set; }
        public decimal DED_AMT { get; set; }
    }
    public class PayslipService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_Summary = new DataTable();
        private DataTable Dt_HEAD = null;
        DataTable Dt_Parent;
        string sHtml = "";
        string Msg = "";
        Boolean bMail = false;
        decimal nTotal = 0;

        DataTable Dt_Email = new DataTable();

        DataTable Dt_Os = new DataTable();

        ExcelFile WB;
        ExcelWorksheet WS = null;
        int iRow = 0;
        int iCol = 0;

        private string SAL_ROW_ID = "";
        private string SAL_MAIL_ERROR = "";
        private int SAL_MAIL_SENT = 0;
        private string SAL_GRADE_CONDTION = "";
        private string MSG_SUBJECT = "";

        string File_Display_Name = "";
        string File_Name = "";
        string report_folder = "";
        string PKID = "";
        string SALPKID = "";
        string subject = "";
        string message = "";
        string email_type = "";
        string company_code = "";
        string branch_code = "";
        string user_code = "";
        string user_name = "";
        string user_pkid = "";
        string empstatus = "";
        int salyear = 100;
        int salmonth = -1;

        private string RepBody = "";
        private string SStyle = "";
        string comp_name = "", comp_add1 = "", comp_add2 = "", comp_add3 = "",
               comp_tel = "", comp_fax = "", comp_web = "", comp_email = "", comp_cinno = "", comp_gstin = "", Comp_br_name = "";

        public IDictionary<string, object> Payslip_Mail(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            Boolean bOk = true;
            string sFlag = "";
            SAL_MAIL_ERROR = "";
            SAL_MAIL_SENT = 0;
            string errstr = "";
            string to_ids = "";
            try
            {
                if (SearchData.ContainsKey("company_code"))
                    company_code = SearchData["company_code"].ToString();
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();
                if (SearchData.ContainsKey("user_code"))
                    user_code = SearchData["user_code"].ToString();
                if (SearchData.ContainsKey("user_name"))
                    user_name = SearchData["user_name"].ToString();
                if (SearchData.ContainsKey("user_pkid"))
                    user_pkid = SearchData["user_pkid"].ToString();
                if (SearchData.ContainsKey("email_type"))
                    email_type = SearchData["email_type"].ToString();
                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();
                if (SearchData.ContainsKey("empstatus"))
                    empstatus = SearchData["empstatus"].ToString();
                if (SearchData.ContainsKey("salyear"))
                    salyear = Lib.Conv2Integer(SearchData["salyear"].ToString());
                if (SearchData.ContainsKey("salmonth"))
                    salmonth = Lib.Conv2Integer(SearchData["salmonth"].ToString());
                if (SearchData.ContainsKey("salpkid"))
                    SALPKID = SearchData["salpkid"].ToString();
                
                Con_Oracle = new DBConnection();

                if (email_type == "PAYSLIP-ALL")
                {

                    SALPKID = SALPKID.Replace(",", "','");

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
                    if (SALPKID.Trim() != "")
                    {
                        sql += " where sal_pkid in ('" + SALPKID + "')";
                    }
                    else
                    {
                        sql += "  where a.rec_company_code = '" + company_code + "'";
                        sql += "  and a.rec_branch_code = '" + branch_code + "'";
                        sql += "  and a.sal_month = " + salmonth.ToString();
                        sql += "  and a.sal_year = " + salyear.ToString();
                    }
                    sql += " order by emp_no,sal_year,sal_month";
                    Dt_Parent= new DataTable();
                    Dt_Parent= Con_Oracle.ExecuteQuery(sql);


                    sql = "select * from salaryheadm where rec_company_code ='" + company_code + "' and sal_head is not null order by sal_code";
                    Dt_HEAD = new DataTable();
                    Dt_HEAD = Con_Oracle.ExecuteQuery(sql);
                  

                    SearchData.Add("to_ids", "");
                    SearchData.Add("cc_ids", "");
                    SearchData.Add("bcc_ids", "");
                    SearchData.Add("subject", "");
                    SearchData.Add("message", "");
                    SearchData.Add("filename", "");
                    SearchData.Add("filedisplayname", "");
                    sql = "select ml_to_ids, ml_cc_ids, ml_bcc_ids from maillist where ml_pkid ='3D7D7573-62E4-4001-B8B2-209402AC7FA8'";
                    DataTable dt_temp = new DataTable();
                    dt_temp = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in dt_temp.Rows)
                    {
                        SearchData["to_ids"] = Dr["ml_to_ids"].ToString();
                        SearchData["cc_ids"] = Dr["ml_cc_ids"].ToString();
                        SearchData["bcc_ids"] = Dr["ml_bcc_ids"].ToString();
                        break;
                    }
                    dt_temp = null;
                    Con_Oracle.CloseConnection();

                    PrepareStyleSheet();
                    ReadCompanyDetails();
                    errstr = "";
                    foreach (DataRow Dr in Dt_Parent.Rows)
                    {
                        bOk = true;
                        RepBody = "";
                        sFlag = Dr["SAL_MAIL_SENT"].ToString();
                        if (sFlag.Trim() == "")
                            sFlag = "N";
                        SAL_ROW_ID = Dr["SAL_PKID"].ToString();
                        if (Dr["EMP_EMAIL_PERSONAL"].ToString().Trim().Length <= 0 && Dr["EMP_EMAIL_OFFICE"].ToString().Trim().Length <= 0)
                        {
                            bOk = false;
                            SAL_MAIL_ERROR += "\n Email Id Not Found " + Dr["EMP_NAME"].ToString();
                        }
                        if (bOk && sFlag != "N")
                            SAL_MAIL_ERROR += "\n Email Already Sent " + Dr["EMP_NAME"].ToString();
                        if (bOk && sFlag != "Y")
                        {
                            if (PrepareHtml(Dr))
                            {
                                to_ids = Dr["EMP_EMAIL_PERSONAL"].ToString().Trim();
                                if (to_ids.Length <= 0)
                                    to_ids = Dr["EMP_EMAIL_OFFICE"].ToString();

                                if (branch_code == "HOCPL")
                                {
                                    to_ids = Dr["EMP_EMAIL_OFFICE"].ToString().Trim();
                                    if (to_ids.Length <= 0)
                                        to_ids = Dr["EMP_EMAIL_PERSONAL"].ToString();
                                }

                                SearchData["to_ids"] = to_ids;
                                if (SearchData["to_ids"].ToString().Length > 0)
                                {
                                    sHtml = GetMailBody(Dr);

                                    SearchData["subject"] = MSG_SUBJECT;
                                    SearchData["message"] = sHtml;
                                    SearchData["filename"] = File_Name;
                                    SearchData["filedisplayname"] = File_Display_Name;

                                    Msg = "";
                                    SmtpMail smail = new SmtpMail();
                                    bMail = smail.SendEmail(SearchData, out Msg);
                                    string mStatus = "SENT";
                                    if (bMail == false)
                                    {
                                        mStatus = "FAIL";
                                        if (errstr != "")
                                            errstr += ",";
                                        errstr += Dr["EMP_NAME"].ToString() + " Failed Err:" + Msg;
                                    }
                                    else
                                    {
                                        try
                                        {
                                            sql = "update salarym set sal_mail_sent = 'Y' where sal_pkid = '" + SAL_ROW_ID + "'";
                                            Con_Oracle = new DBConnection();
                                            Con_Oracle.BeginTransaction();
                                            Con_Oracle.ExecuteNonQuery(sql);
                                            Con_Oracle.CommitTransaction();
                                            Con_Oracle.CloseConnection();
                                        }
                                        catch (Exception)
                                        {
                                            if (Con_Oracle != null)
                                            {
                                                Con_Oracle.RollbackTransaction();
                                                Con_Oracle.CloseConnection();
                                            }
                                        }
                                    }
                                    string sRemarks = "EMP.NO-" + Dr["EMP_NO"].ToString() + ", " + MSG_SUBJECT;
                                    if (Msg != "")
                                        sRemarks += " Error " + Msg;
                                    if (to_ids.Length > 100)
                                        to_ids = to_ids.Substring(0, 100);
                                    Lib.AuditLog("MAIL", email_type, mStatus, company_code, branch_code, user_code, user_pkid, to_ids, sRemarks);
                                }
                            }
                        }
                    }
                    if (SAL_MAIL_ERROR.Trim().Length > 0)
                        Msg = SAL_MAIL_ERROR;
                    else
                        Msg = errstr;
                }

                Con_Oracle.CloseConnection();
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
            if (Dt_List != null)
            {
                Dt_List.Rows.Clear();
                Dt_List = null;
            }
            if (Dt_Os != null)
            {
                Dt_Os.Rows.Clear();
                Dt_Os = null;
            }
            if (Dt_Email != null)
            {
                Dt_Email.Rows.Clear();
                Dt_Email = null;
            }

            RetData.Add("retvalue", bMail);
            RetData.Add("error", Msg);
            RetData.Add("subject", MSG_SUBJECT);
            RetData.Add("message", message);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", "EXCEL");
            RetData.Add("filedisplayname", File_Display_Name);

            return RetData;
        }

        private void PrepareStyleSheet()
        {
            SStyle = "";
            SStyle += "<style type='text/css'> ";
            SStyle += "body {font-size:12px;font-family:Ariel;}";
            SStyle += "table {border: 1px solid black;border-collapse:collapse;width:600px;font-size:13px;font-family:Ariel;cellpadding='5'}";
            SStyle += ".bigfont{font-size:18px;font-family:Ariel;}";
            SStyle += ".ht{height:10px;} ";
            SStyle += ".border{border:1px solid black;} ";
            SStyle += ".lborder{border-left:1px solid black;} ";
            SStyle += ".tr1{border:1px solid black;} ";
            SStyle += ".bold{font-weight:bold;} ";
            SStyle += ".lalign{text-align:left;} ";
            SStyle += ".ralign{text-align:right;} ";
            SStyle += ".td1 {width:35%;} ";
            SStyle += ".td2 {width:15%;} ";
            SStyle += ".td3 {width:35%;} ";
            SStyle += ".td4 {width:15%;} ";
            SStyle += "</style> ";
        }
        private bool PrepareHtml(DataRow dr)
        {

            decimal DaysWork = 0;
            Boolean bRet = true;
            string str = "";
            string sFormCaption = "";
           // PaySlip payslip = new PaySlip();
            //String SalMonth = Txt_Month.Text;
            //String SalYear = Txt_Year.Text;



            DaysWork = Lib.Convert2Decimal(dr["SAL_DAYS_WORKED"].ToString());
            if (dr["SAL_DATE"].ToString().Trim() != "")
                str = ((DateTime)dr["SAL_DATE"]).ToString("MMMM").ToUpper() + ", " + dr["SAL_YEAR"].ToString();
            else
                str = new DateTime(salyear, salmonth, 01).ToString("MMMM").ToUpper() + ", " + salyear.ToString();


            MSG_SUBJECT = "PAY SLIP - " + str;


            RepBody += "Dear Sir/Madam,";
            RepBody += " <br/><br/>";
            RepBody += " Please find your pay slip for " + str;
            RepBody += " <br/><br/>";

            RepBody += "<table cellpadding='3'>";

            sFormCaption = Lib.GetFormNumber(dr["REC_BRANCH_CODE"].ToString(), "PAYSLIP");
            if (sFormCaption.Trim() != "")
            {
                RepBody += "<tr><td colspan='4'>" + sFormCaption + "</td></tr>";
            }
            str = "PAY SLIP FOR " + str;
            RepBody += "<tr><td colspan='4' class='border bold'>" + str + "</td></tr>";
            RepBody += "<tr><td colspan='1'>EMP NO</td><td colspan='3'>" + dr["EMP_NO"] + "</td> </tr>";
            RepBody += "<tr><td colspan='1'>NAME</td><td colspan='3'>" + dr["EMP_NAME"] + "</td> </tr>";
            RepBody += "<tr><td colspan='1'>COMPANY</td><td colspan='3'>" + comp_name + "</td> </tr>";
            RepBody += "<tr><td colspan='1'>DESIGNATION</td><td colspan='3'>" + dr["EMP_DESIGNATION"] + "</td> </tr>";
            if ((DaysWork - Decimal.Floor(DaysWork)) != 0)
                RepBody += "<tr><td colspan='1'>DAYS WORKED</td><td colspan='3'>" + Decimal.Floor(DaysWork) + " ½" + "</td> </tr>";
            else
                RepBody += "<tr><td colspan='1'>DAYS WORKED</td><td colspan='3'>" + Decimal.Floor(DaysWork) + "</td> </tr>";


            RepBody += "<tr class='tr1'>";
            RepBody += "<td class='td1 bold border' >EARNINGS</td>";
            RepBody += "<td class='td2 bold border ralign' >AMOUNT</td>";
            RepBody += "<td class='td3 bold border'>DEDUCTIONS</td>";
            RepBody += "<td class='td4 bold border ralign' >AMOUNT</td>";
            RepBody += "</tr>";


            string HedColName = "";
            decimal TotBasicDa = 0;
            decimal Pf_Wage_Bal = 0;
            List<_SALRECORD> SALRECORDS = new List<_SALRECORD>();

            foreach (DataRow dh in Dt_HEAD.Select("SAL_CODE LIKE 'A%'", "SAL_HEAD_ORDER"))
            {
                HedColName = dh["SAL_CODE"].ToString().ToUpper();
                if (HedColName == "A20")//specialbasic already print with basic
                    continue;

                //if (HedColName == "A01" || HedColName == "A02")
                //    TotBasicDa += Lib.Convert2Decimal(dr[HedColName].ToString());
                //if (Convert.ToDateTime(dr["SAL_DATE"]) >= Convert.ToDateTime("01/12/2012"))
                //{
                //    if (HedColName == "A11")
                //        TotBasicDa += Lib.Convert2Decimal(dr[HedColName].ToString());
                //}

                //if (Convert.ToDateTime(dr["SAL_DATE"]) >= Convert.ToDateTime("01/12/2014"))
                //{
                //    TotBasicDa = Common.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString());
                //    if (TotBasicDa > 15000)
                //        TotBasicDa = 15000;
                //}

                
                 
                    //Pf_Wage_Bal = Lib.Convert2Decimal(dr["SAL_PF_WAGE_BAL"].ToString());
                    //TotBasicDa = Pf_Wage_Bal;
                    //if (Lib.Convert2Decimal(dr["SAL_PF_LIMIT"].ToString()) > 0)//Special pf limit
                    //    TotBasicDa += Lib.Convert2Decimal(dr["SAL_PF_LIMIT"].ToString());
                    //else if (Lib.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString()) > Lib.Convert2Decimal(dr["SAL_PF_CEL_LIMIT"].ToString()))
                    //    TotBasicDa += Lib.Convert2Decimal(dr["SAL_PF_CEL_LIMIT"].ToString());
                    //else
                    //    TotBasicDa += Lib.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString());
                 

                if (Lib.Convert2Decimal(dr[HedColName].ToString()) != 0)
                {
                    _SALRECORD mREC = new _SALRECORD();
                    mREC.ALLOW = dh["SAL_HEAD"].ToString();
                    if (HedColName == "A01")//for adding special basic(A20) to BASIC
                        mREC.ALLOW_AMT = Lib.Convert2Decimal(dr[HedColName].ToString()) + Lib.Convert2Decimal(dr["A20"].ToString());
                    else
                        mREC.ALLOW_AMT = Lib.Convert2Decimal(dr[HedColName].ToString());
                    SALRECORDS.Add(mREC);
                }
            }

            int iRow = -1;

            foreach (DataRow dh in Dt_HEAD.Select("SAL_CODE LIKE 'D%'", "SAL_HEAD_ORDER"))
            {
                HedColName = dh["SAL_CODE"].ToString().ToUpper();
                if (Lib.Convert2Decimal(dr[HedColName].ToString()) != 0)
                {
                    iRow++;
                    if (iRow < SALRECORDS.Count)
                    {
                        SALRECORDS[iRow].DED = dh["SAL_HEAD"].ToString();
                        SALRECORDS[iRow].DED_AMT = Lib.Convert2Decimal(dr[HedColName].ToString());
                    }
                    else
                    {
                        _SALRECORD mREC = new _SALRECORD();
                        mREC.DED = dh["SAL_HEAD"].ToString();
                        mREC.DED_AMT = Lib.Convert2Decimal(dr[HedColName].ToString());
                        SALRECORDS.Add(mREC);
                    }
                }
            }

            foreach (_SALRECORD _Rec in SALRECORDS)
            {
                WriteRow(_Rec.ALLOW, _Rec.ALLOW_AMT.ToString(), _Rec.DED, _Rec.DED_AMT.ToString());
            }

            //WriteRow("BASIC SALARY", "1000.00", "PF", "200");
            //WriteRow("HRA", "1500.00", "TDS", "100");
            //WriteRow("OTHER ALLOWANCE", "500.00", "", "");

            WriteRow_Blank();
            WriteRow_Blank();
            WriteRow_Blank();
            WriteRow_Blank();
            WriteRow_Blank();


            RepBody += "<tr class='tr1'>";
            RepBody += "<td class='td1 bold border' >GROSS SALARY</td>";
            str = Lib.NumericFormat(Lib.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString()).ToString(), 2);
            RepBody += "<td class='td2 bold border ralign'>" + str + "</td>";
            RepBody += "<td class='td3 bold border' >TOTAL DEDUCTIONS</td>";
            str = Lib.NumericFormat(Lib.Convert2Decimal(dr["SAL_GROSS_DEDUCT"].ToString()).ToString(), 2);
            RepBody += "<td class='td4 bold border ralign'>" + str + "</td>";
            RepBody += "</tr>";


            str = Lib.NumericFormat(Lib.Convert2Decimal(dr["SAL_NET"].ToString()).ToString(), 2);
            RepBody += "<tr>";
            RepBody += "<td class='bold'>NET SALARY</td>";
            RepBody += "<td class='bigfont bold ralign'>" + str + "</td>";
            RepBody += "<td></td>";
            RepBody += "</tr>";

            /*
            str = Common.NumericFormat(TotBasicDa.ToString(), 2);
            RepBody += "<tr>";
            RepBody += "<td >PF CONTRIBUTION(12%) CALCULATED ON RS: </td>";
            RepBody += "<td class='ralign'>" + str + "</td>";
            RepBody += "<td></td>";
            RepBody += "</tr>";
            str = "NET SALARY HAS BEEN CREDITED TO YOUR A/C";
            RepBody += "<tr>";
            RepBody += "<td>" + str + "</td>";
            RepBody += "<td></td>";
            RepBody += "<td></td>";
            RepBody += "</tr>";
            */
            RepBody += "</table>";

            RepBody += "<br/>";


            //decimal Pf_Wage_Bal = 0;
            //if (Convert.ToDateTime(dr["SAL_DATE"]) >= Convert.ToDateTime("01/12/2014"))
            //{
            //    Pf_Wage_Bal = Lib.Convert2Decimal(dr["SAL_PF_WAGE_BAL"].ToString());
            //    TotBasicDa = Pf_Wage_Bal;
            //    if (Common.Convert2Decimal(dr["SAL_PF_LIMIT"].ToString()) > 0)
            //        TotBasicDa += Common.Convert2Decimal(dr["SAL_PF_LIMIT"].ToString());
            //    else if (Common.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString()) > 15000)
            //        TotBasicDa += 15000;
            //    else
            //        TotBasicDa += Common.Convert2Decimal(dr["SAL_GROSS_EARN"].ToString());
            //}
            //if (Convert.ToDateTime(dr["SAL_DATE"]) >= Convert.ToDateTime("01/04/2015"))
            //{
            //    TotBasicDa = Common.Convert2Decimal(dr["SAL_PF_BASE"].ToString());
            //}

            Pf_Wage_Bal = Lib.Convert2Decimal(dr["SAL_PF_WAGE_BAL"].ToString());
            TotBasicDa = Lib.Convert2Decimal(dr["SAL_PF_BASE"].ToString());

            str = "PF CONTRIBUTION(12%) CALCULATED ON RS: " + TotBasicDa.ToString();
            if (Pf_Wage_Bal > 0)
            {
                string[] pMnth = dr["SAL_PF_MON_YEAR"].ToString().Split(',');
                str += " (" + Convert.ToDateTime("01/" + pMnth[0] + "/" + pMnth[1]).ToString("MMMM").ToUpper() + ": " + Pf_Wage_Bal.ToString();
                str += ", " + Convert.ToDateTime(dr["SAL_DATE"]).ToString("MMMM").ToUpper() + ": " + (TotBasicDa - Pf_Wage_Bal).ToString() + ")";
            }

            //str = Common.NumericFormat(TotBasicDa.ToString(), 2);
            //RepBody += "PF CONTRIBUTION(12%) CALCULATED ON RS: " + str;
            RepBody += str;
            RepBody += "<br/>";
            str = dr["EMP_BANK_ACNO"].ToString();
            RepBody += "NET SALARY HAS BEEN CREDITED TO YOUR BANK A/C NO. " + str;
            RepBody += "<br/>";
            RepBody += "<br/>";
            RepBody += " This is a system generated report and hence signature is not required.".ToUpper();


            RepBody += " ";

            return bRet;
        }

        private void WriteRow(string A1, string A2, string D1, string D2)
        {
            if (Lib.Convert2Decimal(A2) != 0)
                A2 = Lib.NumericFormat(Lib.Convert2Decimal(A2).ToString(), 2);
            else
                A2 = "";

            if (Lib.Convert2Decimal(D2) != 0)
                D2 = Lib.NumericFormat(Lib.Convert2Decimal(D2).ToString(), 2);
            else
                D2 = "";

            RepBody += "<tr>";
            RepBody += "<td class='td1 '>" + A1 + "</td>";
            RepBody += "<td class='td2 lborder ralign'>" + A2 + "</td>";
            RepBody += "<td class='td3 lborder'>" + D1 + "</td>";
            RepBody += "<td class='td4 lborder ralign'>" + D2 + "</td>";
            RepBody += "</tr>";
        }
        private void WriteRow_Blank()
        {
            RepBody += "<tr class='ht'>";
            RepBody += "<td class='td1 '></td>";
            RepBody += "<td class='td2 lborder ralign'></td>";
            RepBody += "<td class='td3 lborder'></td>";
            RepBody += "<td class='td4 lborder ralign'></td>";
            RepBody += "</tr>";
        }


        private string GetMailBody(DataRow Dr1)
        {
            string str = "";

            string sHtml = "";

            string MailID = "";

            sHtml += "<!DOCTYPE html>";
            sHtml += "<html>";
            sHtml += "<head>";

            sHtml += SStyle;

            sHtml += "</head>";
            sHtml += "<body>";

            sHtml += RepBody;

            sHtml += " <br/><br/>";

            sHtml += "Thanks & Regards<br><br/>";
            sHtml += user_name + "<br>";


            //ADDRESS
            sHtml += comp_name + "<br>";
            str = comp_add1;
            str += "<br />" + comp_add2;
            str += "<br />" + comp_add3;
            sHtml += str + "<br>";
            str = "TEL : " + comp_tel;
            str += "  FAX : " + comp_fax;
            sHtml += str + "<br>";
            sHtml += "Web : " + comp_web.ToLower() + "<br/>";


            sHtml += "</body>";
            sHtml += "</html>";







            //string from_id = GlobalConstants.User_Email;

            //MailID = Dr1["EMP_EMAIL_PERSONAL"].ToString().Trim();
            //if (MailID.Length <= 0)
            //    MailID = Dr1["EMP_EMAIL_OFFICE"].ToString();

            //if (GlobalConstants.Branch_Code == "HOCPL")
            //{

            //    MailID = Dr1["EMP_EMAIL_OFFICE"].ToString().Trim();
            //    if (MailID.Length <= 0)
            //        MailID = Dr1["EMP_EMAIL_PERSONAL"].ToString();
            //}


            //Cmar.Outlook.Smtp_Emailer se = new Cmar.Outlook.Smtp_Emailer();
            //string MailStatus = "";
            //try
            //{
            //    this.Cursor = Cursors.WaitCursor;

            //    se.SentEmail(from_id, MailID, "", from_id, MSG_SUBJECT, sHtml, "", "", "", "", out MailStatus, true, null);
            //    this.Cursor = Cursors.Default;
            //    if (MailStatus == "Mail Sent Successfully")
            //    {
            //        str = "update salarym set sal_mail_sent = 'Y' where sal_pkid = '" + SAL_ROW_ID + "'";
            //        orConnection.RunSql(str);
            //        Dr1["SAL_MAIL_SENT"] = "Y";
            //        SAL_MAIL_ERROR += "\n Mail Sent " + Dr1["EMP_NAME"].ToString();
            //    }
            //    else
            //        SAL_MAIL_ERROR += "\n" + MailStatus.ToString();
            //    //MessageBox.Show(MailStatus.Trim(), "Mail Send");
            //}
            //catch (Exception ex)
            //{
            //    this.Cursor = Cursors.Default;
            //    MessageBox.Show(ex.ToString(), "Mail Send");
            //}
            //this.Cursor = Cursors.Default;
            return sHtml;
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
                    comp_add1 = Dr["COMP_ADDRESS1"].ToString();
                    comp_add2 = Dr["COMP_ADDRESS2"].ToString();
                    comp_add3 = Dr["COMP_ADDRESS3"].ToString();
                    comp_tel = Dr["COMP_TEL"].ToString();
                    comp_fax = Dr["COMP_FAX"].ToString();
                    comp_web = Dr["COMP_WEB"].ToString();
                    comp_email = Dr["COMP_EMAIL"].ToString();
                    //comp_cinno = Dr["COMP_CINNO"].ToString();
                    //comp_gstin = Dr["COMP_GSTIN"].ToString();
                    break;
                }
            }
        }


        //    private void CreateHTML(string Mail_Type)
        //{
        //    decimal _nage1 = 0, _nage2 = 0, _nage3 = 0, _nage4 = 0, _nage5 = 0, _nage6 = 0, _nadv = 0, _nbal = 0, _noverdue =0, _noneyear=0;

        //    decimal nage1 = 0, nage2 = 0, nage3 = 0, nage4 = 0, nage5 = 0, nage6 = 0, nadv = 0, nbal = 0, noverdue = 0, noneyear = 0;

        //    string sCaption = "debtors Ageing Report Of All Branches as on " + System.DateTime.Now.ToString("dd/MM/yyyy");

        //    sHtml = "";



        //    sHtml += "<html>";
        //    sHtml += "<head>";

        //    sHtml += "<style type='text/css'> ";
        //    sHtml += "table{border-collapse:collapse;font-family: Calibri; font-size: 11pt;table-layout:fixed;} ";

        //    sHtml += "td.f {background-color: #ddffdd;border:1px solid black;text-align: left;width:110px;} ";
        //    sHtml += "td.f1 {background-color: #ddffdd;border:1px solid black;text-align: right;width:110px;} ";

        //    sHtml += "th {background-color: #ddffdd;border:1px solid black;text-align: center;} ";
        //    sHtml += "td {border:1px solid black;}";

        //    sHtml += "th.col1 {width:110px;}";
        //    sHtml += "th.col2 {width:110px;}";
        //    sHtml += "th.col3 {width:110px;}";
        //    sHtml += "th.col4 {width:110px;}";
        //    sHtml += "th.col5 {width:110px;}";
        //    sHtml += "th.col6 {width:110px;}";
        //    sHtml += "th.col7 {width:110px;}";
        //    sHtml += "th.col8 {width:110px;}";
        //    sHtml += "th.col9 {width:110px;}";
        //    sHtml += "th.col10 {width:110px;}";
        //    sHtml += "th.col11 {width:110px;}";


        //    sHtml += "td.col1 {background-color: white;text-align: left;} ";
        //    sHtml += "td.col2 {background-color: lightgreen;text-align: right;} ";
        //    sHtml += "td.col3 {background-color: yellow;text-align: right;} ";
        //    sHtml += "td.col4 {background-color: skyblue;text-align: right;} ";
        //    sHtml += "td.col5 {background-color: aqua;text-align: right;} ";
        //    sHtml += "td.col6 {background-color: lightcyan;text-align: right;} ";
        //    sHtml += "td.col7 {background-color: #ffb3b3;text-align: right;} ";
        //    sHtml += "td.col8 {background-color: linen;text-align: right;} ";
        //    sHtml += "td.col9 {background-color: ivory;text-align: right;} ";

        //    sHtml += "td.col10 {background-color: lightblue;text-align: right;} ";
        //    sHtml += "td.col11 {background-color: #ffe6e6;text-align: right;} ";

        //    sHtml += "</style> ";

        //    sHtml += "</head>";
        //    sHtml += "<body>";
        //    sHtml += "Dear Sir,";
        //    sHtml += " <br/><br/>";

        //    if (Mail_Type == "OS-ALL")
        //        sHtml += "Please find all india debtors age wise balance as on " + System.DateTime.Now.ToString("dd/MM/yyyy");
        //    else
        //        sHtml += "Please find debtors age wise balance of Delhi Air & Delhi Sea as on " + System.DateTime.Now.ToString("dd/MM/yyyy");

        //    sHtml += "<br/><br/>";

        //    if (Mail_Type == "OS-ALL")
        //        sHtml += "AIR BRANCHES";

        //    sHtml += "<table>";

            
        //    sHtml += "<tr>";
        //    sHtml += "<th class='col1'>" + "BRANCH" + "</th>";
        //    sHtml += "<th class='col2'>" + "0-15" + "</th>";
        //    sHtml += "<th class='col3'>" + "16-30" + "</th>";
        //    sHtml += "<th class='col4'>" + "31-60" + "</th>";
        //    sHtml += "<th class='col5'>" + "61-90" + "</th>";
        //    sHtml += "<th class='col6'>" + "91-180" + "</th>";
        //    sHtml += "<th class='col7'>" + "180+" + "</th>";
        //    sHtml += "<th class='col8'>" + "TOTAL-OS" + "</th>";
        //    sHtml += "<th class='col9'>" + "ADVANCE" + "</th>";

        //    sHtml += "<th class='col10'>" + "OVERDUE" + "</th>";
        //    sHtml += "<th class='col11'>" + "1YEAR" + "</th>";

        //    sHtml += "</tr>";

        //    foreach (DataRow Dr in Dt_List.Rows)
        //    {
        //        if (Dr["BRANCH_TYPE"].ToString() == "AIR")
        //        {
        //            sHtml += "<tr>";
        //            sHtml += "<td class='col1'>" + Dr["BRANCH"].ToString() + "</td>";
        //            sHtml += "<td class='col2'>" + Lib.NumFormat(Dr["AGE1"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col3'>" + Lib.NumFormat(Dr["AGE2"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col4'>" + Lib.NumFormat(Dr["AGE3"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col5'>" + Lib.NumFormat(Dr["AGE4"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col6'>" + Lib.NumFormat(Dr["AGE5"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col7'>" + Lib.NumFormat(Dr["AGE6"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col8'>" + Lib.NumFormat(Dr["BALANCE"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col9'>" + Lib.NumFormat(Dr["ADVANCE"].ToString(), 2, true) + "</td>";

        //            sHtml += "<td class='col10'>" + Lib.NumFormat(Dr["OVERDUE"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col11'>" + Lib.NumFormat(Dr["ONEYEAR"].ToString(), 2, true) + "</td>";


        //            sHtml += "</tr>";

        //            _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
        //            _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
        //            _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
        //            _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
        //            _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
        //            _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());
        //            _nbal += Lib.Conv2Decimal(Dr["balance"].ToString());
        //            _nadv += Lib.Conv2Decimal(Dr["advance"].ToString());

        //            _noverdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
        //            _noneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());

        //        }
        //    }
        //    if (Mail_Type == "OS-ALL")
        //    {
        //        sHtml += "<tr>";
        //        sHtml += "<td class='f'>" + "TOTAL" + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage1.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage2.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage3.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage4.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage5.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage6.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nbal.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nadv.ToString(), 2, true) + "</td>";

        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_noverdue.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_noneyear.ToString(), 2, true) + "</td>";

        //        sHtml += "</tr>";
        //    }

        //    nage1 += _nage1; nage2 += _nage2; nage3 += _nage3; nage4 += _nage4; nage5 += _nage5; nage6 += _nage6; nbal += _nbal; nadv += _nadv; noverdue += _noverdue; noneyear += _noneyear;
        //    _nage1 = 0; _nage2 = 0; _nage3 = 0; _nage4 = 0; _nage5 = 0; _nage6 = 0; _nbal = 0; _nadv = 0; _noverdue = 0; _noneyear = 0;

        //    if (Mail_Type == "OS-ALL")
        //    {
        //        sHtml += "</table>";

        //        sHtml += "<br/>";

        //        sHtml += "SEA BRANCHES";

        //        sHtml += "<table>";

        //        sHtml += "<tr>";
        //        sHtml += "<th class='col1'>" + "BRANCH" + "</th>";
        //        sHtml += "<th class='col2'>" + "0-15" + "</th>";
        //        sHtml += "<th class='col3'>" + "16-30" + "</th>";
        //        sHtml += "<th class='col4'>" + "31-60" + "</th>";
        //        sHtml += "<th class='col5'>" + "61-90" + "</th>";
        //        sHtml += "<th class='col6'>" + "91-180" + "</th>";
        //        sHtml += "<th class='col7'>" + "180+" + "</th>";
        //        sHtml += "<th class='col8'>" + "TOTAL-OS" + "</th>";
        //        sHtml += "<th class='col9'>" + "ADVANCE" + "</th>";

        //        sHtml += "<th class='col10'>" + "OVERDUE" + "</th>";
        //        sHtml += "<th class='col11'>" + "1YEAR" + "</th>";

        //        sHtml += "</tr>";

        //    }

        //  //  sHtml += "</tr>";

        //    foreach (DataRow Dr in Dt_List.Rows)
        //    {
        //        if (Dr["BRANCH_TYPE"].ToString() == "SEA")
        //        {
        //            sHtml += "<tr>";
        //            sHtml += "<td class='col1'>" + Dr["BRANCH"].ToString() + "</td>";
        //            sHtml += "<td class='col2'>" + Lib.NumFormat(Dr["AGE1"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col3'>" + Lib.NumFormat(Dr["AGE2"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col4'>" + Lib.NumFormat(Dr["AGE3"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col5'>" + Lib.NumFormat(Dr["AGE4"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col6'>" + Lib.NumFormat(Dr["AGE5"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col7'>" + Lib.NumFormat(Dr["AGE6"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col8'>" + Lib.NumFormat(Dr["BALANCE"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col9'>" + Lib.NumFormat(Dr["ADVANCE"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col10'>" + Lib.NumFormat(Dr["OVERDUE"].ToString(), 2, true) + "</td>";
        //            sHtml += "<td class='col11'>" + Lib.NumFormat(Dr["ONEYEAR"].ToString(), 2, true) + "</td>";


        //            sHtml += "</tr>";

        //            _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
        //            _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
        //            _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
        //            _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
        //            _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
        //            _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());
        //            _nbal += Lib.Conv2Decimal(Dr["balance"].ToString());
        //            _nadv += Lib.Conv2Decimal(Dr["advance"].ToString());

        //            _noverdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
        //            _noneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());

        //        }
        //    }

        //    if (Mail_Type == "OS-ALL")
        //    {
        //        sHtml += "<tr>";
        //        sHtml += "<td class='f'>" + "TOTAL" + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage1.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage2.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage3.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage4.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage5.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nage6.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nbal.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_nadv.ToString(), 2, true) + "</td>";

        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_noverdue.ToString(), 2, true) + "</td>";
        //        sHtml += "<td class='f1'>" + Lib.NumFormat(_noneyear.ToString(), 2, true) + "</td>";

        //        sHtml += "</tr>";
        //    }

        //    nage1 += _nage1; nage2 += _nage2; nage3 += _nage3; nage4 += _nage4; nage5 += _nage5; nage6 += _nage6; nbal += _nbal; nadv += _nadv; noverdue += _noverdue; noneyear += _noneyear;
        //    _nage1 = 0; _nage2 = 0; _nage3 = 0; _nage4 = 0; _nage5 = 0; _nage6 = 0; _nbal = 0; _nadv = 0; _noverdue = 0; _noneyear = 0;

        //    string str = "TOTAL";
        //    if (Mail_Type == "OS-ALL")
        //    {
        //        str = "GRAND TOTAL";

        //        sHtml += "</table>";

        //        sHtml += "<br/>";

        //        sHtml += "<table>";
        //    }

        //    sHtml += "<tr>";
        //    sHtml += "<td class='f'>" + str + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nage1.ToString(),2,true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nage2.ToString(),2,true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nage3.ToString(),2,true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nage4.ToString(),2,true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nage5.ToString(),2,true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nage6.ToString(),2,true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nbal.ToString(),2,true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(nadv.ToString(),2,true)  + "</td>";

        //    sHtml += "<td class='f1'>" + Lib.NumFormat(noverdue.ToString(), 2, true) + "</td>";
        //    sHtml += "<td class='f1'>" + Lib.NumFormat(noneyear.ToString(), 2, true) + "</td>";

        //    sHtml += "</tr>";
            

        //    sHtml += "</table>";

        //    //sHtml += " <br/><br/>";
        //    //sHtml += "Thanks & Regards<br><br/>";
        //    //sHtml += "Software Support" + "<br>";
        //    //sHtml += "EDP Division" + "<br>";
        //    //sHtml += "Cargomar (P) Ltd" + "<br>";
        //    //sHtml += "Maradu, Cochin, Kerala" + "<br>";
        //    sHtml += "</body>";
        //    sHtml += "</html>";

        //    nTotal = nbal;

        //}
        //private void CreateAttachment()
        //{

        //    string str = "";
        //    string COMPNAME = "";
        //    string COMPADD1 = "";
        //    string COMPADD2 = "";
        //    string COMPTEL = "";
        //    string COMPFAX = "";
        //    string COMPWEB = "";


        //    Color _Color = Color.Black;
        //    int _Size = 10;

        //    iRow = 0;
        //    iCol = 0;
        //    try
        //    {

        //        Dictionary<string, object> mSearchData = new Dictionary<string, object>();
        //        LovService mService = new LovService();
        //        mSearchData.Add("table", "ADDRESS");
        //        mSearchData.Add("branch_code", "HOCPL");

        //        DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
        //        if (Dt_CompAddress != null)
        //        {
        //            foreach (DataRow Dr in Dt_CompAddress.Rows)
        //            {
        //                COMPNAME = Dr["COMP_NAME"].ToString();
        //                COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
        //                COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
        //                COMPTEL = Dr["COMP_TEL"].ToString();
        //                COMPFAX = Dr["COMP_FAX"].ToString();
        //                COMPWEB = Dr["COMP_WEB"].ToString();
        //                break;
        //            }
        //        }

        //        File_Display_Name = "os.xls";
        //        File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

        //        string sName = "OS";
        //        WB = new ExcelFile();
        //        WB.Worksheets.Add(sName);
        //        WS = WB.Worksheets[sName];


        //        WS.Columns[0].Width = 256 * 2;
        //        WS.Columns[1].Width = 256 * 8;
        //        WS.Columns[2].Width = 256 * 15;
        //        WS.Columns[3].Width = 256 * 25;
        //        WS.Columns[4].Width = 256 * 25;
        //        WS.Columns[5].Width = 256 * 15;
        //        WS.Columns[6].Width = 256 * 15;
        //        WS.Columns[7].Width = 256 * 15;
        //        WS.Columns[8].Width = 256 * 15;
        //        WS.Columns[9].Width = 256 * 15;
        //        WS.Columns[10].Width = 256 * 15;
        //        WS.Columns[11].Width = 256 * 15;
        //        WS.Columns[12].Width = 256 * 15;
        //        WS.Columns[13].Width = 256 * 15;
        //        WS.Columns[14].Width = 256 * 15;

        //        iRow = 0; iCol = 1;

                
        //        WS.Columns[5].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[6].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[7].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[8].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[9].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[10].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[11].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[12].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[13].Style.NumberFormat = "#0,0.00";
        //        WS.Columns[14].Style.NumberFormat = "#0,0.00";

        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", 12, false, 325, "", true);
        //        _Size = 10;
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, true, "", "L", "", _Size, false, 325, "", true);
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, true, "", "L", "", _Size, false, 325, "", true);
        //        iRow++;
        //        str = "";
        //        if (COMPTEL.Trim() != "")
        //            str = "TEL : " + COMPTEL;
        //        if (COMPFAX.Trim() != "")
        //            str += " FAX : " + COMPFAX;

        //        Lib.WriteData(WS, iRow, 1, str, _Color, true, "", "L", "", _Size, false, 325, "", true);
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, true, "", "L", "", _Size, false, 325, "", true);

        //        iRow++;
        //        iRow++;

        //        Lib.WriteData(WS, iRow, 1, "AGEWISE REPORT", _Color, true, "", "L", "", 12, false, 325, "", true);

        //        iRow++;
        //        iRow++;

        //        iCol = 1;

        //        Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "0-15", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "16-30", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "OVERDUE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "1YEAR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

        //        foreach (DataRow Dr in Dt_Os.Rows)
        //        {
        //            iRow++;
        //            iCol = 1;
        //            Lib.WriteData(WS, iRow, iCol++, Dr["branch_type"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Dr["branch"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Dr["cust_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Dr["op_sman_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE1"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE2"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE3"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE4"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE5"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE6"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);

        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["BALANCE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ADVANCE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["OVERDUE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ONEYEAR"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
        //        }

        //        WB.SaveXls(File_Name);
        //    }
        //    catch (Exception Ex)
        //    {
        //        throw Ex;
        //    }
        //}
        //private void CreateSalesmanHTML(string sManName)
        //{
        //    string str = "";
        //    nTotal = 0;

        //    sHtml = "";
        //    sHtml += "<html>";
        //    sHtml += "<head>";
        //    sHtml += "<style type='text/css'> ";
        //    sHtml += "table{border-collapse:collapse;font-family: Calibri; font-size: 09pt;table-layout:fixed;} ";
        //    sHtml += "td.f {background-color: lightblue;border:1px solid black;text-align: left;width:110px;} ";
        //    sHtml += "td.f1 {background-color: lightblue;border:1px solid black;text-align: right;width:110px;} ";
        //    sHtml += "td.f2 {background-color: lightblue;border:1px solid black;text-align: center;width:110px;} ";
        //    sHtml += "td.f3 {background-color: white;border:1px solid black;text-align: right;font-size: 10pt ;} ";
        //    sHtml += "th {background-color: lightblue;border:1px solid black;text-align: center;} ";
        //    sHtml += "td {border:1px solid black;}";

        //    sHtml += "th.col1 {width:100px;background-color: lightblue;text-align: left;} ";
        //    sHtml += "th.col2 {width:90px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col3 {width:50px;background-color: lightblue;text-align: center;} ";
        //    sHtml += "th.col4 {width:90px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col5 {width:90px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col6 {width:100px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col7 {width:100px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col8 {width:80px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col9 {width:80px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col10 {width:80px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col11 {width:80px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col12 {width:80px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col13 {width:80px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col14 {width:90px;background-color: lightblue;text-align: right;} ";
        //    sHtml += "th.col15 {width:90px;background-color: lightblue;text-align: right;} ";

        //    sHtml += "td.col1 {background-color: white;text-align: left;} ";
        //    sHtml += "td.col2 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col3 {background-color: white;text-align: center;} ";
        //    sHtml += "td.col4 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col5 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col6 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col7 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col8 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col9 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col10 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col11 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col12 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col13 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col14 {background-color: white;text-align: right;} ";
        //    sHtml += "td.col15 {background-color: white;text-align: right;} ";

        //    sHtml += "</style> ";
        //    sHtml += "</head>";
        //    sHtml += "<body>";
        //    sHtml += "Dear Sir,";
        //    sHtml += " <br/><br/>";

        //    sHtml += "Please find debtors age wise balance as on " + System.DateTime.Now.ToString("dd/MM/yyyy");

        //    sHtml += "<br/><br/>";

        //    str = "111";
        //    foreach (DataRow Dr in Dt_Summary.Select("op_sman_name='" + sManName + "'", "cust_name"))
        //    {
        //        if (str != Dr["cust_name"].ToString())
        //        {
        //            str = Dr["cust_name"].ToString();
        //            CreateCustomerHTML(sManName, str);
        //        }
        //    }

        //    sHtml += " <br/><br/>";
        //    sHtml += "Thanks & Regards<br><br/>";
        //    sHtml += "NARAYANAN" + "<br>";
        //    sHtml += "CARGOMAR PRIVATE LTD" + "<br>";
        //    sHtml += "CARGOMAR HOUSE, III / 695 - B," + "<br>";
        //    sHtml += "KOTTARAM JUNCTION, MARADU" + "<br>";
        //    sHtml += "COCHIN - 682304" + "<br>";
        //    sHtml += "Tel : +91 - 484 - 4131600,2706224  Fax : +91 - 484 - 2706242" + "<br>";
        //    sHtml += "Email : hogen@cargomar.in" + "<br>";
        //    sHtml += "Web : www.cargomar.in" + "<br>";
        //    sHtml += "</body>";
        //    sHtml += "</html>";
        //}
        //private void CreateCustomerHTML(string sManName,string CustomerName)
        //{
        //    decimal _nodage1 = 0, _nodage2 = 0, _nodage3 = 0, _nodage4 = 0, _nodage5 = 0, _nodage6 = 0;
        //    decimal _nage1 = 0, _nage2 = 0, _nage3 = 0, _nage4 = 0, _nage5 = 0, _nage6 = 0;
        //    decimal ndebit = 0, ncredit = 0, nbalance = 0, nadvance = 0;
        //    decimal noverdueamt = 0,nlegalamt=0;
        //    string str = "";
             
        //    sHtml += "<br/>";
        //    sHtml += "" + CustomerName + "";
        //    sHtml += "<table>";
        //    sHtml += "<tr>";
        //    sHtml += "<th class='col1'>" + "BRANCH" + "</th>";//CODE
        //    sHtml += "<th class='col2'>" + "CR-LIMIT" + "</th>";
        //    sHtml += "<th class='col3'>" + "CR-DAYS" + "</th>";
        //    sHtml += "<th class='col4'>" + "DEBIT" + "</th>";
        //    sHtml += "<th class='col5'>" + "CREDIT" + "</th>";
        //    sHtml += "<th class='col6'>" + "BALANCE" + "</th>";
        //    sHtml += "<th class='col7'>" + "ADVANCE" + "</th>";
        //    sHtml += "<th class='col8'>" + "0-15" + "</th>";
        //    sHtml += "<th class='col9'>" + "16-30" + "</th>";
        //    sHtml += "<th class='col10'>" + "31-60" + "</th>";
        //    sHtml += "<th class='col11'>" + "61-90" + "</th>";
        //    sHtml += "<th class='col12'>" + "91-180" + "</th>";
        //    sHtml += "<th class='col13'>" + "180+" + "</th>";
        //    sHtml += "<th class='col14'>" + "OVERDUE-AMT" + "</th>";
        //    sHtml += "<th class='col15'>" + "LEGAL-AMT" + "</th>";
        //    sHtml += "</tr>";         
        //    foreach (DataRow Dr in Dt_Summary.Select("cust_name = '" + CustomerName + "' and op_sman_name='" + sManName + "'","Branch"))
        //    {
        //        str = "style='color: red'";

        //        sHtml += "<tr >";

        //        sHtml += "<td class='col1'>" + Dr["BRANCH"].ToString() + "</td>";
        //        sHtml += "<td class='col2'>" + GetFormatNum(Dr["CUST_CRLIMIT"].ToString()) + "</td>";
        //        sHtml += "<td class='col3'>" + GetFormatNum(Dr["CUST_CRDAYS"].ToString(), true) + "</td>";
        //        sHtml += "<td class='col4'>" + GetFormatNum(Dr["JV_DEBIT"].ToString()) + "</td>";
        //        sHtml += "<td class='col5'>" + GetFormatNum(Dr["JV_CREDIT"].ToString()) + "</td>";
        //        sHtml += "<td class='col6'>" + GetFormatNum(Dr["BALANCE"].ToString()) + "</td>";
        //        sHtml += "<td class='col7'>" + GetFormatNum(Dr["ADVANCE"].ToString()) + "</td>";
        //        sHtml += "<td class='col8'>" + GetFormatNum(Dr["AGE1"].ToString()) + "</td>";
        //        sHtml += "<td class='col9'>" + GetFormatNum(Dr["AGE2"].ToString()) + "</td>";
        //        sHtml += "<td class='col10'>" + GetFormatNum(Dr["AGE3"].ToString()) + "</td>";
        //        sHtml += "<td class='col11'>" + GetFormatNum(Dr["AGE4"].ToString()) + "</td>";
        //        sHtml += "<td class='col12'>" + GetFormatNum(Dr["AGE5"].ToString()) + "</td>";
        //        sHtml += "<td class='col13'>" + GetFormatNum(Dr["AGE6"].ToString()) + "</td>";
        //        sHtml += "<td " + str + " class='col14'>" + GetFormatNum(Dr["OVERDUEAMT"].ToString()) + "</td>";
        //        sHtml += "<td " + str + " class='col15'>" + GetFormatNum(Dr["LEGALAMT"].ToString()) + "</td>";
        //        sHtml += "</tr>";

        //        ndebit += Lib.Conv2Decimal(Dr["JV_DEBIT"].ToString());
        //        ncredit += Lib.Conv2Decimal(Dr["JV_CREDIT"].ToString());
        //        nbalance += Lib.Conv2Decimal(Dr["BALANCE"].ToString());
        //        nadvance += Lib.Conv2Decimal(Dr["ADVANCE"].ToString());
        //        noverdueamt += Lib.Conv2Decimal(Dr["OVERDUEAMT"].ToString());
        //        nlegalamt += Lib.Conv2Decimal(Dr["LEGALAMT"].ToString());

        //        _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
        //        _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
        //        _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
        //        _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
        //        _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
        //        _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());

        //        //_nodage1 += Lib.Conv2Decimal(Dr["odage1"].ToString());
        //        //_nodage2 += Lib.Conv2Decimal(Dr["odage2"].ToString());
        //        //_nodage3 += Lib.Conv2Decimal(Dr["odage3"].ToString());
        //        //_nodage4 += Lib.Conv2Decimal(Dr["odage4"].ToString());
        //        //_nodage5 += Lib.Conv2Decimal(Dr["odage5"].ToString());
        //        //_nodage6 += Lib.Conv2Decimal(Dr["odage6"].ToString());
        //    }

        //    sHtml += "<tr>";
        //    sHtml += "<td class='f'>" + "TOTAL" + "</td>";
        //    sHtml += "<td class='f'>" + "" + "</td>";
        //    sHtml += "<td class='f'>" + "" + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(ndebit.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(ncredit.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(nbalance.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(nadvance.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(_nage1.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(_nage2.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(_nage3.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(_nage4.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(_nage5.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(_nage6.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(noverdueamt.ToString()) + "</td>";
        //    sHtml += "<td class='f1'>" + GetFormatNum(nlegalamt.ToString()) + "</td>";
        //    sHtml += "</tr>";
        //    sHtml += "</table>";

        //    nTotal += nbalance;

        //    /*
        //    string sHtml1 = "";
        //    sHtml1 = "";
        //    sHtml1 += "<table>";

        //    sHtml1 += "<tr>";
        //    sHtml1 += "<th class='col2'>" + "SALESPERSON" + "</th>";
        //    sHtml1 += "<th class='col7'>" + "DEBIT" + "</th>";
        //    sHtml1 += "<th class='col8'>" + "CREDIT" + "</th>";
        //    sHtml1 += "<th class='col9'>" + "BALANCE" + "</th>";
        //    sHtml1 += "<th class='col10'>" + "ADVANCE" + "</th>";
        //    sHtml1 += "<th class='col12'>" + "OVERDUE-AMT" + "</th>";
        //    sHtml1 += "<th class='col13'>" + "0-15" + "</th>";
        //    sHtml1 += "<th class='col14'>" + "16-30" + "</th>";
        //    sHtml1 += "<th class='col15'>" + "31-60" + "</th>";
        //    sHtml1 += "<th class='col16'>" + "61-90" + "</th>";
        //    sHtml1 += "<th class='col17'>" + "91-180" + "</th>";
        //    sHtml1 += "<th class='col18'>" + "180+" + "</th>";
        //    sHtml1 += "</tr>";

        //    sHtml1 += "<tr >";
        //    sHtml1 += "<td class='col2'>" + sManName + "</td>";
        //    sHtml1 += "<td class='f3'>" + GetFormatNum(ndebit.ToString()) + "</td>";
        //    sHtml1 += "<td class='f3'>" + GetFormatNum(ncredit.ToString()) + "</td>";
        //    sHtml1 += "<td class='f3'>" + GetFormatNum(nbalance.ToString()) + "</td>";
        //    sHtml1 += "<td class='f3'>" + GetFormatNum(nadvance.ToString()) + "</td>";
        //    sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(noverdueamt.ToString()) + "</td>";
        //    sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage1.ToString()) + "</td>";
        //    sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage2.ToString()) + "</td>";
        //    sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage3.ToString()) + "</td>";
        //    sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage4.ToString()) + "</td>";
        //    sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage5.ToString()) + "</td>";
        //    sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage6.ToString()) + "</td>";
        //    sHtml1 += "</tr>";

        //    sHtml1 += "</table>";

        //    sHtml = sHtml.Replace("{SUMMARY}", sHtml1);
        //    */

        //}
        //private void CreateSalesmanAttachment(string sManName)
        //{
        //    string str = "";
        //    string COMPNAME = "";
        //    string COMPADD1 = "";
        //    string COMPADD2 = "";
        //    string COMPTEL = "";
        //    string COMPFAX = "";
        //    string COMPWEB = "";
        //    decimal _nage1 = 0, _nage2 = 0, _nage3 = 0, _nage4 = 0, _nage5 = 0, _nage6 = 0;
        //    decimal ndebit = 0, ncredit = 0, nbalance = 0, nadvance = 0, nlegalamt = 0;

        //    Color _Color = Color.Black;
        //    int _Size = 10;

        //    iRow = 0;
        //    iCol = 0;
        //    try
        //    {


        //        Dictionary<string, object> mSearchData = new Dictionary<string, object>();
        //        LovService mService = new LovService();
        //        mSearchData.Add("table", "ADDRESS");
        //        mSearchData.Add("branch_code", "HOCPL");

        //        DataTable Dt_CompAddress = mService.Search2Datatable(mSearchData);
        //        if (Dt_CompAddress != null)
        //        {
        //            foreach (DataRow Dr in Dt_CompAddress.Rows)
        //            {
        //                COMPNAME = Dr["COMP_NAME"].ToString();
        //                COMPADD1 = Dr["COMP_ADDRESS1"].ToString();
        //                COMPADD2 = Dr["COMP_ADDRESS2"].ToString();
        //                COMPTEL = Dr["COMP_TEL"].ToString();
        //                COMPFAX = Dr["COMP_FAX"].ToString();
        //                COMPWEB = Dr["COMP_WEB"].ToString();
        //                break;
        //            }
        //        }

        //        File_Display_Name = "ossalesreport.xls";
        //        File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

        //        string sName = "Report";
        //        WB = new ExcelFile();
        //        WB.Worksheets.Add(sName);
        //        WS = WB.Worksheets[sName];

        //        WS.Columns[0].Width = 256 * 2;
        //        WS.Columns[1].Width = 256 * 12;
        //        WS.Columns[2].Width = 256 * 12;
        //        WS.Columns[3].Width = 256 * 45;
        //        WS.Columns[4].Width = 256 * 12;
        //        WS.Columns[5].Width = 256 * 10;
        //        WS.Columns[6].Width = 256 * 10;
        //        WS.Columns[7].Width = 256 * 10;
        //        WS.Columns[8].Width = 256 * 12;
        //        WS.Columns[9].Width = 256 * 12;
        //        WS.Columns[10].Width = 256 * 15;
        //        WS.Columns[11].Width = 256 * 15;
        //        WS.Columns[12].Width = 256 * 12;
        //        WS.Columns[13].Width = 256 * 12;
        //        WS.Columns[14].Width = 256 * 12;
        //        WS.Columns[15].Width = 256 * 12;
        //        WS.Columns[16].Width = 256 * 12;
        //        WS.Columns[17].Width = 256 * 12;
        //        WS.Columns[18].Width = 256 * 12;
        //        WS.Columns[19].Width = 256 * 12;
        //        WS.Columns[20].Width = 256 * 12;

        //        iRow = 0; iCol = 1;

        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", 12, false, 325, "", true);
        //        _Size = 10;
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, true, "", "L", "", _Size, false, 325, "", true);
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, true, "", "L", "", _Size, false, 325, "", true);
        //        iRow++;
        //        str = "";
        //        if (COMPTEL.Trim() != "")
        //            str = "TEL : " + COMPTEL;
        //        if (COMPFAX.Trim() != "")
        //            str += " FAX : " + COMPFAX;

        //        Lib.WriteData(WS, iRow, 1, str, _Color, true, "", "L", "", _Size, false, 325, "", true);
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, true, "", "L", "", _Size, false, 325, "", true);

        //        iRow++;
        //        iRow++;
        //        Lib.WriteData(WS, iRow, 1, "OS REPORT - SALESMAN : "+sManName, _Color, true, "", "L", "", 12, false, 325, "", true);

        //        iRow++;
        //        iRow++;


        //        iCol = 1;
        //        Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "CR-LIMIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "CR-DAYS", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "OS-DAYS", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "OVERDUE-DAYS", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "0-15", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "16-30", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "LEGAL", _Color, true, "BT", "C", "", _Size, false, 325, "", true);

        //        _Color = Color.Black;
        //        Color _ColColor;
        //        foreach (DataRow Dr in Dt_List.Select("op_sman_name='" + sManName + "'"))
        //        {
        //            iRow++;
        //            iCol = 1;
        //            if (Lib.Conv2Decimal(Dr["OVERDUE"].ToString()) <= 0)
        //                _ColColor = Color.Black;
        //            else
        //                _ColColor = Color.Red;
        //            Lib.WriteData(WS, iRow, iCol++, Dr["BRANCH"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Dr["CUST_CODE"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Dr["CUST_NAME"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["CUST_CRLIMIT"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["CUST_CRDAYS"].ToString()), _Color, false, "", "C", "", _Size, false, 325, "#0;(#0);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Dr["JVH_VRNO"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["JVH_DATE"]), _Color, false, "", "L", "", _Size, false, 325, "", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["JV_DEBIT"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["JV_CREDIT"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["BALANCE"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ADVANCE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["OS_DAYS"].ToString()), _Color, false, "", "C", "", _Size, false, 325, "#0;(#0);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["OVERDUE"].ToString()), _ColColor, false, "", "C", "", _Size, false, 325, "#0;(#0);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE1"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE2"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE3"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE4"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE5"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE6"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //            Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["LEGALAMT"].ToString()), _ColColor, false, "", "C", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);


        //            ndebit += Lib.Conv2Decimal(Dr["JV_DEBIT"].ToString());
        //            ncredit += Lib.Conv2Decimal(Dr["JV_CREDIT"].ToString());
        //            nbalance += Lib.Conv2Decimal(Dr["BALANCE"].ToString());
        //            nadvance += Lib.Conv2Decimal(Dr["ADVANCE"].ToString());
        //            nlegalamt += Lib.Conv2Decimal(Dr["LEGALAMT"].ToString());

        //            _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
        //            _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
        //            _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
        //            _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
        //            _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
        //            _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());

        //        }
              
        //        iRow++;
        //        iCol = 1;
        //        Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#0;(#0);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
        //        Lib.WriteData(WS, iRow, iCol++, ndebit, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, ncredit, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, nbalance, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, nadvance, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#0;(#0);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#0;(#0);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, _nage1, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, _nage2, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, _nage3, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, _nage4, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, _nage5, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, _nage6, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
        //        Lib.WriteData(WS, iRow, iCol++, nlegalamt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);

        //        WB.SaveXls(File_Name);
        //    }
        //    catch (Exception Ex)
        //    {
        //        if (Con_Oracle != null)
        //            Con_Oracle.CloseConnection();
        //        throw Ex;
        //    }
        //}

        private string GetFormatNum(string sAmt, bool bRound = false)
        {
            string str = "";
            decimal nAmt = Lib.Conv2Decimal(sAmt);
            if (nAmt != 0)
            {
                if (bRound)
                    str = nAmt.ToString("#0;(#0);#");
                else
                    str = nAmt.ToString("#,0.00;(#,0.00);#");
            }
            return str;
        }
    }

}
