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

    public class OsService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        DataTable Dt_Summary = new DataTable();
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

        string File_Display_Name = "";
        string File_Name = "";
        string report_folder = "";
        string PKID = "";
        string subject = "";
        string message = "";
        string email_type = "";
        string company_code = "";
        string branch_code = "";
        string user_code = "";
        string user_pkid = "";
        public IDictionary<string, object> ProcessOSReport(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();

            try
            {

                if (SearchData.ContainsKey("company_code"))
                    company_code = SearchData["company_code"].ToString();
                if (SearchData.ContainsKey("branch_code"))
                    branch_code = SearchData["branch_code"].ToString();
                if (SearchData.ContainsKey("user_code"))
                    user_code = SearchData["user_code"].ToString();
                if (SearchData.ContainsKey("user_pkid"))
                    user_pkid = SearchData["user_pkid"].ToString();

                if (SearchData.ContainsKey("email_type"))
                    email_type = SearchData["email_type"].ToString();

                if (SearchData.ContainsKey("report_folder"))
                    report_folder = SearchData["report_folder"].ToString();

                PKID = System.Guid.NewGuid().ToString().ToUpper();

                Con_Oracle = new DBConnection();

                if (email_type == "OS-ALL" || email_type == "OS-DELHI")
                {
                    sql = " select branch, branch_code,branch_type, ";
                    sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1, ";
                    sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                    sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                    sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                    sql += " sum( case when os_days between 91 and 180  then  balance  else 0 end) as age5, ";
                    sql += " sum( case when os_days >180  then  balance  else 0 end) as age6, ";
                    sql += " sum( case when os_days >=365  then  balance  else 0 end) as oneyear, ";
                    sql += " sum( case when overdue > 0  then  balance  else 0 end) as overdue, ";

                    sql += " sum(balance) as balance, ";
                    sql += " sum(adv) as advance ";
                    sql += " from os ";
                    sql += " where 1=1 ";
                    if (email_type == "OS-DELHI")
                        sql += " and branch_code in ('DELAF','DELSF') ";

                    sql += " group by branch, branch_code, branch_type ";
                    sql += " order by branch, branch_code ";

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    sql = "";
                    sql += " select branch, branch_code, branch_type, op_sman_name, cust_name, ";
                    sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1,  ";
                    sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2,   ";
                    sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                    sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                    sql += " sum( case when os_days between 91 and 180 then  balance  else 0 end) as age5,   ";
                    sql += " sum( case when os_days > 180  then  balance  else 0 end) as age6, ";
                    sql += " sum(balance) as balance,   ";
                    sql += " sum(adv) as advance, ";
                    sql += " sum(case when overdue > 0  then  balance  else 0 end) as overdue, ";
                    sql += " sum( case when os_days >= 365  then  balance  else 0 end) as oneyear ";
                    sql += " from os ";
                    sql += " where 1=1 ";
                    if (email_type == "OS-DELHI")
                        sql += " and branch_code in ('DELAF','DELSF') ";
                    sql += " group by branch,branch_code,branch_type, op_sman_name, cust_name ";
                    sql += " order by branch_type, branch,cust_name ";

                    Dt_Os = new DataTable();
                    Dt_Os = Con_Oracle.ExecuteQuery(sql);

                    /*
                    sql = "select ml_to_ids, ml_cc_ids, ml_bcc_ids from maillist where ml_pkid ='12A34512-6554-33AA-78SD-4228643B232B'";
                    DataTable dt_test = new DataTable();
                    dt_test = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in dt_test.Rows)
                    {
                        SearchData.Add("to_ids", Dr["ml_to_ids"].ToString());
                        SearchData.Add("cc_ids", Dr["ml_cc_ids"].ToString());
                        SearchData.Add("bcc_ids", Dr["ml_bcc_ids"].ToString());
                        break;
                    }
                    dt_test = null;
                    */

                    CreateHTML(email_type);
                    CreateAttachment();

                    subject = "Debtors Age Wise Balance Rs. " + Lib.NumFormat(nTotal.ToString(), 2, true) + " As On " + System.DateTime.Now.ToString("dd/MM/yyyy") + " ";
                    message = sHtml;
                    File_Display_Name = "os.xls";

                    /*
                    SearchData.Add("subject", "Debtors Age Wise Balance Rs. " + Lib.NumFormat(nTotal.ToString(), 2, true) + " As On " + System.DateTime.Now.ToString("dd/MM/yyyy") + " ");
                    SearchData.Add("message", sHtml);
                    SearchData.Add("filename", File_Name);
                    SearchData.Add("filedisplayname", "os.xls");
                    SmtpMail smail = new SmtpMail();
                    bMail = smail.SendEmail(SearchData, out Msg);
                    */
                }
                else
                {

                    sql = "";
                    sql += " select branch,cust_code,cust_name,op_sman_name,cust_crlimit,cust_crdays,";
                    sql += " nvl(jv_od_type,'N') as jv_od_type,jvh_vrno, jvh_date,jv_debit,jv_credit,  ";
                    sql += " balance as balance,adv as advance ,os_days,overdue, ";
                    sql += " case when overdue > 0  then  balance  else 0 end as overdueamt, ";
                    sql += " case when os_days <= 15 then  balance  else 0 end as age1, ";
                    sql += " case when os_days between 16 and 30 then  balance  else 0 end as age2, ";
                    sql += " case when os_days between 31 and 60 then  balance  else 0 end as age3, ";
                    sql += " case when os_days between 61 and 90 then  balance  else 0 end as age4, ";
                    sql += " case when os_days between 91 and 365  then  balance  else 0 end as age5, ";
                    sql += " case when os_days >= 366  then  balance  else 0 end as age6,  ";
                    sql += " case when overdue > 0 and os_days <= 15 then  balance  else 0 end as odage1, ";
                    sql += " case when overdue > 0 and os_days between 16 and 30 then  balance  else 0 end as odage2, ";
                    sql += " case when overdue > 0 and os_days between 31 and 60 then  balance  else 0 end as odage3, ";
                    sql += " case when overdue > 0 and os_days between 61 and 90 then  balance  else 0 end as odage4, ";
                    sql += " case when overdue > 0 and os_days between 91 and 365  then  balance  else 0 end as odage5, ";
                    sql += " case when overdue > 0 and os_days >= 366  then  balance  else 0 end as odage6 , ";
                    sql += " case when nvl(jv_od_type,'N') <> 'N' then  balance  else 0 end as legalamt ";
                    sql += " from os ";
                    sql += " order by overdue desc,cust_name,jvh_vrno, jvh_date ";

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);

                    sql = " select cust_name,op_sman_name,branch,";
                    sql += "  max(cust_crlimit) as cust_crlimit,";
                    sql += "  max(cust_crdays) as cust_crdays,";
                    sql += "  sum(jv_debit) as jv_debit,";
                    sql += "  sum(jv_credit) as jv_credit,";
                    sql += "  sum(balance) as balance,";
                    sql += "  sum(adv) as advance,";
                    sql += "  sum(case when overdue > 0  then  balance  else 0 end) as overdueamt, ";
                    sql += "  sum(case when os_days <= 15 then  balance  else 0 end) as age1, ";
                    sql += "  sum(case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                    sql += "  sum(case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                    sql += "  sum(case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                    sql += "  sum(case when os_days between 91 and 365  then  balance  else 0 end) as age5, ";
                    sql += "  sum(case when os_days >= 366  then  balance  else 0 end) as age6,  ";
                    //sql += "  sum(case when overdue > 0 and os_days <= 15 then  balance  else 0 end) as odage1, ";
                    //sql += "  sum(case when overdue > 0 and os_days between 16 and 30 then  balance  else 0 end) as odage2, ";
                    //sql += "  sum(case when overdue > 0 and os_days between 31 and 60 then  balance  else 0 end) as odage3, ";
                    //sql += "  sum(case when overdue > 0 and os_days between 61 and 90 then  balance  else 0 end) as odage4, ";
                    //sql += "  sum(case when overdue > 0 and os_days between 91 and 365  then  balance  else 0 end) as odage5, ";
                    //sql += "  sum(case when overdue > 0 and os_days >= 366  then  balance  else 0 end) as odage6 , ";
                    sql += "  sum(case when nvl(jv_od_type,'N') <> 'N' then  balance  else 0 end) as legalamt ";
                    sql += "  from os ";
                    sql += "  group by cust_name,branch,op_sman_name";
                    sql += "  order by overdueamt desc,cust_name";

                    Dt_Summary = new DataTable();
                    Dt_Summary = Con_Oracle.ExecuteQuery(sql);


                    Dt_Os = new DataTable();
                    Dt_Os = Dt_List.DefaultView.ToTable(true, "op_sman_name");


                    sql = "select param_pkid,param_email,param_name from param ";
                    sql += " where nvl(rec_locked,'N') != 'Y' and param_type ='SALESMAN' and param_email is not null";
                    Dt_Email = new DataTable();
                    Dt_Email = Con_Oracle.ExecuteQuery(sql);


                    SearchData.Add("to_ids", "");
                    SearchData.Add("cc_ids", "");
                    SearchData.Add("bcc_ids", "");
                    SearchData.Add("subject", "");
                    SearchData.Add("message", "");
                    SearchData.Add("filename", "");
                    SearchData.Add("filedisplayname", "");

                    sql = "select ml_to_ids, ml_cc_ids, ml_bcc_ids from maillist where ml_pkid ='FA289890-B465-48DE-B5D7-6278D4FA9000'";
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

                    string errstr = "";
                    string to_ids = "";
                    
                    foreach (DataRow Dr in Dt_Os.Rows)
                    {
                        SearchData["to_ids"] = "";
                        foreach (DataRow Dremail in Dt_Email.Select("param_name='" + Dr["op_sman_name"].ToString() + "'"))
                        {
                            to_ids = Dremail["param_email"].ToString();
                            SearchData["to_ids"] = Dremail["param_email"].ToString();
                            break;
                        }

                        if (SearchData["to_ids"].ToString().Length > 0)
                        {
                            PKID = System.Guid.NewGuid().ToString().ToUpper();
                            CreateSalesmanHTML(Dr["op_sman_name"].ToString());
                            CreateSalesmanAttachment(Dr["op_sman_name"].ToString());

                            subject = "Debtors Balance Rs. " + Lib.NumFormat(nTotal.ToString(), 2, true) + " Of " + Dr["op_sman_name"].ToString() + " As On " + System.DateTime.Now.ToString("dd/MM/yyyy") + " ";
                            SearchData["subject"] = subject;
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
                                errstr += Dr["op_sman_name"].ToString() + " Failed Err:" + Msg;
                            }
         
                            string sRemarks = subject;
                            if (Msg != "")
                                sRemarks += " Error " + Msg;
                            if (to_ids.Length > 100)
                                to_ids = to_ids.Substring(0, 100);
                            Lib.AuditLog("MAIL", email_type, mStatus, company_code, branch_code, user_code, user_pkid, to_ids, sRemarks);
                        }
                    }
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
            RetData.Add("subject", subject);
            RetData.Add("message", message);
            RetData.Add("filename", File_Name);
            RetData.Add("filetype", "EXCEL");
            RetData.Add("filedisplayname", File_Display_Name);

            return RetData;
        }
        
        private void CreateHTML(string Mail_Type)
        {
            decimal _nage1 = 0, _nage2 = 0, _nage3 = 0, _nage4 = 0, _nage5 = 0, _nage6 = 0, _nadv = 0, _nbal = 0, _noverdue =0, _noneyear=0;

            decimal nage1 = 0, nage2 = 0, nage3 = 0, nage4 = 0, nage5 = 0, nage6 = 0, nadv = 0, nbal = 0, noverdue = 0, noneyear = 0;

            string sCaption = "debtors Ageing Report Of All Branches as on " + System.DateTime.Now.ToString("dd/MM/yyyy");

            sHtml = "";



            sHtml += "<html>";
            sHtml += "<head>";

            sHtml += "<style type='text/css'> ";
            sHtml += "table{border-collapse:collapse;font-family: Calibri; font-size: 11pt;table-layout:fixed;} ";

            sHtml += "td.f {background-color: #ddffdd;border:1px solid black;text-align: left;width:110px;} ";
            sHtml += "td.f1 {background-color: #ddffdd;border:1px solid black;text-align: right;width:110px;} ";

            sHtml += "th {background-color: #ddffdd;border:1px solid black;text-align: center;} ";
            sHtml += "td {border:1px solid black;}";

            sHtml += "th.col1 {width:110px;}";
            sHtml += "th.col2 {width:110px;}";
            sHtml += "th.col3 {width:110px;}";
            sHtml += "th.col4 {width:110px;}";
            sHtml += "th.col5 {width:110px;}";
            sHtml += "th.col6 {width:110px;}";
            sHtml += "th.col7 {width:110px;}";
            sHtml += "th.col8 {width:110px;}";
            sHtml += "th.col9 {width:110px;}";
            sHtml += "th.col10 {width:110px;}";
            sHtml += "th.col11 {width:110px;}";


            sHtml += "td.col1 {background-color: white;text-align: left;} ";
            sHtml += "td.col2 {background-color: lightgreen;text-align: right;} ";
            sHtml += "td.col3 {background-color: yellow;text-align: right;} ";
            sHtml += "td.col4 {background-color: skyblue;text-align: right;} ";
            sHtml += "td.col5 {background-color: aqua;text-align: right;} ";
            sHtml += "td.col6 {background-color: lightcyan;text-align: right;} ";
            sHtml += "td.col7 {background-color: #ffb3b3;text-align: right;} ";
            sHtml += "td.col8 {background-color: linen;text-align: right;} ";
            sHtml += "td.col9 {background-color: ivory;text-align: right;} ";

            sHtml += "td.col10 {background-color: lightblue;text-align: right;} ";
            sHtml += "td.col11 {background-color: #ffe6e6;text-align: right;} ";

            sHtml += "</style> ";

            sHtml += "</head>";
            sHtml += "<body>";
            sHtml += "Dear Sir,";
            sHtml += " <br/><br/>";

            if (Mail_Type == "OS-ALL")
                sHtml += "Please find all india debtors age wise balance as on " + System.DateTime.Now.ToString("dd/MM/yyyy");
            else
                sHtml += "Please find debtors age wise balance of Delhi Air & Delhi Sea as on " + System.DateTime.Now.ToString("dd/MM/yyyy");

            sHtml += "<br/><br/>";

            if (Mail_Type == "OS-ALL")
                sHtml += "AIR BRANCHES";

            sHtml += "<table>";

            
            sHtml += "<tr>";
            sHtml += "<th class='col1'>" + "BRANCH" + "</th>";
            sHtml += "<th class='col2'>" + "0-15" + "</th>";
            sHtml += "<th class='col3'>" + "16-30" + "</th>";
            sHtml += "<th class='col4'>" + "31-60" + "</th>";
            sHtml += "<th class='col5'>" + "61-90" + "</th>";
            sHtml += "<th class='col6'>" + "91-180" + "</th>";
            sHtml += "<th class='col7'>" + "180+" + "</th>";
            sHtml += "<th class='col8'>" + "TOTAL-OS" + "</th>";
            sHtml += "<th class='col9'>" + "ADVANCE" + "</th>";

            sHtml += "<th class='col10'>" + "OVERDUE" + "</th>";
            sHtml += "<th class='col11'>" + "1YEAR" + "</th>";

            sHtml += "</tr>";

            foreach (DataRow Dr in Dt_List.Rows)
            {
                if (Dr["BRANCH_TYPE"].ToString() == "AIR")
                {
                    sHtml += "<tr>";
                    sHtml += "<td class='col1'>" + Dr["BRANCH"].ToString() + "</td>";
                    sHtml += "<td class='col2'>" + Lib.NumFormat(Dr["AGE1"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col3'>" + Lib.NumFormat(Dr["AGE2"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col4'>" + Lib.NumFormat(Dr["AGE3"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col5'>" + Lib.NumFormat(Dr["AGE4"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col6'>" + Lib.NumFormat(Dr["AGE5"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col7'>" + Lib.NumFormat(Dr["AGE6"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col8'>" + Lib.NumFormat(Dr["BALANCE"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col9'>" + Lib.NumFormat(Dr["ADVANCE"].ToString(), 2, true) + "</td>";

                    sHtml += "<td class='col10'>" + Lib.NumFormat(Dr["OVERDUE"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col11'>" + Lib.NumFormat(Dr["ONEYEAR"].ToString(), 2, true) + "</td>";


                    sHtml += "</tr>";

                    _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());
                    _nbal += Lib.Conv2Decimal(Dr["balance"].ToString());
                    _nadv += Lib.Conv2Decimal(Dr["advance"].ToString());

                    _noverdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    _noneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());

                }
            }
            if (Mail_Type == "OS-ALL")
            {
                sHtml += "<tr>";
                sHtml += "<td class='f'>" + "TOTAL" + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage1.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage2.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage3.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage4.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage5.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage6.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nbal.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nadv.ToString(), 2, true) + "</td>";

                sHtml += "<td class='f1'>" + Lib.NumFormat(_noverdue.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_noneyear.ToString(), 2, true) + "</td>";

                sHtml += "</tr>";
            }

            nage1 += _nage1; nage2 += _nage2; nage3 += _nage3; nage4 += _nage4; nage5 += _nage5; nage6 += _nage6; nbal += _nbal; nadv += _nadv; noverdue += _noverdue; noneyear += _noneyear;
            _nage1 = 0; _nage2 = 0; _nage3 = 0; _nage4 = 0; _nage5 = 0; _nage6 = 0; _nbal = 0; _nadv = 0; _noverdue = 0; _noneyear = 0;

            if (Mail_Type == "OS-ALL")
            {
                sHtml += "</table>";

                sHtml += "<br/>";

                sHtml += "SEA BRANCHES";

                sHtml += "<table>";

                sHtml += "<tr>";
                sHtml += "<th class='col1'>" + "BRANCH" + "</th>";
                sHtml += "<th class='col2'>" + "0-15" + "</th>";
                sHtml += "<th class='col3'>" + "16-30" + "</th>";
                sHtml += "<th class='col4'>" + "31-60" + "</th>";
                sHtml += "<th class='col5'>" + "61-90" + "</th>";
                sHtml += "<th class='col6'>" + "91-180" + "</th>";
                sHtml += "<th class='col7'>" + "180+" + "</th>";
                sHtml += "<th class='col8'>" + "TOTAL-OS" + "</th>";
                sHtml += "<th class='col9'>" + "ADVANCE" + "</th>";

                sHtml += "<th class='col10'>" + "OVERDUE" + "</th>";
                sHtml += "<th class='col11'>" + "1YEAR" + "</th>";

                sHtml += "</tr>";

            }

          //  sHtml += "</tr>";

            foreach (DataRow Dr in Dt_List.Rows)
            {
                if (Dr["BRANCH_TYPE"].ToString() == "SEA")
                {
                    sHtml += "<tr>";
                    sHtml += "<td class='col1'>" + Dr["BRANCH"].ToString() + "</td>";
                    sHtml += "<td class='col2'>" + Lib.NumFormat(Dr["AGE1"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col3'>" + Lib.NumFormat(Dr["AGE2"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col4'>" + Lib.NumFormat(Dr["AGE3"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col5'>" + Lib.NumFormat(Dr["AGE4"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col6'>" + Lib.NumFormat(Dr["AGE5"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col7'>" + Lib.NumFormat(Dr["AGE6"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col8'>" + Lib.NumFormat(Dr["BALANCE"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col9'>" + Lib.NumFormat(Dr["ADVANCE"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col10'>" + Lib.NumFormat(Dr["OVERDUE"].ToString(), 2, true) + "</td>";
                    sHtml += "<td class='col11'>" + Lib.NumFormat(Dr["ONEYEAR"].ToString(), 2, true) + "</td>";


                    sHtml += "</tr>";

                    _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());
                    _nbal += Lib.Conv2Decimal(Dr["balance"].ToString());
                    _nadv += Lib.Conv2Decimal(Dr["advance"].ToString());

                    _noverdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    _noneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());

                }
            }

            if (Mail_Type == "OS-ALL")
            {
                sHtml += "<tr>";
                sHtml += "<td class='f'>" + "TOTAL" + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage1.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage2.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage3.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage4.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage5.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nage6.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nbal.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_nadv.ToString(), 2, true) + "</td>";

                sHtml += "<td class='f1'>" + Lib.NumFormat(_noverdue.ToString(), 2, true) + "</td>";
                sHtml += "<td class='f1'>" + Lib.NumFormat(_noneyear.ToString(), 2, true) + "</td>";

                sHtml += "</tr>";
            }

            nage1 += _nage1; nage2 += _nage2; nage3 += _nage3; nage4 += _nage4; nage5 += _nage5; nage6 += _nage6; nbal += _nbal; nadv += _nadv; noverdue += _noverdue; noneyear += _noneyear;
            _nage1 = 0; _nage2 = 0; _nage3 = 0; _nage4 = 0; _nage5 = 0; _nage6 = 0; _nbal = 0; _nadv = 0; _noverdue = 0; _noneyear = 0;

            string str = "TOTAL";
            if (Mail_Type == "OS-ALL")
            {
                str = "GRAND TOTAL";

                sHtml += "</table>";

                sHtml += "<br/>";

                sHtml += "<table>";
            }

            sHtml += "<tr>";
            sHtml += "<td class='f'>" + str + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nage1.ToString(),2,true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nage2.ToString(),2,true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nage3.ToString(),2,true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nage4.ToString(),2,true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nage5.ToString(),2,true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nage6.ToString(),2,true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nbal.ToString(),2,true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(nadv.ToString(),2,true)  + "</td>";

            sHtml += "<td class='f1'>" + Lib.NumFormat(noverdue.ToString(), 2, true) + "</td>";
            sHtml += "<td class='f1'>" + Lib.NumFormat(noneyear.ToString(), 2, true) + "</td>";

            sHtml += "</tr>";
            

            sHtml += "</table>";

            //sHtml += " <br/><br/>";
            //sHtml += "Thanks & Regards<br><br/>";
            //sHtml += "Software Support" + "<br>";
            //sHtml += "EDP Division" + "<br>";
            //sHtml += "Cargomar (P) Ltd" + "<br>";
            //sHtml += "Maradu, Cochin, Kerala" + "<br>";
            sHtml += "</body>";
            sHtml += "</html>";

            nTotal = nbal;

        }
        private void CreateAttachment()
        {

            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";


            Color _Color = Color.Black;
            int _Size = 10;

            iRow = 0;
            iCol = 0;
            try
            {

                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                mSearchData.Add("branch_code", "HOCPL");

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

                File_Display_Name = "os.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "OS";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];


                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 8;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 25;
                WS.Columns[4].Width = 256 * 25;
                WS.Columns[5].Width = 256 * 15;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;
                WS.Columns[9].Width = 256 * 15;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 15;
                WS.Columns[13].Width = 256 * 15;
                WS.Columns[14].Width = 256 * 15;

                iRow = 0; iCol = 1;

                
                WS.Columns[5].Style.NumberFormat = "#0,0.00";
                WS.Columns[6].Style.NumberFormat = "#0,0.00";
                WS.Columns[7].Style.NumberFormat = "#0,0.00";
                WS.Columns[8].Style.NumberFormat = "#0,0.00";
                WS.Columns[9].Style.NumberFormat = "#0,0.00";
                WS.Columns[10].Style.NumberFormat = "#0,0.00";
                WS.Columns[11].Style.NumberFormat = "#0,0.00";
                WS.Columns[12].Style.NumberFormat = "#0,0.00";
                WS.Columns[13].Style.NumberFormat = "#0,0.00";
                WS.Columns[14].Style.NumberFormat = "#0,0.00";

                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", 12, false, 325, "", true);
                _Size = 10;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;

                Lib.WriteData(WS, iRow, 1, str, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, true, "", "L", "", _Size, false, 325, "", true);

                iRow++;
                iRow++;

                Lib.WriteData(WS, iRow, 1, "AGEWISE REPORT", _Color, true, "", "L", "", 12, false, 325, "", true);

                iRow++;
                iRow++;

                iCol = 1;

                Lib.WriteData(WS, iRow, iCol++, "TYPE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "0-15", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "16-30", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OVERDUE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "1YEAR", _Color, true, "BT", "R", "", _Size, false, 325, "", true);

                foreach (DataRow Dr in Dt_Os.Rows)
                {
                    iRow++;
                    iCol = 1;
                    Lib.WriteData(WS, iRow, iCol++, Dr["branch_type"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["branch"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["cust_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["op_sman_name"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE1"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE2"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE3"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE4"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE5"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE6"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);

                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["BALANCE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ADVANCE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["OVERDUE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ONEYEAR"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "", true);
                }

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
        private void CreateSalesmanHTML(string sManName)
        {
            string str = "";
            nTotal = 0;

            sHtml = "";
            sHtml += "<html>";
            sHtml += "<head>";
            sHtml += "<style type='text/css'> ";
            sHtml += "table{border-collapse:collapse;font-family: Calibri; font-size: 09pt;table-layout:fixed;} ";
            sHtml += "td.f {background-color: lightblue;border:1px solid black;text-align: left;width:110px;} ";
            sHtml += "td.f1 {background-color: lightblue;border:1px solid black;text-align: right;width:110px;} ";
            sHtml += "td.f2 {background-color: lightblue;border:1px solid black;text-align: center;width:110px;} ";
            sHtml += "td.f3 {background-color: white;border:1px solid black;text-align: right;font-size: 10pt ;} ";
            sHtml += "th {background-color: lightblue;border:1px solid black;text-align: center;} ";
            sHtml += "td {border:1px solid black;}";

            sHtml += "th.col1 {width:100px;background-color: lightblue;text-align: left;} ";
            sHtml += "th.col2 {width:90px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col3 {width:50px;background-color: lightblue;text-align: center;} ";
            sHtml += "th.col4 {width:90px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col5 {width:90px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col6 {width:100px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col7 {width:100px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col8 {width:80px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col9 {width:80px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col10 {width:80px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col11 {width:80px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col12 {width:80px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col13 {width:80px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col14 {width:90px;background-color: lightblue;text-align: right;} ";
            sHtml += "th.col15 {width:90px;background-color: lightblue;text-align: right;} ";

            sHtml += "td.col1 {background-color: white;text-align: left;} ";
            sHtml += "td.col2 {background-color: white;text-align: right;} ";
            sHtml += "td.col3 {background-color: white;text-align: center;} ";
            sHtml += "td.col4 {background-color: white;text-align: right;} ";
            sHtml += "td.col5 {background-color: white;text-align: right;} ";
            sHtml += "td.col6 {background-color: white;text-align: right;} ";
            sHtml += "td.col7 {background-color: white;text-align: right;} ";
            sHtml += "td.col8 {background-color: white;text-align: right;} ";
            sHtml += "td.col9 {background-color: white;text-align: right;} ";
            sHtml += "td.col10 {background-color: white;text-align: right;} ";
            sHtml += "td.col11 {background-color: white;text-align: right;} ";
            sHtml += "td.col12 {background-color: white;text-align: right;} ";
            sHtml += "td.col13 {background-color: white;text-align: right;} ";
            sHtml += "td.col14 {background-color: white;text-align: right;} ";
            sHtml += "td.col15 {background-color: white;text-align: right;} ";

            sHtml += "</style> ";
            sHtml += "</head>";
            sHtml += "<body>";
            sHtml += "Dear Sir,";
            sHtml += " <br/><br/>";

            sHtml += "Please find debtors age wise balance as on " + System.DateTime.Now.ToString("dd/MM/yyyy");

            sHtml += "<br/><br/>";

            str = "111";
            foreach (DataRow Dr in Dt_Summary.Select("op_sman_name='" + sManName + "'", "cust_name"))
            {
                if (str != Dr["cust_name"].ToString())
                {
                    str = Dr["cust_name"].ToString();
                    CreateCustomerHTML(sManName, str);
                }
            }

            sHtml += " <br/><br/>";
            sHtml += "Thanks & Regards<br><br/>";
            sHtml += "NARAYANAN" + "<br>";
            sHtml += "CARGOMAR PRIVATE LTD" + "<br>";
            sHtml += "CARGOMAR HOUSE, III / 695 - B," + "<br>";
            sHtml += "KOTTARAM JUNCTION, MARADU" + "<br>";
            sHtml += "COCHIN - 682304" + "<br>";
            sHtml += "Tel : +91 - 484 - 4131600,2706224  Fax : +91 - 484 - 2706242" + "<br>";
            sHtml += "Email : hogen@cargomar.in" + "<br>";
            sHtml += "Web : www.cargomar.in" + "<br>";
            sHtml += "</body>";
            sHtml += "</html>";
        }
        private void CreateCustomerHTML(string sManName,string CustomerName)
        {
            decimal _nodage1 = 0, _nodage2 = 0, _nodage3 = 0, _nodage4 = 0, _nodage5 = 0, _nodage6 = 0;
            decimal _nage1 = 0, _nage2 = 0, _nage3 = 0, _nage4 = 0, _nage5 = 0, _nage6 = 0;
            decimal ndebit = 0, ncredit = 0, nbalance = 0, nadvance = 0;
            decimal noverdueamt = 0,nlegalamt=0;
            string str = "";
             
            sHtml += "<br/>";
            sHtml += "" + CustomerName + "";
            sHtml += "<table>";
            sHtml += "<tr>";
            sHtml += "<th class='col1'>" + "BRANCH" + "</th>";//CODE
            sHtml += "<th class='col2'>" + "CR-LIMIT" + "</th>";
            sHtml += "<th class='col3'>" + "CR-DAYS" + "</th>";
            sHtml += "<th class='col4'>" + "DEBIT" + "</th>";
            sHtml += "<th class='col5'>" + "CREDIT" + "</th>";
            sHtml += "<th class='col6'>" + "BALANCE" + "</th>";
            sHtml += "<th class='col7'>" + "ADVANCE" + "</th>";
            sHtml += "<th class='col8'>" + "0-15" + "</th>";
            sHtml += "<th class='col9'>" + "16-30" + "</th>";
            sHtml += "<th class='col10'>" + "31-60" + "</th>";
            sHtml += "<th class='col11'>" + "61-90" + "</th>";
            sHtml += "<th class='col12'>" + "91-180" + "</th>";
            sHtml += "<th class='col13'>" + "180+" + "</th>";
            sHtml += "<th class='col14'>" + "OVERDUE-AMT" + "</th>";
            sHtml += "<th class='col15'>" + "LEGAL-AMT" + "</th>";
            sHtml += "</tr>";         
            foreach (DataRow Dr in Dt_Summary.Select("cust_name = '" + CustomerName + "' and op_sman_name='" + sManName + "'","Branch"))
            {
                str = "style='color: red'";

                sHtml += "<tr >";

                sHtml += "<td class='col1'>" + Dr["BRANCH"].ToString() + "</td>";
                sHtml += "<td class='col2'>" + GetFormatNum(Dr["CUST_CRLIMIT"].ToString()) + "</td>";
                sHtml += "<td class='col3'>" + GetFormatNum(Dr["CUST_CRDAYS"].ToString(), true) + "</td>";
                sHtml += "<td class='col4'>" + GetFormatNum(Dr["JV_DEBIT"].ToString()) + "</td>";
                sHtml += "<td class='col5'>" + GetFormatNum(Dr["JV_CREDIT"].ToString()) + "</td>";
                sHtml += "<td class='col6'>" + GetFormatNum(Dr["BALANCE"].ToString()) + "</td>";
                sHtml += "<td class='col7'>" + GetFormatNum(Dr["ADVANCE"].ToString()) + "</td>";
                sHtml += "<td class='col8'>" + GetFormatNum(Dr["AGE1"].ToString()) + "</td>";
                sHtml += "<td class='col9'>" + GetFormatNum(Dr["AGE2"].ToString()) + "</td>";
                sHtml += "<td class='col10'>" + GetFormatNum(Dr["AGE3"].ToString()) + "</td>";
                sHtml += "<td class='col11'>" + GetFormatNum(Dr["AGE4"].ToString()) + "</td>";
                sHtml += "<td class='col12'>" + GetFormatNum(Dr["AGE5"].ToString()) + "</td>";
                sHtml += "<td class='col13'>" + GetFormatNum(Dr["AGE6"].ToString()) + "</td>";
                sHtml += "<td " + str + " class='col14'>" + GetFormatNum(Dr["OVERDUEAMT"].ToString()) + "</td>";
                sHtml += "<td " + str + " class='col15'>" + GetFormatNum(Dr["LEGALAMT"].ToString()) + "</td>";
                sHtml += "</tr>";

                ndebit += Lib.Conv2Decimal(Dr["JV_DEBIT"].ToString());
                ncredit += Lib.Conv2Decimal(Dr["JV_CREDIT"].ToString());
                nbalance += Lib.Conv2Decimal(Dr["BALANCE"].ToString());
                nadvance += Lib.Conv2Decimal(Dr["ADVANCE"].ToString());
                noverdueamt += Lib.Conv2Decimal(Dr["OVERDUEAMT"].ToString());
                nlegalamt += Lib.Conv2Decimal(Dr["LEGALAMT"].ToString());

                _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                //_nodage1 += Lib.Conv2Decimal(Dr["odage1"].ToString());
                //_nodage2 += Lib.Conv2Decimal(Dr["odage2"].ToString());
                //_nodage3 += Lib.Conv2Decimal(Dr["odage3"].ToString());
                //_nodage4 += Lib.Conv2Decimal(Dr["odage4"].ToString());
                //_nodage5 += Lib.Conv2Decimal(Dr["odage5"].ToString());
                //_nodage6 += Lib.Conv2Decimal(Dr["odage6"].ToString());
            }

            sHtml += "<tr>";
            sHtml += "<td class='f'>" + "TOTAL" + "</td>";
            sHtml += "<td class='f'>" + "" + "</td>";
            sHtml += "<td class='f'>" + "" + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(ndebit.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(ncredit.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(nbalance.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(nadvance.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(_nage1.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(_nage2.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(_nage3.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(_nage4.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(_nage5.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(_nage6.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(noverdueamt.ToString()) + "</td>";
            sHtml += "<td class='f1'>" + GetFormatNum(nlegalamt.ToString()) + "</td>";
            sHtml += "</tr>";
            sHtml += "</table>";

            nTotal += nbalance;

            /*
            string sHtml1 = "";
            sHtml1 = "";
            sHtml1 += "<table>";

            sHtml1 += "<tr>";
            sHtml1 += "<th class='col2'>" + "SALESPERSON" + "</th>";
            sHtml1 += "<th class='col7'>" + "DEBIT" + "</th>";
            sHtml1 += "<th class='col8'>" + "CREDIT" + "</th>";
            sHtml1 += "<th class='col9'>" + "BALANCE" + "</th>";
            sHtml1 += "<th class='col10'>" + "ADVANCE" + "</th>";
            sHtml1 += "<th class='col12'>" + "OVERDUE-AMT" + "</th>";
            sHtml1 += "<th class='col13'>" + "0-15" + "</th>";
            sHtml1 += "<th class='col14'>" + "16-30" + "</th>";
            sHtml1 += "<th class='col15'>" + "31-60" + "</th>";
            sHtml1 += "<th class='col16'>" + "61-90" + "</th>";
            sHtml1 += "<th class='col17'>" + "91-180" + "</th>";
            sHtml1 += "<th class='col18'>" + "180+" + "</th>";
            sHtml1 += "</tr>";

            sHtml1 += "<tr >";
            sHtml1 += "<td class='col2'>" + sManName + "</td>";
            sHtml1 += "<td class='f3'>" + GetFormatNum(ndebit.ToString()) + "</td>";
            sHtml1 += "<td class='f3'>" + GetFormatNum(ncredit.ToString()) + "</td>";
            sHtml1 += "<td class='f3'>" + GetFormatNum(nbalance.ToString()) + "</td>";
            sHtml1 += "<td class='f3'>" + GetFormatNum(nadvance.ToString()) + "</td>";
            sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(noverdueamt.ToString()) + "</td>";
            sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage1.ToString()) + "</td>";
            sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage2.ToString()) + "</td>";
            sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage3.ToString()) + "</td>";
            sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage4.ToString()) + "</td>";
            sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage5.ToString()) + "</td>";
            sHtml1 += "<td style='color: red' class='f3'>" + GetFormatNum(_nodage6.ToString()) + "</td>";
            sHtml1 += "</tr>";

            sHtml1 += "</table>";

            sHtml = sHtml.Replace("{SUMMARY}", sHtml1);
            */

        }
        private void CreateSalesmanAttachment(string sManName)
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";
            decimal _nage1 = 0, _nage2 = 0, _nage3 = 0, _nage4 = 0, _nage5 = 0, _nage6 = 0;
            decimal ndebit = 0, ncredit = 0, nbalance = 0, nadvance = 0, nlegalamt = 0;

            Color _Color = Color.Black;
            int _Size = 10;

            iRow = 0;
            iCol = 0;
            try
            {


                Dictionary<string, object> mSearchData = new Dictionary<string, object>();
                LovService mService = new LovService();
                mSearchData.Add("table", "ADDRESS");
                mSearchData.Add("branch_code", "HOCPL");

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

                File_Display_Name = "ossalesreport.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 12;
                WS.Columns[2].Width = 256 * 12;
                WS.Columns[3].Width = 256 * 45;
                WS.Columns[4].Width = 256 * 12;
                WS.Columns[5].Width = 256 * 10;
                WS.Columns[6].Width = 256 * 10;
                WS.Columns[7].Width = 256 * 10;
                WS.Columns[8].Width = 256 * 12;
                WS.Columns[9].Width = 256 * 12;
                WS.Columns[10].Width = 256 * 15;
                WS.Columns[11].Width = 256 * 15;
                WS.Columns[12].Width = 256 * 12;
                WS.Columns[13].Width = 256 * 12;
                WS.Columns[14].Width = 256 * 12;
                WS.Columns[15].Width = 256 * 12;
                WS.Columns[16].Width = 256 * 12;
                WS.Columns[17].Width = 256 * 12;
                WS.Columns[18].Width = 256 * 12;
                WS.Columns[19].Width = 256 * 12;
                WS.Columns[20].Width = 256 * 12;

                iRow = 0; iCol = 1;

                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", 12, false, 325, "", true);
                _Size = 10;
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD1, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPADD2, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                str = "";
                if (COMPTEL.Trim() != "")
                    str = "TEL : " + COMPTEL;
                if (COMPFAX.Trim() != "")
                    str += " FAX : " + COMPFAX;

                Lib.WriteData(WS, iRow, 1, str, _Color, true, "", "L", "", _Size, false, 325, "", true);
                iRow++;
                Lib.WriteData(WS, iRow, 1, COMPWEB, _Color, true, "", "L", "", _Size, false, 325, "", true);

                iRow++;
                iRow++;
                Lib.WriteData(WS, iRow, 1, "OS REPORT - SALESMAN : "+sManName, _Color, true, "", "L", "", 12, false, 325, "", true);

                iRow++;
                iRow++;


                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CODE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CUSTOMER", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-LIMIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CR-DAYS", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "VRNO", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OS-DAYS", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OVERDUE-DAYS", _Color, true, "BT", "C", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "0-15", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "16-30", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "LEGAL", _Color, true, "BT", "C", "", _Size, false, 325, "", true);

                _Color = Color.Black;
                Color _ColColor;
                foreach (DataRow Dr in Dt_List.Select("op_sman_name='" + sManName + "'"))
                {
                    iRow++;
                    iCol = 1;
                    if (Lib.Conv2Decimal(Dr["OVERDUE"].ToString()) <= 0)
                        _ColColor = Color.Black;
                    else
                        _ColColor = Color.Red;
                    Lib.WriteData(WS, iRow, iCol++, Dr["BRANCH"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["CUST_CODE"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["CUST_NAME"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["CUST_CRLIMIT"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["CUST_CRDAYS"].ToString()), _Color, false, "", "C", "", _Size, false, 325, "#0;(#0);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Dr["JVH_VRNO"].ToString(), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.DatetoStringDisplayformat(Dr["JVH_DATE"]), _Color, false, "", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["JV_DEBIT"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["JV_CREDIT"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["BALANCE"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["ADVANCE"].ToString()), _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["OS_DAYS"].ToString()), _Color, false, "", "C", "", _Size, false, 325, "#0;(#0);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["OVERDUE"].ToString()), _ColColor, false, "", "C", "", _Size, false, 325, "#0;(#0);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE1"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE2"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE3"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE4"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE5"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["AGE6"].ToString()), _ColColor, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                    Lib.WriteData(WS, iRow, iCol++, Lib.Conv2Decimal(Dr["LEGALAMT"].ToString()), _ColColor, false, "", "C", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);


                    ndebit += Lib.Conv2Decimal(Dr["JV_DEBIT"].ToString());
                    ncredit += Lib.Conv2Decimal(Dr["JV_CREDIT"].ToString());
                    nbalance += Lib.Conv2Decimal(Dr["BALANCE"].ToString());
                    nadvance += Lib.Conv2Decimal(Dr["ADVANCE"].ToString());
                    nlegalamt += Lib.Conv2Decimal(Dr["LEGALAMT"].ToString());

                    _nage1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    _nage2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    _nage3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    _nage4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    _nage5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    _nage6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                }
              
                iRow++;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "TOTAL", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#0;(#0);#", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, ndebit, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, ncredit, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, nbalance, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, nadvance, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#0;(#0);#", true);
                Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "R", "", _Size, false, 325, "#0;(#0);#", true);
                Lib.WriteData(WS, iRow, iCol++, _nage1, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, _nage2, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, _nage3, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, _nage4, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, _nage5, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, _nage6, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);
                Lib.WriteData(WS, iRow, iCol++, nlegalamt, _Color, false, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", true);

                WB.SaveXls(File_Name);
            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }
        }

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
