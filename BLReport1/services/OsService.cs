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
    public class OsService : BL_Base
    {
        DataTable Dt_List = new DataTable();
        ExcelFile WB;
        ExcelWorksheet WS = null;
        List<OsRep> mList = new List<OsRep>();
        OsRep  mrow;
        int iRow = 0;
        int iCol = 0;

        string type = "";
        string subtype = "";
        string report_folder = "";
        string File_Name = "";
        string File_Type = "EXCEL";
        string File_Display_Name = "osreport.xls";
        string PKID = "";
        string company_code = "";
        string branch_code = "";
        string branch = "";
        string year_code = "";
        string sman = "";
        string party = "";
        string previous_component_type = "";

        Boolean isadmin =false;
        Boolean iscompany =false;
        string filter_branch_code = "";
        string filter_sman_id = "";
        string filter_sman_name = "";

        string sCondition = "";

        string period = "";

        string sCaption = "";
       
      
        decimal age1 = 0;
        decimal age2 = 0;
        decimal age3 = 0;
        decimal age4 = 0;
        decimal age5 = 0;
        decimal age6 = 0;
        decimal balance = 0;
        decimal advance = 0;
        decimal overdue = 0;
        decimal legal = 0;
        decimal oneyear = 0;


        decimal dr = 0;
        decimal cr = 0;


        public IDictionary<string, object> List(Dictionary<string, object> SearchData)
        {
            subtype = SearchData["subtype"].ToString();

            isadmin = (Boolean)SearchData["isadmin"];
            iscompany = (Boolean)SearchData["iscompany"];
            filter_branch_code = SearchData["filter_branch_code"].ToString();
            filter_sman_id = SearchData["filter_sman_id"].ToString();
            filter_sman_name = SearchData["filter_sman_name"].ToString();


            if (iscompany == false)
            {
                if (isadmin)
                    sCondition = " branch_code = '" + filter_branch_code + "' ";
                else
                    sCondition = " op_sman_name = '" + filter_sman_name + "' ";
            }

            if (subtype == "OSLIST")
                return OsList(SearchData);
            else if (subtype == "SMANLIST")
                return OsSmanList(SearchData);
            else if (subtype == "SMANPARTYBRANCHWISE")
                return OsSmanPartyBranchList(SearchData);
            else if (subtype == "BRANCHWISE")
                return BranchWiseList(SearchData);
            else if (subtype == "SMANWISE")
                return SmanWiseList(SearchData);
            else if (subtype == "INVWISE")
                return InvWiseList(SearchData);
            else if (subtype == "LEGAL")
                return LegalWiseList(SearchData);
            else
                return null;
        }


        public IDictionary<string, object> OsList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();
            
            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();

               
                

                sCaption = "Branch Wise";

                Con_Oracle = new DBConnection();


                sql += " select branch, branch_code, ";
                sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1, ";
                sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                sql += " sum( case when os_days between 91 and 180  then  balance  else 0 end) as age5, ";
                sql += " sum( case when os_days > 180  then  balance  else 0 end) as age6, ";

                sql += " sum( case when nvl(jv_od_type,'N') <> 'N'  then  balance  else 0 end) as legal, ";

                sql += " sum( case when overdue > 0  then  balance  else 0 end) as overdue, ";

                sql += " sum( case when os_days >= 365  then  balance  else 0 end) as oneyear, ";

                sql += " sum(balance) as balance, ";
                sql += " sum(adv) as advance ";
                sql += " from os ";

                if (sCondition != "")
                    sql += " where " + sCondition;


                sql += " group by branch, branch_code ";
                sql += " order by branch, branch_code ";

                //sql += "   and a.rec_branch_code = '{BRCODE}'";


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new OsRep();
                    mrow.row_type = "DETAIL";
                    mrow.branch = Dr["branch"].ToString();
                    mrow.branch_code = Dr["branch_code"].ToString();
                    mrow.age1 = Lib.Conv2Decimal( Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    mrow.advance= Lib.Conv2Decimal(Dr["advance"].ToString());
                    mrow.overdue= Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.legal = Lib.Conv2Decimal(Dr["legal"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    age1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    age2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    age3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    age4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    age5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    age6 += Lib.Conv2Decimal(Dr["age6"].ToString());
                    

                    overdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                    advance += Lib.Conv2Decimal(Dr["advance"].ToString());
                    legal += Lib.Conv2Decimal(Dr["legal"].ToString());
                    oneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());
                    mList.Add(mrow);

                }

                mrow = new OsRep();
                mrow.row_type = "TOTAL";
                mrow.branch = "TOTAL";
                mrow.branch_code = "TOTAL";
                mrow.row_colour = "RED";
                mrow.age1 = age1;
                mrow.age2 = age2;
                mrow.age3 = age3;
                mrow.age4 = age4;
                mrow.age5 = age5;
                mrow.age6 = age6;
                
                mrow.balance = balance ;
                mrow.advance = advance ;
                mrow.overdue = overdue;
                mrow.legal = legal;
                mrow.oneyear = oneyear;
                mList.Add(mrow);

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintOsReport();
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

        public IDictionary<string, object> OsSmanList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();

            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();
               
                sCaption = "Salesman Wise";
                
                Con_Oracle = new DBConnection();


                sql += " select op_sman_name , ";
                sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1, ";
                sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                sql += " sum( case when os_days between 91 and 180  then  balance  else 0 end) as age5, ";
                sql += " sum( case when os_days > 180  then  balance  else 0 end) as age6, ";

                sql += " sum( case when nvl(jv_od_type,'N') <> 'N'  then  balance  else 0 end) as legal, ";

                sql += " sum( case when overdue > 0  then  balance  else 0 end) as overdue, ";

                sql += " sum( case when os_days >= 365  then  balance  else 0 end) as oneyear, ";

                sql += " sum(balance) as balance, ";
                sql += " sum(adv) as advance ";
                sql += " from os ";

                if (sCondition != "")
                    sql += " where " + sCondition;

                sql += " group by op_sman_name ";
                sql += " order by op_sman_name ";

                //sql += "   and a.rec_branch_code = '{BRCODE}'";


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new OsRep();
                    mrow.row_type = "DETAIL";
                    mrow.sman = Dr["op_sman_name"].ToString();

                    mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    mrow.advance = Lib.Conv2Decimal(Dr["advance"].ToString());
                    mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.legal = Lib.Conv2Decimal(Dr["legal"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    age1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    age2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    age3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    age4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    age5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    age6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                    overdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                    advance += Lib.Conv2Decimal(Dr["advance"].ToString());
                    legal += Lib.Conv2Decimal(Dr["legal"].ToString());
                    oneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    mList.Add(mrow);

                }

                mrow = new OsRep();
                mrow.row_type = "TOTAL";
                mrow.row_colour = "RED";
                mrow.branch = "TOTAL";
                mrow.branch_code = "TOTAL";
                mrow.sman = "TOTAL";
                mrow.age1 = age1;
                mrow.age2 = age2;
                mrow.age3 = age3;
                mrow.age4 = age4;
                mrow.age5 = age5;
                mrow.age6 = age6;
                mrow.balance = balance;
                mrow.advance = advance;
                mrow.overdue = overdue;
                mrow.legal = legal;
                mrow.oneyear = oneyear;
                mList.Add(mrow);

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintOsReport();
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

        public IDictionary<string, object> BranchWiseList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();

            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                branch = SearchData["branch"].ToString();
                year_code = SearchData["year_code"].ToString();
               
                sCaption = branch;
              

                Con_Oracle = new DBConnection();


                sql += " select branch, branch_code,op_sman_name, ";
                sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1, ";
                sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                sql += " sum( case when os_days between 91 and 180  then  balance  else 0 end) as age5, ";
                sql += " sum( case when os_days > 180  then  balance  else 0 end) as age6, ";

                sql += " sum( case when nvl(jv_od_type,'N') <> 'N'  then  balance  else 0 end) as legal, ";

                sql += " sum( case when overdue > 0  then  balance  else 0 end) as overdue, ";

                sql += " sum( case when os_days >= 365  then  balance  else 0 end) as oneyear, ";

                sql += " sum(balance) as balance, ";
                sql += " sum(adv) as advance ";
                sql += " from os where branch_code = '" + branch_code + "'" ;

                if (sCondition != "")
                    sql += " and " + sCondition;

                sql += " group by branch, branch_code, op_sman_name ";
                sql += " order by op_sman_name ";


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new OsRep();
                    mrow.row_type = "DETAIL";
                    mrow.branch = Dr["branch"].ToString();
                    mrow.branch_code = Dr["branch_code"].ToString();
                    mrow.sman = Dr["op_sman_name"].ToString();
                    mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    mrow.advance = Lib.Conv2Decimal(Dr["advance"].ToString());
                    mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.legal = Lib.Conv2Decimal(Dr["legal"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    age1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    age2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    age3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    age4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    age5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    age6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                    overdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                    advance += Lib.Conv2Decimal(Dr["advance"].ToString());
                    legal += Lib.Conv2Decimal(Dr["legal"].ToString());
                    oneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());
                    mList.Add(mrow);

                }

                mrow = new OsRep();
                mrow.row_type = "TOTAL";
                mrow.row_colour = "RED";
                mrow.branch = "TOTAL";
                mrow.branch_code = "TOTAL";
                mrow.sman  = "TOTAL";
                mrow.age1 = age1;
                mrow.age2 = age2;
                mrow.age3 = age3;
                mrow.age4 = age4;
                mrow.age5 = age5;
                mrow.age6 = age6;
                mrow.balance = balance;
                mrow.advance = advance;
                mrow.overdue = overdue;
                mrow.legal = legal;
                mrow.oneyear = oneyear;
                mList.Add(mrow);

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintOsReport();
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

        public IDictionary<string, object> OsSmanPartyBranchList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();

            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch = SearchData["branch"].ToString();
                sman = SearchData["sman"].ToString();


                sCaption =  sman ;


                Con_Oracle = new DBConnection();


                sql += " select cust_name,branch, branch_code, ";
                sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1, ";
                sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                sql += " sum( case when os_days between 91 and 180  then  balance  else 0 end) as age5, ";
                sql += " sum( case when os_days > 180  then  balance  else 0 end) as age6, ";

                sql += " sum( case when nvl(jv_od_type,'N') <> 'N'  then  balance  else 0 end) as legal, ";

                sql += " sum( case when overdue > 0  then  balance  else 0 end) as overdue, ";

                sql += " sum( case when os_days >= 365  then  balance  else 0 end) as oneyear, ";

                sql += " sum(balance) as balance, ";
                sql += " sum(adv) as advance ";

                sql += " from os where ";
                if ( sman == "")
                    sql += " op_sman_name is null ";
                else 
                    sql += " op_sman_name = '" + sman + "'";

                if (sCondition != "")
                    sql += " and " + sCondition;

                sql += " group by cust_name,branch, branch_code ";
                sql += " order by cust_name, branch ";


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new OsRep();
                    mrow.row_type = "DETAIL";
                    mrow.branch = Dr["branch"].ToString();
                    mrow.branch_code = Dr["branch_code"].ToString();
                    mrow.party = Dr["cust_name"].ToString();
                    mrow.sman = sman;
                    mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    mrow.advance = Lib.Conv2Decimal(Dr["advance"].ToString());
                    mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.legal = Lib.Conv2Decimal(Dr["legal"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    age1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    age2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    age3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    age4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    age5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    age6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                    overdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                    advance += Lib.Conv2Decimal(Dr["advance"].ToString());
                    legal += Lib.Conv2Decimal(Dr["legal"].ToString());
                    oneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());
                    mList.Add(mrow);

                }

                mrow = new OsRep();
                mrow.row_type = "TOTAL";
                mrow.row_colour = "RED";
                mrow.branch = "TOTAL";
                mrow.branch_code = "TOTAL";
                mrow.party = "TOTAL";
                mrow.age1 = age1;
                mrow.age2 = age2;
                mrow.age3 = age3;
                mrow.age4 = age4;
                mrow.age5 = age5;
                mrow.age6 = age6;
                mrow.balance = balance;
                mrow.advance = advance;
                mrow.overdue = overdue;
                mrow.legal = legal;
                mrow.oneyear = oneyear;
                mList.Add(mrow);

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintOsReport();
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

        public IDictionary<string, object> SmanWiseList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();

            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                branch = SearchData["branch"].ToString();
                sman = SearchData["sman"].ToString();
                year_code = SearchData["year_code"].ToString();
              
                sCaption = branch + " / " + sman;

                Con_Oracle = new DBConnection();



                sql += " select branch, branch_code,op_sman_name,cust_name, ";
                sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1, ";
                sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                sql += " sum( case when os_days between 91 and 180  then  balance  else 0 end) as age5, ";
                sql += " sum( case when os_days > 180  then  balance  else 0 end) as age6, ";

                sql += " sum( case when nvl(jv_od_type,'N') <> 'N'  then  balance  else 0 end) as legal, ";

                sql += " sum( case when overdue > 0  then  balance  else 0 end) as overdue, ";

                sql += " sum( case when os_days >= 365  then  balance  else 0 end) as oneyear, ";

                sql += " sum(balance) as balance, ";
                sql += " sum(adv) as advance ";
                sql += " from os where branch_code = '" + branch_code + "'  ";

                if ( sman == "")
                    sql += " and op_sman_name is null ";
                else
                    sql += " and op_sman_name = '" + sman + "'";

                if (sCondition != "")
                    sql += " and " + sCondition;

                sql += " group by branch, branch_code, op_sman_name, cust_name ";
                sql += " order by op_sman_name, cust_name ";


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new OsRep();
                    mrow.row_type = "DETAIL";
                    mrow.branch = Dr["branch"].ToString();
                    mrow.branch_code = Dr["branch_code"].ToString();
                    mrow.sman = Dr["op_sman_name"].ToString();
                    mrow.party = Dr["cust_name"].ToString();
                    mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    mrow.advance = Lib.Conv2Decimal(Dr["advance"].ToString());
                    mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.legal = Lib.Conv2Decimal(Dr["legal"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    age1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    age2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    age3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    age4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    age5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    age6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                    overdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                    advance += Lib.Conv2Decimal(Dr["advance"].ToString());
                    legal += Lib.Conv2Decimal(Dr["legal"].ToString());
                    oneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());
                    mList.Add(mrow);

                }

                mrow = new OsRep();
                mrow.row_type = "TOTAL";
                mrow.row_colour = "RED";
                mrow.branch = "TOTAL";
                mrow.branch_code = "TOTAL";
                mrow.sman = "TOTAL";
                mrow.party = "TOTAL";
                mrow.age1 = age1;
                mrow.age2 = age2;
                mrow.age3 = age3;
                mrow.age4 = age4;
                mrow.age5 = age5;
                mrow.age6 = age6;
                mrow.balance = balance;
                mrow.advance = advance;
                mrow.overdue = overdue;
                mrow.legal = legal;
                mrow.oneyear = oneyear;
                mList.Add(mrow);

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintOsReport();
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

        public IDictionary<string, object> InvWiseList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();

            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                branch = SearchData["branch"].ToString();
                sman = SearchData["sman"].ToString();
                year_code = SearchData["year_code"].ToString();
                party = SearchData["party"].ToString();

                if (SearchData.ContainsKey("previos_component_type"))
                {
                    previous_component_type = SearchData["previos_component_type"].ToString();
                }

                sCaption = branch + " / " + sman + " / " + party ;

                Con_Oracle = new DBConnection();

                sql = "";
                sql += " select branch, branch_code,op_sman_name, jv_pkid,jvh_type, jvh_vrno, jvh_date, cust_code,cust_name, ";
                sql += " case when os_days <= 15 then  balance  else 0 end as age1, ";
                sql += " case when os_days between 16 and 30 then  balance  else 0 end as age2, ";
                sql += " case when os_days between 31 and 60 then  balance  else 0 end as age3, ";
                sql += " case when os_days between 61 and 90 then  balance  else 0 end as age4, ";
                sql += " case when os_days between 91 and 180  then  balance  else 0 end as age5, ";
                sql += " case when os_days > 180  then  balance  else 0 end as age6, ";

                sql += " case when nvl(jv_od_type,'N') <> 'N'  then  balance  else 0 end as legal, ";

                sql += " case when overdue > 0  then  balance  else 0 end as overdue, ";

                sql += "  case when os_days >= 365  then  balance  else 0 end as oneyear, ";

                sql += " invtype as jv_inv_category, ";

                sql += " balance as balance, ";
                sql += " adv as advance ";
                
                sql += " from os where branch_code = '" + branch_code + "'  ";
                if (sman == "")
                    sql += " and op_sman_name is null ";
                else
                    sql += " and op_sman_name = '" + sman + "'";
                sql += " and cust_name = '" + party + "'";

                if (sCondition != "")
                    sql += " and " + sCondition;

                sql += " order by jvh_date, jvh_vrno ";

                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new OsRep();
                    mrow.row_type = "DETAIL";
                    mrow.branch = Dr["branch"].ToString();
                    mrow.branch_code = Dr["branch_code"].ToString();
                    mrow.sman = Dr["op_sman_name"].ToString();
                    mrow.party = Dr["cust_name"].ToString();

                    mrow.pkid = Dr["jv_pkid"].ToString();


                    mrow.jv_inv_category = Dr["jv_inv_category"].ToString();

                    mrow.invno = Dr["jvh_type"].ToString() + "-" + Dr["jvh_vrno"].ToString();

                    if (!Dr["jvh_date"].Equals(DBNull.Value))
                        mrow.invdate = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);

                    mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    mrow.advance = Lib.Conv2Decimal(Dr["advance"].ToString());
                    mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.legal = Lib.Conv2Decimal(Dr["legal"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    age1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    age2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    age3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    age4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    age5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    age6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                    overdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                    advance += Lib.Conv2Decimal(Dr["advance"].ToString());
                    legal += Lib.Conv2Decimal(Dr["legal"].ToString());
                    oneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());
                    mList.Add(mrow);

                }

                mrow = new OsRep();
                mrow.row_type = "TOTAL";
                mrow.row_colour = "RED";
                mrow.branch = "TOTAL";
                mrow.branch_code = "TOTAL";
                mrow.invno = "TOTAL";
                mrow.sman = "TOTAL";
                mrow.party = "TOTAL";
                mrow.age1 = age1;
                mrow.age2 = age2;
                mrow.age3 = age3;
                mrow.age4 = age4;
                mrow.age5 = age5;
                mrow.age6 = age6;
                mrow.balance = balance;
                mrow.advance = advance;
                mrow.overdue = overdue;
                mrow.legal = legal;
                mrow.oneyear = oneyear;
                mList.Add(mrow);

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintOsReport();
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


        public IDictionary<string, object> LegalWiseList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();
            
            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch = SearchData["branch"].ToString();
                sman = SearchData["sman"].ToString();
               // sman = "LEGAL";
               
                
                sCaption = "LEGAL";


                Con_Oracle = new DBConnection();


                sql += " select cust_name,branch, branch_code,op_sman_name, ";
                sql += " sum( case when os_days <= 15 then  balance  else 0 end) as age1, ";
                sql += " sum( case when os_days between 16 and 30 then  balance  else 0 end) as age2, ";
                sql += " sum( case when os_days between 31 and 60 then  balance  else 0 end) as age3, ";
                sql += " sum( case when os_days between 61 and 90 then  balance  else 0 end) as age4, ";
                sql += " sum( case when os_days between 91 and 180  then  balance  else 0 end) as age5, ";
                sql += " sum( case when os_days > 180  then  balance  else 0 end) as age6, ";

                sql += " sum( case when nvl(jv_od_type,'N') <> 'N'  then  balance  else 0 end) as legal, ";

                sql += " sum( case when overdue > 0  then  balance  else 0 end) as overdue, ";

                sql += " sum( case when os_days >= 365  then  balance  else 0 end) as oneyear, ";

                sql += " sum(balance) as balance, ";
                sql += " sum(adv) as advance ";

                sql += " from os where nvl(jv_od_type,'N') <> 'N' ";
                
                sql += " group by cust_name,op_sman_name,branch, branch_code ";
                sql += " order by cust_name,op_sman_name, branch ";


                sql = sql.Replace("{COMPCODE}", company_code);
                sql = sql.Replace("{BRCODE}", branch_code);
                sql = sql.Replace("{YEARCODE}", year_code);
                Con_Oracle = new DBConnection();
                Dt_List = new DataTable();
                Dt_List = Con_Oracle.ExecuteQuery(sql);
                Con_Oracle.CloseConnection();

                foreach (DataRow Dr in Dt_List.Rows)
                {
                    mrow = new OsRep();
                    mrow.row_type = "DETAIL";
                    mrow.branch = Dr["branch"].ToString();
                    mrow.branch_code = Dr["branch_code"].ToString();
                    mrow.party = Dr["cust_name"].ToString();

                    mrow.sman = Dr["op_sman_name"].ToString();
                    mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                    mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                    mrow.age3 = Lib.Conv2Decimal(Dr["age3"].ToString());
                    mrow.age4 = Lib.Conv2Decimal(Dr["age4"].ToString());
                    mrow.age5 = Lib.Conv2Decimal(Dr["age5"].ToString());
                    mrow.age6 = Lib.Conv2Decimal(Dr["age6"].ToString());
                    mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                    mrow.advance = Lib.Conv2Decimal(Dr["advance"].ToString());
                    mrow.overdue = Lib.Conv2Decimal(Dr["overdue"].ToString());
                    mrow.legal = Lib.Conv2Decimal(Dr["legal"].ToString());
                    mrow.oneyear = Lib.Conv2Decimal(Dr["oneyear"].ToString());

                    age1 += Lib.Conv2Decimal(Dr["age1"].ToString());
                    age2 += Lib.Conv2Decimal(Dr["age2"].ToString());
                    age3 += Lib.Conv2Decimal(Dr["age3"].ToString());
                    age4 += Lib.Conv2Decimal(Dr["age4"].ToString());
                    age5 += Lib.Conv2Decimal(Dr["age5"].ToString());
                    age6 += Lib.Conv2Decimal(Dr["age6"].ToString());

                    overdue += Lib.Conv2Decimal(Dr["overdue"].ToString());
                    balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                    advance += Lib.Conv2Decimal(Dr["advance"].ToString());
                    legal += Lib.Conv2Decimal(Dr["legal"].ToString());
                    oneyear += Lib.Conv2Decimal(Dr["oneyear"].ToString());
                    mList.Add(mrow);

                }

                mrow = new OsRep();
                mrow.row_type = "TOTAL";
                mrow.row_colour = "RED";
                mrow.branch = "TOTAL";
                mrow.branch_code = "TOTAL";
                mrow.party = "TOTAL";
                mrow.age1 = age1;
                mrow.age2 = age2;
                mrow.age3 = age3;
                mrow.age4 = age4;
                mrow.age5 = age5;
                mrow.age6 = age6;
                mrow.balance = balance;
                mrow.advance = advance;
                mrow.overdue = overdue;
                mrow.legal = legal;
                mrow.oneyear = oneyear;
                mList.Add(mrow);

                if (type == "EXCEL")
                {
                    if (mList != null)
                        PrintOsReport();
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

        private void PrintOsReport()
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

                File_Display_Name = "osrep.xls";
                File_Name = Lib.GetFileName(report_folder, PKID, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];
                


                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 20;
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
                WS.Columns[13].Width = 256 * 15;

                iRow = 0; iCol = 1;

                WS.Columns[2].Style.NumberFormat = "#0,0.00";
                WS.Columns[3].Style.NumberFormat = "#0,0.00";
                WS.Columns[4].Style.NumberFormat = "#0,0.00";
                WS.Columns[5].Style.NumberFormat = "#0,0.00";
                WS.Columns[6].Style.NumberFormat = "#0,0.00";
                WS.Columns[7].Style.NumberFormat = "#0,0.00";
                WS.Columns[8].Style.NumberFormat = "#0,0.00";
                WS.Columns[9].Style.NumberFormat = "#0,0.00";
                WS.Columns[10].Style.NumberFormat = "#0,0.00";
                WS.Columns[11].Style.NumberFormat = "#0,0.00";
                WS.Columns[12].Style.NumberFormat = "#0,0.00";

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

                if ( sCaption != "")
                    Lib.WriteData(WS, iRow, 1, "OS REPORT - " + sCaption, _Color, true, "", "L", "", 12, false, 325, "", true);
                else 
                    Lib.WriteData(WS, iRow, 1, "OS REPORT ", _Color, true, "", "L", "", 12, false, 325, "", true);

                iRow++;
                iRow++;

                iCol = 1;
                if (subtype == "SMANLIST")
                    Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                else if (subtype == "BRANCHWISE")
                    Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                else if (subtype == "SMANWISE")
                    Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                else if (subtype == "INVWISE")
                {
                    Lib.WriteData(WS, iRow, iCol++, "INVOICE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "CATEGORY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                else if (subtype == "SMANPARTYBRANCHWISE")
                {
                    Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }
                else if (subtype == "LEGAL")
                {
                    Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "SALESMAN", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                }

                else
                    Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                Lib.WriteData(WS, iRow, iCol++, "0-15", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "16-30", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "31-60", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "61-90", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "91-180", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "180+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                
                
                Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "ADVANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "OVERDUE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "1YEAR+", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                if (subtype == "LEGAL" || previous_component_type == "LEGAL")
                {
                    Lib.WriteData(WS, iRow, iCol++, "", _Color, true, " ", "R", "", _Size, false, 325, "", true);
                }
                else
                {
                    Lib.WriteData(WS, iRow, iCol++, "LEGAL", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                }
                    

                foreach (OsRep  Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        if (subtype == "SMANLIST")
                            Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        else if (subtype == "BRANCHWISE")
                            Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        else if (subtype == "SMANWISE")
                            Lib.WriteData(WS, iRow, iCol++, Rec.party, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        else if (subtype == "INVWISE")
                        {                        
                            Lib.WriteData(WS, iRow, iCol++, Rec.invno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.invdate, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.jv_inv_category, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        else if (subtype == "SMANPARTYBRANCHWISE")
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.party, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        else if (subtype == "LEGAL")
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.party, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, false, "", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        }
                        else
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age1 , _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age2, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age3, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age4, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age5, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age6, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        

                        
                        Lib.WriteData(WS, iRow, iCol++, Rec.balance, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.advance, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.overdue, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.oneyear, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        if (subtype == "LEGAL" || previous_component_type == "LEGAL")
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "", "R", "", _Size, false, 325, "", true);
                        }
                        else
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.legal, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        }
                        

                    }
                    if (Rec.row_type == "TOTAL")
                    {
                        if (subtype == "SMANLIST")
                            Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, true, "BT", "L", "", _Size, false, 325, "", true);

                        else if (subtype == "BRANCHWISE")
                            Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        else if (subtype == "SMANWISE")
                            Lib.WriteData(WS, iRow, iCol++, Rec.sman, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        else if (subtype == "INVWISE")
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.invno, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        else if (subtype == "SMANPARTYBRANCHWISE")
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.party, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++,"", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        else if (subtype == "LEGAL")
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.party, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        }
                        else
                            Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age1, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age2, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age3, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age4, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age5, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age6, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        

                        
                        Lib.WriteData(WS, iRow, iCol++, Rec.balance, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.advance, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.overdue, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.oneyear, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        if (subtype == "LEGAL" || previous_component_type == "LEGAL")
                        {
                            Lib.WriteData(WS, iRow, iCol++, "", _Color, true, "", "R", "", _Size, false, 325, "", true);
                        }
                        else
                        {
                            Lib.WriteData(WS, iRow, iCol++, Rec.legal, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
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


        public IDictionary<string, object> AirPaymentList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();
            string sql1 = "";
            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();


                isadmin = (Boolean)SearchData["isadmin"];
                iscompany = (Boolean)SearchData["iscompany"];
                filter_branch_code = SearchData["filter_branch_code"].ToString();
                filter_sman_id = SearchData["filter_sman_id"].ToString();
                filter_sman_name = SearchData["filter_sman_name"].ToString();




                string sDate = "";
                string eDate = "";
                
                DateTime Dt = System.DateTime.Now;

                Dt = Dt.AddMonths(-1);

                if (Dt.Day < 15)
                {

                    sDate = "01-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                    eDate = "15-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                }
                else
                {
                    sDate = "16-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                    eDate = DateTime.DaysInMonth(Dt.Year, Dt.Month) + "-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                }
                period = "Period " + sDate + " TO " + eDate;


                sCaption = "Branch Wise";

                sql1 = "";
                sql1 += " select jvh_pkid,jvh_cc_id, m.hbl_pkid,hbl_exp_id, a.rec_branch_code,jvh_date, jvh_acc_id, acc_name, e.param_name as sman_name,";
                sql1 += " g.param_name as carrier_name, ";
                sql1 += " jvh_net_amt as total,";
                sql1 += " count(*) over (partition by m.hbl_pkid)  as reccount";
                sql1 += " from ledgerh a";
                sql1 += " inner join hblm m on a.jvh_cc_id = m.hbl_pkid ";
                sql1 += " left join hblm h on m.hbl_pkid = h.hbl_mbl_id ";
                sql1 += " left join customerm shpr on h.hbl_exp_id = shpr.cust_pkid     ";
                sql1 += " left join param e on shpr.cust_sman_id = e.param_pkid ";
                sql1 += " left join acctm f on a.jvh_acc_id = f.acc_pkid  ";
                sql1 += " left join param  g on m.hbl_carrier_id = g.param_pkid";
                sql1 += " where  a.rec_company_code = '{COMPCODE}' and a.rec_branch_code not in('HOCPL','KOLAF') and jvh_type  =  'PN'";
                sql1 += " and m.hbl_type ='MBL-AE' and jvh_date between '{SDATE}' and '{EDATE}'";

                sql1 = sql1.Replace("{COMPCODE}", company_code);
                sql1 = sql1.Replace("{SDATE}", sDate);
                sql1 = sql1.Replace("{EDATE}", eDate);

                if (iscompany == false)
                {
                    if (isadmin)
                        sql1 += " and a.rec_branch_code = '" + filter_branch_code + "' ";
                    else
                        sql1 += " cust_sman_id = '" + filter_sman_id + "' ";
                }

                Con_Oracle = new DBConnection();

                if (subtype == "BRANCH")
                {
                    sql = "";
                    sql += " select * from (			";
                    sql += " select rec_branch_code as branch,";
                    sql += " sum(round(total * 1/reccount)) as A1, '2018' as  period ";
                    sql += " from (";

                    sql += sql1;

                    sql += " ) a ";
                    sql += " group by rec_branch_code";
                    sql += " ) a where a1 > 0";
                    sql += " order by branch";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new OsRep();
                        mrow.row_type = "DETAIL";
                        mrow.branch = Dr["branch"].ToString();
                        mrow.balance = Lib.Conv2Decimal(Dr["a1"].ToString());
                        balance += Lib.Conv2Decimal(Dr["a1"].ToString());
                        mList.Add(mrow);
                    }
                    mrow = new OsRep();
                    mrow.row_type = "TOTAL";
                    mrow.branch = "TOTAL";
                    mrow.branch_code = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.balance = balance;
                    mList.Add(mrow);
                }

                if (subtype == "CARRIER")
                {
                    sql = "";
                    sql += " select * from (			";
                    sql += " select carrier_name,";
                    sql += " sum(round(total * 1/reccount)) as A1, '2018' as  period ";
                    sql += " from (";

                    sql += sql1;

                    sql += " ) a ";
                    sql += " group by carrier_name";
                    sql += " ) a where a1 > 0";
                    sql += " order by carrier_name";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new OsRep();
                        mrow.row_type = "DETAIL";
                        mrow.branch = Dr["carrier_name"].ToString();
                        mrow.balance = Lib.Conv2Decimal(Dr["a1"].ToString());
                        balance += Lib.Conv2Decimal(Dr["a1"].ToString());
                        mList.Add(mrow);
                    }
                    mrow = new OsRep();
                    mrow.row_type = "TOTAL";
                    mrow.branch = "TOTAL";
                    mrow.branch_code = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.balance = balance;
                    mList.Add(mrow);
                }

                if (subtype == "SALESMAN")
                {
                    sql = "";
                    sql += " select * from (			";
                    sql += " select sman_name,";
                    sql += " sum(round(total * 1/reccount)) as A1, '2018' as  period ";
                    sql += " from (";

                    sql += sql1;

                    sql += " ) a ";
                    sql += " group by sman_name";
                    sql += " ) a where a1 > 0";
                    sql += " order by sman_name";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new OsRep();
                        mrow.row_type = "DETAIL";
                        mrow.branch = Dr["sman_name"].ToString();
                        mrow.balance = Lib.Conv2Decimal(Dr["a1"].ToString());
                        balance += Lib.Conv2Decimal(Dr["a1"].ToString());
                        mList.Add(mrow);
                    }
                    mrow = new OsRep();
                    mrow.row_type = "TOTAL";
                    mrow.branch = "TOTAL";
                    mrow.branch_code = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.balance = balance;
                    mList.Add(mrow);
                }


                if (subtype == "BRANCH-SALESMAN")
                {
                    sql = "";
                    sql += " select * from ( ";
                    sql += " select rec_branch_code as branch,sman_name,";
                    sql += " sum(round(total * 1/reccount)) as A1, '2018' as  period ";
                    sql += " from (";

                    sql += sql1;

                    sql += " ) a ";
                    sql += " group by rec_branch_code,sman_name";
                    sql += " ) a where a1 > 0";
                    sql += " order by branch,sman_name";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new OsRep();
                        mrow.row_type = "DETAIL";
                        mrow.branch = Dr["branch"].ToString();
                        mrow.sman = Dr["sman_name"].ToString();
                        mrow.balance = Lib.Conv2Decimal(Dr["a1"].ToString());
                        balance += Lib.Conv2Decimal(Dr["a1"].ToString());
                        mList.Add(mrow);
                    }
                    mrow = new OsRep();
                    mrow.row_type = "TOTAL";
                    mrow.branch = "TOTAL";
                    mrow.branch_code = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.balance = balance;
                    mList.Add(mrow);
                }


                if (subtype == "BRANCH-CARRIER")
                {
                    sql = "";
                    sql += " select * from ( ";
                    sql += " select rec_branch_code as branch,carrier_name,";
                    sql += " sum(round(total * 1/reccount)) as A1, '2018' as  period ";
                    sql += " from (";

                    sql += sql1;

                    sql += " ) a ";
                    sql += " group by rec_branch_code,carrier_name";
                    sql += " ) a where a1 > 0";
                    sql += " order by branch,carrier_name";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new OsRep();
                        mrow.row_type = "DETAIL";
                        mrow.branch = Dr["branch"].ToString();
                        mrow.sman = Dr["carrier_name"].ToString();
                        mrow.balance = Lib.Conv2Decimal(Dr["a1"].ToString());
                        balance += Lib.Conv2Decimal(Dr["a1"].ToString());
                        mList.Add(mrow);
                    }
                    mrow = new OsRep();
                    mrow.row_type = "TOTAL";
                    mrow.branch = "TOTAL";
                    mrow.branch_code = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.balance = balance;
                    mList.Add(mrow);
                }




                Con_Oracle.CloseConnection();



                foreach (OsRep Dr in mList)
                {
                    if (Dr.branch == "ABDSF")
                        Dr.branch = "AHAMADABAD";
                    if (Dr.branch == "DELAF")
                        Dr.branch = "DELHI AIR";
                    if (Dr.branch == "MBYAF")
                        Dr.branch = "MUMBAI AIR";
                    if (Dr.branch == "CHNAF")
                        Dr.branch = "CHENNAI AIR";
                    if (Dr.branch == "COKAF")
                        Dr.branch = "KOCHI AIR";
                    if (Dr.branch == "BLRAF")
                        Dr.branch = "BANGALORE";

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
            RetData.Add("period", period);
            return RetData;
        }

        public IDictionary<string, object> AirInvList(Dictionary<string, object> SearchData)
        {
            Dictionary<string, object> RetData = new Dictionary<string, object>();
            mList = new List<OsRep>();
            string sql1 = "";
            try
            {

                type = SearchData["type"].ToString();
                subtype = SearchData["subtype"].ToString();
                report_folder = SearchData["report_folder"].ToString();
                PKID = SearchData["pkid"].ToString();
                company_code = SearchData["company_code"].ToString();
                branch_code = SearchData["branch_code"].ToString();
                year_code = SearchData["year_code"].ToString();


                isadmin = (Boolean)SearchData["isadmin"];
                iscompany = (Boolean)SearchData["iscompany"];
                filter_branch_code = SearchData["filter_branch_code"].ToString();
                filter_sman_id = SearchData["filter_sman_id"].ToString();
                filter_sman_name = SearchData["filter_sman_name"].ToString();

                balance = 0;
                dr = 0;
                cr = 0;


                string sDate = "";
                string eDate = "";

                DateTime Dt = System.DateTime.Now;

                Dt = Dt.AddMonths(-1);

                if (Dt.Day < 15)
                {

                    sDate = "01-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                    eDate = "15-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                }
                else
                {
                    sDate = "16-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                    eDate = DateTime.DaysInMonth(Dt.Year, Dt.Month) + "-" + Dt.ToString("MMM") + "-" + Dt.Year.ToString();
                }
                period = "Period " + sDate + " TO " + eDate;


                sCaption = "Branch Wise";

                

                sql1 = " select h.rec_branch_code as branch_code, master.hbl_bl_no, jv_pkid,jv_acc_id,jvh_cc_id, jvh_vrno,jvh_docno,jvh_date,max(L.rec_category) as INVTYPE, ";
                sql1 += " jv_debit, nvl(sum(xref_amt), 0) as jv_credit, jv_debit - nvl(sum(xref_amt), 0) as balance ";
                sql1 += " from ledgerh h ";
                sql1 += " inner join ledgert L on (h.jvh_pkid = L.jv_parent_id) ";
                sql1 += " inner join hblm house on jvh_cc_id = house.hbl_pkid ";
                sql1 += " inner join hblm master on house.hbl_mbl_id = master.hbl_pkid ";
                sql1 += " inner join Acctm a on (L.jv_acc_id = A.acc_pkid) ";
                sql1 += " inner join acgroupm g on a.acc_group_id = g.acgrp_pkid ";
                sql1 += " left join ledgerxref X on (L.jv_pkid = X.xref_dr_jv_id) ";
                sql1 += " left join param s on (jv_acc_id = param_pkid) ";
                sql1 += " where h.rec_company_code = '{COMPCODE}' and h.rec_branch_code not in('HOCPL','KOLAF') ";
                sql1 += " and master.hbl_date between '{SDATE}' and  '{EDATE}' and master.hbl_type = 'MBL-AE' ";
                sql1 += " and L.jv_debit > 0 and h.rec_deleted = 'N' and acc_against_invoice = 'D' and jvh_type not in('OP', 'OB', 'OC') ";

                if (subtype != "BRANCH" && subtype != "TOTAL")
                    sql1 += " and h.rec_branch_code = '" + subtype  + "' ";
                sql1 += " group by h.rec_branch_code, master.hbl_bl_no,jv_pkid,jv_acc_id,jvh_cc_id,jvh_date,jvh_vrno,jvh_docno,jv_debit ";

                sql1 = sql1.Replace("{COMPCODE}", company_code);
                sql1 = sql1.Replace("{SDATE}", sDate);
                sql1 = sql1.Replace("{EDATE}", eDate);

                if (iscompany == false)
                {
                    if (isadmin)
                        sql1 += " and h.rec_branch_code = '" + filter_branch_code + "' ";
                }

                Con_Oracle = new DBConnection();

                if (subtype == "BRANCH")
                {
                    sql = "";
                    sql += " select branch_code,comp_name,  sum(jv_debit) as age1, sum(jv_credit) as age2, sum(balance) as balance  from ( ";
                    sql += sql1;
                    sql += " ) a ";
                    sql += " inner join companym on branch_code = comp_code ";
                    sql += " group by branch_code,comp_name ";
                    sql += " order by comp_name ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new OsRep();
                        mrow.row_type = "DETAIL";
                        mrow.branch = Dr["comp_name"].ToString();
                        mrow.branch_code = Dr["branch_code"].ToString();
                        mrow.age1 = Lib.Conv2Decimal(Dr["age1"].ToString());
                        mrow.age2 = Lib.Conv2Decimal(Dr["age2"].ToString());
                        mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                        dr += Lib.Conv2Decimal(Dr["age1"].ToString());
                        cr += Lib.Conv2Decimal(Dr["age2"].ToString());
                        balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                        mList.Add(mrow);
                    }
                    mrow = new OsRep();
                    mrow.row_type = "TOTAL";
                    mrow.branch = "TOTAL";
                    mrow.branch_code = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.age1 = dr;
                    mrow.age2 = cr;
                    mrow.balance = balance;
                    mList.Add(mrow);
                }

                else
                {
                    sql = "";
                    sql += " select branch_code,comp_name,cust_name,hbl_bl_no,jvh_docno,jvh_date,jv_debit, jv_credit, balance ";
                    sql += " from ( ";
                    sql += sql1;
                    sql += " ) a ";
                    sql += " left join customerm on jv_acc_id = cust_pkid ";
                    sql += " inner join companym on branch_code = comp_code ";
                    sql += " order by comp_name, cust_name, jvh_docno ";

                    sql = sql.Replace("{COMPCODE}", company_code);
                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{YEARCODE}", year_code);

                    Dt_List = new DataTable();
                    Dt_List = Con_Oracle.ExecuteQuery(sql);
                    foreach (DataRow Dr in Dt_List.Rows)
                    {
                        mrow = new OsRep();
                        mrow.row_type = "DETAIL";
                        mrow.branch = Dr["comp_name"].ToString();
                        mrow.branch_code = Dr["branch_code"].ToString();
                        mrow.awbno = Dr["hbl_bl_no"].ToString();
                        mrow.party = Dr["cust_name"].ToString();
                        mrow.invno = Dr["jvh_docno"].ToString();
                        mrow.invdate = Lib.DatetoStringDisplayformat(Dr["jvh_date"]);
                        mrow.age1 = Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                        mrow.age2 = Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        mrow.balance = Lib.Conv2Decimal(Dr["balance"].ToString());
                        dr += Lib.Conv2Decimal(Dr["jv_debit"].ToString());
                        cr += Lib.Conv2Decimal(Dr["jv_credit"].ToString());
                        balance += Lib.Conv2Decimal(Dr["balance"].ToString());
                        mList.Add(mrow);
                    }
                    mrow = new OsRep();
                    mrow.row_type = "TOTAL";
                    mrow.branch = "TOTAL";
                    mrow.branch_code = "TOTAL";
                    mrow.row_colour = "RED";
                    mrow.age1 = dr;
                    mrow.age2 = cr;
                    mrow.balance = balance;
                    mList.Add(mrow);
                }

                Con_Oracle.CloseConnection();

                if(type== "EXCEL")
                {
                    PrintInvListReport();
                }
          
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
            RetData.Add("period", period);
            return RetData;
        }


        private void PrintInvListReport()
        {
            string str = "";
            string COMPNAME = "";
            string COMPADD1 = "";
            string COMPADD2 = "";
            string COMPTEL = "";
            string COMPFAX = "";
            string COMPWEB = "";

            string _Border = "";
            Boolean _Bold = false;
            Color _Color = Color.Black;
            int _Size = 10;

            string FolderId = "";
            iRow = 0;
            iCol = 0;
            try
            {
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


                File_Display_Name = "Report.xls";
                FolderId = Guid.NewGuid().ToString().ToUpper();
                File_Name = Lib.GetFileName(report_folder, FolderId, File_Display_Name);

                string sName = "Report";
                WB = new ExcelFile();
                WB.Worksheets.Add(sName);
                WS = WB.Worksheets[sName];

                WS.ViewOptions.ShowGridLines = false;
                WS.PrintOptions.FitWorksheetWidthToPages = 1;

                WS.Columns[0].Width = 256 * 2;
                WS.Columns[1].Width = 256 * 30;
                WS.Columns[2].Width = 256 * 15;
                WS.Columns[3].Width = 256 * 40;
                WS.Columns[4].Width = 256 * 15;
                WS.Columns[5].Width = 256 * 10;
                WS.Columns[6].Width = 256 * 15;
                WS.Columns[7].Width = 256 * 15;
                WS.Columns[8].Width = 256 * 15;


                iRow = 0; iCol = 1;

                iRow++;
                _Size = 14;
                Lib.WriteData(WS, iRow, 1, COMPNAME, _Color, true, "", "L", "", _Size, false, 325, "", true);
                _Size = 10;
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
                Lib.WriteData(WS, iRow, 1, "O/S Against AirLine Payment - " + period, _Color, true, "", "L", "", 15, false, 325, "", true);
                iRow++;
                iRow++;
                _Size = 10;
                iCol = 1;
                Lib.WriteData(WS, iRow, iCol++, "BRANCH", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "AWB", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "PARTY", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "INVNO", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DATE", _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "DEBIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "CREDIT", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                Lib.WriteData(WS, iRow, iCol++, "BALANCE", _Color, true, "BT", "R", "", _Size, false, 325, "", true);
                foreach (OsRep Rec in mList)
                {
                    iRow++;
                    iCol = 1;
                    if (Rec.row_type == "DETAIL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.awbno , _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.party, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.invno, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.invdate, _Color, false, "", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age1, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age2, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.balance, _Color, false, "", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
                    else if (Rec.row_type == "TOTAL")
                    {
                        Lib.WriteData(WS, iRow, iCol++, Rec.branch, _Color, true, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, "", _Color, false, "BT", "L", "", _Size, false, 325, "", true);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age1, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.age2, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                        Lib.WriteData(WS, iRow, iCol++, Rec.balance, _Color, true, "BT", "R", "", _Size, false, 325, "#,0.00;(#,0.00);#", false);
                    }
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

